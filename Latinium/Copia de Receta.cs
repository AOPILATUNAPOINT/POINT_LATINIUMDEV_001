using System;
using System.Data;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using Infragistics.Win.UltraWinGrid;
using C1.Data;
using CrystalDecisions.Shared;
using System.Data.SqlClient;

namespace Latinium
{
	/// <summary>
	/// Descripción breve de Receta1.
	/// </summary>
	public class Receta : DevExpress.XtraEditors.XtraForm
	{
		private C1.Data.C1DataSet cdsProducto;
		private C1.Data.C1DataView cdvProducto;
		private Infragistics.Win.UltraWinGrid.UltraGrid ultraGrid1;
		private C1.Data.C1DataView cdvArtCodigo;
		private C1.Data.C1DataSet cdsArticulo;
		private C1.Data.C1DataView cdvArticulo;
		private C1.Data.C1DataSet cdsSeguridad;
		private C1.Data.C1DataView cdvProdNombre;
		private C1.Data.C1DataView cdvProdCod;
		private System.Windows.Forms.ErrorProvider errorProvider;
		private BarraDatoSQL.BarraDatoSQL barraDato1;
		private Infragistics.Win.UltraWinGrid.UltraDropDown cmbArticulo;
		private Infragistics.Win.UltraWinGrid.UltraDropDown cmbCodArt;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor ultraTextEditor1;
		private System.Windows.Forms.Label label2;
		private Infragistics.Win.Misc.UltraButton btBuscaArticulo;
		private System.Windows.Forms.ImageList imageList;
		private C1.Data.C1DataSet cdsSeteoF;
		private C1.Data.C1DataSet cdsTerminado;
		private System.Windows.Forms.ToolTip toolTip1;
		private Infragistics.Win.Misc.UltraButton btTotal;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtTotal;
		private Infragistics.Win.Misc.UltraButton btImprime;
		private System.Windows.Forms.Label label5;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtidArticulo;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtIdReceta;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtArticulo;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtCodigo;
		private System.ComponentModel.IContainer components;

		public Receta()
		{
			InitializeComponent();
		}
		int IdReceta = 0;
		C1.Data.C1DataRow drReceta;
		bool bEliminaAlValidar = false;

		public Receta(int idReceta)
		{
			InitializeComponent();
			IdReceta = idReceta;
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

		#region Windows Form Designer generated code
		/// <summary>
		/// Método necesario para admitir el Diseñador, no se puede modificar
		/// el contenido del método con el editor de código.
		/// </summary>
		private void InitializeComponent()
		{
			this.components = new System.ComponentModel.Container();
			Infragistics.Win.Appearance appearance1 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand1 = new Infragistics.Win.UltraWinGrid.UltraGridBand("_ProductoDetalle", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn1 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idProdDetalle");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn2 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idProducto");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn3 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idArticulo", -1, "cmbArticulo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn4 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Cantidad");
			Infragistics.Win.Appearance appearance2 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn5 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Precio");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn6 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Costo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn7 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Impuesto");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn8 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Importante");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn9 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Borrar");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn10 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("FechaIngreso");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn11 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idSucursal");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn12 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Proceso");
			Infragistics.Win.Appearance appearance3 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn13 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Proporcion");
			Infragistics.Win.Appearance appearance4 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn14 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Materia");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn15 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("CodMateria");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn16 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCodArt", 0, "cmbCodArt", 0, Infragistics.Win.UltraWinGrid.SortIndicator.Ascending, false);
			Infragistics.Win.Appearance appearance5 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance6 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance7 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance8 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance9 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand2 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Articulo", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn17 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Descontinuado");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn18 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idArticulo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn19 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("dCompra");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn20 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("dVenta");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn21 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idTipoGrupo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn22 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Articulo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn23 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Codigo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn24 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idGrupoArticulo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn25 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("ExigeLote");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn26 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Cantidad6");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn27 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("E_idEducativo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn28 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("CantConf");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn29 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Origen");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn30 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("CodAlterno");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn31 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("SeriePropiaAlterna");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn32 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Cantidad1");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn33 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Cantidad2");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn34 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Cantidad3");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn35 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Cantidad4");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn36 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Cantidad5");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn37 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("ManejaSerial");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn38 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("AplicaDscto");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn39 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Facturable");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn40 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("ValTransfTemporal");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn41 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Modelo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn42 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCategoria");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn43 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("PrecioManual");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn44 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("PeriodoGarantia");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn45 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("CodArtAntiguo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn46 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("CodigoReemplazo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn47 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Tamaño");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn48 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Combo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn49 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Vida_Util");
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand3 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Articulo", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn50 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Descontinuado");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn51 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idArticulo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn52 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("dCompra");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn53 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("dVenta");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn54 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idTipoGrupo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn55 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Articulo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn56 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Codigo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn57 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idGrupoArticulo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn58 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("ExigeLote");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn59 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Cantidad6");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn60 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("E_idEducativo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn61 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("CantConf");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn62 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Origen");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn63 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("CodAlterno");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn64 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("SeriePropiaAlterna");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn65 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Cantidad1");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn66 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Cantidad2");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn67 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Cantidad3");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn68 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Cantidad4");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn69 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Cantidad5");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn70 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("ManejaSerial");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn71 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("AplicaDscto");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn72 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Facturable");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn73 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("ValTransfTemporal");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn74 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Modelo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn75 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCategoria");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn76 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("PrecioManual");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn77 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("PeriodoGarantia");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn78 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("CodArtAntiguo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn79 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("CodigoReemplazo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn80 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Tamaño");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn81 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Combo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn82 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Vida_Util");
			Infragistics.Win.Appearance appearance10 = new Infragistics.Win.Appearance();
			System.Resources.ResourceManager resources = new System.Resources.ResourceManager(typeof(Receta));
			Infragistics.Win.Appearance appearance11 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance12 = new Infragistics.Win.Appearance();
			this.cdsProducto = new C1.Data.C1DataSet();
			this.cdvProducto = new C1.Data.C1DataView();
			this.ultraGrid1 = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.cdvArtCodigo = new C1.Data.C1DataView();
			this.cdsArticulo = new C1.Data.C1DataSet();
			this.cdvArticulo = new C1.Data.C1DataView();
			this.cmbArticulo = new Infragistics.Win.UltraWinGrid.UltraDropDown();
			this.cmbCodArt = new Infragistics.Win.UltraWinGrid.UltraDropDown();
			this.cdsSeguridad = new C1.Data.C1DataSet();
			this.cdvProdNombre = new C1.Data.C1DataView();
			this.cdsTerminado = new C1.Data.C1DataSet();
			this.cdvProdCod = new C1.Data.C1DataView();
			this.errorProvider = new System.Windows.Forms.ErrorProvider();
			this.barraDato1 = new BarraDatoSQL.BarraDatoSQL();
			this.ultraTextEditor1 = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.label2 = new System.Windows.Forms.Label();
			this.btBuscaArticulo = new Infragistics.Win.Misc.UltraButton();
			this.imageList = new System.Windows.Forms.ImageList(this.components);
			this.cdsSeteoF = new C1.Data.C1DataSet();
			this.toolTip1 = new System.Windows.Forms.ToolTip(this.components);
			this.btTotal = new Infragistics.Win.Misc.UltraButton();
			this.txtTotal = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.btImprime = new Infragistics.Win.Misc.UltraButton();
			this.label5 = new System.Windows.Forms.Label();
			this.txtidArticulo = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtIdReceta = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtArticulo = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.txtCodigo = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			((System.ComponentModel.ISupportInitialize)(this.cdsProducto)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdvProducto)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraGrid1)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdvArtCodigo)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsArticulo)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdvArticulo)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbArticulo)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbCodArt)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeguridad)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdvProdNombre)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsTerminado)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdvProdCod)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraTextEditor1)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtTotal)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtidArticulo)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIdReceta)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtArticulo)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtCodigo)).BeginInit();
			this.SuspendLayout();
			// 
			// cdsProducto
			// 
			this.cdsProducto.BindingContextCtrl = this;
			this.cdsProducto.DataLibrary = "LibFacturacion";
			this.cdsProducto.DataLibraryUrl = "";
			this.cdsProducto.DataSetDef = "dsProducto";
			this.cdsProducto.FillOnRequest = false;
			this.cdsProducto.Locale = new System.Globalization.CultureInfo("es-EC");
			this.cdsProducto.Name = "cdsProducto";
			this.cdsProducto.SchemaClassName = "LibFacturacion.DataClass";
			this.cdsProducto.SchemaDef = null;
			this.cdsProducto.BeforeFill += new C1.Data.FillEventHandler(this.cdsProducto_BeforeFill);
			// 
			// cdvProducto
			// 
			this.cdvProducto.BindingContextCtrl = this;
			this.cdvProducto.DataSet = this.cdsProducto;
			this.cdvProducto.Sort = "idProducto";
			this.cdvProducto.TableName = "";
			this.cdvProducto.TableViewName = "Producto";
			// 
			// ultraGrid1
			// 
			this.ultraGrid1.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.ultraGrid1.Cursor = System.Windows.Forms.Cursors.Default;
			this.ultraGrid1.DataMember = "_Producto.Producto - ProductoDetalle";
			this.ultraGrid1.DataSource = this.cdsProducto;
			appearance1.BackColor = System.Drawing.Color.White;
			this.ultraGrid1.DisplayLayout.Appearance = appearance1;
			this.ultraGrid1.DisplayLayout.AutoFitColumns = true;
			ultraGridBand1.AddButtonCaption = "_Producto";
			ultraGridColumn1.Header.VisiblePosition = 11;
			ultraGridColumn1.Hidden = true;
			ultraGridColumn2.Header.VisiblePosition = 10;
			ultraGridColumn2.Hidden = true;
			ultraGridColumn3.Header.Caption = "Artículo";
			ultraGridColumn3.Header.VisiblePosition = 4;
			ultraGridColumn3.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.DropDownList;
			ultraGridColumn3.Width = 343;
			appearance2.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn4.CellAppearance = appearance2;
			ultraGridColumn4.Format = "#,##0.00";
			ultraGridColumn4.Header.VisiblePosition = 0;
			ultraGridColumn4.Width = 99;
			ultraGridColumn5.Header.VisiblePosition = 13;
			ultraGridColumn5.Hidden = true;
			ultraGridColumn6.Header.VisiblePosition = 9;
			ultraGridColumn6.Hidden = true;
			ultraGridColumn7.Header.VisiblePosition = 12;
			ultraGridColumn7.Hidden = true;
			ultraGridColumn8.Header.Caption = "Imp.";
			ultraGridColumn8.Header.VisiblePosition = 5;
			ultraGridColumn8.Hidden = true;
			ultraGridColumn8.Width = 34;
			ultraGridColumn9.Header.VisiblePosition = 8;
			ultraGridColumn9.Hidden = true;
			ultraGridColumn10.Header.VisiblePosition = 14;
			ultraGridColumn10.Hidden = true;
			ultraGridColumn11.Header.VisiblePosition = 15;
			ultraGridColumn11.Hidden = true;
			appearance3.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn12.CellAppearance = appearance3;
			ultraGridColumn12.Header.VisiblePosition = 6;
			ultraGridColumn12.Hidden = true;
			ultraGridColumn12.Width = 53;
			appearance4.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn13.CellAppearance = appearance4;
			ultraGridColumn13.Header.VisiblePosition = 7;
			ultraGridColumn13.Hidden = true;
			ultraGridColumn13.Width = 64;
			ultraGridColumn14.Header.Caption = "Artículo";
			ultraGridColumn14.Header.VisiblePosition = 2;
			ultraGridColumn14.Hidden = true;
			ultraGridColumn14.Width = 388;
			ultraGridColumn15.Header.Caption = "Codigo";
			ultraGridColumn15.Header.VisiblePosition = 1;
			ultraGridColumn15.Hidden = true;
			ultraGridColumn15.Width = 200;
			ultraGridColumn16.Header.Caption = "Codigo";
			ultraGridColumn16.Header.VisiblePosition = 3;
			ultraGridColumn16.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.DropDownList;
			ultraGridColumn16.Width = 217;
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
																										 ultraGridColumn16});
			this.ultraGrid1.DisplayLayout.BandsSerializer.Add(ultraGridBand1);
			appearance5.ForeColor = System.Drawing.Color.Black;
			appearance5.ForeColorDisabled = System.Drawing.Color.Black;
			this.ultraGrid1.DisplayLayout.Override.ActiveRowAppearance = appearance5;
			this.ultraGrid1.DisplayLayout.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.TemplateOnBottom;
			this.ultraGrid1.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.True;
			this.ultraGrid1.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.True;
			appearance6.BackColor = System.Drawing.Color.Transparent;
			this.ultraGrid1.DisplayLayout.Override.CardAreaAppearance = appearance6;
			appearance7.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance7.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance7.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance7.FontData.BoldAsString = "True";
			appearance7.FontData.Name = "Arial";
			appearance7.FontData.SizeInPoints = 10F;
			appearance7.ForeColor = System.Drawing.Color.White;
			appearance7.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.ultraGrid1.DisplayLayout.Override.HeaderAppearance = appearance7;
			this.ultraGrid1.DisplayLayout.Override.HeaderClickAction = Infragistics.Win.UltraWinGrid.HeaderClickAction.SortSingle;
			appearance8.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance8.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance8.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ultraGrid1.DisplayLayout.Override.RowSelectorAppearance = appearance8;
			appearance9.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance9.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance9.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ultraGrid1.DisplayLayout.Override.SelectedRowAppearance = appearance9;
			this.ultraGrid1.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.ultraGrid1.Location = new System.Drawing.Point(8, 40);
			this.ultraGrid1.Name = "ultraGrid1";
			this.ultraGrid1.Size = new System.Drawing.Size(680, 320);
			this.ultraGrid1.TabIndex = 1;
			this.ultraGrid1.InitializeRow += new Infragistics.Win.UltraWinGrid.InitializeRowEventHandler(this.ultraGrid1_InitializeRow);
			this.ultraGrid1.BeforeRowsDeleted += new Infragistics.Win.UltraWinGrid.BeforeRowsDeletedEventHandler(this.ultraGrid1_BeforeRowsDeleted);
			this.ultraGrid1.KeyDown += new System.Windows.Forms.KeyEventHandler(this.ultraGrid1_KeyDown);
			this.ultraGrid1.AfterCellUpdate += new Infragistics.Win.UltraWinGrid.CellEventHandler(this.ultraGrid1_AfterCellUpdate);
			this.ultraGrid1.InitializeLayout += new Infragistics.Win.UltraWinGrid.InitializeLayoutEventHandler(this.ultraGrid1_InitializeLayout);
			// 
			// cdvArtCodigo
			// 
			this.cdvArtCodigo.BindingContextCtrl = this;
			this.cdvArtCodigo.DataSet = this.cdsArticulo;
			this.cdvArtCodigo.Sort = "Codigo";
			this.cdvArtCodigo.TableName = "";
			this.cdvArtCodigo.TableViewName = "Articulo";
			// 
			// cdsArticulo
			// 
			this.cdsArticulo.BindingContextCtrl = this;
			this.cdsArticulo.CaseSensitive = false;
			this.cdsArticulo.DataLibrary = "LibFacturacion";
			this.cdsArticulo.DataLibraryUrl = "";
			this.cdsArticulo.DataSetDef = "dsArticulos";
			this.cdsArticulo.Locale = new System.Globalization.CultureInfo("es-EC");
			this.cdsArticulo.Name = "cdsArticulo";
			this.cdsArticulo.SchemaClassName = "LibFacturacion.DataClass";
			this.cdsArticulo.SchemaDef = null;
			this.cdsArticulo.BeforeFill += new C1.Data.FillEventHandler(this.cdsArticulo_BeforeFill);
			// 
			// cdvArticulo
			// 
			this.cdvArticulo.BindingContextCtrl = this;
			this.cdvArticulo.DataSet = this.cdsArticulo;
			this.cdvArticulo.Sort = "Articulo";
			this.cdvArticulo.TableName = "";
			this.cdvArticulo.TableViewName = "Articulo";
			// 
			// cmbArticulo
			// 
			this.cmbArticulo.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbArticulo.DataSource = this.cdvArticulo;
			ultraGridColumn17.Header.VisiblePosition = 2;
			ultraGridColumn17.Hidden = true;
			ultraGridColumn18.Header.VisiblePosition = 3;
			ultraGridColumn18.Hidden = true;
			ultraGridColumn19.Header.VisiblePosition = 4;
			ultraGridColumn19.Hidden = true;
			ultraGridColumn20.Header.VisiblePosition = 5;
			ultraGridColumn20.Hidden = true;
			ultraGridColumn21.Header.VisiblePosition = 6;
			ultraGridColumn21.Hidden = true;
			ultraGridColumn22.Header.VisiblePosition = 0;
			ultraGridColumn22.Width = 281;
			ultraGridColumn23.Header.VisiblePosition = 1;
			ultraGridColumn24.Header.VisiblePosition = 7;
			ultraGridColumn24.Hidden = true;
			ultraGridColumn25.Header.VisiblePosition = 8;
			ultraGridColumn25.Hidden = true;
			ultraGridColumn26.Header.VisiblePosition = 9;
			ultraGridColumn27.Header.VisiblePosition = 10;
			ultraGridColumn28.Header.VisiblePosition = 11;
			ultraGridColumn29.Header.VisiblePosition = 12;
			ultraGridColumn30.Header.VisiblePosition = 13;
			ultraGridColumn31.Header.VisiblePosition = 14;
			ultraGridColumn32.Header.VisiblePosition = 15;
			ultraGridColumn33.Header.VisiblePosition = 16;
			ultraGridColumn34.Header.VisiblePosition = 17;
			ultraGridColumn35.Header.VisiblePosition = 18;
			ultraGridColumn36.Header.VisiblePosition = 19;
			ultraGridColumn37.Header.VisiblePosition = 20;
			ultraGridColumn38.Header.VisiblePosition = 21;
			ultraGridColumn39.Header.VisiblePosition = 22;
			ultraGridColumn40.Header.VisiblePosition = 23;
			ultraGridColumn41.Header.VisiblePosition = 24;
			ultraGridColumn42.Header.VisiblePosition = 25;
			ultraGridColumn43.Header.VisiblePosition = 26;
			ultraGridColumn43.Hidden = true;
			ultraGridColumn44.Header.VisiblePosition = 27;
			ultraGridColumn44.Hidden = true;
			ultraGridColumn45.Header.VisiblePosition = 28;
			ultraGridColumn46.Header.VisiblePosition = 29;
			ultraGridColumn47.Header.VisiblePosition = 30;
			ultraGridColumn48.Header.VisiblePosition = 31;
			ultraGridColumn49.Header.VisiblePosition = 32;
			ultraGridBand2.Columns.AddRange(new object[] {
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
																										 ultraGridColumn27,
																										 ultraGridColumn28,
																										 ultraGridColumn29,
																										 ultraGridColumn30,
																										 ultraGridColumn31,
																										 ultraGridColumn32,
																										 ultraGridColumn33,
																										 ultraGridColumn34,
																										 ultraGridColumn35,
																										 ultraGridColumn36,
																										 ultraGridColumn37,
																										 ultraGridColumn38,
																										 ultraGridColumn39,
																										 ultraGridColumn40,
																										 ultraGridColumn41,
																										 ultraGridColumn42,
																										 ultraGridColumn43,
																										 ultraGridColumn44,
																										 ultraGridColumn45,
																										 ultraGridColumn46,
																										 ultraGridColumn47,
																										 ultraGridColumn48,
																										 ultraGridColumn49});
			this.cmbArticulo.DisplayLayout.BandsSerializer.Add(ultraGridBand2);
			this.cmbArticulo.DisplayMember = "Articulo";
			this.cmbArticulo.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbArticulo.Location = new System.Drawing.Point(48, 168);
			this.cmbArticulo.Name = "cmbArticulo";
			this.cmbArticulo.Size = new System.Drawing.Size(280, 86);
			this.cmbArticulo.TabIndex = 11;
			this.cmbArticulo.Text = "cmbArticulo";
			this.cmbArticulo.ValueMember = "idArticulo";
			this.cmbArticulo.Visible = false;
			// 
			// cmbCodArt
			// 
			this.cmbCodArt.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbCodArt.DataSource = this.cdvArtCodigo;
			ultraGridColumn50.Header.VisiblePosition = 2;
			ultraGridColumn50.Hidden = true;
			ultraGridColumn51.Header.VisiblePosition = 3;
			ultraGridColumn51.Hidden = true;
			ultraGridColumn52.Header.VisiblePosition = 4;
			ultraGridColumn52.Hidden = true;
			ultraGridColumn53.Header.VisiblePosition = 5;
			ultraGridColumn53.Hidden = true;
			ultraGridColumn54.Header.VisiblePosition = 6;
			ultraGridColumn54.Hidden = true;
			ultraGridColumn55.Header.VisiblePosition = 1;
			ultraGridColumn55.Width = 245;
			ultraGridColumn56.Header.VisiblePosition = 0;
			ultraGridColumn56.Width = 113;
			ultraGridColumn57.Header.VisiblePosition = 7;
			ultraGridColumn57.Hidden = true;
			ultraGridColumn58.Header.VisiblePosition = 8;
			ultraGridColumn58.Hidden = true;
			ultraGridColumn59.Header.VisiblePosition = 9;
			ultraGridColumn60.Header.VisiblePosition = 10;
			ultraGridColumn61.Header.VisiblePosition = 11;
			ultraGridColumn62.Header.VisiblePosition = 12;
			ultraGridColumn63.Header.VisiblePosition = 13;
			ultraGridColumn64.Header.VisiblePosition = 14;
			ultraGridColumn65.Header.VisiblePosition = 15;
			ultraGridColumn66.Header.VisiblePosition = 16;
			ultraGridColumn67.Header.VisiblePosition = 17;
			ultraGridColumn68.Header.VisiblePosition = 18;
			ultraGridColumn69.Header.VisiblePosition = 19;
			ultraGridColumn70.Header.VisiblePosition = 20;
			ultraGridColumn71.Header.VisiblePosition = 21;
			ultraGridColumn72.Header.VisiblePosition = 22;
			ultraGridColumn73.Header.VisiblePosition = 23;
			ultraGridColumn74.Header.VisiblePosition = 24;
			ultraGridColumn75.Header.VisiblePosition = 25;
			ultraGridColumn76.Header.VisiblePosition = 26;
			ultraGridColumn76.Hidden = true;
			ultraGridColumn77.Header.VisiblePosition = 27;
			ultraGridColumn77.Hidden = true;
			ultraGridColumn78.Header.VisiblePosition = 28;
			ultraGridColumn79.Header.VisiblePosition = 29;
			ultraGridColumn80.Header.VisiblePosition = 30;
			ultraGridColumn81.Header.VisiblePosition = 31;
			ultraGridColumn82.Header.VisiblePosition = 32;
			ultraGridBand3.Columns.AddRange(new object[] {
																										 ultraGridColumn50,
																										 ultraGridColumn51,
																										 ultraGridColumn52,
																										 ultraGridColumn53,
																										 ultraGridColumn54,
																										 ultraGridColumn55,
																										 ultraGridColumn56,
																										 ultraGridColumn57,
																										 ultraGridColumn58,
																										 ultraGridColumn59,
																										 ultraGridColumn60,
																										 ultraGridColumn61,
																										 ultraGridColumn62,
																										 ultraGridColumn63,
																										 ultraGridColumn64,
																										 ultraGridColumn65,
																										 ultraGridColumn66,
																										 ultraGridColumn67,
																										 ultraGridColumn68,
																										 ultraGridColumn69,
																										 ultraGridColumn70,
																										 ultraGridColumn71,
																										 ultraGridColumn72,
																										 ultraGridColumn73,
																										 ultraGridColumn74,
																										 ultraGridColumn75,
																										 ultraGridColumn76,
																										 ultraGridColumn77,
																										 ultraGridColumn78,
																										 ultraGridColumn79,
																										 ultraGridColumn80,
																										 ultraGridColumn81,
																										 ultraGridColumn82});
			this.cmbCodArt.DisplayLayout.BandsSerializer.Add(ultraGridBand3);
			this.cmbCodArt.DisplayMember = "Codigo";
			this.cmbCodArt.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbCodArt.Location = new System.Drawing.Point(344, 168);
			this.cmbCodArt.Name = "cmbCodArt";
			this.cmbCodArt.Size = new System.Drawing.Size(288, 86);
			this.cmbCodArt.TabIndex = 12;
			this.cmbCodArt.Text = "cmbArtCodigo";
			this.cmbCodArt.ValueMember = "idArticulo";
			this.cmbCodArt.Visible = false;
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
			this.cdsSeguridad.BeforeFill += new C1.Data.FillEventHandler(this.cdsSeguridad_BeforeFill);
			// 
			// cdvProdNombre
			// 
			this.cdvProdNombre.BindingContextCtrl = this;
			this.cdvProdNombre.DataSet = this.cdsTerminado;
			this.cdvProdNombre.Sort = "Articulo";
			this.cdvProdNombre.TableName = "";
			this.cdvProdNombre.TableViewName = "Articulo";
			// 
			// cdsTerminado
			// 
			this.cdsTerminado.BindingContextCtrl = this;
			this.cdsTerminado.CaseSensitive = false;
			this.cdsTerminado.DataLibrary = "LibFacturacion";
			this.cdsTerminado.DataLibraryUrl = "";
			this.cdsTerminado.DataSetDef = "dsArticulos";
			this.cdsTerminado.Locale = new System.Globalization.CultureInfo("es-EC");
			this.cdsTerminado.Name = "cdsTerminado";
			this.cdsTerminado.SchemaClassName = "LibFacturacion.DataClass";
			this.cdsTerminado.SchemaDef = null;
			this.cdsTerminado.BeforeFill += new C1.Data.FillEventHandler(this.cdsTerminado_BeforeFill);
			// 
			// cdvProdCod
			// 
			this.cdvProdCod.BindingContextCtrl = this;
			this.cdvProdCod.DataSet = this.cdsTerminado;
			this.cdvProdCod.Sort = "Codigo";
			this.cdvProdCod.TableName = "";
			this.cdvProdCod.TableViewName = "Articulo";
			// 
			// errorProvider
			// 
			this.errorProvider.ContainerControl = this;
			// 
			// barraDato1
			// 
			this.barraDato1.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.barraDato1.BarraMovimiento = true;
			this.barraDato1.DataMember = "_Producto";
			this.barraDato1.DataNombreId = "idProducto";
			this.barraDato1.DataOrden = "";
			this.barraDato1.DataSource = this.cdsProducto;
			this.barraDato1.DataTabla = "Producto";
			this.barraDato1.DataTablaHija = "ProductoDetalle";
			this.barraDato1.Location = new System.Drawing.Point(8, 448);
			this.barraDato1.Name = "barraDato1";
			this.barraDato1.Size = new System.Drawing.Size(304, 25);
			this.barraDato1.TabIndex = 15;
			this.barraDato1.VisibleBorrar = true;
			this.barraDato1.VisibleBuscar = true;
			this.barraDato1.VisibleEditar = true;
			this.barraDato1.VisibleImprimir = true;
			this.barraDato1.VisibleNuevo = true;
			this.barraDato1.Buscar += new System.EventHandler(this.barraDato1_Buscar);
			this.barraDato1.Load += new System.EventHandler(this.barraDato1_Load);
			this.barraDato1.Borrar += new System.EventHandler(this.barraDato1_Borrar);
			this.barraDato1.Imprimir += new System.EventHandler(this.barraDato1_Imprimir);
			this.barraDato1.Restaurar += new System.EventHandler(this.barraDato1_Restaurar);
			this.barraDato1.Refresca += new System.EventHandler(this.barraDatoSQL1_Refresca);
			this.barraDato1.Crear += new System.EventHandler(this.barraDato1_Crear);
			this.barraDato1.Deshacer += new System.EventHandler(this.barraDato1_Deshacer);
			this.barraDato1.Copiar += new System.EventHandler(this.barraDato1_Copiar);
			this.barraDato1.Impresora += new System.EventHandler(this.barraDato1_Impresora);
			this.barraDato1.Editar += new System.EventHandler(this.barraDato1_Editar);
			this.barraDato1.Grabar += new System.EventHandler(this.barraDato1_Grabar);
			// 
			// ultraTextEditor1
			// 
			this.ultraTextEditor1.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.ultraTextEditor1.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.ultraTextEditor1.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsProducto, "Producto.Nota"));
			this.ultraTextEditor1.Location = new System.Drawing.Point(64, 368);
			this.ultraTextEditor1.Multiline = true;
			this.ultraTextEditor1.Name = "ultraTextEditor1";
			this.ultraTextEditor1.Scrollbars = System.Windows.Forms.ScrollBars.Vertical;
			this.ultraTextEditor1.Size = new System.Drawing.Size(624, 64);
			this.ultraTextEditor1.TabIndex = 17;
			// 
			// label2
			// 
			this.label2.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.label2.AutoSize = true;
			this.label2.Location = new System.Drawing.Point(8, 368);
			this.label2.Name = "label2";
			this.label2.Size = new System.Drawing.Size(37, 16);
			this.label2.TabIndex = 18;
			this.label2.Text = "Notas:";
			// 
			// btBuscaArticulo
			// 
			this.btBuscaArticulo.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			appearance10.Image = 6;
			this.btBuscaArticulo.Appearance = appearance10;
			this.btBuscaArticulo.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			this.btBuscaArticulo.ImageList = this.imageList;
			this.btBuscaArticulo.Location = new System.Drawing.Point(328, 448);
			this.btBuscaArticulo.Name = "btBuscaArticulo";
			this.btBuscaArticulo.Size = new System.Drawing.Size(24, 24);
			this.btBuscaArticulo.TabIndex = 19;
			this.toolTip1.SetToolTip(this.btBuscaArticulo, "Ver Producta / Materia Prima");
			this.btBuscaArticulo.Click += new System.EventHandler(this.btBuscaArticulo_Click);
			this.btBuscaArticulo.MouseDown += new System.Windows.Forms.MouseEventHandler(this.btBuscaArticulo_MouseDown);
			// 
			// imageList
			// 
			this.imageList.ImageSize = new System.Drawing.Size(16, 16);
			this.imageList.ImageStream = ((System.Windows.Forms.ImageListStreamer)(resources.GetObject("imageList.ImageStream")));
			this.imageList.TransparentColor = System.Drawing.Color.Transparent;
			// 
			// cdsSeteoF
			// 
			this.cdsSeteoF.BindingContextCtrl = this;
			this.cdsSeteoF.CaseSensitive = false;
			this.cdsSeteoF.DataLibrary = "LibFacturacion";
			this.cdsSeteoF.DataLibraryUrl = "";
			this.cdsSeteoF.DataSetDef = "dsSeteoF";
			this.cdsSeteoF.Locale = new System.Globalization.CultureInfo("es-EC");
			this.cdsSeteoF.Name = "cdsSeteoF";
			this.cdsSeteoF.SchemaClassName = "LibFacturacion.DataClass";
			this.cdsSeteoF.SchemaDef = null;
			// 
			// btTotal
			// 
			this.btTotal.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			this.btTotal.Location = new System.Drawing.Point(560, 448);
			this.btTotal.Name = "btTotal";
			this.btTotal.Size = new System.Drawing.Size(48, 26);
			this.btTotal.TabIndex = 20;
			this.btTotal.Text = "Total";
			this.btTotal.Visible = false;
			this.btTotal.Click += new System.EventHandler(this.btTotal_Click);
			// 
			// txtTotal
			// 
			this.txtTotal.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			this.txtTotal.Location = new System.Drawing.Point(616, 448);
			this.txtTotal.Name = "txtTotal";
			this.txtTotal.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtTotal.PromptChar = ' ';
			this.txtTotal.Size = new System.Drawing.Size(64, 22);
			this.txtTotal.TabIndex = 21;
			this.txtTotal.Visible = false;
			// 
			// btImprime
			// 
			this.btImprime.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			this.btImprime.Location = new System.Drawing.Point(480, 448);
			this.btImprime.Name = "btImprime";
			this.btImprime.Size = new System.Drawing.Size(72, 26);
			this.btImprime.TabIndex = 22;
			this.btImprime.Text = "Imprime";
			this.btImprime.Visible = false;
			this.btImprime.Click += new System.EventHandler(this.btImprime_Click);
			this.btImprime.MouseUp += new System.Windows.Forms.MouseEventHandler(this.btImprime_MouseUp);
			// 
			// label5
			// 
			this.label5.AutoSize = true;
			this.label5.Font = new System.Drawing.Font("Tahoma", 8.25F);
			this.label5.ForeColor = System.Drawing.Color.FromArgb(((System.Byte)(50)), ((System.Byte)(50)), ((System.Byte)(50)));
			this.label5.Location = new System.Drawing.Point(8, 8);
			this.label5.Name = "label5";
			this.label5.Size = new System.Drawing.Size(42, 17);
			this.label5.TabIndex = 162;
			this.label5.Text = "Artículo";
			this.label5.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// txtidArticulo
			// 
			this.txtidArticulo.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			this.txtidArticulo.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsProducto, "_Producto.idArticulo"));
			this.txtidArticulo.Location = new System.Drawing.Point(368, 448);
			this.txtidArticulo.Name = "txtidArticulo";
			this.txtidArticulo.PromptChar = ' ';
			this.txtidArticulo.ReadOnly = true;
			this.txtidArticulo.Size = new System.Drawing.Size(48, 22);
			this.txtidArticulo.TabIndex = 165;
			this.txtidArticulo.Visible = false;
			// 
			// txtIdReceta
			// 
			this.txtIdReceta.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			this.txtIdReceta.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsProducto, "_Producto.idProducto"));
			this.txtIdReceta.Location = new System.Drawing.Point(424, 448);
			this.txtIdReceta.Name = "txtIdReceta";
			this.txtIdReceta.PromptChar = ' ';
			this.txtIdReceta.ReadOnly = true;
			this.txtIdReceta.Size = new System.Drawing.Size(48, 22);
			this.txtIdReceta.TabIndex = 166;
			this.txtIdReceta.Visible = false;
			// 
			// txtArticulo
			// 
			appearance11.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtArticulo.Appearance = appearance11;
			this.txtArticulo.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtArticulo.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtArticulo.Location = new System.Drawing.Point(320, 5);
			this.txtArticulo.Name = "txtArticulo";
			this.txtArticulo.Size = new System.Drawing.Size(368, 22);
			this.txtArticulo.TabIndex = 169;
			// 
			// txtCodigo
			// 
			appearance12.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtCodigo.Appearance = appearance12;
			this.txtCodigo.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtCodigo.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtCodigo.Location = new System.Drawing.Point(64, 5);
			this.txtCodigo.Name = "txtCodigo";
			this.txtCodigo.Size = new System.Drawing.Size(248, 22);
			this.txtCodigo.TabIndex = 168;
			this.txtCodigo.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtCodigo_KeyDown);
			// 
			// Receta
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 14);
			this.ClientSize = new System.Drawing.Size(696, 478);
			this.Controls.Add(this.txtArticulo);
			this.Controls.Add(this.txtCodigo);
			this.Controls.Add(this.txtIdReceta);
			this.Controls.Add(this.txtidArticulo);
			this.Controls.Add(this.label5);
			this.Controls.Add(this.label2);
			this.Controls.Add(this.btImprime);
			this.Controls.Add(this.txtTotal);
			this.Controls.Add(this.btTotal);
			this.Controls.Add(this.btBuscaArticulo);
			this.Controls.Add(this.ultraTextEditor1);
			this.Controls.Add(this.barraDato1);
			this.Controls.Add(this.cmbCodArt);
			this.Controls.Add(this.cmbArticulo);
			this.Controls.Add(this.ultraGrid1);
			this.KeyPreview = true;
			this.Name = "Receta";
			this.Text = "Receta";
			this.KeyDown += new System.Windows.Forms.KeyEventHandler(this.Receta_KeyDown);
			this.Closing += new System.ComponentModel.CancelEventHandler(this.Receta_Closing);
			this.Load += new System.EventHandler(this.Receta1_Load);
			((System.ComponentModel.ISupportInitialize)(this.cdsProducto)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdvProducto)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraGrid1)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdvArtCodigo)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsArticulo)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdvArticulo)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbArticulo)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbCodArt)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeguridad)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdvProdNombre)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsTerminado)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdvProdCod)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraTextEditor1)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtTotal)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtidArticulo)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIdReceta)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtArticulo)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtCodigo)).EndInit();
			this.ResumeLayout(false);

		}
		#endregion

		private void barraDato1_Borrar(object sender, System.EventArgs e)
		{
			if (!miAcceso.Eliminar)
			{
				MessageBox.Show("No tiene acceso a Eliminar Recetas",
					"Información", MessageBoxButtons.OK, MessageBoxIcon.Information);
				return;
			}
			if (DialogResult.No == MessageBox.Show(
				"¿Está seguro que desea borrar la Receta?"
				+ ".\nEste proceso no se puede reversar", 
				"Confirmación", MessageBoxButtons.YesNo, MessageBoxIcon.Stop, 
				MessageBoxDefaultButton.Button2)) return;
			
			this.barraDato1.BorraRegistro();
		}

		private void BuscaIdArt(int Id)
		{
			string stSelect = "SELECT Count(*) FROM Producto "+
				" WHERE idArticulo = " + Id.ToString();
			int Cuenta = Funcion.iEscalarSQL(cdsProducto, stSelect, true);
			if (Cuenta == 0)
			{
				MessageBox.Show("No existe esa receta", "Información");
				return;
			}
			stSelect = "SELECT idProducto FROM Producto "+
				" WHERE idArticulo = " + Id.ToString();
			int idProd = Funcion.iEscalarSQL(cdsProducto, stSelect);
			if (idProd > 0)
			{
				barraDato1.IdRegistro = idProd;
				barraDato1.ProximoId(5);
			}
		}

		private void barraDato1_Buscar(object sender, System.EventArgs e)
		{
			if (!miAcceso.Buscar)
			{
				MessageBox.Show("No tiene acceso a Buscar Recetas",
					"Información", MessageBoxButtons.OK, MessageBoxIcon.Information);
				return;
			}

			using (BuscaReceta miBusqueda = new BuscaReceta())
			{
				if (miBusqueda.ShowDialog() == DialogResult.OK)
				{
					if (miBusqueda.idBusca.Text.Length == 0) return;
					int IdReceta = int.Parse(miBusqueda.idBusca.Text);
					BuscaIdArt(IdReceta);
				}
			}
		}

		private void barraDato1_Crear(object sender, System.EventArgs e)
		{
			if (!miAcceso.Nuevo)
			{
				MessageBox.Show("No tiene acceso a Crear Nuevas Recetas", "Información", MessageBoxButtons.OK, MessageBoxIcon.Information);
				return;
			}
			this.barraDato1.CrearRegistro();
			ModificaGrilla(true);
			this.txtArticulo.Enabled = false;	
		}

		private void barraDato1_Grabar(object sender, System.EventArgs e)
		{
			#region Validacion
			if (this.txtCodigo.Text.ToString().Length == 0 || this.txtArticulo.Text.ToString().Length == 0 || drReceta["idArticulo"] == System.DBNull.Value)
			{
				MessageBox.Show("Seleccione un articulo.", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				this.txtCodigo.Focus();
				return;
			}
				
			string stSelect = string.Format("Select Count(*) From Producto Where idArticulo = {0} And idProducto <> {1}", (int)drReceta["idArticulo"], (int)drReceta["idProducto"]);
			int iCuenta = Funcion.iEscalarSQL(cdsProducto, stSelect);
			if (iCuenta > 0)
			{
				MessageBox.Show("Ya existe una receta con este producto.", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				this.txtCodigo.Focus();
				return;
			}
			
			#region Fila vacia
			foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in ultraGrid1.Rows.All)
			{
				if (dr.Cells["idArticulo"].Value == System.DBNull.Value || dr.Cells["idCodArt"].Value == System.DBNull.Value)
				{
					bEliminaAlValidar = true;
					dr.Delete();
				}
			}
			#endregion Fila vacia
			#endregion Validacion

			this.barraDato1.GrabaRegistro();

			ModificaGrilla(false);
			bEliminaAlValidar = false;
		}

		private void barraDato1_Editar(object sender, System.EventArgs e)
		{
			if (!miAcceso.Editar)
			{
				MessageBox.Show("No tiene acceso a Editar Recetas", "Información", MessageBoxButtons.OK, MessageBoxIcon.Information);
				return;
			}
			this.barraDato1.EditarRegistro();
			ModificaGrilla(true);
			this.txtArticulo.Enabled = false;
		}

		private void barraDato1_Deshacer(object sender, System.EventArgs e)
		{
			this.barraDato1.DeshacerRegistro();
			ModificaGrilla(false);
		}

		private void ModificaGrilla(bool bModifica)
		{
			ultraGrid1.Enabled = true;
			if (bModifica)
			{
				ultraGrid1.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.True;
				ultraGrid1.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.True;
			}
			else
			{
				ultraGrid1.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.False;
				ultraGrid1.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.False;
			}
		}

		private Acceso miAcceso;
		private C1DataRow drSeteoF;

		private void Receta1_Load(object sender, System.EventArgs e)
		{
			drSeteoF = cdsSeteoF.TableViews["SeteoF"].Rows[0];

			if (drSeteoF["VerDesplegable"] != System.DBNull.Value && !(bool) drSeteoF["VerDesplegable"])
			{
				ultraGrid1.DisplayLayout.Bands[0].Columns["Materia"].Hidden = false;
				ultraGrid1.DisplayLayout.Bands[0].Columns["CodMateria"].Hidden = false;
				ultraGrid1.DisplayLayout.Bands[0].Columns["idArticulo"].Hidden = true;
				ultraGrid1.DisplayLayout.Bands[0].Columns["idCodArt"].Hidden = true;
			}

			string stMoneda = "#,##0.";
			string stCeros = "0000000000";
			
			int iDecProd = Funcion.iEscalarSQL(cdsArticulo, "Select IsNull(DecProduccion, 0) From SeteoF", true);
			string stFormato = stMoneda + stCeros.Substring(0, iDecProd);
			ultraGrid1.DisplayLayout.Bands[0].Columns["Cantidad"].Format = stFormato; 

			string stAudita = "Exec AuditaCrear 53, 6, 'Receta'";
			Funcion.EjecutaSQL(cdsArticulo, stAudita, true);
			Cursor = Cursors.Default;
			miAcceso =new Acceso("0901", cdsSeguridad.StorageDataSet.Tables["Seguridad"]);
			this.txtidArticulo.Width = 0;
			if (IdReceta > 0) 
				BuscaIdArt(IdReceta);
			else
				barraDato1.PosUltima();

			barraDato1.HabilitaControles(false);
			ModificaGrilla(false);
		}

		private void cdsProducto_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			this.cdsProducto.Schema.Connections[0].ConnectionString = 
				miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);
		}
    bool bCargaArticulo=true;

		private void cdsArticulo_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			this.cdsArticulo.Schema.Connections[0].ConnectionString = 
				miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);		

			string stFiltro = "[Descontinuado] = 0 And dCompra <> 1 ";//And dVenta <> 1";
			string stSelect = "Select IsNull(VerDesplegable, 1) From SeteoF";
			bCargaArticulo = Funcion.bEscalarSQL(cdsArticulo, stSelect, true);
			if (!bCargaArticulo)
				stFiltro = "idArticulo < 2";
			C1.Data.SchemaObjects.DataSetDef dataSetDef = 
				e.DataSet.Schema.DataSetDefs["dsArticulos"];
			e.Filter.Add(new C1.Data.FilterCondition(dataSetDef.TableViews["Articulo"],
				stFiltro));		
		}

		private void barraDato1_Imprimir(object sender, System.EventArgs e)
		{
			this.Imprime(0);		
		}

		private void Imprime(int iTipo)
		{
			if (!miAcceso.Imprimir)
			{
				MessageBox.Show("No tiene acceso a Imprimir Recetas",
					"Información", MessageBoxButtons.OK, MessageBoxIcon.Information);
				return;
			}
			Cursor = Cursors.WaitCursor;
			string stFiltro = "";
			string stReporte = "Receta.rpt";
			ParameterFields paramFields = new ParameterFields ();
			ParameterField paramField = new ParameterField ();

			if (iTipo == 0)	stFiltro = "{Producto.IdProducto} = " + this.txtidArticulo.Text;
			else
			{
				stReporte = "RecetaSP.rpt";

				ParameterDiscreteValue discreteVal = new ParameterDiscreteValue ();
				paramField.ParameterFieldName = "@idProducto";
				discreteVal.Value = int.Parse(this.txtidArticulo.Text);
				paramField.CurrentValues.Add (discreteVal);
				paramFields.Add (paramField);
				// Fin de parametros
			}
			Reporte miReporte = new Reporte(stReporte, stFiltro);
			miReporte.MdiParent = this.MdiParent;
			if (iTipo == 1) miReporte.crVista.ParameterFieldInfo = paramFields;
			Cursor = Cursors.Default;
			miReporte.Show();
		}

		private void barraDato1_Impresora(object sender, System.EventArgs e)
		{
			this.Imprime(1);
		}

		private void barraDato1_Copiar(object sender, System.EventArgs e)
		{
			if (!miAcceso.Nuevo)
			{
				MessageBox.Show("No tiene acceso a Copiar o Crear Articulos",
					"Información", MessageBoxButtons.OK, MessageBoxIcon.Information);
				return;
			}

			if (DialogResult.No == MessageBox.Show("¿Desea Copiar Receta?",
				"Confirmación", MessageBoxButtons.YesNo, MessageBoxIcon.Question,
				MessageBoxDefaultButton.Button2)) return;

			string stSelect = "Insert Into Producto (idArticulo) Values (0)";
			string stCod = this.txtidArticulo.Text.ToString();
			Funcion.EjecutaSQL(cdsProducto, stSelect);
			barraDato1.PosUltima();

			// Copia Receta Cuyo id = txtIdProducto
			stSelect = "Exec CopiaRegistros 'ProductoDetalle', "
				+ "'idProducto = "	+ stCod
				+ "', 'idProducto,', '" + this.txtidArticulo.Text.ToString() 
				+  "', 'idProdDetalle,'";
			Funcion.EjecutaSQL(cdsProducto, stSelect);
			barraDato1.PosUltima();
			barraDato1.EditarRegistro();
			cmbArticulo.Focus();
			MessageBox.Show("Copia de Recetas Terminada", "Información");
			ModificaGrilla(true);
		}

		private void barraDato1_Restaurar(object sender, System.EventArgs e)
		{
		
		}

		private void cdsSeguridad_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			this.cdsSeguridad.Schema.Connections[0].ConnectionString = 
				miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);		
		}

		private void barraDatoSQL1_Refresca(object sender, System.EventArgs e)
		{
			drReceta = cdsProducto.TableViews["Producto"].Rows[0];

			this.txtCodigo.Text = "";
			this.txtArticulo.Text = "";

			int idArticulo = 0;
			if (drReceta["idArticulo"] != System.DBNull.Value) idArticulo = (int)drReceta["idArticulo"];
			string sSQL = string.Format("Select Codigo, Articulo From Articulo Where idArticulo = {0}", idArticulo);
			SqlDataReader drArt = Funcion.rEscalarSQL(cdsSeteoF, sSQL, true);
			drArt.Read();
			if (drArt.HasRows)
			{
				this.txtCodigo.Text = drArt.GetString(0);
				this.txtArticulo.Text = drArt.GetString(1);
			}
			drArt.Close();

			txtTotal.Value = 0;
		}
		private bool bActualiza = true;

		private void ultraGrid1_AfterCellUpdate(object sender, Infragistics.Win.UltraWinGrid.CellEventArgs e)
		{
			if (!bActualiza) return;
			bActualiza = false;
			if (e.Cell.IsInEditMode)
				if (e.Cell.Column.ToString() == "idCodArt")
					if (e.Cell.OriginalValue != e.Cell.Value)
					{
						e.Cell.Row.Cells["idArticulo"].Value = e.Cell.Row.Cells["idCodArt"].Value;
					}

			string stFiltro="";
			int IdTipoFactura = 0;
			if (e.Cell.Column.ToString() == "CodMateria")
			{
				string stCod = e.Cell.Row.Cells["CodMateria"].Value.ToString();
				string stSelect = "Select Count(*) From Articulo Where Codigo like '" + stCod + "%'" + stFiltro;
				int iCuenta = Funcion.iEscalarSQL(cdsArticulo, stSelect, true);
				int IdArt = 0;
				if (iCuenta != 1)
				{
					using (BuscaArticulos miBusqueda = new BuscaArticulos(stCod, 1, IdTipoFactura))
					{
						if (DialogResult.OK == miBusqueda.ShowDialog())
						{
							IdArt = miBusqueda.idArticulo;
							e.Cell.Row.Cells["Materia"].Value = Funcion.sEscalarSQL(cdsArticulo, "Select a.Articulo From Articulo a Where a.idArticulo = " + IdArt);
							e.Cell.Row.Cells["CodMateria"].Value = Funcion.sEscalarSQL(cdsArticulo, "Select a.Codigo From Articulo a Where a.idArticulo = " + IdArt);
						}
						else
						{
							bActualiza = true;
							return;						
						}
					}
				}
				else
				{
					stSelect = "Select idArticulo From Articulo Where Codigo like '" + stCod + "%'" + stFiltro;
					IdArt = Funcion.iEscalarSQL(cdsArticulo, stSelect, true);
					stSelect = "Select Codigo From Articulo Where Codigo like '" + stCod + "%'";
					e.Cell.Row.Cells["CodMateria"].Value = Funcion.sEscalarSQL(cdsArticulo, stSelect, true);
					stSelect = "Select Articulo From Articulo Where Codigo like '" + stCod + "%'";
					e.Cell.Row.Cells["Materia"].Value = Funcion.sEscalarSQL(cdsArticulo, stSelect, true);
				}
				e.Cell.Row.Cells["idArticulo"].Value = IdArt;
			}
			if (e.Cell.Column.ToString() == "Materia")
			{
				string stCod = e.Cell.Row.Cells["Materia"].Value.ToString();
				string stSelect = "Select Count(*) From Articulo Where Articulo like '"
					+ stCod + "%'";
				int iCuenta = Funcion.iEscalarSQL(cdsArticulo, stSelect, true);
				int IdArt = 0;
				if (iCuenta != 1)
				{
					using (BuscaArticulos miBusqueda = new BuscaArticulos(stCod, 0, IdTipoFactura))
					{
						if (DialogResult.OK == miBusqueda.ShowDialog())
						{
							IdArt = miBusqueda.idArticulo;
							e.Cell.Row.Cells["Materia"].Value = Funcion.sEscalarSQL(cdsArticulo, 
								"Select a.Articulo From Articulo a Where a.idArticulo = " + IdArt);
							e.Cell.Row.Cells["CodMateria"].Value = Funcion.sEscalarSQL(cdsArticulo, 
								"Select a.Codigo From Articulo a Where a.idArticulo = " + IdArt);
						}
						else
						{
							bActualiza = true;
							return;						
						}
					}
				}
				else
				{
					stSelect = "Select idArticulo From Articulo Where Articulo like '"
						+ stCod + "%'";
					IdArt = Funcion.iEscalarSQL(cdsArticulo, stSelect, true);
					stSelect = "Select Codigo From Articulo Where Articulo like '" + stCod + "%'";
					e.Cell.Row.Cells["CodMateria"].Value = Funcion.sEscalarSQL(cdsArticulo, stSelect, true);
					stSelect = "Select Articulo From Articulo Where Articulo like '" + stCod + "%'";
					e.Cell.Row.Cells["Materia"].Value = Funcion.sEscalarSQL(cdsArticulo, stSelect, true);
				}
				e.Cell.Row.Cells["idArticulo"].Value = IdArt;
			}
			bActualiza = true;
		}

		private void ultraGrid1_InitializeRow(object sender, Infragistics.Win.UltraWinGrid.InitializeRowEventArgs e)
		{
			e.Row.Cells["idCodArt"].Value = e.Row.Cells["idArticulo"].Value;
		}

		private void Receta_Closing(object sender, System.ComponentModel.CancelEventArgs e)
		{
			string stAudita = "Exec AuditaCrear 53, 7, 'Receta'";
			Funcion.EjecutaSQL(cdsArticulo, stAudita, true);
		}

		private void ultraGrid1_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
      if (ultraGrid1.ActiveCell == null) return;
			try
			{
				if (ultraGrid1.ActiveCell.Column.Key.ToString().ToUpper() == "NOTAS") return;
				switch(e.KeyValue)
				{
					case (int) Keys.F3:
 						#region F3 Inventario	
						if (this.ultraGrid1.ActiveCell.Column.Key.ToString().ToUpper() == "CODMATERIA" && barraDato1.bEditar)
						{
							string sArticulo = ""; 
							sArticulo = this.ultraGrid1.ActiveRow.Cells["CodMateria"].Value.ToString();
							int IdArt = 0;
					
							using (BuscaArticulos miBusqueda = new BuscaArticulos(sArticulo, 1, 1))
							{
								if (DialogResult.OK == miBusqueda.ShowDialog())
								{
									IdArt = miBusqueda.idArticulo;
									ultraGrid1.ActiveRow.Cells["CodMateria"].Value = miBusqueda.sCodigo;
								}
							}		
						}
						break;
						#endregion F3 Inventario	

					case (int) Keys.PageDown:
						if(ultraGrid1.ActiveCell.DroppedDown == false)
							ultraGrid1.ActiveCell.DroppedDown = true;
						break;

					case (int) Keys.Up:
						if (ultraGrid1.ActiveCell.DroppedDown) return;

						ultraGrid1.PerformAction(UltraGridAction.ExitEditMode, false, false);
						ultraGrid1.PerformAction(UltraGridAction.AboveCell, false, false);
						e.Handled = true;
						ultraGrid1.PerformAction(UltraGridAction.EnterEditMode, false, false);
						break;
					case (int) Keys.Left:
						ultraGrid1.PerformAction(UltraGridAction.ExitEditMode, false, false);
						ultraGrid1.PerformAction(UltraGridAction.PrevCellByTab, false, false);
						e.Handled = true;
						ultraGrid1.PerformAction(UltraGridAction.EnterEditMode, false, false);
						break;
					case (int) Keys.Down:
						if (ultraGrid1.ActiveCell.DroppedDown) return;
			
						ultraGrid1.PerformAction(UltraGridAction.ExitEditMode, false, false);
						ultraGrid1.PerformAction(UltraGridAction.BelowCell, false, false);
						e.Handled = true;
						ultraGrid1.PerformAction(UltraGridAction.EnterEditMode, false, false);
						break;
					case (int) Keys.Enter:
						ultraGrid1.PerformAction(UltraGridAction.ExitEditMode, false, false);
						ultraGrid1.PerformAction(UltraGridAction.NextCellByTab, false, false);
						e.Handled = true;
						ultraGrid1.PerformAction(UltraGridAction.EnterEditMode, false, false);
						if(ultraGrid1.ActiveCell.DroppedDown == false)
							ultraGrid1.ActiveCell.DroppedDown = true;
						//						ultraGrid1.PerformAction(UltraGridAction.EnterEditModeAndDropdown, false, false);
						break;
				}
			}
			catch(System.Exception ex)
			{
				Console.WriteLine(ex.Message);
				//				MessageBox.Show(ex.Message, "Tecla");
			}

		}

		private void btBuscaArticulo_Click(object sender, System.EventArgs e)
		{
//			if (cmbArtCodigo.Value == System.DBNull.Value)
//			{
//				MessageBox.Show("Seleccione el Articulo para mostrar", "Información");
//				return;
//			}
			Cursor = Cursors.WaitCursor;
			//int idArticulo = (int) cmbArtCodigo.Value;
//			Articulo miArticulo = new Articulo((int)this.txtidArticulo.Value);
//			miArticulo.MdiParent = this.MdiParent;
//			miArticulo.Show();
			Cursor = Cursors.Default;
		}

		private void btImprimirTodo_Click(object sender, System.EventArgs e)
		{
			this.Imprime(0);
		}

		private void cdsTerminado_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			this.cdsTerminado.Schema.Connections[0].ConnectionString = miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);		

			string stFiltro = "idTipoGrupo In (6, 8)"; //"idSubGrupo = 30";
			C1.Data.SchemaObjects.DataSetDef dataSetDef = e.DataSet.Schema.DataSetDefs["dsArticulos"];
			e.Filter.Add(new C1.Data.FilterCondition(dataSetDef.TableViews["Articulo"], stFiltro));		
		}

		private void btBuscaArticulo_MouseDown(object sender, System.Windows.Forms.MouseEventArgs e)
		{
			if (e.Button == MouseButtons.Right)
			{
				if (this.ultraGrid1.ActiveRow.Cells["idArticulo"].Value == null
					|| this.ultraGrid1.ActiveRow.Cells["idArticulo"].Value == System.DBNull.Value)
					return;

				Cursor = Cursors.WaitCursor;
				int idArticulo = (int) ultraGrid1.ActiveRow.Cells["idArticulo"].Value;
//				Articulo miArticulo = new Articulo(idArticulo);
//				miArticulo.MdiParent = this.MdiParent;
//				miArticulo.Show();
				Cursor = Cursors.Default;
			}
		}

		private void btTotal_Click(object sender, System.EventArgs e)
		{
			int idProducto = 0;
			if (this.txtidArticulo.Text.Trim().Length > 0)
				idProducto = int.Parse(txtidArticulo.Text);
			string stExec = "Exec RecetaCostoMP " + idProducto.ToString();
			double dTotal = Funcion.dEscalarSQL(cdsArticulo, stExec, true);
			txtTotal.Value = dTotal;
		}

		private void btImprime_Click(object sender, System.EventArgs e)
		{
			Cursor = Cursors.WaitCursor;
			Reporte miReporte = new Reporte("RecetaTotal.rpt");
			miReporte.MdiParent = this.MdiParent;
			miReporte.Show();
			Cursor = Cursors.Default;
		}

		private void btImprime_MouseUp(object sender, System.Windows.Forms.MouseEventArgs e)
		{
			if(e.Button!=MouseButtons.Right)return;
			Cursor = Cursors.WaitCursor;
			RecetaArticulo miRecetaArticulo = new RecetaArticulo();
			miRecetaArticulo.MdiParent = this.MdiParent;
			miRecetaArticulo.Show();
			Cursor = Cursors.Default;
																				
		}

		private void ultraGrid1_InitializeLayout(object sender, Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs e)
		{
		
		}

		private void Receta_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Escape) 
			{
				if (this.barraDato1.bEditar)
				{
					this.barraDato1_Deshacer(sender, e);
				}
				else
					Close();
			}
			if (e.KeyCode == Keys.N)
			{
				if (!this.barraDato1.bEditar)
				{
					this.barraDato1_Crear(sender, e);
				}
			}
			if (e.KeyCode == Keys.E)
			{
				if (!this.barraDato1.bEditar)
				{
					this.barraDato1_Editar(sender, e);
				}
			}
			if (e.KeyCode == Keys.C)
			{
				if (!this.barraDato1.bEditar)
				{
					this.barraDato1_Buscar(sender, e);
				}
			}
			if (e.KeyCode == Keys.I)
			{
				if (!this.barraDato1.bEditar)
				{
					this.barraDato1_Impresora(sender, e);
				}
			}
			if (e.KeyCode == Keys.Delete)
			{
				if (!this.barraDato1.bEditar)
				{
					this.barraDato1_Borrar(sender, e);
				}
			}
			if (e.KeyCode == Keys.F2)
			{
				if (this.barraDato1.bEditar)
				{
					this.barraDato1_Grabar(sender, e);
				}
			}
		}

		private void barraDato1_Load(object sender, System.EventArgs e)
		{
		
		}

		private void ultraGrid1_BeforeRowsDeleted(object sender, Infragistics.Win.UltraWinGrid.BeforeRowsDeletedEventArgs e)
		{
			if (!bEliminaAlValidar)
			{
				if (DialogResult.No ==	MessageBox.Show("¿Desea Borrar Líneas Seleccionadas?", "Point Technology", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2))
					e.Cancel = true;
				else 					
					e.DisplayPromptMsg = false;
			}
			else
				e.DisplayPromptMsg = false;
		}

		private void txtCodigo_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.F3)
			{
				using (BuscaArticulos miBusqueda = new BuscaArticulos(this.txtCodigo.Text.ToString(), 1, 0))
				{
					if (DialogResult.OK == miBusqueda.ShowDialog())
					{
						drReceta["idArticulo"] = (int) miBusqueda.idArticulo;
						this.txtCodigo.Text = miBusqueda.sCodigo;
						this.txtArticulo.Text = miBusqueda.sArticulo;			
					}
				}
			}
		}
	}
}
