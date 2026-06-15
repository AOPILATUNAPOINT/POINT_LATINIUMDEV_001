using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;

namespace Latinium
{
	/// <summary>
	/// Descripción breve de ArticuloArbol.
	/// </summary>
	public class ArticuloArbol : DevExpress.XtraEditors.XtraForm
	{
		private C1.Data.C1DataSet cdsArticulo;
		private C1.Data.C1DataView cdvArticulo;
		private DevExpress.XtraEditors.SimpleButton ultraButton1;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbArt;
		private DevExpress.XtraTreeList.TreeList tlsArticulo;
		private DevExpress.XtraTreeList.Columns.TreeListColumn colArticulo;
		private DevExpress.XtraTreeList.Columns.TreeListColumn colCodigo;
		/// <summary>
		/// Variable del diseñador requerida.
		/// </summary>
		private System.ComponentModel.Container components = null;

		public ArticuloArbol()
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
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand1 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Articulo", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn1 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idArticulo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn2 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Articulo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn3 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idArticuloSRI");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn4 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idColor");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn5 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idModelo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn6 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idGrupoArticulo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn7 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idSubGrupo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn8 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idArancel");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn9 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idMarca");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn10 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idUnidad");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn11 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idPromocion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn12 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idTipoGrupo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn13 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idClase");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn14 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idSeccion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn15 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idSucursal");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn16 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idPadre");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn17 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("FechaIngreso");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn18 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Codigo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn19 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("CodAlterno");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn20 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Saldo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn21 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Peso");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn22 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Caracteristica");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn23 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Aplicacion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn24 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Iva");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn25 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Ice");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn26 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Precio1");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn27 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Precio2");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn28 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Precio3");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn29 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Precio4");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn30 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Precio5");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn31 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Precio6");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn32 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("CantidadMinima");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn33 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("CantidadMaxima");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn34 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Residual");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn35 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("CostoUltimo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn36 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("FechaUltimo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn37 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCompra1");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn38 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCompra2");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn39 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCompra3");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn40 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("CostoPromedio");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn41 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Gramos");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn42 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("ConsumoProm");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn43 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("TiempoReposicion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn44 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("UnidadPeso");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn45 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Notas");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn46 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Depreciacion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn47 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Incr1");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn48 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Incr2");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn49 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Incr3");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn50 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Incr4");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn51 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Incr5");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn52 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Incr6");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn53 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Ubicacion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn54 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("NumeroParte");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn55 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("ComisionValor");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn56 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("ComisionPorcentaje");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn57 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("CantidadRequerida");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn58 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("PrecioPromocion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn59 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("CantidadPromocion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn60 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("DescuentoContado");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn61 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("DescuentoCredito");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn62 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("DescPorcContado");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn63 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("DescPorcCre");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn64 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("CantReqDesc");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn65 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("CantReqPorc");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn66 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("TipoPromocion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn67 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Responsable");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn68 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Comision30");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn69 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Comision60");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn70 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Comision90");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn71 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Comision120");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn72 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Bodega1");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn73 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Bodega2");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn74 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Bodega3");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn75 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Bodega4");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn76 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Bodega5");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn77 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Bodega6");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn78 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Bodega7");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn79 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Bodega8");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn80 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Bodega9");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn81 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("SaldoT");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn82 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Alto");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn83 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Ancho");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn84 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Largo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn85 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Color");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn86 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Descontinuado");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn87 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Borrar");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn88 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Habilitado");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn89 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("ValorPeso");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn90 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("BodegaPredef");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn91 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Tiempo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn92 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Cantidad1");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn93 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Cantidad2");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn94 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Cantidad3");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn95 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Cantidad4");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn96 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Cantidad5");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn97 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("PaqueteUnidad");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn98 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Editar");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn99 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nuevo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn100 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("FechaEditado");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn101 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("CambiarPrecio");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn102 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("dCompra");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn103 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("dVenta");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn104 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Foto");
			this.tlsArticulo = new DevExpress.XtraTreeList.TreeList();
			this.colArticulo = new DevExpress.XtraTreeList.Columns.TreeListColumn();
			this.colCodigo = new DevExpress.XtraTreeList.Columns.TreeListColumn();
			this.cdvArticulo = new C1.Data.C1DataView();
			this.cdsArticulo = new C1.Data.C1DataSet();
			this.ultraButton1 = new DevExpress.XtraEditors.SimpleButton();
			this.cmbArt = new Infragistics.Win.UltraWinGrid.UltraCombo();
			((System.ComponentModel.ISupportInitialize)(this.tlsArticulo)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdvArticulo)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsArticulo)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbArt)).BeginInit();
			this.SuspendLayout();
			// 
			// tlsArticulo
			// 
			this.tlsArticulo.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.tlsArticulo.Columns.AddRange(new DevExpress.XtraTreeList.Columns.TreeListColumn[] {
																																															 this.colArticulo,
																																															 this.colCodigo});
			this.tlsArticulo.DataSource = this.cdvArticulo;
			this.tlsArticulo.KeyFieldName = "idArticulo";
			this.tlsArticulo.Location = new System.Drawing.Point(16, 56);
			this.tlsArticulo.Name = "tlsArticulo";
			this.tlsArticulo.ParentFieldName = "idPadre";
			this.tlsArticulo.RootValue = "0";
			this.tlsArticulo.Size = new System.Drawing.Size(528, 304);
			this.tlsArticulo.Styles.AddReplace("Preview", new DevExpress.Utils.ViewStyle("Preview", "TreeList", new System.Drawing.Font("Tahoma", 8F), "", DevExpress.Utils.StyleOptions.StyleEnabled, true, true, false, DevExpress.Utils.HorzAlignment.Near, DevExpress.Utils.VertAlignment.Center, null, System.Drawing.Color.FromArgb(((System.Byte)(255)), ((System.Byte)(255)), ((System.Byte)(255))), System.Drawing.Color.FromArgb(((System.Byte)(0)), ((System.Byte)(0)), ((System.Byte)(255)))));
			this.tlsArticulo.Styles.AddReplace("FooterPanel", new DevExpress.Utils.ViewStyle("FooterPanel", "TreeList", new System.Drawing.Font("Tahoma", 8F), "", DevExpress.Utils.StyleOptions.StyleEnabled, true, false, false, DevExpress.Utils.HorzAlignment.Far, DevExpress.Utils.VertAlignment.Center, null, System.Drawing.Color.FromArgb(((System.Byte)(0)), ((System.Byte)(206)), ((System.Byte)(209))), System.Drawing.Color.FromArgb(((System.Byte)(0)), ((System.Byte)(0)), ((System.Byte)(0)))));
			this.tlsArticulo.Styles.AddReplace("GroupButton", new DevExpress.Utils.ViewStyle("GroupButton", "TreeList", new System.Drawing.Font("Tahoma", 8F), "", DevExpress.Utils.StyleOptions.StyleEnabled, true, false, false, DevExpress.Utils.HorzAlignment.Default, DevExpress.Utils.VertAlignment.Center, null, System.Drawing.Color.FromArgb(((System.Byte)(255)), ((System.Byte)(255)), ((System.Byte)(255))), System.Drawing.Color.FromArgb(((System.Byte)(0)), ((System.Byte)(0)), ((System.Byte)(0)))));
			this.tlsArticulo.Styles.AddReplace("EvenRow", new DevExpress.Utils.ViewStyle("EvenRow", "TreeList", new System.Drawing.Font("Tahoma", 8F), "", DevExpress.Utils.StyleOptions.StyleEnabled, true, false, false, DevExpress.Utils.HorzAlignment.Default, DevExpress.Utils.VertAlignment.Center, null, System.Drawing.Color.FromArgb(((System.Byte)(248)), ((System.Byte)(248)), ((System.Byte)(255))), System.Drawing.Color.FromArgb(((System.Byte)(0)), ((System.Byte)(0)), ((System.Byte)(0)))));
			this.tlsArticulo.Styles.AddReplace("TreeLine", new DevExpress.Utils.ViewStyle("TreeLine", "TreeList", new System.Drawing.Font("Microsoft Sans Serif", 8F), "", DevExpress.Utils.StyleOptions.StyleEnabled, true, false, false, DevExpress.Utils.HorzAlignment.Default, DevExpress.Utils.VertAlignment.Center, null, System.Drawing.Color.FromArgb(((System.Byte)(255)), ((System.Byte)(255)), ((System.Byte)(255))), System.Drawing.Color.FromArgb(((System.Byte)(128)), ((System.Byte)(128)), ((System.Byte)(128)))));
			this.tlsArticulo.Styles.AddReplace("PressedColumn", new DevExpress.Utils.ViewStyle("PressedColumn", "TreeList", new System.Drawing.Font("Arial", 8F, System.Drawing.FontStyle.Bold), "HeaderPanel", ((DevExpress.Utils.StyleOptions)(((DevExpress.Utils.StyleOptions.StyleEnabled | DevExpress.Utils.StyleOptions.UseBackColor) 
				| DevExpress.Utils.StyleOptions.UseForeColor))), true, false, false, DevExpress.Utils.HorzAlignment.Near, DevExpress.Utils.VertAlignment.Center, null, System.Drawing.Color.FromArgb(((System.Byte)(128)), ((System.Byte)(128)), ((System.Byte)(128))), System.Drawing.Color.FromArgb(((System.Byte)(255)), ((System.Byte)(255)), ((System.Byte)(255)))));
			this.tlsArticulo.Styles.AddReplace("HideSelectionRow", new DevExpress.Utils.ViewStyle("HideSelectionRow", "TreeList", new System.Drawing.Font("Tahoma", 8F), "", ((DevExpress.Utils.StyleOptions)((((((DevExpress.Utils.StyleOptions.StyleEnabled | DevExpress.Utils.StyleOptions.UseBackColor) 
				| DevExpress.Utils.StyleOptions.UseDrawFocusRect) 
				| DevExpress.Utils.StyleOptions.UseFont) 
				| DevExpress.Utils.StyleOptions.UseForeColor) 
				| DevExpress.Utils.StyleOptions.UseImage))), true, false, false, DevExpress.Utils.HorzAlignment.Default, DevExpress.Utils.VertAlignment.Center, null, System.Drawing.Color.FromArgb(((System.Byte)(128)), ((System.Byte)(128)), ((System.Byte)(128))), System.Drawing.Color.FromArgb(((System.Byte)(212)), ((System.Byte)(208)), ((System.Byte)(200)))));
			this.tlsArticulo.Styles.AddReplace("GroupFooter", new DevExpress.Utils.ViewStyle("GroupFooter", "TreeList", new System.Drawing.Font("Tahoma", 8F), "", DevExpress.Utils.StyleOptions.StyleEnabled, true, false, false, DevExpress.Utils.HorzAlignment.Far, DevExpress.Utils.VertAlignment.Center, null, System.Drawing.Color.FromArgb(((System.Byte)(10)), ((System.Byte)(216)), ((System.Byte)(219))), System.Drawing.Color.FromArgb(((System.Byte)(0)), ((System.Byte)(0)), ((System.Byte)(0)))));
			this.tlsArticulo.Styles.AddReplace("HorzLine", new DevExpress.Utils.ViewStyle("HorzLine", "TreeList", new System.Drawing.Font("Tahoma", 8F), "", DevExpress.Utils.StyleOptions.StyleEnabled, true, false, false, DevExpress.Utils.HorzAlignment.Default, DevExpress.Utils.VertAlignment.Center, null, System.Drawing.Color.FromArgb(((System.Byte)(128)), ((System.Byte)(128)), ((System.Byte)(128))), System.Drawing.Color.FromArgb(((System.Byte)(128)), ((System.Byte)(128)), ((System.Byte)(128)))));
			this.tlsArticulo.Styles.AddReplace("HeaderPanel", new DevExpress.Utils.ViewStyle("HeaderPanel", "TreeList", new System.Drawing.Font("Arial", 8F, System.Drawing.FontStyle.Bold), "", DevExpress.Utils.StyleOptions.StyleEnabled, true, false, false, DevExpress.Utils.HorzAlignment.Near, DevExpress.Utils.VertAlignment.Center, null, System.Drawing.Color.FromArgb(((System.Byte)(0)), ((System.Byte)(206)), ((System.Byte)(209))), System.Drawing.Color.FromArgb(((System.Byte)(0)), ((System.Byte)(0)), ((System.Byte)(0)))));
			this.tlsArticulo.Styles.AddReplace("Empty", new DevExpress.Utils.ViewStyle("Empty", "TreeList", new System.Drawing.Font("Microsoft Sans Serif", 8F), "", DevExpress.Utils.StyleOptions.StyleEnabled, true, false, false, DevExpress.Utils.HorzAlignment.Default, DevExpress.Utils.VertAlignment.Center, null, System.Drawing.Color.FromArgb(((System.Byte)(255)), ((System.Byte)(255)), ((System.Byte)(255))), System.Drawing.Color.FromArgb(((System.Byte)(255)), ((System.Byte)(255)), ((System.Byte)(255)))));
			this.tlsArticulo.Styles.AddReplace("VertLine", new DevExpress.Utils.ViewStyle("VertLine", "TreeList", new System.Drawing.Font("Tahoma", 8F), "", DevExpress.Utils.StyleOptions.StyleEnabled, true, false, false, DevExpress.Utils.HorzAlignment.Near, DevExpress.Utils.VertAlignment.Center, null, System.Drawing.Color.FromArgb(((System.Byte)(128)), ((System.Byte)(128)), ((System.Byte)(128))), System.Drawing.Color.FromArgb(((System.Byte)(128)), ((System.Byte)(128)), ((System.Byte)(128)))));
			this.tlsArticulo.Styles.AddReplace("FocusedCell", new DevExpress.Utils.ViewStyle("FocusedCell", "TreeList", new System.Drawing.Font("Tahoma", 8F), "", ((DevExpress.Utils.StyleOptions)((((((DevExpress.Utils.StyleOptions.StyleEnabled | DevExpress.Utils.StyleOptions.UseBackColor) 
				| DevExpress.Utils.StyleOptions.UseDrawFocusRect) 
				| DevExpress.Utils.StyleOptions.UseFont) 
				| DevExpress.Utils.StyleOptions.UseForeColor) 
				| DevExpress.Utils.StyleOptions.UseImage))), true, false, false, DevExpress.Utils.HorzAlignment.Default, DevExpress.Utils.VertAlignment.Center, null, System.Drawing.Color.FromArgb(((System.Byte)(255)), ((System.Byte)(228)), ((System.Byte)(225))), System.Drawing.Color.FromArgb(((System.Byte)(0)), ((System.Byte)(0)), ((System.Byte)(0)))));
			this.tlsArticulo.Styles.AddReplace("OddRow", new DevExpress.Utils.ViewStyle("OddRow", "TreeList", new System.Drawing.Font("Arial Black", 7F), "", DevExpress.Utils.StyleOptions.StyleEnabled, true, false, false, DevExpress.Utils.HorzAlignment.Default, DevExpress.Utils.VertAlignment.Center, null, System.Drawing.Color.FromArgb(((System.Byte)(230)), ((System.Byte)(230)), ((System.Byte)(250))), System.Drawing.Color.FromArgb(((System.Byte)(0)), ((System.Byte)(0)), ((System.Byte)(0)))));
			this.tlsArticulo.Styles.AddReplace("SelectedRow", new DevExpress.Utils.ViewStyle("SelectedRow", "TreeList", new System.Drawing.Font("Tahoma", 8F), "", ((DevExpress.Utils.StyleOptions)((((((DevExpress.Utils.StyleOptions.StyleEnabled | DevExpress.Utils.StyleOptions.UseBackColor) 
				| DevExpress.Utils.StyleOptions.UseDrawFocusRect) 
				| DevExpress.Utils.StyleOptions.UseFont) 
				| DevExpress.Utils.StyleOptions.UseForeColor) 
				| DevExpress.Utils.StyleOptions.UseImage))), true, false, false, DevExpress.Utils.HorzAlignment.Default, DevExpress.Utils.VertAlignment.Center, null, System.Drawing.Color.FromArgb(((System.Byte)(255)), ((System.Byte)(10)), ((System.Byte)(10))), System.Drawing.Color.FromArgb(((System.Byte)(255)), ((System.Byte)(255)), ((System.Byte)(255)))));
			this.tlsArticulo.Styles.AddReplace("FocusedRow", new DevExpress.Utils.ViewStyle("FocusedRow", "TreeList", new System.Drawing.Font("Tahoma", 8F), "", ((DevExpress.Utils.StyleOptions)((((((DevExpress.Utils.StyleOptions.StyleEnabled | DevExpress.Utils.StyleOptions.UseBackColor) 
				| DevExpress.Utils.StyleOptions.UseDrawFocusRect) 
				| DevExpress.Utils.StyleOptions.UseFont) 
				| DevExpress.Utils.StyleOptions.UseForeColor) 
				| DevExpress.Utils.StyleOptions.UseImage))), true, false, false, DevExpress.Utils.HorzAlignment.Default, DevExpress.Utils.VertAlignment.Center, null, System.Drawing.Color.FromArgb(((System.Byte)(255)), ((System.Byte)(0)), ((System.Byte)(0))), System.Drawing.Color.FromArgb(((System.Byte)(255)), ((System.Byte)(255)), ((System.Byte)(255)))));
			this.tlsArticulo.Styles.AddReplace("Row", new DevExpress.Utils.ViewStyle("Row", "TreeList", new System.Drawing.Font("Tahoma", 8F), "", DevExpress.Utils.StyleOptions.StyleEnabled, true, false, false, DevExpress.Utils.HorzAlignment.Default, DevExpress.Utils.VertAlignment.Center, null, System.Drawing.Color.FromArgb(((System.Byte)(255)), ((System.Byte)(255)), ((System.Byte)(255))), System.Drawing.Color.FromArgb(((System.Byte)(0)), ((System.Byte)(0)), ((System.Byte)(0)))));
			this.tlsArticulo.TabIndex = 0;
			this.tlsArticulo.FocusedNodeChanged += new DevExpress.XtraTreeList.FocusedNodeChangedEventHandler(this.tlsArticulo_FocusedNodeChanged);
			// 
			// colArticulo
			// 
			this.colArticulo.Caption = "Articulo";
			this.colArticulo.FieldName = "Articulo";
			this.colArticulo.Name = "colArticulo";
			this.colArticulo.VisibleIndex = 0;
			this.colArticulo.Width = 303;
			// 
			// colCodigo
			// 
			this.colCodigo.Caption = "Codigo";
			this.colCodigo.FieldName = "Codigo";
			this.colCodigo.Name = "colCodigo";
			this.colCodigo.VisibleIndex = 1;
			this.colCodigo.Width = 160;
			// 
			// cdvArticulo
			// 
			this.cdvArticulo.BindingContextCtrl = this;
			this.cdvArticulo.DataSet = this.cdsArticulo;
			this.cdvArticulo.Sort = "Articulo";
			this.cdvArticulo.TableName = "";
			this.cdvArticulo.TableViewName = "Articulo";
			// 
			// cdsArticulo
			// 
			this.cdsArticulo.BindingContextCtrl = this;
			this.cdsArticulo.DataLibrary = "LibFacturacion";
			this.cdsArticulo.DataLibraryUrl = "";
			this.cdsArticulo.DataSetDef = "dsArticulo";
			this.cdsArticulo.Locale = new System.Globalization.CultureInfo("es-EC");
			this.cdsArticulo.Name = "cdsArticulo";
			this.cdsArticulo.SchemaClassName = "LibFacturacion.DataClass";
			this.cdsArticulo.SchemaDef = null;
			// 
			// ultraButton1
			// 
			this.ultraButton1.Location = new System.Drawing.Point(208, 16);
			this.ultraButton1.Name = "ultraButton1";
			this.ultraButton1.Size = new System.Drawing.Size(56, 24);
			this.ultraButton1.TabIndex = 1;
			this.ultraButton1.Text = "Buscar";
			this.ultraButton1.Click += new System.EventHandler(this.ultraButton1_Click);
			// 
			// cmbArt
			// 
			appearance1.ForeColorDisabled = System.Drawing.Color.Black;
			this.cmbArt.Appearance = appearance1;
			this.cmbArt.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbArt.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbArt.DataSource = this.cdvArticulo;
			ultraGridColumn1.Header.VisiblePosition = 44;
			ultraGridColumn1.Hidden = true;
			ultraGridColumn2.Header.VisiblePosition = 0;
			ultraGridColumn2.Width = 260;
			ultraGridColumn3.Header.VisiblePosition = 100;
			ultraGridColumn3.Hidden = true;
			ultraGridColumn4.Header.VisiblePosition = 93;
			ultraGridColumn4.Hidden = true;
			ultraGridColumn5.Header.VisiblePosition = 94;
			ultraGridColumn5.Hidden = true;
			ultraGridColumn6.Header.VisiblePosition = 46;
			ultraGridColumn6.Hidden = true;
			ultraGridColumn7.Header.VisiblePosition = 50;
			ultraGridColumn7.Hidden = true;
			ultraGridColumn8.Header.VisiblePosition = 43;
			ultraGridColumn8.Hidden = true;
			ultraGridColumn9.Header.VisiblePosition = 47;
			ultraGridColumn9.Hidden = true;
			ultraGridColumn10.Header.VisiblePosition = 52;
			ultraGridColumn10.Hidden = true;
			ultraGridColumn11.Header.VisiblePosition = 48;
			ultraGridColumn11.Hidden = true;
			ultraGridColumn12.Header.VisiblePosition = 51;
			ultraGridColumn12.Hidden = true;
			ultraGridColumn13.Header.VisiblePosition = 45;
			ultraGridColumn13.Hidden = true;
			ultraGridColumn14.Header.VisiblePosition = 49;
			ultraGridColumn14.Hidden = true;
			ultraGridColumn15.Header.VisiblePosition = 91;
			ultraGridColumn15.Hidden = true;
			ultraGridColumn16.Header.VisiblePosition = 99;
			ultraGridColumn16.Hidden = true;
			ultraGridColumn17.Header.VisiblePosition = 84;
			ultraGridColumn17.Hidden = true;
			ultraGridColumn18.Header.VisiblePosition = 1;
			ultraGridColumn19.Header.VisiblePosition = 23;
			ultraGridColumn19.Hidden = true;
			ultraGridColumn20.Header.VisiblePosition = 72;
			ultraGridColumn20.Hidden = true;
			ultraGridColumn21.Header.VisiblePosition = 62;
			ultraGridColumn21.Hidden = true;
			ultraGridColumn22.Header.VisiblePosition = 22;
			ultraGridColumn22.Hidden = true;
			ultraGridColumn23.Header.VisiblePosition = 4;
			ultraGridColumn23.Hidden = true;
			ultraGridColumn24.Header.VisiblePosition = 59;
			ultraGridColumn24.Hidden = true;
			ultraGridColumn25.Header.VisiblePosition = 42;
			ultraGridColumn25.Hidden = true;
			ultraGridColumn26.Header.VisiblePosition = 63;
			ultraGridColumn26.Hidden = true;
			ultraGridColumn27.Header.VisiblePosition = 64;
			ultraGridColumn27.Hidden = true;
			ultraGridColumn28.Header.VisiblePosition = 65;
			ultraGridColumn28.Hidden = true;
			ultraGridColumn29.Header.VisiblePosition = 66;
			ultraGridColumn29.Hidden = true;
			ultraGridColumn30.Header.VisiblePosition = 67;
			ultraGridColumn30.Hidden = true;
			ultraGridColumn31.Header.VisiblePosition = 68;
			ultraGridColumn31.Hidden = true;
			ultraGridColumn32.Header.VisiblePosition = 17;
			ultraGridColumn32.Hidden = true;
			ultraGridColumn33.Header.VisiblePosition = 16;
			ultraGridColumn33.Hidden = true;
			ultraGridColumn34.Header.VisiblePosition = 70;
			ultraGridColumn34.Hidden = true;
			ultraGridColumn35.Header.VisiblePosition = 33;
			ultraGridColumn35.Hidden = true;
			ultraGridColumn36.Header.VisiblePosition = 83;
			ultraGridColumn36.Hidden = true;
			ultraGridColumn37.Header.VisiblePosition = 80;
			ultraGridColumn37.Hidden = true;
			ultraGridColumn38.Header.VisiblePosition = 81;
			ultraGridColumn38.Hidden = true;
			ultraGridColumn39.Header.VisiblePosition = 82;
			ultraGridColumn39.Hidden = true;
			ultraGridColumn40.Header.VisiblePosition = 32;
			ultraGridColumn40.Hidden = true;
			ultraGridColumn41.Header.VisiblePosition = 40;
			ultraGridColumn41.Hidden = true;
			ultraGridColumn42.Header.VisiblePosition = 31;
			ultraGridColumn42.Hidden = true;
			ultraGridColumn43.Header.VisiblePosition = 75;
			ultraGridColumn43.Hidden = true;
			ultraGridColumn44.Header.VisiblePosition = 78;
			ultraGridColumn44.Hidden = true;
			ultraGridColumn45.Header.VisiblePosition = 92;
			ultraGridColumn45.Hidden = true;
			ultraGridColumn46.Header.VisiblePosition = 34;
			ultraGridColumn46.Hidden = true;
			ultraGridColumn47.Header.VisiblePosition = 53;
			ultraGridColumn47.Hidden = true;
			ultraGridColumn48.Header.VisiblePosition = 54;
			ultraGridColumn48.Hidden = true;
			ultraGridColumn49.Header.VisiblePosition = 55;
			ultraGridColumn49.Hidden = true;
			ultraGridColumn50.Header.VisiblePosition = 56;
			ultraGridColumn50.Hidden = true;
			ultraGridColumn51.Header.VisiblePosition = 57;
			ultraGridColumn51.Hidden = true;
			ultraGridColumn52.Header.VisiblePosition = 58;
			ultraGridColumn52.Hidden = true;
			ultraGridColumn53.Header.VisiblePosition = 77;
			ultraGridColumn53.Hidden = true;
			ultraGridColumn54.Header.VisiblePosition = 61;
			ultraGridColumn54.Hidden = true;
			ultraGridColumn55.Header.VisiblePosition = 30;
			ultraGridColumn55.Hidden = true;
			ultraGridColumn56.Header.VisiblePosition = 29;
			ultraGridColumn56.Hidden = true;
			ultraGridColumn57.Header.VisiblePosition = 19;
			ultraGridColumn57.Hidden = true;
			ultraGridColumn58.Header.VisiblePosition = 69;
			ultraGridColumn58.Hidden = true;
			ultraGridColumn59.Header.VisiblePosition = 18;
			ultraGridColumn59.Hidden = true;
			ultraGridColumn60.Header.VisiblePosition = 38;
			ultraGridColumn60.Hidden = true;
			ultraGridColumn61.Header.VisiblePosition = 39;
			ultraGridColumn61.Hidden = true;
			ultraGridColumn62.Header.VisiblePosition = 36;
			ultraGridColumn62.Hidden = true;
			ultraGridColumn63.Header.VisiblePosition = 37;
			ultraGridColumn63.Hidden = true;
			ultraGridColumn64.Header.VisiblePosition = 20;
			ultraGridColumn64.Hidden = true;
			ultraGridColumn65.Header.VisiblePosition = 21;
			ultraGridColumn65.Hidden = true;
			ultraGridColumn66.Header.VisiblePosition = 76;
			ultraGridColumn66.Hidden = true;
			ultraGridColumn67.Header.VisiblePosition = 71;
			ultraGridColumn67.Hidden = true;
			ultraGridColumn68.Header.VisiblePosition = 26;
			ultraGridColumn68.Hidden = true;
			ultraGridColumn69.Header.VisiblePosition = 27;
			ultraGridColumn69.Hidden = true;
			ultraGridColumn70.Header.VisiblePosition = 28;
			ultraGridColumn70.Hidden = true;
			ultraGridColumn71.Header.VisiblePosition = 25;
			ultraGridColumn71.Hidden = true;
			ultraGridColumn72.Header.VisiblePosition = 5;
			ultraGridColumn72.Hidden = true;
			ultraGridColumn73.Header.VisiblePosition = 6;
			ultraGridColumn73.Hidden = true;
			ultraGridColumn74.Header.VisiblePosition = 7;
			ultraGridColumn74.Hidden = true;
			ultraGridColumn75.Header.VisiblePosition = 8;
			ultraGridColumn75.Hidden = true;
			ultraGridColumn76.Header.VisiblePosition = 9;
			ultraGridColumn76.Hidden = true;
			ultraGridColumn77.Header.VisiblePosition = 10;
			ultraGridColumn77.Hidden = true;
			ultraGridColumn78.Header.VisiblePosition = 11;
			ultraGridColumn78.Hidden = true;
			ultraGridColumn79.Header.VisiblePosition = 12;
			ultraGridColumn79.Hidden = true;
			ultraGridColumn80.Header.VisiblePosition = 13;
			ultraGridColumn80.Hidden = true;
			ultraGridColumn81.Header.VisiblePosition = 73;
			ultraGridColumn81.Hidden = true;
			ultraGridColumn82.Header.VisiblePosition = 2;
			ultraGridColumn82.Hidden = true;
			ultraGridColumn83.Header.VisiblePosition = 3;
			ultraGridColumn83.Hidden = true;
			ultraGridColumn84.Header.VisiblePosition = 60;
			ultraGridColumn84.Hidden = true;
			ultraGridColumn85.Header.VisiblePosition = 24;
			ultraGridColumn85.Hidden = true;
			ultraGridColumn86.Header.VisiblePosition = 35;
			ultraGridColumn86.Hidden = true;
			ultraGridColumn87.Header.VisiblePosition = 15;
			ultraGridColumn87.Hidden = true;
			ultraGridColumn88.Header.VisiblePosition = 41;
			ultraGridColumn88.Hidden = true;
			ultraGridColumn89.Header.VisiblePosition = 79;
			ultraGridColumn89.Hidden = true;
			ultraGridColumn90.Header.VisiblePosition = 14;
			ultraGridColumn90.Hidden = true;
			ultraGridColumn91.Header.VisiblePosition = 74;
			ultraGridColumn91.Hidden = true;
			ultraGridColumn92.Header.VisiblePosition = 85;
			ultraGridColumn92.Hidden = true;
			ultraGridColumn93.Header.VisiblePosition = 86;
			ultraGridColumn93.Hidden = true;
			ultraGridColumn94.Header.VisiblePosition = 87;
			ultraGridColumn94.Hidden = true;
			ultraGridColumn95.Header.VisiblePosition = 88;
			ultraGridColumn95.Hidden = true;
			ultraGridColumn96.Header.VisiblePosition = 89;
			ultraGridColumn96.Hidden = true;
			ultraGridColumn97.Header.VisiblePosition = 90;
			ultraGridColumn97.Hidden = true;
			ultraGridColumn98.Header.VisiblePosition = 95;
			ultraGridColumn98.Hidden = true;
			ultraGridColumn99.Header.VisiblePosition = 96;
			ultraGridColumn99.Hidden = true;
			ultraGridColumn100.Header.VisiblePosition = 97;
			ultraGridColumn100.Hidden = true;
			ultraGridColumn101.Header.VisiblePosition = 98;
			ultraGridColumn101.Hidden = true;
			ultraGridColumn102.Header.VisiblePosition = 101;
			ultraGridColumn102.Hidden = true;
			ultraGridColumn103.Header.VisiblePosition = 102;
			ultraGridColumn103.Hidden = true;
			ultraGridColumn104.Header.VisiblePosition = 103;
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
																										 ultraGridColumn49,
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
																										 ultraGridColumn82,
																										 ultraGridColumn83,
																										 ultraGridColumn84,
																										 ultraGridColumn85,
																										 ultraGridColumn86,
																										 ultraGridColumn87,
																										 ultraGridColumn88,
																										 ultraGridColumn89,
																										 ultraGridColumn90,
																										 ultraGridColumn91,
																										 ultraGridColumn92,
																										 ultraGridColumn93,
																										 ultraGridColumn94,
																										 ultraGridColumn95,
																										 ultraGridColumn96,
																										 ultraGridColumn97,
																										 ultraGridColumn98,
																										 ultraGridColumn99,
																										 ultraGridColumn100,
																										 ultraGridColumn101,
																										 ultraGridColumn102,
																										 ultraGridColumn103,
																										 ultraGridColumn104});
			this.cmbArt.DisplayLayout.BandsSerializer.Add(ultraGridBand1);
			this.cmbArt.DisplayMember = "Articulo";
			this.cmbArt.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbArt.Location = new System.Drawing.Point(16, 16);
			this.cmbArt.Name = "cmbArt";
			this.cmbArt.Size = new System.Drawing.Size(176, 21);
			this.cmbArt.TabIndex = 2;
			this.cmbArt.ValueMember = "idArticulo";
			this.cmbArt.MouseUp += new System.Windows.Forms.MouseEventHandler(this.cmbArt_MouseUp);
			// 
			// ArticuloArbol
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 13);
			this.ClientSize = new System.Drawing.Size(560, 414);
			this.Controls.Add(this.cmbArt);
			this.Controls.Add(this.ultraButton1);
			this.Controls.Add(this.tlsArticulo);
			this.Name = "ArticuloArbol";
			this.Text = "Articulos en Arbol";
			this.Closing += new System.ComponentModel.CancelEventHandler(this.ArticuloArbol_Closing);
			this.Load += new System.EventHandler(this.ArticuloArbol_Load);
			((System.ComponentModel.ISupportInitialize)(this.tlsArticulo)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdvArticulo)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsArticulo)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbArt)).EndInit();
			this.ResumeLayout(false);

		}
		#endregion

		private void ultraButton1_Click(object sender, System.EventArgs e)
		{
			if (this.cmbArt.Value == null)
			{
				this.cmbArt.Select();
				MessageBox.Show("Escoja Articulo a buscar", "Información");
				return;
			}
			int IdArt = (int) this.cmbArt.Value;
			DevExpress.XtraTreeList.Nodes.TreeListNode nd = this.tlsArticulo.FindNodeByKeyID(IdArt);
			this.tlsArticulo.FocusedNode = nd;

		}

		private void cmbArt_MouseUp(object sender, System.Windows.Forms.MouseEventArgs e)
		{
			if (e.Button == MouseButtons.Right)
			{
				Cursor = Cursors.WaitCursor;
				string stPrimero = "Articulo";
				if (cmbArt.DisplayLayout.Bands[0].Columns[stPrimero].Header.VisiblePosition == 0)
					stPrimero = "Codigo";
				cdvArticulo.Sort = stPrimero;
				cmbArt.DisplayLayout.Bands[0].Columns[stPrimero].Header.VisiblePosition = 0;
				cmbArt.DisplayMember = stPrimero;
				Cursor = Cursors.Default;
			}
		}

		private void tlsArticulo_FocusedNodeChanged(object sender, DevExpress.XtraTreeList.FocusedNodeChangedEventArgs e)
		{
		
		}

		private void ArticuloArbol_Closing(object sender, System.ComponentModel.CancelEventArgs e)
		{
			string stAudita = "Exec AuditaCrear 53, 7, 'Art. Arbol'";
			Funcion.EjecutaSQL(cdsArticulo, stAudita, true);
		}

		private void ArticuloArbol_Load(object sender, System.EventArgs e)
		{
			string stAudita = "Exec AuditaCrear 53, 6, 'Art. Arbol'";
			Funcion.EjecutaSQL(cdsArticulo, stAudita, true);
		}
	}
}
