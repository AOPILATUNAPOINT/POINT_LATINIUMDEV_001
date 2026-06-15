using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using Infragistics.Win.UltraWinGrid;
namespace Latinium
{
	/// <summary>
	/// Descripción breve de MarceloReparteC.
	/// </summary>
	public class MarceloReparteC : DevExpress.XtraEditors.XtraForm
	{
		private System.Windows.Forms.Label label3;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtDescripcion;
		private BarraDatoSQL.BarraDatoSQL barraDato1;
		private Infragistics.Win.UltraWinGrid.UltraGrid grdTalla;
		private C1.Data.C1DataSet cdsMarceloCodigoD;
		private System.Windows.Forms.Label label1;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtCodigo;
		private C1.Data.C1DataSet cdsMAlmacen;
		private C1.Data.C1DataView cdvMAlmacen;
		private Infragistics.Win.UltraWinGrid.UltraDropDown cmbAlmacen;
		private Infragistics.Win.Misc.UltraButton btTotal;
		int seexcede=0;
		private System.Windows.Forms.Label label2;
		private System.Windows.Forms.Label lblStock;
		private System.Windows.Forms.Label lbl1;
		/// <summary>
		/// Variable del diseñador requerida.
		/// </summary>
		private System.ComponentModel.Container components = null;

		public MarceloReparteC()
		{
			//
			// Necesario para admitir el Diseñador de Windows Forms
			//
			InitializeComponent();

			//
			// TODO: agregar código de constructor después de llamar a InitializeComponent
			//
		}
		int idMarceloCod;
		public MarceloReparteC(int iDCod)
		{
			InitializeComponent();
	    idMarceloCod=iDCod ;
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
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand1 = new Infragistics.Win.UltraWinGrid.UltraGridBand("_MarceloDetCodigo", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn1 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idMarceloDetCodigo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn2 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Codigo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn3 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idMarceloAlmacen", -1, "cmbAlmacen");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn4 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idMarceloCodigo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn5 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Talla1");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn6 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Talla2");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn7 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Talla3");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn8 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Talla4");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn9 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Talla5");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn10 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Talla6");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn11 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Talla7");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn12 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Talla8");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn13 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Talla9");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn14 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Talla10");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn15 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Talla11");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn16 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Talla12");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn17 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Talla13");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn18 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Talla14");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn19 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Talla15");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn20 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Talla16");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn21 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Talla17");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn22 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Talla18");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn23 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Talla19");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn24 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Talla20");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn25 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Talla21");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn26 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Talla22");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn27 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Talla23");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn28 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Talla24");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn29 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Talla25");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn30 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("TotalF");
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings1 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "Talla1", 4, true, "_MarceloDetCodigo", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "Talla1", 4, true);
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings2 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "Talla2", 5, true, "_MarceloDetCodigo", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "Talla2", 5, true);
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings3 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "Talla3", 6, true, "_MarceloDetCodigo", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "Talla3", 6, true);
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings4 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "Talla4", 7, true, "_MarceloDetCodigo", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "Talla4", 7, true);
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings5 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "Talla5", 8, true, "_MarceloDetCodigo", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "Talla5", 8, true);
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings6 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "Talla6", 9, true, "_MarceloDetCodigo", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "Talla6", 9, true);
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings7 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "Talla7", 10, true, "_MarceloDetCodigo", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "Talla7", 10, true);
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings8 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "Talla8", 11, true, "_MarceloDetCodigo", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "Talla8", 11, true);
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings9 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "Talla9", 12, true, "_MarceloDetCodigo", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "Talla9", 12, true);
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings10 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "Talla10", 13, true, "_MarceloDetCodigo", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "Talla10", 13, true);
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings11 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "Talla11", 14, true, "_MarceloDetCodigo", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "Talla11", 14, true);
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings12 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "Talla12", 15, true, "_MarceloDetCodigo", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "Talla12", 15, true);
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings13 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "Talla13", 16, true, "_MarceloDetCodigo", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "Talla13", 16, true);
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings14 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "Talla14", 17, true, "_MarceloDetCodigo", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "Talla14", 17, true);
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings15 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "Talla15", 18, true, "_MarceloDetCodigo", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "Talla15", 18, true);
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings16 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "Talla16", 19, true, "_MarceloDetCodigo", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "Talla16", 19, true);
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings17 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "Talla17", 20, true, "_MarceloDetCodigo", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "Talla17", 20, true);
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings18 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "Talla18", 21, true, "_MarceloDetCodigo", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "Talla18", 21, true);
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings19 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "Talla19", 22, true, "_MarceloDetCodigo", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "Talla19", 22, true);
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings20 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "Talla20", 23, true, "_MarceloDetCodigo", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "Talla20", 23, true);
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings21 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "Talla21", 24, true, "_MarceloDetCodigo", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "Talla21", 24, true);
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings22 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "Talla22", 25, true, "_MarceloDetCodigo", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "Talla22", 25, true);
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings23 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "Talla23", 26, true, "_MarceloDetCodigo", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "Talla23", 26, true);
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings24 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "Talla24", 27, true, "_MarceloDetCodigo", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "Talla24", 27, true);
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings25 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "Talla25", 28, true, "_MarceloDetCodigo", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "Talla25", 28, true);
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings26 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "TotalF", 29, true, "_MarceloDetCodigo", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, "TotalF", 29, true);
			Infragistics.Win.Appearance appearance3 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance4 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance5 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance6 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance7 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand2 = new Infragistics.Win.UltraWinGrid.UltraGridBand("MarceloAlmacen", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn31 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idMarceloAlmacen");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn32 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Orden");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn33 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Almacen");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn34 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("existe");
			this.label3 = new System.Windows.Forms.Label();
			this.txtDescripcion = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.cdsMarceloCodigoD = new C1.Data.C1DataSet();
			this.barraDato1 = new BarraDatoSQL.BarraDatoSQL();
			this.grdTalla = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.label1 = new System.Windows.Forms.Label();
			this.txtCodigo = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.cdsMAlmacen = new C1.Data.C1DataSet();
			this.cdvMAlmacen = new C1.Data.C1DataView();
			this.cmbAlmacen = new Infragistics.Win.UltraWinGrid.UltraDropDown();
			this.btTotal = new Infragistics.Win.Misc.UltraButton();
			this.label2 = new System.Windows.Forms.Label();
			this.lblStock = new System.Windows.Forms.Label();
			this.lbl1 = new System.Windows.Forms.Label();
			((System.ComponentModel.ISupportInitialize)(this.txtDescripcion)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsMarceloCodigoD)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.grdTalla)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtCodigo)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsMAlmacen)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdvMAlmacen)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbAlmacen)).BeginInit();
			this.SuspendLayout();
			// 
			// label3
			// 
			this.label3.AutoSize = true;
			this.label3.Location = new System.Drawing.Point(14, 8);
			this.label3.Name = "label3";
			this.label3.Size = new System.Drawing.Size(43, 16);
			this.label3.TabIndex = 146;
			this.label3.Text = "Codigo:";
			// 
			// txtDescripcion
			// 
			appearance1.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtDescripcion.Appearance = appearance1;
			this.txtDescripcion.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsMarceloCodigoD, "_MarceloCodigo.Descripcion"));
			this.txtDescripcion.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtDescripcion.Location = new System.Drawing.Point(96, 40);
			this.txtDescripcion.Name = "txtDescripcion";
			this.txtDescripcion.Size = new System.Drawing.Size(528, 21);
			this.txtDescripcion.TabIndex = 145;
			// 
			// cdsMarceloCodigoD
			// 
			this.cdsMarceloCodigoD.BindingContextCtrl = this;
			this.cdsMarceloCodigoD.DataLibrary = "LibFacturacion";
			this.cdsMarceloCodigoD.DataLibraryUrl = "";
			this.cdsMarceloCodigoD.DataSetDef = "MarceloCodigoD";
			this.cdsMarceloCodigoD.FillOnRequest = false;
			this.cdsMarceloCodigoD.Locale = new System.Globalization.CultureInfo("es-MX");
			this.cdsMarceloCodigoD.Name = "cdsMarceloCodigoD";
			this.cdsMarceloCodigoD.SchemaClassName = "LibFacturacion.DataClass";
			this.cdsMarceloCodigoD.SchemaDef = null;
			this.cdsMarceloCodigoD.BeforeFill += new C1.Data.FillEventHandler(this.cdsMarceloCodigoD_BeforeFill);
			// 
			// barraDato1
			// 
			this.barraDato1.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.barraDato1.BarraMovimiento = true;
			this.barraDato1.DataMember = "_MarceloCodigo";
			this.barraDato1.DataNombreId = "idMarceloCodigo";
			this.barraDato1.DataOrden = "";
			this.barraDato1.DataSource = this.cdsMarceloCodigoD;
			this.barraDato1.DataTabla = "MarceloCodigo";
			this.barraDato1.DataTablaHija = "MarceloDetCodigo";
			this.barraDato1.Location = new System.Drawing.Point(168, 488);
			this.barraDato1.Name = "barraDato1";
			this.barraDato1.Size = new System.Drawing.Size(314, 24);
			this.barraDato1.TabIndex = 144;
			this.barraDato1.VisibleBorrar = true;
			this.barraDato1.VisibleBuscar = true;
			this.barraDato1.VisibleEditar = true;
			this.barraDato1.VisibleImprimir = true;
			this.barraDato1.VisibleNuevo = true;
			this.barraDato1.Buscar += new System.EventHandler(this.barraDato1_Buscar);
			this.barraDato1.Borrar += new System.EventHandler(this.barraDato1_Borrar);
			this.barraDato1.Imprimir += new System.EventHandler(this.barraDato1_Imprimir);
			this.barraDato1.Refresca += new System.EventHandler(this.barraDato1_Refresca);
			this.barraDato1.Crear += new System.EventHandler(this.barraDato1_Crear);
			this.barraDato1.Deshacer += new System.EventHandler(this.barraDato1_Deshacer);
			this.barraDato1.Editar += new System.EventHandler(this.barraDato1_Editar);
			this.barraDato1.Grabar += new System.EventHandler(this.barraDato1_Grabar);
			// 
			// grdTalla
			// 
			this.grdTalla.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.grdTalla.Cursor = System.Windows.Forms.Cursors.Default;
			this.grdTalla.DataMember = "_MarceloCodigo.MarceloCodigo - MarceloDetCodigo";
			this.grdTalla.DataSource = this.cdsMarceloCodigoD;
			appearance2.BackColor = System.Drawing.Color.White;
			this.grdTalla.DisplayLayout.Appearance = appearance2;
			ultraGridBand1.AddButtonCaption = "Articulo_Filtro";
			ultraGridColumn1.Header.VisiblePosition = 1;
			ultraGridColumn1.Hidden = true;
			ultraGridColumn2.Header.VisiblePosition = 3;
			ultraGridColumn2.Hidden = true;
			ultraGridColumn2.Width = 146;
			ultraGridColumn3.AutoEdit = false;
			ultraGridColumn3.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			ultraGridColumn3.Header.Caption = "Almacen";
			ultraGridColumn3.Header.VisiblePosition = 0;
			ultraGridColumn3.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.DropDownList;
			ultraGridColumn3.Width = 140;
			ultraGridColumn4.Header.VisiblePosition = 2;
			ultraGridColumn4.Hidden = true;
			ultraGridColumn5.Header.VisiblePosition = 4;
			ultraGridColumn5.Width = 46;
			ultraGridColumn6.Header.VisiblePosition = 5;
			ultraGridColumn6.Width = 36;
			ultraGridColumn7.Header.VisiblePosition = 6;
			ultraGridColumn7.Width = 36;
			ultraGridColumn8.Header.VisiblePosition = 7;
			ultraGridColumn8.Width = 30;
			ultraGridColumn9.Header.VisiblePosition = 8;
			ultraGridColumn9.Width = 36;
			ultraGridColumn10.Header.VisiblePosition = 9;
			ultraGridColumn10.Width = 36;
			ultraGridColumn11.Header.VisiblePosition = 10;
			ultraGridColumn11.Width = 36;
			ultraGridColumn12.Header.VisiblePosition = 11;
			ultraGridColumn12.Width = 36;
			ultraGridColumn13.Header.VisiblePosition = 12;
			ultraGridColumn13.Width = 36;
			ultraGridColumn14.Header.VisiblePosition = 13;
			ultraGridColumn14.Width = 36;
			ultraGridColumn15.Header.VisiblePosition = 14;
			ultraGridColumn15.Width = 36;
			ultraGridColumn16.Header.VisiblePosition = 15;
			ultraGridColumn16.Width = 36;
			ultraGridColumn17.Header.VisiblePosition = 16;
			ultraGridColumn17.Width = 36;
			ultraGridColumn18.Header.VisiblePosition = 17;
			ultraGridColumn18.Width = 34;
			ultraGridColumn19.Header.VisiblePosition = 18;
			ultraGridColumn19.Width = 36;
			ultraGridColumn20.Header.VisiblePosition = 19;
			ultraGridColumn20.Width = 36;
			ultraGridColumn21.Header.VisiblePosition = 20;
			ultraGridColumn21.Width = 36;
			ultraGridColumn22.Header.VisiblePosition = 21;
			ultraGridColumn22.Width = 36;
			ultraGridColumn23.Header.VisiblePosition = 22;
			ultraGridColumn23.Width = 36;
			ultraGridColumn24.Header.VisiblePosition = 23;
			ultraGridColumn24.Width = 36;
			ultraGridColumn25.Header.VisiblePosition = 24;
			ultraGridColumn25.Width = 36;
			ultraGridColumn26.Header.VisiblePosition = 25;
			ultraGridColumn26.Width = 36;
			ultraGridColumn27.Header.VisiblePosition = 26;
			ultraGridColumn27.Width = 36;
			ultraGridColumn28.Header.VisiblePosition = 27;
			ultraGridColumn28.Width = 36;
			ultraGridColumn29.Header.VisiblePosition = 28;
			ultraGridColumn29.Width = 36;
			ultraGridColumn30.Header.VisiblePosition = 29;
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
																										 ultraGridColumn30});
			summarySettings1.DisplayFormat = "{0: #,##0}";
			summarySettings1.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			summarySettings2.DisplayFormat = "{0: #,##0}";
			summarySettings2.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			summarySettings3.DisplayFormat = "{0: #,##0}";
			summarySettings3.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			summarySettings4.DisplayFormat = "{0: #,##0}";
			summarySettings4.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			summarySettings5.DisplayFormat = "{0: #,##0}";
			summarySettings5.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			summarySettings6.DisplayFormat = "{0: #,##0}";
			summarySettings6.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			summarySettings7.DisplayFormat = "{0: #,##0}";
			summarySettings7.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			summarySettings8.DisplayFormat = "{0: #,##0}";
			summarySettings8.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			summarySettings9.DisplayFormat = "{0: #,##0}";
			summarySettings9.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			summarySettings10.DisplayFormat = "{0: #,##0}";
			summarySettings10.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			summarySettings11.DisplayFormat = "{0: #,##0}";
			summarySettings11.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			summarySettings12.DisplayFormat = "{0: #,##0}";
			summarySettings12.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			summarySettings13.DisplayFormat = "{0: #,##0}";
			summarySettings13.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			summarySettings14.DisplayFormat = "{0: #,##0}";
			summarySettings14.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			summarySettings15.DisplayFormat = "{0: #,##0}";
			summarySettings15.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			summarySettings16.DisplayFormat = "{0: #,##0}";
			summarySettings16.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			summarySettings17.DisplayFormat = "{0: #,##0}";
			summarySettings17.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			summarySettings18.DisplayFormat = "{0: #,##0}";
			summarySettings18.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			summarySettings19.DisplayFormat = "{0: #,##0}";
			summarySettings19.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			summarySettings20.DisplayFormat = "{0: #,##0}";
			summarySettings20.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			summarySettings21.DisplayFormat = "{0: #,##0}";
			summarySettings21.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			summarySettings22.DisplayFormat = "{0: #,##0}";
			summarySettings22.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			summarySettings23.DisplayFormat = "{0: #,##0}";
			summarySettings23.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			summarySettings24.DisplayFormat = "{0: #,##0}";
			summarySettings24.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			summarySettings25.DisplayFormat = "{0: #,##0}";
			summarySettings25.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			summarySettings26.DisplayFormat = "{0: #,##0}";
			summarySettings26.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			ultraGridBand1.Summaries.AddRange(new Infragistics.Win.UltraWinGrid.SummarySettings[] {
																																															summarySettings1,
																																															summarySettings2,
																																															summarySettings3,
																																															summarySettings4,
																																															summarySettings5,
																																															summarySettings6,
																																															summarySettings7,
																																															summarySettings8,
																																															summarySettings9,
																																															summarySettings10,
																																															summarySettings11,
																																															summarySettings12,
																																															summarySettings13,
																																															summarySettings14,
																																															summarySettings15,
																																															summarySettings16,
																																															summarySettings17,
																																															summarySettings18,
																																															summarySettings19,
																																															summarySettings20,
																																															summarySettings21,
																																															summarySettings22,
																																															summarySettings23,
																																															summarySettings24,
																																															summarySettings25,
																																															summarySettings26});
			this.grdTalla.DisplayLayout.BandsSerializer.Add(ultraGridBand1);
			this.grdTalla.DisplayLayout.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.No;
			this.grdTalla.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.True;
			this.grdTalla.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.True;
			appearance3.BackColor = System.Drawing.Color.Transparent;
			this.grdTalla.DisplayLayout.Override.CardAreaAppearance = appearance3;
			appearance4.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance4.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance4.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance4.FontData.BoldAsString = "True";
			appearance4.FontData.Name = "Arial";
			appearance4.FontData.SizeInPoints = 10F;
			appearance4.ForeColor = System.Drawing.Color.White;
			appearance4.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.grdTalla.DisplayLayout.Override.HeaderAppearance = appearance4;
			appearance5.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance5.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance5.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.grdTalla.DisplayLayout.Override.RowSelectorAppearance = appearance5;
			appearance6.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance6.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance6.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.grdTalla.DisplayLayout.Override.SelectedRowAppearance = appearance6;
			this.grdTalla.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.grdTalla.Location = new System.Drawing.Point(16, 112);
			this.grdTalla.Name = "grdTalla";
			this.grdTalla.Size = new System.Drawing.Size(664, 360);
			this.grdTalla.TabIndex = 143;
			this.grdTalla.BeforeColPosChanged += new Infragistics.Win.UltraWinGrid.BeforeColPosChangedEventHandler(this.grdTalla_BeforeColPosChanged);
			this.grdTalla.Validated += new System.EventHandler(this.grdTalla_Validated);
			this.grdTalla.BeforeCellActivate += new Infragistics.Win.UltraWinGrid.CancelableCellEventHandler(this.grdTalla_BeforeCellActivate);
			this.grdTalla.KeyDown += new System.Windows.Forms.KeyEventHandler(this.grdTalla_KeyDown);
			this.grdTalla.KeyUp += new System.Windows.Forms.KeyEventHandler(this.grdTalla_KeyUp);
			this.grdTalla.AfterCellUpdate += new Infragistics.Win.UltraWinGrid.CellEventHandler(this.grdTalla_AfterCellUpdate);
			this.grdTalla.InitializeLayout += new Infragistics.Win.UltraWinGrid.InitializeLayoutEventHandler(this.grdTalla_InitializeLayout);
			// 
			// label1
			// 
			this.label1.AutoSize = true;
			this.label1.Location = new System.Drawing.Point(16, 40);
			this.label1.Name = "label1";
			this.label1.Size = new System.Drawing.Size(67, 16);
			this.label1.TabIndex = 149;
			this.label1.Text = "Descripcion:";
			// 
			// txtCodigo
			// 
			appearance7.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtCodigo.Appearance = appearance7;
			this.txtCodigo.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsMarceloCodigoD, "_MarceloCodigo.Codigo"));
			this.txtCodigo.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtCodigo.Location = new System.Drawing.Point(96, 8);
			this.txtCodigo.Name = "txtCodigo";
			this.txtCodigo.Size = new System.Drawing.Size(170, 21);
			this.txtCodigo.TabIndex = 150;
			// 
			// cdsMAlmacen
			// 
			this.cdsMAlmacen.BindingContextCtrl = this;
			this.cdsMAlmacen.DataLibrary = "LibFacturacion";
			this.cdsMAlmacen.DataLibraryUrl = "";
			this.cdsMAlmacen.DataSetDef = "dsMarceloAlmacen";
			this.cdsMAlmacen.Locale = new System.Globalization.CultureInfo("es-MX");
			this.cdsMAlmacen.Name = "cdsMAlmacen";
			this.cdsMAlmacen.SchemaClassName = "LibFacturacion.DataClass";
			this.cdsMAlmacen.SchemaDef = null;
			// 
			// cdvMAlmacen
			// 
			this.cdvMAlmacen.BindingContextCtrl = this;
			this.cdvMAlmacen.DataSet = this.cdsMAlmacen;
			this.cdvMAlmacen.TableName = "";
			this.cdvMAlmacen.TableViewName = "MarceloAlmacen";
			// 
			// cmbAlmacen
			// 
			this.cmbAlmacen.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbAlmacen.DataMember = "MarceloAlmacen";
			this.cmbAlmacen.DataSource = this.cdsMAlmacen;
			ultraGridBand2.AddButtonCaption = "Proyecto";
			ultraGridColumn31.Header.VisiblePosition = 1;
			ultraGridColumn31.Hidden = true;
			ultraGridColumn32.Header.VisiblePosition = 2;
			ultraGridColumn32.Hidden = true;
			ultraGridColumn33.Header.VisiblePosition = 0;
			ultraGridColumn33.Width = 221;
			ultraGridColumn34.Header.VisiblePosition = 3;
			ultraGridColumn34.Hidden = true;
			ultraGridBand2.Columns.AddRange(new object[] {
																										 ultraGridColumn31,
																										 ultraGridColumn32,
																										 ultraGridColumn33,
																										 ultraGridColumn34});
			this.cmbAlmacen.DisplayLayout.BandsSerializer.Add(ultraGridBand2);
			this.cmbAlmacen.DisplayMember = "Almacen";
			this.cmbAlmacen.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbAlmacen.Location = new System.Drawing.Point(200, 256);
			this.cmbAlmacen.Name = "cmbAlmacen";
			this.cmbAlmacen.Size = new System.Drawing.Size(240, 96);
			this.cmbAlmacen.TabIndex = 152;
			this.cmbAlmacen.ValueMember = "idMarceloAlmacen";
			this.cmbAlmacen.Visible = false;
			// 
			// btTotal
			// 
			this.btTotal.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btTotal.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			this.btTotal.Location = new System.Drawing.Point(624, 488);
			this.btTotal.Name = "btTotal";
			this.btTotal.Size = new System.Drawing.Size(56, 23);
			this.btTotal.TabIndex = 156;
			this.btTotal.Text = "Total";
			this.btTotal.Click += new System.EventHandler(this.btTotal_Click);
			// 
			// label2
			// 
			this.label2.Location = new System.Drawing.Point(16, 72);
			this.label2.Name = "label2";
			this.label2.Size = new System.Drawing.Size(56, 23);
			this.label2.TabIndex = 157;
			this.label2.Text = "Stock :";
			this.label2.Click += new System.EventHandler(this.label2_Click);
			// 
			// lblStock
			// 
			this.lblStock.Location = new System.Drawing.Point(88, 72);
			this.lblStock.Name = "lblStock";
			this.lblStock.Size = new System.Drawing.Size(56, 23);
			this.lblStock.TabIndex = 158;
			this.lblStock.Text = "label4";
			// 
			// lbl1
			// 
			this.lbl1.Location = new System.Drawing.Point(168, 72);
			this.lbl1.Name = "lbl1";
			this.lbl1.Size = new System.Drawing.Size(56, 23);
			this.lbl1.TabIndex = 159;
			this.lbl1.Text = "label4";
			// 
			// MarceloReparteC
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 13);
			this.ClientSize = new System.Drawing.Size(712, 517);
			this.Controls.Add(this.lbl1);
			this.Controls.Add(this.lblStock);
			this.Controls.Add(this.label2);
			this.Controls.Add(this.btTotal);
			this.Controls.Add(this.cmbAlmacen);
			this.Controls.Add(this.txtCodigo);
			this.Controls.Add(this.label1);
			this.Controls.Add(this.label3);
			this.Controls.Add(this.txtDescripcion);
			this.Controls.Add(this.barraDato1);
			this.Controls.Add(this.grdTalla);
			this.Name = "MarceloReparteC";
			this.Text = "MarceloReparteC";
			this.Load += new System.EventHandler(this.MarceloReparteC_Load);
			((System.ComponentModel.ISupportInitialize)(this.txtDescripcion)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsMarceloCodigoD)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.grdTalla)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtCodigo)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsMAlmacen)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdvMAlmacen)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbAlmacen)).EndInit();
			this.ResumeLayout(false);

		}
		#endregion

		private void cdsMarceloCodigoD_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			this.cdsMarceloCodigoD.Schema.Connections[0].ConnectionString = 
				miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);
		}

		private void barraDato1_Borrar(object sender, System.EventArgs e)
		{
		
		}

		private void barraDato1_Buscar(object sender, System.EventArgs e)
		{
		
		}

		private void barraDato1_Crear(object sender, System.EventArgs e)
		{
				barraDato1.CrearRegistro();  
		}

		private void barraDato1_Deshacer(object sender, System.EventArgs e)
		{
				this.barraDato1.DeshacerRegistro();  
		}

		private void barraDato1_Editar(object sender, System.EventArgs e)
		{
				this.barraDato1.EditarRegistro();  
		}

		private void barraDato1_Grabar(object sender, System.EventArgs e)
		{
			  cmdTotal();
				this.barraDato1.GrabaRegistro();  
		}

		private void barraDato1_Imprimir(object sender, System.EventArgs e)
		{
		
		}

		private void barraDato1_Refresca(object sender, System.EventArgs e)
		{
		
		}

		private void MarceloReparteC_Load(object sender, System.EventArgs e)
		{
		

			foreach(Infragistics.Win.UltraWinGrid.UltraGridColumn dc in grdTalla.DisplayLayout.Bands[0].Columns)
			{
				dc.Hidden = true;
			}
			grdTalla.DisplayLayout.Bands[0].Columns["idMarceloAlmacen"].Hidden = false;
			grdTalla.DisplayLayout.Bands[0].Columns["idMarceloAlmacen"].Header.VisiblePosition = 1;
			grdTalla.DisplayLayout.Bands[0].Columns["idMarceloAlmacen"].Header.Caption = "Almacen" ;
			grdTalla.DisplayLayout.Bands[0].Columns["idMarceloAlmacen"].CellActivation = Infragistics.Win.UltraWinGrid.Activation.NoEdit;

////////////

			string stSelect="select count(*) from marcelodetcodigo where idmarcelocodigo="+ idMarceloCod;
			int cant=Funcion.iEscalarSQL(cdsMarceloCodigoD,stSelect);
  
			if (cant==0 )	
			{	
				stSelect="insert into marcelodetcodigo (idmarceloalmacen,idmarcelocodigo) "
				+ " select idmarceloalmacen , "
				+	idMarceloCod   + " as idmarcelocodigo  from marceloalmacen where existe=0 order by orden ";
				Funcion.EjecutaSQL(cdsMarceloCodigoD,stSelect )  ;
			}

			Llena(); 

			this.barraDato1.HabilitaControles(false);
			this.barraDato1.IdRegistro =idMarceloCod;
			this.barraDato1.ProximoId(5); 
		}

		private void btReparte_Click(object sender, System.EventArgs e)
		{

		}

		private void grdTalla_InitializeLayout(object sender, Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs e)
		{
		
		}
 
		private void Llena()
		{
			string stColumna = "";
			string stTalla="";
			int iPos=1;
			int i=1;			
			while (i<=25 )
			{
				stTalla="";
				string stLista = "Select calla"+i +   " From marcelocodigo where idmarcelocodigo="+ idMarceloCod; 
				stTalla=Funcion.sEscalarSQL(cdsMarceloCodigoD,stLista);  
		
				if ( stTalla.StartsWith("Talla"))
				{
				}
				else
				{
					stColumna = "talla" + iPos.ToString().Trim();
					grdTalla.DisplayLayout.Bands[0].Columns[stColumna].Hidden = false;
					grdTalla.DisplayLayout.Bands[0].Columns[stColumna].Header.VisiblePosition = iPos+1;
					grdTalla.DisplayLayout.Bands[0].Columns[stColumna].Header.Caption = stTalla ;
					grdTalla.DisplayLayout.Bands[0].Columns[stColumna].CellActivation = Infragistics.Win.UltraWinGrid.Activation.AllowEdit;
					iPos =iPos +1;
				}
			i=i+1;
			}//while

			grdTalla.DisplayLayout.Bands[0].Columns["TotalF"].Hidden = false;
			grdTalla.DisplayLayout.Bands[0].Columns["TotalF"].Header.VisiblePosition = i;
			grdTalla.DisplayLayout.Bands[0].Columns["TotalF"].Header.Caption = "Total" ;
			grdTalla.DisplayLayout.Bands[0].Columns["TotalF"].CellActivation = Infragistics.Win.UltraWinGrid.Activation.NoEdit;


		}

		private void grdTalla_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{

			if (e.Alt) 
				return;
			
			if (grdTalla.ActiveCell == null) return;
			
			
			switch(e.KeyValue)
			{
				case (int) Keys.PageDown:
					if(grdTalla.ActiveCell.DroppedDown == false)
						grdTalla.ActiveCell.DroppedDown = true;
					break;

				case (int) Keys.Up:
					if (grdTalla.ActiveCell.DroppedDown) return;

					grdTalla.PerformAction(UltraGridAction.ExitEditMode, false, false);
					grdTalla.PerformAction(UltraGridAction.AboveCell, false, false);
					e.Handled = true;
					grdTalla.PerformAction(UltraGridAction.EnterEditMode, false, false);
					break;
				case (int) Keys.Down:
					if (grdTalla.ActiveCell.DroppedDown) return;
			
					grdTalla.PerformAction(UltraGridAction.ExitEditMode, false, false);
					grdTalla.PerformAction(UltraGridAction.BelowCell, false, false);
					e.Handled = true;
					grdTalla.PerformAction(UltraGridAction.EnterEditMode, false, false);
					break;
				case (int) Keys.Enter:
					grdTalla.PerformAction(UltraGridAction.ExitEditMode, false, false);
					grdTalla.PerformAction(UltraGridAction.NextCellByTab, false, false);
					e.Handled = true;
					grdTalla.PerformAction(UltraGridAction.EnterEditMode, false, false);
					break;
			}		
		}
		int t1=0,t2=0,t3=0,t4=0,t5=0,t6=0,t7=0,t8=0,t9=0,t10=0,t11=0,t12=0,t13=0,t14=0,t15=0,t16=0,t17=0,t18=0,t19=0,t20=0;
		int t21=0,t22=0,t23=0,t24=0,t25=0;

		private void btTotal_Click(object sender, System.EventArgs e)
		{
			cmdTotal();
			 
		}
			
			private void cmdTotal()
			{
			  int bod=17;
				t1=0;t2=0;t3=0;t4=0;t5=0;t6=0;t7=0;t8=0;t9=0;t10=0;t11=0;t12=0;t13=0;t14=0;t15=0;t16=0;t17=0;t18=0;t19=0;t20=0;
				t21=0;t22=0;t23=0;t24=0;t25=0;

				foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in grdTalla.Rows.All)
				{
					if ((int)dr.Cells["idMarceloAlmacen"].Value !=bod)
 					{
					if (dr.Cells["Talla1"].Value != System.DBNull.Value)
						t1=t1 +(int)dr.Cells["Talla1"].Value;
     			if (dr.Cells["Talla2"].Value != System.DBNull.Value)
						t2=t2 +(int)dr.Cells["Talla2"].Value;
					if (dr.Cells["Talla3"].Value != System.DBNull.Value)
						t3=t3 +(int)dr.Cells["Talla3"].Value;
					if (dr.Cells["Talla4"].Value != System.DBNull.Value)
						t4=t4 +(int)dr.Cells["Talla4"].Value;
					if (dr.Cells["Talla5"].Value != System.DBNull.Value)
						t5=t5 +(int)dr.Cells["Talla5"].Value;
					if (dr.Cells["Talla6"].Value != System.DBNull.Value)
						t6=t6 +(int)dr.Cells["Talla6"].Value;
					if (dr.Cells["Talla7"].Value != System.DBNull.Value)
						t7=t7 +(int)dr.Cells["Talla7"].Value;
					if (dr.Cells["Talla8"].Value != System.DBNull.Value)
						t8=t8 +(int)dr.Cells["Talla8"].Value;
					if (dr.Cells["Talla9"].Value != System.DBNull.Value)
						t9=t9 +(int)dr.Cells["Talla9"].Value;
					if (dr.Cells["Talla10"].Value != System.DBNull.Value)
						t10=t10 +(int)dr.Cells["Talla10"].Value;
					if (dr.Cells["Talla11"].Value != System.DBNull.Value)
						t11=t11 +(int)dr.Cells["Talla11"].Value;
					if (dr.Cells["Talla12"].Value != System.DBNull.Value)
						t12=t12 +(int)dr.Cells["Talla12"].Value;
					if (dr.Cells["Talla13"].Value != System.DBNull.Value)
						t13=t13 +(int)dr.Cells["Talla13"].Value;
					if (dr.Cells["Talla14"].Value != System.DBNull.Value)
						t14=t14 +(int)dr.Cells["Talla14"].Value;
					if (dr.Cells["Talla15"].Value != System.DBNull.Value)
						t15=t15 +(int)dr.Cells["Talla15"].Value;
					if (dr.Cells["Talla16"].Value != System.DBNull.Value)
						t16=t16 +(int)dr.Cells["Talla16"].Value;
					if (dr.Cells["Talla17"].Value != System.DBNull.Value)
						t17=t17 +(int)dr.Cells["Talla17"].Value;
					if (dr.Cells["Talla18"].Value != System.DBNull.Value)
						t18=t18 +(int)dr.Cells["Talla18"].Value;
					if (dr.Cells["Talla19"].Value != System.DBNull.Value)
						t19=t19 +(int)dr.Cells["Talla19"].Value;
					if (dr.Cells["Talla20"].Value != System.DBNull.Value)
						t20=t20 +(int)dr.Cells["Talla20"].Value;
					if (dr.Cells["Talla21"].Value != System.DBNull.Value)
						t21=t21 +(int)dr.Cells["Talla21"].Value;
					if (dr.Cells["Talla22"].Value != System.DBNull.Value)
						t22=t22 +(int)dr.Cells["Talla22"].Value;
					if (dr.Cells["Talla23"].Value != System.DBNull.Value)
						t23=t23 +(int)dr.Cells["Talla23"].Value;
					if (dr.Cells["Talla24"].Value != System.DBNull.Value)
						t24=t24 +(int)dr.Cells["Talla24"].Value;
					if (dr.Cells["Talla25"].Value != System.DBNull.Value)
						t25=t25 +(int)dr.Cells["Talla25"].Value;
					}	
				}	
				string stSelect ="select isnull(talla1,0) from marcelocodigo where idmarcelocodigo=" + idMarceloCod.ToString();
				int tc1=Funcion.iEscalarSQL(cdsMAlmacen,stSelect);  
				if (t1>tc1) 
				{
					int excede=t1-tc1;
					seexcede=1; 
					MessageBox.Show("Se excede en " + excede.ToString() + " - Solo hay " + tc1.ToString());
				}	
				else
				{
					foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in grdTalla.Rows.All)
					{
						if ((int)dr.Cells["idMarceloAlmacen"].Value ==bod)//es bodega
						{
							dr.Cells["Talla1"].Value=tc1-t1;
						}
					}
				}	


				stSelect ="select isnull(talla2,0) from marcelocodigo where idmarcelocodigo=" + idMarceloCod.ToString();
				int tc2=Funcion.iEscalarSQL(cdsMAlmacen,stSelect);  
				if (t2>tc2) 
				{
					int excede=t2-tc2;
					seexcede=1; 
					MessageBox.Show("Se excede en " + excede.ToString() + " - Solo hay " + tc2.ToString());
				}	
				else
				{
					foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in grdTalla.Rows.All)
					{
						if ((int)dr.Cells["idMarceloAlmacen"].Value ==bod)//es bodega
						{
							dr.Cells["Talla2"].Value=tc2-t2;
						}
					}
				}	
	
				stSelect ="select isnull(talla3,0) from marcelocodigo where idmarcelocodigo=" + idMarceloCod.ToString();
				int tc3=Funcion.iEscalarSQL(cdsMAlmacen,stSelect);  
				if (t3>tc3) 
				{
					int excede=t3-tc3;
					seexcede=1; 
					MessageBox.Show("Se excede en " + excede.ToString() + " - Solo hay " + tc3.ToString());
				}	
				else
				{
					foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in grdTalla.Rows.All)
					{
						if ((int)dr.Cells["idMarceloAlmacen"].Value ==bod)//es bodega
						{
							dr.Cells["Talla3"].Value=tc3-t3;
						}
					}
				}	

				stSelect ="select isnull(talla4,0) from marcelocodigo where idmarcelocodigo=" + idMarceloCod.ToString();
				int tc4=Funcion.iEscalarSQL(cdsMAlmacen,stSelect);  
				if (t4>tc4) 
				{
					int excede=t4-tc4;
					seexcede=1; 
					MessageBox.Show("Se excede en " + excede.ToString() + " - Solo hay " + tc4.ToString());
				}	
				else
				{
					foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in grdTalla.Rows.All)
					{
						if ((int)dr.Cells["idMarceloAlmacen"].Value ==bod)//es bodega
						{
							dr.Cells["Talla4"].Value=tc4-t4;
						}
					}
				}	
		
				stSelect ="select isnull(talla5,0) from marcelocodigo where idmarcelocodigo=" + idMarceloCod.ToString();
				int tc5=Funcion.iEscalarSQL(cdsMAlmacen,stSelect);  
				if (t5>tc5) 
				{
					int excede=t5-tc5;
					seexcede=1; 
					MessageBox.Show("Se excede en " + excede.ToString() + " - Solo hay " + tc5.ToString());
				}	
				else
				{
					foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in grdTalla.Rows.All)
					{
						if ((int)dr.Cells["idMarceloAlmacen"].Value ==bod)//es bodega
						{
							dr.Cells["Talla5"].Value=tc5-t5;
						}
					}
				}	
		
				stSelect ="select isnull(talla6,0) from marcelocodigo where idmarcelocodigo=" + idMarceloCod.ToString();
				int tc6=Funcion.iEscalarSQL(cdsMAlmacen,stSelect);  
				if (t6>tc6) 
				{
					int excede=t6-tc6;
					seexcede=1; 
					MessageBox.Show("Se excede en " + excede.ToString() + " - Solo hay " + tc6.ToString());
				}	
				else
				{
					foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in grdTalla.Rows.All)
					{
						if ((int)dr.Cells["idMarceloAlmacen"].Value ==bod)//es bodega
						{
							dr.Cells["Talla6"].Value=tc6-t6;
						}
					}
				}	
		
				stSelect ="select isnull(talla7,0) from marcelocodigo where idmarcelocodigo=" + idMarceloCod.ToString();
				int tc7=Funcion.iEscalarSQL(cdsMAlmacen,stSelect);  
				if (t7>tc7) 
				{
					int excede=t7-tc7;
					seexcede=1; 
					MessageBox.Show("Se excede en " + excede.ToString() + " - Solo hay " + tc7.ToString());
				}	
				else
				{
					foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in grdTalla.Rows.All)
					{
						if ((int)dr.Cells["idMarceloAlmacen"].Value ==bod)//es bodega
						{
							dr.Cells["Talla7"].Value=tc7-t7;
						}
					}
				}	

				stSelect ="select isnull(talla8,0) from marcelocodigo where idmarcelocodigo=" + idMarceloCod.ToString();
				int tc8=Funcion.iEscalarSQL(cdsMAlmacen,stSelect);  
				if (t8>tc8) 
				{
					int excede=t8-tc8;
					seexcede=1; 
					MessageBox.Show("Se excede en " + excede.ToString() + " - Solo hay " + tc8.ToString());
				}	
				else
				{
					foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in grdTalla.Rows.All)
					{
						if ((int)dr.Cells["idMarceloAlmacen"].Value ==bod)//es bodega
						{
							dr.Cells["Talla8"].Value=tc8-t8;
						}
					}
				}	
		
				stSelect ="select isnull(talla9,0) from marcelocodigo where idmarcelocodigo=" + idMarceloCod.ToString();
				int tc9=Funcion.iEscalarSQL(cdsMAlmacen,stSelect);  
				if (t9>tc9) 
				{
					int excede=t9-tc9;
					seexcede=1; 
					MessageBox.Show("Se excede en " + excede.ToString() + " - Solo hay " + tc9.ToString());
				}	
				else
				{
					foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in grdTalla.Rows.All)
					{
						if ((int)dr.Cells["idMarceloAlmacen"].Value ==bod)//es bodega
						{
							dr.Cells["Talla9"].Value=tc9-t9;
						}
					}
				}	
		
				stSelect ="select isnull(talla10,0) from marcelocodigo where idmarcelocodigo=" + idMarceloCod.ToString();
				int tc10=Funcion.iEscalarSQL(cdsMAlmacen,stSelect);  
				if (t10>tc10) 
				{
					int excede=t10-tc10;
					seexcede=1; 
					MessageBox.Show("Se excede en " + excede.ToString() + " - Solo hay " + tc10.ToString());
				}	
				else
				{
					foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in grdTalla.Rows.All)
					{
						if ((int)dr.Cells["idMarceloAlmacen"].Value ==bod)//es bodega
						{
							dr.Cells["Talla10"].Value=tc10-t10;
						}
					}
				}	
		
				stSelect ="select isnull(talla11,0) from marcelocodigo where idmarcelocodigo=" + idMarceloCod.ToString();
				int tc11=Funcion.iEscalarSQL(cdsMAlmacen,stSelect);  
				if (t11>tc11) 
				{
					int excede=t11-tc11;
					seexcede=1; 
					MessageBox.Show("Se excede en " + excede.ToString() + " - Solo hay " + tc11.ToString());
				}	
				else
				{
					foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in grdTalla.Rows.All)
					{
						if ((int)dr.Cells["idMarceloAlmacen"].Value ==bod)//es bodega
						{
							dr.Cells["Talla11"].Value=tc11-t11;
						}
					}
				}	

				stSelect ="select isnull(talla12,0) from marcelocodigo where idmarcelocodigo=" + idMarceloCod.ToString();
				int tc12=Funcion.iEscalarSQL(cdsMAlmacen,stSelect);  
				if (t12>tc12) 
				{
					int excede=t12-tc12;
					seexcede=1; 
					MessageBox.Show("Se excede en " + excede.ToString() + " - Solo hay " + tc12.ToString());
				}	
				else
				{
					foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in grdTalla.Rows.All)
					{
						if ((int)dr.Cells["idMarceloAlmacen"].Value ==bod)//es bodega
						{
							dr.Cells["Talla12"].Value=tc12-t12;
						}
					}
				}	
		
				stSelect ="select isnull(talla13,0) from marcelocodigo where idmarcelocodigo=" + idMarceloCod.ToString();
				int tc13=Funcion.iEscalarSQL(cdsMAlmacen,stSelect);  
				if (t13>tc13) 
				{
					int excede=t13-tc13;
					seexcede=1; 
					MessageBox.Show("Se excede en " + excede.ToString() + " - Solo hay " + tc13.ToString());
				}	
				else
				{
					foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in grdTalla.Rows.All)
					{
						if ((int)dr.Cells["idMarceloAlmacen"].Value ==bod)//es bodega
						{
							dr.Cells["Talla13"].Value=tc13-t13;
						}
					}
				}	

				stSelect ="select isnull(talla14,0) from marcelocodigo where idmarcelocodigo=" + idMarceloCod.ToString();
				int tc14=Funcion.iEscalarSQL(cdsMAlmacen,stSelect);  
				if (t14>tc14) 
				{
					int excede=t14-tc14;
					seexcede=1; 
					MessageBox.Show("Se excede en " + excede.ToString() + " - Solo hay " + tc14.ToString());
				}	
				else
				{
					foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in grdTalla.Rows.All)
					{
						if ((int)dr.Cells["idMarceloAlmacen"].Value ==bod)//es bodega
						{
							dr.Cells["Talla14"].Value=tc14-t14;
						}
					}
				}	

				stSelect ="select isnull(talla15,0) from marcelocodigo where idmarcelocodigo=" + idMarceloCod.ToString();
				int tc15=Funcion.iEscalarSQL(cdsMAlmacen,stSelect);  
				if (t15>tc15) 
				{
					int excede=t15-tc15;
					seexcede=1; 
					MessageBox.Show("Se excede en " + excede.ToString() + " - Solo hay " + tc15.ToString());
				}	
				else
				{
					foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in grdTalla.Rows.All)
					{
						if ((int)dr.Cells["idMarceloAlmacen"].Value ==bod)//es bodega
						{
							dr.Cells["Talla15"].Value=tc15-t15;
						}
					}
				}	
		
				stSelect ="select isnull(talla16,0) from marcelocodigo where idmarcelocodigo=" + idMarceloCod.ToString();
				int tc16=Funcion.iEscalarSQL(cdsMAlmacen,stSelect);  
				if (t16>tc16) 
				{
					int excede=t16-tc16;
					seexcede=1; 
					MessageBox.Show("Se excede en " + excede.ToString() + " - Solo hay " + tc16.ToString());
				}	
				else
				{
					foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in grdTalla.Rows.All)
					{
						if ((int)dr.Cells["idMarceloAlmacen"].Value ==bod)//es bodega
						{
							dr.Cells["Talla16"].Value=tc16-t16;
						}
					}
				}	
		
				stSelect ="select isnull(talla17,0) from marcelocodigo where idmarcelocodigo=" + idMarceloCod.ToString();
				int tc17=Funcion.iEscalarSQL(cdsMAlmacen,stSelect);  
				if (t17>tc17) 
				{
					int excede=t17-tc17;
					seexcede=1; 
					MessageBox.Show("Se excede en " + excede.ToString() + " - Solo hay " + tc17.ToString());
				}	
				else
				{
					foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in grdTalla.Rows.All)
					{
						if ((int)dr.Cells["idMarceloAlmacen"].Value ==bod)//es bodega
						{
							dr.Cells["Talla17"].Value=tc17-t17;
						}
					}
				}	

				stSelect ="select isnull(talla18,0) from marcelocodigo where idmarcelocodigo=" + idMarceloCod.ToString();
				int tc18=Funcion.iEscalarSQL(cdsMAlmacen,stSelect);  
				if (t18>tc18) 
				{
					int excede=t18-tc18;
					seexcede=1; 
					MessageBox.Show("Se excede en " + excede.ToString() + " - Solo hay " + tc18.ToString());
				}	
				else
				{
					foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in grdTalla.Rows.All)
					{
						if ((int)dr.Cells["idMarceloAlmacen"].Value ==bod)//es bodega
						{
							dr.Cells["Talla18"].Value=tc18-t18;
						}
					}
				}	
		
				stSelect ="select isnull(talla19,0) from marcelocodigo where idmarcelocodigo=" + idMarceloCod.ToString();
				int tc19=Funcion.iEscalarSQL(cdsMAlmacen,stSelect);  
				if (t19>tc19) 
				{
					int excede=t19-tc19;
					seexcede=1; 
					MessageBox.Show("Se excede en " + excede.ToString() + " - Solo hay " + tc19.ToString());
				}	
				else
				{
					foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in grdTalla.Rows.All)
					{
						if ((int)dr.Cells["idMarceloAlmacen"].Value ==bod)//es bodega
						{
							dr.Cells["Talla19"].Value=tc19-t19;
						}
					}
				}	
		
				stSelect ="select isnull(talla20,0) from marcelocodigo where idmarcelocodigo=" + idMarceloCod.ToString();
				int tc20=Funcion.iEscalarSQL(cdsMAlmacen,stSelect);  
				if (t20>tc20) 
				{
					int excede=t20-tc20;
					seexcede=1; 
					MessageBox.Show("Se excede en " + excede.ToString() + " - Solo hay " + tc20.ToString());
				}	
				else
				{
					foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in grdTalla.Rows.All)
					{
						if ((int)dr.Cells["idMarceloAlmacen"].Value ==bod)//es bodega
						{
							dr.Cells["Talla20"].Value=tc20-t20;
						}
					}
				}	
		
				stSelect ="select isnull(talla21,0) from marcelocodigo where idmarcelocodigo=" + idMarceloCod.ToString();
				int tc21=Funcion.iEscalarSQL(cdsMAlmacen,stSelect);  
				if (t21>tc20) 
				{
					int excede=t21-tc21;
					seexcede=1; 
					MessageBox.Show("Se excede en " + excede.ToString() + " - Solo hay " + tc21.ToString());
				}		
				else
				{
					foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in grdTalla.Rows.All)
					{
						if ((int)dr.Cells["idMarceloAlmacen"].Value ==bod)//es bodega
						{
							dr.Cells["Talla21"].Value=tc21-t21;
						}
					}
				}	

				stSelect ="select isnull(talla22,0) from marcelocodigo where idmarcelocodigo=" + idMarceloCod.ToString();
				int tc22=Funcion.iEscalarSQL(cdsMAlmacen,stSelect);  
				if (t22>tc22) 
				{
					int excede=t22-tc22;
					seexcede=1; 
					MessageBox.Show("Se excede en " + excede.ToString() + " - Solo hay " + tc22.ToString());
				}
				else
				{
					foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in grdTalla.Rows.All)
					{
						if ((int)dr.Cells["idMarceloAlmacen"].Value ==bod)//es bodega
						{
							dr.Cells["Talla22"].Value=tc22-t22;
						}
					}
				}	
		
				stSelect ="select isnull(talla23,0) from marcelocodigo where idmarcelocodigo=" + idMarceloCod.ToString();
				int tc23=Funcion.iEscalarSQL(cdsMAlmacen,stSelect);  
				if (t23>tc23) 
				{
					int excede=t23-tc23;
					seexcede=1; 
					MessageBox.Show("Se excede en " + excede.ToString() + " - Solo hay " + tc23.ToString());
				}		
				else
				{
					foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in grdTalla.Rows.All)
					{
						if ((int)dr.Cells["idMarceloAlmacen"].Value ==bod)//es bodega
						{
							dr.Cells["Talla23"].Value=tc23-t23;
						}
					}
				}	

				stSelect ="select isnull(talla24,0) from marcelocodigo where idmarcelocodigo=" + idMarceloCod.ToString();
				int tc24=Funcion.iEscalarSQL(cdsMAlmacen,stSelect);  
				if (t24>tc24) 
				{
					int excede=t24-tc24;
					seexcede=1; 
					MessageBox.Show("Se excede en " + excede.ToString() + " - Solo hay " + tc24.ToString());
				}		
				else
				{
					foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in grdTalla.Rows.All)
					{
						if ((int)dr.Cells["idMarceloAlmacen"].Value ==bod)//es bodega
						{
							dr.Cells["Talla24"].Value=tc24-t24;
						}
					}
				}	

				stSelect ="select isnull(talla25,0) from marcelocodigo where idmarcelocodigo=" + idMarceloCod.ToString();
				int tc25=Funcion.iEscalarSQL(cdsMAlmacen,stSelect);  
				if (t25>tc25) 
				{
					int excede=t25-tc25;
					seexcede=1; 
					MessageBox.Show("Se excede en " + excede.ToString() + " - Solo hay " + tc25.ToString());
				}		
				else
				{
					foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in grdTalla.Rows.All)
					{
						if ((int)dr.Cells["idMarceloAlmacen"].Value ==bod)//es bodega
						{
							dr.Cells["Talla25"].Value=tc25-t25;
						}
					}
				}	

				
			}

		private void grdTalla_KeyUp(object sender, System.Windows.Forms.KeyEventArgs e)
		{

		}

		private void grdTalla_Validated(object sender, System.EventArgs e)
		{
	
		}

		private void grdTalla_AfterCellUpdate(object sender, Infragistics.Win.UltraWinGrid.CellEventArgs e)
		{
			if ((int)e.Cell.Value==0 )
			{
				cmdTotal(); 
			}	
			else				    
			{
				cmdTotal(); 
				if(seexcede==1) 
					e.Cell.Value =0;  
				seexcede=0; 
        

			}

			}

		private void label2_Click(object sender, System.EventArgs e)
		{
		
		}

		private void grdTalla_BeforeColPosChanged(object sender, Infragistics.Win.UltraWinGrid.BeforeColPosChangedEventArgs e)
		{

		}

		private void grdTalla_BeforeCellActivate(object sender, Infragistics.Win.UltraWinGrid.CancelableCellEventArgs e)
		{
				string stSelect;
				if (e.Cell.Column.ToString() == "Talla1")
				{
					stSelect ="select isnull(talla1,0) from marcelocodigo where idmarcelocodigo=" + idMarceloCod.ToString();
					lblStock.Text =Funcion.iEscalarSQL(cdsMAlmacen,stSelect).ToString() ;  
					lbl1.Text="1"; 
				}
				if (e.Cell.Column.ToString() == "Talla2")
				{
					stSelect ="select isnull(talla2,0) from marcelocodigo where idmarcelocodigo=" + idMarceloCod.ToString();
					lblStock.Text =Funcion.iEscalarSQL(cdsMAlmacen,stSelect).ToString() ;  
					lbl1.Text="2"; 
				}
				if (e.Cell.Column.ToString() == "Talla3")
				{
					stSelect ="select isnull(talla3,0) from marcelocodigo where idmarcelocodigo=" + idMarceloCod.ToString();
					lblStock.Text =Funcion.iEscalarSQL(cdsMAlmacen,stSelect).ToString() ;  
					lbl1.Text="3"; 
				}
				if (e.Cell.Column.ToString() == "Talla4")
				{
					stSelect ="select isnull(talla4,0) from marcelocodigo where idmarcelocodigo=" + idMarceloCod.ToString();
					lblStock.Text =Funcion.iEscalarSQL(cdsMAlmacen,stSelect).ToString() ;  
					lbl1.Text="4"; 
				}	
			if (e.Cell.Column.ToString() == "Talla5")
			{
				stSelect ="select isnull(talla5,0) from marcelocodigo where idmarcelocodigo=" + idMarceloCod.ToString();
				lblStock.Text =Funcion.iEscalarSQL(cdsMAlmacen,stSelect).ToString() ;  
				lbl1.Text="5"; 
			}
			if (e.Cell.Column.ToString() == "Talla6")
			{
				stSelect ="select isnull(talla6,0) from marcelocodigo where idmarcelocodigo=" + idMarceloCod.ToString();
				lblStock.Text =Funcion.iEscalarSQL(cdsMAlmacen,stSelect).ToString() ;  
				lbl1.Text="6"; 
			}
			if (e.Cell.Column.ToString() == "Talla7")
			{
				stSelect ="select isnull(talla7,0) from marcelocodigo where idmarcelocodigo=" + idMarceloCod.ToString();
				lblStock.Text =Funcion.iEscalarSQL(cdsMAlmacen,stSelect).ToString() ;  
				lbl1.Text="7"; 
			}
			if (e.Cell.Column.ToString() == "Talla8")
			{
				stSelect ="select isnull(talla8,0) from marcelocodigo where idmarcelocodigo=" + idMarceloCod.ToString();
				lblStock.Text =Funcion.iEscalarSQL(cdsMAlmacen,stSelect).ToString() ;  
				lbl1.Text="8"; 
			}	
			if (e.Cell.Column.ToString() == "Talla9")
			{
				stSelect ="select isnull(talla9,0) from marcelocodigo where idmarcelocodigo=" + idMarceloCod.ToString();
				lblStock.Text =Funcion.iEscalarSQL(cdsMAlmacen,stSelect).ToString() ;  
				lbl1.Text="9"; 
			}
			if (e.Cell.Column.ToString() == "Talla10")
			{
				stSelect ="select isnull(talla10,0) from marcelocodigo where idmarcelocodigo=" + idMarceloCod.ToString();
				lblStock.Text =Funcion.iEscalarSQL(cdsMAlmacen,stSelect).ToString() ;  
				lbl1.Text="10"; 
			}
			if (e.Cell.Column.ToString() == "Talla11")
			{
				stSelect ="select isnull(talla11,0) from marcelocodigo where idmarcelocodigo=" + idMarceloCod.ToString();
				lblStock.Text =Funcion.iEscalarSQL(cdsMAlmacen,stSelect).ToString() ;  
				lbl1.Text="11"; 
			}
			if (e.Cell.Column.ToString() == "Talla12")
			{
				stSelect ="select isnull(talla12,0) from marcelocodigo where idmarcelocodigo=" + idMarceloCod.ToString();
				lblStock.Text =Funcion.iEscalarSQL(cdsMAlmacen,stSelect).ToString() ;  
				lbl1.Text="12"; 
			}	
			if (e.Cell.Column.ToString() == "Talla13")
			{
				stSelect ="select isnull(talla13,0) from marcelocodigo where idmarcelocodigo=" + idMarceloCod.ToString();
				lblStock.Text =Funcion.iEscalarSQL(cdsMAlmacen,stSelect).ToString() ;  
				lbl1.Text="13"; 
			}
			if (e.Cell.Column.ToString() == "Talla14")
			{
				stSelect ="select isnull(talla14,0) from marcelocodigo where idmarcelocodigo=" + idMarceloCod.ToString();
				lblStock.Text =Funcion.iEscalarSQL(cdsMAlmacen,stSelect).ToString() ;  
				lbl1.Text="14"; 
			}
			if (e.Cell.Column.ToString() == "Talla15")
			{
				stSelect ="select isnull(talla15,0) from marcelocodigo where idmarcelocodigo=" + idMarceloCod.ToString();
				lblStock.Text =Funcion.iEscalarSQL(cdsMAlmacen,stSelect).ToString() ;  
				lbl1.Text="15"; 
			}
			if (e.Cell.Column.ToString() == "Talla16")
			{
				stSelect ="select isnull(talla16,0) from marcelocodigo where idmarcelocodigo=" + idMarceloCod.ToString();
				lblStock.Text =Funcion.iEscalarSQL(cdsMAlmacen,stSelect).ToString() ;  
				lbl1.Text="16"; 
			}	
			if (e.Cell.Column.ToString() == "Talla17")
			{
				stSelect ="select isnull(talla17,0) from marcelocodigo where idmarcelocodigo=" + idMarceloCod.ToString();
				lblStock.Text =Funcion.iEscalarSQL(cdsMAlmacen,stSelect).ToString() ;  
				lbl1.Text="17"; 
			}
			if (e.Cell.Column.ToString() == "Talla18")
			{
				stSelect ="select isnull(talla18,0) from marcelocodigo where idmarcelocodigo=" + idMarceloCod.ToString();
				lblStock.Text =Funcion.iEscalarSQL(cdsMAlmacen,stSelect).ToString() ;  
				lbl1.Text="18"; 
			}
			if (e.Cell.Column.ToString() == "Talla19")
			{
				stSelect ="select isnull(talla19,0) from marcelocodigo where idmarcelocodigo=" + idMarceloCod.ToString();
				lblStock.Text =Funcion.iEscalarSQL(cdsMAlmacen,stSelect).ToString() ;  
				lbl1.Text="19"; 
			}
			if (e.Cell.Column.ToString() == "Talla20")
			{
				stSelect ="select isnull(talla20,0) from marcelocodigo where idmarcelocodigo=" + idMarceloCod.ToString();
				lblStock.Text =Funcion.iEscalarSQL(cdsMAlmacen,stSelect).ToString() ;  
				lbl1.Text="20"; 
			}	
			if (e.Cell.Column.ToString() == "Talla21")
			{
				stSelect ="select isnull(talla21,0) from marcelocodigo where idmarcelocodigo=" + idMarceloCod.ToString();
				lblStock.Text =Funcion.iEscalarSQL(cdsMAlmacen,stSelect).ToString() ;  
				lbl1.Text="21"; 
			}	
			if (e.Cell.Column.ToString() == "Talla22")
			{
				stSelect ="select isnull(talla22,0) from marcelocodigo where idmarcelocodigo=" + idMarceloCod.ToString();
				lblStock.Text =Funcion.iEscalarSQL(cdsMAlmacen,stSelect).ToString() ;  
				lbl1.Text="22"; 
			}	
			if (e.Cell.Column.ToString() == "Talla23")
			{
				stSelect ="select isnull(talla23,0) from marcelocodigo where idmarcelocodigo=" + idMarceloCod.ToString();
				lblStock.Text =Funcion.iEscalarSQL(cdsMAlmacen,stSelect).ToString() ;  
				lbl1.Text="23"; 
			}	
			if (e.Cell.Column.ToString() == "Talla24")
			{
				stSelect ="select isnull(talla24,0) from marcelocodigo where idmarcelocodigo=" + idMarceloCod.ToString();
				lblStock.Text =Funcion.iEscalarSQL(cdsMAlmacen,stSelect).ToString() ;  
				lbl1.Text="24"; 
			}	
			if (e.Cell.Column.ToString() == "Talla25")
			{
				stSelect ="select isnull(talla25,0) from marcelocodigo where idmarcelocodigo=" + idMarceloCod.ToString();
				lblStock.Text =Funcion.iEscalarSQL(cdsMAlmacen,stSelect).ToString() ;  
				lbl1.Text="24"; 
			}	
		}
	}
}
