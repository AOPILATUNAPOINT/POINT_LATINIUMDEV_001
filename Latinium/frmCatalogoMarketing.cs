using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using DevExpress.XtraEditors;
using System.IO;
using System.Data.SqlClient;
using System.Data;
using System.Text;
using System.Threading;
using System.Collections.Specialized;

namespace Latinium
{
	/// <summary>
	/// Descripción breve de frmCatalogoMarketing.
	/// </summary>
	public class frmCatalogoMarketing : System.Windows.Forms.Form
	{
			private Acceso miAcceso;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource1;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource2;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtObservaciones;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtBusqueda;
		private System.Windows.Forms.GroupBox groupBox1;
		private Infragistics.Win.UltraWinGrid.ExcelExport.UltraGridExcelExporter ultraGridExcelExporter1;
		private Infragistics.Win.UltraWinEditors.UltraOptionSet optEstado;
		private System.Windows.Forms.Label lblContador;
		private System.Windows.Forms.Button btnExcel;
		private System.Windows.Forms.Button btnVer;
		private Infragistics.Win.UltraWinEditors.UltraOptionSet optEstadoCaja;
		private Infragistics.Win.UltraWinTabControl.UltraTabControl utcEstadoCajas;
		private Infragistics.Win.UltraWinTabControl.UltraTabSharedControlsPage ultraTabSharedControlsPage1;
		private Infragistics.Win.UltraWinTabControl.UltraTabPageControl pgEntrega;
		private Infragistics.Win.UltraWinGrid.UltraGrid ultraGrid2;
		private Infragistics.Win.UltraWinTabControl.UltraTabPageControl ultraTabPageControl9;
		private Infragistics.Win.UltraWinGrid.UltraGrid ultraGrid1;
		private System.Windows.Forms.Label label17;
		private System.Windows.Forms.Label label18;
		private Infragistics.Win.UltraWinSchedule.UltraCalendarCombo dtHasta;
		private Infragistics.Win.UltraWinSchedule.UltraCalendarCombo dtDesde;
		private System.Windows.Forms.Label label36;
		private C1.Data.C1DataSet cdsSeteoF;
		private System.Windows.Forms.SaveFileDialog saveFileDialog1;
		private System.Windows.Forms.MainMenu mainMenu1;
		private Infragistics.Win.UltraWinTabControl.UltraTabPageControl ultraTabPageControl1;
		private System.Windows.Forms.PictureBox pbImagen;
		private System.Windows.Forms.Button btnImagen;
		private System.Windows.Forms.Button btnGuardar;
		private System.Windows.Forms.Button btnFinalizar;
		private System.Windows.Forms.RichTextBox rtbDescripcion;
		private System.Windows.Forms.Label lblCliente;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtPromocion;
		private System.Windows.Forms.Label label1;
		private System.Windows.Forms.OpenFileDialog openFileDialog1;
		private System.Windows.Forms.Label lblImagen;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtidCatalogo;
		private System.Windows.Forms.Button btnFuente;
		private System.Windows.Forms.Button btnColor;
		private System.Windows.Forms.Label lblEstado;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtEstadoCatalogo;
		/// <summary>
		/// Variable del diseñador requerida.
		/// </summary>
		private System.ComponentModel.Container components = null;

		public frmCatalogoMarketing()
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
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand1 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn1 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCatalogo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn2 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("NombreCatalogo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn3 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("PaginaCatalogo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn4 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("EstadoCatalogo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn5 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Desde");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn6 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Hasta");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn7 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idTipoCatalogo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn8 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Seleccion", 0);
			Infragistics.Win.Appearance appearance2 = new Infragistics.Win.Appearance();
			System.Resources.ResourceManager resources = new System.Resources.ResourceManager(typeof(frmCatalogoMarketing));
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn9 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Anular", 1, null, 0, Infragistics.Win.UltraWinGrid.SortIndicator.Ascending, false);
			Infragistics.Win.Appearance appearance3 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn10 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Sel1", 2);
			Infragistics.Win.Appearance appearance4 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance5 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance6 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance7 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance8 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance9 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance10 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance11 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn1 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idCatalogo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn2 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("NombreCatalogo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn3 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("PaginaCatalogo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn4 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("EstadoCatalogo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn5 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Desde");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn6 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Hasta");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn7 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idTipoCatalogo");
			Infragistics.Win.Appearance appearance12 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance13 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand2 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn11 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCatalogo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn12 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Marca");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn13 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Desde");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn14 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Hasta");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn15 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("EstadoCatalogo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn16 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Codigo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn17 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Articulo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn18 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Pagina");
			Infragistics.Win.Appearance appearance14 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn19 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Cantidad");
			Infragistics.Win.Appearance appearance15 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn20 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("CostoUltimo");
			Infragistics.Win.Appearance appearance16 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn21 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("UtilidadEfectivo");
			Infragistics.Win.Appearance appearance17 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn22 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("UtilidadTC");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn23 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Precio3");
			Infragistics.Win.Appearance appearance18 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn24 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("cuo6");
			Infragistics.Win.Appearance appearance19 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn25 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("cuo9");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn26 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("cuo12");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn27 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("cuo24");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn28 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("cuc3");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn29 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("cuc6");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn30 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("cuc9");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn31 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("cuc12");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn32 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("cuc18");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn33 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("entrada");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn34 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("notaportada");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn35 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("nota1");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn36 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("nota2");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn37 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("valorportada");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn38 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idTipoCatalogo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn39 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("CuotasE");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn40 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("CuotasC");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn41 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("cuotap", 0);
			Infragistics.Win.Appearance appearance20 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance21 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance22 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance23 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance24 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance25 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance26 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance27 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance28 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn8 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idCatalogo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn9 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Marca");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn10 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Desde");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn11 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Hasta");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn12 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("EstadoCatalogo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn13 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Codigo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn14 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Articulo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn15 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Pagina");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn16 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Cantidad");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn17 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("CostoUltimo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn18 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("UtilidadEfectivo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn19 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("UtilidadTC");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn20 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Precio3");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn21 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("cuo6");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn22 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("cuo9");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn23 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("cuo12");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn24 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("cuo24");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn25 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("cuc3");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn26 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("cuc6");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn27 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("cuc9");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn28 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("cuc12");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn29 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("cuc18");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn30 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("entrada");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn31 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("notaportada");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn32 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("nota1");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn33 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("nota2");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn34 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("valorportada");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn35 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idTipoCatalogo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn36 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("CuotasE");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn37 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("CuotasC");
			Infragistics.Win.Appearance appearance29 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance30 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance31 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance32 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance33 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance34 = new Infragistics.Win.Appearance();
			Infragistics.Win.ValueListItem valueListItem1 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.ValueListItem valueListItem2 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.ValueListItem valueListItem3 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.Appearance appearance35 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance36 = new Infragistics.Win.Appearance();
			Infragistics.Win.ValueListItem valueListItem4 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.ValueListItem valueListItem5 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.ValueListItem valueListItem6 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.Appearance appearance37 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinTabControl.UltraTab ultraTab1 = new Infragistics.Win.UltraWinTabControl.UltraTab();
			Infragistics.Win.UltraWinTabControl.UltraTab ultraTab2 = new Infragistics.Win.UltraWinTabControl.UltraTab();
			Infragistics.Win.UltraWinTabControl.UltraTab ultraTab3 = new Infragistics.Win.UltraWinTabControl.UltraTab();
			Infragistics.Win.Appearance appearance38 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance39 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton dateButton1 = new Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton();
			Infragistics.Win.Appearance appearance40 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton dateButton2 = new Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton();
			this.ultraTabPageControl9 = new Infragistics.Win.UltraWinTabControl.UltraTabPageControl();
			this.ultraGrid1 = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.ultraDataSource1 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.pgEntrega = new Infragistics.Win.UltraWinTabControl.UltraTabPageControl();
			this.ultraGrid2 = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.ultraDataSource2 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.ultraTabPageControl1 = new Infragistics.Win.UltraWinTabControl.UltraTabPageControl();
			this.txtEstadoCatalogo = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.btnColor = new System.Windows.Forms.Button();
			this.btnFuente = new System.Windows.Forms.Button();
			this.txtidCatalogo = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.label1 = new System.Windows.Forms.Label();
			this.btnFinalizar = new System.Windows.Forms.Button();
			this.btnImagen = new System.Windows.Forms.Button();
			this.pbImagen = new System.Windows.Forms.PictureBox();
			this.btnGuardar = new System.Windows.Forms.Button();
			this.rtbDescripcion = new System.Windows.Forms.RichTextBox();
			this.txtPromocion = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.lblCliente = new System.Windows.Forms.Label();
			this.lblImagen = new System.Windows.Forms.Label();
			this.lblEstado = new System.Windows.Forms.Label();
			this.txtObservaciones = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.txtBusqueda = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.groupBox1 = new System.Windows.Forms.GroupBox();
			this.ultraGridExcelExporter1 = new Infragistics.Win.UltraWinGrid.ExcelExport.UltraGridExcelExporter();
			this.optEstado = new Infragistics.Win.UltraWinEditors.UltraOptionSet();
			this.lblContador = new System.Windows.Forms.Label();
			this.btnExcel = new System.Windows.Forms.Button();
			this.btnVer = new System.Windows.Forms.Button();
			this.optEstadoCaja = new Infragistics.Win.UltraWinEditors.UltraOptionSet();
			this.utcEstadoCajas = new Infragistics.Win.UltraWinTabControl.UltraTabControl();
			this.ultraTabSharedControlsPage1 = new Infragistics.Win.UltraWinTabControl.UltraTabSharedControlsPage();
			this.label17 = new System.Windows.Forms.Label();
			this.label18 = new System.Windows.Forms.Label();
			this.dtHasta = new Infragistics.Win.UltraWinSchedule.UltraCalendarCombo();
			this.dtDesde = new Infragistics.Win.UltraWinSchedule.UltraCalendarCombo();
			this.label36 = new System.Windows.Forms.Label();
			this.cdsSeteoF = new C1.Data.C1DataSet();
			this.saveFileDialog1 = new System.Windows.Forms.SaveFileDialog();
			this.mainMenu1 = new System.Windows.Forms.MainMenu();
			this.openFileDialog1 = new System.Windows.Forms.OpenFileDialog();
			this.ultraTabPageControl9.SuspendLayout();
			((System.ComponentModel.ISupportInitialize)(this.ultraGrid1)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource1)).BeginInit();
			this.pgEntrega.SuspendLayout();
			((System.ComponentModel.ISupportInitialize)(this.ultraGrid2)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource2)).BeginInit();
			this.ultraTabPageControl1.SuspendLayout();
			((System.ComponentModel.ISupportInitialize)(this.txtEstadoCatalogo)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtidCatalogo)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtPromocion)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtObservaciones)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtBusqueda)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.optEstado)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.optEstadoCaja)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.utcEstadoCajas)).BeginInit();
			this.utcEstadoCajas.SuspendLayout();
			((System.ComponentModel.ISupportInitialize)(this.dtHasta)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.dtDesde)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).BeginInit();
			this.SuspendLayout();
			// 
			// ultraTabPageControl9
			// 
			this.ultraTabPageControl9.Controls.Add(this.ultraGrid1);
			this.ultraTabPageControl9.Location = new System.Drawing.Point(1, 20);
			this.ultraTabPageControl9.Name = "ultraTabPageControl9";
			this.ultraTabPageControl9.Size = new System.Drawing.Size(1430, 335);
			// 
			// ultraGrid1
			// 
			this.ultraGrid1.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left)));
			this.ultraGrid1.Cursor = System.Windows.Forms.Cursors.Default;
			this.ultraGrid1.DataSource = this.ultraDataSource1;
			appearance1.BackColor = System.Drawing.Color.White;
			this.ultraGrid1.DisplayLayout.Appearance = appearance1;
			this.ultraGrid1.DisplayLayout.AutoFitColumns = true;
			ultraGridColumn1.Header.VisiblePosition = 4;
			ultraGridColumn1.Hidden = true;
			ultraGridColumn1.Width = 63;
			ultraGridColumn2.Header.VisiblePosition = 3;
			ultraGridColumn2.Width = 210;
			ultraGridColumn3.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			ultraGridColumn3.Header.VisiblePosition = 7;
			ultraGridColumn3.Width = 151;
			ultraGridColumn4.Header.VisiblePosition = 5;
			ultraGridColumn4.Width = 210;
			ultraGridColumn5.Header.VisiblePosition = 6;
			ultraGridColumn5.Width = 192;
			ultraGridColumn6.Header.VisiblePosition = 8;
			ultraGridColumn6.Width = 192;
			ultraGridColumn7.Header.VisiblePosition = 9;
			ultraGridColumn7.Hidden = true;
			ultraGridColumn7.Width = 86;
			ultraGridColumn8.ButtonDisplayStyle = Infragistics.Win.UltraWinGrid.ButtonDisplayStyle.Always;
			appearance2.Image = ((object)(resources.GetObject("appearance2.Image")));
			ultraGridColumn8.CellButtonAppearance = appearance2;
			ultraGridColumn8.Header.Caption = "...";
			ultraGridColumn8.Header.VisiblePosition = 1;
			ultraGridColumn8.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.Button;
			ultraGridColumn8.Width = 18;
			ultraGridColumn9.ButtonDisplayStyle = Infragistics.Win.UltraWinGrid.ButtonDisplayStyle.Always;
			appearance3.Image = ((object)(resources.GetObject("appearance3.Image")));
			ultraGridColumn9.CellButtonAppearance = appearance3;
			ultraGridColumn9.CellDisplayStyle = Infragistics.Win.UltraWinGrid.CellDisplayStyle.PlainText;
			ultraGridColumn9.Header.Caption = "...";
			ultraGridColumn9.Header.VisiblePosition = 2;
			ultraGridColumn9.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.Button;
			ultraGridColumn9.Width = 24;
			ultraGridColumn10.ButtonDisplayStyle = Infragistics.Win.UltraWinGrid.ButtonDisplayStyle.Always;
			appearance4.Image = ((object)(resources.GetObject("appearance4.Image")));
			ultraGridColumn10.CellButtonAppearance = appearance4;
			ultraGridColumn10.Header.Caption = "Ver";
			ultraGridColumn10.Header.VisiblePosition = 0;
			ultraGridColumn10.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.Button;
			ultraGridColumn10.Width = 14;
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
			ultraGridBand1.SummaryFooterCaption = "";
			this.ultraGrid1.DisplayLayout.BandsSerializer.Add(ultraGridBand1);
			this.ultraGrid1.DisplayLayout.GroupByBox.Hidden = true;
			appearance5.ForeColor = System.Drawing.Color.Black;
			appearance5.ForeColorDisabled = System.Drawing.Color.Black;
			this.ultraGrid1.DisplayLayout.Override.ActiveRowAppearance = appearance5;
			this.ultraGrid1.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.False;
			this.ultraGrid1.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.False;
			appearance6.BackColor = System.Drawing.Color.Transparent;
			this.ultraGrid1.DisplayLayout.Override.CardAreaAppearance = appearance6;
			appearance7.ForeColor = System.Drawing.Color.Black;
			appearance7.ForeColorDisabled = System.Drawing.Color.Black;
			this.ultraGrid1.DisplayLayout.Override.CellAppearance = appearance7;
			appearance8.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance8.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance8.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance8.FontData.BoldAsString = "True";
			appearance8.FontData.Name = "Arial";
			appearance8.FontData.SizeInPoints = 8.5F;
			appearance8.ForeColor = System.Drawing.Color.White;
			appearance8.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.ultraGrid1.DisplayLayout.Override.HeaderAppearance = appearance8;
			appearance9.BackColor = System.Drawing.SystemColors.ControlLight;
			appearance9.BackColor2 = System.Drawing.Color.Gainsboro;
			appearance9.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ultraGrid1.DisplayLayout.Override.RowAlternateAppearance = appearance9;
			appearance10.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance10.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance10.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ultraGrid1.DisplayLayout.Override.RowSelectorAppearance = appearance10;
			appearance11.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance11.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance11.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ultraGrid1.DisplayLayout.Override.SelectedRowAppearance = appearance11;
			this.ultraGrid1.DisplayLayout.ViewStyleBand = Infragistics.Win.UltraWinGrid.ViewStyleBand.OutlookGroupBy;
			this.ultraGrid1.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.ultraGrid1.Location = new System.Drawing.Point(8, 8);
			this.ultraGrid1.Name = "ultraGrid1";
			this.ultraGrid1.Size = new System.Drawing.Size(1032, 324);
			this.ultraGrid1.TabIndex = 11;
			this.ultraGrid1.MouseMove += new System.Windows.Forms.MouseEventHandler(this.ultraGrid1_MouseMove);
			this.ultraGrid1.ClickCellButton += new Infragistics.Win.UltraWinGrid.CellEventHandler(this.ultraGrid1_ClickCellButton);
			this.ultraGrid1.InitializeLayout += new Infragistics.Win.UltraWinGrid.InitializeLayoutEventHandler(this.ultraGrid1_InitializeLayout);
			// 
			// ultraDataSource1
			// 
			ultraDataColumn1.DataType = typeof(int);
			ultraDataColumn3.DataType = typeof(int);
			ultraDataColumn5.DataType = typeof(System.DateTime);
			ultraDataColumn6.DataType = typeof(System.DateTime);
			ultraDataColumn7.DataType = typeof(int);
			this.ultraDataSource1.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn1,
																																 ultraDataColumn2,
																																 ultraDataColumn3,
																																 ultraDataColumn4,
																																 ultraDataColumn5,
																																 ultraDataColumn6,
																																 ultraDataColumn7});
			// 
			// pgEntrega
			// 
			this.pgEntrega.Controls.Add(this.ultraGrid2);
			this.pgEntrega.Location = new System.Drawing.Point(-10000, -10000);
			this.pgEntrega.Name = "pgEntrega";
			this.pgEntrega.Size = new System.Drawing.Size(1430, 335);
			// 
			// ultraGrid2
			// 
			this.ultraGrid2.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left)));
			this.ultraGrid2.Cursor = System.Windows.Forms.Cursors.Default;
			this.ultraGrid2.DataSource = this.ultraDataSource2;
			appearance12.TextHAlign = Infragistics.Win.HAlign.Right;
			this.ultraGrid2.DisplayLayout.AddNewBox.Appearance = appearance12;
			appearance13.BackColor = System.Drawing.Color.White;
			appearance13.TextHAlign = Infragistics.Win.HAlign.Right;
			this.ultraGrid2.DisplayLayout.Appearance = appearance13;
			ultraGridColumn11.Header.VisiblePosition = 0;
			ultraGridColumn11.Hidden = true;
			ultraGridColumn12.Header.VisiblePosition = 1;
			ultraGridColumn13.Header.VisiblePosition = 26;
			ultraGridColumn14.Header.VisiblePosition = 27;
			ultraGridColumn15.Header.VisiblePosition = 29;
			ultraGridColumn16.Header.VisiblePosition = 2;
			ultraGridColumn17.Header.VisiblePosition = 3;
			appearance14.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn18.CellAppearance = appearance14;
			ultraGridColumn18.CellDisplayStyle = Infragistics.Win.UltraWinGrid.CellDisplayStyle.FullEditorDisplay;
			ultraGridColumn18.Header.VisiblePosition = 28;
			ultraGridColumn18.UseEditorMaskSettings = true;
			appearance15.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn19.CellAppearance = appearance15;
			ultraGridColumn19.Header.VisiblePosition = 5;
			appearance16.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn20.CellAppearance = appearance16;
			ultraGridColumn20.CellDisplayStyle = Infragistics.Win.UltraWinGrid.CellDisplayStyle.FullEditorDisplay;
			ultraGridColumn20.Header.Caption = "Costo C/IVA";
			ultraGridColumn20.Header.VisiblePosition = 4;
			ultraGridColumn20.Hidden = true;
			appearance17.ForeColor = System.Drawing.Color.Black;
			appearance17.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn21.CellAppearance = appearance17;
			ultraGridColumn21.Header.Caption = "Utilidad Efectivo";
			ultraGridColumn21.Header.VisiblePosition = 6;
			ultraGridColumn21.Hidden = true;
			ultraGridColumn22.Header.Caption = "Utilidad TC";
			ultraGridColumn22.Header.VisiblePosition = 7;
			ultraGridColumn22.Hidden = true;
			appearance18.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn23.CellAppearance = appearance18;
			ultraGridColumn23.CellDisplayStyle = Infragistics.Win.UltraWinGrid.CellDisplayStyle.FullEditorDisplay;
			ultraGridColumn23.Header.Caption = "Precio Oferta";
			ultraGridColumn23.Header.VisiblePosition = 8;
			appearance19.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn24.CellAppearance = appearance19;
			ultraGridColumn24.Header.Caption = "6 Cuotas con interes";
			ultraGridColumn24.Header.VisiblePosition = 12;
			ultraGridColumn24.Hidden = true;
			ultraGridColumn25.Header.Caption = "9 Cuotas con interes";
			ultraGridColumn25.Header.VisiblePosition = 13;
			ultraGridColumn25.Hidden = true;
			ultraGridColumn26.Header.Caption = "12 Cuotas con interes";
			ultraGridColumn26.Header.VisiblePosition = 14;
			ultraGridColumn26.Hidden = true;
			ultraGridColumn27.Header.Caption = "24 Cuotas con interes";
			ultraGridColumn27.Header.VisiblePosition = 15;
			ultraGridColumn27.Hidden = true;
			ultraGridColumn28.Header.Caption = "6 Cuotas ";
			ultraGridColumn28.Header.VisiblePosition = 17;
			ultraGridColumn28.Hidden = true;
			ultraGridColumn29.Header.Caption = "6 Cuotas ";
			ultraGridColumn29.Header.VisiblePosition = 18;
			ultraGridColumn29.Hidden = true;
			ultraGridColumn30.Header.Caption = "9 Cuotas ";
			ultraGridColumn30.Header.VisiblePosition = 19;
			ultraGridColumn30.Hidden = true;
			ultraGridColumn31.Header.Caption = "12 Cuotas ";
			ultraGridColumn31.Header.VisiblePosition = 20;
			ultraGridColumn31.Hidden = true;
			ultraGridColumn32.Header.Caption = "18 Cuotas ";
			ultraGridColumn32.Header.VisiblePosition = 21;
			ultraGridColumn32.Hidden = true;
			ultraGridColumn33.Header.Caption = "Entrada";
			ultraGridColumn33.Header.VisiblePosition = 16;
			ultraGridColumn33.Hidden = true;
			ultraGridColumn34.Header.Caption = "Portada";
			ultraGridColumn34.Header.VisiblePosition = 22;
			ultraGridColumn35.Header.Caption = "Adicionales/ Obsequios";
			ultraGridColumn35.Header.Enabled = false;
			ultraGridColumn35.Header.VisiblePosition = 24;
			ultraGridColumn36.Header.Caption = "Adicional";
			ultraGridColumn36.Header.VisiblePosition = 25;
			ultraGridColumn37.Header.Caption = "Valor Portada";
			ultraGridColumn37.Header.VisiblePosition = 23;
			ultraGridColumn38.Header.VisiblePosition = 30;
			ultraGridColumn38.Hidden = true;
			ultraGridColumn39.Header.Caption = "Valor Cuota";
			ultraGridColumn39.Header.VisiblePosition = 10;
			ultraGridColumn39.Hidden = true;
			ultraGridColumn40.Header.Caption = "Valor Cuota";
			ultraGridColumn40.Header.VisiblePosition = 11;
			ultraGridColumn40.Hidden = true;
			appearance20.BackColor = System.Drawing.Color.LimeGreen;
			ultraGridColumn41.CellAppearance = appearance20;
			ultraGridColumn41.DataType = typeof(System.Double);
			ultraGridColumn41.Header.Caption = "Cuota Publicada";
			ultraGridColumn41.Header.VisiblePosition = 9;
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
																										 ultraGridColumn41});
			appearance21.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridBand2.Override.EditCellAppearance = appearance21;
			ultraGridBand2.SummaryFooterCaption = "";
			this.ultraGrid2.DisplayLayout.BandsSerializer.Add(ultraGridBand2);
			this.ultraGrid2.DisplayLayout.GroupByBox.Hidden = true;
			appearance22.ForeColor = System.Drawing.Color.Black;
			appearance22.ForeColorDisabled = System.Drawing.Color.Black;
			this.ultraGrid2.DisplayLayout.Override.ActiveRowAppearance = appearance22;
			this.ultraGrid2.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.False;
			this.ultraGrid2.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.True;
			appearance23.BackColor = System.Drawing.Color.Transparent;
			this.ultraGrid2.DisplayLayout.Override.CardAreaAppearance = appearance23;
			appearance24.ForeColor = System.Drawing.Color.Black;
			appearance24.ForeColorDisabled = System.Drawing.Color.Black;
			this.ultraGrid2.DisplayLayout.Override.CellAppearance = appearance24;
			appearance25.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance25.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance25.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance25.FontData.BoldAsString = "True";
			appearance25.FontData.Name = "Arial";
			appearance25.FontData.SizeInPoints = 8.25F;
			appearance25.ForeColor = System.Drawing.Color.White;
			appearance25.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.ultraGrid2.DisplayLayout.Override.HeaderAppearance = appearance25;
			appearance26.BackColor = System.Drawing.SystemColors.ControlLight;
			appearance26.BackColor2 = System.Drawing.Color.Gainsboro;
			appearance26.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ultraGrid2.DisplayLayout.Override.RowAlternateAppearance = appearance26;
			appearance27.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance27.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance27.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ultraGrid2.DisplayLayout.Override.RowSelectorAppearance = appearance27;
			appearance28.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance28.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance28.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ultraGrid2.DisplayLayout.Override.SelectedRowAppearance = appearance28;
			this.ultraGrid2.DisplayLayout.ViewStyleBand = Infragistics.Win.UltraWinGrid.ViewStyleBand.OutlookGroupBy;
			this.ultraGrid2.Font = new System.Drawing.Font("Tahoma", 8F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.ultraGrid2.Location = new System.Drawing.Point(7, 5);
			this.ultraGrid2.Name = "ultraGrid2";
			this.ultraGrid2.Size = new System.Drawing.Size(1313, 324);
			this.ultraGrid2.TabIndex = 12;
			this.ultraGrid2.InitializeLayout += new Infragistics.Win.UltraWinGrid.InitializeLayoutEventHandler(this.ultraGrid2_InitializeLayout);
			// 
			// ultraDataSource2
			// 
			ultraDataColumn8.DataType = typeof(int);
			ultraDataColumn10.DataType = typeof(System.DateTime);
			ultraDataColumn10.DefaultValue = new System.DateTime(((long)(0)));
			ultraDataColumn11.DataType = typeof(System.DateTime);
			ultraDataColumn11.DefaultValue = new System.DateTime(((long)(0)));
			ultraDataColumn15.DataType = typeof(int);
			ultraDataColumn16.DataType = typeof(System.Double);
			ultraDataColumn17.DataType = typeof(System.Double);
			ultraDataColumn18.DataType = typeof(System.Double);
			ultraDataColumn19.DataType = typeof(System.Double);
			ultraDataColumn20.DataType = typeof(System.Double);
			ultraDataColumn21.DataType = typeof(System.Double);
			ultraDataColumn22.DataType = typeof(System.Double);
			ultraDataColumn23.DataType = typeof(System.Double);
			ultraDataColumn24.DataType = typeof(System.Double);
			ultraDataColumn25.DataType = typeof(System.Double);
			ultraDataColumn26.DataType = typeof(System.Double);
			ultraDataColumn27.DataType = typeof(System.Double);
			ultraDataColumn28.DataType = typeof(System.Double);
			ultraDataColumn29.DataType = typeof(System.Double);
			ultraDataColumn30.DataType = typeof(System.Double);
			ultraDataColumn34.DataType = typeof(System.Double);
			ultraDataColumn35.DataType = typeof(int);
			ultraDataColumn36.DataType = typeof(System.Double);
			ultraDataColumn37.DataType = typeof(System.Double);
			this.ultraDataSource2.Band.Columns.AddRange(new object[] {
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
																																 ultraDataColumn24,
																																 ultraDataColumn25,
																																 ultraDataColumn26,
																																 ultraDataColumn27,
																																 ultraDataColumn28,
																																 ultraDataColumn29,
																																 ultraDataColumn30,
																																 ultraDataColumn31,
																																 ultraDataColumn32,
																																 ultraDataColumn33,
																																 ultraDataColumn34,
																																 ultraDataColumn35,
																																 ultraDataColumn36,
																																 ultraDataColumn37});
			// 
			// ultraTabPageControl1
			// 
			this.ultraTabPageControl1.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D;
			this.ultraTabPageControl1.Controls.Add(this.txtEstadoCatalogo);
			this.ultraTabPageControl1.Controls.Add(this.btnColor);
			this.ultraTabPageControl1.Controls.Add(this.btnFuente);
			this.ultraTabPageControl1.Controls.Add(this.txtidCatalogo);
			this.ultraTabPageControl1.Controls.Add(this.label1);
			this.ultraTabPageControl1.Controls.Add(this.btnFinalizar);
			this.ultraTabPageControl1.Controls.Add(this.btnImagen);
			this.ultraTabPageControl1.Controls.Add(this.pbImagen);
			this.ultraTabPageControl1.Controls.Add(this.btnGuardar);
			this.ultraTabPageControl1.Controls.Add(this.rtbDescripcion);
			this.ultraTabPageControl1.Controls.Add(this.txtPromocion);
			this.ultraTabPageControl1.Controls.Add(this.lblCliente);
			this.ultraTabPageControl1.Controls.Add(this.lblImagen);
			this.ultraTabPageControl1.Controls.Add(this.lblEstado);
			this.ultraTabPageControl1.Location = new System.Drawing.Point(-10000, -10000);
			this.ultraTabPageControl1.Name = "ultraTabPageControl1";
			this.ultraTabPageControl1.Size = new System.Drawing.Size(1430, 335);
			this.ultraTabPageControl1.Paint += new System.Windows.Forms.PaintEventHandler(this.ultraTabPageControl1_Paint);
			// 
			// txtEstadoCatalogo
			// 
			appearance29.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtEstadoCatalogo.Appearance = appearance29;
			this.txtEstadoCatalogo.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtEstadoCatalogo.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtEstadoCatalogo.Enabled = false;
			this.txtEstadoCatalogo.Location = new System.Drawing.Point(496, 80);
			this.txtEstadoCatalogo.MaxLength = 300;
			this.txtEstadoCatalogo.Multiline = true;
			this.txtEstadoCatalogo.Name = "txtEstadoCatalogo";
			this.txtEstadoCatalogo.Size = new System.Drawing.Size(40, 24);
			this.txtEstadoCatalogo.TabIndex = 821;
			this.txtEstadoCatalogo.Visible = false;
			// 
			// btnColor
			// 
			this.btnColor.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btnColor.CausesValidation = false;
			this.btnColor.Enabled = false;
			this.btnColor.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnColor.Location = new System.Drawing.Point(488, 8);
			this.btnColor.Name = "btnColor";
			this.btnColor.Size = new System.Drawing.Size(56, 23);
			this.btnColor.TabIndex = 820;
			this.btnColor.Text = "&Color";
			this.btnColor.Click += new System.EventHandler(this.btnColor_Click);
			// 
			// btnFuente
			// 
			this.btnFuente.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btnFuente.CausesValidation = false;
			this.btnFuente.Enabled = false;
			this.btnFuente.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnFuente.Location = new System.Drawing.Point(424, 8);
			this.btnFuente.Name = "btnFuente";
			this.btnFuente.Size = new System.Drawing.Size(56, 23);
			this.btnFuente.TabIndex = 819;
			this.btnFuente.Text = "&Fuente";
			this.btnFuente.Click += new System.EventHandler(this.btnFuente_Click);
			// 
			// txtidCatalogo
			// 
			this.txtidCatalogo.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			this.txtidCatalogo.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtidCatalogo.Enabled = false;
			this.txtidCatalogo.FormatString = "#,##0.00";
			this.txtidCatalogo.Location = new System.Drawing.Point(504, 48);
			this.txtidCatalogo.Name = "txtidCatalogo";
			this.txtidCatalogo.PromptChar = ' ';
			this.txtidCatalogo.Size = new System.Drawing.Size(40, 21);
			this.txtidCatalogo.TabIndex = 818;
			this.txtidCatalogo.Visible = false;
			// 
			// label1
			// 
			this.label1.AutoSize = true;
			this.label1.Enabled = false;
			this.label1.FlatStyle = System.Windows.Forms.FlatStyle.System;
			this.label1.Location = new System.Drawing.Point(16, 56);
			this.label1.Name = "label1";
			this.label1.Size = new System.Drawing.Size(68, 16);
			this.label1.TabIndex = 817;
			this.label1.Text = "Observacion";
			this.label1.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// btnFinalizar
			// 
			this.btnFinalizar.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btnFinalizar.CausesValidation = false;
			this.btnFinalizar.Enabled = false;
			this.btnFinalizar.Image = ((System.Drawing.Image)(resources.GetObject("btnFinalizar.Image")));
			this.btnFinalizar.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnFinalizar.Location = new System.Drawing.Point(240, 232);
			this.btnFinalizar.Name = "btnFinalizar";
			this.btnFinalizar.Size = new System.Drawing.Size(104, 23);
			this.btnFinalizar.TabIndex = 816;
			this.btnFinalizar.Text = "&Autorizar";
			this.btnFinalizar.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnFinalizar.Click += new System.EventHandler(this.btnFinalizar_Click);
			// 
			// btnImagen
			// 
			this.btnImagen.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btnImagen.CausesValidation = false;
			this.btnImagen.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnImagen.Location = new System.Drawing.Point(16, 232);
			this.btnImagen.Name = "btnImagen";
			this.btnImagen.Size = new System.Drawing.Size(104, 23);
			this.btnImagen.TabIndex = 226;
			this.btnImagen.Text = "&Cargar Imagen";
			this.btnImagen.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnImagen.Click += new System.EventHandler(this.btnImagen_Click);
			// 
			// pbImagen
			// 
			this.pbImagen.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
			this.pbImagen.Enabled = false;
			this.pbImagen.Location = new System.Drawing.Point(552, 8);
			this.pbImagen.Name = "pbImagen";
			this.pbImagen.Size = new System.Drawing.Size(368, 312);
			this.pbImagen.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage;
			this.pbImagen.TabIndex = 224;
			this.pbImagen.TabStop = false;
			// 
			// btnGuardar
			// 
			this.btnGuardar.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btnGuardar.Enabled = false;
			this.btnGuardar.Image = ((System.Drawing.Image)(resources.GetObject("btnGuardar.Image")));
			this.btnGuardar.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnGuardar.Location = new System.Drawing.Point(128, 232);
			this.btnGuardar.Name = "btnGuardar";
			this.btnGuardar.Size = new System.Drawing.Size(104, 23);
			this.btnGuardar.TabIndex = 815;
			this.btnGuardar.Text = "&Guardar";
			this.btnGuardar.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnGuardar.Click += new System.EventHandler(this.btnGuardar_Click);
			// 
			// rtbDescripcion
			// 
			this.rtbDescripcion.AcceptsTab = true;
			this.rtbDescripcion.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
			this.rtbDescripcion.Location = new System.Drawing.Point(16, 80);
			this.rtbDescripcion.Name = "rtbDescripcion";
			this.rtbDescripcion.Size = new System.Drawing.Size(456, 120);
			this.rtbDescripcion.TabIndex = 815;
			this.rtbDescripcion.Text = "";
			// 
			// txtPromocion
			// 
			appearance30.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtPromocion.Appearance = appearance30;
			this.txtPromocion.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtPromocion.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtPromocion.Enabled = false;
			this.txtPromocion.Location = new System.Drawing.Point(80, 16);
			this.txtPromocion.Name = "txtPromocion";
			this.txtPromocion.Size = new System.Drawing.Size(336, 21);
			this.txtPromocion.TabIndex = 815;
			// 
			// lblCliente
			// 
			this.lblCliente.AutoSize = true;
			this.lblCliente.Enabled = false;
			this.lblCliente.FlatStyle = System.Windows.Forms.FlatStyle.System;
			this.lblCliente.Location = new System.Drawing.Point(16, 16);
			this.lblCliente.Name = "lblCliente";
			this.lblCliente.Size = new System.Drawing.Size(58, 16);
			this.lblCliente.TabIndex = 816;
			this.lblCliente.Text = "Promoción";
			this.lblCliente.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// lblImagen
			// 
			this.lblImagen.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			this.lblImagen.AutoSize = true;
			this.lblImagen.Font = new System.Drawing.Font("Microsoft Sans Serif", 15F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.lblImagen.ForeColor = System.Drawing.Color.Firebrick;
			this.lblImagen.Location = new System.Drawing.Point(736, 296);
			this.lblImagen.Name = "lblImagen";
			this.lblImagen.Size = new System.Drawing.Size(0, 26);
			this.lblImagen.TabIndex = 815;
			this.lblImagen.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
			this.lblImagen.Visible = false;
			// 
			// lblEstado
			// 
			this.lblEstado.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			this.lblEstado.AutoSize = true;
			this.lblEstado.Font = new System.Drawing.Font("Microsoft Sans Serif", 15F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.lblEstado.ForeColor = System.Drawing.Color.Firebrick;
			this.lblEstado.Location = new System.Drawing.Point(416, 224);
			this.lblEstado.Name = "lblEstado";
			this.lblEstado.Size = new System.Drawing.Size(0, 26);
			this.lblEstado.TabIndex = 815;
			this.lblEstado.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
			// 
			// txtObservaciones
			// 
			appearance31.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtObservaciones.Appearance = appearance31;
			this.txtObservaciones.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtObservaciones.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtObservaciones.Enabled = false;
			this.txtObservaciones.Location = new System.Drawing.Point(712, 24);
			this.txtObservaciones.MaxLength = 300;
			this.txtObservaciones.Multiline = true;
			this.txtObservaciones.Name = "txtObservaciones";
			this.txtObservaciones.Size = new System.Drawing.Size(24, 24);
			this.txtObservaciones.TabIndex = 814;
			this.txtObservaciones.Visible = false;
			// 
			// txtBusqueda
			// 
			appearance32.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtBusqueda.Appearance = appearance32;
			this.txtBusqueda.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtBusqueda.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtBusqueda.Location = new System.Drawing.Point(80, 56);
			this.txtBusqueda.Name = "txtBusqueda";
			this.txtBusqueda.Size = new System.Drawing.Size(208, 21);
			this.txtBusqueda.TabIndex = 813;
			// 
			// groupBox1
			// 
			this.groupBox1.Location = new System.Drawing.Point(8, 88);
			this.groupBox1.Name = "groupBox1";
			this.groupBox1.Size = new System.Drawing.Size(1328, 8);
			this.groupBox1.TabIndex = 812;
			this.groupBox1.TabStop = false;
			// 
			// optEstado
			// 
			appearance33.ForeColorDisabled = System.Drawing.Color.Black;
			this.optEstado.Appearance = appearance33;
			this.optEstado.BorderStyle = Infragistics.Win.UIElementBorderStyle.Solid;
			this.optEstado.ItemAppearance = appearance34;
			this.optEstado.ItemOrigin = new System.Drawing.Point(10, 0);
			valueListItem1.DataValue = 0;
			valueListItem1.DisplayText = "Todos";
			valueListItem2.DataValue = 1;
			valueListItem2.DisplayText = "Inicial";
			valueListItem3.DataValue = 2;
			valueListItem3.DisplayText = "Aprobado";
			this.optEstado.Items.Add(valueListItem1);
			this.optEstado.Items.Add(valueListItem2);
			this.optEstado.Items.Add(valueListItem3);
			this.optEstado.ItemSpacingVertical = 10;
			this.optEstado.Location = new System.Drawing.Point(656, 24);
			this.optEstado.Name = "optEstado";
			this.optEstado.Size = new System.Drawing.Size(8, 24);
			this.optEstado.TabIndex = 810;
			this.optEstado.Visible = false;
			// 
			// lblContador
			// 
			this.lblContador.AutoSize = true;
			this.lblContador.Font = new System.Drawing.Font("Tahoma", 10F);
			this.lblContador.ForeColor = System.Drawing.Color.Firebrick;
			this.lblContador.Location = new System.Drawing.Point(592, 56);
			this.lblContador.Name = "lblContador";
			this.lblContador.Size = new System.Drawing.Size(0, 20);
			this.lblContador.TabIndex = 811;
			this.lblContador.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
			// 
			// btnExcel
			// 
			this.btnExcel.Image = ((System.Drawing.Image)(resources.GetObject("btnExcel.Image")));
			this.btnExcel.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnExcel.Location = new System.Drawing.Point(792, 8);
			this.btnExcel.Name = "btnExcel";
			this.btnExcel.Size = new System.Drawing.Size(88, 24);
			this.btnExcel.TabIndex = 809;
			this.btnExcel.Text = "Excel";
			this.btnExcel.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnExcel.Click += new System.EventHandler(this.btnExcel_Click);
			// 
			// btnVer
			// 
			this.btnVer.Image = ((System.Drawing.Image)(resources.GetObject("btnVer.Image")));
			this.btnVer.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnVer.Location = new System.Drawing.Point(312, 48);
			this.btnVer.Name = "btnVer";
			this.btnVer.Size = new System.Drawing.Size(75, 24);
			this.btnVer.TabIndex = 808;
			this.btnVer.Text = "Ver";
			this.btnVer.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnVer.Click += new System.EventHandler(this.btnVer_Click);
			// 
			// optEstadoCaja
			// 
			appearance35.ForeColorDisabled = System.Drawing.Color.Black;
			this.optEstadoCaja.Appearance = appearance35;
			this.optEstadoCaja.BorderStyle = Infragistics.Win.UIElementBorderStyle.Solid;
			this.optEstadoCaja.CheckedIndex = 0;
			this.optEstadoCaja.ItemAppearance = appearance36;
			this.optEstadoCaja.ItemOrigin = new System.Drawing.Point(10, 0);
			valueListItem4.DataValue = 7;
			valueListItem4.DisplayText = "Todos";
			valueListItem5.DataValue = 1;
			valueListItem5.DisplayText = "Aprobado";
			valueListItem6.DataValue = 2;
			valueListItem6.DisplayText = "Cerrado";
			this.optEstadoCaja.Items.Add(valueListItem4);
			this.optEstadoCaja.Items.Add(valueListItem5);
			this.optEstadoCaja.Items.Add(valueListItem6);
			this.optEstadoCaja.ItemSpacingVertical = 10;
			this.optEstadoCaja.Location = new System.Drawing.Point(424, 16);
			this.optEstadoCaja.Name = "optEstadoCaja";
			this.optEstadoCaja.Size = new System.Drawing.Size(224, 24);
			this.optEstadoCaja.TabIndex = 807;
			this.optEstadoCaja.Text = "Todos";
			// 
			// utcEstadoCajas
			// 
			this.utcEstadoCajas.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			appearance37.ForeColorDisabled = System.Drawing.Color.Black;
			this.utcEstadoCajas.Appearance = appearance37;
			this.utcEstadoCajas.Controls.Add(this.ultraTabSharedControlsPage1);
			this.utcEstadoCajas.Controls.Add(this.pgEntrega);
			this.utcEstadoCajas.Controls.Add(this.ultraTabPageControl9);
			this.utcEstadoCajas.Controls.Add(this.ultraTabPageControl1);
			this.utcEstadoCajas.Location = new System.Drawing.Point(8, 104);
			this.utcEstadoCajas.Name = "utcEstadoCajas";
			this.utcEstadoCajas.SharedControlsPage = this.ultraTabSharedControlsPage1;
			this.utcEstadoCajas.Size = new System.Drawing.Size(1432, 356);
			this.utcEstadoCajas.Style = Infragistics.Win.UltraWinTabControl.UltraTabControlStyle.Excel;
			this.utcEstadoCajas.TabIndex = 806;
			ultraTab1.TabPage = this.ultraTabPageControl9;
			ultraTab1.Text = "Catalogo";
			ultraTab2.TabPage = this.pgEntrega;
			ultraTab2.Text = "Detalle Catalogo";
			appearance38.BackColor = System.Drawing.Color.GhostWhite;
			appearance38.BackColor2 = System.Drawing.Color.MidnightBlue;
			appearance38.BackGradientStyle = Infragistics.Win.GradientStyle.Elliptical;
			ultraTab3.ClientAreaAppearance = appearance38;
			ultraTab3.TabPage = this.ultraTabPageControl1;
			ultraTab3.Text = "Imagen de Portada";
			this.utcEstadoCajas.Tabs.AddRange(new Infragistics.Win.UltraWinTabControl.UltraTab[] {
																																														 ultraTab1,
																																														 ultraTab2,
																																														 ultraTab3});
			this.utcEstadoCajas.ViewStyle = Infragistics.Win.UltraWinTabControl.ViewStyle.Office2003;
			// 
			// ultraTabSharedControlsPage1
			// 
			this.ultraTabSharedControlsPage1.Location = new System.Drawing.Point(-10000, -10000);
			this.ultraTabSharedControlsPage1.Name = "ultraTabSharedControlsPage1";
			this.ultraTabSharedControlsPage1.Size = new System.Drawing.Size(1430, 335);
			// 
			// label17
			// 
			this.label17.AutoSize = true;
			this.label17.Location = new System.Drawing.Point(216, 24);
			this.label17.Name = "label17";
			this.label17.Size = new System.Drawing.Size(34, 16);
			this.label17.TabIndex = 805;
			this.label17.Text = "Hasta";
			this.label17.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// label18
			// 
			this.label18.AutoSize = true;
			this.label18.Location = new System.Drawing.Point(24, 24);
			this.label18.Name = "label18";
			this.label18.Size = new System.Drawing.Size(37, 16);
			this.label18.TabIndex = 804;
			this.label18.Text = "Desde";
			this.label18.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// dtHasta
			// 
			appearance39.ForeColorDisabled = System.Drawing.Color.Black;
			this.dtHasta.Appearance = appearance39;
			this.dtHasta.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			dateButton1.Caption = "Today";
			this.dtHasta.DateButtons.Add(dateButton1);
			this.dtHasta.Format = "dd/MM/yyyy";
			this.dtHasta.Location = new System.Drawing.Point(272, 24);
			this.dtHasta.Name = "dtHasta";
			this.dtHasta.NonAutoSizeHeight = 23;
			this.dtHasta.Size = new System.Drawing.Size(112, 21);
			this.dtHasta.SpinButtonsVisible = true;
			this.dtHasta.TabIndex = 802;
			this.dtHasta.Value = ((object)(resources.GetObject("dtHasta.Value")));
			// 
			// dtDesde
			// 
			appearance40.ForeColorDisabled = System.Drawing.Color.Black;
			this.dtDesde.Appearance = appearance40;
			this.dtDesde.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			dateButton2.Caption = "Today";
			this.dtDesde.DateButtons.Add(dateButton2);
			this.dtDesde.Format = "dd/MM/yyyy";
			this.dtDesde.Location = new System.Drawing.Point(80, 24);
			this.dtDesde.Name = "dtDesde";
			this.dtDesde.NonAutoSizeHeight = 23;
			this.dtDesde.Size = new System.Drawing.Size(112, 21);
			this.dtDesde.SpinButtonsVisible = true;
			this.dtDesde.TabIndex = 801;
			this.dtDesde.Value = ((object)(resources.GetObject("dtDesde.Value")));
			// 
			// label36
			// 
			this.label36.AutoSize = true;
			this.label36.Location = new System.Drawing.Point(24, 56);
			this.label36.Name = "label36";
			this.label36.Size = new System.Drawing.Size(50, 16);
			this.label36.TabIndex = 803;
			this.label36.Text = "Catalogo";
			this.label36.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
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
			// frmCatalogoMarketing
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 13);
			this.ClientSize = new System.Drawing.Size(1352, 518);
			this.Controls.Add(this.btnVer);
			this.Controls.Add(this.optEstadoCaja);
			this.Controls.Add(this.utcEstadoCajas);
			this.Controls.Add(this.label17);
			this.Controls.Add(this.label18);
			this.Controls.Add(this.label36);
			this.Controls.Add(this.dtHasta);
			this.Controls.Add(this.dtDesde);
			this.Controls.Add(this.txtObservaciones);
			this.Controls.Add(this.txtBusqueda);
			this.Controls.Add(this.groupBox1);
			this.Controls.Add(this.optEstado);
			this.Controls.Add(this.lblContador);
			this.Controls.Add(this.btnExcel);
			this.Menu = this.mainMenu1;
			this.Name = "frmCatalogoMarketing";
			this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
			this.Text = "CatalogoMarketing";
			this.Load += new System.EventHandler(this.frmCatalogoMarketing_Load);
			this.ultraTabPageControl9.ResumeLayout(false);
			((System.ComponentModel.ISupportInitialize)(this.ultraGrid1)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource1)).EndInit();
			this.pgEntrega.ResumeLayout(false);
			((System.ComponentModel.ISupportInitialize)(this.ultraGrid2)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource2)).EndInit();
			this.ultraTabPageControl1.ResumeLayout(false);
			((System.ComponentModel.ISupportInitialize)(this.txtEstadoCatalogo)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtidCatalogo)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtPromocion)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtObservaciones)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtBusqueda)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.optEstado)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.optEstadoCaja)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.utcEstadoCajas)).EndInit();
			this.utcEstadoCajas.ResumeLayout(false);
			((System.ComponentModel.ISupportInitialize)(this.dtHasta)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.dtDesde)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).EndInit();
			this.ResumeLayout(false);

		}
		#endregion

		private void btnImagen_Click(object sender, System.EventArgs e)
		{
			this.btnGuardar.Enabled= true;
			this.txtPromocion.Enabled = true;
			this.rtbDescripcion.Enabled = true;
			this.openFileDialog1.CheckFileExists = true;
			this.openFileDialog1.ShowReadOnly = false;
			this.btnFuente.Enabled = true;
			this.btnColor.Enabled = true;
			this.openFileDialog1.Filter = "All Files|*.*|Bitmap Files (*)|*.bmp;*.gif;*.jpg";
			this.openFileDialog1.FilterIndex = 2;
			
			if (this.openFileDialog1.ShowDialog() == DialogResult.OK)
			{
				this.pbImagen.Image = Image.FromFile(this.openFileDialog1.FileName);
				this.lblImagen.Text = this.openFileDialog1.FileName;
			}			 
			
		}

		private void btnGuardar_Click(object sender, System.EventArgs e)
		{
			if (this.txtPromocion.Text.Length == 0)
			{
				MessageBox.Show("Escriba el Nombre de la Portada", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				this.txtPromocion.Focus();
			}
			else if (this.rtbDescripcion.Text.Length == 0)
			{
				MessageBox.Show("Escriba la Descripcion de la Portada", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				this.rtbDescripcion.Focus();
			}
			else if (this.pbImagen.Image == null)
			{
				MessageBox.Show("Seleccione una Imagen", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				this.txtPromocion.Focus();
			}
			else
			{				
						MemoryStream ms = new MemoryStream();
						FileStream fs = new FileStream(this.lblImagen.Text, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
 
						ms.SetLength(fs.Length);
						fs.Read(ms.GetBuffer(), 0, (int)fs.Length);
 
						byte[] arrImg = ms.GetBuffer();
						ms.Flush();
						fs.Close();

						int aa = 0;
						aa =(int)this.ultraGrid1.ActiveRow.Cells["idCatalogo"].Value;
				
//						string sSQL1 = string.Format("Update Catalogo set Descripcion = '{0}', NombrePortada = '{1}' where idCatalogo = {2} ", this.rtbDescripcion.Text.ToString(), this.txtPromocion.Text.ToString(), aa);
//						Funcion.EjecutaSQL(cdsSeteoF, sSQL1);
						
						string sSQL1 = string.Format("Update Catalogo set Imagen = @Imagen, Descripcion = @Descripcion, NombrePortada = @NombrePortada, EstadoCatalogo =@EstadoCatalogo  where idCatalogo = @idCatalogo ");
//						Funcion.EjecutaSQL(cdsSeteoF, sSQL1);

				
						C1.Data.SchemaObjects.Connection miConeccion = cdsSeteoF.Schema.Connections[0];
							miConeccion.Open();
						SqlCommand cmdGrabaFoto = new SqlCommand (sSQL1,	
							(SqlConnection) miConeccion.DbConnection);

						cmdGrabaFoto.Parameters.Add("@Imagen", SqlDbType.VarBinary).Value = arrImg;
						cmdGrabaFoto.Parameters.Add("@Descripcion", SqlDbType.VarChar).Value = this.rtbDescripcion.Text.ToString();
						cmdGrabaFoto.Parameters.Add("@NombrePortada", SqlDbType.VarChar).Value = this.txtPromocion.Text.ToString();			
						cmdGrabaFoto.Parameters.Add("@idCatalogo", SqlDbType.Int).Value = aa;
						cmdGrabaFoto.Parameters.Add("@EstadoCatalogo",SqlDbType.VarChar).Value= "4";
					cmdGrabaFoto.ExecuteNonQuery();
					
						this.txtPromocion.Enabled = false;
						this.rtbDescripcion.Enabled = false;
						this.pbImagen.Enabled = false;


						this.btnGuardar.Enabled = false;
						this.btnFinalizar.Enabled = true;			
						this.btnImagen.Enabled = false;
						
			}
		}

		private void btnVer_Click(object sender, System.EventArgs e)
		{
//			string estadop = "";
//			if((int)this.optEstadoCaja.Value == 0) estadop="AUTORIZADO";
//			else if((int)this.optEstadoCaja.Value == 1) estadop="APROBADO";
//			else if((int)this.optEstadoCaja.Value == 2) estadop="CANCELADO";

			string sSQL = string.Format("Exec CatalogoGeneral '{0}', '{1}', '{2}',{3}", 
				Convert.ToDateTime(this.dtDesde.Value).ToString("yyyyMMdd"), Convert.ToDateTime(this.dtHasta.Value).ToString("yyyyMMdd"), this.txtBusqueda.Text.ToString(),(int)this.optEstadoCaja.Value);
			FuncionesProcedimientos.dsGeneral(sSQL, this.ultraGrid1);			
		}

		private void frmCatalogoMarketing_Load(object sender, System.EventArgs e)
		{
			this.dtDesde.CalendarInfo.MinDate = DateTime.Parse("01/01/2015");
			this.dtDesde.Value = "01/" + DateTime.Now.Month.ToString() + "/" + DateTime.Now.Year.ToString();			
			this.dtHasta.Value = DateTime.Today;
		}

		private void ultraGrid1_InitializeLayout(object sender, Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs e)
		{
				
		}

		private void ultraGrid1_ClickCellButton(object sender, Infragistics.Win.UltraWinGrid.CellEventArgs e)
		{
		this.ultraGrid1.ActiveRow.Cells["Anular"].ToolTipText= "Cerrar_Catalogo";
			string sSQLc  = string.Format("Exec BuscarCPortada {0}", (int)this.ultraGrid1.ActiveRow.Cells["idCatalogo"].Value);
			SqlDataReader dr = Funcion.rEscalarSQL(cdsSeteoF, sSQLc, true);
//		this.ultraGrid2.DataSource = null;
			dr.Read();
			if (dr.HasRows)
			{
				this.txtidCatalogo.Value = dr.GetInt32(0);
				this.txtPromocion.Text = dr.GetString(1);
				this.rtbDescripcion.Text = dr.GetString(2);
				this.txtEstadoCatalogo.Text = dr.GetString(4);
					try
					{
						byte[] arrImg = (byte[])dr.GetValue(3);
						MemoryStream ms = new MemoryStream(arrImg);
						this.pbImagen.Image = System.Drawing.Image.FromStream(ms);
						ms.Close();
					}
					catch(Exception){
						MessageBox.Show("No existe imagen", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
						this.pbImagen.Image = null;
						this.txtPromocion.Text = "";
						this.rtbDescripcion.Text = "";
					}
					
			}
			dr.Close();
			this.utcEstadoCajas.Tabs[2].Active= true;
			
			string sSQL  = string.Format("Exec ReporteCatalogo {0}", (int)this.ultraGrid1.ActiveRow.Cells["idCatalogo"].Value);
			FuncionesProcedimientos.dsGeneral(sSQL, this.ultraGrid2);
			this.ultraGrid2.DisplayLayout.Bands[0].SortedColumns.Add("Pagina", false, true);			
			this.ultraGrid2.Rows.ExpandAll(true);
			this.utcEstadoCajas.Tabs[1].Active= true;
			this.utcEstadoCajas.Tabs[1].Selected = true;
			this.ultraGrid2.DisplayLayout.Bands[0].Columns["Imagen"].Hidden = true;	
			this.ultraGrid2.DisplayLayout.Bands[0].Columns["Descripcion"].Hidden = true;	
			this.ultraGrid2.DisplayLayout.Bands[0].Columns["NombrePortada"].Hidden = true;	
//	    decimal cp=0.00m;
//			cp= Convert.ToDecimal(this.ultraGrid2.ActiveRow.Cells["cuotap"].Value);
			if (e.Cell.Column.ToString() == "Sel1") 
			{
				if (this.ultraGrid1.ActiveRow != null) 
					ultraGrid1.ActiveRow.Selected = false;
				//				int idcatalogo = (int)this.ultraGrid1.ActiveRow.Cells["idCatalogo"].Value;
				if((int)this.ultraGrid1.ActiveRow.Cells["idTipoCatalogo"].Value == 2)
				{
					this.ultraGrid2.DisplayLayout.Bands[0].Columns["CuotasE"].Hidden = false;
//					this.ultraGrid2.DisplayLayout.Bands[0].Columns["cuo6"].Hidden = false;	
//					this.ultraGrid2.DisplayLayout.Bands[0].Columns["cuo9"].Hidden = false;
//					this.ultraGrid2.DisplayLayout.Bands[0].Columns["cuo12"].Hidden = false;
//					this.ultraGrid2.DisplayLayout.Bands[0].Columns["cuo24"].Hidden = false;
				}
				else
				{
						this.ultraGrid2.DisplayLayout.Bands[0].Columns["CuotasC"].Hidden = false;
					this.ultraGrid2.DisplayLayout.Bands[0].Columns["Entrada"].Hidden = false;
//					this.ultraGrid2.DisplayLayout.Bands[0].Columns["cuc3"].Hidden = false;
//					this.ultraGrid2.DisplayLayout.Bands[0].Columns["cuc6"].Hidden = false;	
//					this.ultraGrid2.DisplayLayout.Bands[0].Columns["cuc9"].Hidden = false;	
//					this.ultraGrid2.DisplayLayout.Bands[0].Columns["cuc12"].Hidden = false;	
//					this.ultraGrid2.DisplayLayout.Bands[0].Columns["cuc18"].Hidden = false;	
				}
			}
			if (e.Cell.Column.ToString() == "Seleccion") 
			{
				if (this.ultraGrid1.ActiveRow != null) 
					ultraGrid1.ActiveRow.Selected = false;
				if((int)this.ultraGrid1.ActiveRow.Cells["idTipoCatalogo"].Value == 2)
				{
				this.ultraGrid2.DisplayLayout.Bands[0].Columns["CuotasE"].Hidden = false;
					//					this.ultraGrid2.DisplayLayout.Bands[0].Columns["cuo6"].Hidden = false;	
					//					this.ultraGrid2.DisplayLayout.Bands[0].Columns["cuo9"].Hidden = false;
					//					this.ultraGrid2.DisplayLayout.Bands[0].Columns["cuo12"].Hidden = false;
					//					this.ultraGrid2.DisplayLayout.Bands[0].Columns["cuo24"].Hidden = false;
				}
				else
				{
					this.ultraGrid2.DisplayLayout.Bands[0].Columns["CuotasC"].Hidden = false;
					this.ultraGrid2.DisplayLayout.Bands[0].Columns["Entrada"].Hidden = false;
					//					this.ultraGrid2.DisplayLayout.Bands[0].Columns["cuc3"].Hidden = false;
					//					this.ultraGrid2.DisplayLayout.Bands[0].Columns["cuc6"].Hidden = false;	
					//					this.ultraGrid2.DisplayLayout.Bands[0].Columns["cuc9"].Hidden = false;	
					//					this.ultraGrid2.DisplayLayout.Bands[0].Columns["cuc12"].Hidden = false;	
					//					this.ultraGrid2.DisplayLayout.Bands[0].Columns["cuc18"].Hidden = false;	
				}
				if (DialogResult.Yes == MessageBox.Show("Desea APROBAR el Catalogo con su portada de Marketing", "Point Technology", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2))
				{
					using (frmNotas Notas  = new frmNotas())
					{
						if (DialogResult.OK == Notas.ShowDialog())
						{
							this.txtObservaciones.Text = Notas.txtNotas.Text;
						}
					}
					string sSQL1 = string.Format("Update Catalogo set EstadoCatalogo = 5  where idCatalogo = {0}",(int)this.ultraGrid1.ActiveRow.Cells["idCatalogo"].Value);
					Funcion.EjecutaSQL(cdsSeteoF, sSQL1);

				}
			}
				
			if (e.Cell.Column.ToString() == "Anular")
			{				        
				
				if (this.ultraGrid1.ActiveRow != null) 
					ultraGrid1.ActiveRow.Selected = false;
				
				if((int)this.ultraGrid1.ActiveRow.Cells["idTipoCatalogo"].Value == 2)
				{
				this.ultraGrid2.DisplayLayout.Bands[0].Columns["CuotasE"].Hidden = false;
//					this.ultraGrid2.DisplayLayout.Bands[0].Columns["cuo6"].Hidden = false;	
//					this.ultraGrid2.DisplayLayout.Bands[0].Columns["cuo9"].Hidden = false;
//					this.ultraGrid2.DisplayLayout.Bands[0].Columns["cuo12"].Hidden = false;
//					this.ultraGrid2.DisplayLayout.Bands[0].Columns["cuo24"].Hidden = false;
				}
				else
				{
					this.ultraGrid2.DisplayLayout.Bands[0].Columns["CuotasC"].Hidden = false;
					this.ultraGrid2.DisplayLayout.Bands[0].Columns["Entrada"].Hidden = false;
//					this.ultraGrid2.DisplayLayout.Bands[0].Columns["cuc3"].Hidden = false;
//					this.ultraGrid2.DisplayLayout.Bands[0].Columns["cuc6"].Hidden = false;	
//					this.ultraGrid2.DisplayLayout.Bands[0].Columns["cuc9"].Hidden = false;	
//					this.ultraGrid2.DisplayLayout.Bands[0].Columns["cuc12"].Hidden = false;	
//					this.ultraGrid2.DisplayLayout.Bands[0].Columns["cuc18"].Hidden = false;	
				}
				if (this.ultraGrid1.ActiveRow.Cells["Estado"].Value.ToString() == "APROBADO_MARKETING")
				{
					if (DialogResult.Yes == MessageBox.Show("Desea CERRAR el Catalogo ", "Point Technology", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2))
					{
						using (frmNotas Notas  = new frmNotas())
						{
							if (DialogResult.OK == Notas.ShowDialog())
							{
								this.txtObservaciones.Text = Notas.txtNotas.Text;
							}
						}
						string sSQL2 = string.Format("Update Catalogo set EstadoCatalogo = 6 where idCatalogo = {0}",(int)this.ultraGrid1.ActiveRow.Cells["idCatalogo"].Value);
						Funcion.EjecutaSQL(cdsSeteoF, sSQL2);
	
					}
				}
				else{
					MessageBox.Show("No se puede Cerrar el catalogo en este estado debe estar autorizado por Marketing", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
		
				}
		
			}

				
		}

		private void cdsSeteoF_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			cdsSeteoF.Schema.Connections[0].ConnectionString = miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);		

		}

		private void ultraGrid2_InitializeLayout(object sender, Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs e)
		{



		}

		private void btnFinalizar_Click(object sender, System.EventArgs e)
		{
				if (this.txtEstadoCatalogo.Text == "5")
				{
					this.btnFinalizar.Enabled = true;
					Funcion.EjecutaSQL(cdsSeteoF, string.Format("Update Catalogo Set EstadoCatalogo = 4 Where idCatalogo =  {0}", (int)this.txtidCatalogo.Value));
					this.lblEstado.Text = "INICIADO";
					this.btnFinalizar.Text = "APROBAR";
				}
				else if (this.txtEstadoCatalogo.Text == "4")
				{
					this.btnFinalizar.Enabled = true;
					Funcion.EjecutaSQL(cdsSeteoF, string.Format("Update Catalogo Set EstadoCatalogo = 5 Where  idCatalogo = {0}", (int)this.txtidCatalogo.Value));
					this.lblEstado.Text = "APROBAR";
					this.btnFinalizar.Text = "Finalizar";
				}



		}

		private void ultraGrid1_MouseMove(object sender, System.Windows.Forms.MouseEventArgs e)
		{
//					this.ultraGrid1.ActiveRow.Cells["Anular"].ToolTipText= "Cerrar Catalogo";
		}

		private void btnFuente_Click(object sender, System.EventArgs e)
		{
			FontDialog font = new FontDialog();
			font.Font = rtbDescripcion.Font;
			if (font.ShowDialog() == DialogResult.OK)
			{
				rtbDescripcion.Font = font.Font;
			}
		}

		private void ultraTabPageControl1_Paint(object sender, System.Windows.Forms.PaintEventArgs e)
		{
		
		}

		private void btnColor_Click(object sender, System.EventArgs e)
		{
			ColorDialog color = new ColorDialog();
			if (color.ShowDialog() == DialogResult.OK)
			{
				rtbDescripcion.ForeColor = color.Color;
			}
		}

		private void txtEstadoCatalogo_ValueChanged(object sender, System.EventArgs e)
		{
		
		}

		private void btnExcel_Click(object sender, System.EventArgs e)
		{
			if (this.utcEstadoCajas.Tabs[0].Selected) FuncionesProcedimientos.ExportaExcel(this.ultraGrid1);
			else if (this.utcEstadoCajas.Tabs[1].Selected) FuncionesProcedimientos.ExportaExcel(this.ultraGrid2);

		}

	





	}
}
