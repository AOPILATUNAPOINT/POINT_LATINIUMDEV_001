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

namespace Latinium
{
	/// <summary>
	/// Descripción breve de frmPedidoPresupuesto.
	/// </summary>
	public class frmPedidoPresupuesto : System.Windows.Forms.Form
	{
		private System.Windows.Forms.Label label6;
		private System.Windows.Forms.Label lblBodega;
		private Infragistics.Win.UltraWinSchedule.UltraCalendarCombo dtFecha;
		private System.Windows.Forms.Label label2;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtNumero;
		private System.Windows.Forms.Label label8;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbFormaPago;
		private System.Windows.Forms.Label label3;
		private System.Windows.Forms.Label label1;
		private System.Windows.Forms.Label lblSubTotal;
		private System.Windows.Forms.Label lblIva;
		private System.Windows.Forms.Label lblDescuento;
		private System.Windows.Forms.Label lblIvaSubt;
		private System.Windows.Forms.Label lblIva0;
		private Infragistics.Win.UltraWinGrid.UltraDropDown cmbColor;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtIdProyecto;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtIdCredTributario;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtIdComprobante;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtContadoCredito;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtIdGrupoCliente;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtIdTipoRuc;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtBodegaPredef;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtEstado;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtIdCliente;
		private System.Windows.Forms.GroupBox gbBotones;
		private System.Windows.Forms.Button btnNuevo;
		private System.Windows.Forms.Button btnConsultar;
		private System.Windows.Forms.Button btnEditar;
		private System.Windows.Forms.Button btnImprimir;
		private System.Windows.Forms.Button btnGuardar;
		private System.Windows.Forms.Button btnAnular;
		private System.Windows.Forms.Button btnCancelar;
		private System.Windows.Forms.Button btnSalir;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtIdCompra;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtTotal;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtIvaTotal;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtDescIva0Total;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtDescIvaTotal;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtIva;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtIva0;
		private Infragistics.Win.UltraWinGrid.UltraGrid ultraGrid1;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtNotas;
		private System.Windows.Forms.Label lblUsuario;
		private System.Windows.Forms.Label lblEstado;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource10;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource2;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource1;
		private C1.Data.C1DataSet cdsSeteoF;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource3;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource4;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource5;
		private Infragistics.Win.UltraWinGrid.UltraDropDown cmbProyecto1;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbBodega;
		private System.Windows.Forms.Button btnDistribucion;
		/// <summary>
		/// Variable del diseñador requerida.
		/// </summary>
		private System.ComponentModel.Container components = null;

		public frmPedidoPresupuesto()
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
			Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton dateButton1 = new Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton();
			System.Resources.ResourceManager resources = new System.Resources.ResourceManager(typeof(frmPedidoPresupuesto));
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn1 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idProyecto");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn2 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Nombre");
			Infragistics.Win.Appearance appearance2 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance3 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand1 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn1 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idNomina");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn2 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn3 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idNomina");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn4 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Nombre");
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand2 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn3 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idColor");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn4 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Color");
			Infragistics.Win.Appearance appearance4 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance5 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance6 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance7 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance8 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance9 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance10 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand3 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn5 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idDetCompra");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn6 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idOrigen");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn7 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCompra");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn8 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idArticulo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn9 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idUnidad");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn10 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Codigo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn11 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Articulo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn12 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Cantidad");
			Infragistics.Win.Appearance appearance11 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn13 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("CantAereo");
			Infragistics.Win.Appearance appearance12 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn14 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Precio");
			Infragistics.Win.Appearance appearance13 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn15 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Impuesto");
			Infragistics.Win.Appearance appearance14 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn16 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("DescuentoArt");
			Infragistics.Win.Appearance appearance15 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn17 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("DescuentoPorc");
			Infragistics.Win.Appearance appearance16 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn18 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("SubTotal");
			Infragistics.Win.Appearance appearance17 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn19 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Valores");
			Infragistics.Win.Appearance appearance18 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn20 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idProyecto", -1, "cmbProyecto1");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn21 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idImportacion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn22 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Bodega");
			Infragistics.Win.Appearance appearance19 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn23 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Signo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn24 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Tipo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn25 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("EstadoSerial");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn26 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Courier");
			Infragistics.Win.Appearance appearance20 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn27 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Peso");
			Infragistics.Win.Appearance appearance21 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn28 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Arancel");
			Infragistics.Win.Appearance appearance22 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn29 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idTipoCalculo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn30 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("DescNC");
			Infragistics.Win.Appearance appearance23 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn31 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idColor", -1, "cmbColor");
			Infragistics.Win.Appearance appearance24 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance25 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance26 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance27 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance28 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance29 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance30 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance31 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn5 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idColor");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn6 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Color");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn7 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idImportacionNombre");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn8 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Nombre");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn9 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idDetCompra");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn10 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idOrigen");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn11 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idCompra");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn12 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idArticulo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn13 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idUnidad");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn14 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Codigo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn15 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Articulo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn16 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Cantidad");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn17 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("CantAereo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn18 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Precio");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn19 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Impuesto");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn20 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("DescuentoArt");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn21 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("DescuentoPorc");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn22 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("SubTotal");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn23 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Valores");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn24 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idProyecto");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn25 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idImportacion");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn26 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Bodega");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn27 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Signo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn28 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Tipo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn29 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("EstadoSerial");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn30 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Courier");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn31 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Peso");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn32 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Arancel");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn33 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idTipoCalculo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn34 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("DescNC");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn35 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idColor");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn36 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idProyecto");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn37 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Nombre");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn38 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("CentroCosto");
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand4 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn32 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idProyecto");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn33 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn34 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("CentroCosto");
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand5 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn35 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idProyecto");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn36 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			this.label6 = new System.Windows.Forms.Label();
			this.lblBodega = new System.Windows.Forms.Label();
			this.dtFecha = new Infragistics.Win.UltraWinSchedule.UltraCalendarCombo();
			this.ultraDataSource3 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.label2 = new System.Windows.Forms.Label();
			this.txtNumero = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.label8 = new System.Windows.Forms.Label();
			this.cmbFormaPago = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.ultraDataSource4 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.label3 = new System.Windows.Forms.Label();
			this.label1 = new System.Windows.Forms.Label();
			this.lblSubTotal = new System.Windows.Forms.Label();
			this.lblIva = new System.Windows.Forms.Label();
			this.lblDescuento = new System.Windows.Forms.Label();
			this.lblIvaSubt = new System.Windows.Forms.Label();
			this.lblIva0 = new System.Windows.Forms.Label();
			this.cmbColor = new Infragistics.Win.UltraWinGrid.UltraDropDown();
			this.txtIdProyecto = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtIdCredTributario = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtIdComprobante = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtContadoCredito = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtIdGrupoCliente = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtIdTipoRuc = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtBodegaPredef = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtEstado = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtIdCliente = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.gbBotones = new System.Windows.Forms.GroupBox();
			this.btnNuevo = new System.Windows.Forms.Button();
			this.btnConsultar = new System.Windows.Forms.Button();
			this.btnEditar = new System.Windows.Forms.Button();
			this.btnImprimir = new System.Windows.Forms.Button();
			this.btnGuardar = new System.Windows.Forms.Button();
			this.btnAnular = new System.Windows.Forms.Button();
			this.btnCancelar = new System.Windows.Forms.Button();
			this.btnSalir = new System.Windows.Forms.Button();
			this.txtIdCompra = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtTotal = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtIvaTotal = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtDescIva0Total = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtDescIvaTotal = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtIva = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtIva0 = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.ultraGrid1 = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.txtNotas = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.lblUsuario = new System.Windows.Forms.Label();
			this.lblEstado = new System.Windows.Forms.Label();
			this.ultraDataSource10 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.ultraDataSource2 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.ultraDataSource1 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.cdsSeteoF = new C1.Data.C1DataSet();
			this.ultraDataSource5 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.cmbProyecto1 = new Infragistics.Win.UltraWinGrid.UltraDropDown();
			this.cmbBodega = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.btnDistribucion = new System.Windows.Forms.Button();
			((System.ComponentModel.ISupportInitialize)(this.dtFecha)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource3)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNumero)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbFormaPago)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource4)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbColor)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIdProyecto)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIdCredTributario)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIdComprobante)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtContadoCredito)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIdGrupoCliente)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIdTipoRuc)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtBodegaPredef)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtEstado)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIdCliente)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIdCompra)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtTotal)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIvaTotal)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtDescIva0Total)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtDescIvaTotal)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIva)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIva0)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraGrid1)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNotas)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource10)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource2)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource1)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource5)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbProyecto1)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbBodega)).BeginInit();
			this.SuspendLayout();
			// 
			// label6
			// 
			this.label6.AutoSize = true;
			this.label6.Location = new System.Drawing.Point(16, 18);
			this.label6.Name = "label6";
			this.label6.Size = new System.Drawing.Size(36, 16);
			this.label6.TabIndex = 371;
			this.label6.Text = "Fecha";
			this.label6.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// lblBodega
			// 
			this.lblBodega.AutoSize = true;
			this.lblBodega.Location = new System.Drawing.Point(264, 18);
			this.lblBodega.Name = "lblBodega";
			this.lblBodega.Size = new System.Drawing.Size(49, 16);
			this.lblBodega.TabIndex = 372;
			this.lblBodega.Text = "Proyecto";
			this.lblBodega.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.lblBodega.Visible = false;
			// 
			// dtFecha
			// 
			appearance1.ForeColorDisabled = System.Drawing.Color.Black;
			this.dtFecha.Appearance = appearance1;
			this.dtFecha.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			dateButton1.Caption = "Today";
			this.dtFecha.DateButtons.Add(dateButton1);
			this.dtFecha.Enabled = false;
			this.dtFecha.Format = "dd/MM/yyyy";
			this.dtFecha.Location = new System.Drawing.Point(72, 16);
			this.dtFecha.Name = "dtFecha";
			this.dtFecha.NonAutoSizeHeight = 23;
			this.dtFecha.Size = new System.Drawing.Size(112, 21);
			this.dtFecha.SpinButtonsVisible = true;
			this.dtFecha.TabIndex = 374;
			this.dtFecha.Value = ((object)(resources.GetObject("dtFecha.Value")));
			this.dtFecha.KeyDown += new System.Windows.Forms.KeyEventHandler(this.dtFecha_KeyDown);
			this.dtFecha.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.dtFecha_KeyPress);
			this.dtFecha.Validating += new System.ComponentModel.CancelEventHandler(this.dtFecha_Validating);
			this.dtFecha.Validated += new System.EventHandler(this.dtFecha_Validated);
			this.dtFecha.ValueChanged += new System.EventHandler(this.dtFecha_ValueChanged);
			// 
			// ultraDataSource3
			// 
			ultraDataColumn1.DataType = typeof(int);
			this.ultraDataSource3.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn1,
																																 ultraDataColumn2});
			// 
			// label2
			// 
			this.label2.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			this.label2.AutoSize = true;
			this.label2.Location = new System.Drawing.Point(832, 16);
			this.label2.Name = "label2";
			this.label2.Size = new System.Drawing.Size(44, 16);
			this.label2.TabIndex = 376;
			this.label2.Text = "Número";
			this.label2.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// txtNumero
			// 
			this.txtNumero.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			appearance2.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtNumero.Appearance = appearance2;
			this.txtNumero.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtNumero.Enabled = false;
			this.txtNumero.Location = new System.Drawing.Point(896, 16);
			this.txtNumero.Name = "txtNumero";
			this.txtNumero.Size = new System.Drawing.Size(112, 21);
			this.txtNumero.TabIndex = 375;
			// 
			// label8
			// 
			this.label8.AutoSize = true;
			this.label8.Enabled = false;
			this.label8.Location = new System.Drawing.Point(16, 51);
			this.label8.Name = "label8";
			this.label8.Size = new System.Drawing.Size(41, 16);
			this.label8.TabIndex = 382;
			this.label8.Text = "Solicita";
			this.label8.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// cmbFormaPago
			// 
			appearance3.ForeColorDisabled = System.Drawing.Color.Black;
			this.cmbFormaPago.Appearance = appearance3;
			this.cmbFormaPago.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbFormaPago.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbFormaPago.DataSource = this.ultraDataSource4;
			ultraGridColumn1.Header.VisiblePosition = 0;
			ultraGridColumn1.Hidden = true;
			ultraGridColumn2.Header.VisiblePosition = 1;
			ultraGridColumn2.Width = 245;
			ultraGridBand1.Columns.AddRange(new object[] {
																										 ultraGridColumn1,
																										 ultraGridColumn2});
			this.cmbFormaPago.DisplayLayout.BandsSerializer.Add(ultraGridBand1);
			this.cmbFormaPago.DisplayMember = "Nombre";
			this.cmbFormaPago.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbFormaPago.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbFormaPago.Enabled = false;
			this.cmbFormaPago.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbFormaPago.Location = new System.Drawing.Point(72, 48);
			this.cmbFormaPago.Name = "cmbFormaPago";
			this.cmbFormaPago.Size = new System.Drawing.Size(248, 22);
			this.cmbFormaPago.TabIndex = 381;
			this.cmbFormaPago.ValueMember = "idNomina";
			// 
			// ultraDataSource4
			// 
			ultraDataColumn3.DataType = typeof(int);
			this.ultraDataSource4.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn3,
																																 ultraDataColumn4});
			// 
			// label3
			// 
			this.label3.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.label3.AutoSize = true;
			this.label3.Location = new System.Drawing.Point(12, 272);
			this.label3.Name = "label3";
			this.label3.Size = new System.Drawing.Size(34, 16);
			this.label3.TabIndex = 398;
			this.label3.Text = "Notas";
			this.label3.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// label1
			// 
			this.label1.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			this.label1.AutoSize = true;
			this.label1.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label1.Location = new System.Drawing.Point(756, 368);
			this.label1.Name = "label1";
			this.label1.Size = new System.Drawing.Size(41, 16);
			this.label1.TabIndex = 396;
			this.label1.Text = "TOTAL";
			this.label1.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// lblSubTotal
			// 
			this.lblSubTotal.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			this.lblSubTotal.AutoSize = true;
			this.lblSubTotal.Location = new System.Drawing.Point(756, 296);
			this.lblSubTotal.Name = "lblSubTotal";
			this.lblSubTotal.Size = new System.Drawing.Size(64, 16);
			this.lblSubTotal.TabIndex = 389;
			this.lblSubTotal.Text = "SUBTOTAL";
			this.lblSubTotal.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// lblIva
			// 
			this.lblIva.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			this.lblIva.AutoSize = true;
			this.lblIva.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.lblIva.Location = new System.Drawing.Point(756, 344);
			this.lblIva.Name = "lblIva";
			this.lblIva.Size = new System.Drawing.Size(22, 16);
			this.lblIva.TabIndex = 388;
			this.lblIva.Text = "IVA";
			this.lblIva.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// lblDescuento
			// 
			this.lblDescuento.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			this.lblDescuento.AutoSize = true;
			this.lblDescuento.Location = new System.Drawing.Point(756, 320);
			this.lblDescuento.Name = "lblDescuento";
			this.lblDescuento.Size = new System.Drawing.Size(75, 16);
			this.lblDescuento.TabIndex = 387;
			this.lblDescuento.Text = "DESCUENTO";
			this.lblDescuento.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// lblIvaSubt
			// 
			this.lblIvaSubt.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			this.lblIvaSubt.AutoSize = true;
			this.lblIvaSubt.Location = new System.Drawing.Point(988, 264);
			this.lblIvaSubt.Name = "lblIvaSubt";
			this.lblIvaSubt.Size = new System.Drawing.Size(22, 16);
			this.lblIvaSubt.TabIndex = 386;
			this.lblIvaSubt.Text = "IVA";
			this.lblIvaSubt.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// lblIva0
			// 
			this.lblIva0.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			this.lblIva0.AutoSize = true;
			this.lblIva0.Location = new System.Drawing.Point(884, 264);
			this.lblIva0.Name = "lblIva0";
			this.lblIva0.Size = new System.Drawing.Size(32, 16);
			this.lblIva0.TabIndex = 385;
			this.lblIva0.Text = "IVA 0";
			this.lblIva0.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// cmbColor
			// 
			this.cmbColor.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbColor.DisplayLayout.AutoFitColumns = true;
			ultraGridBand2.AddButtonCaption = "ImportacionTipo";
			ultraGridColumn3.Header.VisiblePosition = 0;
			ultraGridColumn3.Hidden = true;
			ultraGridColumn3.Width = 120;
			ultraGridColumn4.Header.VisiblePosition = 1;
			ultraGridColumn4.Width = 278;
			ultraGridBand2.Columns.AddRange(new object[] {
																										 ultraGridColumn3,
																										 ultraGridColumn4});
			this.cmbColor.DisplayLayout.BandsSerializer.Add(ultraGridBand2);
			this.cmbColor.DisplayMember = "Color";
			this.cmbColor.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbColor.Location = new System.Drawing.Point(384, 152);
			this.cmbColor.Name = "cmbColor";
			this.cmbColor.Size = new System.Drawing.Size(280, 112);
			this.cmbColor.TabIndex = 419;
			this.cmbColor.ValueMember = "idColor";
			this.cmbColor.Visible = false;
			// 
			// txtIdProyecto
			// 
			this.txtIdProyecto.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			this.txtIdProyecto.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtIdProyecto.Enabled = false;
			this.txtIdProyecto.Location = new System.Drawing.Point(860, 408);
			this.txtIdProyecto.Name = "txtIdProyecto";
			this.txtIdProyecto.PromptChar = ' ';
			this.txtIdProyecto.Size = new System.Drawing.Size(16, 21);
			this.txtIdProyecto.TabIndex = 418;
			this.txtIdProyecto.Visible = false;
			// 
			// txtIdCredTributario
			// 
			this.txtIdCredTributario.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			this.txtIdCredTributario.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtIdCredTributario.Enabled = false;
			this.txtIdCredTributario.Location = new System.Drawing.Point(1004, 408);
			this.txtIdCredTributario.Name = "txtIdCredTributario";
			this.txtIdCredTributario.PromptChar = ' ';
			this.txtIdCredTributario.Size = new System.Drawing.Size(16, 21);
			this.txtIdCredTributario.TabIndex = 416;
			this.txtIdCredTributario.Visible = false;
			// 
			// txtIdComprobante
			// 
			this.txtIdComprobante.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			this.txtIdComprobante.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtIdComprobante.Enabled = false;
			this.txtIdComprobante.Location = new System.Drawing.Point(988, 408);
			this.txtIdComprobante.Name = "txtIdComprobante";
			this.txtIdComprobante.PromptChar = ' ';
			this.txtIdComprobante.Size = new System.Drawing.Size(16, 21);
			this.txtIdComprobante.TabIndex = 415;
			this.txtIdComprobante.Visible = false;
			// 
			// txtContadoCredito
			// 
			this.txtContadoCredito.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			this.txtContadoCredito.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtContadoCredito.Enabled = false;
			this.txtContadoCredito.Location = new System.Drawing.Point(972, 408);
			this.txtContadoCredito.Name = "txtContadoCredito";
			this.txtContadoCredito.PromptChar = ' ';
			this.txtContadoCredito.Size = new System.Drawing.Size(16, 21);
			this.txtContadoCredito.TabIndex = 414;
			this.txtContadoCredito.Visible = false;
			// 
			// txtIdGrupoCliente
			// 
			this.txtIdGrupoCliente.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			this.txtIdGrupoCliente.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtIdGrupoCliente.Enabled = false;
			this.txtIdGrupoCliente.Location = new System.Drawing.Point(940, 408);
			this.txtIdGrupoCliente.Name = "txtIdGrupoCliente";
			this.txtIdGrupoCliente.PromptChar = ' ';
			this.txtIdGrupoCliente.Size = new System.Drawing.Size(16, 21);
			this.txtIdGrupoCliente.TabIndex = 413;
			this.txtIdGrupoCliente.Visible = false;
			// 
			// txtIdTipoRuc
			// 
			this.txtIdTipoRuc.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			this.txtIdTipoRuc.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtIdTipoRuc.Enabled = false;
			this.txtIdTipoRuc.Location = new System.Drawing.Point(924, 408);
			this.txtIdTipoRuc.Name = "txtIdTipoRuc";
			this.txtIdTipoRuc.PromptChar = ' ';
			this.txtIdTipoRuc.Size = new System.Drawing.Size(16, 21);
			this.txtIdTipoRuc.TabIndex = 412;
			this.txtIdTipoRuc.Visible = false;
			// 
			// txtBodegaPredef
			// 
			this.txtBodegaPredef.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			this.txtBodegaPredef.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtBodegaPredef.Enabled = false;
			this.txtBodegaPredef.Location = new System.Drawing.Point(956, 408);
			this.txtBodegaPredef.Name = "txtBodegaPredef";
			this.txtBodegaPredef.PromptChar = ' ';
			this.txtBodegaPredef.Size = new System.Drawing.Size(16, 21);
			this.txtBodegaPredef.TabIndex = 411;
			this.txtBodegaPredef.Visible = false;
			// 
			// txtEstado
			// 
			this.txtEstado.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			this.txtEstado.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtEstado.Enabled = false;
			this.txtEstado.Location = new System.Drawing.Point(892, 408);
			this.txtEstado.Name = "txtEstado";
			this.txtEstado.PromptChar = ' ';
			this.txtEstado.Size = new System.Drawing.Size(16, 21);
			this.txtEstado.TabIndex = 410;
			this.txtEstado.Visible = false;
			// 
			// txtIdCliente
			// 
			this.txtIdCliente.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			this.txtIdCliente.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtIdCliente.Enabled = false;
			this.txtIdCliente.Location = new System.Drawing.Point(908, 408);
			this.txtIdCliente.Name = "txtIdCliente";
			this.txtIdCliente.PromptChar = ' ';
			this.txtIdCliente.Size = new System.Drawing.Size(16, 21);
			this.txtIdCliente.TabIndex = 409;
			this.txtIdCliente.Visible = false;
			// 
			// gbBotones
			// 
			this.gbBotones.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.gbBotones.Location = new System.Drawing.Point(4, 392);
			this.gbBotones.Name = "gbBotones";
			this.gbBotones.Size = new System.Drawing.Size(1016, 8);
			this.gbBotones.TabIndex = 408;
			this.gbBotones.TabStop = false;
			// 
			// btnNuevo
			// 
			this.btnNuevo.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btnNuevo.CausesValidation = false;
			this.btnNuevo.Image = ((System.Drawing.Image)(resources.GetObject("btnNuevo.Image")));
			this.btnNuevo.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnNuevo.Location = new System.Drawing.Point(12, 408);
			this.btnNuevo.Name = "btnNuevo";
			this.btnNuevo.Size = new System.Drawing.Size(72, 23);
			this.btnNuevo.TabIndex = 399;
			this.btnNuevo.Text = "&Nuevo";
			this.btnNuevo.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnNuevo.Click += new System.EventHandler(this.btnNuevo_Click);
			// 
			// btnConsultar
			// 
			this.btnConsultar.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btnConsultar.CausesValidation = false;
			this.btnConsultar.Image = ((System.Drawing.Image)(resources.GetObject("btnConsultar.Image")));
			this.btnConsultar.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnConsultar.Location = new System.Drawing.Point(87, 408);
			this.btnConsultar.Name = "btnConsultar";
			this.btnConsultar.Size = new System.Drawing.Size(78, 23);
			this.btnConsultar.TabIndex = 406;
			this.btnConsultar.Text = "&Consultar";
			this.btnConsultar.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnConsultar.Click += new System.EventHandler(this.btnConsultar_Click);
			// 
			// btnEditar
			// 
			this.btnEditar.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btnEditar.CausesValidation = false;
			this.btnEditar.Enabled = false;
			this.btnEditar.Image = ((System.Drawing.Image)(resources.GetObject("btnEditar.Image")));
			this.btnEditar.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnEditar.Location = new System.Drawing.Point(168, 408);
			this.btnEditar.Name = "btnEditar";
			this.btnEditar.Size = new System.Drawing.Size(76, 23);
			this.btnEditar.TabIndex = 405;
			this.btnEditar.Text = "&Editar";
			this.btnEditar.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnEditar.Click += new System.EventHandler(this.btnEditar_Click);
			// 
			// btnImprimir
			// 
			this.btnImprimir.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btnImprimir.CausesValidation = false;
			this.btnImprimir.Enabled = false;
			this.btnImprimir.Image = ((System.Drawing.Image)(resources.GetObject("btnImprimir.Image")));
			this.btnImprimir.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnImprimir.Location = new System.Drawing.Point(247, 408);
			this.btnImprimir.Name = "btnImprimir";
			this.btnImprimir.Size = new System.Drawing.Size(77, 23);
			this.btnImprimir.TabIndex = 401;
			this.btnImprimir.Text = "&Imprimir";
			this.btnImprimir.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// btnGuardar
			// 
			this.btnGuardar.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btnGuardar.Enabled = false;
			this.btnGuardar.Image = ((System.Drawing.Image)(resources.GetObject("btnGuardar.Image")));
			this.btnGuardar.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnGuardar.Location = new System.Drawing.Point(328, 408);
			this.btnGuardar.Name = "btnGuardar";
			this.btnGuardar.Size = new System.Drawing.Size(76, 23);
			this.btnGuardar.TabIndex = 400;
			this.btnGuardar.Text = "&Guardar";
			this.btnGuardar.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnGuardar.Click += new System.EventHandler(this.btnGuardar_Click);
			// 
			// btnAnular
			// 
			this.btnAnular.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btnAnular.CausesValidation = false;
			this.btnAnular.Enabled = false;
			this.btnAnular.Image = ((System.Drawing.Image)(resources.GetObject("btnAnular.Image")));
			this.btnAnular.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnAnular.Location = new System.Drawing.Point(672, 408);
			this.btnAnular.Name = "btnAnular";
			this.btnAnular.Size = new System.Drawing.Size(76, 23);
			this.btnAnular.TabIndex = 407;
			this.btnAnular.Text = "&Anular";
			this.btnAnular.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnAnular.Visible = false;
			this.btnAnular.Click += new System.EventHandler(this.btnAnular_Click);
			// 
			// btnCancelar
			// 
			this.btnCancelar.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btnCancelar.CausesValidation = false;
			this.btnCancelar.Enabled = false;
			this.btnCancelar.Image = ((System.Drawing.Image)(resources.GetObject("btnCancelar.Image")));
			this.btnCancelar.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnCancelar.Location = new System.Drawing.Point(408, 408);
			this.btnCancelar.Name = "btnCancelar";
			this.btnCancelar.Size = new System.Drawing.Size(72, 23);
			this.btnCancelar.TabIndex = 402;
			this.btnCancelar.Text = "&Cancelar";
			this.btnCancelar.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnCancelar.Click += new System.EventHandler(this.btnCancelar_Click);
			// 
			// btnSalir
			// 
			this.btnSalir.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btnSalir.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(230)), ((System.Byte)(230)), ((System.Byte)(230)));
			this.btnSalir.CausesValidation = false;
			this.btnSalir.ForeColor = System.Drawing.Color.Black;
			this.btnSalir.Image = ((System.Drawing.Image)(resources.GetObject("btnSalir.Image")));
			this.btnSalir.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnSalir.Location = new System.Drawing.Point(488, 408);
			this.btnSalir.Name = "btnSalir";
			this.btnSalir.TabIndex = 403;
			this.btnSalir.Text = "&Salir";
			this.btnSalir.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnSalir.Click += new System.EventHandler(this.btnSalir_Click);
			// 
			// txtIdCompra
			// 
			this.txtIdCompra.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			this.txtIdCompra.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtIdCompra.Enabled = false;
			this.txtIdCompra.Location = new System.Drawing.Point(876, 408);
			this.txtIdCompra.Name = "txtIdCompra";
			this.txtIdCompra.PromptChar = ' ';
			this.txtIdCompra.Size = new System.Drawing.Size(16, 21);
			this.txtIdCompra.TabIndex = 397;
			this.txtIdCompra.Visible = false;
			// 
			// txtTotal
			// 
			this.txtTotal.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			appearance4.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtTotal.Appearance = appearance4;
			this.txtTotal.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtTotal.Enabled = false;
			this.txtTotal.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.txtTotal.FormatString = "#,##0.00";
			this.txtTotal.Location = new System.Drawing.Point(836, 360);
			this.txtTotal.Name = "txtTotal";
			this.txtTotal.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtTotal.PromptChar = ' ';
			this.txtTotal.Size = new System.Drawing.Size(176, 28);
			this.txtTotal.TabIndex = 395;
			// 
			// txtIvaTotal
			// 
			this.txtIvaTotal.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			appearance5.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtIvaTotal.Appearance = appearance5;
			this.txtIvaTotal.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtIvaTotal.Enabled = false;
			this.txtIvaTotal.FormatString = "#,##0.00";
			this.txtIvaTotal.Location = new System.Drawing.Point(932, 336);
			this.txtIvaTotal.Name = "txtIvaTotal";
			this.txtIvaTotal.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtIvaTotal.PromptChar = ' ';
			this.txtIvaTotal.Size = new System.Drawing.Size(80, 21);
			this.txtIvaTotal.TabIndex = 394;
			// 
			// txtDescIva0Total
			// 
			this.txtDescIva0Total.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			appearance6.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtDescIva0Total.Appearance = appearance6;
			this.txtDescIva0Total.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtDescIva0Total.Enabled = false;
			this.txtDescIva0Total.FormatString = "#,##0.00";
			this.txtDescIva0Total.Location = new System.Drawing.Point(836, 312);
			this.txtDescIva0Total.Name = "txtDescIva0Total";
			this.txtDescIva0Total.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtDescIva0Total.PromptChar = ' ';
			this.txtDescIva0Total.Size = new System.Drawing.Size(80, 21);
			this.txtDescIva0Total.TabIndex = 393;
			// 
			// txtDescIvaTotal
			// 
			this.txtDescIvaTotal.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			appearance7.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtDescIvaTotal.Appearance = appearance7;
			this.txtDescIvaTotal.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtDescIvaTotal.Enabled = false;
			this.txtDescIvaTotal.FormatString = "#,##0.00";
			this.txtDescIvaTotal.Location = new System.Drawing.Point(932, 312);
			this.txtDescIvaTotal.MaskInput = "{LOC}nn,nnn,nnn.nn";
			this.txtDescIvaTotal.Name = "txtDescIvaTotal";
			this.txtDescIvaTotal.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtDescIvaTotal.PromptChar = ' ';
			this.txtDescIvaTotal.Size = new System.Drawing.Size(80, 21);
			this.txtDescIvaTotal.TabIndex = 392;
			// 
			// txtIva
			// 
			this.txtIva.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			appearance8.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtIva.Appearance = appearance8;
			this.txtIva.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtIva.Enabled = false;
			this.txtIva.FormatString = "#,##0.00";
			this.txtIva.Location = new System.Drawing.Point(932, 288);
			this.txtIva.MaskInput = "{LOC}nn,nnn,nnn.nn";
			this.txtIva.Name = "txtIva";
			this.txtIva.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtIva.PromptChar = ' ';
			this.txtIva.Size = new System.Drawing.Size(80, 21);
			this.txtIva.TabIndex = 391;
			// 
			// txtIva0
			// 
			this.txtIva0.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			appearance9.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtIva0.Appearance = appearance9;
			this.txtIva0.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtIva0.Enabled = false;
			this.txtIva0.FormatString = "#,##0.00";
			this.txtIva0.Location = new System.Drawing.Point(836, 288);
			this.txtIva0.Name = "txtIva0";
			this.txtIva0.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtIva0.PromptChar = ' ';
			this.txtIva0.Size = new System.Drawing.Size(80, 21);
			this.txtIva0.TabIndex = 390;
			// 
			// ultraGrid1
			// 
			this.ultraGrid1.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.ultraGrid1.Cursor = System.Windows.Forms.Cursors.Default;
			appearance10.BackColor = System.Drawing.Color.White;
			this.ultraGrid1.DisplayLayout.Appearance = appearance10;
			this.ultraGrid1.DisplayLayout.AutoFitColumns = true;
			ultraGridBand3.AddButtonCaption = "_DetCompra";
			ultraGridColumn5.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn5.Header.VisiblePosition = 8;
			ultraGridColumn5.Hidden = true;
			ultraGridColumn6.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn6.Header.VisiblePosition = 15;
			ultraGridColumn6.Hidden = true;
			ultraGridColumn7.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn7.Header.VisiblePosition = 0;
			ultraGridColumn7.Hidden = true;
			ultraGridColumn8.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn8.Header.Caption = "Artículo";
			ultraGridColumn8.Header.VisiblePosition = 1;
			ultraGridColumn8.Hidden = true;
			ultraGridColumn8.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.DropDownList;
			ultraGridColumn8.Width = 194;
			ultraGridColumn9.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn9.Header.Caption = "Unidad";
			ultraGridColumn9.Header.VisiblePosition = 3;
			ultraGridColumn9.Hidden = true;
			ultraGridColumn9.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.DropDownValidate;
			ultraGridColumn9.Width = 53;
			ultraGridColumn10.Case = Infragistics.Win.UltraWinGrid.Case.Upper;
			ultraGridColumn10.Header.VisiblePosition = 2;
			ultraGridColumn10.Width = 182;
			ultraGridColumn11.Case = Infragistics.Win.UltraWinGrid.Case.Upper;
			ultraGridColumn11.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			ultraGridColumn11.Header.VisiblePosition = 4;
			ultraGridColumn11.Width = 324;
			appearance11.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn12.CellAppearance = appearance11;
			ultraGridColumn12.Format = "#,##0";
			ultraGridColumn12.Header.VisiblePosition = 6;
			ultraGridColumn12.MaskInput = "";
			ultraGridColumn12.NullText = "0";
			ultraGridColumn12.PromptChar = ' ';
			ultraGridColumn12.Width = 65;
			ultraGridColumn13.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance12.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn13.CellAppearance = appearance12;
			ultraGridColumn13.Format = "#,##0";
			ultraGridColumn13.Header.Caption = "Avión";
			ultraGridColumn13.Header.VisiblePosition = 7;
			ultraGridColumn13.Hidden = true;
			ultraGridColumn13.MaskInput = "";
			ultraGridColumn13.NullText = "";
			ultraGridColumn13.PromptChar = ' ';
			ultraGridColumn13.Width = 70;
			appearance13.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn14.CellAppearance = appearance13;
			ultraGridColumn14.Format = "#,##0.00";
			ultraGridColumn14.Header.VisiblePosition = 9;
			ultraGridColumn14.MaskInput = "{LOC}nn,nnn,nnn.nn";
			ultraGridColumn14.PromptChar = ' ';
			ultraGridColumn14.Width = 78;
			appearance14.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn15.CellAppearance = appearance14;
			ultraGridColumn15.Format = "#,##0.00";
			ultraGridColumn15.Header.Caption = "IVA";
			ultraGridColumn15.Header.VisiblePosition = 10;
			ultraGridColumn15.MaskInput = "{LOC}nnnnnnn.nn";
			ultraGridColumn15.PromptChar = ' ';
			ultraGridColumn15.Width = 46;
			appearance15.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn16.CellAppearance = appearance15;
			ultraGridColumn16.Format = "#,##0.00";
			ultraGridColumn16.Header.Caption = "Descuento";
			ultraGridColumn16.Header.VisiblePosition = 11;
			ultraGridColumn16.Hidden = true;
			ultraGridColumn16.MaskInput = "{LOC}nn,nnn,nnn.nn";
			ultraGridColumn16.PromptChar = ' ';
			ultraGridColumn16.Width = 14;
			appearance16.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn17.CellAppearance = appearance16;
			ultraGridColumn17.Format = "#,##0.00";
			ultraGridColumn17.Header.Caption = "% Desc";
			ultraGridColumn17.Header.VisiblePosition = 12;
			ultraGridColumn17.MaskInput = "{LOC}nnnnnnn.nn";
			ultraGridColumn17.PromptChar = ' ';
			ultraGridColumn17.Width = 56;
			ultraGridColumn18.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance17.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn18.CellAppearance = appearance17;
			ultraGridColumn18.Format = "#,##0.00";
			ultraGridColumn18.Header.VisiblePosition = 13;
			ultraGridColumn18.PromptChar = ' ';
			ultraGridColumn18.Width = 93;
			appearance18.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn19.CellAppearance = appearance18;
			ultraGridColumn19.Format = "#,##0.00";
			ultraGridColumn19.Header.VisiblePosition = 26;
			ultraGridColumn19.Hidden = true;
			ultraGridColumn20.ButtonDisplayStyle = Infragistics.Win.UltraWinGrid.ButtonDisplayStyle.Always;
			ultraGridColumn20.Header.Caption = "Centro de costo";
			ultraGridColumn20.Header.VisiblePosition = 17;
			ultraGridColumn20.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.DropDownList;
			ultraGridColumn20.Width = 135;
			ultraGridColumn21.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn21.Header.VisiblePosition = 18;
			ultraGridColumn21.Hidden = true;
			ultraGridColumn22.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance19.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn22.CellAppearance = appearance19;
			ultraGridColumn22.Header.VisiblePosition = 14;
			ultraGridColumn22.Hidden = true;
			ultraGridColumn22.MaskInput = "nnnnnnnnn";
			ultraGridColumn22.Width = 14;
			ultraGridColumn23.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn23.Header.VisiblePosition = 16;
			ultraGridColumn23.Hidden = true;
			ultraGridColumn24.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn24.Header.VisiblePosition = 19;
			ultraGridColumn24.Hidden = true;
			ultraGridColumn25.Header.VisiblePosition = 20;
			ultraGridColumn25.Hidden = true;
			appearance20.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn26.CellAppearance = appearance20;
			ultraGridColumn26.Format = "#,##0.00";
			ultraGridColumn26.Header.VisiblePosition = 21;
			ultraGridColumn26.Hidden = true;
			ultraGridColumn27.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance21.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn27.CellAppearance = appearance21;
			ultraGridColumn27.Format = "#,##0.00";
			ultraGridColumn27.Header.VisiblePosition = 22;
			ultraGridColumn27.Hidden = true;
			ultraGridColumn28.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance22.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn28.CellAppearance = appearance22;
			ultraGridColumn28.Format = "#,##0.00";
			ultraGridColumn28.Header.VisiblePosition = 23;
			ultraGridColumn28.Hidden = true;
			ultraGridColumn29.Header.VisiblePosition = 24;
			ultraGridColumn29.Hidden = true;
			appearance23.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn30.CellAppearance = appearance23;
			ultraGridColumn30.Format = "#,##0.00";
			ultraGridColumn30.Header.Caption = "Dscto Prom";
			ultraGridColumn30.Header.VisiblePosition = 25;
			ultraGridColumn30.Hidden = true;
			ultraGridColumn30.Width = 87;
			ultraGridColumn31.Case = Infragistics.Win.UltraWinGrid.Case.Upper;
			ultraGridColumn31.Header.Caption = "Color";
			ultraGridColumn31.Header.VisiblePosition = 5;
			ultraGridColumn31.Hidden = true;
			ultraGridColumn31.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.DropDownList;
			ultraGridColumn31.Width = 80;
			ultraGridBand3.Columns.AddRange(new object[] {
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
																										 ultraGridColumn31});
			this.ultraGrid1.DisplayLayout.BandsSerializer.Add(ultraGridBand3);
			appearance24.ForeColor = System.Drawing.Color.Black;
			appearance24.ForeColorDisabled = System.Drawing.Color.Black;
			this.ultraGrid1.DisplayLayout.Override.ActiveRowAppearance = appearance24;
			this.ultraGrid1.DisplayLayout.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.TemplateOnBottom;
			this.ultraGrid1.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.True;
			this.ultraGrid1.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.True;
			appearance25.BackColor = System.Drawing.Color.Transparent;
			this.ultraGrid1.DisplayLayout.Override.CardAreaAppearance = appearance25;
			appearance26.ForeColor = System.Drawing.Color.Black;
			appearance26.ForeColorDisabled = System.Drawing.Color.Black;
			this.ultraGrid1.DisplayLayout.Override.CellAppearance = appearance26;
			appearance27.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance27.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance27.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance27.FontData.BoldAsString = "True";
			appearance27.FontData.Name = "Arial";
			appearance27.FontData.SizeInPoints = 8.5F;
			appearance27.ForeColor = System.Drawing.Color.White;
			appearance27.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.ultraGrid1.DisplayLayout.Override.HeaderAppearance = appearance27;
			appearance28.BackColor = System.Drawing.SystemColors.ControlLight;
			appearance28.BackColor2 = System.Drawing.Color.Gainsboro;
			appearance28.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ultraGrid1.DisplayLayout.Override.RowAlternateAppearance = appearance28;
			appearance29.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance29.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance29.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ultraGrid1.DisplayLayout.Override.RowSelectorAppearance = appearance29;
			appearance30.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance30.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance30.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ultraGrid1.DisplayLayout.Override.SelectedRowAppearance = appearance30;
			this.ultraGrid1.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.ultraGrid1.Location = new System.Drawing.Point(12, 88);
			this.ultraGrid1.Name = "ultraGrid1";
			this.ultraGrid1.Size = new System.Drawing.Size(1000, 168);
			this.ultraGrid1.TabIndex = 384;
			this.ultraGrid1.AfterRowsDeleted += new System.EventHandler(this.ultraGrid1_AfterRowsDeleted);
			this.ultraGrid1.AfterRowInsert += new Infragistics.Win.UltraWinGrid.RowEventHandler(this.ultraGrid1_AfterRowInsert);
			this.ultraGrid1.BeforeCellUpdate += new Infragistics.Win.UltraWinGrid.BeforeCellUpdateEventHandler(this.ultraGrid1_BeforeCellUpdate);
			this.ultraGrid1.BeforeRowsDeleted += new Infragistics.Win.UltraWinGrid.BeforeRowsDeletedEventHandler(this.ultraGrid1_BeforeRowsDeleted);
			this.ultraGrid1.KeyDown += new System.Windows.Forms.KeyEventHandler(this.ultraGrid1_KeyDown);
			this.ultraGrid1.AfterCellUpdate += new Infragistics.Win.UltraWinGrid.CellEventHandler(this.ultraGrid1_AfterCellUpdate);
			this.ultraGrid1.InitializeLayout += new Infragistics.Win.UltraWinGrid.InitializeLayoutEventHandler(this.ultraGrid1_InitializeLayout);
			// 
			// txtNotas
			// 
			this.txtNotas.AcceptsReturn = true;
			this.txtNotas.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			appearance31.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtNotas.Appearance = appearance31;
			this.txtNotas.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtNotas.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtNotas.Enabled = false;
			this.txtNotas.Location = new System.Drawing.Point(12, 296);
			this.txtNotas.Multiline = true;
			this.txtNotas.Name = "txtNotas";
			this.txtNotas.Size = new System.Drawing.Size(550, 88);
			this.txtNotas.TabIndex = 383;
			// 
			// lblUsuario
			// 
			this.lblUsuario.AutoSize = true;
			this.lblUsuario.BackColor = System.Drawing.Color.Transparent;
			this.lblUsuario.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
			this.lblUsuario.Font = new System.Drawing.Font("Tahoma", 10F, System.Drawing.FontStyle.Bold);
			this.lblUsuario.Location = new System.Drawing.Point(728, 53);
			this.lblUsuario.Name = "lblUsuario";
			this.lblUsuario.Size = new System.Drawing.Size(0, 23);
			this.lblUsuario.TabIndex = 421;
			this.lblUsuario.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
			// 
			// lblEstado
			// 
			this.lblEstado.AutoSize = true;
			this.lblEstado.Font = new System.Drawing.Font("Microsoft Sans Serif", 15F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.lblEstado.ForeColor = System.Drawing.Color.Firebrick;
			this.lblEstado.Location = new System.Drawing.Point(408, 45);
			this.lblEstado.Name = "lblEstado";
			this.lblEstado.Size = new System.Drawing.Size(0, 26);
			this.lblEstado.TabIndex = 420;
			this.lblEstado.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
			// 
			// ultraDataSource10
			// 
			ultraDataColumn5.DataType = typeof(int);
			this.ultraDataSource10.Band.Columns.AddRange(new object[] {
																																	ultraDataColumn5,
																																	ultraDataColumn6});
			// 
			// ultraDataSource2
			// 
			ultraDataColumn7.DataType = typeof(int);
			this.ultraDataSource2.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn7,
																																 ultraDataColumn8});
			// 
			// ultraDataSource1
			// 
			ultraDataColumn9.DataType = typeof(int);
			ultraDataColumn10.DataType = typeof(int);
			ultraDataColumn11.DataType = typeof(int);
			ultraDataColumn12.DataType = typeof(int);
			ultraDataColumn13.DataType = typeof(int);
			ultraDataColumn16.DataType = typeof(int);
			ultraDataColumn16.DefaultValue = 0;
			ultraDataColumn17.DataType = typeof(int);
			ultraDataColumn17.DefaultValue = 0;
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
			ultraDataColumn24.DataType = typeof(int);
			ultraDataColumn25.DataType = typeof(int);
			ultraDataColumn26.DataType = typeof(int);
			ultraDataColumn27.DataType = typeof(int);
			ultraDataColumn28.DataType = typeof(int);
			ultraDataColumn29.DataType = typeof(bool);
			ultraDataColumn29.DefaultValue = false;
			ultraDataColumn30.DataType = typeof(System.Decimal);
			ultraDataColumn30.DefaultValue = new System.Decimal(new int[] {
																																			0,
																																			0,
																																			0,
																																			0});
			ultraDataColumn31.DataType = typeof(System.Decimal);
			ultraDataColumn31.DefaultValue = new System.Decimal(new int[] {
																																			0,
																																			0,
																																			0,
																																			0});
			ultraDataColumn32.DataType = typeof(System.Decimal);
			ultraDataColumn32.DefaultValue = new System.Decimal(new int[] {
																																			0,
																																			0,
																																			0,
																																			0});
			ultraDataColumn33.DataType = typeof(int);
			ultraDataColumn34.DataType = typeof(System.Decimal);
			ultraDataColumn34.DefaultValue = new System.Decimal(new int[] {
																																			0,
																																			0,
																																			0,
																																			0});
			ultraDataColumn35.DataType = typeof(int);
			this.ultraDataSource1.Band.Columns.AddRange(new object[] {
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
																																 ultraDataColumn35});
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
			// ultraDataSource5
			// 
			ultraDataColumn36.DataType = typeof(int);
			this.ultraDataSource5.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn36,
																																 ultraDataColumn37,
																																 ultraDataColumn38});
			// 
			// cmbProyecto1
			// 
			this.cmbProyecto1.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbProyecto1.DataSource = this.ultraDataSource5;
			ultraGridColumn32.Header.VisiblePosition = 0;
			ultraGridColumn32.Hidden = true;
			ultraGridColumn33.Header.VisiblePosition = 1;
			ultraGridColumn33.Width = 200;
			ultraGridColumn34.Header.VisiblePosition = 2;
			ultraGridBand4.Columns.AddRange(new object[] {
																										 ultraGridColumn32,
																										 ultraGridColumn33,
																										 ultraGridColumn34});
			this.cmbProyecto1.DisplayLayout.BandsSerializer.Add(ultraGridBand4);
			this.cmbProyecto1.DisplayMember = "Nombre";
			this.cmbProyecto1.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbProyecto1.Location = new System.Drawing.Point(568, 280);
			this.cmbProyecto1.Name = "cmbProyecto1";
			this.cmbProyecto1.Size = new System.Drawing.Size(256, 112);
			this.cmbProyecto1.TabIndex = 422;
			this.cmbProyecto1.ValueMember = "idProyecto";
			this.cmbProyecto1.Visible = false;
			// 
			// cmbBodega
			// 
			this.cmbBodega.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbBodega.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbBodega.DataSource = this.ultraDataSource3;
			ultraGridColumn35.Header.VisiblePosition = 0;
			ultraGridColumn36.Header.VisiblePosition = 1;
			ultraGridColumn36.Width = 180;
			ultraGridBand5.Columns.AddRange(new object[] {
																										 ultraGridColumn35,
																										 ultraGridColumn36});
			this.cmbBodega.DisplayLayout.BandsSerializer.Add(ultraGridBand5);
			this.cmbBodega.DisplayMember = "Nombre";
			this.cmbBodega.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbBodega.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbBodega.Enabled = false;
			this.cmbBodega.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbBodega.Location = new System.Drawing.Point(328, 16);
			this.cmbBodega.MaxDropDownItems = 30;
			this.cmbBodega.Name = "cmbBodega";
			this.cmbBodega.Size = new System.Drawing.Size(180, 21);
			this.cmbBodega.TabIndex = 423;
			this.cmbBodega.ValueMember = "idProyecto";
			this.cmbBodega.Visible = false;
			// 
			// btnDistribucion
			// 
			this.btnDistribucion.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			this.btnDistribucion.CausesValidation = false;
			this.btnDistribucion.Enabled = false;
			this.btnDistribucion.Image = ((System.Drawing.Image)(resources.GetObject("btnDistribucion.Image")));
			this.btnDistribucion.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnDistribucion.Location = new System.Drawing.Point(904, 48);
			this.btnDistribucion.Name = "btnDistribucion";
			this.btnDistribucion.Size = new System.Drawing.Size(104, 24);
			this.btnDistribucion.TabIndex = 424;
			this.btnDistribucion.Text = "Distribucion";
			this.btnDistribucion.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnDistribucion.Click += new System.EventHandler(this.btnDistribucion_Click);
			// 
			// frmPedidoPresupuesto
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 13);
			this.ClientSize = new System.Drawing.Size(1024, 438);
			this.Controls.Add(this.btnDistribucion);
			this.Controls.Add(this.cmbBodega);
			this.Controls.Add(this.cmbProyecto1);
			this.Controls.Add(this.lblUsuario);
			this.Controls.Add(this.lblEstado);
			this.Controls.Add(this.label3);
			this.Controls.Add(this.label1);
			this.Controls.Add(this.lblSubTotal);
			this.Controls.Add(this.lblIva);
			this.Controls.Add(this.lblDescuento);
			this.Controls.Add(this.lblIvaSubt);
			this.Controls.Add(this.lblIva0);
			this.Controls.Add(this.label8);
			this.Controls.Add(this.label2);
			this.Controls.Add(this.label6);
			this.Controls.Add(this.lblBodega);
			this.Controls.Add(this.cmbColor);
			this.Controls.Add(this.txtIdProyecto);
			this.Controls.Add(this.txtIdCredTributario);
			this.Controls.Add(this.txtIdComprobante);
			this.Controls.Add(this.txtContadoCredito);
			this.Controls.Add(this.txtIdGrupoCliente);
			this.Controls.Add(this.txtIdTipoRuc);
			this.Controls.Add(this.txtBodegaPredef);
			this.Controls.Add(this.txtEstado);
			this.Controls.Add(this.txtIdCliente);
			this.Controls.Add(this.gbBotones);
			this.Controls.Add(this.btnNuevo);
			this.Controls.Add(this.btnConsultar);
			this.Controls.Add(this.btnEditar);
			this.Controls.Add(this.btnImprimir);
			this.Controls.Add(this.btnGuardar);
			this.Controls.Add(this.btnAnular);
			this.Controls.Add(this.btnCancelar);
			this.Controls.Add(this.btnSalir);
			this.Controls.Add(this.txtIdCompra);
			this.Controls.Add(this.txtTotal);
			this.Controls.Add(this.txtIvaTotal);
			this.Controls.Add(this.txtDescIva0Total);
			this.Controls.Add(this.txtDescIvaTotal);
			this.Controls.Add(this.txtIva);
			this.Controls.Add(this.txtIva0);
			this.Controls.Add(this.ultraGrid1);
			this.Controls.Add(this.txtNotas);
			this.Controls.Add(this.cmbFormaPago);
			this.Controls.Add(this.txtNumero);
			this.Controls.Add(this.dtFecha);
			this.KeyPreview = true;
			this.Name = "frmPedidoPresupuesto";
			this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
			this.Text = "Pedido Presupuesto";
			this.KeyDown += new System.Windows.Forms.KeyEventHandler(this.frmPedidoPresupuesto_KeyDown);
			this.Click += new System.EventHandler(this.su);
			this.Load += new System.EventHandler(this.frmPedidoPresupuesto_Load);
			((System.ComponentModel.ISupportInitialize)(this.dtFecha)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource3)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNumero)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbFormaPago)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource4)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbColor)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIdProyecto)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIdCredTributario)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIdComprobante)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtContadoCredito)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIdGrupoCliente)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIdTipoRuc)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtBodegaPredef)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtEstado)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIdCliente)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIdCompra)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtTotal)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIvaTotal)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtDescIva0Total)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtDescIvaTotal)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIva)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIva0)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraGrid1)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNotas)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource10)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource2)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource1)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource5)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbProyecto1)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbBodega)).EndInit();
			this.ResumeLayout(false);

		}
		#endregion

		bool bNuevo = false;
		bool bEdicion = false;
		bool bAnulado = false;
		bool bEliminaAlValidar = false;
		private Acceso miAcceso;
		bool bRastreo = false;
		int IdSigno = 1;
		private C1DataRow drSeteoF;
		private bool bActualiza = true;
		decimal dPorcIva = 0.00m;
		int idBloqueaTransacciones = 0;
		int idProyecto = 0;
		bool bCarga;
		bool ODCDiaCredito = false;
		int iOrigen = 0;

		private void frmPedidoPresupuesto_Load(object sender, System.EventArgs e)
		{
			miAcceso = new Acceso(cdsSeteoF, "2608");

			bCarga=true;

			if (!Funcion.Permiso("1730", cdsSeteoF))
			{				
				MessageBox.Show("No Puede Ingresar a Presupuesto de Pedidos", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				BeginInvoke(new MethodInvoker(this.Close));				

				return;
			}
			drSeteoF = cdsSeteoF.TableViews["SeteoF"].Rows[0];

			Funcion.EjecutaSQL(cdsSeteoF, "Exec AuditaCrear 53, 6, 'ORDEN DE COMPRA'", true);

			bRastreo = Funcion.bEjecutaSQL(cdsSeteoF, "Exec SeteoGExiste 'RASTREO'");

			IdSigno = Funcion.iEscalarSQL(cdsSeteoF, "SELECT Signo FROM CompraNumero WHERE idTipoFactura = 2");
				
				
	
		}
		private void Rastreo(string stMensaje)
		{
			string stSelect = "Insert into ErrorSalida (Mensaje) Values ('" + stMensaje + "')";
			Funcion.EjecutaSQL(cdsSeteoF, stSelect);
		}

		private void UnloadMe()
		{
			this.Close();
		}
		//private Acceso miAcceso;


		private void cmdTotal()
		{
			#region Creacion de Variables			
			
			decimal dDescuentoLineal = 0.00m;
			decimal dIvaLineal = 0.00m;

			int dRedondeo = (int) drSeteoF["DecRedondeo"];			
			#endregion Creacion de Variables
			
			this.txtIva.Value = 0.00m;
			this.txtIva0.Value = 0.00m;
			this.txtDescIvaTotal.Value = 0.00m;
			this.txtDescIva0Total.Value = 0.00m;
			this.txtIvaTotal.Value = 0.00m;

			#region Desplazamiento de grilla
			foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in ultraGrid1.Rows.All)
			{
				dDescuentoLineal = 0.00m;
				dIvaLineal = 0.00m;

				if (dr.Cells["Cantidad"].Value != DBNull.Value && dr.Cells["Precio"].Value != DBNull.Value && dr.Cells["DescuentoPorc"].Value != DBNull.Value && dr.Cells["DescNC"].Value != DBNull.Value)
				{
					dr.Cells["Subtotal"].Value = Math.Round(Convert.ToDecimal(dr.Cells["Cantidad"].Value) * Convert.ToDecimal(dr.Cells["Precio"].Value), 2);
					dr.Cells["Subtotal"].Value = Convert.ToDecimal(dr.Cells["Subtotal"].Value) - Convert.ToDecimal(dr.Cells["DescNC"].Value);
					dDescuentoLineal = Math.Round((Convert.ToDecimal(dr.Cells["Subtotal"].Value) * Convert.ToDecimal(dr.Cells["DescuentoPorc"].Value))/100, 2);
					dr.Cells["DescuentoArt"].Value = Math.Round(dDescuentoLineal, 2);

					if (Convert.ToDecimal(dr.Cells["Impuesto"].Value) > 0.00m)
					{						
						this.txtIva.Value = Convert.ToDecimal(this.txtIva.Value) + Convert.ToDecimal(dr.Cells["Subtotal"].Value);
						this.txtDescIvaTotal.Value = Convert.ToDecimal(this.txtDescIvaTotal.Value) + Convert.ToDecimal(dr.Cells["DescuentoArt"].Value);
					}
					else
					{				
						this.txtIva0.Value = Convert.ToDecimal(this.txtIva0.Value) + Convert.ToDecimal(dr.Cells["Subtotal"].Value);
						this.txtDescIva0Total.Value = Convert.ToDecimal(this.txtDescIva0Total.Value) + Convert.ToDecimal(dr.Cells["DescuentoArt"].Value);
					}
				}
								
				decimal Subtotal = Convert.ToDecimal(this.txtIva.Value);
				decimal Descuento = Convert.ToDecimal(this.txtDescIvaTotal.Value);
				decimal IVA = Math.Round((Subtotal - Descuento) * dPorcIva /100, 2);

				this.txtIvaTotal.Value = IVA;

				this.txtTotal.Value = Convert.ToDecimal(this.txtIva0.Value)
					+ Convert.ToDecimal(this.txtIva.Value)
					- Convert.ToDecimal(this.txtDescIva0Total.Value) 
					- Convert.ToDecimal(this.txtDescIvaTotal.Value)
					+ Convert.ToDecimal(this.txtIvaTotal.Value);
			} 
			#endregion Desplazamiento de grilla
		}
		private void btnCancelar_Click(object sender, System.EventArgs e)
		{
			if (bEdicion && !bNuevo)
			{
				string sSQLBloqueo = string.Format("Exec RegistraBloqueoTransacciones {0}, 0, {1}, 1", idBloqueaTransacciones, (int)this.txtIdCompra.Value);
				Funcion.EjecutaSQL(cdsSeteoF, sSQLBloqueo, true);
			}

	

			this.txtIdCompra.Value = 0;
			this.txtIdCliente.Value = 0;
			this.txtEstado.Value = 0;
			this.txtIdProyecto.Value = 0;
			this.txtIdGrupoCliente.Value = 0;
			this.txtIdTipoRuc.Value = 0;
			this.txtContadoCredito.Value = 0;
			this.txtIdComprobante.Value = 0;
			this.txtIdCredTributario.Value = 0;


			this.cmbFormaPago.Enabled = false;
			this.cmbFormaPago.Value = System.DBNull.Value;


			this.cmbBodega.Value = System.DBNull.Value;
			this.dtFecha.Value = System.DBNull.Value;
			this.dtFecha.CalendarInfo.MinDate = Funcion.dtEscalarSQL(cdsSeteoF, "Select CONVERT(Date, FechaInicio) From SeteoF");

			this.txtNumero.Text = "";

			this.lblEstado.Text = "";
			this.lblUsuario.Text = "";
			
			this.ultraGrid1.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "Exec ConsultaDetalleOrdenCompra 0");

			this.txtNotas.Text = "";
			
			this.txtIva0.Value = 0.00m;
			this.txtIva.Value = 0.00m;
			this.txtDescIva0Total.Value = 0.00m;
			this.txtDescIvaTotal.Value = 0.00m;
			this.txtIvaTotal.Value = 0.00m;
			this.txtTotal.Value = 0.00m;

			bAnulado = false;
			bEliminaAlValidar = false;


			this.cmbBodega.Enabled = false;
			this.dtFecha.Enabled = false;

			this.txtNumero.Enabled = false;

			
			FuncionesProcedimientos.EstadoGrids(false, this.ultraGrid1);
			
			this.txtNotas.Enabled = false;

			if (miAcceso.Nuevo) this.btnNuevo.Enabled = true;
			if (miAcceso.Buscar) this.btnConsultar.Enabled = true;			
			this.btnEditar.Enabled = false;
			this.btnImprimir.Enabled = false;
			this.btnGuardar.Enabled = false;			

			this.btnAnular.Enabled = false;
			this.btnCancelar.Enabled = false;
			this.btnAnular.Text = "Anular";
			this.btnDistribucion.Enabled = false;

			this.ultraGrid1.DisplayLayout.Bands[0].Columns["CantAereo"].Hidden = true;
			this.ultraGrid1.DisplayLayout.Bands[0].Columns["Valores"].Hidden = true;
			this.ultraGrid1.DisplayLayout.Bands[0].Columns["Courier"].Hidden = true;

			idBloqueaTransacciones = 0;
			//dPorcIva = 0.00m;

			if (iOrigen == 1) this.Close();
		}

		private void cdsSeteoF_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			cdsSeteoF.Schema.Connections[0].ConnectionString = miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);	
		
		}

		private void btnNuevo_Click(object sender, System.EventArgs e)
		{
			this.btnCancelar_Click(sender, e);

			this.txtEstado.Value = 1;						
			
			//FuncionesProcedimientos.FechasDocumentos(this.dtFecha);//this.dtFecha.Value = DateTime.Today;
			FuncionesProcedimientos.RetornaFechaServidor(this.dtFecha, cdsSeteoF, false);
			dPorcIva = Funcion.decEscalarSQL(cdsSeteoF, string.Format("Select dbo.RetornaImpuestoIVAVigente('{0}', 1)", Convert.ToDateTime(this.dtFecha.Value).ToString("yyyyMMdd")));

			
			this.cmbBodega.Enabled = true;
			this.txtNotas.Enabled = true;
		
			FuncionesProcedimientos.EstadoGrids(true, this.ultraGrid1);
			
			this.btnNuevo.Enabled = false;
			this.btnConsultar.Enabled = false;
			this.btnEditar.Enabled = false;
			this.btnImprimir.Enabled = false;
			this.btnGuardar.Enabled = true;
	
			this.btnAnular.Enabled = false;
			this.btnCancelar.Enabled = true;
			if (miAcceso.BtnDistribucion) this.btnDistribucion.Enabled = true;

			this.cmbFormaPago.Value = 6;
			this.cmbFormaPago.Enabled = true;
				
			bNuevo = true;
			bEdicion = true;
		}

		private void dtFecha_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
			if (e.KeyChar == 8) e.Handled = true;
		}

		private void dtFecha_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Delete) e.Handled = true;

			if (e.KeyCode == Keys.Enter)
				this.cmbBodega.Focus();
		}

		private void dtFecha_Validated(object sender, System.EventArgs e)
		{
			if (bEdicion)
			{				

				dPorcIva = Funcion.decEscalarSQL(cdsSeteoF, string.Format("Select dbo.RetornaImpuestoIVAVigente('{0}', 1)", Convert.ToDateTime(this.dtFecha.Value).ToString("yyyyMMdd")));
				ActualizaIVA();
			}
		}

		private void ActualizaIVA()
		{
			if (this.ultraGrid1.Rows.Count > 0)
			{		
				if (this.dtFecha.Value != System.DBNull.Value)
				{
					foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.ultraGrid1.Rows.All)
					{
						string sSQL = string.Format("Exec RetornaIVAClienteArticulo {0}, {1}, '{2}'", (int)this.txtIdGrupoCliente.Value, (int)dr.Cells["idArticulo"].Value, Convert.ToDateTime(this.dtFecha.Value).ToString("yyyyMMdd"));
						dr.Cells["Impuesto"].Value = Funcion.decEscalarSQL(cdsSeteoF, sSQL); 					
					}
				}
			}			
		}

		private void dtFecha_Validating(object sender, System.ComponentModel.CancelEventArgs e)
		{
			if (bEdicion && this.dtFecha.Value == DBNull.Value)
			{
				MessageBox.Show("Ingrese la Fecha del Documento", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				this.dtFecha.Focus();
				e.Cancel = true;				
			}
			else if (Funcion.ValidaPeriodo((DateTime)this.dtFecha.Value, cdsSeteoF, "Compra"))
			{
				DateTime dtFecha = (DateTime)this.dtFecha.Value;
				MessageBox.Show(string.Format("No puede ingresar Pedidos de Presupuesto en esta Fecha, El Período '{0}' esta Cerrado, \n\n Consulte al Administrador", dtFecha.ToString("MMMM").ToUpper()), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				this.dtFecha.Focus();
				e.Cancel = true;					
			}
		}

		private void dtFecha_ValueChanged(object sender, System.EventArgs e)
		{
			if (this.dtFecha.Value != System.DBNull.Value)
			{	
				ListaBodegas(this.cmbBodega, MenuLatinium.IdUsuario,  (DateTime)this.dtFecha.Value,  cdsSeteoF);
				this.cmbProyecto1.DataSource = cmbBodega.DataSource;
				if (this.cmbBodega.Value != System.DBNull.Value)
				{	
					ListaSoliictante(this.cmbFormaPago, MenuLatinium.IdUsuario,  (DateTime)this.dtFecha.Value,  cdsSeteoF);
				}
			}
		}

		private void ListaBodegas (Infragistics.Win.UltraWinGrid.UltraCombo uCombo, int idUsuario,  DateTime dtFecha,  C1.Data.C1DataSet ds)
		{
			uCombo.DataSource = Funcion.dvProcedimiento(ds, string.Format("Exec PresupuestoListaCCPed {0}, '{1}'", idUsuario,  dtFecha.ToString("yyyyMMdd")));
		}
		private void ListaSoliictante (Infragistics.Win.UltraWinGrid.UltraCombo uCombo, int idUsuario,  DateTime dtFecha,  C1.Data.C1DataSet ds)
		{
			uCombo.DataSource = Funcion.dvProcedimiento(ds, string.Format("Exec [ConsultaPErsonalSolicita] {0}, '{1}'", idUsuario,  dtFecha.ToString("yyyyMMdd")));
		}

		private void ultraGrid1_AfterRowInsert(object sender, Infragistics.Win.UltraWinGrid.RowEventArgs e)
		{
			if (bEdicion)
			{
				Facturacion.IniciaValoresColumnas(sender, e, this.ultraGrid1, 2, 1);

				e.Row.Cells["idImportacion"].Value = 0;
				e.Row.Cells["idTipoCalculo"].Value = 0;
			}
		}

		private void ultraGrid1_InitializeLayout(object sender, Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs e)
		{
			#region Formato numerico
			CultureInfo cultureENUS = CultureInfo.CreateSpecificCulture("ja-jp");
			string stMoneda = "#,##0.";
			string stMascara = "nnn,nnn,nnn.";
			string stCeros = "0000000000";
			string stNumero = "nnnnnnnnnn";
			string stFormato = "";
			string stInput = "";
			#endregion Formato numerico
			#region Decimales para el total
			stFormato = stMoneda + stCeros.Substring(0, 2);
			stInput = stMascara + stNumero.Substring(0, 2);
			e.Layout.Bands[0].Columns["SubTotal"].Format = stFormato; 
			e.Layout.Bands[0].Columns["DescuentoArt"].Format = stFormato;

			this.txtDescIvaTotal.FormatString = stFormato;
			this.txtIvaTotal.FormatString = stFormato;
			this.txtIva.FormatString = stFormato;
			this.txtIva0.FormatString = stFormato;
			this.txtTotal.FormatString = stFormato;
			this.txtDescIva0Total.FormatString = stFormato;

			this.txtDescIvaTotal.MaskInput = stInput;
			this.txtIvaTotal.MaskInput = stInput;
			this.txtIva.MaskInput = stInput;
			this.txtIva0.MaskInput = stInput;
			this.txtTotal.MaskInput = stInput;
			this.txtDescIva0Total.MaskInput = stInput;
			#endregion
			
			FuncionesProcedimientos.FormatoGrid(e, "Precio", 5);
			FuncionesProcedimientos.FormatoGrid(e, "Impuesto", 2);						
			FuncionesProcedimientos.FormatoGrid(e, "Valores", 2);						
			FuncionesProcedimientos.FormatoGrid(e, "DescuentoPorc", 5);
			FuncionesProcedimientos.FormatoGrid(e, "Courier", 5);
			FuncionesProcedimientos.FormatoGrid(e, "Peso", 5);
			FuncionesProcedimientos.FormatoGrid(e, "Arancel", 5);
			FuncionesProcedimientos.FormatoGrid(e, "DescNC", 5);
			FuncionesProcedimientos.FormatoGrid(e, "Subtotal", 5);	
		}

		private void ultraGrid1_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (bEdicion)
			{
				if (ultraGrid1.ActiveCell != null) 
				{					
					FuncionesProcedimientos.DesplazamientoGrids(sender, e, this.ultraGrid1);	

					#region Articulos
					if (e.KeyCode == Keys.F3)
					{		
						try
						{	
							if (this.ultraGrid1.ActiveCell.Column.Key.ToString().ToUpper() == "CODIGO" || this.ultraGrid1.ActiveCell.Column.Key.ToString().ToUpper() == "ARTICULO")
							{									
								int idDetCompra = 0;
								if (this.ultraGrid1.ActiveRow.Cells["idDetCompra"].Value != System.DBNull.Value) idDetCompra = (int)this.ultraGrid1.ActiveRow.Cells["idDetCompra"].Value;
								int idArticulo = 0;
								if (this.ultraGrid1.ActiveRow.Cells["idArticulo"].Value != System.DBNull.Value) idArticulo = (int)this.ultraGrid1.ActiveRow.Cells["idArticulo"].Value;

								//								string sSQL = string.Format("Exec ValidaConteoSeriales {0}, {1}", idDetCompra, idArticulo);
								//								int iConteo = Funcion.iEscalarSQL(cdsSeteoF, sSQL);
								//										
								//								if (!bNuevo && (iConteo > 0 && !miAcceso.BModArtSerializado))
								//								{
								//									MessageBox.Show(string.Format("Este Articulo tiene {0} Seriales Registradas no puede Modificarlo\n\nConsulte con el Administrador", iConteo), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);										
								//									return;
								//								}
									
								FuncionesInventario.AbreBusquedaArticulos(this.ultraGrid1, 1, 80, 6, Convert.ToDateTime(this.dtFecha.Value), false);									
							}						
						}
						catch(System.Exception ex)
						{
							MessageBox.Show(ex.Message, "Error en Ingreso de Productos", MessageBoxButtons.OK, MessageBoxIcon.Error);
						}
					}
					#endregion Articulos
				}
			}
		}

		private void ultraGrid1_AfterCellUpdate(object sender, Infragistics.Win.UltraWinGrid.CellEventArgs e)
		{
			if (bEdicion)
			{
				if (!bActualiza) return;
						
				if (bActualiza)
				{
					bActualiza = false;
					#region Verifica Cambio
					int VerificaCambio = 0;
					#region Codigo o Articulo
					if (e.Cell.Column.ToString() == "Codigo")// || e.Cell.Column.ToString() == "Articulo")
					{
						#region Variables
						string stCod = "";
						int IdArt = 0;
						int iCuenta = 0;
						if (bRastreo) Rastreo("Inicio Cambio de " + e.Cell.Column.ToString());
						#endregion Variables
						#region Codigo
						if (e.Cell.Column.ToString() == "Codigo")
						{
							if (e.Cell.Row.Cells["Codigo"].Value.ToString().Length == 0) e.Cell.Row.Cells["Codigo"].Value = e.Cell.OriginalValue.ToString();

							stCod = e.Cell.Row.Cells["Codigo"].Value.ToString();						
						}
						#endregion Codigo
						#region Articulo
						if (e.Cell.Column.ToString() == "Articulo")
						{
							stCod = e.Cell.Row.Cells["Articulo"].Value.ToString();
						}
						#endregion Articulo

						int iTipoArticulo = 0;
								
						#region Cuenta coincidencias del codigo
						string stSelect = string.Format("Exec VentaArticuloCuenta '{0}', {1}, {2}, {3}, 0", stCod, 80, 1, iTipoArticulo);
						if (bRastreo) Rastreo(String.Format("Cuenta Coincidencias de {0}: {1}", e.Cell.Column.ToString(), stCod));
						iCuenta = Funcion.iEscalarSQL(cdsSeteoF, stSelect, true);
						#endregion Cuenta coincidencias del codigo
						#region Mas de una coincidencia o ninguna muestra la lista
						//iCuenta = 1;
						if (iCuenta != 1)
						{ 
							if (bRastreo) Rastreo(string.Format("# de Coincidencias != de 1 en {0} : {1}", e.Cell.Column.ToString(), stCod));
					
							int Bodega = (int) this.cmbBodega.Value;

							using (BuscaArticulos miBusqueda = new BuscaArticulos(stCod, Bodega, 80))
							{
								if (DialogResult.OK == miBusqueda.ShowDialog())
								{
									if (bRastreo) Rastreo("Articulo escogido de la lista");
									IdArt = miBusqueda.idArticulo;
								}
								else
								{
									e.Cell.Row.Cells["idArticulo"].Value = 0;
									e.Cell.Row.Cells["Precio"].Value = 0.00m;
									e.Cell.Row.Cells["Codigo"].Value = "";
									e.Cell.Row.Cells["Articulo"].Value = "";
									cmdTotal();

									if (bRastreo) Rastreo("Cancela la selección de Artículos en lista");
									bActualiza = true;											
								}
							}
						}
							#endregion Mas de una coincidencia o ninguna muestra la lista
							#region Encuentra un solo articulo Asigna Codigo y Unidad e idArticulo
						else 
						{
							if (bRastreo) Rastreo("Un codigo de Articulo Encontrado");
							stSelect = string.Format("Exec VentaArticuloEncontrado '{0}', {1}, {2}, '', '', 0, 0, 0", stCod, 1, 80);
							IdArt = Funcion.iEscalarSQL(cdsSeteoF, stSelect, true);			
						}
						#endregion Encuentra un solo articulo
						#region Asignaciones
						if (bRastreo) Rastreo("Asigna Codigo, Articulo, Unidad, idArticulo");
						stSelect = string.Format("Exec CargaInformArticuloCompra {0}, {1}, 80, '{2}', 1", IdArt, 1, Convert.ToDateTime(this.dtFecha.Value).ToString("yyyyMMdd"));
						//@idArticulo Int, @idClienteGrupo Int, @idTipoFactura Int, @Fecha Datetime, @Fiscal Bit
						SqlDataReader dr = Funcion.rEscalarSQL(cdsSeteoF, stSelect, true);
						dr.Read();
						stCod = dr.GetString(0);
						e.Cell.Row.Cells["Articulo"].Value = dr.GetString(1);
						e.Cell.Row.Cells["idUnidad"].Value = dr.GetInt32(2);
						e.Cell.Row.Cells["idArticulo"].Value = IdArt;
						e.Cell.Row.Cells["Precio"].Value = dr.GetDecimal(3);
						e.Cell.Row.Cells["Impuesto"].Value = dr.GetDecimal(4);						
						e.Cell.Row.Cells["Peso"].Value = dr.GetDecimal(5);						
						e.Cell.Row.Cells["Arancel"].Value = dr.GetDecimal(6);
						dr.Close();
						e.Cell.Row.Cells["Codigo"].Value = stCod;

				
						VerificaCambio = 2; // Verificaciones y Cambio de precio
						#endregion Asignaciones
					} 
					#endregion Codigo o articulo
						
					#region Cantidad
					if (e.Cell.Column.ToString() == "Cantidad")
					{
						if (e.Cell.Row.Cells["Cantidad"].Value == System.DBNull.Value) e.Cell.Row.Cells["Cantidad"].Value = (int)e.Cell.OriginalValue;
						
						if (e.Cell.Row.Cells["idArticulo"].Value == System.DBNull.Value)
						{							
							if ((int)e.Cell.Row.Cells["Cantidad"].Value > 1)
							{
								e.Cell.Value = 0;
								MessageBox.Show("Seleccione un Producto antes de Ingresar Cantidad", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Error);
							}
						}

						e.Cell.Row.Cells["DescNC"].Value = 0.00m;

						#region Valida Seriales ingresadas
						if (!bNuevo)
						{
							int idDetCompra = 0;
							if (e.Cell.Row.Cells["idDetCompra"].Value != System.DBNull.Value) idDetCompra = (int)e.Cell.Row.Cells["idDetCompra"].Value;
							int idArticulo = 0;
							if (e.Cell.Row.Cells["idArticulo"].Value != System.DBNull.Value) idArticulo = (int)e.Cell.Row.Cells["idArticulo"].Value;

							//							string sSQL = string.Format("Exec ValidaConteoSeriales {0}, {1}", idDetCompra, idArticulo);
							//							int iConteo = Funcion.iEscalarSQL(cdsSeteoF, sSQL);
							//
							//							int iCantidadM = (int)e.Cell.Row.Cells["Cantidad"].Value;
							// 
							//							if (iCantidadM < iConteo)
							//							{
							//								e.Cell.Row.Cells["Cantidad"].Value = (int)e.Cell.OriginalValue;
							//								MessageBox.Show(string.Format("El Artículo '{0}' tiene {1} seriales registradas", e.Cell.Row.Cells["Codigo"].Value.ToString().Trim(), iConteo), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Error);						
							//							}								
						}
						#endregion Valida Seriales ingresadas
						this.cmdTotal();
						VerificaCambio = 1;
					}
					#endregion Cantidad
					#region Cantidad Aereo
					if (e.Cell.Column.ToString() == "CantAereo")
					{
						if (e.Cell.Row.Cells["CantAereo"].Value == System.DBNull.Value) e.Cell.Row.Cells["CantAereo"].Value = (int)e.Cell.OriginalValue;

						string sSQLCantCmp = string.Format("Select dc.Cantidad From DetCompra dc Inner Join Compra c On c.idCompra = dc.idCompra Where dc.idOrigen = {0} And c.idTipoFactura = 4", (int)e.Cell.Row.Cells["idDetCompra"].Value);
						int dCantCMP = Funcion.iEscalarSQL(cdsSeteoF, sSQLCantCmp);
						int dCantAer = 0;
						if (e.Cell.Row.Cells["CantAereo"].Value != System.DBNull.Value) dCantAer = (int)e.Cell.Row.Cells["CantAereo"].Value;

						if ((dCantCMP - dCantAer) < 0)
						{
							e.Cell.Row.Cells["CantAereo"].Value = 0;
							MessageBox.Show(string.Format("La Cantidad Máxima es {0}", dCantCMP), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Error);
						}
						else e.Cell.Row.Cells["Cantidad"].Value = dCantCMP - dCantAer;					
					}
					#endregion Cantidad Aereo								
					#endregion Fin de Verificacion de Cambio
					#region DescuentoPorc o DescuentoArt
					if (e.Cell.Column.ToString() == "DescuentoPorc")
					{
						bActualiza = false;
						if (e.Cell.Row.Cells["DescuentoPorc"].Value == System.DBNull.Value) e.Cell.Row.Cells["DescuentoPorc"].Value = (decimal)e.Cell.OriginalValue;
						
						if (e.Cell.Row.Cells["idArticulo"].Value == System.DBNull.Value)
						{							
							if ((decimal)e.Cell.Row.Cells["DescuentoPorc"].Value > 0.00m)
							{
								e.Cell.Value = 0.00m;
								MessageBox.Show("Seleccione un Producto antes de Ingresar el Descuento", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Error);
							}
						}

						if ((decimal)e.Cell.Row.Cells["DescuentoPorc"].Value > 100.00m)
						{								
							e.Cell.Value = 0.00m;
							MessageBox.Show("El Porcentaje de Descuento Maximo es 100%", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Error);
						}
						
						cmdTotal();
						bActualiza = true;
					}

					if (e.Cell.Column.ToString() == "DescuentoArt")
					{
						bActualiza = false;
						cmdTotal();
						bActualiza = true;
					}
					#endregion DescuentoPorc o DescuentoArt
					#region Precio
					
					if (e.Cell.Column.ToString() == "Precio")
					{
						bActualiza = false;
						if (e.Cell.Row.Cells["Precio"].Value == System.DBNull.Value) e.Cell.Row.Cells["Precio"].Value = (decimal)e.Cell.OriginalValue;
						
						if (e.Cell.Row.Cells["idArticulo"].Value == System.DBNull.Value)
						{							
							if ((decimal)e.Cell.Row.Cells["Precio"].Value > 0.00m)
							{
								e.Cell.Row.Cells["DescNC"].Value = 0.00m;
								e.Cell.Value = 0.00m;
								MessageBox.Show("Seleccione un Producto antes de Ingresar el Precio", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Error);
							}
						}	
					
						cmdTotal();
						bActualiza = true;
							
					}
					#endregion Precio					
					#region IVA
					if (e.Cell.Column.ToString() == "Impuesto")
					{
						bActualiza = false;		
						if (e.Cell.Row.Cells["Impuesto"].Value == System.DBNull.Value) e.Cell.Row.Cells["Impuesto"].Value = Convert.ToDecimal(e.Cell.OriginalValue);
						
						if (e.Cell.Row.Cells["idArticulo"].Value == System.DBNull.Value)
						{							
							if ((decimal)e.Cell.Row.Cells["Impuesto"].Value > 0.00m)
							{
								e.Cell.Value = 0.00m;
								MessageBox.Show("Seleccione un Producto antes de Ingresar el Porcentaje de I.V.A.", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Error);
							}
						}
						if ((decimal)e.Cell.Row.Cells["Impuesto"].Value > 0.00m)
						{								
							if ((decimal)e.Cell.Row.Cells["Impuesto"].Value != dPorcIva)
							{

								e.Cell.Value = (decimal)e.Cell.OriginalValue;;
								MessageBox.Show(string.Format("El Porcentaje de I.V.A. debe ser 0% o {0}%", dPorcIva), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Error);
							}
						}
						
						cmdTotal();
						bActualiza = true;
				
					}
					#endregion IVA
					#region Descuento Promocion
					if (e.Cell.Column.ToString() == "DescNC")
					{
						bActualiza = false;
						if (e.Cell.Row.Cells["DescNC"].Value == System.DBNull.Value) e.Cell.Row.Cells["DescNC"].Value = (decimal)e.Cell.OriginalValue;
						
						if (e.Cell.Row.Cells["idArticulo"].Value == System.DBNull.Value)
						{							
							if ((decimal)e.Cell.Row.Cells["DescNC"].Value > 0.00m)
							{
								e.Cell.Value = 0.00m;
								MessageBox.Show("Seleccione un Producto antes de Ingresar el Descuento Promoción", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Error);
							}
						}		
				
						cmdTotal();
						bActualiza = true;
					}
					#endregion Descuento Promocion
					#region Valores
					if (e.Cell.Column.ToString() == "Valores")
					{
						bActualiza = false;	
						if (e.Cell.Row.Cells["Valores"].Value == System.DBNull.Value) e.Cell.Row.Cells["Valores"].Value = (decimal)e.Cell.OriginalValue;
						
						if (e.Cell.Row.Cells["idArticulo"].Value == System.DBNull.Value)
						{							
							if ((decimal)e.Cell.Row.Cells["Valores"].Value > 0.00m)
							{
								e.Cell.Value = 0.00m;
								MessageBox.Show("Seleccione un Producto antes de Ingresar Valores", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Error);
							}
						}
						bActualiza = true;
					}
					#endregion Valores
				}
				bActualiza = true;
			}
		}

		private void ultraGrid1_AfterRowsDeleted(object sender, System.EventArgs e)
		{
			this.cmdTotal();
		}

		private void ultraGrid1_BeforeCellUpdate(object sender, Infragistics.Win.UltraWinGrid.BeforeCellUpdateEventArgs e)
		{
			if (e.Cell.Column.ToString() == "Codigo")
			{	
				int idDetCompra = 0;
				if (e.Cell.Row.Cells["idDetCompra"].Value != System.DBNull.Value) idDetCompra = (int)e.Cell.Row.Cells["idDetCompra"].Value;
				int idArticulo = 0;
				if (e.Cell.Row.Cells["idArticulo"].Value != System.DBNull.Value) idArticulo = (int)e.Cell.Row.Cells["idArticulo"].Value;

				//				string sSQL = string.Format("Exec ValidaConteoSeriales {0}, {1}", idDetCompra, idArticulo);
				//				int iConteo = Funcion.iEscalarSQL(cdsSeteoF, sSQL);
				//										
				//				if (!bNuevo && (iConteo > 0 && !miAcceso.BModArtSerializado))
				//				{
				//					MessageBox.Show(string.Format("Este Articulo tiene {0} Seriales Registradas no puede Modificarlo\n\nConsulte con el Administrador", iConteo), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);										
				//					e.Cancel = true;
				//				}
			}
		}

		private void ultraGrid1_BeforeRowsDeleted(object sender, Infragistics.Win.UltraWinGrid.BeforeRowsDeletedEventArgs e)
		{
			if (bEdicion)
			{
				#region Elimina
				if (DialogResult.Yes ==	MessageBox.Show("¿Esta seguro de borrar ?", "Point Technology", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2))
				{				
					for (int i=0; i < e.Rows.Length; i++)
					{	
						if ((int)e.Rows[i].Cells["idDetCompra"].Value > 0)
						{
							string sSQL = string.Format("Delete From Detcompra where idDetcompra = {0} ", (int)e.Rows[i].Cells["idDetCompra"].Value);
							Funcion.EjecutaSQL(cdsSeteoF, sSQL, true);

							e.DisplayPromptMsg = false;
						}
						else e.DisplayPromptMsg = false;
					}				
				}
				else e.Cancel = true;
				#endregion Elimina	
			}
		}
		
		public void ConsultaMaestro(int IdCompra)
		{
			try
			{
				if (miAcceso.Nuevo) this.btnNuevo.Enabled = true;
				if (miAcceso.Buscar) this.btnConsultar.Enabled = true;			
				this.btnEditar.Enabled = false;
				this.btnImprimir.Enabled = false;
				this.btnGuardar.Enabled = false;			
		
				this.btnAnular.Enabled = false;				
				
				int IdBodegaC = 0;

				int idImportacion = 0;
				string sSQL = string.Format("Exec [ConsultaMaestroPresupuestoPedido] {0}", IdCompra);
				SqlDataReader dr = Funcion.rEscalarSQL(cdsSeteoF, sSQL, true);
				dr.Read();
				if (dr.HasRows)
				{
					this.txtIdCompra.Value = IdCompra;
					this.txtContadoCredito.Value = dr.GetInt32(1);
					this.txtIdComprobante.Value = dr.GetInt32(2); 
					this.txtIdCredTributario.Value = dr.GetInt32(3);
					//this.chkConsignacion.Checked = dr.GetBoolean(4);
					this.cmbFormaPago.Value = dr.GetInt32(5);
					this.txtIdTipoRuc.Value = 0;
					this.txtIdGrupoCliente.Value = 0;
					//this.txtRuc.Text = dr.GetString(8);
					//this.txtNombre.Text = dr.GetString(9);
					IdBodegaC = dr.GetInt32(10);
					this.dtFecha.Value = dr.GetDateTime(11);
					//if (dr.GetValue(12) != System.DBNull.Value) this.dtFechaLlegada.Value = dr.GetDateTime(12);	
					//if (dr.GetValue(13) != System.DBNull.Value) this.dthorallegada.Value = dr.GetDateTime(13);	
					this.lblUsuario.Text = dr.GetString(14);
					this.txtNumero.Text = dr.GetString(15);
					if (dr.GetValue(16) != System.DBNull.Value) idImportacion = dr.GetInt32(16);
					this.txtEstado.Value = dr.GetInt32(17);
					this.txtNotas.Text = dr.GetValue(18).ToString();
					this.txtIva0.Value = dr.GetDecimal(19);
					this.txtDescIva0Total.Value = dr.GetDecimal(20);
					this.txtIva.Value = dr.GetDecimal(21);
					this.txtDescIvaTotal.Value = dr.GetDecimal(22);
					this.txtIvaTotal.Value = dr.GetDecimal(23);
					this.txtTotal.Value = dr.GetDecimal(24);
					bAnulado = dr.GetBoolean(25);
					
				}
				dr.Close();
				this.cmbBodega.Value= IdBodegaC;

				if ((int)this.txtIdGrupoCliente.Value == 2 || (int)this.txtIdGrupoCliente.Value == 3)
				{
					this.ultraGrid1.DisplayLayout.Bands[0].Columns["CantAereo"].Hidden = false;
					this.ultraGrid1.DisplayLayout.Bands[0].Columns["Valores"].Hidden = false;
				}

			

				this.ultraGrid1.DataSource = Funcion.dvProcedimiento(cdsSeteoF, string.Format("Exec ConsultaDetalleOrdenCompra {0}", IdCompra));


					
				if ((int)this.txtEstado.Value == 1) this.lblEstado.Text = "PENDIENTE";				
				if ((int)this.txtEstado.Value == 2) this.lblEstado.Text = "APROBADO";
				if ((int)this.txtEstado.Value == 3) this.lblEstado.Text = "RECHAZADO";




				if (miAcceso.Editar)
				{
					if ((int)this.txtEstado.Value == 1) this.btnEditar.Enabled = true;
					else  this.btnEditar.Enabled = false;
				}

				if (miAcceso.Imprimir) this.btnImprimir.Enabled = true;

				if (miAcceso.Anular && !bAnulado)
				{
					if ((int)this.txtEstado.Value == 1) this.btnAnular.Enabled = true;
					else if ((int)this.txtEstado.Value == 3 && miAcceso.BAnularProcesado)this.btnAnular.Enabled = true;
				}

				this.btnCancelar.Enabled = true;
			}			
			catch(Exception Exc)
			{
				MessageBox.Show(string.Format("Error al Consultar: {0}", Exc), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Error);
			}	
		}


		private void btnConsultar_Click(object sender, System.EventArgs e)
		{
			this.btnCancelar_Click(sender, e);
			
			using (frmBuscaPResupuestoCXC Busqueda = new frmBuscaPResupuestoCXC())
				if (DialogResult.OK == Busqueda.ShowDialog())
					this.ConsultaMaestro((int)Busqueda.uGridCreditos.ActiveRow.Cells["idCompra"].Value);
		}

		private void btnEditar_Click(object sender, System.EventArgs e)
		{
			this.ConsultaMaestro((int)this.txtIdCompra.Value);
			
			if (bAnulado)
			{
				MessageBox.Show("No puede editar registro ANULADO", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Error);
				return;
			}

			if ((int)this.txtEstado.Value == 3 && !miAcceso.BEditarProcesado)
			{
				MessageBox.Show("No puede editar una orden de compra PROCESADA, Consulte con el Administrador", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Error);				
				return;
			}

			if ((int)this.txtEstado.Value == 9 || (int)this.txtEstado.Value == 13 || (int)this.txtEstado.Value == 12)
			{
				MessageBox.Show("No Puede Editar una orden de compra FACTURADA", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Error);				
				return;
			}
			
			//if (Funcion.ValidaBloqueoERMEnTransacciones((int)this.txtIdCompra.Value, cdsSeteoF)) return;
			
			//	if (Funcion.ValidaBloqueo((int)this.txtIdCompra.Value, 1, cdsSeteoF)) return;
				
			dPorcIva = Funcion.decEscalarSQL(cdsSeteoF, string.Format("Select dbo.RetornaImpuestoIVAVigente('{0}', 1)", Convert.ToDateTime(this.dtFecha.Value).ToString("yyyyMMdd")));

			string sSQLBloqueo = string.Format("Exec RegistraBloqueoTransacciones 0, 1, {0}, 1", (int)this.txtIdCompra.Value);
			idBloqueaTransacciones = Funcion.iEscalarSQL(cdsSeteoF, sSQLBloqueo, true);

			#region Controles
			//this.txtRuc.Enabled = true;

			if (miAcceso.CambiarFecha) this.dtFecha.Enabled = true;			
			 
					
			if ((int)this.txtEstado.Value == 1) this.cmbBodega.Enabled = true;
			FuncionesProcedimientos.EstadoGrids(true, this.ultraGrid1);
			this.txtNotas.Enabled = true;
			
			this.btnNuevo.Enabled = false;
			this.btnConsultar.Enabled = false;
			this.btnEditar.Enabled = false;
			this.btnImprimir.Enabled = false;
			this.btnGuardar.Enabled = true;

			this.btnAnular.Enabled = true;
			this.btnCancelar.Enabled = true;


			this.cmbFormaPago.Enabled = true;

			bNuevo = false;
			bEdicion = true;
			#endregion Controles			
		}

		private void btnGuardar_Click(object sender, System.EventArgs e)
		{
			#region Elimina Filas Vacias
		
			bEliminaAlValidar = true;
			Validacion.vEliminaFilasVacias(this.ultraGrid1, "idArticulo");
			bEliminaAlValidar = false;

			
			#endregion Elimina Filas Vacias

			this.ActualizaIVA();
			this.cmdTotal();
		

			
			if (!Validacion.ValidaTotales((decimal)this.txtTotal.Value)) return;
			

			if (!Validacion.vbFechaEnDocumentos(this.dtFecha, "Ingrese la fecha del documento", true, "Compra", cdsSeteoF)) return;

			//if (!Validacion.vbComboVacio(this.cmbBodega, "Seleccione una Bodega")) return;

			if (!Validacion.vbComboVacio(this.cmbFormaPago, "Seleccione El solicitante")) return;
	
			if (!Validacion.vbGrids(this.ultraGrid1, "Codigo", 2)) return;	

			#region Valida Centro de Costo
			foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.ultraGrid1.Rows.All)
			{
				if (dr.Cells["idProyecto"].Value == DBNull.Value)
				{
					MessageBox.Show(string.Format("Seleccione el Centro de Costo : '{0}'", dr.Cells["Articulo"].Value.ToString().Trim()), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Error);
					this.ultraGrid1.ActiveRow = dr;
					this.ultraGrid1.ActiveRow.Selected = true;
					return;
				}
			}
			#endregion Valida Centro de Costo		

			foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.ultraGrid1.Rows.All)
			{
				int iValidaPresupuesto = 0;
				iValidaPresupuesto = Funcion.iEscalarSQL(cdsSeteoF, string.Format("Exec [PresupuestoValidaValores] '{0}', {1}, {2}, {3}",
					Convert.ToDateTime(this.dtFecha.Value).ToString("yyyyMMdd"), 0, (int) dr.Cells["idProyecto"].Value, 2));	

				if ( iValidaPresupuesto == 0 )
				{
					MessageBox.Show("El presupuesto para la bodega y periodo seleccionado \n no esta creado comuniquese con el area de contabilidad.", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
					this.ultraGrid1.ActiveRow = dr;
					this.ultraGrid1.ActiveRow.Selected = true;
					return;
				}
			}

			int icont = 0;
			foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.ultraGrid1.Rows.All)
			{
				int idCuponesWeb 	= Funcion.iEscalarSQL(cdsSeteoF, string.Format("Exec [PresupuestoValidaValores] '{0}', {1}, {2},{3}",
					Convert.ToDateTime(this.dtFecha.Value).ToString("yyyyMMdd"), 0, (int) dr.Cells["idProyecto"].Value, 2));	

				if ( idCuponesWeb == 0)
				{
					icont = icont +1 ; 
					this.ultraGrid1.ActiveRow = dr;
					this.ultraGrid1.ActiveRow.Selected = true;
				}
			}

			if ( icont > 0)
			{
				MessageBox.Show("Las cuentas seleccionadas no estan asignadas  \n a la bodega seleccionada.", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				return;
			}

			#region valida cuenta con valores en 0
			int icontCero = 0;
			foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.ultraGrid1.Rows.All)
			{
				int idCuponesWeb 	= Funcion.iEscalarSQL(cdsSeteoF, string.Format("Exec [PresupuestoValidaValores] '{0}', {1}, {2},{3}",
					Convert.ToDateTime(this.dtFecha.Value).ToString("yyyyMMdd"), (int)dr.Cells["idArticulo"].Value, dr.Cells["idProyecto"].Value , 1));	

				if ( idCuponesWeb > 0)
				{
					icontCero = icontCero +1 ; 
					this.ultraGrid1.ActiveRow = dr;
					this.ultraGrid1.ActiveRow.Selected = true;
				}
			}

			if ( icontCero > 0)
			{
				MessageBox.Show("Las cuentas seleccionadas tienen un presupuesto de 0 \n Comuniquese con el área de contabilidad.", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				return;
			}
			#endregion valida cuenta con valores en 0

			#region valida presupuesto

			foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.ultraGrid1.Rows.All)
			{
				decimal dValorarticulo = 0;
				foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dra in this.ultraGrid1.Rows.All)
				{
					if ((int)dr.Cells["idArticulo"].Value == (int)dra.Cells["idArticulo"].Value)
					{
						dValorarticulo = dValorarticulo + (decimal)dra.Cells["Subtotal"].Value ;
					}
				}
				decimal Presupuesto = 0;
				decimal Utilizado = 0;
				decimal Disponible = 0;
				if ( dValorarticulo > 0 )
				{
					string sSQL = string.Format("Exec [PresupuestoValidaValores] '{0}', {1}, {2}, {3}", 
						Convert.ToDateTime(this.dtFecha.Value).ToString("yyyyMMdd"), (int)dr.Cells["idArticulo"].Value, (int) dr.Cells["idProyecto"].Value, 3  );
					SqlDataReader dra = Funcion.rEscalarSQL(cdsSeteoF, sSQL, true);
					dra.Read();
					if (dra.HasRows)
					{
						if (dra.GetValue(0) != System.DBNull.Value)Presupuesto = dra.GetDecimal (0);
						if (dra.GetValue(1) != System.DBNull.Value)Utilizado = dra.GetDecimal (1);
						if (dra.GetValue(2) != System.DBNull.Value)Disponible = dra.GetDecimal (2);
					}
					dra.Close();
				}
				if ( dValorarticulo > Disponible )
				{
					MessageBox.Show("El presupuesto para la cuenta seleccionado a sido superada.", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
					this.ultraGrid1.ActiveRow = dr;
					this.ultraGrid1.ActiveRow.Selected = true;
					return;
				}
			}
			#endregion valida presupuesto



			if (DialogResult.Yes == MessageBox.Show("Antes de Guardar. \n\n VERIQUE QUE LOS DATOS REGISTRADOS SON LOS CORRECTOS \n\n ¿Esta seguro de Guardar?", "Point Technology", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2)) 
			{
				using (SqlConnection oConexion = new SqlConnection("Password=Bl45o6$9;Persist Security Info=True;User ID= " + MenuLatinium.stUsuario + ";Initial Catalog=" + MenuLatinium.stDirFacturacion + ";Data Source=" + MenuLatinium.stDirServidor + "; max pool size=1024"))
				{
					oConexion.Open();

					SqlTransaction oTransaction = null;
					SqlCommand oCmdActualiza = oConexion.CreateCommand();
					oCmdActualiza.Connection = oConexion;
					oCmdActualiza.CommandType = CommandType.Text;
					oCmdActualiza.CommandTimeout = 120;
						
					try
					{							
						oTransaction = oConexion.BeginTransaction();
						oCmdActualiza.Transaction = oTransaction;
						
						#region Numeracion
						if (bNuevo )
						{
							string sSQLNumero = string.Format("Exec NumeracionLocales 80, {0}, 1", 0);//(int)this.cmbBodega.Value);
							oCmdActualiza.CommandText = sSQLNumero;
							this.txtNumero.Text = oCmdActualiza.ExecuteScalar().ToString(); 
						}
						#endregion Numeracion

						#region Maestro
						bool bFechaLlegada = false;
						bool hFechaLlegada = false;
						DateTime dtFechaLlegada = DateTime.Today;
						DateTime dthorallegada= DateTime.Today;

					
						
						int idImportacion = 0;
						int idFormaPAgo = 0;
						//if (this.cmbImportacion.Value != System.DBNull.Value) idImportacion = (int)this.cmbImportacion.Value;
						//if (this.cmbFormaPago.Value != System.DBNull.Value) idFormaPAgo = (int)this.cmbFormaPago.Value;
												
						string sSQLMaestro = string.Format("Exec [GrabaMaestroPedidoPresupuesto] {0}, {1}, {2}, {3}, {4}, {5}, {6}, '{7}', '{8}', '{9}', '{10}', {11}, {12}, '{13}', {14}, {15}, {16}, {17}, {18}, {19}, {20}, {21},{22},{23},{24},{25},{26}", 
							(int)this.txtIdCompra.Value, (int)this.txtContadoCredito.Value, (int)this.txtIdComprobante.Value, (int)this.txtIdCredTributario.Value, false,
							(int)this.cmbFormaPago.Value, (int)this.cmbBodega.Value, this.txtNumero.Text.ToString(), 
							Convert.ToDateTime(this.dtFecha.Value).ToString("yyyyMMdd"), dtFechaLlegada.ToString("yyyyMMdd HH:mm"), dthorallegada.ToString("yyyyMMdd HH:mm"), 
							idImportacion, (int)this.txtEstado.Value, this.txtNotas.Text.ToString(),
							(decimal)this.txtIva0.Value, (decimal)this.txtDescIva0Total.Value, (decimal)this.txtIva.Value, (decimal)this.txtDescIvaTotal.Value, (decimal)this.txtIvaTotal.Value, (decimal)this.txtTotal.Value, 
							bFechaLlegada, hFechaLlegada, idFormaPAgo, 0, 0,0,0);
						oCmdActualiza.CommandText = sSQLMaestro;
						this.txtIdCompra.Value = (int)oCmdActualiza.ExecuteScalar();						
						#endregion Maestro

						#region Detalle
						int idArticulo = 0; 
						foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.ultraGrid1.Rows.All)
						{		
							
							
							string sSQLDetalle = string.Format("Exec GuardaDetalleOrdenDeCompra {0}, {1}, {2}, {3}, {4}, {5}, {6}, {7}, {8}, {9}, {10}, {11}, {12}, {13}, {14}, {15}, {16}, {17}, {18}, {19}, {20}, {21}, {22}, {23}",
								(int)dr.Cells["idDetCompra"].Value, (int)this.txtIdCompra.Value, (int)dr.Cells["idArticulo"].Value, (int)dr.Cells["idUnidad"].Value, 														
								(int)dr.Cells["Cantidad"].Value, (int)dr.Cells["CantAereo"].Value, (decimal)dr.Cells["Precio"].Value, (decimal)dr.Cells["Impuesto"].Value, 
								(decimal)dr.Cells["DescuentoArt"].Value, (decimal)dr.Cells["DescuentoPorc"].Value, (decimal)dr.Cells["Subtotal"].Value,
								(int)this.cmbFormaPago.Value, (int)dr.Cells["idImportacion"].Value, 
								(int)dr.Cells["idProyecto"].Value, 0, (int)dr.Cells["Tipo"].Value, (bool)dr.Cells["EstadoSerial"].Value,
								(decimal)dr.Cells["Courier"].Value, (decimal)dr.Cells["Peso"].Value, (decimal)dr.Cells["Arancel"].Value, (int)dr.Cells["idTipoCalculo"].Value, 
								(decimal)dr.Cells["DescNC"].Value, (decimal)dr.Cells["Valores"].Value,0);
							oCmdActualiza.CommandText = sSQLDetalle;
							oCmdActualiza.ExecuteNonQuery();
						}

						
												 
						string sSQLActDetalle = string.Format("Exec ConsultaDetalleOrdenCompra {0}", (int)this.txtIdCompra.Value);
						oCmdActualiza.CommandText = sSQLActDetalle;
						SqlDataAdapter oDADetalle = new SqlDataAdapter(oCmdActualiza);
						DataTable oDTDetalle = new DataTable();
						oDADetalle.Fill(oDTDetalle);
						this.ultraGrid1.DataSource = oDTDetalle;											
						#endregion Detalle	

						if (!bNuevo)
						{
							string sSQLBloqueo = string.Format("Exec RegistraBloqueoTransacciones {0}, 0, {1}, 1", idBloqueaTransacciones, (int)this.txtIdCompra.Value);
							oCmdActualiza.CommandText = sSQLBloqueo;
							oCmdActualiza.ExecuteNonQuery();
						}

						oTransaction.Commit();

						if ((int)this.txtEstado.Value == 1) this.lblEstado.Text = "PENDIENTE";				
						if ((int)this.txtEstado.Value == 3) this.lblEstado.Text = "PROCESADO";
												
						#region Controles
						bNuevo = false;
						bEdicion = false;
						bEliminaAlValidar = false;


						this.dtFecha.Enabled = false;
						this.cmbBodega.Enabled = false;


						this.cmbFormaPago.Enabled = false;

														
						FuncionesProcedimientos.EstadoGrids(false, this.ultraGrid1);
																	
						this.txtNotas.Enabled = false;
						
						if (miAcceso.Nuevo) this.btnNuevo.Enabled = true;
						if (miAcceso.Buscar) this.btnConsultar.Enabled = true;
						if (miAcceso.Editar) this.btnEditar.Enabled = true;
						if (miAcceso.Imprimir) this.btnImprimir.Enabled = true;
						
						this.btnGuardar.Enabled = false;												
						if (miAcceso.Anular && !bAnulado) this.btnAnular.Enabled = true;																																																
						#endregion Controles						
					}
					catch (Exception ex)
					{
						try
						{
							oTransaction.Rollback();

							MessageBox.Show(string.Format("Commit Exception Type : {0} {1}", ex.GetType(), ex.Message), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Error);
						}
						catch (Exception ex2)
						{
							MessageBox.Show(string.Format("Rollback Exception Type: {0} {1}", ex2.GetType(), ex2.Message), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Error);							
						}
					}
					finally
					{
						oConexion.Close();
					}
				}
			}
		}

		private void cmbBodega_ValueChanged(object sender, System.EventArgs e)
		{
			if (this.cmbBodega.Value != System.DBNull.Value)
			{	
				ListaSoliictante(this.cmbFormaPago, MenuLatinium.IdUsuario,  (DateTime)this.dtFecha.Value,  cdsSeteoF);
			}
		}

		private void su(object sender, System.EventArgs e)
		{
		
		}

		private void btnDistribucion_Click(object sender, System.EventArgs e)
		{
			if (!Validacion.vbComboVacio(this.cmbFormaPago, "Seleccione El solicitante")) return;
			Distribucion(1);
		}

		private void mnuAgencias_Click(object sender, System.EventArgs e)
		{
			Distribucion(1);
		}

		private void mnuAdministrativos_Click(object sender, System.EventArgs e)
		{
			Distribucion(2);
		}

		private void mnuGeneral_Click(object sender, System.EventArgs e)
		{
			Distribucion(3);
		}

		public void  Distribucion (int sTipo)
		{
			DateTime dtFechaFac = (DateTime) this.dtFecha.Value;
		
			using (frmCompraDistribucion ComprasDistribucion = new frmCompraDistribucion())
			{
				if (DialogResult.OK == ComprasDistribucion.ShowDialog())
				{
					if (ComprasDistribucion.uGridDetalle.Rows.Count == 1)
					{	
						#region comentado
						//							foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in ComprasDistribucion.uGridDetalle.Rows.All)
						//							{
						//								int iCuenta =(int)dr.Cells["idCuenta"].Value ;	
						//								decimal iValor  = (decimal)dr.Cells["Valor"].Value;
						//								try
						//								{	
						//									int IBodega = 1;
						//									this.cmbFormaPago.Value = 6;
						//									//this.txtNumEstado.Value = 7;
						//									this.dtFecha.Value = DateTime.Today;
						//				
						//									string sSQLInf = string.Format("Exec GeneraCompraDistribucion {0},{1},'{2}',{3}", iValor,iCuenta,dtFechaFac.ToString("yyyyMMdd"),sTipo);
						//									SqlDataReader drInf = Funcion.rEscalarSQL(cdsSeteoF, sSQLInf, true);
						//									while (drInf.Read())
						//									{	
						//										bActualiza=false;
						//										this.ultraGrid1.Rows.Band.AddNew();
						//										this.ultraGrid1.ActiveRow.Cells["Codigo"].Value = drInf.GetString(5);
						//										this.ultraGrid1.ActiveRow.Cells["Articulo"].Value = drInf.GetString(6);
						//										this.ultraGrid1.ActiveRow.Cells["Cantidad"].Value = drInf.GetInt32(7);
						//										this.ultraGrid1.ActiveRow.Cells["Precio"].Value = drInf.GetDecimal(8);
						//										this.ultraGrid1.ActiveRow.Cells["Impuesto"].Value = drInf.GetDecimal(9);
						//										this.ultraGrid1.ActiveRow.Cells["idProyecto"].Value = drInf.GetInt32(13);
						//										this.ultraGrid1.ActiveRow.Cells["idArticulo"].Value = drInf.GetInt32(3);
						//										bActualiza=true;
						//									}
						//									drInf.Close();	
						//									cmdTotal();
						//
						//									bNuevo = true;
						//									bEdicion = true;
						//
						//									#region Controles
						//									this.cmbBodega.Value = IBodega;
						//									//	this.dtFechadeEmicion.Enabled = true;
						//
						//									this.btnNuevo.Enabled = false;
						//									this.btnConsultar.Enabled = false;
						//									this.btnGuardar.Enabled = true;
						//									this.btnCancelar.Enabled = true;
						//									this.btnDistribucion.Enabled = false;
						//									#endregion Controles
						//						
						//									bNuevo = true;
						//									bEdicion = true;
						//					
						//									this.ultraGrid1.Rows.Band.AddNew();
						//								}
						//								catch(Exception Exc)
						//								{
						//									if (this.ultraGrid1.ActiveRow.Cells["Codigo"].Value == System.DBNull.Value)
						//									{
						//										MessageBox.Show("Cuenta no asignada", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);	
						//										return;
						//									}
						//								
						//									MessageBox.Show(string.Format("Error al Consultar: {0}", Exc), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Error);
						//
						//								}
						//							}
						#endregion comentado
						
						foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in ComprasDistribucion.uGridDetalle.Rows.All)
						{
							int iCuenta =(int)dr.Cells["idCuenta"].Value ;	
							decimal iValor  = (decimal)dr.Cells["Valor"].Value;
							OpenFileDialog sArchivo = new OpenFileDialog();

							sArchivo.Filter = "Microsoft Excel (*.xls)|*.xls";
							string sCuentasNot ="";
							string sCuentaseXIS ="";
							if (sArchivo.ShowDialog() == System.Windows.Forms.DialogResult.OK && sArchivo.FileName.Length > 0)
							{
								this.Cursor = Cursors.WaitCursor;                                  
								foreach (DataRow row in FuncionesProcedimientos.ImportaDeExcel(sArchivo.FileName, "SELECT BODEGA, PORCENTAJE FROM [PRESUPUESTO$] ", "PRESUPUESTO").Tables[0].Rows)
								{
									try
									{
										string BODEGA = "";
										decimal PORCENTAJE = 0.00m;
										if (row["BODEGA"] != System.DBNull.Value) BODEGA = row["BODEGA"].ToString();
										if (row["PORCENTAJE"] != System.DBNull.Value) PORCENTAJE = Convert.ToDecimal( row["PORCENTAJE"].ToString());
										if ( PORCENTAJE > 0 )
										{
											int icontProeyecto = 0;
											icontProeyecto = Funcion.iEscalarSQL(cdsSeteoF, string.Format("Exec [ValidaBodegaPorcPresupuesto] {0}, '{1}', '{2}'", MenuLatinium.IdUsuario,Convert.ToDateTime(this.dtFecha.Value).ToString("yyyyMMdd") , BODEGA));	
											if( icontProeyecto > 0 )
											{
												string sSQLInf = string.Format("Exec GeneraCompraDistribucionPresupuesto {0},{1},'{2}',{3}", iValor,iCuenta,Convert.ToDateTime(this.dtFecha.Value).ToString("yyyyMMdd"),PORCENTAJE);
												SqlDataReader drInf = Funcion.rEscalarSQL(cdsSeteoF, sSQLInf, true);
												while (drInf.Read())
												{	
													bActualiza=false;
													this.ultraGrid1.Rows.Band.AddNew();
													this.ultraGrid1.ActiveRow.Cells["Codigo"].Value = drInf.GetString(5);
													this.ultraGrid1.ActiveRow.Cells["Articulo"].Value = drInf.GetString(6);
													this.ultraGrid1.ActiveRow.Cells["Cantidad"].Value = drInf.GetInt32(7);
													this.ultraGrid1.ActiveRow.Cells["Precio"].Value = drInf.GetDecimal(8);
													this.ultraGrid1.ActiveRow.Cells["Impuesto"].Value = drInf.GetDecimal(9);
													this.ultraGrid1.ActiveRow.Cells["idProyecto"].Value = icontProeyecto;
													this.ultraGrid1.ActiveRow.Cells["idArticulo"].Value = drInf.GetInt32(3);
													bActualiza=true;
												}
												drInf.Close();	
												cmdTotal();
				
												bNuevo = true;
												bEdicion = true; 
											}
										}

									}
									catch(Exception Exc)
									{
										if (this.ultraGrid1.ActiveRow.Cells["Codigo"].Value == System.DBNull.Value)
										{
											MessageBox.Show("Cuenta no asignada", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);	
											return;
										}
															
										MessageBox.Show(string.Format("Error al Consultar: {0}", Exc), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Error);
									}
								}
							}
							bEliminaAlValidar = true;
							Validacion.vEliminaFilasVacias(this.ultraGrid1, "idArticulo");
							bEliminaAlValidar = false;

						}
					}
				}
			}
	

		}

		private void btnSalir_Click(object sender, System.EventArgs e)
		{
		this.Close ();
		}

		private void frmPedidoPresupuesto_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			
			if (!bEdicion)
			{
				if (e.KeyCode == Keys.N) if (miAcceso.Nuevo) this.btnNuevo_Click(sender, e);
				if (e.KeyCode == Keys.B) if (miAcceso.Buscar) this.btnConsultar_Click(sender, e);
				if ((int)this.txtIdCompra.Value > 0)
				{
					if (e.KeyCode == Keys.E) if (miAcceso.Editar) this.btnEditar_Click(sender, e);
					//if (e.KeyCode == Keys.I) if (miAcceso.Imprimir) this.btnImprimir_Click(sender, e);
					//if (e.KeyCode == Keys.A) if (miAcceso.Anular && !bAnulado) this.btnAnular_Click(sender, e);
					//if (e.KeyCode == Keys.X) if (miAcceso.Anular && bAnulado) this.btnAnular_Click(sender, e);
					if (e.KeyCode == Keys.F5)
					{
						miAcceso = new Acceso(cdsSeteoF, "2608");
						this.ConsultaMaestro((int)this.txtIdCompra.Value);
					}
				}
				if (e.KeyCode == Keys.Escape) this.Close();					
			}
			else
			{
				if (e.KeyCode == Keys.Escape)
					if (DialogResult.Yes == MessageBox.Show("Esta seguro de Cancelar la Edición", "Point Technology", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2))
						this.btnCancelar_Click(sender, e);

				if (e.KeyCode == Keys.F12) this.btnGuardar_Click(sender, e);
			}
		}

		private void btnAnular_Click(object sender, System.EventArgs e)
		{
		
		}


	}
}
