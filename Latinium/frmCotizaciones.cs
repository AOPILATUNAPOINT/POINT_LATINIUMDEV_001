using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using System.Globalization;
using System.Data;
using System.Data.SqlClient;
using CrystalDecisions;
using CrystalDecisions.CrystalReports.Engine;
using CrystalDecisions.ReportSource;
using CrystalDecisions.Shared;

namespace Latinium
{
	/// <summary>
	/// Descripción breve de frmCotizaciones.
	/// </summary>
	public class frmCotizaciones : System.Windows.Forms.Form
	{
		private System.Windows.Forms.Button btnCliente;
		private System.Windows.Forms.Label label4;
		private System.Windows.Forms.Label label3;
		private System.Windows.Forms.Label lblSubTotal;
		private System.Windows.Forms.Label lblIva;
		private System.Windows.Forms.Label lblDescuento;
		private System.Windows.Forms.Label lblIvaSubt;
		private System.Windows.Forms.Label lblIva0;
		private System.Windows.Forms.Label label2;
		private System.Windows.Forms.Label lblIdentificacion;
		private System.Windows.Forms.Label label8;
		private System.Windows.Forms.Label lblComprobante;
		private System.Windows.Forms.Label lblVendedor;
		private System.Windows.Forms.Label lblCliente;
		private System.Windows.Forms.Label label6;
		private System.Windows.Forms.Label label1;
		private System.Windows.Forms.Label lblBodega;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtNotas;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtTotal;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtIvaTotal;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtDescIva0Total;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtDescIvaTotal;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtIva;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtIva0;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtContadoCredito;
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
		private Infragistics.Win.UltraWinGrid.UltraGrid ultraGrid1;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbFormaPago;
		private System.Windows.Forms.Label lblUsuario;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtComprobante;
		private System.Windows.Forms.Label lblEstado;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbVendedor;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbTipoRuc;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtRuc;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtNombre;
		private Infragistics.Win.UltraWinSchedule.UltraCalendarCombo dtFecha;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbBodega;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtNumero;
		private System.Windows.Forms.GroupBox groupBox1;
		private System.Windows.Forms.Label label67;
		private System.Windows.Forms.Label label66;
		private System.Windows.Forms.Label label19;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbDirecciones;
		private Infragistics.Win.UltraWinEditors.UltraCheckEditor chkEntregado;
		private Infragistics.Win.UltraWinSchedule.UltraCalendarCombo dtFechaEntrega;
		private Infragistics.Win.UltraWinSchedule.UltraCalendarCombo dtpHoraEntrega;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbProyecto;
		private System.Windows.Forms.Label label10;
		private C1.Data.C1DataSet cdsSeteoF;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource1;
		private System.Windows.Forms.Button btnEMail;
		private System.Windows.Forms.Button btnProyecto;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource2;
		private System.Windows.Forms.Label label5;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtBaseImponible;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtJornadas;
		private System.Windows.Forms.Label label7;
		private System.Windows.Forms.Label label9;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtDescuento;
		private System.Windows.Forms.Button btnAplicarDescuento;
		private System.Windows.Forms.Button btnDuplicar;
		private System.Windows.Forms.Button btnJornadas;
		/// <summary>
		/// Variable del diseñador requerida.
		/// </summary>
		private System.ComponentModel.Container components = null;

		public frmCotizaciones()
		{
			//
			// Necesario para admitir el Diseñador de Windows Forms
			//
			InitializeComponent();

			//
			// TODO: agregar código de constructor después de llamar a InitializeComponent
			//
		}

		int IdCompra = 0;

		public frmCotizaciones(int idCompra)
		{
			//
			// Necesario para admitir el Diseñador de Windows Forms
			//
			InitializeComponent();
			IdCompra = idCompra;
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
			Infragistics.Win.Appearance appearance2 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance3 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance4 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance5 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance6 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance7 = new Infragistics.Win.Appearance();
			System.Resources.ResourceManager resources = new System.Resources.ResourceManager(typeof(frmCotizaciones));
			Infragistics.Win.Appearance appearance8 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand1 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn1 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idDetCompra");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn2 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idOrigen");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn3 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCompra");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn4 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idArticulo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn5 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idUnidad");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn6 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Codigo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn7 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Articulo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn8 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Jornadas");
			Infragistics.Win.Appearance appearance9 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn9 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Cantidad");
			Infragistics.Win.Appearance appearance10 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn10 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Precio");
			Infragistics.Win.Appearance appearance11 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn11 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Impuesto");
			Infragistics.Win.Appearance appearance12 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn12 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("DescuentoArt");
			Infragistics.Win.Appearance appearance13 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn13 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("DescuentoPorc");
			Infragistics.Win.Appearance appearance14 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn14 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("SubTotal");
			Infragistics.Win.Appearance appearance15 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn15 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("EstadoSerial");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn16 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Tipo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn17 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Entregado");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn18 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Posicion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn19 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Principal");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn20 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Bloqueado");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn21 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCombo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn22 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idArticuloGE");
			Infragistics.Win.Appearance appearance16 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance17 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance18 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance19 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance20 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance21 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn1 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idDetCompra");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn2 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idOrigen");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn3 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idCompra");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn4 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idArticulo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn5 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idUnidad");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn6 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Codigo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn7 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Articulo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn8 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Jornadas");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn9 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Cantidad");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn10 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Precio");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn11 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Impuesto");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn12 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("DescuentoArt");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn13 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("DescuentoPorc");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn14 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("SubTotal");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn15 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("EstadoSerial");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn16 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Tipo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn17 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Entregado");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn18 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Posicion");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn19 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Principal");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn20 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Bloqueado");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn21 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idCombo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn22 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idArticuloGE");
			Infragistics.Win.Appearance appearance22 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand2 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn23 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idFormaPago");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn24 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("FormaPago");
			Infragistics.Win.Appearance appearance23 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance24 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand3 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn25 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idPersonal");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn26 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn27 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Codigo");
			Infragistics.Win.Appearance appearance25 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand4 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn28 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idTipoRuc");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn29 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("TipoRuc");
			Infragistics.Win.Appearance appearance26 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance27 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance28 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton dateButton1 = new Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand5 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn30 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Bodega");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn31 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn32 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Codigo", 0);
			Infragistics.Win.Appearance appearance29 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance30 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand6 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn33 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idDreccion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn34 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Direccion");
			Infragistics.Win.Appearance appearance31 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance32 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton dateButton2 = new Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton();
			Infragistics.Win.Appearance appearance33 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton dateButton3 = new Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand7 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn35 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idProyectoZoom");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn36 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Numero");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn37 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn23 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idProyectoZoom");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn24 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Numero");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn25 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Nombre");
			Infragistics.Win.Appearance appearance34 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance35 = new Infragistics.Win.Appearance();
			this.btnCliente = new System.Windows.Forms.Button();
			this.label4 = new System.Windows.Forms.Label();
			this.label3 = new System.Windows.Forms.Label();
			this.lblSubTotal = new System.Windows.Forms.Label();
			this.lblIva = new System.Windows.Forms.Label();
			this.lblDescuento = new System.Windows.Forms.Label();
			this.lblIvaSubt = new System.Windows.Forms.Label();
			this.lblIva0 = new System.Windows.Forms.Label();
			this.label2 = new System.Windows.Forms.Label();
			this.lblIdentificacion = new System.Windows.Forms.Label();
			this.label8 = new System.Windows.Forms.Label();
			this.lblComprobante = new System.Windows.Forms.Label();
			this.lblVendedor = new System.Windows.Forms.Label();
			this.lblCliente = new System.Windows.Forms.Label();
			this.label6 = new System.Windows.Forms.Label();
			this.label1 = new System.Windows.Forms.Label();
			this.lblBodega = new System.Windows.Forms.Label();
			this.txtNotas = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.txtTotal = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtIvaTotal = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtDescIva0Total = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtDescIvaTotal = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtIva = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtIva0 = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtContadoCredito = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
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
			this.ultraGrid1 = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.ultraDataSource1 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.cmbFormaPago = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.lblUsuario = new System.Windows.Forms.Label();
			this.txtComprobante = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.lblEstado = new System.Windows.Forms.Label();
			this.cmbVendedor = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.cmbTipoRuc = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.txtRuc = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.txtNombre = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.dtFecha = new Infragistics.Win.UltraWinSchedule.UltraCalendarCombo();
			this.cmbBodega = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.txtNumero = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.groupBox1 = new System.Windows.Forms.GroupBox();
			this.label67 = new System.Windows.Forms.Label();
			this.label66 = new System.Windows.Forms.Label();
			this.label19 = new System.Windows.Forms.Label();
			this.cmbDirecciones = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.chkEntregado = new Infragistics.Win.UltraWinEditors.UltraCheckEditor();
			this.dtFechaEntrega = new Infragistics.Win.UltraWinSchedule.UltraCalendarCombo();
			this.dtpHoraEntrega = new Infragistics.Win.UltraWinSchedule.UltraCalendarCombo();
			this.cmbProyecto = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.ultraDataSource2 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.label10 = new System.Windows.Forms.Label();
			this.cdsSeteoF = new C1.Data.C1DataSet();
			this.btnEMail = new System.Windows.Forms.Button();
			this.btnProyecto = new System.Windows.Forms.Button();
			this.label5 = new System.Windows.Forms.Label();
			this.txtBaseImponible = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtJornadas = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.label7 = new System.Windows.Forms.Label();
			this.label9 = new System.Windows.Forms.Label();
			this.txtDescuento = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.btnAplicarDescuento = new System.Windows.Forms.Button();
			this.btnDuplicar = new System.Windows.Forms.Button();
			this.btnJornadas = new System.Windows.Forms.Button();
			((System.ComponentModel.ISupportInitialize)(this.txtNotas)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtTotal)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIvaTotal)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtDescIva0Total)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtDescIvaTotal)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIva)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIva0)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtContadoCredito)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtBodegaPredef)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtEstado)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIdCliente)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIdCompra)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraGrid1)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource1)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbFormaPago)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtComprobante)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbVendedor)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbTipoRuc)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtRuc)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNombre)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.dtFecha)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbBodega)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNumero)).BeginInit();
			this.groupBox1.SuspendLayout();
			((System.ComponentModel.ISupportInitialize)(this.cmbDirecciones)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.dtFechaEntrega)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.dtpHoraEntrega)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbProyecto)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource2)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtBaseImponible)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtJornadas)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtDescuento)).BeginInit();
			this.SuspendLayout();
			// 
			// btnCliente
			// 
			this.btnCliente.CausesValidation = false;
			this.btnCliente.Enabled = false;
			this.btnCliente.Location = new System.Drawing.Point(816, 40);
			this.btnCliente.Name = "btnCliente";
			this.btnCliente.Size = new System.Drawing.Size(24, 21);
			this.btnCliente.TabIndex = 457;
			this.btnCliente.Text = "...";
			this.btnCliente.Click += new System.EventHandler(this.btnCliente_Click);
			// 
			// label4
			// 
			this.label4.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.label4.AutoSize = true;
			this.label4.BackColor = System.Drawing.Color.Transparent;
			this.label4.Location = new System.Drawing.Point(8, 288);
			this.label4.Name = "label4";
			this.label4.Size = new System.Drawing.Size(76, 17);
			this.label4.TabIndex = 454;
			this.label4.Text = "Observaciones";
			this.label4.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// label3
			// 
			this.label3.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			this.label3.AutoSize = true;
			this.label3.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label3.Location = new System.Drawing.Point(768, 382);
			this.label3.Name = "label3";
			this.label3.Size = new System.Drawing.Size(38, 17);
			this.label3.TabIndex = 452;
			this.label3.Text = "TOTAL";
			this.label3.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// lblSubTotal
			// 
			this.lblSubTotal.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			this.lblSubTotal.AutoSize = true;
			this.lblSubTotal.Location = new System.Drawing.Point(768, 282);
			this.lblSubTotal.Name = "lblSubTotal";
			this.lblSubTotal.Size = new System.Drawing.Size(58, 17);
			this.lblSubTotal.TabIndex = 445;
			this.lblSubTotal.Text = "SUBTOTAL";
			this.lblSubTotal.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// lblIva
			// 
			this.lblIva.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			this.lblIva.AutoSize = true;
			this.lblIva.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.lblIva.Location = new System.Drawing.Point(768, 354);
			this.lblIva.Name = "lblIva";
			this.lblIva.Size = new System.Drawing.Size(22, 17);
			this.lblIva.TabIndex = 444;
			this.lblIva.Text = "IVA";
			this.lblIva.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// lblDescuento
			// 
			this.lblDescuento.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			this.lblDescuento.AutoSize = true;
			this.lblDescuento.Location = new System.Drawing.Point(768, 306);
			this.lblDescuento.Name = "lblDescuento";
			this.lblDescuento.Size = new System.Drawing.Size(67, 17);
			this.lblDescuento.TabIndex = 443;
			this.lblDescuento.Text = "DESCUENTO";
			this.lblDescuento.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// lblIvaSubt
			// 
			this.lblIvaSubt.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			this.lblIvaSubt.AutoSize = true;
			this.lblIvaSubt.Location = new System.Drawing.Point(992, 264);
			this.lblIvaSubt.Name = "lblIvaSubt";
			this.lblIvaSubt.Size = new System.Drawing.Size(22, 17);
			this.lblIvaSubt.TabIndex = 442;
			this.lblIvaSubt.Text = "IVA";
			this.lblIvaSubt.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// lblIva0
			// 
			this.lblIva0.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			this.lblIva0.AutoSize = true;
			this.lblIva0.Location = new System.Drawing.Point(896, 264);
			this.lblIva0.Name = "lblIva0";
			this.lblIva0.Size = new System.Drawing.Size(32, 17);
			this.lblIva0.TabIndex = 441;
			this.lblIva0.Text = "IVA 0";
			this.lblIva0.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// label2
			// 
			this.label2.AutoSize = true;
			this.label2.Location = new System.Drawing.Point(248, 42);
			this.label2.Name = "label2";
			this.label2.Size = new System.Drawing.Size(86, 17);
			this.label2.TabIndex = 425;
			this.label2.Text = "N. Identificación";
			this.label2.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// lblIdentificacion
			// 
			this.lblIdentificacion.AutoSize = true;
			this.lblIdentificacion.Location = new System.Drawing.Point(8, 42);
			this.lblIdentificacion.Name = "lblIdentificacion";
			this.lblIdentificacion.Size = new System.Drawing.Size(71, 17);
			this.lblIdentificacion.TabIndex = 424;
			this.lblIdentificacion.Text = "Identificación";
			this.lblIdentificacion.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// label8
			// 
			this.label8.AutoSize = true;
			this.label8.Location = new System.Drawing.Point(8, 74);
			this.label8.Name = "label8";
			this.label8.Size = new System.Drawing.Size(80, 17);
			this.label8.TabIndex = 423;
			this.label8.Text = "Forma de Pago";
			this.label8.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// lblComprobante
			// 
			this.lblComprobante.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			this.lblComprobante.AutoSize = true;
			this.lblComprobante.BackColor = System.Drawing.Color.Transparent;
			this.lblComprobante.Location = new System.Drawing.Point(848, 42);
			this.lblComprobante.Name = "lblComprobante";
			this.lblComprobante.Size = new System.Drawing.Size(41, 17);
			this.lblComprobante.TabIndex = 419;
			this.lblComprobante.Text = "Factura";
			this.lblComprobante.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// lblVendedor
			// 
			this.lblVendedor.AutoSize = true;
			this.lblVendedor.Location = new System.Drawing.Point(488, 10);
			this.lblVendedor.Name = "lblVendedor";
			this.lblVendedor.Size = new System.Drawing.Size(52, 17);
			this.lblVendedor.TabIndex = 416;
			this.lblVendedor.Text = "Vendedor";
			this.lblVendedor.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// lblCliente
			// 
			this.lblCliente.AutoSize = true;
			this.lblCliente.Location = new System.Drawing.Point(488, 42);
			this.lblCliente.Name = "lblCliente";
			this.lblCliente.Size = new System.Drawing.Size(44, 17);
			this.lblCliente.TabIndex = 414;
			this.lblCliente.Text = "Nombre";
			this.lblCliente.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// label6
			// 
			this.label6.AutoSize = true;
			this.label6.Location = new System.Drawing.Point(8, 10);
			this.label6.Name = "label6";
			this.label6.Size = new System.Drawing.Size(34, 17);
			this.label6.TabIndex = 413;
			this.label6.Text = "Fecha";
			this.label6.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// label1
			// 
			this.label1.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			this.label1.AutoSize = true;
			this.label1.Location = new System.Drawing.Point(848, 10);
			this.label1.Name = "label1";
			this.label1.Size = new System.Drawing.Size(44, 17);
			this.label1.TabIndex = 409;
			this.label1.Text = "Número";
			this.label1.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// lblBodega
			// 
			this.lblBodega.AutoSize = true;
			this.lblBodega.Location = new System.Drawing.Point(248, 10);
			this.lblBodega.Name = "lblBodega";
			this.lblBodega.Size = new System.Drawing.Size(30, 17);
			this.lblBodega.TabIndex = 407;
			this.lblBodega.Text = "Local";
			this.lblBodega.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// txtNotas
			// 
			this.txtNotas.AcceptsReturn = true;
			this.txtNotas.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			appearance1.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtNotas.Appearance = appearance1;
			this.txtNotas.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtNotas.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtNotas.Enabled = false;
			this.txtNotas.Location = new System.Drawing.Point(8, 312);
			this.txtNotas.Multiline = true;
			this.txtNotas.Name = "txtNotas";
			this.txtNotas.Size = new System.Drawing.Size(296, 96);
			this.txtNotas.TabIndex = 453;
			this.txtNotas.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtNotas_KeyDown);
			// 
			// txtTotal
			// 
			this.txtTotal.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			appearance2.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtTotal.Appearance = appearance2;
			this.txtTotal.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtTotal.Enabled = false;
			this.txtTotal.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.txtTotal.FormatString = "#,##0.00";
			this.txtTotal.Location = new System.Drawing.Point(848, 376);
			this.txtTotal.Name = "txtTotal";
			this.txtTotal.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtTotal.PromptChar = ' ';
			this.txtTotal.Size = new System.Drawing.Size(168, 28);
			this.txtTotal.TabIndex = 451;
			// 
			// txtIvaTotal
			// 
			this.txtIvaTotal.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			appearance3.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtIvaTotal.Appearance = appearance3;
			this.txtIvaTotal.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtIvaTotal.Enabled = false;
			this.txtIvaTotal.FormatString = "#,##0.00";
			this.txtIvaTotal.Location = new System.Drawing.Point(936, 352);
			this.txtIvaTotal.Name = "txtIvaTotal";
			this.txtIvaTotal.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtIvaTotal.PromptChar = ' ';
			this.txtIvaTotal.Size = new System.Drawing.Size(80, 22);
			this.txtIvaTotal.TabIndex = 450;
			// 
			// txtDescIva0Total
			// 
			this.txtDescIva0Total.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			appearance4.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtDescIva0Total.Appearance = appearance4;
			this.txtDescIva0Total.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtDescIva0Total.Enabled = false;
			this.txtDescIva0Total.FormatString = "#,##0.00";
			this.txtDescIva0Total.Location = new System.Drawing.Point(848, 304);
			this.txtDescIva0Total.Name = "txtDescIva0Total";
			this.txtDescIva0Total.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtDescIva0Total.PromptChar = ' ';
			this.txtDescIva0Total.Size = new System.Drawing.Size(80, 22);
			this.txtDescIva0Total.TabIndex = 449;
			// 
			// txtDescIvaTotal
			// 
			this.txtDescIvaTotal.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			appearance5.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtDescIvaTotal.Appearance = appearance5;
			this.txtDescIvaTotal.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtDescIvaTotal.Enabled = false;
			this.txtDescIvaTotal.FormatString = "#,##0.00";
			this.txtDescIvaTotal.Location = new System.Drawing.Point(936, 304);
			this.txtDescIvaTotal.Name = "txtDescIvaTotal";
			this.txtDescIvaTotal.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtDescIvaTotal.PromptChar = ' ';
			this.txtDescIvaTotal.Size = new System.Drawing.Size(80, 22);
			this.txtDescIvaTotal.TabIndex = 448;
			this.txtDescIvaTotal.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtDescIvaTotal_KeyDown);
			this.txtDescIvaTotal.Click += new System.EventHandler(this.txtDescIvaTotal_Click);
			this.txtDescIvaTotal.Validated += new System.EventHandler(this.txtDescIvaTotal_Validated);
			this.txtDescIvaTotal.Enter += new System.EventHandler(this.txtDescIvaTotal_Enter);
			// 
			// txtIva
			// 
			this.txtIva.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			appearance6.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtIva.Appearance = appearance6;
			this.txtIva.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtIva.Enabled = false;
			this.txtIva.FormatString = "#,##0.00";
			this.txtIva.Location = new System.Drawing.Point(936, 280);
			this.txtIva.Name = "txtIva";
			this.txtIva.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtIva.PromptChar = ' ';
			this.txtIva.Size = new System.Drawing.Size(80, 22);
			this.txtIva.TabIndex = 447;
			// 
			// txtIva0
			// 
			this.txtIva0.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			appearance7.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtIva0.Appearance = appearance7;
			this.txtIva0.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtIva0.Enabled = false;
			this.txtIva0.FormatString = "#,##0.00";
			this.txtIva0.Location = new System.Drawing.Point(848, 280);
			this.txtIva0.Name = "txtIva0";
			this.txtIva0.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtIva0.PromptChar = ' ';
			this.txtIva0.Size = new System.Drawing.Size(80, 22);
			this.txtIva0.TabIndex = 446;
			// 
			// txtContadoCredito
			// 
			this.txtContadoCredito.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			this.txtContadoCredito.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtContadoCredito.Enabled = false;
			this.txtContadoCredito.Location = new System.Drawing.Point(992, 432);
			this.txtContadoCredito.Name = "txtContadoCredito";
			this.txtContadoCredito.PromptChar = ' ';
			this.txtContadoCredito.Size = new System.Drawing.Size(16, 22);
			this.txtContadoCredito.TabIndex = 440;
			this.txtContadoCredito.Visible = false;
			// 
			// txtBodegaPredef
			// 
			this.txtBodegaPredef.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			this.txtBodegaPredef.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtBodegaPredef.Enabled = false;
			this.txtBodegaPredef.Location = new System.Drawing.Point(976, 432);
			this.txtBodegaPredef.Name = "txtBodegaPredef";
			this.txtBodegaPredef.PromptChar = ' ';
			this.txtBodegaPredef.Size = new System.Drawing.Size(16, 22);
			this.txtBodegaPredef.TabIndex = 439;
			this.txtBodegaPredef.Visible = false;
			// 
			// txtEstado
			// 
			this.txtEstado.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			this.txtEstado.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtEstado.Enabled = false;
			this.txtEstado.Location = new System.Drawing.Point(944, 432);
			this.txtEstado.Name = "txtEstado";
			this.txtEstado.PromptChar = ' ';
			this.txtEstado.Size = new System.Drawing.Size(16, 22);
			this.txtEstado.TabIndex = 438;
			this.txtEstado.Visible = false;
			// 
			// txtIdCliente
			// 
			this.txtIdCliente.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			this.txtIdCliente.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtIdCliente.Enabled = false;
			this.txtIdCliente.Location = new System.Drawing.Point(960, 432);
			this.txtIdCliente.Name = "txtIdCliente";
			this.txtIdCliente.PromptChar = ' ';
			this.txtIdCliente.Size = new System.Drawing.Size(16, 22);
			this.txtIdCliente.TabIndex = 437;
			this.txtIdCliente.Visible = false;
			// 
			// gbBotones
			// 
			this.gbBotones.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.gbBotones.Location = new System.Drawing.Point(0, 416);
			this.gbBotones.Name = "gbBotones";
			this.gbBotones.Size = new System.Drawing.Size(1024, 8);
			this.gbBotones.TabIndex = 436;
			this.gbBotones.TabStop = false;
			// 
			// btnNuevo
			// 
			this.btnNuevo.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btnNuevo.CausesValidation = false;
			this.btnNuevo.Image = ((System.Drawing.Image)(resources.GetObject("btnNuevo.Image")));
			this.btnNuevo.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnNuevo.Location = new System.Drawing.Point(8, 432);
			this.btnNuevo.Name = "btnNuevo";
			this.btnNuevo.Size = new System.Drawing.Size(72, 23);
			this.btnNuevo.TabIndex = 428;
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
			this.btnConsultar.Location = new System.Drawing.Point(81, 432);
			this.btnConsultar.Name = "btnConsultar";
			this.btnConsultar.Size = new System.Drawing.Size(78, 23);
			this.btnConsultar.TabIndex = 434;
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
			this.btnEditar.Location = new System.Drawing.Point(160, 432);
			this.btnEditar.Name = "btnEditar";
			this.btnEditar.Size = new System.Drawing.Size(76, 23);
			this.btnEditar.TabIndex = 433;
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
			this.btnImprimir.Location = new System.Drawing.Point(313, 432);
			this.btnImprimir.Name = "btnImprimir";
			this.btnImprimir.Size = new System.Drawing.Size(77, 23);
			this.btnImprimir.TabIndex = 430;
			this.btnImprimir.Text = "&Imprimir";
			this.btnImprimir.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnImprimir.Click += new System.EventHandler(this.btnImprimir_Click);
			// 
			// btnGuardar
			// 
			this.btnGuardar.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btnGuardar.Enabled = false;
			this.btnGuardar.Image = ((System.Drawing.Image)(resources.GetObject("btnGuardar.Image")));
			this.btnGuardar.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnGuardar.Location = new System.Drawing.Point(469, 432);
			this.btnGuardar.Name = "btnGuardar";
			this.btnGuardar.Size = new System.Drawing.Size(76, 23);
			this.btnGuardar.TabIndex = 429;
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
			this.btnAnular.Location = new System.Drawing.Point(546, 432);
			this.btnAnular.Name = "btnAnular";
			this.btnAnular.Size = new System.Drawing.Size(76, 23);
			this.btnAnular.TabIndex = 435;
			this.btnAnular.Text = "&Anular";
			this.btnAnular.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnAnular.Click += new System.EventHandler(this.btnAnular_Click);
			// 
			// btnCancelar
			// 
			this.btnCancelar.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btnCancelar.CausesValidation = false;
			this.btnCancelar.Enabled = false;
			this.btnCancelar.Image = ((System.Drawing.Image)(resources.GetObject("btnCancelar.Image")));
			this.btnCancelar.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnCancelar.Location = new System.Drawing.Point(623, 432);
			this.btnCancelar.Name = "btnCancelar";
			this.btnCancelar.Size = new System.Drawing.Size(72, 23);
			this.btnCancelar.TabIndex = 431;
			this.btnCancelar.Text = "&Cancelar";
			this.btnCancelar.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnCancelar.Click += new System.EventHandler(this.btnCancelar_Click);
			// 
			// btnSalir
			// 
			this.btnSalir.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btnSalir.CausesValidation = false;
			this.btnSalir.Image = ((System.Drawing.Image)(resources.GetObject("btnSalir.Image")));
			this.btnSalir.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnSalir.Location = new System.Drawing.Point(696, 432);
			this.btnSalir.Name = "btnSalir";
			this.btnSalir.TabIndex = 432;
			this.btnSalir.Text = "&Salir";
			this.btnSalir.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnSalir.Click += new System.EventHandler(this.btnSalir_Click);
			// 
			// txtIdCompra
			// 
			this.txtIdCompra.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			this.txtIdCompra.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtIdCompra.Enabled = false;
			this.txtIdCompra.Location = new System.Drawing.Point(928, 432);
			this.txtIdCompra.Name = "txtIdCompra";
			this.txtIdCompra.PromptChar = ' ';
			this.txtIdCompra.Size = new System.Drawing.Size(16, 22);
			this.txtIdCompra.TabIndex = 427;
			this.txtIdCompra.Visible = false;
			// 
			// ultraGrid1
			// 
			this.ultraGrid1.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.ultraGrid1.Cursor = System.Windows.Forms.Cursors.Default;
			this.ultraGrid1.DataSource = this.ultraDataSource1;
			appearance8.BackColor = System.Drawing.Color.White;
			this.ultraGrid1.DisplayLayout.Appearance = appearance8;
			this.ultraGrid1.DisplayLayout.AutoFitColumns = true;
			ultraGridBand1.AddButtonCaption = "_DetCompra";
			ultraGridColumn1.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn1.Header.VisiblePosition = 14;
			ultraGridColumn1.Hidden = true;
			ultraGridColumn2.Header.VisiblePosition = 1;
			ultraGridColumn2.Hidden = true;
			ultraGridColumn2.Width = 42;
			ultraGridColumn3.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn3.Header.VisiblePosition = 0;
			ultraGridColumn3.Hidden = true;
			ultraGridColumn4.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn4.Header.VisiblePosition = 2;
			ultraGridColumn4.Hidden = true;
			ultraGridColumn4.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.DropDownList;
			ultraGridColumn4.Width = 194;
			ultraGridColumn5.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn5.Header.VisiblePosition = 8;
			ultraGridColumn5.Hidden = true;
			ultraGridColumn5.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.DropDownValidate;
			ultraGridColumn5.Width = 53;
			ultraGridColumn6.Case = Infragistics.Win.UltraWinGrid.Case.Upper;
			ultraGridColumn6.Header.VisiblePosition = 3;
			ultraGridColumn6.Width = 200;
			ultraGridColumn7.Case = Infragistics.Win.UltraWinGrid.Case.Upper;
			ultraGridColumn7.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn7.Header.VisiblePosition = 5;
			ultraGridColumn7.Width = 321;
			appearance9.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn8.CellAppearance = appearance9;
			ultraGridColumn8.Format = "#,##0.00";
			ultraGridColumn8.Header.VisiblePosition = 6;
			ultraGridColumn8.Width = 71;
			appearance10.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn9.CellAppearance = appearance10;
			ultraGridColumn9.Header.VisiblePosition = 7;
			ultraGridColumn9.MaskInput = "";
			ultraGridColumn9.NullText = "0";
			ultraGridColumn9.PromptChar = ' ';
			ultraGridColumn9.Width = 76;
			appearance11.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn10.CellAppearance = appearance11;
			ultraGridColumn10.Format = "#,##0.00";
			ultraGridColumn10.Header.VisiblePosition = 9;
			ultraGridColumn10.MaskInput = "{LOC}nn,nnn,nnn.nn";
			ultraGridColumn10.PromptChar = ' ';
			ultraGridColumn10.Width = 87;
			ultraGridColumn11.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance12.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn11.CellAppearance = appearance12;
			ultraGridColumn11.Format = "#,##0.00";
			ultraGridColumn11.Header.Caption = "I.V.A.";
			ultraGridColumn11.Header.VisiblePosition = 10;
			ultraGridColumn11.MaskInput = "{LOC}nnnnnnn.nn";
			ultraGridColumn11.PromptChar = ' ';
			ultraGridColumn11.Width = 49;
			appearance13.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn12.CellAppearance = appearance13;
			ultraGridColumn12.Format = "#,##0.00";
			ultraGridColumn12.Header.Caption = "Descuento";
			ultraGridColumn12.Header.VisiblePosition = 11;
			ultraGridColumn12.Hidden = true;
			ultraGridColumn12.MaskInput = "{LOC}nn,nnn,nnn.nn";
			ultraGridColumn12.PromptChar = ' ';
			ultraGridColumn12.Width = 14;
			appearance14.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn13.CellAppearance = appearance14;
			ultraGridColumn13.Format = "#,##0.00";
			ultraGridColumn13.Header.Caption = "% Desc.";
			ultraGridColumn13.Header.VisiblePosition = 12;
			ultraGridColumn13.MaskInput = "";
			ultraGridColumn13.PromptChar = ' ';
			ultraGridColumn13.Width = 90;
			ultraGridColumn14.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance15.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn14.CellAppearance = appearance15;
			ultraGridColumn14.Format = "#,##0.00";
			ultraGridColumn14.Header.VisiblePosition = 13;
			ultraGridColumn14.PromptChar = ' ';
			ultraGridColumn14.Width = 93;
			ultraGridColumn15.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn15.Header.VisiblePosition = 4;
			ultraGridColumn15.Hidden = true;
			ultraGridColumn16.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn16.Header.VisiblePosition = 15;
			ultraGridColumn16.Hidden = true;
			ultraGridColumn17.Header.VisiblePosition = 16;
			ultraGridColumn17.Hidden = true;
			ultraGridColumn17.Width = 48;
			ultraGridColumn18.Header.VisiblePosition = 17;
			ultraGridColumn18.Hidden = true;
			ultraGridColumn18.Width = 42;
			ultraGridColumn19.Header.VisiblePosition = 19;
			ultraGridColumn19.Hidden = true;
			ultraGridColumn19.Width = 42;
			ultraGridColumn20.Header.VisiblePosition = 20;
			ultraGridColumn20.Hidden = true;
			ultraGridColumn20.Width = 49;
			ultraGridColumn21.Header.VisiblePosition = 18;
			ultraGridColumn21.Hidden = true;
			ultraGridColumn22.Header.VisiblePosition = 21;
			ultraGridColumn22.Hidden = true;
			ultraGridColumn22.Width = 63;
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
																										 ultraGridColumn22});
			this.ultraGrid1.DisplayLayout.BandsSerializer.Add(ultraGridBand1);
			appearance16.ForeColor = System.Drawing.Color.Black;
			appearance16.ForeColorDisabled = System.Drawing.Color.Black;
			this.ultraGrid1.DisplayLayout.Override.ActiveRowAppearance = appearance16;
			this.ultraGrid1.DisplayLayout.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.TemplateOnBottom;
			this.ultraGrid1.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.True;
			this.ultraGrid1.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.True;
			appearance17.BackColor = System.Drawing.Color.Transparent;
			this.ultraGrid1.DisplayLayout.Override.CardAreaAppearance = appearance17;
			appearance18.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance18.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance18.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance18.FontData.BoldAsString = "True";
			appearance18.FontData.Name = "Arial";
			appearance18.FontData.SizeInPoints = 8.5F;
			appearance18.ForeColor = System.Drawing.Color.White;
			appearance18.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.ultraGrid1.DisplayLayout.Override.HeaderAppearance = appearance18;
			appearance19.BackColor = System.Drawing.SystemColors.ControlLight;
			appearance19.BackColor2 = System.Drawing.Color.Gainsboro;
			appearance19.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ultraGrid1.DisplayLayout.Override.RowAlternateAppearance = appearance19;
			appearance20.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance20.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance20.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ultraGrid1.DisplayLayout.Override.RowSelectorAppearance = appearance20;
			appearance21.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance21.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance21.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ultraGrid1.DisplayLayout.Override.SelectedRowAppearance = appearance21;
			this.ultraGrid1.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.ultraGrid1.Location = new System.Drawing.Point(8, 104);
			this.ultraGrid1.Name = "ultraGrid1";
			this.ultraGrid1.Size = new System.Drawing.Size(1008, 152);
			this.ultraGrid1.TabIndex = 426;
			this.ultraGrid1.AfterRowsDeleted += new System.EventHandler(this.ultraGrid1_AfterRowsDeleted);
			this.ultraGrid1.AfterRowInsert += new Infragistics.Win.UltraWinGrid.RowEventHandler(this.ultraGrid1_AfterRowInsert);
			this.ultraGrid1.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.ultraGrid1_KeyPress);
			this.ultraGrid1.BeforeCellUpdate += new Infragistics.Win.UltraWinGrid.BeforeCellUpdateEventHandler(this.ultraGrid1_BeforeCellUpdate);
			this.ultraGrid1.BeforeRowsDeleted += new Infragistics.Win.UltraWinGrid.BeforeRowsDeletedEventHandler(this.ultraGrid1_BeforeRowsDeleted);
			this.ultraGrid1.KeyDown += new System.Windows.Forms.KeyEventHandler(this.ultraGrid1_KeyDown);
			this.ultraGrid1.AfterCellUpdate += new Infragistics.Win.UltraWinGrid.CellEventHandler(this.ultraGrid1_AfterCellUpdate);
			this.ultraGrid1.InitializeLayout += new Infragistics.Win.UltraWinGrid.InitializeLayoutEventHandler(this.ultraGrid1_InitializeLayout);
			// 
			// ultraDataSource1
			// 
			ultraDataColumn1.DataType = typeof(int);
			ultraDataColumn1.DefaultValue = 0;
			ultraDataColumn2.DataType = typeof(int);
			ultraDataColumn2.DefaultValue = 0;
			ultraDataColumn3.DataType = typeof(int);
			ultraDataColumn3.DefaultValue = 0;
			ultraDataColumn4.DataType = typeof(int);
			ultraDataColumn4.DefaultValue = 0;
			ultraDataColumn5.DataType = typeof(int);
			ultraDataColumn5.DefaultValue = 0;
			ultraDataColumn8.DataType = typeof(System.Decimal);
			ultraDataColumn8.DefaultValue = new System.Decimal(new int[] {
																																		 0,
																																		 0,
																																		 0,
																																		 0});
			ultraDataColumn9.DataType = typeof(int);
			ultraDataColumn9.DefaultValue = 0;
			ultraDataColumn10.DataType = typeof(System.Decimal);
			ultraDataColumn10.DefaultValue = new System.Decimal(new int[] {
																																			0,
																																			0,
																																			0,
																																			0});
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
			ultraDataColumn15.DataType = typeof(bool);
			ultraDataColumn15.DefaultValue = false;
			ultraDataColumn16.DataType = typeof(int);
			ultraDataColumn16.DefaultValue = 0;
			ultraDataColumn17.DataType = typeof(bool);
			ultraDataColumn17.DefaultValue = false;
			ultraDataColumn18.DataType = typeof(int);
			ultraDataColumn18.DefaultValue = 0;
			ultraDataColumn19.DataType = typeof(bool);
			ultraDataColumn19.DefaultValue = false;
			ultraDataColumn20.DataType = typeof(bool);
			ultraDataColumn20.DefaultValue = false;
			ultraDataColumn21.DataType = typeof(int);
			ultraDataColumn21.DefaultValue = 0;
			ultraDataColumn22.DataType = typeof(int);
			ultraDataColumn22.DefaultValue = 0;
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
																																 ultraDataColumn22});
			// 
			// cmbFormaPago
			// 
			appearance22.ForeColorDisabled = System.Drawing.Color.Black;
			this.cmbFormaPago.Appearance = appearance22;
			this.cmbFormaPago.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbFormaPago.Cursor = System.Windows.Forms.Cursors.Default;
			ultraGridColumn23.Header.VisiblePosition = 0;
			ultraGridColumn23.Hidden = true;
			ultraGridColumn24.Header.VisiblePosition = 1;
			ultraGridColumn24.Width = 200;
			ultraGridBand2.Columns.AddRange(new object[] {
																										 ultraGridColumn23,
																										 ultraGridColumn24});
			this.cmbFormaPago.DisplayLayout.BandsSerializer.Add(ultraGridBand2);
			this.cmbFormaPago.DisplayMember = "FormaPago";
			this.cmbFormaPago.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbFormaPago.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbFormaPago.Enabled = false;
			this.cmbFormaPago.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbFormaPago.Location = new System.Drawing.Point(96, 72);
			this.cmbFormaPago.Name = "cmbFormaPago";
			this.cmbFormaPago.Size = new System.Drawing.Size(144, 21);
			this.cmbFormaPago.TabIndex = 422;
			this.cmbFormaPago.ValueMember = "idFormaPago";
			this.cmbFormaPago.KeyDown += new System.Windows.Forms.KeyEventHandler(this.cmbFormaPago_KeyDown);
			this.cmbFormaPago.Validating += new System.ComponentModel.CancelEventHandler(this.cmbFormaPago_Validating);
			// 
			// lblUsuario
			// 
			this.lblUsuario.AutoSize = true;
			this.lblUsuario.BackColor = System.Drawing.Color.Transparent;
			this.lblUsuario.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
			this.lblUsuario.Font = new System.Drawing.Font("Tahoma", 10F, System.Drawing.FontStyle.Bold);
			this.lblUsuario.Location = new System.Drawing.Point(992, 1872);
			this.lblUsuario.Name = "lblUsuario";
			this.lblUsuario.Size = new System.Drawing.Size(0, 23);
			this.lblUsuario.TabIndex = 421;
			this.lblUsuario.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
			// 
			// txtComprobante
			// 
			this.txtComprobante.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			appearance23.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtComprobante.Appearance = appearance23;
			this.txtComprobante.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtComprobante.Enabled = false;
			this.txtComprobante.Location = new System.Drawing.Point(904, 40);
			this.txtComprobante.Name = "txtComprobante";
			this.txtComprobante.Size = new System.Drawing.Size(112, 22);
			this.txtComprobante.TabIndex = 420;
			// 
			// lblEstado
			// 
			this.lblEstado.AutoSize = true;
			this.lblEstado.Font = new System.Drawing.Font("Microsoft Sans Serif", 15F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.lblEstado.ForeColor = System.Drawing.Color.Firebrick;
			this.lblEstado.Location = new System.Drawing.Point(416, 384);
			this.lblEstado.Name = "lblEstado";
			this.lblEstado.Size = new System.Drawing.Size(0, 26);
			this.lblEstado.TabIndex = 418;
			this.lblEstado.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
			// 
			// cmbVendedor
			// 
			appearance24.ForeColorDisabled = System.Drawing.Color.Black;
			this.cmbVendedor.Appearance = appearance24;
			this.cmbVendedor.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbVendedor.Cursor = System.Windows.Forms.Cursors.Default;
			ultraGridColumn25.Header.VisiblePosition = 0;
			ultraGridColumn25.Hidden = true;
			ultraGridColumn26.Header.VisiblePosition = 1;
			ultraGridColumn26.Width = 335;
			ultraGridColumn27.Header.VisiblePosition = 2;
			ultraGridBand3.Columns.AddRange(new object[] {
																										 ultraGridColumn25,
																										 ultraGridColumn26,
																										 ultraGridColumn27});
			this.cmbVendedor.DisplayLayout.BandsSerializer.Add(ultraGridBand3);
			this.cmbVendedor.DisplayMember = "Nombre";
			this.cmbVendedor.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbVendedor.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbVendedor.Enabled = false;
			this.cmbVendedor.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbVendedor.Location = new System.Drawing.Point(552, 8);
			this.cmbVendedor.Name = "cmbVendedor";
			this.cmbVendedor.Size = new System.Drawing.Size(264, 21);
			this.cmbVendedor.TabIndex = 417;
			this.cmbVendedor.ValueMember = "idPersonal";
			this.cmbVendedor.KeyDown += new System.Windows.Forms.KeyEventHandler(this.cmbVendedor_KeyDown);
			this.cmbVendedor.Validating += new System.ComponentModel.CancelEventHandler(this.cmbVendedor_Validating);
			// 
			// cmbTipoRuc
			// 
			appearance25.ForeColorDisabled = System.Drawing.Color.Black;
			this.cmbTipoRuc.Appearance = appearance25;
			this.cmbTipoRuc.CausesValidation = false;
			this.cmbTipoRuc.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbTipoRuc.Cursor = System.Windows.Forms.Cursors.Default;
			ultraGridColumn28.Header.VisiblePosition = 0;
			ultraGridColumn28.Hidden = true;
			ultraGridColumn29.Header.VisiblePosition = 1;
			ultraGridColumn29.Width = 233;
			ultraGridBand4.Columns.AddRange(new object[] {
																										 ultraGridColumn28,
																										 ultraGridColumn29});
			this.cmbTipoRuc.DisplayLayout.BandsSerializer.Add(ultraGridBand4);
			this.cmbTipoRuc.DisplayMember = "TipoRuc";
			this.cmbTipoRuc.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbTipoRuc.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbTipoRuc.Enabled = false;
			this.cmbTipoRuc.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbTipoRuc.Location = new System.Drawing.Point(96, 40);
			this.cmbTipoRuc.Name = "cmbTipoRuc";
			this.cmbTipoRuc.Size = new System.Drawing.Size(144, 21);
			this.cmbTipoRuc.TabIndex = 415;
			this.cmbTipoRuc.ValueMember = "idTipoRuc";
			this.cmbTipoRuc.KeyDown += new System.Windows.Forms.KeyEventHandler(this.cmbTipoRuc_KeyDown);
			this.cmbTipoRuc.Validating += new System.ComponentModel.CancelEventHandler(this.cmbTipoRuc_Validating);
			this.cmbTipoRuc.ValueChanged += new System.EventHandler(this.cmbTipoRuc_ValueChanged);
			// 
			// txtRuc
			// 
			appearance26.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtRuc.Appearance = appearance26;
			this.txtRuc.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtRuc.Enabled = false;
			this.txtRuc.Location = new System.Drawing.Point(344, 40);
			this.txtRuc.MaxLength = 3;
			this.txtRuc.Name = "txtRuc";
			this.txtRuc.Size = new System.Drawing.Size(136, 22);
			this.txtRuc.TabIndex = 410;
			this.txtRuc.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtRuc_KeyDown);
			this.txtRuc.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.txtRuc_KeyPress);
			this.txtRuc.Validating += new System.ComponentModel.CancelEventHandler(this.txtRuc_Validating);
			this.txtRuc.Validated += new System.EventHandler(this.txtRuc_Validated);
			this.txtRuc.ValueChanged += new System.EventHandler(this.txtRuc_ValueChanged);
			// 
			// txtNombre
			// 
			appearance27.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtNombre.Appearance = appearance27;
			this.txtNombre.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtNombre.Enabled = false;
			this.txtNombre.Location = new System.Drawing.Point(552, 40);
			this.txtNombre.Name = "txtNombre";
			this.txtNombre.Size = new System.Drawing.Size(264, 22);
			this.txtNombre.TabIndex = 411;
			// 
			// dtFecha
			// 
			appearance28.ForeColorDisabled = System.Drawing.Color.Black;
			this.dtFecha.Appearance = appearance28;
			this.dtFecha.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			dateButton1.Caption = "Today";
			this.dtFecha.DateButtons.Add(dateButton1);
			this.dtFecha.Enabled = false;
			this.dtFecha.Format = "dd/MM/yyyy";
			this.dtFecha.Location = new System.Drawing.Point(96, 8);
			this.dtFecha.Name = "dtFecha";
			this.dtFecha.NonAutoSizeHeight = 23;
			this.dtFecha.Size = new System.Drawing.Size(104, 21);
			this.dtFecha.SpinButtonsVisible = true;
			this.dtFecha.TabIndex = 412;
			this.dtFecha.Value = ((object)(resources.GetObject("dtFecha.Value")));
			this.dtFecha.KeyDown += new System.Windows.Forms.KeyEventHandler(this.dtFecha_KeyDown);
			this.dtFecha.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.dtFecha_KeyPress);
			this.dtFecha.Validating += new System.ComponentModel.CancelEventHandler(this.dtFecha_Validating);
			this.dtFecha.BeforeDropDown += new System.ComponentModel.CancelEventHandler(this.dtFecha_BeforeDropDown);
			this.dtFecha.ValueChanged += new System.EventHandler(this.dtFecha_ValueChanged);
			// 
			// cmbBodega
			// 
			this.cmbBodega.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbBodega.Cursor = System.Windows.Forms.Cursors.Default;
			ultraGridColumn30.Header.VisiblePosition = 2;
			ultraGridColumn31.Header.VisiblePosition = 1;
			ultraGridColumn31.Width = 180;
			ultraGridColumn32.Header.VisiblePosition = 0;
			ultraGridBand5.Columns.AddRange(new object[] {
																										 ultraGridColumn30,
																										 ultraGridColumn31,
																										 ultraGridColumn32});
			this.cmbBodega.DisplayLayout.BandsSerializer.Add(ultraGridBand5);
			this.cmbBodega.DisplayMember = "Nombre";
			this.cmbBodega.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbBodega.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbBodega.Enabled = false;
			this.cmbBodega.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbBodega.Location = new System.Drawing.Point(288, 8);
			this.cmbBodega.MaxDropDownItems = 30;
			this.cmbBodega.Name = "cmbBodega";
			this.cmbBodega.Size = new System.Drawing.Size(192, 21);
			this.cmbBodega.TabIndex = 408;
			this.cmbBodega.ValueMember = "Bodega";
			this.cmbBodega.ValueChanged += new System.EventHandler(this.cmbBodega_ValueChanged);
			this.cmbBodega.InitializeLayout += new Infragistics.Win.UltraWinGrid.InitializeLayoutEventHandler(this.cmbBodega_InitializeLayout);
			// 
			// txtNumero
			// 
			this.txtNumero.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			appearance29.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtNumero.Appearance = appearance29;
			this.txtNumero.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtNumero.Enabled = false;
			this.txtNumero.Location = new System.Drawing.Point(904, 8);
			this.txtNumero.Name = "txtNumero";
			this.txtNumero.Size = new System.Drawing.Size(112, 22);
			this.txtNumero.TabIndex = 406;
			// 
			// groupBox1
			// 
			this.groupBox1.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.groupBox1.Controls.Add(this.label67);
			this.groupBox1.Controls.Add(this.label66);
			this.groupBox1.Controls.Add(this.label19);
			this.groupBox1.Controls.Add(this.cmbDirecciones);
			this.groupBox1.Controls.Add(this.chkEntregado);
			this.groupBox1.Controls.Add(this.dtFechaEntrega);
			this.groupBox1.Controls.Add(this.dtpHoraEntrega);
			this.groupBox1.Location = new System.Drawing.Point(320, 288);
			this.groupBox1.Name = "groupBox1";
			this.groupBox1.Size = new System.Drawing.Size(440, 88);
			this.groupBox1.TabIndex = 455;
			this.groupBox1.TabStop = false;
			// 
			// label67
			// 
			this.label67.AutoSize = true;
			this.label67.BackColor = System.Drawing.Color.Transparent;
			this.label67.Location = new System.Drawing.Point(304, 26);
			this.label67.Name = "label67";
			this.label67.Size = new System.Drawing.Size(28, 17);
			this.label67.TabIndex = 257;
			this.label67.Text = "Hora";
			this.label67.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// label66
			// 
			this.label66.AutoSize = true;
			this.label66.BackColor = System.Drawing.Color.Transparent;
			this.label66.Location = new System.Drawing.Point(8, 58);
			this.label66.Name = "label66";
			this.label66.Size = new System.Drawing.Size(50, 17);
			this.label66.TabIndex = 255;
			this.label66.Text = "Dirección";
			this.label66.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// label19
			// 
			this.label19.AutoSize = true;
			this.label19.BackColor = System.Drawing.Color.Transparent;
			this.label19.Location = new System.Drawing.Point(136, 26);
			this.label19.Name = "label19";
			this.label19.Size = new System.Drawing.Size(34, 17);
			this.label19.TabIndex = 252;
			this.label19.Text = "Fecha";
			this.label19.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// cmbDirecciones
			// 
			appearance30.ForeColorDisabled = System.Drawing.Color.Black;
			this.cmbDirecciones.Appearance = appearance30;
			this.cmbDirecciones.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbDirecciones.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbDirecciones.DisplayLayout.AutoFitColumns = true;
			ultraGridColumn33.Header.VisiblePosition = 0;
			ultraGridColumn33.Hidden = true;
			ultraGridColumn34.Header.VisiblePosition = 1;
			ultraGridColumn34.Width = 325;
			ultraGridBand6.Columns.AddRange(new object[] {
																										 ultraGridColumn33,
																										 ultraGridColumn34});
			this.cmbDirecciones.DisplayLayout.BandsSerializer.Add(ultraGridBand6);
			this.cmbDirecciones.DisplayMember = "Direccion";
			this.cmbDirecciones.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbDirecciones.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbDirecciones.Enabled = false;
			this.cmbDirecciones.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbDirecciones.Location = new System.Drawing.Point(88, 56);
			this.cmbDirecciones.Name = "cmbDirecciones";
			this.cmbDirecciones.Size = new System.Drawing.Size(344, 21);
			this.cmbDirecciones.TabIndex = 256;
			this.cmbDirecciones.ValueMember = "idDreccion";
			this.cmbDirecciones.KeyDown += new System.Windows.Forms.KeyEventHandler(this.cmbDirecciones_KeyDown);
			this.cmbDirecciones.Validating += new System.ComponentModel.CancelEventHandler(this.cmbDirecciones_Validating);
			// 
			// chkEntregado
			// 
			appearance31.ForeColorDisabled = System.Drawing.Color.Black;
			this.chkEntregado.Appearance = appearance31;
			this.chkEntregado.BackColor = System.Drawing.Color.Transparent;
			this.chkEntregado.CausesValidation = false;
			this.chkEntregado.Enabled = false;
			this.chkEntregado.Location = new System.Drawing.Point(8, 8);
			this.chkEntregado.Name = "chkEntregado";
			this.chkEntregado.Size = new System.Drawing.Size(120, 18);
			this.chkEntregado.TabIndex = 254;
			this.chkEntregado.Text = "Entrega a domicilio";
			this.chkEntregado.KeyDown += new System.Windows.Forms.KeyEventHandler(this.chkEntregado_KeyDown);
			this.chkEntregado.CheckedChanged += new System.EventHandler(this.chkEntregado_CheckedChanged);
			// 
			// dtFechaEntrega
			// 
			appearance32.ForeColorDisabled = System.Drawing.Color.Black;
			this.dtFechaEntrega.Appearance = appearance32;
			this.dtFechaEntrega.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			dateButton2.Caption = "Today";
			this.dtFechaEntrega.DateButtons.Add(dateButton2);
			this.dtFechaEntrega.Enabled = false;
			this.dtFechaEntrega.Format = "dd/MM/yyyy";
			this.dtFechaEntrega.Location = new System.Drawing.Point(184, 24);
			this.dtFechaEntrega.Name = "dtFechaEntrega";
			this.dtFechaEntrega.NonAutoSizeHeight = 23;
			this.dtFechaEntrega.Size = new System.Drawing.Size(112, 21);
			this.dtFechaEntrega.SpinButtonsVisible = true;
			this.dtFechaEntrega.TabIndex = 253;
			this.dtFechaEntrega.Value = ((object)(resources.GetObject("dtFechaEntrega.Value")));
			this.dtFechaEntrega.KeyDown += new System.Windows.Forms.KeyEventHandler(this.dtFechaEntrega_KeyDown);
			this.dtFechaEntrega.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.dtFechaEntrega_KeyPress);
			this.dtFechaEntrega.Validating += new System.ComponentModel.CancelEventHandler(this.dtFechaEntrega_Validating);
			// 
			// dtpHoraEntrega
			// 
			appearance33.ForeColorDisabled = System.Drawing.Color.Black;
			this.dtpHoraEntrega.Appearance = appearance33;
			this.dtpHoraEntrega.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			dateButton3.Caption = "Today";
			this.dtpHoraEntrega.DateButtons.Add(dateButton3);
			this.dtpHoraEntrega.Enabled = false;
			this.dtpHoraEntrega.Format = "HH:mm";
			this.dtpHoraEntrega.Location = new System.Drawing.Point(344, 24);
			this.dtpHoraEntrega.Name = "dtpHoraEntrega";
			this.dtpHoraEntrega.NonAutoSizeHeight = 23;
			this.dtpHoraEntrega.Size = new System.Drawing.Size(88, 21);
			this.dtpHoraEntrega.SpinButtonsVisible = true;
			this.dtpHoraEntrega.TabIndex = 403;
			this.dtpHoraEntrega.Value = ((object)(resources.GetObject("dtpHoraEntrega.Value")));
			this.dtpHoraEntrega.KeyDown += new System.Windows.Forms.KeyEventHandler(this.dtpHoraEntrega_KeyDown);
			this.dtpHoraEntrega.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.dtpHoraEntrega_KeyPress);
			this.dtpHoraEntrega.Validating += new System.ComponentModel.CancelEventHandler(this.dtpHoraEntrega_Validating);
			// 
			// cmbProyecto
			// 
			this.cmbProyecto.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbProyecto.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbProyecto.DataSource = this.ultraDataSource2;
			ultraGridColumn35.Header.VisiblePosition = 0;
			ultraGridColumn35.Hidden = true;
			ultraGridColumn36.Header.VisiblePosition = 1;
			ultraGridColumn37.Header.VisiblePosition = 2;
			ultraGridColumn37.Width = 300;
			ultraGridBand7.Columns.AddRange(new object[] {
																										 ultraGridColumn35,
																										 ultraGridColumn36,
																										 ultraGridColumn37});
			this.cmbProyecto.DisplayLayout.BandsSerializer.Add(ultraGridBand7);
			this.cmbProyecto.DisplayMember = "Nombre";
			this.cmbProyecto.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbProyecto.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbProyecto.Enabled = false;
			this.cmbProyecto.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbProyecto.Location = new System.Drawing.Point(312, 72);
			this.cmbProyecto.MaxDropDownItems = 30;
			this.cmbProyecto.Name = "cmbProyecto";
			this.cmbProyecto.Size = new System.Drawing.Size(250, 21);
			this.cmbProyecto.TabIndex = 466;
			this.cmbProyecto.ValueMember = "idProyectoZoom";
			this.cmbProyecto.KeyDown += new System.Windows.Forms.KeyEventHandler(this.cmbProyecto_KeyDown);
			// 
			// ultraDataSource2
			// 
			ultraDataColumn23.DataType = typeof(int);
			this.ultraDataSource2.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn23,
																																 ultraDataColumn24,
																																 ultraDataColumn25});
			// 
			// label10
			// 
			this.label10.AutoSize = true;
			this.label10.BackColor = System.Drawing.Color.Transparent;
			this.label10.Location = new System.Drawing.Point(248, 74);
			this.label10.Name = "label10";
			this.label10.Size = new System.Drawing.Size(47, 17);
			this.label10.TabIndex = 465;
			this.label10.Text = "Proyecto";
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
			// btnEMail
			// 
			this.btnEMail.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btnEMail.CausesValidation = false;
			this.btnEMail.Enabled = false;
			this.btnEMail.Image = ((System.Drawing.Image)(resources.GetObject("btnEMail.Image")));
			this.btnEMail.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnEMail.Location = new System.Drawing.Point(391, 432);
			this.btnEMail.Name = "btnEMail";
			this.btnEMail.Size = new System.Drawing.Size(77, 23);
			this.btnEMail.TabIndex = 467;
			this.btnEMail.Text = "&E-Mail";
			this.btnEMail.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnEMail.Click += new System.EventHandler(this.btnEMail_Click);
			// 
			// btnProyecto
			// 
			this.btnProyecto.CausesValidation = false;
			this.btnProyecto.Location = new System.Drawing.Point(568, 72);
			this.btnProyecto.Name = "btnProyecto";
			this.btnProyecto.Size = new System.Drawing.Size(24, 21);
			this.btnProyecto.TabIndex = 468;
			this.btnProyecto.Text = "...";
			this.btnProyecto.Click += new System.EventHandler(this.btnProyecto_Click);
			// 
			// label5
			// 
			this.label5.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			this.label5.AutoSize = true;
			this.label5.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label5.ForeColor = System.Drawing.Color.Firebrick;
			this.label5.Location = new System.Drawing.Point(768, 330);
			this.label5.Name = "label5";
			this.label5.Size = new System.Drawing.Size(146, 17);
			this.label5.TabIndex = 469;
			this.label5.Text = "SUBTOTAL - DESCUENTO";
			this.label5.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// txtBaseImponible
			// 
			this.txtBaseImponible.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			appearance34.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtBaseImponible.Appearance = appearance34;
			this.txtBaseImponible.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtBaseImponible.Enabled = false;
			this.txtBaseImponible.FormatString = "#,##0.00";
			this.txtBaseImponible.Location = new System.Drawing.Point(936, 328);
			this.txtBaseImponible.Name = "txtBaseImponible";
			this.txtBaseImponible.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtBaseImponible.PromptChar = ' ';
			this.txtBaseImponible.Size = new System.Drawing.Size(80, 22);
			this.txtBaseImponible.TabIndex = 470;
			// 
			// txtJornadas
			// 
			this.txtJornadas.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtJornadas.Enabled = false;
			this.txtJornadas.FormatString = "#,##0.00";
			this.txtJornadas.Location = new System.Drawing.Point(672, 72);
			this.txtJornadas.MaxValue = 365;
			this.txtJornadas.MinValue = 1;
			this.txtJornadas.Name = "txtJornadas";
			this.txtJornadas.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtJornadas.PromptChar = ' ';
			this.txtJornadas.Size = new System.Drawing.Size(72, 22);
			this.txtJornadas.TabIndex = 471;
			this.txtJornadas.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtJornadas_KeyDown);
			this.txtJornadas.Click += new System.EventHandler(this.txtJornadas_Click);
			this.txtJornadas.Validated += new System.EventHandler(this.txtJornadas_Validated);
			this.txtJornadas.ValueChanged += new System.EventHandler(this.txtJornadas_ValueChanged);
			this.txtJornadas.Enter += new System.EventHandler(this.txtJornadas_Enter);
			// 
			// label7
			// 
			this.label7.AutoSize = true;
			this.label7.BackColor = System.Drawing.Color.Transparent;
			this.label7.Location = new System.Drawing.Point(608, 74);
			this.label7.Name = "label7";
			this.label7.Size = new System.Drawing.Size(49, 17);
			this.label7.TabIndex = 472;
			this.label7.Text = "Jornadas";
			// 
			// label9
			// 
			this.label9.AutoSize = true;
			this.label9.Location = new System.Drawing.Point(784, 74);
			this.label9.Name = "label9";
			this.label9.Size = new System.Drawing.Size(57, 17);
			this.label9.TabIndex = 473;
			this.label9.Text = "Descuento";
			this.label9.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.label9.Visible = false;
			// 
			// txtDescuento
			// 
			appearance35.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtDescuento.Appearance = appearance35;
			this.txtDescuento.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtDescuento.Enabled = false;
			this.txtDescuento.FormatString = "#,##0.00";
			this.txtDescuento.Location = new System.Drawing.Point(856, 72);
			this.txtDescuento.Name = "txtDescuento";
			this.txtDescuento.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtDescuento.PromptChar = ' ';
			this.txtDescuento.Size = new System.Drawing.Size(88, 22);
			this.txtDescuento.TabIndex = 474;
			this.txtDescuento.Visible = false;
			this.txtDescuento.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtDescuento_KeyDown);
			this.txtDescuento.Validated += new System.EventHandler(this.txtDescuento_Validated);
			// 
			// btnAplicarDescuento
			// 
			this.btnAplicarDescuento.CausesValidation = false;
			this.btnAplicarDescuento.Enabled = false;
			this.btnAplicarDescuento.Location = new System.Drawing.Point(952, 72);
			this.btnAplicarDescuento.Name = "btnAplicarDescuento";
			this.btnAplicarDescuento.Size = new System.Drawing.Size(24, 21);
			this.btnAplicarDescuento.TabIndex = 475;
			this.btnAplicarDescuento.Text = "...";
			this.btnAplicarDescuento.Visible = false;
			this.btnAplicarDescuento.Click += new System.EventHandler(this.btnAplicarDescuento_Click);
			// 
			// btnDuplicar
			// 
			this.btnDuplicar.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btnDuplicar.CausesValidation = false;
			this.btnDuplicar.Enabled = false;
			this.btnDuplicar.Image = ((System.Drawing.Image)(resources.GetObject("btnDuplicar.Image")));
			this.btnDuplicar.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnDuplicar.Location = new System.Drawing.Point(237, 432);
			this.btnDuplicar.Name = "btnDuplicar";
			this.btnDuplicar.TabIndex = 476;
			this.btnDuplicar.Text = "&Duplicar";
			this.btnDuplicar.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnDuplicar.Click += new System.EventHandler(this.btnDuplicar_Click);
			// 
			// btnJornadas
			// 
			this.btnJornadas.CausesValidation = false;
			this.btnJornadas.Enabled = false;
			this.btnJornadas.Location = new System.Drawing.Point(752, 72);
			this.btnJornadas.Name = "btnJornadas";
			this.btnJornadas.Size = new System.Drawing.Size(24, 21);
			this.btnJornadas.TabIndex = 477;
			this.btnJornadas.Text = "...";
			this.btnJornadas.Click += new System.EventHandler(this.btnJornadas_Click);
			// 
			// frmCotizaciones
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 14);
			this.ClientSize = new System.Drawing.Size(1024, 462);
			this.Controls.Add(this.btnJornadas);
			this.Controls.Add(this.btnDuplicar);
			this.Controls.Add(this.btnAplicarDescuento);
			this.Controls.Add(this.label9);
			this.Controls.Add(this.label7);
			this.Controls.Add(this.label5);
			this.Controls.Add(this.label10);
			this.Controls.Add(this.label4);
			this.Controls.Add(this.label3);
			this.Controls.Add(this.lblSubTotal);
			this.Controls.Add(this.lblIva);
			this.Controls.Add(this.lblDescuento);
			this.Controls.Add(this.lblIvaSubt);
			this.Controls.Add(this.lblIva0);
			this.Controls.Add(this.label2);
			this.Controls.Add(this.lblIdentificacion);
			this.Controls.Add(this.label8);
			this.Controls.Add(this.lblComprobante);
			this.Controls.Add(this.lblVendedor);
			this.Controls.Add(this.lblCliente);
			this.Controls.Add(this.label6);
			this.Controls.Add(this.label1);
			this.Controls.Add(this.lblBodega);
			this.Controls.Add(this.txtDescuento);
			this.Controls.Add(this.txtJornadas);
			this.Controls.Add(this.txtBaseImponible);
			this.Controls.Add(this.btnProyecto);
			this.Controls.Add(this.btnEMail);
			this.Controls.Add(this.cmbProyecto);
			this.Controls.Add(this.btnCliente);
			this.Controls.Add(this.txtNotas);
			this.Controls.Add(this.txtTotal);
			this.Controls.Add(this.txtIvaTotal);
			this.Controls.Add(this.txtDescIva0Total);
			this.Controls.Add(this.txtDescIvaTotal);
			this.Controls.Add(this.txtIva);
			this.Controls.Add(this.txtIva0);
			this.Controls.Add(this.txtContadoCredito);
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
			this.Controls.Add(this.ultraGrid1);
			this.Controls.Add(this.cmbFormaPago);
			this.Controls.Add(this.lblUsuario);
			this.Controls.Add(this.txtComprobante);
			this.Controls.Add(this.lblEstado);
			this.Controls.Add(this.cmbVendedor);
			this.Controls.Add(this.cmbTipoRuc);
			this.Controls.Add(this.txtRuc);
			this.Controls.Add(this.txtNombre);
			this.Controls.Add(this.dtFecha);
			this.Controls.Add(this.cmbBodega);
			this.Controls.Add(this.txtNumero);
			this.Controls.Add(this.groupBox1);
			this.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.KeyPreview = true;
			this.Name = "frmCotizaciones";
			this.Text = "Cotizaciones";
			this.KeyDown += new System.Windows.Forms.KeyEventHandler(this.frmCotizaciones_KeyDown);
			this.Closing += new System.ComponentModel.CancelEventHandler(this.frmCotizaciones_Closing);
			this.Load += new System.EventHandler(this.frmCotizaciones_Load);
			((System.ComponentModel.ISupportInitialize)(this.txtNotas)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtTotal)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIvaTotal)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtDescIva0Total)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtDescIvaTotal)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIva)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIva0)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtContadoCredito)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtBodegaPredef)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtEstado)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIdCliente)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIdCompra)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraGrid1)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource1)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbFormaPago)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtComprobante)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbVendedor)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbTipoRuc)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtRuc)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNombre)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.dtFecha)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbBodega)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNumero)).EndInit();
			this.groupBox1.ResumeLayout(false);
			((System.ComponentModel.ISupportInitialize)(this.cmbDirecciones)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.dtFechaEntrega)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.dtpHoraEntrega)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbProyecto)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource2)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtBaseImponible)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtJornadas)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtDescuento)).EndInit();
			this.ResumeLayout(false);

		}
		#endregion

		#region Variables
		private Acceso miAcceso;
		decimal dPorcIva = 0.00m;	
		bool bNuevo = false;
		bool bEdicion = false;
		decimal MaximoCredito = 0.00m;
		bool bRastreo = false;		
		private bool bActualiza = true;		
		bool bAnulado = false;
		bool bEliminaAlValidar = false;
		int idGrupoCliente = 0;
		int idBloqueaTransacciones = 0;
		decimal dClavePorcDscto = 0.00m;
		int idTipoPrecio = 3;
		CultureInfo us = new CultureInfo("en-US");
		bool bDescuentoLineal = false;
		int iNumero = 0;	
		#endregion Variables

		private void GeneraInventarioTemporal()
		{
			iNumero = Funcion.iEscalarSQL(cdsSeteoF, "Exec CreaInventarioTemporalZoom");
		}

		private void frmCotizaciones_Load(object sender, System.EventArgs e)
		{
			miAcceso = new Acceso(cdsSeteoF, "0702");

			if (!Funcion.Permiso("117", cdsSeteoF))
			{				
				BeginInvoke(new MethodInvoker(this.Close));				
			}
			else
			{
				Funcion.EjecutaSQL(cdsSeteoF, "Exec AuditaCrear 53, 6, 'PEDIDOS'", true);

				bRastreo = Funcion.bEjecutaSQL(cdsSeteoF, "Exec SeteoGExiste 'RASTREO'");
				
				//if (!bDescuentoLineal) this.ultraGrid1.DisplayLayout.Bands[0].Columns["DescuentoPorc"].CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;

				this.dtFecha.CalendarInfo.MaxDate = DateTime.Today;								
				this.txtBodegaPredef.Value = Funcion.iEscalarSQL(cdsSeteoF, string.Format("Exec BodegaAsignadaPorUsuario {0}", MenuLatinium.IdUsuario));				
				this.cmbFormaPago.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "Exec CompraFormaPagoCarga 1, 0");
				this.cmbProyecto.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "SELECT idProyectoZoom, Numero, Nombre FROM ProyectoZoom Order By Numero");

				FuncionesProcedimientos.EstadoGrids(false, this.ultraGrid1);			

				if (IdCompra > 0) this.Consulta(IdCompra);
			}

			//Actualizar el reporte 
			//Funcion.VersionamientoReporte("OrdenCompra.Rpt");
		}

		private void cdsSeteoF_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			cdsSeteoF.Schema.Connections[0].ConnectionString = miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);	
		}

		private void btnSalir_Click(object sender, System.EventArgs e)
		{
			this.Close();
		}

		private void btnCancelar_Click(object sender, System.EventArgs e)
		{
			if (bEdicion && !bNuevo)
			{
				string sSQLBloqueo = string.Format("Exec RegistraBloqueoTransacciones {0}, 0, {1}, 1", idBloqueaTransacciones, (int)this.txtIdCompra.Value);
				Funcion.EjecutaSQL(cdsSeteoF, sSQLBloqueo, true);
			}

			bNuevo = false;
			bEdicion = false;

			this.txtIdCompra.Value = 0;
			this.txtIdCliente.Value = 0;
			this.txtContadoCredito.Value = 0;
			this.txtEstado.Value = 0;

			this.cmbFormaPago.Value = System.DBNull.Value;
			this.cmbVendedor.Value = System.DBNull.Value;
			this.cmbBodega.Value = System.DBNull.Value;
			this.cmbTipoRuc.Value = System.DBNull.Value;
			this.txtRuc.Text = "";
			this.txtNombre.Text = "";
			this.txtNumero.Text = "";
			this.dtFecha.Value = System.DBNull.Value;
			this.txtComprobante.Text = "";
			this.lblEstado.Text = "";
			this.lblUsuario.Text = "";
			this.cmbProyecto.Value = System.DBNull.Value;
			this.txtJornadas.Value = 1;
			this.txtDescuento.Value = 0;
			this.txtNotas.Text = "";

			this.ultraGrid1.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "Exec ConsultaDetalleCotizaciones 0");
			this.cmbDirecciones.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "Select idDreccion, Direccion From ClientesDirecciones Where idCliente = 0");

			this.chkEntregado.Checked = false;
			this.dtFechaEntrega.Value = System.DBNull.Value;
			this.dtpHoraEntrega.Value = System.DBNull.Value;
			this.cmbDirecciones.Value = System.DBNull.Value;			

			this.txtIva0.Value = 0.00m;
			this.txtDescIva0Total.Value = 0.00m;
			this.txtIva.Value = 0.00m;
			this.txtIvaTotal.Value = 0.00m;			
			this.txtDescIvaTotal.Value = 0.00m;
			this.txtTotal.Value = 0.00m;

			this.cmbFormaPago.Enabled = false;
			this.cmbVendedor.Enabled = false;
			this.cmbBodega.Enabled = false;
			this.cmbTipoRuc.Enabled = false;
			this.txtRuc.Enabled = false;
			this.dtFecha.Enabled = false;
			this.cmbProyecto.Enabled = false;
			this.txtJornadas.Enabled = false;
			this.txtDescuento.Enabled = false;

			FuncionesProcedimientos.EstadoGrids(false, this.ultraGrid1);

			this.txtDescIvaTotal.Enabled = false;
			this.txtNotas.Enabled = false;

			this.chkEntregado.Enabled = false;
			this.dtFechaEntrega.Enabled = false;
			this.dtpHoraEntrega.Enabled = false;
			this.cmbDirecciones.Enabled = false;			

			if (miAcceso.Nuevo) this.btnNuevo.Enabled = true;
			if (miAcceso.Buscar) this.btnConsultar.Enabled = true;
			this.btnEditar.Enabled = false;
			this.btnImprimir.Enabled = false;
			this.btnEMail.Enabled = false;
			this.btnGuardar.Enabled = false;
			this.btnAnular.Enabled = false;
			this.btnCancelar.Enabled = false;
			this.btnCliente.Enabled = false;
			this.btnAplicarDescuento.Enabled = false;
			this.btnJornadas.Enabled = false;
			this.btnDuplicar.Enabled = false;

			idGrupoCliente = 0;
			idBloqueaTransacciones = 0;
			dPorcIva = 0.00m;				
			bAnulado = false;
			Funcion.EjecutaSQL(cdsSeteoF, string.Format("Delete From TemporalArticulos Where Numero = {0}", iNumero));
			iNumero = 0;			
		}

		private void btnAnular_Click(object sender, System.EventArgs e)
		{
			this.Consulta((int)this.txtIdCompra.Value);

			if (!Facturacion.ValidaEdicionAnulacion((int)this.txtIdCompra.Value, this.ultraGrid1, miAcceso.Anular, (int)this.txtEstado.Value, 3, (int)this.cmbBodega.Value, 
				(DateTime)this.dtFecha.Value, true, bAnulado, 0, cdsSeteoF, false, miAcceso.BAnularProcesado, miAcceso.BAnularAprobado, miAcceso.BAnularConSeriales, "VENTA", 0, false)) return;

			if (Facturacion.AnulacionCompraVenta((int)this.txtIdCompra.Value, 3, cdsSeteoF))
				this.Consulta((int)this.txtIdCompra.Value);
		}

		private void btnGuardar_Click(object sender, System.EventArgs e)
		{
			#region Elimina Filas Vacias
			bEliminaAlValidar = true;
			Validacion.vEliminaFilasVacias(this.ultraGrid1, "idArticulo");
			bEliminaAlValidar = false;
			#endregion Elimina Filas Vacias

			this.cmdTotal();			
			
			if (!Validacion.vbFechaEnDocumentos(this.dtFecha, "Ingrese la Fecha del Documento", true, "VENTA", cdsSeteoF)) return;			

			if (!Validacion.vbComboVacio(this.cmbVendedor, "Seleccione un vendedor")) return;
			
			if (!Validacion.vbIdentificacion(this.txtRuc, (int)this.cmbTipoRuc.Value, cdsSeteoF)) return;

			if (!Validacion.vbClienteExiste(this.txtRuc, false, 3, (int)this.cmbFormaPago.Value, cdsSeteoF)) return;			
			
			if (!Validacion.vbComboVacio(this.cmbProyecto, "Seleccione un Proyecto")) return;

			if (!Validacion.EntregaDomicilio(this.chkEntregado, this.cmbDirecciones, this.dtFechaEntrega, this.dtpHoraEntrega, 1)) return;

			if (!Validacion.vbGrids(this.ultraGrid1, "Codigo", 3)) return;

			if (!Validacion.ValidaTotales((decimal)this.txtTotal.Value)) return;		
				
			if (DialogResult.Yes == MessageBox.Show("Antes de Guardar. \n\n VERIQUE QUE LOS DATOS REGISTRADOS SON LOS CORRECTOS \n\n INFORMACIÓN DEL CLIENTE \n PRODUCTOS \n PRECIOS CANTIDADES DESCUENTOS \n\n\n ¿Esta seguro de Guardar?", "Point Technology", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2)) 
			{
				#region Guarda Documento
				#region Valida Entrega domicilio
				DateTime dtFecha = (DateTime)this.dtFecha.Value;
				DateTime dtFechaEntrega = DateTime.Today;
				DateTime dtHoraEntrega = DateTime.Today;
				int idDireccion = 0;
				bool bFechaEntrega = false;
				bool bHoraEntrega = false;

				if (this.chkEntregado.Checked)
				{
					dtFechaEntrega = (DateTime)this.dtFechaEntrega.Value;
					dtHoraEntrega = (DateTime)this.dtpHoraEntrega.Value;
					idDireccion = (int)this.cmbDirecciones.Value;

					bFechaEntrega = true;
					bHoraEntrega = true;
				}
				#endregion Valida Entrega domicilio

				if (!bNuevo) 
				{
					if (Funcion.bEscalarSQL(cdsSeteoF, string.Format("Exec CotizacionValidaModificacion {0}, {1}, {2}, {3}, {4}", 
						(int)this.txtIdCompra.Value, (Convert.ToDecimal(this.txtIva0.Value) + Convert.ToDecimal(this.txtIva.Value)), 
						Convert.ToDecimal(this.txtTotal.Value), (Convert.ToDecimal(this.txtDescIva0Total.Value) + Convert.ToDecimal(this.txtDescIvaTotal.Value)), this.ultraGrid1.Rows.Count), true))
					{						
						foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.ultraGrid1.Rows.All)
							dr.Cells["idDetCompra"].Value = 0;						

						this.txtIdCompra.Value = 0;

						bNuevo = true;
					}
				}

				this.txtIdCompra.Value = 
					Facturacion.GrabaFacturacion((int)this.txtIdCompra.Value, 0, 0, 3, 15, (int)this.txtEstado.Value, false, (int)this.txtContadoCredito.Value, (int)this.cmbFormaPago.Value, 
					(int)this.cmbBodega.Value, 0, (int)this.cmbVendedor.Value, this.txtNumero.Text.ToString(), this.txtComprobante.Text.ToString(), 
					(int)this.cmbTipoRuc.Value,	(int)this.txtIdCliente.Value, this.txtRuc.Text.ToString(), this.txtNombre.Text.ToString(), (DateTime)this.dtFecha.Value, "", "", DateTime.Today,
					this.chkEntregado.Checked, dtFechaEntrega, dtHoraEntrega, idDireccion, this.txtNotas.Text.ToString(),
					false, 0.00m, 0, 0.00m, DateTime.Today, 
					Convert.ToDecimal(this.txtIva0.Value), Convert.ToDecimal(this.txtDescIva0Total.Value), Convert.ToDecimal(this.txtIva.Value), 
					Convert.ToDecimal(this.txtDescIvaTotal.Value), Convert.ToDecimal(this.txtIvaTotal.Value), Convert.ToDecimal(this.txtTotal.Value), 
					false, false, bFechaEntrega, bHoraEntrega, false, true, 0, 0, "", 0.00m, 0.00m, 0.00m, 0.00m, 0.00m, false, 0.00m, 0, 0, 0, 0, 0, "", false, 1, "", "", 0, "", 0.00m, (int)this.cmbProyecto.Value, "", Convert.ToDecimal(this.txtJornadas.Value),
					bNuevo, true, idBloqueaTransacciones, this.ultraGrid1,
					false, 0, this.ultraGrid1, false);					
																		
				#region Controles
				if ((int)this.txtIdCompra.Value == 0) return;

				this.cmbFormaPago.Enabled = false;
				this.cmbVendedor.Enabled = false;
				this.cmbTipoRuc.Enabled = false;
				this.cmbBodega.Enabled = false;
				this.txtRuc.Enabled = false;
				this.dtFecha.Enabled = false;									
				this.cmbProyecto.Enabled = false;
				this.txtJornadas.Enabled = false;
				this.txtDescuento.Enabled = false;
				this.txtDescIvaTotal.Enabled = false;
																				
				FuncionesProcedimientos.EstadoGrids(false, this.ultraGrid1);
			
				this.txtNotas.Enabled = false;
				this.chkEntregado.Enabled = false;
				this.dtFechaEntrega.Enabled = false;
				this.dtpHoraEntrega.Enabled = false;
				this.cmbDirecciones.Enabled = false;				
				this.btnCliente.Enabled = false;
				this.btnAplicarDescuento.Enabled = false;
				this.btnJornadas.Enabled = false;
				this.btnDuplicar.Enabled = true;				
				this.btnEMail.Enabled = true;
				bNuevo = false;
				bEdicion = false;
				bEliminaAlValidar = false;
				idBloqueaTransacciones = 0;
				Funcion.EjecutaSQL(cdsSeteoF, string.Format("Delete From TemporalArticulos Where Numero = {0}", iNumero));
				iNumero = 0;

				this.Consulta((int)this.txtIdCompra.Value);
				#endregion Controles										
				#endregion Guarda Documento
			}							
		}

		private void btnImprimir_Click(object sender, System.EventArgs e)
		{
			this.Consulta((int)this.txtIdCompra.Value);

			//Facturacion.Impresion((int)this.txtIdCompra.Value, bAnulado, (int)this.txtEstado.Value, miAcceso.Imprimir, miAcceso.ReImprimir, true, false, 3, 0.00m, 0, 0, "", DateTime.Today, 0, "", cdsSeteoF);
			ParameterFields paramFields = new ParameterFields ();
				
			ParameterField paramFieldPeriodo = new ParameterField ();
			ParameterDiscreteValue ValIdPeriodo = new ParameterDiscreteValue ();
			paramFieldPeriodo.ParameterFieldName = "@idCompra";
			ValIdPeriodo.Value = (int)this.txtIdCompra.Value;
			paramFieldPeriodo.CurrentValues.Add (ValIdPeriodo);
			paramFields.Add (paramFieldPeriodo);

			Reporte Reporte = new Reporte("CotizacionZOOM.rpt", "");
			Reporte.MdiParent = this.MdiParent;			
			Reporte.crVista.ParameterFieldInfo = paramFields;
			Reporte.Show();
		}

		private void btnEditar_Click(object sender, System.EventArgs e)
		{
			this.Consulta((int)this.txtIdCompra.Value);
			
			if (!Facturacion.ValidaEdicionAnulacion((int)this.txtIdCompra.Value, this.ultraGrid1, miAcceso.Editar, (int)this.txtEstado.Value, 3, (int)this.cmbBodega.Value, 
				(DateTime)this.dtFecha.Value, false, bAnulado, 0, cdsSeteoF, false, miAcceso.BEditarProcesado, miAcceso.BEditarAprobado, false, "VENTA", 0, false)) return;
		
			if (miAcceso.BModifcarFormaDePago) this.cmbFormaPago.Enabled = true;

			this.cmbVendedor.Enabled = true;
			this.cmbTipoRuc.Enabled = true;
			this.cmbProyecto.Enabled = true;
			this.txtRuc.Enabled = true;
			if (miAcceso.CambiarFecha) this.dtFecha.Enabled = true;	
					
			FuncionesProcedimientos.EstadoGrids(true, this.ultraGrid1);

			this.txtNotas.Enabled = true;
			this.chkEntregado.Enabled = true;

			if (this.chkEntregado.Checked)
			{
				this.dtFechaEntrega.Enabled = true;
				this.dtpHoraEntrega.Enabled = true;
				this.cmbDirecciones.Enabled = true;				
			}
					
			//if (!bDescuentoLineal) this.txtDescIvaTotal.Enabled = true;

			this.dtFechaEntrega.CalendarInfo.MinDate = (DateTime) this.dtFecha.Value;
			this.txtJornadas.Enabled = true;
			this.txtDescuento.Enabled = true;

			this.btnNuevo.Enabled = false;
			this.btnConsultar.Enabled = false;
			this.btnEditar.Enabled = false;
			this.btnImprimir.Enabled = false;
			this.btnEMail.Enabled = false;
			this.btnGuardar.Enabled = true;
			this.btnCancelar.Enabled = true;
			this.btnAnular.Enabled = false;			
			this.btnCliente.Enabled = true;
			this.btnAplicarDescuento.Enabled = true;
			this.btnJornadas.Enabled = true;
			this.btnDuplicar.Enabled = false;
			this.btnEMail.Enabled = false;

			string sSQLBloqueo = string.Format("Exec RegistraBloqueoTransacciones 0, 1, {0}, 1", (int)this.txtIdCompra.Value);
			idBloqueaTransacciones = Funcion.iEscalarSQL(cdsSeteoF, sSQLBloqueo, true);

			GeneraInventarioTemporal();

			bNuevo = false;
			bEdicion = true;
		}

		private void Consulta(int idCompra)
		{			
			try
			{
				#region Botones 
				if (miAcceso.Nuevo) this.btnNuevo.Enabled = true;
				if (miAcceso.Buscar) this.btnConsultar.Enabled = true;
				this.btnEditar.Enabled = false;
				this.btnImprimir.Enabled = false;
				this.btnGuardar.Enabled = false;
				this.btnAnular.Enabled = false;
				this.btnCancelar.Enabled = false;
				this.btnCliente.Enabled = false;
				this.btnEMail.Enabled = false;
				#endregion Botones

				string sSQL = string.Format("Exec PedidoConsultaMaestroIndividual {0}", idCompra);
				SqlDataReader dr = Funcion.rEscalarSQL(cdsSeteoF, sSQL, true);
				dr.Read();
				if (dr.HasRows)
				{
					this.txtIdCompra.Value = idCompra;
					this.cmbFormaPago.Value = dr.GetInt32(1);
					this.txtContadoCredito.Value = dr.GetInt32(2);					
					this.dtFecha.Value = dr.GetDateTime(10);
					this.cmbBodega.Value = dr.GetInt32(4);
					this.cmbVendedor.Value = dr.GetInt32(3);
					this.cmbTipoRuc.Value = dr.GetInt32(5);
					this.txtIdCliente.Value = dr.GetInt32(6);
					this.txtRuc.Text = dr.GetString(7);
					this.txtNombre.Text = dr.GetString(8);
					this.txtNumero.Text = dr.GetString(9);					
					this.txtComprobante.Text = dr.GetValue(11).ToString();
					this.txtNotas.Text = dr.GetValue(12).ToString();
					this.chkEntregado.Checked = dr.GetBoolean(13);
					if (this.chkEntregado.Checked)
					{
						this.dtFechaEntrega.Value = dr.GetDateTime(14); 
						this.dtpHoraEntrega.Value = dr.GetDateTime(15); 
						this.cmbDirecciones.Value = dr.GetInt32(16); 						
					}
					this.txtIva0.Value = dr.GetDecimal(18);
					this.txtDescIva0Total.Value = dr.GetDecimal(19);
					this.txtIva.Value = dr.GetDecimal(20);
					this.txtDescIvaTotal.Value = dr.GetDecimal(21);
					this.txtIvaTotal.Value = dr.GetDecimal(22);
					this.txtTotal.Value = dr.GetDecimal(23);
					this.txtEstado.Value = dr.GetInt32(24);
					bAnulado = dr.GetBoolean(25);
					this.cmbProyecto.Value = dr.GetInt32(26);
					//this.txtJornadas.Value = dr.GetInt32(27);
				}
				dr.Close();

				if (bAnulado) this.lblEstado.Text = "ANULADO";
				else
				{
					if ((int)this.txtEstado.Value == 1) this.lblEstado.Text = "PENDIENTE";
					if ((int)this.txtEstado.Value == 2) this.lblEstado.Text = "MODIFICADA";				
					if ((int)this.txtEstado.Value == 3) this.lblEstado.Text = "PROCESADO";
					if ((int)this.txtEstado.Value == 4) this.lblEstado.Text = "NEGADO";
					if ((int)this.txtEstado.Value == 5) this.lblEstado.Text = "APROBADO";
					if ((int)this.txtEstado.Value == 9) this.lblEstado.Text = "FACTURADO";
				}

				this.txtBaseImponible.Value = Math.Round(Convert.ToDecimal(this.txtIva.Value) - Convert.ToDecimal(this.txtDescIvaTotal.Value), 2);

				this.cmbDirecciones.DataSource = Funcion.dvProcedimiento(cdsSeteoF, string.Format("Select idDreccion, Direccion From ClientesDirecciones Where idCliente = {0}", (int)this.txtIdCliente.Value));

				this.cmbVendedor.DataSource = FuncionesProcedimientos.Lista_Vendedores(Convert.ToDateTime(this.dtFecha.Value), (int)this.cmbBodega.Value, cdsSeteoF);

				this.ultraGrid1.DataSource = Funcion.dvProcedimiento(cdsSeteoF, string.Format("Exec ConsultaDetalleCotizaciones {0}", (int)this.txtIdCompra.Value));

				if (!Funcion.ValidaPeriodo((DateTime)this.dtFecha.Value, cdsSeteoF, "Venta"))
				{
					if (!bAnulado)
					{
						if (miAcceso.Editar)
						{						
							if ((int)this.txtEstado.Value == 1) this.btnEditar.Enabled = true;
							else if ((int)this.txtEstado.Value == 4) this.btnEditar.Enabled = true;
							else if ((int)this.txtEstado.Value == 5 && miAcceso.BEditarAprobado) this.btnEditar.Enabled = true;
							else if ((int)this.txtEstado.Value == 3 && miAcceso.BEditarProcesado) this.btnEditar.Enabled = true;
						}

						if (miAcceso.Imprimir) this.btnImprimir.Enabled = true;

						if (miAcceso.Anular && !bAnulado)
						{
							if ((int)this.txtEstado.Value == 1) this.btnAnular.Enabled = true;
							else if ((int)this.txtEstado.Value == 5 && miAcceso.BAnularAprobado) this.btnAnular.Enabled = true;
							else if ((int)this.txtEstado.Value == 3 && miAcceso.BAnularProcesado) this.btnAnular.Enabled = true;
						}	
	
						this.btnDuplicar.Enabled = true;
						this.btnEMail.Enabled = true;
					}
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

			using (frmBuscaPedidos Busqueda = new frmBuscaPedidos(3))
			{
				if (DialogResult.OK == Busqueda.ShowDialog())
				{
					this.Consulta((int)Busqueda.uGridPedidos.ActiveRow.Cells["idCompra"].Value);
				}
			}
		}

		private void btnNuevo_Click(object sender, System.EventArgs e)
		{
			this.btnCancelar_Click(sender, e);
						
			using (frmFormasDePago FormaPago = new frmFormasDePago(3, false, false, 0))
			{				
				if (DialogResult.OK == FormaPago.ShowDialog())
				{
					this.cmbFormaPago.Value = (int) FormaPago.ultraGrid1.ActiveRow.Cells["idFormaPago"].Value;

					this.cmbTipoRuc.Value = 4;

					if ((int)this.cmbFormaPago.Value == 6) this.txtContadoCredito.Value = 2; else this.txtContadoCredito.Value = 1;
					
					this.dtFecha.Value = DateTime.Today;
					this.dtFechaEntrega.CalendarInfo.MinDate = DateTime.Today;
					if ((int)this.txtBodegaPredef.Value == 0) this.cmbBodega.Value = 1; else this.cmbBodega.Value = (int)this.txtBodegaPredef.Value;
					this.txtJornadas.Value = 1;

					#region Habilita controles
					this.cmbVendedor.Enabled = true;
					this.cmbTipoRuc.Enabled = true;
					this.txtRuc.Enabled = true;
					if (miAcceso.CambiarFecha) this.dtFecha.Enabled = true;
					this.cmbProyecto.Enabled = true;
					this.txtJornadas.Enabled = true;
					this.txtDescuento.Enabled = true;

					if ((int)this.txtBodegaPredef.Value == 0) 
					{
						this.cmbBodega.Value = 1;
						this.cmbBodega.Enabled = true;
					}
					else this.cmbBodega.Value = (int)this.txtBodegaPredef.Value;

					this.txtEstado.Value = 1;

					FuncionesProcedimientos.EstadoGrids(true, this.ultraGrid1);
												
					this.chkEntregado.Enabled = true;
					this.txtNotas.Enabled = true;
					this.txtDescIvaTotal.Enabled = true;

					this.btnNuevo.Enabled = false;
					this.btnConsultar.Enabled = false;
					this.btnEditar.Enabled = false;
					this.btnImprimir.Enabled = false;
					this.btnGuardar.Enabled = true;
					this.btnCancelar.Enabled = true;
					this.btnAnular.Enabled = false;
					this.btnCliente.Enabled = true;
					this.btnJornadas.Enabled = true;
					this.btnAplicarDescuento.Enabled = true;
					this.btnDuplicar.Enabled = false;
					this.btnEMail.Enabled = false;					
//					if (!bDescuentoLineal) 
//					{
//						this.txtDescIvaTotal.Enabled = true; 
//					}
					#endregion Habilita controles					

					this.cmbVendedor.Focus();

					GeneraInventarioTemporal();

					bNuevo = true;
					bEdicion = true;											
				}
				else this.btnCancelar_Click(sender, e);
			}      
		}

		private void frmCotizaciones_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (!bEdicion)
			{
				if (e.KeyCode == Keys.N) if (miAcceso.Nuevo)	if (this.btnNuevo.Enabled) this.btnNuevo_Click(sender, e);
				if (e.KeyCode == Keys.B) if (miAcceso.Buscar) if (this.btnConsultar.Enabled) this.btnConsultar_Click(sender, e);
				if (e.KeyCode == Keys.E) if (miAcceso.Editar) if (this.btnEditar.Enabled) this.btnEditar_Click(sender, e);
				if (e.KeyCode == Keys.I) if (miAcceso.Imprimir) if (this.btnImprimir.Enabled) this.btnImprimir_Click(sender, e);
				if (e.KeyCode == Keys.A) if (miAcceso.Anular && !bAnulado) if (this.btnAnular.Enabled) this.btnAnular_Click(sender, e);				
				if (e.KeyCode == Keys.F5) 
				{
					miAcceso = new Acceso(cdsSeteoF, "0706");
					this.Consulta((int)this.txtIdCompra.Value);					
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

		private void dtFecha_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Delete) e.Handled = true;

			if (e.KeyCode == Keys.Enter)
				if (this.cmbBodega.Enabled) this.cmbBodega.Focus(); else this.cmbVendedor.Focus();
		}

		private void dtFecha_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
			if (e.KeyChar == 8) e.Handled = true;
		}

		private void dtFecha_Validating(object sender, System.ComponentModel.CancelEventArgs e)
		{
			if (bEdicion)
				if (!Validacion.vbFechaEnDocumentos(this.dtFecha, "Ingrese la Fecha", true, "Venta", cdsSeteoF)) e.Cancel = true;
		}

		private void dtFecha_ValueChanged(object sender, System.EventArgs e)
		{
			if (this.dtFecha.Value != System.DBNull.Value)
				this.cmbBodega.DataSource = Funcion.dvProcedimiento(cdsSeteoF, string.Format("Exec FacturacionListaBodegas {0}, 3, '{1}'", MenuLatinium.IdUsuario, Convert.ToDateTime(this.dtFecha.Value).ToString("yyyyMMdd")));
		}

		private void cmbBodega_ValueChanged(object sender, System.EventArgs e)
		{
			if (this.cmbBodega.ActiveRow != null)
			{
				if (bNuevo) this.txtNumero.Text = "";
								
				if (this.cmbVendedor.ActiveRow != null) this.cmbVendedor.Value = System.DBNull.Value;

				this.cmbVendedor.DataSource = FuncionesProcedimientos.Lista_Vendedores(Convert.ToDateTime(this.dtFecha.Value), (int)this.cmbBodega.Value, cdsSeteoF);
				this.cmbTipoRuc.DataSource = FuncionesProcedimientos.Lista_TiposDeIdentificacion((int)this.cmbBodega.Value, 1, cdsSeteoF);

				if (this.cmbTipoRuc.ActiveRow == null) this.cmbTipoRuc.Value = 4;			
			}
		}

		private void cmbVendedor_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Enter) this.txtRuc.Focus();
		}

		private void cmbVendedor_Validating(object sender, System.ComponentModel.CancelEventArgs e)
		{
			if (bEdicion)
				if (!Validacion.vbComboVacio(this.cmbVendedor, "Seleccione un vendedor")) e.Cancel = true;	
		}

		private void cmbTipoRuc_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Enter) this.txtRuc.Focus();
		}

		private void cmbTipoRuc_Validating(object sender, System.ComponentModel.CancelEventArgs e)
		{
			if (bEdicion)
				if (!Validacion.vbComboVacio(this.cmbTipoRuc, "Seleccione un Tipo de Identificación del Cliente")) e.Cancel = true;
		}

		private void cmbTipoRuc_ValueChanged(object sender, System.EventArgs e)
		{
			if (this.cmbTipoRuc.ActiveRow != null)
				if ((int)this.cmbTipoRuc.Value == 4) this.txtRuc.MaxLength = 13; else if ((int)this.cmbTipoRuc.Value == 5) this.txtRuc.MaxLength = 10;
		}

		private void txtRuc_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Enter) this.cmbProyecto.Focus();

			if (e.KeyCode == Keys.F3)
			{
				string sIdentificacion = FuncionesProcedimientos.AbreListaClientesProveedores(false, 3);
				if (sIdentificacion.Length > 0)
				{						
					this.txtRuc.Text = sIdentificacion;
					
					CargaCliente();

					this.ultraGrid1.DisplayLayout.Bands[0].AddNew();
					this.ultraGrid1.ActiveCell = this.ultraGrid1.Rows[0].Cells["Codigo"];		
				}
			}
		}

		private void txtRuc_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
			if ((int)this.cmbTipoRuc.Value == 4 || (int)this.cmbTipoRuc.Value == 5) Funcion.CamposNumericos(sender, e);
		}

		private void txtRuc_Validated(object sender, System.EventArgs e)
		{
			if (bEdicion)
			{
				#region Busca Cliente 
				this.txtIdCliente.Value = 0;
				this.txtNombre.Text = "";

				CargaCliente();

				if ((int)this.txtIdCliente.Value == 0)
				{
					FuncionesProcedimientos.AbreFacturacionClientes((int)this.cmbTipoRuc.Value, this.txtRuc.Text.ToString(), true, false);

					CargaCliente();
				}
				
				this.cmbDirecciones.DataSource = FuncionesProcedimientos.ListaDireccionesClientes((int)this.txtIdCliente.Value, cdsSeteoF);
				#endregion Busca Cliente
			}
		}

		private void txtRuc_Validating(object sender, System.ComponentModel.CancelEventArgs e)
		{
			if (bEdicion)			
			{
				if (!Validacion.vbIdentificacion(this.txtRuc, (int)this.cmbTipoRuc.Value, cdsSeteoF)) 
				{
					e.Cancel = true;
					return;
				}
								
				if (!Validacion.vbClienteExiste(this.txtRuc, false, 3, (int)this.cmbFormaPago.Value, cdsSeteoF)) 
				{
					e.Cancel = true;
					return;
				}
			}
		}

		private void CargaCliente()
		{
			string[] sParametros = new string[11];

			sParametros = FuncionesProcedimientos.RetornaParametrosClientesProveedores(false, 1, this.txtRuc.Text.ToString(), cdsSeteoF);
			
			if (sParametros[0] != null) this.txtIdCliente.Value = int.Parse(sParametros[0]);
			if (sParametros[1] != null) this.cmbTipoRuc.Value = int.Parse(sParametros[1]);
			if (sParametros[2] != null) idGrupoCliente = int.Parse(sParametros[2]);
			if (sParametros[3] != null) this.txtRuc.Text = sParametros[3];
			if (sParametros[4] != null) this.txtNombre.Text = sParametros[4];
			if (sParametros[5] != null) MaximoCredito = Decimal.Parse(sParametros[5]);
		}

		private void btnCliente_Click(object sender, System.EventArgs e)
		{
			if (this.txtRuc.Text.Length == 0)
			{
				MessageBox.Show("Ingrese el Numero de Identificacion del Cliente", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				this.txtRuc.Focus();
			}
			else
			{
				FuncionesProcedimientos.AbreFacturacionClientes((int)this.cmbTipoRuc.Value, this.txtRuc.Text.ToString(), true, false);

				CargaCliente();
			}
		}

		private void cmbFormaPago_Validating(object sender, System.ComponentModel.CancelEventArgs e)
		{
			if (bEdicion)
				if (!Validacion.vbComboVacio(this.cmbFormaPago, "Seleccione la Forma de Pago")) e.Cancel = true;	
		}

		private void cmbFormaPago_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Enter) 
			{
				this.ultraGrid1.DisplayLayout.Bands[0].AddNew();
				this.ultraGrid1.ActiveCell = this.ultraGrid1.Rows[0].Cells["Codigo"];		
			}
		}

		private void chkEntregado_CheckedChanged(object sender, System.EventArgs e)
		{
			if (bEdicion)
			{
				if (this.chkEntregado.Checked)
				{
					if ((int)this.txtIdCliente.Value == 0)
					{
						MessageBox.Show("Ingrese un Cliente", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
						this.chkEntregado.Checked	= false;
						this.txtRuc.Focus();
						return;
					}

					if (this.cmbDirecciones.Rows.Count == 0)
					{
						MessageBox.Show("El Cliente No Tiene Direcciones para la Entrega", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
						this.chkEntregado.Checked	= false;
						this.txtRuc.Focus();
						return;
					}

					this.dtFechaEntrega.Enabled = true;
					this.dtpHoraEntrega.Enabled = true;
					this.cmbDirecciones.Enabled = true;					

					this.dtFechaEntrega.Value = (DateTime)this.dtFecha.Value;
						
					this.dtFechaEntrega.Focus();
				}
				else
				{
					this.dtFechaEntrega.Value = System.DBNull.Value;
					this.dtpHoraEntrega.Value = System.DBNull.Value;
					this.cmbDirecciones.Value = System.DBNull.Value;					

					this.dtFechaEntrega.Enabled = false;
					this.dtpHoraEntrega.Enabled = false;
					this.cmbDirecciones.Enabled = false;					
				}
			}
		}

		private void dtFechaEntrega_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Delete) e.Handled = true;
			if (e.KeyCode == Keys.Enter) this.dtpHoraEntrega.Focus();
		}

		private void dtFechaEntrega_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
			if (e.KeyChar == 8) e.Handled = true;
		}

		private void dtFechaEntrega_Validating(object sender, System.ComponentModel.CancelEventArgs e)
		{
			if (bEdicion)
				if (!Validacion.vbFechaEnDocumentos(this.dtFechaEntrega, "Ingrese la Fecha de Entrega", false, "", cdsSeteoF)) e.Cancel = true;
		}

		private void dtpHoraEntrega_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Delete) e.Handled = true;
			if (e.KeyCode == Keys.Enter) this.cmbDirecciones.Focus();
		}

		private void dtpHoraEntrega_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
			if (e.KeyChar == 8) e.Handled = true;
		}

		private void dtpHoraEntrega_Validating(object sender, System.ComponentModel.CancelEventArgs e)
		{
			if (bEdicion)
				if (!Validacion.vbFechaEnDocumentos(this.dtpHoraEntrega, "Ingrese la hora de entrega", false, "", cdsSeteoF)) e.Cancel = true;		
		}

		private void cmbDirecciones_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Enter) this.btnGuardar.Focus();
		}

		private void cmbDirecciones_Validating(object sender, System.ComponentModel.CancelEventArgs e)
		{
			if (bEdicion)
				if (!Validacion.vbComboVacio(this.cmbDirecciones, "Seleccione la Dirección de Entrega")) e.Cancel = true;
		}

		private void txtNotas_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Enter) this.chkEntregado.Focus();
		}

		private void chkEntregado_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Enter)
			{
				if (this.chkEntregado.Checked) this.dtFechaEntrega.Focus();
				else this.btnGuardar.Focus();
			}
		}

		private void cmdTotal()
		{						
			decimal dDescuentoLineal = 0.00m;

			this.txtIva.Value = 0.00m;
			this.txtIva0.Value = 0.00m;			
			this.txtDescIva0Total.Value = 0.00m;			
			//this.txtDescIvaTotal.Value = 0.00m;						
			this.txtIvaTotal.Value = 0.00m;
			this.txtTotal.Value = 0.00m;

			dPorcIva = Funcion.decEscalarSQL(cdsSeteoF, string.Format("Select dbo.RetornaImpuestoIVAVigente('{0}', 1)", 
				Convert.ToDateTime(this.dtFecha.Value).ToString("yyyyMMdd")));

			#region Desplazamiento de grilla
			foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in ultraGrid1.Rows.All)
			{				
				if (dr.Cells["Cantidad"].Value != DBNull.Value && dr.Cells["Precio"].Value != DBNull.Value && dr.Cells["DescuentoPorc"].Value != DBNull.Value && dr.Cells["Jornadas"].Value != DBNull.Value)
				{
					dr.Cells["Subtotal"].Value = (Math.Round((int)dr.Cells["Cantidad"].Value * Convert.ToDecimal(dr.Cells["Precio"].Value) * (decimal)dr.Cells["Jornadas"].Value, 2));
					
					dDescuentoLineal = Math.Round((((int)dr.Cells["Cantidad"].Value * Convert.ToDecimal(dr.Cells["Precio"].Value) * (decimal)dr.Cells["Jornadas"].Value) * Convert.ToDecimal(dr.Cells["DescuentoPorc"].Value))/100, 2);

					dr.Cells["DescuentoArt"].Value = Math.Round(dDescuentoLineal, 2);

					dr.Cells["Subtotal"].Value = Convert.ToDecimal(dr.Cells["Subtotal"].Value) - dDescuentoLineal;

					if (Convert.ToDecimal(dr.Cells["Impuesto"].Value) > 0.00m)
					{						
						this.txtIva.Value = (Math.Round(Convert.ToDecimal(this.txtIva.Value) + Convert.ToDecimal(dr.Cells["Subtotal"].Value), 2));						
					
						//this.txtDescIvaTotal.Value = (Math.Round(Convert.ToDecimal(this.txtDescIvaTotal.Value) + Convert.ToDecimal(dr.Cells["DescuentoArt"].Value), 2));					
					}
					else
					{				
						this.txtIva0.Value = (Math.Round(Convert.ToDecimal(this.txtIva0.Value) + Convert.ToDecimal(dr.Cells["Subtotal"].Value), 2)) * (decimal)dr.Cells["Jornadas"].Value;
					
						//this.txtDescIva0Total.Value = (Math.Round(Convert.ToDecimal(this.txtDescIva0Total.Value) + Convert.ToDecimal(dr.Cells["DescuentoArt"].Value), 2));
					}
				}				

				this.txtBaseImponible.Value = Math.Round(Convert.ToDecimal(this.txtIva.Value) - Convert.ToDecimal(this.txtDescIvaTotal.Value), 2);

				this.txtIvaTotal.Value = Math.Round(((Convert.ToDecimal(this.txtIva.Value) - Convert.ToDecimal(this.txtDescIvaTotal.Value)) * dPorcIva)/100.00m, 2);				

				this.txtTotal.Value = Math.Round(Convert.ToDecimal(this.txtIva0.Value), 2)
					+ Math.Round(Convert.ToDecimal(this.txtIva.Value), 2)
					- Math.Round(Convert.ToDecimal(this.txtDescIva0Total.Value), 2)
					- Math.Round(Convert.ToDecimal(this.txtDescIvaTotal.Value), 2)
					+ Math.Round(Convert.ToDecimal(this.txtIvaTotal.Value), 2);			
			} 
			#endregion Desplazamiento de grilla			
		}

		private void ultraGrid1_AfterCellUpdate(object sender, Infragistics.Win.UltraWinGrid.CellEventArgs e)
		{
			if (bEdicion)
			{
				if (bActualiza)
				{
					bActualiza = false;
					
					int VerificaCambio = 0;

					#region Codigo
//					if (e.Cell.Column.ToString() == "Codigo")
//					{
//						bActualiza = false;
//						FuncionesInventario.FacturacionCargaArticulos(sender, e, this.ultraGrid1, (int)this.cmbFormaPago.Value, idTipoPrecio, (int)this.cmbBodega.Value, 0, 0, 0, 0, Convert.ToDateTime(this.dtFecha.Value), 3, true, false, false, false, 0, 0, 0, cdsSeteoF);
//						bActualiza = true;
//						this.cmdTotal();						
//					} 
					#endregion Codigo

					#region Cantidad
					if (e.Cell.Column.ToString() == "Cantidad")
					{
						bActualiza = false;
						FuncionesInventario.ValidaCamposGridEnteros(sender, e, "Cantidad");

						if (!bNuevo)
							FuncionesInventario.ValidaSerialesIngresadasModificaCantidad(e, cdsSeteoF);
												
						VerificaCambio = 1;
						bActualiza = true;

						this.cmdTotal(); 
					}
					#endregion Cantidad

					#region Jornadas
					if (e.Cell.Column.ToString() == "Jornadas")
					{
						bActualiza = false;
						if (!FuncionesInventario.ValidaCamposGridDecimal(sender, e, "Jornadas"))
						{																		
							cmdTotal();
							bActualiza = true;
							return;
						}

						cmdTotal();
						bActualiza = true;
					}
					#endregion Jornadas
				
					#region Precio
					if (e.Cell.Column.ToString() == "Precio")
					{
						bActualiza = false;
						if (!FuncionesInventario.ValidaCamposGridDecimal(sender, e, "Precio"))
						{
							cmdTotal();
							bActualiza = true;
							return;
						}
	
						if (!Facturacion.bValidaModificacionPrecios(e, miAcceso.CambiarPrecio, 3, idTipoPrecio, 0, cdsSeteoF)) 
						{							
							cmdTotal();
							bActualiza = true;
							return;
						}

						cmdTotal();
						bActualiza = true;
					}
					#endregion Precio

					#region IVA
					if (e.Cell.Column.ToString() == "Impuesto")
					{
						bActualiza = false;
						if (!FuncionesInventario.ValidaCamposGridDecimal(sender, e, "Impuesto"))
						{							
							cmdTotal();
							bActualiza = true;
							return;
						}

						if (!Facturacion.bValidaImpuestos(e, (DateTime)this.dtFecha.Value, 1, cdsSeteoF))
						{							
							cmdTotal();
							bActualiza = true;
							return;
						}						
						
						cmdTotal();
						bActualiza = true;
					}
					#endregion IVA

					#region DescuentoPorc - DescuentoArt
					if (e.Cell.Column.ToString() == "DescuentoPorc")
					{
						bActualiza = false;
						if (!FuncionesInventario.ValidaCamposGridDecimal(sender, e, "DescuentoPorc"))	
						{
							cmdTotal();
							bActualiza = true;
							return;
						}

//						bActualiza = true;
//
//						bActualiza = false;
//						if (Convert.ToDecimal(this.txtDescIvaTotal.Value) > 0)
//						{							
//							cmdTotal();
//							bActualiza = true;
//							return;
//						}

						cmdTotal();
						bActualiza = true;						
					}

					if (e.Cell.Column.ToString() == "DescuentoArt")
					{						
						cmdTotal();
						bActualiza = true;
					}
					#endregion DescuentoPorc - DescuentoArt
				}

				bActualiza = true;
			}
		}

		private void ultraGrid1_AfterRowInsert(object sender, Infragistics.Win.UltraWinGrid.RowEventArgs e)
		{
			if (bEdicion)
				Facturacion.IniciaValoresColumnas(sender, e, this.ultraGrid1, 3, (int)this.cmbBodega.Value);			
		}

		private void ultraGrid1_BeforeRowsDeleted(object sender, Infragistics.Win.UltraWinGrid.BeforeRowsDeletedEventArgs e)
		{
			if (bEdicion)
			{
				if (!bEliminaAlValidar)
				{
					if (!bNuevo && !FuncionesInventario.ValidaConteoSeriales(this.ultraGrid1, cdsSeteoF))
					{
						e.Cancel = true;
						return;
					}					

					if (!FuncionesInventario.EliminaFilasGrids(this.btnCancelar, sender, e, cdsSeteoF)) e.Cancel = true;
				}
				else e.DisplayPromptMsg = false;
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
			
			this.txtDescIvaTotal.MaskInput = stInput;
			this.txtIvaTotal.MaskInput = stInput;
			this.txtIva.MaskInput = stInput;
			this.txtIva0.MaskInput = stInput;
			this.txtTotal.MaskInput = stInput;		
			#endregion Formato numerico			
			
			FuncionesProcedimientos.FormatoGrid(e, "Jornadas", 2);
			FuncionesProcedimientos.FormatoGrid(e, "Precio", 2);						
			FuncionesProcedimientos.FormatoGrid(e, "Impuesto", 2);						
			FuncionesProcedimientos.FormatoGrid(e, "DescuentoPorc", 2);
		}

		private void ultraGrid1_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.F3)
			{
				if (this.ultraGrid1.ActiveCell.Column.Key.ToString().ToUpper() == "CODIGO" && bEdicion)
				{
					if (!FuncionesInventario.ValidaConteoSeriales(this.ultraGrid1, cdsSeteoF)) return;

					using (frmInventarioZoom Busqueda = new frmInventarioZoom(iNumero))
					{
						if (DialogResult.OK == Busqueda.ShowDialog())
						{				
							int iFila = this.ultraGrid1.ActiveRow.Index;
							foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in Busqueda.ultraGrid1.Rows.All)
							{
								if ((bool)dr.Cells["Sel"].Value)
								{
									#region Asigna valores en el grid
									ultraGrid1.Rows[iFila].Cells["idArticulo"].Value = (int)dr.Cells["idArticulo"].Value;
									ultraGrid1.Rows[iFila].Cells["Codigo"].Value = dr.Cells["Codigo"].Value.ToString();
									ultraGrid1.Rows[iFila].Cells["Articulo"].Value = dr.Cells["Articulo"].Value.ToString();
									ultraGrid1.Rows[iFila].Cells["idUnidad"].Value = 1;
									ultraGrid1.Rows[iFila].Cells["Cantidad"].Value = 1;
									ultraGrid1.Rows[iFila].Cells["Jornadas"].Value = Convert.ToDecimal(this.txtJornadas.Value);
									ultraGrid1.Rows[iFila].Cells["Precio"].Value = Convert.ToDecimal(dr.Cells["Precio3"].Value);
									ultraGrid1.Rows[iFila].Cells["Impuesto"].Value = Facturacion.dPorcentajeIVAVigente(true, false, 0, (DateTime)this.dtFecha.Value, cdsSeteoF);
									ultraGrid1.Rows[iFila].Cells["DescuentoArt"].Value = 0.00m;
									ultraGrid1.Rows[iFila].Cells["DescuentoPorc"].Value = 0.00m;
									ultraGrid1.Rows[iFila].Cells["Principal"].Value = false;
									ultraGrid1.Rows[iFila].Cells["idCombo"].Value = 0;
									ultraGrid1.Rows[iFila].Cells["Tipo"].Value = 0;									

									iFila = iFila + 1;

									ultraGrid1.DisplayLayout.Bands[0].AddNew();									
									#endregion Asigna valores en el grid								
								}
							}							

							Funcion.EjecutaSQL(cdsSeteoF, string.Format("Update TemporalArticulos Set Sel = 0 Where Numero = {0}", iNumero));
						}
					}
					//FuncionesInventario.AbreBusquedaArticulos(this.ultraGrid1, (int)this.cmbBodega.Value, 3, (int)this.cmbFormaPago.Value, Convert.ToDateTime(this.dtFecha.Value));
				}				
			}

			FuncionesProcedimientos.DesplazamientoGrids(sender, e, this.ultraGrid1);	
		}

		private void ultraGrid1_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
			if (ultraGrid1.ActiveCell.Column.Key.ToString().ToUpper() == "CANTIDAD")
				Funcion.CamposNumericos(sender, e);
		}

		private void frmCotizaciones_Closing(object sender, System.ComponentModel.CancelEventArgs e)
		{
			if (bEdicion)
			{
				MessageBox.Show("No puede salir mientras edita un Pedido", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				e.Cancel = true;
				return;
			}
		}
		
		private void txtDescIvaTotal_Validated(object sender, System.EventArgs e)
		{
			if (bEdicion) this.cmdTotal();				
		}		

		private void ultraGrid1_AfterRowsDeleted(object sender, System.EventArgs e)
		{
			this.cmdTotal();
		}

		private void txtDescIvaTotal_Click(object sender, System.EventArgs e)
		{
			this.txtDescIvaTotal.SelectAll();
		}

		private void txtDescIvaTotal_Enter(object sender, System.EventArgs e)
		{
			this.txtDescIvaTotal.SelectAll();
		}

		private void txtDescIvaTotal_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Enter) this.btnGuardar.Focus();
		}

		private void btnEMail_Click(object sender, System.EventArgs e)
		{
			using (frmListaCorreosCotizaciones Lista = new frmListaCorreosCotizaciones())
			{
				if (DialogResult.OK == Lista.ShowDialog())
				{	
					string stFiltro = "{Compra.idCompra} = " + this.txtIdCompra.Value.ToString();

					using (Reporte Reporte = new Reporte("OrdenCompra.Rpt", stFiltro))
					{
						Reporte.ExportaPDF(@"\Cotizaciones\Cotizacion_" + this.txtNumero.Text.ToString() + ".pdf");					
					
						string sArchivo = this.txtNumero.Text.ToString() + ".pdf";
					
						string sSQLMail = string.Format(@"Exec MailCotizaciones 'Cotizacion_{0}', {1}, '{2}'", 
							sArchivo, (int)this.txtIdCompra.Value, Lista.txtDireccion.Text.ToString().Trim());
						string sMensaje = Funcion.sEscalarSQL(cdsSeteoF, sSQLMail, true);

						MessageBox.Show(sMensaje, "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Information);
					}
					/*
					if (!FuncionesProcedimientos.bExportaPDF_PA("CotizacionZOOM.rpt", (int)this.txtIdCompra.Value, this.txtNumero.Text.ToString().Trim() + ".pdf")) return;
					string sMensaje = string.Format("Sr./Srta./Sra. \nPresente.-\n\n\nAdjunto Archivo con la Cotización N. {0}\n\n\n\n\n\nAtte.\n\n\nASOCIACION DE CUENTAS DE PARTICIPACION CROMATIKA", 
						this.txtNumero.Text.ToString().Trim());
					FuncionesProcedimientos.Envio_Mail(Lista.txtDireccion.Text.ToString().Trim(), string.Format("Cotización N.: {0}", this.txtNumero.Text.ToString().Trim()), this.txtNumero.Text.ToString().Trim() + ".pdf", sMensaje, cdsSeteoF); 							
					*/
				}
			}
		}

		private void btnProyecto_Click(object sender, System.EventArgs e)
		{
			using (frmCreaProyectoZOOM Proyecto = new frmCreaProyectoZOOM())
			{
				if (DialogResult.OK == Proyecto.ShowDialog())
				{
					this.cmbProyecto.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "SELECT idProyectoZoom, Numero, Nombre FROM ProyectoZoom Order By Numero");
				}
			}
		}

		private void ultraGrid1_BeforeCellUpdate(object sender, Infragistics.Win.UltraWinGrid.BeforeCellUpdateEventArgs e)
		{
			
		}

		private void cmbProyecto_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Enter) this.txtJornadas.Focus();
		}

		private void txtRuc_ValueChanged(object sender, System.EventArgs e)
		{
		
		}

		private void txtJornadas_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Enter) 
			{
				this.ultraGrid1.DisplayLayout.Bands[0].AddNew();
				this.ultraGrid1.ActiveCell = this.ultraGrid1.Rows[0].Cells["Codigo"];
			}
		}

		private void txtJornadas_Enter(object sender, System.EventArgs e)
		{
			this.txtJornadas.SelectAll();
		}

		private void txtJornadas_Click(object sender, System.EventArgs e)
		{
			this.txtJornadas.SelectAll();
		}

		private void btnAplicarDescuento_Click(object sender, System.EventArgs e)
		{
			if (this.ultraGrid1.Rows.Count == 0)
			{
				MessageBox.Show("No hay items para aplicar descuentos", "Validacion", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				return;
			}

			if (this.txtDescuento.Value == System.DBNull.Value)
			{
				MessageBox.Show("El descuento a aplicar no puede estar vacio", "Validacion", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				this.txtDescuento.Focus();
				return;
			}

			foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.ultraGrid1.Rows.All)
			{
				dr.Cells["DescuentoPorc"].Value = Convert.ToDecimal(this.txtDescuento.Value);
			}

			this.cmdTotal();
		}

		private void txtDescuento_Validated(object sender, System.EventArgs e)
		{
			if (bEdicion)
				if (this.txtDescuento.Value == System.DBNull.Value) 
					this.txtDescuento.Value = 0.00m;
		}

		private void txtJornadas_Validated(object sender, System.EventArgs e)
		{
			if (bEdicion)
				if (this.txtJornadas.Value == System.DBNull.Value) 
					this.txtJornadas.Value = 1;
		}

		private void txtJornadas_ValueChanged(object sender, System.EventArgs e)
		{
			if (bEdicion) this.cmdTotal();
		}

		private void txtDescuento_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Enter) this.btnAplicarDescuento_Click(sender, e);
		}

		private void dtFecha_BeforeDropDown(object sender, System.ComponentModel.CancelEventArgs e)
		{
		
		}

		private void btnJornadas_Click(object sender, System.EventArgs e)
		{
			if (this.ultraGrid1.Rows.Count == 0)
			{
				MessageBox.Show("No hay items para asignar jornadas", "Validacion", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				return;
			}

			if (this.txtJornadas.Value == System.DBNull.Value)
			{
				MessageBox.Show("El numero de jornadas no puede estar vacio", "Validacion", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				this.txtJornadas.Focus();
				return;
			}

			foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.ultraGrid1.Rows.All)
			{
				dr.Cells["Jornadas"].Value = Convert.ToDecimal(this.txtJornadas.Value);
			}

			this.cmdTotal();
		}

		private void btnDuplicar_Click(object sender, System.EventArgs e)
		{	
			this.txtIdCompra.Value = 0;	
			this.dtFecha.Value = DateTime.Today;
			this.dtFechaEntrega.CalendarInfo.MinDate = DateTime.Today;			
			this.txtJornadas.Value = 1;
      
			#region Habilita controles
			this.cmbVendedor.Enabled = true;
			this.cmbTipoRuc.Enabled = true;
			this.cmbFormaPago.Enabled = true;
			this.txtRuc.Enabled = true;
			if (miAcceso.CambiarFecha) this.dtFecha.Enabled = true;
			this.cmbProyecto.Enabled = true;
			this.txtJornadas.Enabled = true;
			this.txtDescuento.Enabled = true;
			this.txtEstado.Value = 1;

			foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.ultraGrid1.Rows.All)
				dr.Cells["idDetCompra"].Value = 0;

			FuncionesProcedimientos.EstadoGrids(true, this.ultraGrid1);
												
			this.chkEntregado.Enabled = true;
			this.txtNotas.Enabled = true;

			this.btnNuevo.Enabled = false;
			this.btnConsultar.Enabled = false;
			this.btnEditar.Enabled = false;
			this.btnImprimir.Enabled = false;
			this.btnGuardar.Enabled = true;
			this.btnCancelar.Enabled = true;
			this.btnAnular.Enabled = false;
			this.btnCliente.Enabled = true;
			this.btnJornadas.Enabled = true;
			this.btnAplicarDescuento.Enabled = true;		
			this.btnDuplicar.Enabled = false;
			this.btnEMail.Enabled = false;
			#endregion Habilita controles					

			this.cmbVendedor.Focus();

			GeneraInventarioTemporal();

			bNuevo = true;
			bEdicion = true;
		}

		private void cmbBodega_InitializeLayout(object sender, Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs e)
		{
		
		}
	}
}
