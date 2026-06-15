using System;
using System.Globalization;
using System.Data;
using System.Data.SqlClient;
using C1.Data;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using Infragistics.Win.UltraWinGrid;
using System.Web.Mail;
using CrystalDecisions.Shared;
using System.Runtime.InteropServices;
using System.Diagnostics;
using System.IO;
using CrystalDecisions;
using CrystalDecisions.CrystalReports.Engine;
using CrystalDecisions.ReportSource;


namespace Latinium
{
	/// <summary>
	/// Descripción breve de frmPrefacturaRemate.
	/// </summary>
	public class frmPrefacturaRemate : System.Windows.Forms.Form
	{
		private Infragistics.Win.UltraWinEditors.UltraCheckEditor chkAutorizar;
		private System.Windows.Forms.Label label2;
		private System.Windows.Forms.Label label63;
		private System.Windows.Forms.Label label64;
		private System.Windows.Forms.Label label65;
		private System.Windows.Forms.Label label23;
		private System.Windows.Forms.Label lblIdentificacion;
		private System.Windows.Forms.Label lblCliente;
		private System.Windows.Forms.Label lblBodega;
		private System.Windows.Forms.Label lblVendedor;
		private System.Windows.Forms.Label label6;    
		private System.Windows.Forms.Label label21;
		private System.Windows.Forms.Label label61;
		private Infragistics.Win.UltraWinEditors.UltraCheckEditor chkVerificar;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtNumero;
		private System.Windows.Forms.Label lblEstado;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtCuotaAsignada;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtCuotaDisponible;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtCuotaUsada;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtValorCuotas;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtRuc;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbBodega;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtNombre;
		private Infragistics.Win.UltraWinSchedule.UltraCalendarCombo dtFecha;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbVendedor;
		private System.Windows.Forms.ComboBox cboFechaPrimerPago;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtNumCuotas;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtCuotaDeEntrada;
		private Infragistics.Win.UltraWinTabControl.UltraTabControl utcTransaccion;
		private Infragistics.Win.UltraWinTabControl.UltraTabSharedControlsPage ultraTabSharedControlsPage1;
		private Infragistics.Win.UltraWinTabControl.UltraTabPageControl ultraTabPageControl1;
		private Infragistics.Win.UltraWinTabControl.UltraTabPageControl ultraTabPageControl9;
		private Infragistics.Win.Misc.UltraButton ultraButton1;
		private Infragistics.Win.UltraWinGrid.UltraGrid grdSerial;
		private System.Windows.Forms.Button btnClaveDescuento;
		private System.Windows.Forms.Button btnAprobacion;
		private System.Windows.Forms.Button btnNuevo;
		private System.Windows.Forms.Button btnConsultar;
		private System.Windows.Forms.Button btnEditar;
		private System.Windows.Forms.Button btnImprimir;
		private System.Windows.Forms.Button btnGuardar;
		private System.Windows.Forms.Button btnAnular;
		private System.Windows.Forms.Button btnCancelar;
		private System.Windows.Forms.Button btnSalir;
		private System.Windows.Forms.Label label1;
		private System.Windows.Forms.Label lblIva;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtTotal;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtIvaTotal;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtDescIvaTotal;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtIva;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtidEntidadFinanciera;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtInteres;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtTotalCuotas;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtidTipoRuc;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtIdFormaPago;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtNumIdProyecto;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtIdComprobante;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtNumEstado;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtIdTipoFactura;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtBodegaPredef;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtIdSucursal;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtNumContadoCredito;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtNumIdCompra;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtNumIdCliente;
		private System.Windows.Forms.Label lblComprobante;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtComprobante;
		private System.Windows.Forms.Label label5;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtObservaciones;
		private System.Windows.Forms.Label label4;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtNotas;
		private System.Windows.Forms.Label label3;
		private System.Windows.Forms.Label label15;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbBodega1;
		private C1.Data.C1DataSet cdsSeteoF;
		private System.Windows.Forms.ContextMenu MnuAprobacion;
		private System.Windows.Forms.MenuItem mnuAprobacionAprobar;
		private System.Windows.Forms.MenuItem mnuAprobacionRechazar;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource2;
		private Infragistics.Win.UltraWinGrid.UltraGrid ultraGrid2;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource1;
		private System.Windows.Forms.Label lblUsuario;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtIngresoSerial;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource4;
		private System.Windows.Forms.Label label7;
		private System.Windows.Forms.GroupBox groupBox1;
		private System.Windows.Forms.Label label8;
		private System.Windows.Forms.Label label9;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtDescIva0Total;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtIva0;
		private System.Windows.Forms.Label lblIva0;
		private System.Windows.Forms.Label label55;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtSerial;
		private System.Windows.Forms.Label label10;
		/// <summary>
		/// Variable del diseñador requerida.
		/// </summary>
		private System.ComponentModel.Container components = null;

		public frmPrefacturaRemate()
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
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn1 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idSerial");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn2 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idDetalleSerial");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn3 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idArticulo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn4 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Serial");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn5 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Observaciones");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn6 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Estado");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn7 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idOrigen");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn8 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Posicion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn9 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idDetCompra");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn10 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Consignacion");
			Infragistics.Win.Appearance appearance2 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance3 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance4 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance5 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance6 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance7 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn1 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idSerial");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn2 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idDetalleSerial");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn3 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idArticulo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn4 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Serial");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn5 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Observaciones");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn6 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Estado");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn7 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idOrigen");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn8 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Posicion");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn9 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idDetCompra");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn10 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Consignacion");
			Infragistics.Win.Appearance appearance8 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance9 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance10 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance11 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance12 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance13 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance14 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance15 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance16 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance17 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance18 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand2 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn11 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idDetCompra");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn12 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idOrigen");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn13 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCompra");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn14 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idArticulo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn15 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idUnidad");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn16 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Codigo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn17 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Articulo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn18 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Cantidad");
			Infragistics.Win.Appearance appearance19 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn19 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Precio");
			Infragistics.Win.Appearance appearance20 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn20 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Impuesto");
			Infragistics.Win.Appearance appearance21 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn21 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("DescuentoArt");
			Infragistics.Win.Appearance appearance22 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn22 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("DescuentoPorc");
			Infragistics.Win.Appearance appearance23 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn23 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("SubTotal");
			Infragistics.Win.Appearance appearance24 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn24 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("EstadoSerial");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn25 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Tipo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn26 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Entregado");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn27 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Posicion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn28 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Principal");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn29 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Bloqueado");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn30 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCombo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn31 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idArticuloGE");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn32 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Jornadas");
			Infragistics.Win.Appearance appearance25 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance26 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance27 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance28 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance29 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance30 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance31 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn11 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idDetCompra");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn12 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idOrigen");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn13 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idCompra");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn14 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idArticulo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn15 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idUnidad");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn16 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Codigo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn17 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Articulo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn18 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Cantidad");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn19 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Precio");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn20 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Impuesto");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn21 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("DescuentoArt");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn22 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("DescuentoPorc");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn23 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("SubTotal");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn24 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("EstadoSerial");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn25 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Tipo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn26 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Entregado");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn27 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Posicion");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn28 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Principal");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn29 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Bloqueado");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn30 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idCombo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn31 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idArticuloGE");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn32 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Jornadas");
			Infragistics.Win.Appearance appearance32 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand3 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn33 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Bodega");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn34 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn35 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Codigo", 0);
			Infragistics.Win.Appearance appearance33 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance34 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton dateButton1 = new Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton();
			System.Resources.ResourceManager resources = new System.Resources.ResourceManager(typeof(frmPrefacturaRemate));
			Infragistics.Win.Appearance appearance35 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand4 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn36 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idPersonal");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn37 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn38 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Codigo");
			Infragistics.Win.Appearance appearance36 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance37 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance38 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinTabControl.UltraTab ultraTab1 = new Infragistics.Win.UltraWinTabControl.UltraTab();
			Infragistics.Win.UltraWinTabControl.UltraTab ultraTab2 = new Infragistics.Win.UltraWinTabControl.UltraTab();
			Infragistics.Win.Appearance appearance39 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance40 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance41 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance42 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance43 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance44 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand5 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn39 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Bodega");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn40 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn41 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Codigo", 0);
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn33 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idEntidadFinanciera");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn34 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Nombre");
			Infragistics.Win.Appearance appearance45 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance46 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance47 = new Infragistics.Win.Appearance();
			this.ultraTabPageControl9 = new Infragistics.Win.UltraWinTabControl.UltraTabPageControl();
			this.ultraButton1 = new Infragistics.Win.Misc.UltraButton();
			this.grdSerial = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.ultraDataSource4 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.txtSerial = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.label10 = new System.Windows.Forms.Label();
			this.ultraTabPageControl1 = new Infragistics.Win.UltraWinTabControl.UltraTabPageControl();
			this.txtNotas = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.label4 = new System.Windows.Forms.Label();
			this.txtObservaciones = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.label5 = new System.Windows.Forms.Label();
			this.chkAutorizar = new Infragistics.Win.UltraWinEditors.UltraCheckEditor();
			this.label2 = new System.Windows.Forms.Label();
			this.label63 = new System.Windows.Forms.Label();
			this.label64 = new System.Windows.Forms.Label();
			this.label65 = new System.Windows.Forms.Label();
			this.label23 = new System.Windows.Forms.Label();
			this.lblIdentificacion = new System.Windows.Forms.Label();
			this.lblCliente = new System.Windows.Forms.Label();
			this.lblBodega = new System.Windows.Forms.Label();
			this.lblVendedor = new System.Windows.Forms.Label();
			this.label6 = new System.Windows.Forms.Label();
			this.label21 = new System.Windows.Forms.Label();
			this.label61 = new System.Windows.Forms.Label();
			this.chkVerificar = new Infragistics.Win.UltraWinEditors.UltraCheckEditor();
			this.txtNumero = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.lblEstado = new System.Windows.Forms.Label();
			this.txtCuotaAsignada = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtCuotaDisponible = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtCuotaUsada = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtValorCuotas = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.ultraGrid2 = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.ultraDataSource2 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.txtRuc = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.cmbBodega = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.txtNombre = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.dtFecha = new Infragistics.Win.UltraWinSchedule.UltraCalendarCombo();
			this.cmbVendedor = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.cboFechaPrimerPago = new System.Windows.Forms.ComboBox();
			this.txtNumCuotas = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtCuotaDeEntrada = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.utcTransaccion = new Infragistics.Win.UltraWinTabControl.UltraTabControl();
			this.ultraTabSharedControlsPage1 = new Infragistics.Win.UltraWinTabControl.UltraTabSharedControlsPage();
			this.btnClaveDescuento = new System.Windows.Forms.Button();
			this.btnAprobacion = new System.Windows.Forms.Button();
			this.btnNuevo = new System.Windows.Forms.Button();
			this.btnConsultar = new System.Windows.Forms.Button();
			this.btnEditar = new System.Windows.Forms.Button();
			this.btnImprimir = new System.Windows.Forms.Button();
			this.btnGuardar = new System.Windows.Forms.Button();
			this.btnAnular = new System.Windows.Forms.Button();
			this.btnCancelar = new System.Windows.Forms.Button();
			this.btnSalir = new System.Windows.Forms.Button();
			this.label1 = new System.Windows.Forms.Label();
			this.lblIva = new System.Windows.Forms.Label();
			this.txtTotal = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtIvaTotal = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtDescIvaTotal = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtIva = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtidEntidadFinanciera = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtInteres = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtTotalCuotas = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtidTipoRuc = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtIdFormaPago = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtNumIdProyecto = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtIdComprobante = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtNumEstado = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtIdTipoFactura = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtBodegaPredef = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtIdSucursal = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtNumContadoCredito = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtNumIdCompra = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtNumIdCliente = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.lblComprobante = new System.Windows.Forms.Label();
			this.txtComprobante = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.label3 = new System.Windows.Forms.Label();
			this.label15 = new System.Windows.Forms.Label();
			this.cmbBodega1 = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.cdsSeteoF = new C1.Data.C1DataSet();
			this.MnuAprobacion = new System.Windows.Forms.ContextMenu();
			this.mnuAprobacionAprobar = new System.Windows.Forms.MenuItem();
			this.mnuAprobacionRechazar = new System.Windows.Forms.MenuItem();
			this.ultraDataSource1 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.lblUsuario = new System.Windows.Forms.Label();
			this.txtIngresoSerial = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.label7 = new System.Windows.Forms.Label();
			this.groupBox1 = new System.Windows.Forms.GroupBox();
			this.label8 = new System.Windows.Forms.Label();
			this.label9 = new System.Windows.Forms.Label();
			this.txtDescIva0Total = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtIva0 = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.lblIva0 = new System.Windows.Forms.Label();
			this.label55 = new System.Windows.Forms.Label();
			this.ultraTabPageControl9.SuspendLayout();
			((System.ComponentModel.ISupportInitialize)(this.grdSerial)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource4)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtSerial)).BeginInit();
			this.ultraTabPageControl1.SuspendLayout();
			((System.ComponentModel.ISupportInitialize)(this.txtNotas)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtObservaciones)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNumero)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtCuotaAsignada)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtCuotaDisponible)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtCuotaUsada)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtValorCuotas)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraGrid2)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource2)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtRuc)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbBodega)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNombre)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.dtFecha)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbVendedor)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNumCuotas)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtCuotaDeEntrada)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.utcTransaccion)).BeginInit();
			this.utcTransaccion.SuspendLayout();
			((System.ComponentModel.ISupportInitialize)(this.txtTotal)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIvaTotal)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtDescIvaTotal)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIva)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtidEntidadFinanciera)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtInteres)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtTotalCuotas)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtidTipoRuc)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIdFormaPago)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNumIdProyecto)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIdComprobante)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNumEstado)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIdTipoFactura)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtBodegaPredef)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIdSucursal)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNumContadoCredito)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNumIdCompra)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNumIdCliente)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtComprobante)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbBodega1)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource1)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIngresoSerial)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtDescIva0Total)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIva0)).BeginInit();
			this.SuspendLayout();
			// 
			// ultraTabPageControl9
			// 
			this.ultraTabPageControl9.Controls.Add(this.ultraButton1);
			this.ultraTabPageControl9.Controls.Add(this.grdSerial);
			this.ultraTabPageControl9.Controls.Add(this.txtSerial);
			this.ultraTabPageControl9.Controls.Add(this.label10);
			this.ultraTabPageControl9.Location = new System.Drawing.Point(1, 20);
			this.ultraTabPageControl9.Name = "ultraTabPageControl9";
			this.ultraTabPageControl9.Size = new System.Drawing.Size(694, 138);
			// 
			// ultraButton1
			// 
			this.ultraButton1.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.ultraButton1.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(247)), ((System.Byte)(245)), ((System.Byte)(241)));
			this.ultraButton1.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			this.ultraButton1.Font = new System.Drawing.Font("Tahoma", 8.25F);
			this.ultraButton1.ForeColor = System.Drawing.Color.Black;
			this.ultraButton1.Location = new System.Drawing.Point(760, 9);
			this.ultraButton1.Name = "ultraButton1";
			this.ultraButton1.Size = new System.Drawing.Size(24, 25);
			this.ultraButton1.TabIndex = 152;
			this.ultraButton1.Text = "&N";
			this.ultraButton1.Visible = false;
			// 
			// grdSerial
			// 
			this.grdSerial.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.grdSerial.Cursor = System.Windows.Forms.Cursors.Default;
			this.grdSerial.DataSource = this.ultraDataSource4;
			appearance1.BackColor = System.Drawing.Color.White;
			this.grdSerial.DisplayLayout.Appearance = appearance1;
			this.grdSerial.DisplayLayout.AutoFitColumns = true;
			ultraGridColumn1.Header.VisiblePosition = 0;
			ultraGridColumn1.Hidden = true;
			ultraGridColumn1.Width = 21;
			ultraGridColumn2.Header.VisiblePosition = 1;
			ultraGridColumn2.Hidden = true;
			ultraGridColumn2.Width = 36;
			ultraGridColumn3.Header.VisiblePosition = 2;
			ultraGridColumn3.Hidden = true;
			ultraGridColumn3.Width = 85;
			ultraGridColumn4.Case = Infragistics.Win.UltraWinGrid.Case.Upper;
			ultraGridColumn4.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn4.Header.VisiblePosition = 3;
			ultraGridColumn4.Width = 236;
			ultraGridColumn5.Case = Infragistics.Win.UltraWinGrid.Case.Upper;
			ultraGridColumn5.Header.VisiblePosition = 4;
			ultraGridColumn5.Width = 423;
			ultraGridColumn6.Header.VisiblePosition = 5;
			ultraGridColumn6.Hidden = true;
			ultraGridColumn6.Width = 34;
			ultraGridColumn7.Header.VisiblePosition = 6;
			ultraGridColumn7.Hidden = true;
			ultraGridColumn8.Header.VisiblePosition = 7;
			ultraGridColumn8.Hidden = true;
			ultraGridColumn8.Width = 55;
			ultraGridColumn9.Header.VisiblePosition = 8;
			ultraGridColumn9.Hidden = true;
			ultraGridColumn9.Width = 80;
			ultraGridColumn10.Header.VisiblePosition = 9;
			ultraGridColumn10.Hidden = true;
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
			ultraGridBand1.Override.CellClickAction = Infragistics.Win.UltraWinGrid.CellClickAction.Edit;
			this.grdSerial.DisplayLayout.BandsSerializer.Add(ultraGridBand1);
			appearance2.ForeColor = System.Drawing.Color.Black;
			appearance2.ForeColorDisabled = System.Drawing.Color.Black;
			this.grdSerial.DisplayLayout.Override.ActiveRowAppearance = appearance2;
			this.grdSerial.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.False;
			this.grdSerial.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.True;
			appearance3.BackColor = System.Drawing.Color.Transparent;
			this.grdSerial.DisplayLayout.Override.CardAreaAppearance = appearance3;
			this.grdSerial.DisplayLayout.Override.CellClickAction = Infragistics.Win.UltraWinGrid.CellClickAction.Edit;
			appearance4.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance4.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance4.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance4.FontData.BoldAsString = "True";
			appearance4.FontData.Name = "Arial";
			appearance4.FontData.SizeInPoints = 9F;
			appearance4.ForeColor = System.Drawing.Color.White;
			appearance4.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.grdSerial.DisplayLayout.Override.HeaderAppearance = appearance4;
			appearance5.BackColor = System.Drawing.SystemColors.ControlLight;
			appearance5.BackColor2 = System.Drawing.Color.Gainsboro;
			appearance5.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.grdSerial.DisplayLayout.Override.RowAlternateAppearance = appearance5;
			appearance6.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance6.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance6.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.grdSerial.DisplayLayout.Override.RowSelectorAppearance = appearance6;
			appearance7.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance7.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance7.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.grdSerial.DisplayLayout.Override.SelectedRowAppearance = appearance7;
			this.grdSerial.Enabled = false;
			this.grdSerial.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.grdSerial.Location = new System.Drawing.Point(8, 40);
			this.grdSerial.Name = "grdSerial";
			this.grdSerial.Size = new System.Drawing.Size(680, 96);
			this.grdSerial.TabIndex = 17;
			this.grdSerial.BeforeRowsDeleted += new Infragistics.Win.UltraWinGrid.BeforeRowsDeletedEventHandler(this.grdSerial_BeforeRowsDeleted);
			this.grdSerial.KeyDown += new System.Windows.Forms.KeyEventHandler(this.grdSerial_KeyDown);
			// 
			// ultraDataSource4
			// 
			ultraDataColumn1.DataType = typeof(int);
			ultraDataColumn2.DataType = typeof(int);
			ultraDataColumn3.DataType = typeof(int);
			ultraDataColumn6.DataType = typeof(int);
			ultraDataColumn7.DataType = typeof(int);
			ultraDataColumn8.DataType = typeof(int);
			ultraDataColumn9.DataType = typeof(int);
			ultraDataColumn10.DataType = typeof(bool);
			this.ultraDataSource4.Band.Columns.AddRange(new object[] {
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
			// txtSerial
			// 
			appearance8.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtSerial.Appearance = appearance8;
			this.txtSerial.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtSerial.Enabled = false;
			this.txtSerial.Location = new System.Drawing.Point(64, 8);
			this.txtSerial.Name = "txtSerial";
			this.txtSerial.Size = new System.Drawing.Size(264, 22);
			this.txtSerial.TabIndex = 437;
			// 
			// label10
			// 
			this.label10.AutoSize = true;
			this.label10.BackColor = System.Drawing.Color.Transparent;
			this.label10.Location = new System.Drawing.Point(8, 11);
			this.label10.Name = "label10";
			this.label10.Size = new System.Drawing.Size(43, 17);
			this.label10.TabIndex = 437;
			this.label10.Text = "Seriales";
			// 
			// ultraTabPageControl1
			// 
			this.ultraTabPageControl1.Controls.Add(this.txtNotas);
			this.ultraTabPageControl1.Controls.Add(this.label4);
			this.ultraTabPageControl1.Controls.Add(this.txtObservaciones);
			this.ultraTabPageControl1.Controls.Add(this.label5);
			this.ultraTabPageControl1.Location = new System.Drawing.Point(-10000, -10000);
			this.ultraTabPageControl1.Name = "ultraTabPageControl1";
			this.ultraTabPageControl1.Size = new System.Drawing.Size(694, 138);
			// 
			// txtNotas
			// 
			this.txtNotas.AcceptsReturn = true;
			this.txtNotas.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			appearance9.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtNotas.Appearance = appearance9;
			this.txtNotas.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtNotas.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtNotas.Enabled = false;
			this.txtNotas.Location = new System.Drawing.Point(8, 31);
			this.txtNotas.Multiline = true;
			this.txtNotas.Name = "txtNotas";
			this.txtNotas.Size = new System.Drawing.Size(320, 84);
			this.txtNotas.TabIndex = 308;
			// 
			// label4
			// 
			this.label4.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.label4.AutoSize = true;
			this.label4.BackColor = System.Drawing.Color.Transparent;
			this.label4.Location = new System.Drawing.Point(8, 7);
			this.label4.Name = "label4";
			this.label4.Size = new System.Drawing.Size(33, 17);
			this.label4.TabIndex = 307;
			this.label4.Text = "Notas";
			// 
			// txtObservaciones
			// 
			this.txtObservaciones.AcceptsReturn = true;
			this.txtObservaciones.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			appearance10.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtObservaciones.Appearance = appearance10;
			this.txtObservaciones.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtObservaciones.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtObservaciones.Enabled = false;
			this.txtObservaciones.Location = new System.Drawing.Point(336, 31);
			this.txtObservaciones.Multiline = true;
			this.txtObservaciones.Name = "txtObservaciones";
			this.txtObservaciones.Size = new System.Drawing.Size(320, 84);
			this.txtObservaciones.TabIndex = 306;
			// 
			// label5
			// 
			this.label5.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.label5.AutoSize = true;
			this.label5.BackColor = System.Drawing.Color.Transparent;
			this.label5.Location = new System.Drawing.Point(336, 7);
			this.label5.Name = "label5";
			this.label5.Size = new System.Drawing.Size(76, 17);
			this.label5.TabIndex = 305;
			this.label5.Text = "Observaciones";
			// 
			// chkAutorizar
			// 
			appearance11.ForeColorDisabled = System.Drawing.Color.Black;
			this.chkAutorizar.Appearance = appearance11;
			this.chkAutorizar.BackColor = System.Drawing.Color.Transparent;
			this.chkAutorizar.Enabled = false;
			this.chkAutorizar.Location = new System.Drawing.Point(920, 56);
			this.chkAutorizar.Name = "chkAutorizar";
			this.chkAutorizar.Size = new System.Drawing.Size(88, 22);
			this.chkAutorizar.TabIndex = 335;
			this.chkAutorizar.Text = "Aut. Entrada";
			this.chkAutorizar.Visible = false;
			// 
			// label2
			// 
			this.label2.AutoSize = true;
			this.label2.BackColor = System.Drawing.Color.Transparent;
			this.label2.Location = new System.Drawing.Point(320, 83);
			this.label2.Name = "label2";
			this.label2.Size = new System.Drawing.Size(44, 17);
			this.label2.TabIndex = 333;
			this.label2.Text = "Número";
			// 
			// label63
			// 
			this.label63.AutoSize = true;
			this.label63.BackColor = System.Drawing.Color.Transparent;
			this.label63.Location = new System.Drawing.Point(696, 35);
			this.label63.Name = "label63";
			this.label63.Size = new System.Drawing.Size(87, 17);
			this.label63.TabIndex = 327;
			this.label63.Text = "Cuota disponible";
			// 
			// label64
			// 
			this.label64.AutoSize = true;
			this.label64.BackColor = System.Drawing.Color.Transparent;
			this.label64.Location = new System.Drawing.Point(512, 59);
			this.label64.Name = "label64";
			this.label64.Size = new System.Drawing.Size(66, 17);
			this.label64.TabIndex = 326;
			this.label64.Text = "Cuota usada";
			// 
			// label65
			// 
			this.label65.AutoSize = true;
			this.label65.BackColor = System.Drawing.Color.Transparent;
			this.label65.Location = new System.Drawing.Point(512, 35);
			this.label65.Name = "label65";
			this.label65.Size = new System.Drawing.Size(81, 17);
			this.label65.TabIndex = 325;
			this.label65.Text = "Cuota asignada";
			// 
			// label23
			// 
			this.label23.AutoSize = true;
			this.label23.BackColor = System.Drawing.Color.Transparent;
			this.label23.Location = new System.Drawing.Point(696, 59);
			this.label23.Name = "label23";
			this.label23.Size = new System.Drawing.Size(62, 17);
			this.label23.TabIndex = 323;
			this.label23.Text = "Valor Cuota";
			// 
			// lblIdentificacion
			// 
			this.lblIdentificacion.AutoSize = true;
			this.lblIdentificacion.Location = new System.Drawing.Point(8, 35);
			this.lblIdentificacion.Name = "lblIdentificacion";
			this.lblIdentificacion.Size = new System.Drawing.Size(38, 17);
			this.lblIdentificacion.TabIndex = 322;
			this.lblIdentificacion.Text = "Cédula";
			// 
			// lblCliente
			// 
			this.lblCliente.AutoSize = true;
			this.lblCliente.Location = new System.Drawing.Point(176, 35);
			this.lblCliente.Name = "lblCliente";
			this.lblCliente.Size = new System.Drawing.Size(44, 17);
			this.lblCliente.TabIndex = 321;
			this.lblCliente.Text = "Nombre";
			// 
			// lblBodega
			// 
			this.lblBodega.AutoSize = true;
			this.lblBodega.Location = new System.Drawing.Point(176, 10);
			this.lblBodega.Name = "lblBodega";
			this.lblBodega.Size = new System.Drawing.Size(30, 17);
			this.lblBodega.TabIndex = 320;
			this.lblBodega.Text = "Local";
			this.lblBodega.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
			// 
			// lblVendedor
			// 
			this.lblVendedor.AutoSize = true;
			this.lblVendedor.Location = new System.Drawing.Point(688, 10);
			this.lblVendedor.Name = "lblVendedor";
			this.lblVendedor.Size = new System.Drawing.Size(52, 17);
			this.lblVendedor.TabIndex = 318;
			this.lblVendedor.Text = "Vendedor";
			// 
			// label6
			// 
			this.label6.AutoSize = true;
			this.label6.Location = new System.Drawing.Point(8, 10);
			this.label6.Name = "label6";
			this.label6.Size = new System.Drawing.Size(34, 17);
			this.label6.TabIndex = 316;
			this.label6.Text = "Fecha";
			// 
			// label21
			// 
			this.label21.AutoSize = true;
			this.label21.BackColor = System.Drawing.Color.Transparent;
			this.label21.Location = new System.Drawing.Point(8, 59);
			this.label21.Name = "label21";
			this.label21.Size = new System.Drawing.Size(38, 17);
			this.label21.TabIndex = 314;
			this.label21.Text = "Cuotas";
			// 
			// label61
			// 
			this.label61.AutoSize = true;
			this.label61.BackColor = System.Drawing.Color.Transparent;
			this.label61.Location = new System.Drawing.Point(176, 59);
			this.label61.Name = "label61";
			this.label61.Size = new System.Drawing.Size(43, 17);
			this.label61.TabIndex = 315;
			this.label61.Text = "Entrada";
			// 
			// chkVerificar
			// 
			appearance12.ForeColorDisabled = System.Drawing.Color.Black;
			this.chkVerificar.Appearance = appearance12;
			this.chkVerificar.BackColor = System.Drawing.Color.Transparent;
			this.chkVerificar.Enabled = false;
			this.chkVerificar.Location = new System.Drawing.Point(920, 32);
			this.chkVerificar.Name = "chkVerificar";
			this.chkVerificar.Size = new System.Drawing.Size(56, 22);
			this.chkVerificar.TabIndex = 334;
			this.chkVerificar.Text = "Carpa";
			// 
			// txtNumero
			// 
			appearance13.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtNumero.Appearance = appearance13;
			this.txtNumero.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtNumero.Enabled = false;
			this.txtNumero.Location = new System.Drawing.Point(384, 80);
			this.txtNumero.Name = "txtNumero";
			this.txtNumero.Size = new System.Drawing.Size(112, 22);
			this.txtNumero.TabIndex = 332;
			// 
			// lblEstado
			// 
			this.lblEstado.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			this.lblEstado.AutoSize = true;
			this.lblEstado.Font = new System.Drawing.Font("Microsoft Sans Serif", 13F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.lblEstado.ForeColor = System.Drawing.Color.Firebrick;
			this.lblEstado.Location = new System.Drawing.Point(816, 80);
			this.lblEstado.Name = "lblEstado";
			this.lblEstado.Size = new System.Drawing.Size(0, 23);
			this.lblEstado.TabIndex = 331;
			this.lblEstado.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
			// 
			// txtCuotaAsignada
			// 
			appearance14.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtCuotaAsignada.Appearance = appearance14;
			this.txtCuotaAsignada.BorderStyle = Infragistics.Win.UIElementBorderStyle.Solid;
			this.txtCuotaAsignada.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtCuotaAsignada.Enabled = false;
			this.txtCuotaAsignada.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.txtCuotaAsignada.FormatString = "#,##0.00";
			this.txtCuotaAsignada.Location = new System.Drawing.Point(604, 33);
			this.txtCuotaAsignada.Name = "txtCuotaAsignada";
			this.txtCuotaAsignada.NullText = "0.00";
			this.txtCuotaAsignada.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtCuotaAsignada.PromptChar = ' ';
			this.txtCuotaAsignada.Size = new System.Drawing.Size(80, 20);
			this.txtCuotaAsignada.TabIndex = 330;
			// 
			// txtCuotaDisponible
			// 
			appearance15.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtCuotaDisponible.Appearance = appearance15;
			this.txtCuotaDisponible.BorderStyle = Infragistics.Win.UIElementBorderStyle.Solid;
			this.txtCuotaDisponible.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtCuotaDisponible.Enabled = false;
			this.txtCuotaDisponible.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.txtCuotaDisponible.FormatString = "#,##0.00";
			this.txtCuotaDisponible.Location = new System.Drawing.Point(800, 33);
			this.txtCuotaDisponible.Name = "txtCuotaDisponible";
			this.txtCuotaDisponible.NullText = "0.00";
			this.txtCuotaDisponible.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtCuotaDisponible.PromptChar = ' ';
			this.txtCuotaDisponible.Size = new System.Drawing.Size(80, 20);
			this.txtCuotaDisponible.TabIndex = 329;
			// 
			// txtCuotaUsada
			// 
			appearance16.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtCuotaUsada.Appearance = appearance16;
			this.txtCuotaUsada.BorderStyle = Infragistics.Win.UIElementBorderStyle.Solid;
			this.txtCuotaUsada.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtCuotaUsada.Enabled = false;
			this.txtCuotaUsada.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.txtCuotaUsada.FormatString = "#,##0.00";
			this.txtCuotaUsada.Location = new System.Drawing.Point(604, 57);
			this.txtCuotaUsada.Name = "txtCuotaUsada";
			this.txtCuotaUsada.NullText = "0.00";
			this.txtCuotaUsada.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtCuotaUsada.PromptChar = ' ';
			this.txtCuotaUsada.Size = new System.Drawing.Size(80, 20);
			this.txtCuotaUsada.TabIndex = 328;
			// 
			// txtValorCuotas
			// 
			appearance17.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtValorCuotas.Appearance = appearance17;
			this.txtValorCuotas.BorderStyle = Infragistics.Win.UIElementBorderStyle.Solid;
			this.txtValorCuotas.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtValorCuotas.Enabled = false;
			this.txtValorCuotas.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.txtValorCuotas.FormatString = "#,##0.00";
			this.txtValorCuotas.Location = new System.Drawing.Point(800, 57);
			this.txtValorCuotas.Name = "txtValorCuotas";
			this.txtValorCuotas.NullText = "0.00";
			this.txtValorCuotas.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtValorCuotas.PromptChar = ' ';
			this.txtValorCuotas.Size = new System.Drawing.Size(80, 20);
			this.txtValorCuotas.TabIndex = 324;
			// 
			// ultraGrid2
			// 
			this.ultraGrid2.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.ultraGrid2.Cursor = System.Windows.Forms.Cursors.Default;
			this.ultraGrid2.DataSource = this.ultraDataSource2;
			appearance18.BackColor = System.Drawing.Color.White;
			this.ultraGrid2.DisplayLayout.Appearance = appearance18;
			this.ultraGrid2.DisplayLayout.AutoFitColumns = true;
			ultraGridBand2.AddButtonCaption = "_DetCompra";
			ultraGridColumn11.Header.VisiblePosition = 8;
			ultraGridColumn11.Hidden = true;
			ultraGridColumn11.Width = 57;
			ultraGridColumn12.Header.VisiblePosition = 12;
			ultraGridColumn12.Hidden = true;
			ultraGridColumn13.Header.VisiblePosition = 0;
			ultraGridColumn13.Hidden = true;
			ultraGridColumn13.Width = 57;
			ultraGridColumn14.Header.VisiblePosition = 1;
			ultraGridColumn14.Hidden = true;
			ultraGridColumn14.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.DropDownList;
			ultraGridColumn14.Width = 186;
			ultraGridColumn15.Header.Caption = "Unidad";
			ultraGridColumn15.Header.VisiblePosition = 4;
			ultraGridColumn15.Hidden = true;
			ultraGridColumn15.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.DropDownValidate;
			ultraGridColumn15.Width = 53;
			ultraGridColumn16.Case = Infragistics.Win.UltraWinGrid.Case.Upper;
			ultraGridColumn16.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn16.Header.VisiblePosition = 2;
			ultraGridColumn16.Width = 220;
			ultraGridColumn17.Case = Infragistics.Win.UltraWinGrid.Case.Upper;
			ultraGridColumn17.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn17.Header.VisiblePosition = 3;
			ultraGridColumn17.Width = 340;
			ultraGridColumn18.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance19.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn18.CellAppearance = appearance19;
			ultraGridColumn18.Format = "";
			ultraGridColumn18.Header.VisiblePosition = 5;
			ultraGridColumn18.MaskInput = "";
			ultraGridColumn18.NullText = "0";
			ultraGridColumn18.PromptChar = ' ';
			ultraGridColumn18.Width = 71;
			appearance20.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn19.CellAppearance = appearance20;
			ultraGridColumn19.Format = "#,##0.00";
			ultraGridColumn19.Header.VisiblePosition = 6;
			ultraGridColumn19.MaskInput = "{LOC}nn,nnn,nnn.nn";
			ultraGridColumn19.PromptChar = ' ';
			ultraGridColumn19.Width = 100;
			appearance21.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn20.CellAppearance = appearance21;
			ultraGridColumn20.Format = "#,##0.00";
			ultraGridColumn20.Header.Caption = "I.V.A.";
			ultraGridColumn20.Header.VisiblePosition = 7;
			ultraGridColumn20.MaskInput = "{LOC}nnnnnnn.nn";
			ultraGridColumn20.PromptChar = ' ';
			ultraGridColumn20.Width = 60;
			ultraGridColumn21.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance22.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn21.CellAppearance = appearance22;
			ultraGridColumn21.Header.Caption = "Descuento";
			ultraGridColumn21.Header.VisiblePosition = 9;
			ultraGridColumn21.Hidden = true;
			ultraGridColumn21.MaskInput = "{LOC}nn,nnn,nnn.nn";
			ultraGridColumn21.PromptChar = ' ';
			ultraGridColumn21.Width = 61;
			appearance23.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn22.CellAppearance = appearance23;
			ultraGridColumn22.Format = "#,##0.0000";
			ultraGridColumn22.Header.Caption = "% Desc.";
			ultraGridColumn22.Header.VisiblePosition = 10;
			ultraGridColumn22.MaskInput = "";
			ultraGridColumn22.PromptChar = ' ';
			ultraGridColumn22.Width = 80;
			ultraGridColumn23.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance24.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn23.CellAppearance = appearance24;
			ultraGridColumn23.Format = "#,##0.00";
			ultraGridColumn23.Header.VisiblePosition = 11;
			ultraGridColumn23.PromptChar = ' ';
			ultraGridColumn23.Width = 100;
			ultraGridColumn24.Header.VisiblePosition = 13;
			ultraGridColumn24.Hidden = true;
			ultraGridColumn25.Header.VisiblePosition = 14;
			ultraGridColumn25.Hidden = true;
			ultraGridColumn26.Header.VisiblePosition = 15;
			ultraGridColumn26.Hidden = true;
			ultraGridColumn27.Header.VisiblePosition = 16;
			ultraGridColumn27.Hidden = true;
			ultraGridColumn28.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn28.Header.VisiblePosition = 17;
			ultraGridColumn28.Hidden = true;
			ultraGridColumn29.Header.VisiblePosition = 18;
			ultraGridColumn29.Hidden = true;
			ultraGridColumn30.Header.VisiblePosition = 19;
			ultraGridColumn30.Hidden = true;
			ultraGridColumn31.Header.VisiblePosition = 20;
			ultraGridColumn31.Hidden = true;
			ultraGridColumn32.Header.VisiblePosition = 21;
			ultraGridColumn32.Hidden = true;
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
																										 ultraGridColumn32});
			ultraGridBand2.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.TemplateOnBottom;
			this.ultraGrid2.DisplayLayout.BandsSerializer.Add(ultraGridBand2);
			appearance25.ForeColor = System.Drawing.Color.Black;
			appearance25.ForeColorDisabled = System.Drawing.Color.Black;
			this.ultraGrid2.DisplayLayout.Override.ActiveRowAppearance = appearance25;
			this.ultraGrid2.DisplayLayout.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.TemplateOnBottom;
			this.ultraGrid2.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.True;
			this.ultraGrid2.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.True;
			appearance26.BackColor = System.Drawing.Color.Transparent;
			this.ultraGrid2.DisplayLayout.Override.CardAreaAppearance = appearance26;
			appearance27.ForeColor = System.Drawing.Color.Black;
			appearance27.ForeColorDisabled = System.Drawing.Color.Black;
			this.ultraGrid2.DisplayLayout.Override.CellAppearance = appearance27;
			appearance28.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance28.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance28.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance28.FontData.BoldAsString = "True";
			appearance28.FontData.Name = "Arial";
			appearance28.FontData.SizeInPoints = 9F;
			appearance28.ForeColor = System.Drawing.Color.White;
			appearance28.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.ultraGrid2.DisplayLayout.Override.HeaderAppearance = appearance28;
			appearance29.BackColor = System.Drawing.SystemColors.ControlLight;
			appearance29.BackColor2 = System.Drawing.Color.Gainsboro;
			appearance29.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ultraGrid2.DisplayLayout.Override.RowAlternateAppearance = appearance29;
			appearance30.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance30.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance30.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ultraGrid2.DisplayLayout.Override.RowSelectorAppearance = appearance30;
			appearance31.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance31.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance31.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ultraGrid2.DisplayLayout.Override.SelectedRowAppearance = appearance31;
			this.ultraGrid2.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.ultraGrid2.Location = new System.Drawing.Point(8, 112);
			this.ultraGrid2.Name = "ultraGrid2";
			this.ultraGrid2.Size = new System.Drawing.Size(992, 144);
			this.ultraGrid2.TabIndex = 313;
			this.ultraGrid2.DoubleClick += new System.EventHandler(this.ultraGrid2_DoubleClick);
			this.ultraGrid2.AfterRowsDeleted += new System.EventHandler(this.ultraGrid1_AfterRowsDeleted);
			this.ultraGrid2.AfterRowInsert += new Infragistics.Win.UltraWinGrid.RowEventHandler(this.ultraGrid1_AfterRowInsert);
			this.ultraGrid2.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.ultraGrid1_KeyPress);
			this.ultraGrid2.BeforeCellUpdate += new Infragistics.Win.UltraWinGrid.BeforeCellUpdateEventHandler(this.ultraGrid1_BeforeCellUpdate);
			this.ultraGrid2.BeforeRowsDeleted += new Infragistics.Win.UltraWinGrid.BeforeRowsDeletedEventHandler(this.ultraGrid1_BeforeRowsDeleted);
			this.ultraGrid2.KeyDown += new System.Windows.Forms.KeyEventHandler(this.ultraGrid1_KeyDown);
			this.ultraGrid2.AfterCellUpdate += new Infragistics.Win.UltraWinGrid.CellEventHandler(this.ultraGrid1_AfterCellUpdate);
			this.ultraGrid2.InitializeLayout += new Infragistics.Win.UltraWinGrid.InitializeLayoutEventHandler(this.ultraGrid1_InitializeLayout);
			// 
			// ultraDataSource2
			// 
			ultraDataColumn11.DataType = typeof(int);
			ultraDataColumn11.DefaultValue = 0;
			ultraDataColumn12.DataType = typeof(int);
			ultraDataColumn12.DefaultValue = 0;
			ultraDataColumn13.DataType = typeof(int);
			ultraDataColumn13.DefaultValue = 0;
			ultraDataColumn14.DataType = typeof(int);
			ultraDataColumn14.DefaultValue = 0;
			ultraDataColumn15.DataType = typeof(int);
			ultraDataColumn15.DefaultValue = 0;
			ultraDataColumn18.DataType = typeof(int);
			ultraDataColumn18.DefaultValue = 0;
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
			ultraDataColumn24.DataType = typeof(bool);
			ultraDataColumn24.DefaultValue = false;
			ultraDataColumn25.DataType = typeof(int);
			ultraDataColumn25.DefaultValue = 0;
			ultraDataColumn26.DataType = typeof(bool);
			ultraDataColumn26.DefaultValue = false;
			ultraDataColumn27.DataType = typeof(int);
			ultraDataColumn27.DefaultValue = 0;
			ultraDataColumn28.DataType = typeof(bool);
			ultraDataColumn28.DefaultValue = false;
			ultraDataColumn29.DataType = typeof(bool);
			ultraDataColumn29.DefaultValue = false;
			ultraDataColumn30.DataType = typeof(int);
			ultraDataColumn30.DefaultValue = 0;
			ultraDataColumn31.DataType = typeof(int);
			ultraDataColumn31.DefaultValue = 0;
			ultraDataColumn32.DataType = typeof(System.Decimal);
			ultraDataColumn32.DefaultValue = new System.Decimal(new int[] {
																																			0,
																																			0,
																																			0,
																																			0});
			this.ultraDataSource2.Band.Columns.AddRange(new object[] {
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
																																 ultraDataColumn32});
			// 
			// txtRuc
			// 
			appearance32.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtRuc.Appearance = appearance32;
			this.txtRuc.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtRuc.Enabled = false;
			this.txtRuc.Location = new System.Drawing.Point(56, 32);
			this.txtRuc.MaxLength = 10;
			this.txtRuc.Name = "txtRuc";
			this.txtRuc.Size = new System.Drawing.Size(112, 22);
			this.txtRuc.TabIndex = 309;
			this.txtRuc.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtRuc_KeyDown);
			this.txtRuc.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.txtRuc_KeyPress);
			this.txtRuc.Validating += new System.ComponentModel.CancelEventHandler(this.txtRuc_Validating);
			this.txtRuc.Validated += new System.EventHandler(this.txtRuc_Validated);
			this.txtRuc.ValueChanged += new System.EventHandler(this.txtRuc_ValueChanged);
			// 
			// cmbBodega
			// 
			this.cmbBodega.CausesValidation = false;
			this.cmbBodega.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbBodega.Cursor = System.Windows.Forms.Cursors.Default;
			ultraGridColumn33.Header.VisiblePosition = 2;
			ultraGridColumn34.Header.VisiblePosition = 1;
			ultraGridColumn34.Width = 180;
			ultraGridColumn35.Header.VisiblePosition = 0;
			ultraGridBand3.Columns.AddRange(new object[] {
																										 ultraGridColumn33,
																										 ultraGridColumn34,
																										 ultraGridColumn35});
			this.cmbBodega.DisplayLayout.BandsSerializer.Add(ultraGridBand3);
			this.cmbBodega.DisplayMember = "Nombre";
			this.cmbBodega.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbBodega.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbBodega.Enabled = false;
			this.cmbBodega.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbBodega.Location = new System.Drawing.Point(232, 7);
			this.cmbBodega.MaxDropDownItems = 30;
			this.cmbBodega.Name = "cmbBodega";
			this.cmbBodega.Size = new System.Drawing.Size(180, 22);
			this.cmbBodega.TabIndex = 307;
			this.cmbBodega.ValueMember = "Bodega";
			this.cmbBodega.KeyDown += new System.Windows.Forms.KeyEventHandler(this.cmbBodega_KeyDown);
			this.cmbBodega.Validating += new System.ComponentModel.CancelEventHandler(this.cmbBodega_Validating);
			this.cmbBodega.ValueChanged += new System.EventHandler(this.cmbBodega_ValueChanged);
			this.cmbBodega.InitializeLayout += new Infragistics.Win.UltraWinGrid.InitializeLayoutEventHandler(this.cmbBodega_InitializeLayout);
			// 
			// txtNombre
			// 
			appearance33.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtNombre.Appearance = appearance33;
			this.txtNombre.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtNombre.Enabled = false;
			this.txtNombre.Location = new System.Drawing.Point(232, 32);
			this.txtNombre.Name = "txtNombre";
			this.txtNombre.Size = new System.Drawing.Size(264, 22);
			this.txtNombre.TabIndex = 319;
			// 
			// dtFecha
			// 
			appearance34.ForeColorDisabled = System.Drawing.Color.Black;
			this.dtFecha.Appearance = appearance34;
			this.dtFecha.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			dateButton1.Caption = "Today";
			this.dtFecha.DateButtons.Add(dateButton1);
			this.dtFecha.Enabled = false;
			this.dtFecha.Format = "dd/MM/yyyy";
			this.dtFecha.Location = new System.Drawing.Point(56, 8);
			this.dtFecha.Name = "dtFecha";
			this.dtFecha.NonAutoSizeHeight = 23;
			this.dtFecha.Size = new System.Drawing.Size(104, 21);
			this.dtFecha.SpinButtonsVisible = true;
			this.dtFecha.TabIndex = 306;
			this.dtFecha.Value = ((object)(resources.GetObject("dtFecha.Value")));
			this.dtFecha.KeyDown += new System.Windows.Forms.KeyEventHandler(this.dtFecha_KeyDown);
			this.dtFecha.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.dtFecha_KeyPress);
			this.dtFecha.Validating += new System.ComponentModel.CancelEventHandler(this.dtFecha_Validating);
			this.dtFecha.Validated += new System.EventHandler(this.dtFecha_Validated);
			this.dtFecha.ValueChanged += new System.EventHandler(this.dtFecha_ValueChanged);
			// 
			// cmbVendedor
			// 
			appearance35.ForeColorDisabled = System.Drawing.Color.Black;
			this.cmbVendedor.Appearance = appearance35;
			this.cmbVendedor.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbVendedor.Cursor = System.Windows.Forms.Cursors.Default;
			ultraGridColumn36.Header.VisiblePosition = 0;
			ultraGridColumn36.Hidden = true;
			ultraGridColumn37.Header.VisiblePosition = 1;
			ultraGridColumn37.Width = 335;
			ultraGridColumn38.Header.VisiblePosition = 2;
			ultraGridBand4.Columns.AddRange(new object[] {
																										 ultraGridColumn36,
																										 ultraGridColumn37,
																										 ultraGridColumn38});
			this.cmbVendedor.DisplayLayout.BandsSerializer.Add(ultraGridBand4);
			this.cmbVendedor.DisplayMember = "Nombre";
			this.cmbVendedor.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbVendedor.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbVendedor.Enabled = false;
			this.cmbVendedor.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbVendedor.Location = new System.Drawing.Point(752, 7);
			this.cmbVendedor.Name = "cmbVendedor";
			this.cmbVendedor.Size = new System.Drawing.Size(250, 22);
			this.cmbVendedor.TabIndex = 308;
			this.cmbVendedor.ValueMember = "idPersonal";
			this.cmbVendedor.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.cmbVendedor_KeyPress);
			this.cmbVendedor.Validating += new System.ComponentModel.CancelEventHandler(this.cmbVendedor_Validating);
			// 
			// cboFechaPrimerPago
			// 
			this.cboFechaPrimerPago.Enabled = false;
			this.cboFechaPrimerPago.Location = new System.Drawing.Point(396, 57);
			this.cboFechaPrimerPago.MaxDropDownItems = 15;
			this.cboFechaPrimerPago.Name = "cboFechaPrimerPago";
			this.cboFechaPrimerPago.Size = new System.Drawing.Size(100, 21);
			this.cboFechaPrimerPago.TabIndex = 312;
			this.cboFechaPrimerPago.KeyDown += new System.Windows.Forms.KeyEventHandler(this.cboFechaPrimerPago_KeyDown);
			// 
			// txtNumCuotas
			// 
			appearance36.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtNumCuotas.Appearance = appearance36;
			this.txtNumCuotas.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtNumCuotas.Enabled = false;
			this.txtNumCuotas.FormatString = "";
			this.txtNumCuotas.Location = new System.Drawing.Point(56, 56);
			this.txtNumCuotas.MaskInput = "nn";
			this.txtNumCuotas.Name = "txtNumCuotas";
			this.txtNumCuotas.PromptChar = ' ';
			this.txtNumCuotas.Size = new System.Drawing.Size(64, 22);
			this.txtNumCuotas.SpinButtonDisplayStyle = Infragistics.Win.ButtonDisplayStyle.Always;
			this.txtNumCuotas.TabIndex = 310;
			this.txtNumCuotas.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtNumCuotas_KeyDown);
			// 
			// txtCuotaDeEntrada
			// 
			appearance37.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtCuotaDeEntrada.Appearance = appearance37;
			this.txtCuotaDeEntrada.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtCuotaDeEntrada.Enabled = false;
			this.txtCuotaDeEntrada.FormatString = "#,##0.00";
			this.txtCuotaDeEntrada.Location = new System.Drawing.Point(232, 56);
			this.txtCuotaDeEntrada.Name = "txtCuotaDeEntrada";
			this.txtCuotaDeEntrada.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtCuotaDeEntrada.PromptChar = ' ';
			this.txtCuotaDeEntrada.Size = new System.Drawing.Size(80, 22);
			this.txtCuotaDeEntrada.TabIndex = 311;
			this.txtCuotaDeEntrada.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtCuotaDeEntrada_KeyDown);
			// 
			// utcTransaccion
			// 
			this.utcTransaccion.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			appearance38.ForeColorDisabled = System.Drawing.Color.Black;
			this.utcTransaccion.Appearance = appearance38;
			this.utcTransaccion.Controls.Add(this.ultraTabSharedControlsPage1);
			this.utcTransaccion.Controls.Add(this.ultraTabPageControl1);
			this.utcTransaccion.Controls.Add(this.ultraTabPageControl9);
			this.utcTransaccion.Location = new System.Drawing.Point(8, 265);
			this.utcTransaccion.Name = "utcTransaccion";
			this.utcTransaccion.SharedControlsPage = this.ultraTabSharedControlsPage1;
			this.utcTransaccion.Size = new System.Drawing.Size(696, 159);
			this.utcTransaccion.Style = Infragistics.Win.UltraWinTabControl.UltraTabControlStyle.Excel;
			this.utcTransaccion.TabIndex = 336;
			ultraTab1.TabPage = this.ultraTabPageControl9;
			ultraTab1.Text = "Seriales";
			ultraTab2.TabPage = this.ultraTabPageControl1;
			ultraTab2.Text = "Notas";
			this.utcTransaccion.Tabs.AddRange(new Infragistics.Win.UltraWinTabControl.UltraTab[] {
																																														 ultraTab1,
																																														 ultraTab2});
			this.utcTransaccion.ViewStyle = Infragistics.Win.UltraWinTabControl.ViewStyle.Office2003;
			// 
			// ultraTabSharedControlsPage1
			// 
			this.ultraTabSharedControlsPage1.Location = new System.Drawing.Point(-10000, -10000);
			this.ultraTabSharedControlsPage1.Name = "ultraTabSharedControlsPage1";
			this.ultraTabSharedControlsPage1.Size = new System.Drawing.Size(694, 138);
			// 
			// btnClaveDescuento
			// 
			this.btnClaveDescuento.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btnClaveDescuento.CausesValidation = false;
			this.btnClaveDescuento.Enabled = false;
			this.btnClaveDescuento.Location = new System.Drawing.Point(711, 448);
			this.btnClaveDescuento.Name = "btnClaveDescuento";
			this.btnClaveDescuento.Size = new System.Drawing.Size(24, 22);
			this.btnClaveDescuento.TabIndex = 346;
			this.btnClaveDescuento.Text = "D";
			this.btnClaveDescuento.Visible = false;
			this.btnClaveDescuento.Click += new System.EventHandler(this.btnClaveDescuento_Click);
			// 
			// btnAprobacion
			// 
			this.btnAprobacion.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btnAprobacion.CausesValidation = false;
			this.btnAprobacion.Enabled = false;
			this.btnAprobacion.Image = ((System.Drawing.Image)(resources.GetObject("btnAprobacion.Image")));
			this.btnAprobacion.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnAprobacion.Location = new System.Drawing.Point(397, 448);
			this.btnAprobacion.Name = "btnAprobacion";
			this.btnAprobacion.Size = new System.Drawing.Size(83, 23);
			this.btnAprobacion.TabIndex = 342;
			this.btnAprobacion.Text = "&Aprobación";
			this.btnAprobacion.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnAprobacion.Click += new System.EventHandler(this.mnuAprobacionAprobar_Click);
			// 
			// btnNuevo
			// 
			this.btnNuevo.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btnNuevo.CausesValidation = false;
			this.btnNuevo.Image = ((System.Drawing.Image)(resources.GetObject("btnNuevo.Image")));
			this.btnNuevo.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnNuevo.Location = new System.Drawing.Point(8, 448);
			this.btnNuevo.Name = "btnNuevo";
			this.btnNuevo.Size = new System.Drawing.Size(72, 23);
			this.btnNuevo.TabIndex = 337;
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
			this.btnConsultar.Location = new System.Drawing.Point(82, 448);
			this.btnConsultar.Name = "btnConsultar";
			this.btnConsultar.Size = new System.Drawing.Size(78, 23);
			this.btnConsultar.TabIndex = 341;
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
			this.btnEditar.Location = new System.Drawing.Point(162, 448);
			this.btnEditar.Name = "btnEditar";
			this.btnEditar.Size = new System.Drawing.Size(76, 23);
			this.btnEditar.TabIndex = 340;
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
			this.btnImprimir.Location = new System.Drawing.Point(240, 448);
			this.btnImprimir.Name = "btnImprimir";
			this.btnImprimir.Size = new System.Drawing.Size(77, 23);
			this.btnImprimir.TabIndex = 338;
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
			this.btnGuardar.Location = new System.Drawing.Point(319, 448);
			this.btnGuardar.Name = "btnGuardar";
			this.btnGuardar.Size = new System.Drawing.Size(76, 23);
			this.btnGuardar.TabIndex = 339;
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
			this.btnAnular.Location = new System.Drawing.Point(482, 448);
			this.btnAnular.Name = "btnAnular";
			this.btnAnular.Size = new System.Drawing.Size(76, 23);
			this.btnAnular.TabIndex = 343;
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
			this.btnCancelar.Location = new System.Drawing.Point(560, 448);
			this.btnCancelar.Name = "btnCancelar";
			this.btnCancelar.Size = new System.Drawing.Size(72, 23);
			this.btnCancelar.TabIndex = 344;
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
			this.btnSalir.Location = new System.Drawing.Point(634, 448);
			this.btnSalir.Name = "btnSalir";
			this.btnSalir.TabIndex = 345;
			this.btnSalir.Text = "&Salir";
			this.btnSalir.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnSalir.Click += new System.EventHandler(this.btnSalir_Click);
			// 
			// label1
			// 
			this.label1.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			this.label1.AutoSize = true;
			this.label1.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label1.Location = new System.Drawing.Point(728, 377);
			this.label1.Name = "label1";
			this.label1.Size = new System.Drawing.Size(29, 17);
			this.label1.TabIndex = 354;
			this.label1.Text = "Total";
			// 
			// lblIva
			// 
			this.lblIva.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			this.lblIva.AutoSize = true;
			this.lblIva.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.lblIva.Location = new System.Drawing.Point(728, 347);
			this.lblIva.Name = "lblIva";
			this.lblIva.Size = new System.Drawing.Size(22, 17);
			this.lblIva.TabIndex = 348;
			this.lblIva.Text = "IVA";
			// 
			// txtTotal
			// 
			this.txtTotal.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			appearance39.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtTotal.Appearance = appearance39;
			this.txtTotal.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtTotal.Enabled = false;
			this.txtTotal.Font = new System.Drawing.Font("Tahoma", 15.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.txtTotal.FormatString = "#,##0.00";
			this.txtTotal.Location = new System.Drawing.Point(816, 368);
			this.txtTotal.Name = "txtTotal";
			this.txtTotal.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtTotal.PromptChar = ' ';
			this.txtTotal.Size = new System.Drawing.Size(184, 35);
			this.txtTotal.TabIndex = 353;
			// 
			// txtIvaTotal
			// 
			this.txtIvaTotal.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			appearance40.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtIvaTotal.Appearance = appearance40;
			this.txtIvaTotal.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtIvaTotal.Enabled = false;
			this.txtIvaTotal.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.txtIvaTotal.FormatString = "#,##0.00";
			this.txtIvaTotal.Location = new System.Drawing.Point(920, 344);
			this.txtIvaTotal.Name = "txtIvaTotal";
			this.txtIvaTotal.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtIvaTotal.PromptChar = ' ';
			this.txtIvaTotal.Size = new System.Drawing.Size(80, 22);
			this.txtIvaTotal.TabIndex = 352;
			// 
			// txtDescIvaTotal
			// 
			this.txtDescIvaTotal.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			appearance41.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtDescIvaTotal.Appearance = appearance41;
			this.txtDescIvaTotal.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtDescIvaTotal.Enabled = false;
			this.txtDescIvaTotal.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.txtDescIvaTotal.FormatString = "#,##0.00";
			this.txtDescIvaTotal.Location = new System.Drawing.Point(920, 320);
			this.txtDescIvaTotal.Name = "txtDescIvaTotal";
			this.txtDescIvaTotal.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtDescIvaTotal.PromptChar = ' ';
			this.txtDescIvaTotal.Size = new System.Drawing.Size(80, 22);
			this.txtDescIvaTotal.TabIndex = 351;
			// 
			// txtIva
			// 
			this.txtIva.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			appearance42.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtIva.Appearance = appearance42;
			this.txtIva.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtIva.Enabled = false;
			this.txtIva.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.txtIva.FormatString = "#,##0.00";
			this.txtIva.Location = new System.Drawing.Point(920, 296);
			this.txtIva.Name = "txtIva";
			this.txtIva.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtIva.PromptChar = ' ';
			this.txtIva.Size = new System.Drawing.Size(80, 22);
			this.txtIva.TabIndex = 350;
			// 
			// txtidEntidadFinanciera
			// 
			this.txtidEntidadFinanciera.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			this.txtidEntidadFinanciera.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtidEntidadFinanciera.Enabled = false;
			this.txtidEntidadFinanciera.FormatString = "#,##0.00";
			this.txtidEntidadFinanciera.Location = new System.Drawing.Point(784, 456);
			this.txtidEntidadFinanciera.Name = "txtidEntidadFinanciera";
			this.txtidEntidadFinanciera.PromptChar = ' ';
			this.txtidEntidadFinanciera.Size = new System.Drawing.Size(16, 22);
			this.txtidEntidadFinanciera.TabIndex = 368;
			this.txtidEntidadFinanciera.Visible = false;
			// 
			// txtInteres
			// 
			this.txtInteres.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			this.txtInteres.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtInteres.Enabled = false;
			this.txtInteres.FormatString = "#,##0.00";
			this.txtInteres.Location = new System.Drawing.Point(800, 456);
			this.txtInteres.Name = "txtInteres";
			this.txtInteres.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtInteres.PromptChar = ' ';
			this.txtInteres.Size = new System.Drawing.Size(16, 22);
			this.txtInteres.TabIndex = 367;
			this.txtInteres.Visible = false;
			// 
			// txtTotalCuotas
			// 
			this.txtTotalCuotas.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			this.txtTotalCuotas.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtTotalCuotas.Enabled = false;
			this.txtTotalCuotas.FormatString = "#,##0.00";
			this.txtTotalCuotas.Location = new System.Drawing.Point(816, 456);
			this.txtTotalCuotas.Name = "txtTotalCuotas";
			this.txtTotalCuotas.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtTotalCuotas.PromptChar = ' ';
			this.txtTotalCuotas.Size = new System.Drawing.Size(16, 22);
			this.txtTotalCuotas.TabIndex = 366;
			this.txtTotalCuotas.Visible = false;
			// 
			// txtidTipoRuc
			// 
			this.txtidTipoRuc.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			this.txtidTipoRuc.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtidTipoRuc.Enabled = false;
			this.txtidTipoRuc.Location = new System.Drawing.Point(992, 456);
			this.txtidTipoRuc.Name = "txtidTipoRuc";
			this.txtidTipoRuc.PromptChar = ' ';
			this.txtidTipoRuc.Size = new System.Drawing.Size(16, 22);
			this.txtidTipoRuc.TabIndex = 365;
			this.txtidTipoRuc.Visible = false;
			// 
			// txtIdFormaPago
			// 
			this.txtIdFormaPago.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			this.txtIdFormaPago.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtIdFormaPago.Enabled = false;
			this.txtIdFormaPago.Location = new System.Drawing.Point(976, 456);
			this.txtIdFormaPago.Name = "txtIdFormaPago";
			this.txtIdFormaPago.PromptChar = ' ';
			this.txtIdFormaPago.Size = new System.Drawing.Size(16, 22);
			this.txtIdFormaPago.TabIndex = 364;
			this.txtIdFormaPago.Visible = false;
			// 
			// txtNumIdProyecto
			// 
			this.txtNumIdProyecto.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			this.txtNumIdProyecto.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtNumIdProyecto.Enabled = false;
			this.txtNumIdProyecto.Location = new System.Drawing.Point(880, 456);
			this.txtNumIdProyecto.Name = "txtNumIdProyecto";
			this.txtNumIdProyecto.PromptChar = ' ';
			this.txtNumIdProyecto.Size = new System.Drawing.Size(16, 22);
			this.txtNumIdProyecto.TabIndex = 363;
			this.txtNumIdProyecto.Visible = false;
			// 
			// txtIdComprobante
			// 
			this.txtIdComprobante.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			this.txtIdComprobante.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtIdComprobante.Enabled = false;
			this.txtIdComprobante.Location = new System.Drawing.Point(960, 456);
			this.txtIdComprobante.Name = "txtIdComprobante";
			this.txtIdComprobante.PromptChar = ' ';
			this.txtIdComprobante.Size = new System.Drawing.Size(16, 22);
			this.txtIdComprobante.TabIndex = 362;
			this.txtIdComprobante.Visible = false;
			// 
			// txtNumEstado
			// 
			this.txtNumEstado.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			this.txtNumEstado.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtNumEstado.Enabled = false;
			this.txtNumEstado.Location = new System.Drawing.Point(944, 456);
			this.txtNumEstado.Name = "txtNumEstado";
			this.txtNumEstado.PromptChar = ' ';
			this.txtNumEstado.Size = new System.Drawing.Size(16, 22);
			this.txtNumEstado.TabIndex = 361;
			this.txtNumEstado.Visible = false;
			// 
			// txtIdTipoFactura
			// 
			this.txtIdTipoFactura.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			this.txtIdTipoFactura.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtIdTipoFactura.Enabled = false;
			this.txtIdTipoFactura.Location = new System.Drawing.Point(928, 456);
			this.txtIdTipoFactura.Name = "txtIdTipoFactura";
			this.txtIdTipoFactura.PromptChar = ' ';
			this.txtIdTipoFactura.Size = new System.Drawing.Size(16, 22);
			this.txtIdTipoFactura.TabIndex = 360;
			this.txtIdTipoFactura.Visible = false;
			// 
			// txtBodegaPredef
			// 
			this.txtBodegaPredef.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			appearance43.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(230)), ((System.Byte)(230)), ((System.Byte)(230)));
			appearance43.ForeColor = System.Drawing.Color.FromArgb(((System.Byte)(50)), ((System.Byte)(50)), ((System.Byte)(50)));
			this.txtBodegaPredef.Appearance = appearance43;
			this.txtBodegaPredef.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(230)), ((System.Byte)(230)), ((System.Byte)(230)));
			this.txtBodegaPredef.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtBodegaPredef.Enabled = false;
			this.txtBodegaPredef.Font = new System.Drawing.Font("Tahoma", 8.25F);
			this.txtBodegaPredef.ForeColor = System.Drawing.Color.FromArgb(((System.Byte)(50)), ((System.Byte)(50)), ((System.Byte)(50)));
			this.txtBodegaPredef.Location = new System.Drawing.Point(912, 456);
			this.txtBodegaPredef.Name = "txtBodegaPredef";
			this.txtBodegaPredef.PromptChar = ' ';
			this.txtBodegaPredef.Size = new System.Drawing.Size(16, 22);
			this.txtBodegaPredef.TabIndex = 359;
			this.txtBodegaPredef.Visible = false;
			// 
			// txtIdSucursal
			// 
			this.txtIdSucursal.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			this.txtIdSucursal.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtIdSucursal.Enabled = false;
			this.txtIdSucursal.Location = new System.Drawing.Point(896, 456);
			this.txtIdSucursal.Name = "txtIdSucursal";
			this.txtIdSucursal.PromptChar = ' ';
			this.txtIdSucursal.Size = new System.Drawing.Size(16, 22);
			this.txtIdSucursal.TabIndex = 358;
			this.txtIdSucursal.Visible = false;
			// 
			// txtNumContadoCredito
			// 
			this.txtNumContadoCredito.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			this.txtNumContadoCredito.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtNumContadoCredito.Enabled = false;
			this.txtNumContadoCredito.Location = new System.Drawing.Point(864, 456);
			this.txtNumContadoCredito.Name = "txtNumContadoCredito";
			this.txtNumContadoCredito.PromptChar = ' ';
			this.txtNumContadoCredito.Size = new System.Drawing.Size(16, 22);
			this.txtNumContadoCredito.TabIndex = 357;
			this.txtNumContadoCredito.Visible = false;
			// 
			// txtNumIdCompra
			// 
			this.txtNumIdCompra.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			this.txtNumIdCompra.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtNumIdCompra.Enabled = false;
			this.txtNumIdCompra.Location = new System.Drawing.Point(832, 456);
			this.txtNumIdCompra.Name = "txtNumIdCompra";
			this.txtNumIdCompra.PromptChar = ' ';
			this.txtNumIdCompra.Size = new System.Drawing.Size(16, 22);
			this.txtNumIdCompra.TabIndex = 356;
			// 
			// txtNumIdCliente
			// 
			this.txtNumIdCliente.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			this.txtNumIdCliente.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtNumIdCliente.Enabled = false;
			this.txtNumIdCliente.Location = new System.Drawing.Point(848, 456);
			this.txtNumIdCliente.Name = "txtNumIdCliente";
			this.txtNumIdCliente.PromptChar = ' ';
			this.txtNumIdCliente.Size = new System.Drawing.Size(16, 22);
			this.txtNumIdCliente.TabIndex = 355;
			// 
			// lblComprobante
			// 
			this.lblComprobante.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.lblComprobante.AutoSize = true;
			this.lblComprobante.BackColor = System.Drawing.Color.Transparent;
			this.lblComprobante.Location = new System.Drawing.Point(512, 83);
			this.lblComprobante.Name = "lblComprobante";
			this.lblComprobante.Size = new System.Drawing.Size(71, 17);
			this.lblComprobante.TabIndex = 369;
			this.lblComprobante.Text = "Comprobante";
			// 
			// txtComprobante
			// 
			appearance44.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtComprobante.Appearance = appearance44;
			this.txtComprobante.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtComprobante.Enabled = false;
			this.txtComprobante.Location = new System.Drawing.Point(604, 80);
			this.txtComprobante.Name = "txtComprobante";
			this.txtComprobante.Size = new System.Drawing.Size(112, 22);
			this.txtComprobante.TabIndex = 370;
			// 
			// label3
			// 
			this.label3.AutoSize = true;
			this.label3.Location = new System.Drawing.Point(8, 83);
			this.label3.Name = "label3";
			this.label3.Size = new System.Drawing.Size(32, 17);
			this.label3.TabIndex = 424;
			this.label3.Text = "Serial";
			this.label3.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// label15
			// 
			this.label15.AutoSize = true;
			this.label15.Location = new System.Drawing.Point(432, 10);
			this.label15.Name = "label15";
			this.label15.Size = new System.Drawing.Size(41, 17);
			this.label15.TabIndex = 427;
			this.label15.Text = "Bodega";
			this.label15.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// cmbBodega1
			// 
			this.cmbBodega1.CausesValidation = false;
			this.cmbBodega1.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbBodega1.Cursor = System.Windows.Forms.Cursors.Default;
			ultraGridColumn39.Header.VisiblePosition = 2;
			ultraGridColumn40.Header.VisiblePosition = 1;
			ultraGridColumn40.Width = 180;
			ultraGridColumn41.Header.VisiblePosition = 0;
			ultraGridBand5.Columns.AddRange(new object[] {
																										 ultraGridColumn39,
																										 ultraGridColumn40,
																										 ultraGridColumn41});
			this.cmbBodega1.DisplayLayout.BandsSerializer.Add(ultraGridBand5);
			this.cmbBodega1.DisplayMember = "Nombre";
			this.cmbBodega1.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbBodega1.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbBodega1.Enabled = false;
			this.cmbBodega1.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbBodega1.Location = new System.Drawing.Point(488, 7);
			this.cmbBodega1.MaxDropDownItems = 30;
			this.cmbBodega1.Name = "cmbBodega1";
			this.cmbBodega1.Size = new System.Drawing.Size(180, 22);
			this.cmbBodega1.TabIndex = 426;
			this.cmbBodega1.ValueMember = "Bodega";
			this.cmbBodega1.KeyDown += new System.Windows.Forms.KeyEventHandler(this.cmbBodega1_KeyDown);
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
			// MnuAprobacion
			// 
			this.MnuAprobacion.MenuItems.AddRange(new System.Windows.Forms.MenuItem[] {
																																									this.mnuAprobacionAprobar,
																																									this.mnuAprobacionRechazar});
			// 
			// mnuAprobacionAprobar
			// 
			this.mnuAprobacionAprobar.Enabled = false;
			this.mnuAprobacionAprobar.Index = 0;
			this.mnuAprobacionAprobar.Text = "Aprobar";
			this.mnuAprobacionAprobar.Click += new System.EventHandler(this.mnuAprobacionAprobar_Click);
			// 
			// mnuAprobacionRechazar
			// 
			this.mnuAprobacionRechazar.Enabled = false;
			this.mnuAprobacionRechazar.Index = 1;
			this.mnuAprobacionRechazar.Text = "Rechazar";
			this.mnuAprobacionRechazar.Click += new System.EventHandler(this.mnuAprobacionRechazar_Click);
			// 
			// ultraDataSource1
			// 
			ultraDataColumn33.DataType = typeof(int);
			this.ultraDataSource1.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn33,
																																 ultraDataColumn34});
			// 
			// lblUsuario
			// 
			this.lblUsuario.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.lblUsuario.AutoSize = true;
			this.lblUsuario.BackColor = System.Drawing.Color.Transparent;
			this.lblUsuario.Font = new System.Drawing.Font("Tahoma", 10F, ((System.Drawing.FontStyle)((System.Drawing.FontStyle.Bold | System.Drawing.FontStyle.Underline))));
			this.lblUsuario.Location = new System.Drawing.Point(728, 408);
			this.lblUsuario.Name = "lblUsuario";
			this.lblUsuario.Size = new System.Drawing.Size(0, 20);
			this.lblUsuario.TabIndex = 428;
			this.lblUsuario.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
			// 
			// txtIngresoSerial
			// 
			appearance45.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtIngresoSerial.Appearance = appearance45;
			this.txtIngresoSerial.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtIngresoSerial.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtIngresoSerial.Enabled = false;
			this.txtIngresoSerial.Location = new System.Drawing.Point(56, 80);
			this.txtIngresoSerial.Name = "txtIngresoSerial";
			this.txtIngresoSerial.Size = new System.Drawing.Size(256, 22);
			this.txtIngresoSerial.TabIndex = 429;
			this.txtIngresoSerial.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtIngresoSerial_KeyDown);
			this.txtIngresoSerial.ValueChanged += new System.EventHandler(this.txtIngresoSerial_ValueChanged);
			// 
			// label7
			// 
			this.label7.AutoSize = true;
			this.label7.BackColor = System.Drawing.Color.Transparent;
			this.label7.Location = new System.Drawing.Point(320, 59);
			this.label7.Name = "label7";
			this.label7.Size = new System.Drawing.Size(64, 17);
			this.label7.TabIndex = 317;
			this.label7.Text = "Día de pago";
			// 
			// groupBox1
			// 
			this.groupBox1.Location = new System.Drawing.Point(0, 432);
			this.groupBox1.Name = "groupBox1";
			this.groupBox1.Size = new System.Drawing.Size(1008, 8);
			this.groupBox1.TabIndex = 430;
			this.groupBox1.TabStop = false;
			// 
			// label8
			// 
			this.label8.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			this.label8.AutoSize = true;
			this.label8.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label8.Location = new System.Drawing.Point(728, 299);
			this.label8.Name = "label8";
			this.label8.Size = new System.Drawing.Size(45, 17);
			this.label8.TabIndex = 432;
			this.label8.Text = "Subtotal";
			this.label8.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// label9
			// 
			this.label9.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			this.label9.AutoSize = true;
			this.label9.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label9.Location = new System.Drawing.Point(728, 323);
			this.label9.Name = "label9";
			this.label9.Size = new System.Drawing.Size(57, 17);
			this.label9.TabIndex = 431;
			this.label9.Text = "Descuento";
			this.label9.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// txtDescIva0Total
			// 
			this.txtDescIva0Total.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			appearance46.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtDescIva0Total.Appearance = appearance46;
			this.txtDescIva0Total.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtDescIva0Total.Enabled = false;
			this.txtDescIva0Total.FormatString = "#,##0.00";
			this.txtDescIva0Total.Location = new System.Drawing.Point(816, 320);
			this.txtDescIva0Total.Name = "txtDescIva0Total";
			this.txtDescIva0Total.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtDescIva0Total.PromptChar = ' ';
			this.txtDescIva0Total.Size = new System.Drawing.Size(80, 22);
			this.txtDescIva0Total.TabIndex = 434;
			// 
			// txtIva0
			// 
			this.txtIva0.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			appearance47.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtIva0.Appearance = appearance47;
			this.txtIva0.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtIva0.Enabled = false;
			this.txtIva0.FormatString = "#,##0.00";
			this.txtIva0.Location = new System.Drawing.Point(816, 296);
			this.txtIva0.Name = "txtIva0";
			this.txtIva0.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtIva0.PromptChar = ' ';
			this.txtIva0.Size = new System.Drawing.Size(80, 22);
			this.txtIva0.TabIndex = 433;
			// 
			// lblIva0
			// 
			this.lblIva0.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			this.lblIva0.AutoSize = true;
			this.lblIva0.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.lblIva0.Location = new System.Drawing.Point(816, 272);
			this.lblIva0.Name = "lblIva0";
			this.lblIva0.Size = new System.Drawing.Size(32, 17);
			this.lblIva0.TabIndex = 435;
			this.lblIva0.Text = "IVA 0";
			this.lblIva0.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// label55
			// 
			this.label55.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			this.label55.AutoSize = true;
			this.label55.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label55.Location = new System.Drawing.Point(920, 272);
			this.label55.Name = "label55";
			this.label55.Size = new System.Drawing.Size(22, 17);
			this.label55.TabIndex = 436;
			this.label55.Text = "IVA";
			this.label55.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// frmPrefacturaRemate
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 14);
			this.ClientSize = new System.Drawing.Size(1008, 478);
			this.Controls.Add(this.lblIva0);
			this.Controls.Add(this.label55);
			this.Controls.Add(this.label8);
			this.Controls.Add(this.label9);
			this.Controls.Add(this.label15);
			this.Controls.Add(this.label3);
			this.Controls.Add(this.lblComprobante);
			this.Controls.Add(this.label1);
			this.Controls.Add(this.lblIva);
			this.Controls.Add(this.label2);
			this.Controls.Add(this.label63);
			this.Controls.Add(this.label64);
			this.Controls.Add(this.label65);
			this.Controls.Add(this.label23);
			this.Controls.Add(this.lblIdentificacion);
			this.Controls.Add(this.lblCliente);
			this.Controls.Add(this.lblBodega);
			this.Controls.Add(this.lblVendedor);
			this.Controls.Add(this.label6);
			this.Controls.Add(this.label21);
			this.Controls.Add(this.label61);
			this.Controls.Add(this.label7);
			this.Controls.Add(this.txtDescIva0Total);
			this.Controls.Add(this.txtIva0);
			this.Controls.Add(this.groupBox1);
			this.Controls.Add(this.txtIngresoSerial);
			this.Controls.Add(this.lblUsuario);
			this.Controls.Add(this.cmbBodega1);
			this.Controls.Add(this.txtComprobante);
			this.Controls.Add(this.txtidEntidadFinanciera);
			this.Controls.Add(this.txtInteres);
			this.Controls.Add(this.txtTotalCuotas);
			this.Controls.Add(this.txtidTipoRuc);
			this.Controls.Add(this.txtIdFormaPago);
			this.Controls.Add(this.txtNumIdProyecto);
			this.Controls.Add(this.txtIdComprobante);
			this.Controls.Add(this.txtNumEstado);
			this.Controls.Add(this.txtIdTipoFactura);
			this.Controls.Add(this.txtBodegaPredef);
			this.Controls.Add(this.txtIdSucursal);
			this.Controls.Add(this.txtNumContadoCredito);
			this.Controls.Add(this.txtNumIdCompra);
			this.Controls.Add(this.txtNumIdCliente);
			this.Controls.Add(this.txtTotal);
			this.Controls.Add(this.txtIvaTotal);
			this.Controls.Add(this.txtDescIvaTotal);
			this.Controls.Add(this.txtIva);
			this.Controls.Add(this.btnClaveDescuento);
			this.Controls.Add(this.btnAprobacion);
			this.Controls.Add(this.btnNuevo);
			this.Controls.Add(this.btnConsultar);
			this.Controls.Add(this.btnEditar);
			this.Controls.Add(this.btnImprimir);
			this.Controls.Add(this.btnGuardar);
			this.Controls.Add(this.btnAnular);
			this.Controls.Add(this.btnCancelar);
			this.Controls.Add(this.btnSalir);
			this.Controls.Add(this.utcTransaccion);
			this.Controls.Add(this.chkAutorizar);
			this.Controls.Add(this.chkVerificar);
			this.Controls.Add(this.txtNumero);
			this.Controls.Add(this.lblEstado);
			this.Controls.Add(this.txtCuotaAsignada);
			this.Controls.Add(this.txtCuotaDisponible);
			this.Controls.Add(this.txtCuotaUsada);
			this.Controls.Add(this.txtValorCuotas);
			this.Controls.Add(this.ultraGrid2);
			this.Controls.Add(this.txtRuc);
			this.Controls.Add(this.cmbBodega);
			this.Controls.Add(this.txtNombre);
			this.Controls.Add(this.dtFecha);
			this.Controls.Add(this.cmbVendedor);
			this.Controls.Add(this.cboFechaPrimerPago);
			this.Controls.Add(this.txtNumCuotas);
			this.Controls.Add(this.txtCuotaDeEntrada);
			this.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.Name = "frmPrefacturaRemate";
			this.Text = "Pre Factura De Remates";
			this.KeyDown += new System.Windows.Forms.KeyEventHandler(this.frmPrefactura_KeyDown);
			this.Resize += new System.EventHandler(this.frmPrefactura_Resize);
			this.Closing += new System.ComponentModel.CancelEventHandler(this.frmPrefactura_Closing);
			this.Load += new System.EventHandler(this.frmPrefactura_Load);
			this.ultraTabPageControl9.ResumeLayout(false);
			((System.ComponentModel.ISupportInitialize)(this.grdSerial)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource4)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtSerial)).EndInit();
			this.ultraTabPageControl1.ResumeLayout(false);
			((System.ComponentModel.ISupportInitialize)(this.txtNotas)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtObservaciones)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNumero)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtCuotaAsignada)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtCuotaDisponible)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtCuotaUsada)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtValorCuotas)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraGrid2)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource2)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtRuc)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbBodega)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNombre)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.dtFecha)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbVendedor)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNumCuotas)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtCuotaDeEntrada)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.utcTransaccion)).EndInit();
			this.utcTransaccion.ResumeLayout(false);
			((System.ComponentModel.ISupportInitialize)(this.txtTotal)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIvaTotal)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtDescIvaTotal)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIva)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtidEntidadFinanciera)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtInteres)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtTotalCuotas)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtidTipoRuc)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIdFormaPago)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNumIdProyecto)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIdComprobante)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNumEstado)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIdTipoFactura)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtBodegaPredef)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIdSucursal)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNumContadoCredito)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNumIdCompra)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNumIdCliente)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtComprobante)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbBodega1)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource1)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIngresoSerial)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtDescIva0Total)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIva0)).EndInit();
			this.ResumeLayout(false);

		}
		#endregion

		#region Variables Globales
//		int idTipoOrigen = 0;
//		bool ModSerial = false;
//		decimal dPorcEnt = 0.00m; 
//		decimal dPorcEntMin = 0.00m;
//		decimal dEntrada = 0.00m;
//		decimal dEntMin = 0.00m;
//		int idPromocionCombo = 0;
//		int IdArtProm = 0;
		bool bEliminaAlValidar = false;
		private Acceso miAcceso;
		private C1DataRow drSeteoF;
//		private C1DataRow drSeteoC;
//		private C1DataRow drSeteo;
//		private int IdSigno = 2;
		private bool bIncluyeIva = false;
//		string controlW="";
//		string controlW1="";		
		private bool bActualiza = true;
		CultureInfo us = new CultureInfo("en-US");
		bool bRastreo = false;
		bool bNuevo = false;
		bool bEdicion = false;
		bool bAnulado = false;
		bool bImpreso = false;
		bool bPromocionCombo = false;
		bool bPromocion = false;
		bool bCombo = false;
		int idTipoPrecio = 4;
//		decimal dCuotaAsignada = 0.00m;
//		decimal dInteres = 0.00m;
//		decimal dCuotaDisponible = 0.00m;
//		decimal dCuotaUsada = 0.00m;
//    bool AutorizaEntrada = false;
		bool bCrediPoint = false;
		bool bDescuentoActivo = false;
		int idBloqueaTransacciones = 0;
		decimal dClavePorcDscto = 0.00m;
		int iDiaMaximoDePago = 0;
		int iMinimoDiasPago = 0;
		int iMaximoDiasPago = 0;			
		decimal dPorcIva = 0.00m;
		int idInstitucion = 0;
		bool bModificaColor = false;
		bool Combo = false;
		bool bArtPromocion = false;
		decimal dPorcentajeCuponWeb = 0.00m;
		int idCuponWeb = 0;
		decimal dDsctoReiterativo = 0.00m;
		bool bElCombo = false;
		int idComboElimina = 0;
//		int idArticuloElimina = 0;
		int RegistroIdSerial = 0;
		bool bEliminando = false;
		#endregion Variables Globales

		#region DataSets
		private void cdsSeteoF_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			cdsSeteoF.Schema.Connections[0].ConnectionString = miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);		
		}		
		#endregion DataSets

		#region Funciones de inicio
		private int WM_CLOSE = 0x0010;
		protected override void WndProc(ref System.Windows.Forms.Message msg)
		{
			if (msg.Msg == WM_CLOSE) 
			{
				ActiveControl.CausesValidation = false;
			}			
			base.WndProc(ref msg);			
		}
		
		private void UnloadMe()
		{
			this.Close();
		}				
		#endregion Funciones de inicio

		#region Funciones de la Barra datos
		private void ValidaNumeroCuotas()
		{
			int iMinimoNumeroCuotas = 0;
			int iMaximoNumeroCuotas = 0;
			
			DateTime dtFecha = Convert.ToDateTime(this.dtFecha.Value);
			string sSQLParametros = string.Format("Exec Cre_RetornaCuotasDias {0}, '{1}'", (int)this.txtidEntidadFinanciera.Value, dtFecha.ToString("yyyyMMdd"));
			SqlDataReader dr = Funcion.rEscalarSQL(cdsSeteoF, sSQLParametros, true);
			
			dr.Read();
			if (dr.HasRows)
			{				
				iMinimoNumeroCuotas = dr.GetInt32(0);
				iMaximoNumeroCuotas = dr.GetInt32(1);
			}
			dr.Close();

			this.txtNumCuotas.MinValue = iMinimoNumeroCuotas;
			this.txtNumCuotas.MaxValue = 18;//iMaximoNumeroCuotas;
			this.txtNumCuotas.Value = 18;//iMaximoNumeroCuotas;
		}

		private void DiasDePago()
		{
			DateTime dtFecha = Convert.ToDateTime(this.dtFecha.Value);
			string sSQLParametros = string.Format("Exec Cre_RetornaCuotasDias {0}, '{1}'", (int)this.txtidEntidadFinanciera.Value, dtFecha.ToString("yyyyMMdd"));
			SqlDataReader dr = Funcion.rEscalarSQL(cdsSeteoF, sSQLParametros, true);
			
			dr.Read();
			if (dr.HasRows)
			{				
				iMinimoDiasPago = dr.GetInt32(2);
				iMaximoDiasPago = dr.GetInt32(3);
				iDiaMaximoDePago = dr.GetInt32(4);
			}
			dr.Close();
		}

		private void CargaFecha()
		{			
			this.cboFechaPrimerPago.Items.Clear();
						
			DiasDePago();

			DateTime fecha = (DateTime) this.dtFecha.Value;
			DateTime fechaIni = fecha.AddDays(iMinimoDiasPago);
			DateTime fechaFin = fecha.AddDays(iMaximoDiasPago);
			
			int DiaIni = fechaIni.Date.Day;

			if (DiaIni == 28)
			{
				fechaIni = fechaIni.AddDays(1);
				iMaximoDiasPago--;
			}

			DiaIni = fechaIni.Date.Day;

			if (DiaIni == 29) 
			{
				fechaIni = fechaIni.AddDays(1);
				iMaximoDiasPago--;
			}

			DiaIni = fechaIni.Date.Day;

			if (DiaIni == 30) 
			{
				fechaIni = fechaIni.AddDays(1);
				iMaximoDiasPago--;
			}

			DiaIni = fechaIni.Date.Day;

			if (DiaIni == 31) 
			{
				fechaIni = fechaIni.AddDays(1);
				iMaximoDiasPago--;
			}

			DiaIni = fechaIni.Date.Day;
			
			this.cboFechaPrimerPago.Items.Add(fechaIni.ToString("dd/MM/yyyy"));

			for (int i = iMinimoDiasPago; i < iMaximoDiasPago; i++)
			{
				DateTime Fecha = fechaIni.Date.AddDays(1);

				int Dia = Fecha.Date.Day;
				
				if (Dia <= iDiaMaximoDePago) this.cboFechaPrimerPago.Items.Add(Fecha.ToString("dd/MM/yyyy"));
				
				fechaIni = Fecha;
			}
		}

		private decimal SumaPrecios()
		{
			decimal dPrecio = 0.00m;

			foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in ultraGrid2.Rows.All)
			{
				if (dr.Cells["Tipo"].Value != System.DBNull.Value)	
				{
					if (dr.Cells["Subtotal"].Value != System.DBNull.Value)
					{
						if ((int)dr.Cells["Tipo"].Value != 3)	dPrecio = dPrecio + Convert.ToDecimal(dr.Cells["Subtotal"].Value);
					}
				}
			}

			dPrecio = dPrecio * dPorcIva;

			return dPrecio;
		}

		private void cmdTotal()
		{
			#region Creacion de Variables
			decimal dDescuentoLineal = 0.00m;
			int dRedondeo = (int) drSeteoF["DecRedondeo"];
			#endregion Creacion de Variables
			
			this.txtIva0.Value = 0.00m;
			this.txtIva.Value = 0.00m;
			this.txtDescIva0Total.Value = 0.00m;
			this.txtDescIvaTotal.Value = 0.00m;
			this.txtIvaTotal.Value = 0.00m;
			this.txtTotal.Value = 0.00m;
			this.txtValorCuotas.Value = 0.00m;
			this.txtTotalCuotas.Value = 0.00m;

			dPorcIva = Funcion.decEscalarSQL(cdsSeteoF, string.Format("Select dbo.RetornaImpuestoIVAVigente('{0}', 1)", 
				Convert.ToDateTime(this.dtFecha.Value).ToString("yyyyMMdd")));

			#region Desplazamiento de grilla
			foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in ultraGrid2.Rows.All)
			{
				dDescuentoLineal = 0.00m;
				
				if (dr.Cells["Cantidad"].Value != DBNull.Value && dr.Cells["Precio"].Value != DBNull.Value && dr.Cells["DescuentoPorc"].Value != DBNull.Value)
				{
					dr.Cells["Subtotal"].Value = Math.Round(Convert.ToDecimal( dr.Cells["Cantidad"].Value) * Convert.ToDecimal(dr.Cells["Precio"].Value), 2);
					dDescuentoLineal = Math.Round((Convert.ToDecimal(dr.Cells["Cantidad"].Value) * Convert.ToDecimal(dr.Cells["Precio"].Value) * Convert.ToDecimal(dr.Cells["DescuentoPorc"].Value))/100.00m, 2);
					dr.Cells["DescuentoArt"].Value = dDescuentoLineal;

					if (Convert.ToDecimal(dr.Cells["Impuesto"].Value) > 0.00m)
					{
						this.txtIva.Value = Math.Round(Convert.ToDecimal(this.txtIva.Value) + Convert.ToDecimal(dr.Cells["Subtotal"].Value), 2);
						this.txtDescIvaTotal.Value = Math.Round(Convert.ToDecimal(this.txtDescIvaTotal.Value) + Convert.ToDecimal(dr.Cells["DescuentoArt"].Value), 2);
					}
					else					
					{
						this.txtIva0.Value = Math.Round(Convert.ToDecimal(this.txtIva0.Value) + Convert.ToDecimal(dr.Cells["Subtotal"].Value), 2);
						this.txtDescIva0Total.Value = Math.Round(Convert.ToDecimal(this.txtDescIva0Total.Value) + Convert.ToDecimal(dr.Cells["DescuentoArt"].Value), 2);
					}
				}
				
				this.txtIvaTotal.Value = Math.Round(((Convert.ToDecimal(this.txtIva.Value) - Convert.ToDecimal(this.txtDescIvaTotal.Value)) * dPorcIva)/100.00m, 2);

				this.txtTotal.Value = Math.Round(Convert.ToDecimal(this.txtIva0.Value), 2)
					+ Math.Round(Convert.ToDecimal(this.txtIva.Value), 2)
					- Math.Round(Convert.ToDecimal(this.txtDescIva0Total.Value), 2)
					- Math.Round(Convert.ToDecimal(this.txtDescIvaTotal.Value), 2)
					+ Math.Round(Convert.ToDecimal(this.txtIvaTotal.Value), 2);
			} 
			#endregion Desplazamiento de grilla

			#region Calculo Cuotas
			decimal dMontoCredito = Math.Round(Convert.ToDecimal(this.txtTotal.Value), 2) - Math.Round(Convert.ToDecimal(this.txtCuotaDeEntrada.Value), 2);

			if (dMontoCredito > 0.00m)
			{
				string sSQL1 = string.Format("Exec CalculaCuotaInicial {0}, {1}, {2}, '{3}', {4}", 
					dMontoCredito, (int)this.txtNumCuotas.Value, Convert.ToInt32(this.txtNumIdCliente.Value), Convert.ToDateTime(this.dtFecha.Value).ToString("yyyyMMdd"), Convert.ToInt32(this.txtidEntidadFinanciera.Value));
				SqlDataReader dReader =	Funcion.rEscalarSQL(cdsSeteoF, sSQL1, true);

				dReader.Read();
				decimal dValorc = Convert.ToDecimal(dReader.GetValue(2));
				this.txtValorCuotas.Value = Convert.ToDecimal(dReader.GetValue(2));
				dReader.Close();				
			
				this.txtTotalCuotas.Value = Math.Round(Convert.ToDecimal(this.txtValorCuotas.Value), 2) * (int)this.txtNumCuotas.Value;			
			}		
			#endregion Calculo Cuotas
		}
		#endregion Funciones de la Barra datos

		#region Funciones del grid
		private void Rastreo(string stMensaje)
		{
			string stSelect = "Insert into ErrorSalida (Mensaje) Values ('" + stMensaje + "')";
			Funcion.EjecutaSQL(cdsSeteoF, stSelect);
		}
			
		private string revisaCadena(string cadenaOrg)
		{
			string letra = "";
			string stL = "'";
			string cadena = cadenaOrg;
			string cadena1 = "";
			for (int i=0;i<cadena.Length;i++)
			{
				letra = cadena.Substring(i,1);				
				if(letra == stL) cadena1 += "-";
				else cadena1 += letra;
			}
			return cadena1;
		}
		#endregion Funciones del grid

		private void frmPrefactura_Load(object sender, System.EventArgs e)
		{
			bRastreo = true;		
			
			miAcceso = new Acceso(cdsSeteoF, "0704");
			
			if (!Funcion.Permiso("119", cdsSeteoF))
			{				
				MessageBox.Show("No puede ingresar a Prefactura", "Informacion");
				BeginInvoke(new MethodInvoker(UnloadMe));				
			}			
			else
			{
				if (miAcceso.Nuevo) this.btnNuevo.Enabled = true;
				if (miAcceso.Buscar) this.btnConsultar.Enabled = true;
				if (this.miAcceso.CambiarDescuento) this.btnClaveDescuento.Visible = true;			
								
				drSeteoF = cdsSeteoF.TableViews["SeteoF"].Rows[0];				
				bIncluyeIva = (bool) drSeteoF["IncluyeIva"];
							
				string sSQL = string.Format("Exec BodegaAsignadaPorUsuario {0}", MenuLatinium.IdUsuario);		
				this.txtBodegaPredef.Value = Funcion.iEscalarSQL(cdsSeteoF, sSQL);			
				
				string stAudita = "Exec AuditaCrear 53, 6, 'PREFACTURA'";
				Funcion.EjecutaSQL(cdsSeteoF, stAudita, true);
			
				FuncionesProcedimientos.EstadoGrids(false, this.ultraGrid2);
				
				this.txtIdTipoFactura.Value = 43;
				this.txtIdComprobante.Value = 15;
				this.txtIdFormaPago.Value = 9;
				this.txtNumContadoCredito.Value = 2;
				this.txtIdSucursal.Value = 1;
				this.dtFecha.CalendarInfo.MaxDate = DateTime.Today;
			}
		}

		#region Ultragrid
		private void ultraGrid1_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (!bEdicion) return;
			
			FuncionesProcedimientos.DesplazamientoGrids(sender, e, this.ultraGrid2);						
		}

		private void ultraGrid1_InitializeLayout(object sender, Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs e)
		{
			drSeteoF = cdsSeteoF.TableViews["SeteoF"].Rows[0];
			
			#region Iva
			if (!(bool)drSeteoF["CambiarIva"]) e.Layout.Bands[0].Columns["Impuesto"].CellActivation = Activation.Disabled;
			#endregion Iva
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
			stFormato = stMoneda + stCeros.Substring(0, (int) drSeteoF["decTotal"]);
			stInput = stMascara + stNumero.Substring(0, (int) drSeteoF["decTotal"]);
			e.Layout.Bands[0].Columns["SubTotal"].Format = stFormato; 
			e.Layout.Bands[0].Columns["DescuentoArt"].Format = stFormato;

			this.txtDescIvaTotal.FormatString = stFormato;
			this.txtIvaTotal.FormatString = stFormato;
			this.txtIva.FormatString = stFormato;
			this.txtTotal.FormatString = stFormato;
			
			this.txtDescIvaTotal.MaskInput = stInput;
			this.txtIvaTotal.MaskInput = stInput;
			this.txtIva.MaskInput = stInput;
			this.txtTotal.MaskInput = stInput;
			#endregion

			FuncionesProcedimientos.FormatoGrid(e, "Precio", 2);						
			FuncionesProcedimientos.FormatoGrid(e, "Impuesto", 2);						
			FuncionesProcedimientos.FormatoGrid(e, "DescuentoPorc", 2);	
		}
		
		private void ultraGrid1_BeforeRowsDeleted(object sender, Infragistics.Win.UltraWinGrid.BeforeRowsDeletedEventArgs e)
		{
			if (bEdicion)
			{			
				try
				{
					if (!bEliminaAlValidar)
					{
						//if (!FuncionesInventario.ValidaEliminaFilasGrids(sender, e, 1, this.grdSerial, cdsSeteoF)) return;
						if (!bEliminando) 
						{
							bEliminando = true;
							
							if (!Facturacion.FacturacionEliminaFilas(this.ultraGrid2, this.grdSerial, e, this.btnCancelar, 43, bNuevo, bImpreso, miAcceso.BEliArtCmbProm, miAcceso.BEliArt, (int)this.cmbBodega.Value, idTipoPrecio, 9, (decimal)this.txtTotal.Value,
								0, 0, 0, 0, (int)this.txtNumCuotas.Value, false, true, cdsSeteoF)) e.Cancel = true;
              
							e.DisplayPromptMsg = false;
							bEliminando = false;						
						}
						else e.DisplayPromptMsg = false;
					}
					else e.DisplayPromptMsg = false;
				}
				catch (Exception ex2)
				{
					MessageBox.Show(string.Format("Error al eliminar filas\n\nLa Pantalla se cerrara vuelva a abrirla para continuar", ex2.GetType(), ex2.Message), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);													
					this.btnCancelar_Click(sender, e);
					this.Close();
					e.DisplayPromptMsg = false;
				}
			}
		}

		private void ultraGrid1_AfterRowInsert(object sender, Infragistics.Win.UltraWinGrid.RowEventArgs e)
		{
			   if (bEdicion)
				Facturacion.IniciaValoresColumnas(sender, e, this.ultraGrid2, 43, (int)this.cmbBodega.Value);
		}	
		
		private void ultraGrid1_AfterCellUpdate(object sender, Infragistics.Win.UltraWinGrid.CellEventArgs e)
		{		
			if (bEdicion)
			{		
				try
				{
					if (bActualiza) 
					{
						bActualiza = false;
						int VerificaCambio = 0;

						#region Codigo
						if (e.Cell.Column.ToString() == "Codigo")
						{
							bActualiza = false;
							if (!bModificaColor)
								FuncionesInventario.FacturacionCargaArticulos(sender, e, this.ultraGrid2, 9, 4, (int)this.cmbBodega.Value, (int)this.cmbBodega1.Value, 0, 0, 0, 0, Convert.ToDateTime(this.dtFecha.Value), 43, true, true, true, true, (int)this.txtNumCuotas.Value, (int)this.txtidEntidadFinanciera.Value, (int)this.txtNumIdCliente.Value, 0, RegistroIdSerial, false, false, 0, cdsSeteoF, true, 0, false);
						
							bActualiza = true;
							this.cmdTotal();

							foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.ultraGrid2.Rows.All)
							{
								if ((int)dr.Cells["idCombo"].Value > 0 && ((int)dr.Cells["Tipo"].Value == 3 || (int)dr.Cells["Tipo"].Value == 4))
								{
									if (Funcion.bEscalarSQL(cdsSeteoF, string.Format("Exec PromocionRetornaSiAplicaCuotas {0}", (int)dr.Cells["idCombo"].Value), true)) this.txtNumCuotas.Enabled = false;
								}
							}
						}
						#endregion Codigo

						#region Cantidad
						if (e.Cell.Column.ToString() == "Cantidad")
						{
							bActualiza = false;
							FuncionesInventario.ValidaCamposGridEnteros(sender, e, "Cantidad");
						
							VerificaCambio = 1;
							bActualiza = true;

							this.cmdTotal();
						}
						#endregion Cantidad					

						this.cmdTotal();

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
	
							bActualiza = true;

							bActualiza = false;

							if (!Facturacion.bValidaModificacionPrecios(e, miAcceso.CambiarPrecio, 43, 4, 0, cdsSeteoF)) 
							{							
								cmdTotal();
								bActualiza = true;
								return;
							}

							bActualiza = true;
							cmdTotal();												
						}
						#endregion Precio

						this.cmdTotal();

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
						
							bActualiza = true;
						}
						#endregion IVA

						cmdTotal();

						#region Descuento Porcentaje
						if (e.Cell.Column.ToString() == "DescuentoPorc")
						{		
							bActualiza = false;
						
							if (!FuncionesInventario.ValidaCamposGridDecimal(sender, e, "DescuentoPorc"))
							{							
								cmdTotal();
								bActualiza = true;
								return;
							}

							bActualiza = true;

							bActualiza = false;
					
							if (!Facturacion.bValidaDescuentos(e, this.ultraGrid2, (int)this.cmbBodega.Value, bDescuentoActivo, dClavePorcDscto, 43, miAcceso.CambiarDescuento, 4, 0, 0, 0, 0, dPorcentajeCuponWeb, dDsctoReiterativo, false, 0, false, cdsSeteoF))
							{							
								cmdTotal();
								bActualiza = true;
								return;
							}						
						}

						if (e.Cell.Column.ToString() == "DescuentoArt")
						{						
							cmdTotal();
							bActualiza = true;
						}
						#endregion Descuento Porcentaje
					
						this.cmdTotal();

						bActualiza = true;					
					}
				}
				catch (Exception ex2)
				{
					MessageBox.Show(string.Format("Error al insertar filas\n\nLa Pantalla se cerrara vuelva a abrirla para continuar", ex2.GetType(), ex2.Message), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);													
					
					
					
					
					
					
					
					this.btnCancelar_Click(sender, e);
					this.Close();					
				}
			}


		}		
		#endregion Ultragrid

		private void CargaCliente()
		{
			string[] sParametros = new string[11];

			sParametros = FuncionesProcedimientos.RetornaParametrosClientesProveedores(false, 43, this.txtRuc.Text.ToString(), cdsSeteoF);
			
			if (sParametros[0] != null) this.txtNumIdCliente.Value = int.Parse(sParametros[0]);
			if (sParametros[1] != null) this.txtidTipoRuc.Value = int.Parse(sParametros[1]);
			if (sParametros[3] != null) this.txtRuc.Text = sParametros[3];
			if (sParametros[4] != null) this.txtNombre.Text = sParametros[4];	
			if (sParametros[11] != null) idInstitucion = int.Parse(sParametros[11]);

			dDsctoReiterativo = Funcion.decEscalarSQL(cdsSeteoF, string.Format("Select dbo.ValidaDescuentoGestionVentas({0})", (int)this.txtNumIdCliente.Value));
		}

		private void txtRuc_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.F3)
			{
				string sIdentificacion = FuncionesProcedimientos.AbreListaClientesProveedores(false, 43);
				if (sIdentificacion.Length > 0)
				{						
					this.txtRuc.Text = sIdentificacion;
					
					CargaCliente();

					this.txtNumCuotas.Focus();
				}
			}
		}

		private void txtRuc_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
			Funcion.CamposNumericos(sender, e);
			if (e.KeyChar == 13) this.txtNumCuotas.Focus();
		}

		private void txtRuc_Validating(object sender, System.ComponentModel.CancelEventArgs e)
		{
			if (bEdicion)			
			{
				if (!Validacion.vbIdentificacion(this.txtRuc, (int)this.txtidTipoRuc.Value, cdsSeteoF)) 
				{
					e.Cancel = true;
					return;
				}
								
				if (!Validacion.vbClienteExiste(this.txtRuc, false, 43, 9, cdsSeteoF)) 
				{
					e.Cancel = true;
					return;
				}
						
				if (Funcion.iEscalarSQL(cdsSeteoF, string.Format("Exec RetornaIdEFSolCredito '{0}'", this.txtRuc.Value.ToString())) == 0)
				{
					MessageBox.Show("Cliente le falta un ´Parametro\n\nComuniquese con Credito", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Information);
					this.txtRuc.Focus();
					e.Cancel = true;
				}
			}
		}

		private void txtRuc_Validated(object sender, System.EventArgs e)
		{
			if (bEdicion)			
			{
				this.txtNumIdCliente.Value = 0;
				this.txtidTipoRuc.Value = 0;
				this.txtNombre.Text = "";
				idInstitucion = 0;
				dDsctoReiterativo = 0.00m;

				CargaCliente();			
			
				#region Parametros de Credito
				if ((int)this.txtNumIdCliente.Value > 0)
				{
					DateTime dtFecha = Convert.ToDateTime(this.dtFecha.Value);
					string sSQLParametros = string.Format("Exec Cre_RetornaParametros {0}, '{1}'", (int)this.txtNumIdCliente.Value, dtFecha.ToString("yyyyMMdd"));
					SqlDataReader dr = Funcion.rEscalarSQL(cdsSeteoF, sSQLParametros, true);
			
					dr.Read();
					if (dr.HasRows)
					{
						if (bNuevo) this.txtidEntidadFinanciera.Value = dr.GetInt32(0);
						this.txtCuotaAsignada.Value = dr.GetDecimal(1);
						this.txtCuotaUsada.Value = dr.GetDecimal(2);
						this.txtCuotaDisponible.Value = dr.GetDecimal(3);
						this.txtInteres.Value = dr.GetDecimal(4);				
					}
					dr.Close();
			
					ValidaNumeroCuotas();

					CargaFecha();

					if (idInstitucion > 0) idTipoPrecio = 2; else if (idInstitucion == 0) idTipoPrecio = 4;
				}
				#endregion Parametros de Credito			
			}
		}

		private void txtNumCuotas_Enter(object sender, System.EventArgs e)
		{
			this.txtNumCuotas.SelectAll();
		}

		private void txtNumCuotas_Click(object sender, System.EventArgs e)
		{
			this.txtNumCuotas.SelectAll();
		}

		private void txtNumCuotas_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
			if (e.KeyChar == 8) e.Handled = true;			
			if (e.KeyChar == 13) this.txtCuotaDeEntrada.Focus();
		}

		private void txtNumCuotas_Validated(object sender, System.EventArgs e)
		{
			
		}

		private void txtNumCuotas_Validating(object sender, System.ComponentModel.CancelEventArgs e)
		{
			if (bEdicion)
				if (!Validacion.vbCampoEnteroVacio(this.txtNumCuotas, "Escriba el número de cuotas", (int)this.txtNumCuotas.MinValue, (int)this.txtNumCuotas.MaxValue)) e.Cancel = true;
		}

		private void txtCuotaDeEntrada_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
			if (e.KeyChar == 8) e.Handled = true;			
			if (e.KeyChar == 13)
			{
				this.cboFechaPrimerPago.DroppedDown = true;
				this.cboFechaPrimerPago.Focus();				
			}
		}

		private void txtCuotaDeEntrada_Click(object sender, System.EventArgs e)
		{
			this.txtCuotaDeEntrada.SelectAll();
		}

		private void txtCuotaDeEntrada_Enter(object sender, System.EventArgs e)
		{
			this.txtCuotaDeEntrada.SelectAll();
		}

		private void txtCuotaDeEntrada_Validated(object sender, System.EventArgs e)
		{
			if (bEdicion)	cmdTotal();
		}

		private void dtFecha_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
			if (e.KeyChar == 8) e.Handled = true;			
		}

		private void dtFecha_Validated(object sender, System.EventArgs e)
		{
			if (bEdicion)	
			{				
				if (this.cmbVendedor.ActiveRow != null) this.cmbVendedor.Value = System.DBNull.Value;
				if (this.cmbBodega.ActiveRow != null) 
					this.cmbVendedor.DataSource = FuncionesProcedimientos.Lista_Vendedores(Convert.ToDateTime(this.dtFecha.Value), (int)this.cmbBodega.Value, cdsSeteoF);

				CargaFecha();
			}
		}

		private void cboFechaPrimerPago_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
			if (e.KeyChar == 13) 
			{
				this.ultraGrid2.Focus();
				this.ultraGrid2.DisplayLayout.Bands[0].AddNew();
				this.ultraGrid2.ActiveCell = this.ultraGrid2.Rows[0].Cells["Codigo"];
			}
		}

		private void cmbVendedor_Validating(object sender, System.ComponentModel.CancelEventArgs e)
		{
			if (bEdicion)
				if (!Validacion.vbComboVacio(this.cmbVendedor, "Seleccione un Vendedor")) e.Cancel = true;
		}

		private void cboFechaPrimerPago_Validating(object sender, System.ComponentModel.CancelEventArgs e)
		{
			if (bEdicion && this.cboFechaPrimerPago.Text.ToString() == "")
			{
				MessageBox.Show("Seleccione la fecha de pago.", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Error);
				this.cboFechaPrimerPago.Focus();
				e.Cancel = true;				
			}		
		}

		private void txtNumCuotas_ValueChanged(object sender, System.EventArgs e)
		{
			if (bEdicion) cmdTotal();
		}

		private void dtFecha_Validating(object sender, System.ComponentModel.CancelEventArgs e)
		{
			if (bEdicion)
				if (!Validacion.vbFechaEnDocumentos(this.dtFecha, "Ingrese la Fecha", true, "Venta", cdsSeteoF)) e.Cancel = true;
		}

		private void frmPrefactura_Closing(object sender, System.ComponentModel.CancelEventArgs e)
		{
			if (bEdicion)
			{
				if (bEdicion)
				{
					MessageBox.Show("No puede salir mientras edita una PreFactura", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
					e.Cancel = true;
					return;
				}				
			}

			int iLong = this.Text.Trim().Length;
			if (iLong > 20) iLong = 20;

			string stAudita = "Exec AuditaCrear 53, 7, '" + this.Text.Trim().Substring(0, iLong) + "'";
			Funcion.EjecutaSQL(cdsSeteoF, stAudita, true);
		}

		private void frmPrefactura_Resize(object sender, System.EventArgs e)
		{
			this.ultraGrid2.DisplayLayout.AutoFitColumns = true;
		}

		private void frmPrefactura_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (!bEdicion)
			{
				if (e.KeyCode == Keys.N) if (miAcceso.Nuevo)	if (this.btnNuevo.Enabled) this.btnNuevo_Click(sender, e);
				if (e.KeyCode == Keys.B) if (miAcceso.Buscar) if (this.btnConsultar.Enabled) this.btnConsultar_Click(sender, e);
				if ((int)this.txtNumIdCompra.Value > 0)
				{
					if (e.KeyCode == Keys.E) if (miAcceso.Editar) if (this.btnEditar.Enabled) this.btnEditar_Click(sender, e);
					if (e.KeyCode == Keys.I) if (bImpreso) 
																	 {
																		 if (miAcceso.Imprimir) if (this.btnImprimir.Enabled) this.btnImprimir_Click(sender, e);
																	 }
																	 else if (miAcceso.ReImprimir) if (this.btnImprimir.Enabled) this.btnImprimir.Enabled = true;
					if (e.KeyCode == Keys.F5) 
					{
						miAcceso = new Acceso(cdsSeteoF, "0704");

						this.Consulta((int)this.txtNumIdCompra.Value);
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

		private void mnuAprobacionAprobar_Click(object sender, System.EventArgs e)
		{
			this.Consulta((int)this.txtNumIdCompra.Value);

			#region Verificaciones
			if (!bImpreso)
			{
				MessageBox.Show("No puede Aprobar si no ha impreso los documentos.", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				return;
			}
			
			if (bAnulado)
			{
				MessageBox.Show("No puede Aprobar una Prefactura Anulada.", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);				
				return;
			}

			if (Funcion.ValidaBloqueo((int)this.txtNumIdCompra.Value, 1, cdsSeteoF)) return;

			if (Funcion.ValidaPeriodo((DateTime)this.dtFecha.Value, cdsSeteoF, "VENTA"))
			{				
				MessageBox.Show(string.Format("No puede modificar documentos en esta fecha, El período '{0}' esta cerrado,\n\nConsulte al Administrador", Convert.ToDateTime(this.dtFecha.Value).ToString("MMMM").ToUpper()), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);									
				return;
			}

			if ((int) this.txtNumEstado.Value == 5)
			{
				MessageBox.Show("Este documento ya fue Aprobado.", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);				
				return;
			}
			if ((int) this.txtNumEstado.Value == 9)
			{
				MessageBox.Show("Este documento ya se Facturó.", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);				
				return;
			}

			#region Valida Creditos Anulados
			string sSQLCA = string.Format("Exec Cre_ValidaCreditosAnulados {0}", (int)this.txtNumIdCliente.Value);
			int iAnulados = Funcion.iEscalarSQL(cdsSeteoF, sSQLCA);
			bool bRevisado = false;
			if (iAnulados > 0)
			{
				if (DialogResult.OK == MessageBox.Show(string.Format("El Cliente tiene {0} Pre-Facturas ANULADAS/RECHAZADAS\n\nVerifique que el PAGARE y la TABLA DE AMORTIZACIÓN\n\nCORRESPONDAN AL NUEVO CRÉDITO", iAnulados), "Point Technology", MessageBoxButtons.OKCancel, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button2))
				{
					bRevisado = true;
				}
				else return;
			}			
			#endregion Valida Creditos Anulados
			#endregion Verificaciones			
							
			using (frmCre_CarpaAlmacen CA = new frmCre_CarpaAlmacen())
			{
				if (DialogResult.OK == CA.ShowDialog())
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

							#region Crea Factura
							string sSQLTransforma = string.Format("Exec CompraTransformaCopia {0}, 1, {1}, {2}, {3}, {4}, '{5}'",
								(int)this.txtNumIdCompra.Value, (int)this.cmbBodega.Value, (bool)CA.optTipos.Value, bRevisado, CA.chkVerificacion.Checked, CA.txtNotas.Text.ToString().Trim());
							oCmdActualiza.CommandText = sSQLTransforma;
							int idCompraDest = (int)oCmdActualiza.ExecuteScalar();						
							#endregion Crea Factura

							string sSQLD = string.Format("Exec MovimientoSerialPrefacturaRemates {0}", (int)this.txtNumIdCompra.Value);
							oCmdActualiza.CommandText = sSQLD;
							oCmdActualiza.ExecuteNonQuery();						
			
							#region Guarda la tabla de amortización de la factura
//							decimal dValor = Convert.ToDecimal(this.txtTotal.Value) - Convert.ToDecimal(this.txtCuotaDeEntrada.Value);
//							DateTime Fecha = Convert.ToDateTime(this.cboFechaPrimerPago.Text.ToString());
//							DateTime dtFecha = (DateTime)this.dtFecha.Value;
//
//							string sSQL = string.Format("Exec GuardaTablaDeAmortizacion {0}, {1}, '{2}', '{3}', {4}", 
//								dValor, (int)this.txtNumCuotas.Value, dtFecha.ToString("yyyyMMdd"), Fecha.ToString("yyyyMMdd"), idCompraDest);						
//							oCmdActualiza.CommandText = sSQL;
//							oCmdActualiza.ExecuteNonQuery();
							#endregion Guarda la tabla de amortización de la factura
							
							oTransaction.Commit();						
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

							this.Consulta((int)this.txtNumIdCompra.Value);
						}
					}
				}
			}			
		}

		private void mnuAprobacionRechazar_Click(object sender, System.EventArgs e)
		{
			this.Consulta((int)this.txtNumIdCompra.Value);

			if (bAnulado)
			{
				MessageBox.Show("No puede Aprobar una Prefactura Anulada.", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);				
				return;
			}

			if ((int) this.txtNumEstado.Value == 4)
			{
				MessageBox.Show("Este documento ya fue Rechazado.", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);				
				return;
			}
			
			if ((int) this.txtNumEstado.Value == 9)
			{
				MessageBox.Show("Este documento ya se Facturó.", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);				
				return;
			}
			
			if (Funcion.ValidaBloqueo((int)this.txtNumIdCompra.Value, 1, cdsSeteoF)) return;
			
			using (frmJustAnulaVenta Nota = new frmJustAnulaVenta())
			{
				if (DialogResult.OK == Nota.ShowDialog())
				{
					string sSQL = string.Format("Exec ActualizaEdoPedido {0}, 4, '', 0, ''", (int)this.txtNumIdCompra.Value);
					Funcion.EjecutaSQL(cdsSeteoF, sSQL, true);

					Funcion.EjecutaSQL(cdsSeteoF, string.Format("Update Compra Set Borrar = 1, Notas = '{1}' Where idOrigen = {0} And idTipoFactura = 1", (int)this.txtNumIdCompra.Value, Nota.txtNotas.Text.ToString()));

					Funcion.EjecutaSQL(cdsSeteoF, string.Format("Update Compra Set FechaRechazado = GETDATE(), Notas = '{1}' Where idCompra = {0}", (int)this.txtNumIdCompra.Value, Nota.txtNotas.Text.ToString()));						

					this.Consulta((int)this.txtNumIdCompra.Value);
				}
			}			
		}

		private void cmbVendedor_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
			if (e.KeyChar == 13) this.txtRuc.Focus();			
		}
		
		private void EliminaElementosCombo()
		{
			int Tipo = 1;

			foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.ultraGrid2.Rows.All)
			{
				if (bElCombo)
				{
					if (dr.Cells["Tipo"].Value != System.DBNull.Value) Tipo = (int)dr.Cells["Tipo"].Value;

					if (Tipo > 0)// && !(bool)dr.Cells["Principal"].Value) 
					{
						if (idComboElimina == (int)dr.Cells["idCombo"].Value) 
						{
							string sSQLBorraLinea = string.Format("Delete From DetCompra Where idDetCompra = {0}", (int)dr.Cells["idDetCompra"].Value);
							Funcion.EjecutaSQL(cdsSeteoF, sSQLBorraLinea, true);

							dr.Delete();
						}
					}
				}
			}
			
			bElCombo = false;
			idComboElimina = 0;

			this.ultraGrid2.DisplayLayout.Bands[0].AddNew();
			this.ultraGrid2.ActiveCell = this.ultraGrid2.Rows[0].Cells["Codigo"];			
		}

		private void ultraGrid1_AfterRowsDeleted(object sender, System.EventArgs e)
		{
			if (bEdicion)cmdTotal();		
		}

		private void btnSalir_Click(object sender, System.EventArgs e)
		{
			this.Close();
		}

		private void btnCancelar_Click(object sender, System.EventArgs e)
		{
			if (bEdicion && !bNuevo)
			{
				string sSQLBloqueo = string.Format("Exec RegistraBloqueoTransacciones {0}, 0, {1}, 1", idBloqueaTransacciones, (int)this.txtNumIdCompra.Value);
				Funcion.EjecutaSQL(cdsSeteoF, sSQLBloqueo, true);
			}
			
			if ((int)this.txtNumEstado.Value == 6)
			{
				if (Funcion.iEscalarSQL(cdsSeteoF, string.Format("Select Count(*) From DetCompra Where idCompra = {0}", (int)this.txtNumIdCompra.Value)) == 0)
				{
					MessageBox.Show("Se eliminaron los productos de la factura vuelva a cargarlos y Guarde la Factura, para poder cancelar", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
					return;
				}
			}

			#region Audita Descuento
			if (bDescuentoActivo)
			{
				string sMensaje = "";
					
				int iContador = 1;

				foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in ultraGrid2.Rows.All)
				{
					if (iContador > 1) sMensaje = " - ";

					sMensaje = sMensaje + " Codigo : " + dr.Cells["Codigo"].Value.ToString() + " % Dscto : " + dr.Cells["DescuentoPorc"].Value.ToString();

					iContador++;
				}

				Funcion.EjecutaSQL(cdsSeteoF, string.Format("Exec RegistraAuditoriaClaveDesCuento {0}, 'PREFACTURA', {1}, 4, '{2}'", MenuLatinium.IdUsuario, (int)this.txtNumIdCompra.Value, sMensaje));
			}
			#endregion Audita Descuento

			bEdicion = false;
			bNuevo = false;

			this.txtidEntidadFinanciera.Value = 0;
			this.cmbBodega.Value = System.DBNull.Value;
			this.cmbVendedor.Value = System.DBNull.Value;
			this.cmbBodega1.Value = System.DBNull.Value;
			this.txtNumero.Text = "";
			this.txtRuc.Text = "";
			this.txtNombre.Text = "";
			this.txtComprobante.Text = "";
			this.lblEstado.Text = "";
			this.lblUsuario.Text = "";
			this.txtNumCuotas.MinValue = 0;
			this.txtNumCuotas.Value = 0;
			this.txtCuotaDeEntrada.Value = 0;
			this.cboFechaPrimerPago.Text = "";
			this.cboFechaPrimerPago.Items.Clear();
			this.cmbVendedor.Value = System.DBNull.Value;
			this.txtCuotaAsignada.Value = 0.00m;
			this.txtCuotaDisponible.Value = 0.00m;
			this.txtCuotaUsada.Value = 0.00m;
			this.txtValorCuotas.Value = 0.00m;
			this.txtTotalCuotas.Value = 0.00m;
			this.txtIngresoSerial.Text = "";
			this.txtIngresoSerial.Enabled = false;
			this.chkAutorizar.Checked = false;
			this.chkVerificar.Checked = false;
			this.dtFecha.Value = System.DBNull.Value;
			this.txtNumEstado.Value = 0;
			this.txtNumIdCliente.Value = 0;
			this.txtNumIdCompra.Value = 0;
			this.txtNumIdProyecto.Value = 0;
			this.txtidTipoRuc.Value = 0;
			this.lblUsuario.Text = "";
			this.txtIva0.Value = 0.00m;
			this.txtIva.Value = 0.00m;
			this.txtDescIva0Total.Value = 0.00m;
			this.txtDescIvaTotal.Value = 0.00m;
			this.txtIvaTotal.Value = 0.00m;
			this.txtTotal.Value = 0.00m;
			this.txtidTipoRuc.Value = 0;
			this.txtInteres.Value = 0.00m;
			this.txtSerial.Text = "";			
			this.cmbBodega.Enabled = false;
			this.cmbVendedor.Enabled = false;			
			this.cmbBodega1.Enabled = false;
			this.txtNumero.Enabled = false;
			this.txtRuc.Enabled = false;
			this.dtFecha.Enabled = false;
			this.txtNumCuotas.Enabled = false;
			this.txtCuotaDeEntrada.Enabled = false;
			this.cboFechaPrimerPago.Enabled = false;
      			
			this.txtNotas.Enabled = false;
			this.txtObservaciones.Enabled = false;
			this.txtSerial.Enabled = false;
			this.chkAutorizar.Enabled = false;
			this.chkVerificar.Enabled = false;
			this.btnAprobacion.Enabled = false;			

			if (miAcceso.Nuevo) this.btnNuevo.Enabled = true;
			if (miAcceso.Buscar) this.btnConsultar.Enabled = true;
			this.btnEditar.Enabled = false;
			this.btnImprimir.Enabled = false;
			this.btnGuardar.Enabled = false;
			this.btnAnular.Enabled = false;
			this.btnCancelar.Enabled = false;
			this.btnAprobacion.Enabled = false;
			this.btnClaveDescuento.Enabled = false;
			bDescuentoActivo = false;
			FuncionesProcedimientos.EstadoGrids(false, this.ultraGrid2);
			FuncionesProcedimientos.EstadoGrids(false, this.grdSerial);
		
			this.ultraGrid2.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "Exec ConsultaDetalleFactura 0");
			this.grdSerial.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "Exec RefrescaSerialVenta 0");	

			this.txtNotas.Text = "";	
			this.txtObservaciones.Text = "";	
			
			if (idCuponWeb > 0) 
				Facturacion.ActualizaEstadoCuponWeb(idCuponWeb, 0, (int)this.txtNumIdCompra.Value, cdsSeteoF);

			#region Variables
			idInstitucion = 0;
			bImpreso = false;
			bEliminaAlValidar = false;			
			bPromocionCombo = false;
			bPromocion = false;
			bCombo = false;			
			bAnulado = false;
			bModificaColor = false;
			bDescuentoActivo = false;
			idTipoPrecio = 4;
			bArtPromocion = false;
			idBloqueaTransacciones = 0;
			dClavePorcDscto = 0.00m;
			iDiaMaximoDePago = 0;
			iMinimoDiasPago = 0;
			iMaximoDiasPago = 0;
			idCuponWeb = 0;
			dDsctoReiterativo = 0.00m;
			dPorcentajeCuponWeb = 0.00m;
			#endregion Variables
		}

		private void btnAnular_Click(object sender, System.EventArgs e)
		{
			this.Consulta((int)this.txtNumIdCompra.Value);

			if (!Facturacion.ValidaEdicionAnulacion((int)this.txtNumIdCompra.Value, this.ultraGrid2, miAcceso.Anular, (int)this.txtNumEstado.Value, 43, (int)this.cmbBodega.Value, 
				(DateTime)this.dtFecha.Value, true, bAnulado, 0, cdsSeteoF, false, false, false, false, "VENTA", 0, true)) return;
			
			if (DialogResult.Yes == MessageBox.Show("¿Esta seguro de anular esta Prefactura?.", "Point Technology", MessageBoxButtons.YesNo, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button2))
			{
				using (frmJustAnulaVenta Nota = new frmJustAnulaVenta())
				{
					if (DialogResult.OK == Nota.ShowDialog())
					{
						#region Anulacion	
						if (Facturacion.AnulacionCompraVenta ((int)this.txtNumIdCompra.Value, 43, cdsSeteoF))
							this.Consulta((int)this.txtNumIdCompra.Value);
						#endregion Anulacion	
					}
				}
			}			
		}

		private void btnEditar_Click(object sender, System.EventArgs e)
		{
			this.Consulta((int)this.txtNumIdCompra.Value);
			
			if (!Facturacion.ValidaEdicionAnulacion((int)this.txtNumIdCompra.Value, this.ultraGrid2, miAcceso.Editar, (int)this.txtNumEstado.Value, 43, (int)this.cmbBodega.Value, Convert.ToDateTime(this.dtFecha.Value), false, bAnulado, 0, cdsSeteoF, bImpreso, false, false, false, "VENTA", 0, true)) return;
			
			#region Habilita
			this.txtNumCuotas.Enabled = true;
			this.txtCuotaDeEntrada.Enabled = true;
			if (miAcceso.CambiarFecha) this.dtFecha.Enabled = true;
			if (miAcceso.BAutEntrada)
			{
				this.chkAutorizar.Enabled = true;
				this.chkVerificar.Enabled = true;
				this.txtObservaciones.Enabled = true;
			}

			if (miAcceso.BEditarNombre)	this.txtRuc.Enabled = true;

			this.cboFechaPrimerPago.Enabled = true;
			this.cmbVendedor.Enabled = true;

			this.txtNotas.Enabled = true;
			this.cmbVendedor.Enabled = true;

			//this.txtSerial.Enabled = true;

			this.btnNuevo.Enabled = false;
			this.btnConsultar.Enabled = false;
			this.btnEditar.Enabled = false;
			this.btnAnular.Enabled = false;
			this.btnImprimir.Enabled = false;
			this.btnGuardar.Enabled = true;
			this.btnAprobacion.Enabled = false;
			this.btnCancelar.Enabled = true;

			FuncionesProcedimientos.EstadoGrids(true, this.ultraGrid2);

			bNuevo = false;
			bEdicion = true;	

			if ((int)this.txtNumEstado.Value == 1 && this.miAcceso.CambiarDescuento) this.btnClaveDescuento.Enabled = true;

			if ((int)this.txtNumEstado.Value == 1)
			{
				string[] sPrmCW = new string[2];
				sPrmCW = Facturacion.sEditarCuponWeb((int)this.txtNumIdCompra.Value, cdsSeteoF);
				idCuponWeb = int.Parse(sPrmCW[0]);
				if (idCuponWeb > 0)
					dPorcentajeCuponWeb = decimal.Parse(sPrmCW[1]);
			}

			string sSQLBloqueo = string.Format("Exec RegistraBloqueoTransacciones 0, 1, {0}, 1", (int)this.txtNumIdCompra.Value);
			idBloqueaTransacciones = Funcion.iEscalarSQL(cdsSeteoF, sSQLBloqueo, true);
			#endregion Habilita 
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
				this.btnAprobacion.Enabled = false;
				this.btnClaveDescuento.Enabled = false;
				bDescuentoActivo = false;
				#endregion Botones

				string sSQL = string.Format("Exec ConsultaIndividualPreFactura {0}", idCompra);
				SqlDataReader drPreFactura = Funcion.rEscalarSQL(cdsSeteoF, sSQL, true);
				drPreFactura.Read();
				if (drPreFactura.HasRows)
				{
					#region Maestro
					this.txtNumIdCompra.Value = idCompra;
					this.txtIdSucursal.Value  = drPreFactura.GetInt32(1);
					this.txtIdTipoFactura.Value = drPreFactura.GetInt32(2);
					this.txtIdComprobante.Value  = drPreFactura.GetInt32(3);
					this.txtNumEstado.Value = drPreFactura.GetInt32(4);			

					this.txtNumContadoCredito.Value = drPreFactura.GetInt32(5);
					this.txtIdFormaPago.Value = drPreFactura.GetInt32(6);
					this.txtidEntidadFinanciera.Value = drPreFactura.GetInt32(7);
					this.cmbBodega.Value = drPreFactura.GetInt32(8);
					this.txtNumIdProyecto.Value = drPreFactura.GetInt32(9);
					this.cmbVendedor.Value = drPreFactura.GetInt32(10);
					this.txtNumero.Text = drPreFactura.GetValue(11).ToString();
					this.txtComprobante.Text = drPreFactura.GetValue(12).ToString();

					this.txtidTipoRuc.Value = drPreFactura.GetInt32(13);
					this.txtNumIdCliente.Value = drPreFactura.GetInt32(14);
					this.txtRuc.Text = drPreFactura.GetString(15);
					this.txtNombre.Text = drPreFactura.GetString(16);
					this.dtFecha.Value = drPreFactura.GetDateTime(17);
					this.txtNotas.Text = drPreFactura.GetString(18);

					bCrediPoint = drPreFactura.GetBoolean(19);
					this.txtTotalCuotas.Value = drPreFactura.GetDecimal(20);
					this.txtNumCuotas.Value = drPreFactura.GetInt32(21);
					this.txtCuotaDeEntrada.Value = drPreFactura.GetDecimal(22);
					if (drPreFactura.GetValue(23) != System.DBNull.Value) this.cboFechaPrimerPago.Text = drPreFactura.GetValue(23).ToString();

					this.txtIva0.Value = drPreFactura.GetDecimal(40);
					this.txtDescIva0Total.Value = drPreFactura.GetDecimal(41);
					this.txtIva.Value = drPreFactura.GetDecimal(24);
					this.txtDescIvaTotal.Value = drPreFactura.GetDecimal(25);
					this.txtIvaTotal.Value = drPreFactura.GetDecimal(26);
					this.txtTotal.Value = drPreFactura.GetDecimal(27);

					this.txtCuotaAsignada.Value = drPreFactura.GetDecimal(28);
					this.txtCuotaDisponible.Value = drPreFactura.GetDecimal(29);
					this.txtCuotaUsada.Value = drPreFactura.GetDecimal(30);
					this.txtInteres.Value = drPreFactura.GetDecimal(31);
					this.txtValorCuotas.Value = drPreFactura.GetDecimal(32);

					if (drPreFactura.GetValue(33) != System.DBNull.Value) this.chkAutorizar.Checked = drPreFactura.GetBoolean(33);
					bImpreso = drPreFactura.GetBoolean(34);
					bAnulado = drPreFactura.GetBoolean(35);
					this.lblUsuario.Text = drPreFactura.GetValue(36).ToString().ToUpper();
					this.chkVerificar.Checked = drPreFactura.GetBoolean(37);
					this.txtObservaciones.Text = drPreFactura.GetString(38);
					if (drPreFactura.GetValue(39) != System.DBNull.Value) this.cmbBodega1.Value = drPreFactura.GetInt32(39);
					#endregion Maestro 
				}
				drPreFactura.Close();

				if (bAnulado) this.lblEstado.Text = "ANULADO";
				else
				{
					if ((int)this.txtNumEstado.Value == 1) this.lblEstado.Text = "PENDIENTE";
					if ((int)this.txtNumEstado.Value == 2) this.lblEstado.Text = "VERIFICACIÓN";
					if ((int)this.txtNumEstado.Value == 4) this.lblEstado.Text = "RECHAZADO";
					if ((int)this.txtNumEstado.Value == 5) this.lblEstado.Text = "APROBADO";
					if ((int)this.txtNumEstado.Value == 9) this.lblEstado.Text = "FACTURADO";
				}

				this.ultraGrid2.DataSource = Funcion.dvProcedimiento(cdsSeteoF, string.Format("Exec ConsultaDetalleFactura {0}", (int)this.txtNumIdCompra.Value));					
				
				this.grdSerial.DataSource = Funcion.dvProcedimiento(cdsSeteoF, string.Format("Exec RefrescaSerialVenta {0}", (int)this.txtNumIdCompra.Value));

				this.cmbVendedor.DataSource = FuncionesProcedimientos.Lista_Vendedores(Convert.ToDateTime(this.dtFecha.Value), (int)this.cmbBodega.Value, cdsSeteoF);

				if (!bAnulado)
				{
					if (miAcceso.Editar && (int)this.txtNumEstado.Value == 1 && !bImpreso) this.btnEditar.Enabled = true;					
					if (!bImpreso) if (miAcceso.Imprimir) this.btnImprimir.Enabled = true;
					if (bImpreso) if (miAcceso.ReImprimir) this.btnImprimir.Enabled = true;
					if (miAcceso.Anular && !bAnulado && ((int)this.txtNumEstado.Value == 1 || (int)this.txtNumEstado.Value == 4)) this.btnAnular.Enabled = true;											
					if (miAcceso.BAprobarPrefactura) if ((int)this.txtNumEstado.Value == 1) this.btnAprobacion.Enabled = true;
					if (miAcceso.BRechazarPrefactura) if ((int)this.txtNumEstado.Value == 1 || (int)this.txtNumEstado.Value == 5)  this.mnuAprobacionRechazar.Enabled = true;				
					
			

			
					CargaFecha();
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
			btnCancelar_Click(sender, e);

			using (frmBuscaTransacciones Busqueda = new frmBuscaTransacciones((int)this.txtIdTipoFactura.Value, (int)this.txtBodegaPredef.Value))
			{
				if (DialogResult.OK == Busqueda.ShowDialog())
				{
					this.Consulta((int)Busqueda.grdFacturas.ActiveRow.Cells["idCompra"].Value);
				}
			}
		}

		private void btnNuevo_Click(object sender, System.EventArgs e)
		{			
			this.btnCancelar_Click(sender, e);									

			#region Habilita Controles
			if (miAcceso.BAutEntrada)
			{
				this.chkAutorizar.Enabled = true;
				this.txtObservaciones.Enabled = true;
				this.chkVerificar.Enabled = true;
			}
						
			this.btnNuevo.Enabled = false;
			this.btnConsultar.Enabled = false;
			this.btnEditar.Enabled = false;
			this.btnImprimir.Enabled = false;
			this.btnGuardar.Enabled = true;
			this.btnCancelar.Enabled = true;
		
			FuncionesProcedimientos.EstadoGrids(true, this.ultraGrid2);
			FuncionesProcedimientos.EstadoGrids(true, this.grdSerial);

			bNuevo = true;
			bEdicion = true;

			this.txtSerial.Enabled = true;
			this.txtNumEstado.Value = 1;
			this.txtidTipoRuc.Value	= 5;
			//FuncionesProcedimientos.FechasDocumentos(this.dtFecha);//this.dtFecha.Value = DateTime.Today;
			FuncionesProcedimientos.RetornaFechaServidor(this.dtFecha, cdsSeteoF, false);

			if (miAcceso.CambiarFecha) this.dtFecha.Enabled = true;

			if ((int)this.txtBodegaPredef.Value == 0) 
			{
				this.cmbBodega.Enabled = true;				
				this.cmbBodega.Value = 31;
			}

			if ((int)this.txtBodegaPredef.Value > 0) this.cmbBodega.Value = (int)this.txtBodegaPredef.Value;

			this.txtRuc.Enabled = true;
			this.txtNumCuotas.Enabled = true;
			this.txtCuotaDeEntrada.Enabled = true;
			this.cboFechaPrimerPago.Enabled = true;
			this.cmbVendedor.Enabled = true;
			this.txtNotas.Enabled = true;
			this.cmbBodega1.Enabled = true;			
			this.txtIngresoSerial.Enabled = true;	

			if (this.dtFecha.Enabled) this.dtFecha.Focus();
			else if ((int)this.txtBodegaPredef.Value == 0) this.cmbBodega.Focus();
			else this.cmbVendedor.Focus();
			#endregion Habilita Controles
		}

		private void cmbBodega_ValueChanged(object sender, System.EventArgs e)
		{
			if (bEdicion) 
			{
				if (this.cmbBodega.ActiveRow != null)
				{
					this.cmbBodega1.DataSource= Funcion.dvProcedimiento(cdsSeteoF, string.Format("Exec ListaBodegasPorLocal {0}, 1", (int)this.cmbBodega.Value));
					
					this.txtNumIdProyecto.Value = Funcion.iEscalarSQL(cdsSeteoF, string.Format("Select isNull(idProyecto, 0) From Bodega Where Bodega = {0}", (int)this.cmbBodega.Value));

					if (this.cmbVendedor.ActiveRow != null) this.cmbVendedor.Value = System.DBNull.Value;
					this.cmbVendedor.DataSource = FuncionesProcedimientos.Lista_Vendedores(Convert.ToDateTime(this.dtFecha.Value), (int)this.cmbBodega.Value, cdsSeteoF);

//					if (bNuevo)
//					{						
//						string[] sPrmCW = new string[2];
//						sPrmCW = Facturacion.sCuponWeb(cdsSeteoF);
//						idCuponWeb = int.Parse(sPrmCW[0]);
//						if (idCuponWeb > 0) dPorcentajeCuponWeb = decimal.Parse(sPrmCW[1]);
//					}
				}
			}				
		}

		private void btnImprimir_Click(object sender, System.EventArgs e)
		{							
			this.Consulta(Convert.ToInt32(this.txtNumIdCompra.Value));

			decimal dValor = Convert.ToDecimal(this.txtTotal.Value) - Convert.ToDecimal(this.txtCuotaDeEntrada.Value);
			string Fecha = this.cboFechaPrimerPago.Text;

			Facturacion.Impresion((int)this.txtNumIdCompra.Value, bAnulado, (int)this.txtNumEstado.Value, miAcceso.Imprimir, miAcceso.ReImprimir, true, bImpreso, 43, dValor, 
				(int)this.txtNumCuotas.Value, (int)this.txtNumIdCliente.Value, Fecha, (DateTime)this.dtFecha.Value, (int)this.txtidEntidadFinanciera.Value, this.txtRuc.Text.ToString(), cdsSeteoF);
		}

		private void btnGuardar_Click(object sender, System.EventArgs e)
		{
			#region Elimina Filas Vacias
			bEliminaAlValidar = true;
			Validacion.vEliminaFilasVacias(this.ultraGrid2, "idArticulo");
			bEliminaAlValidar = false;
			#endregion Elimina Filas Vacias

			#region Validaciones
			if (!Validacion.vbFechaEnDocumentos(this.dtFecha, "Ingrese la Fecha del Documento", true, "VENTA", cdsSeteoF)) return;

			if (!Validacion.vbComboVacio(this.cmbVendedor, "Seleccione un vendedor")) return;
			
			if (!Validacion.vbIdentificacion(this.txtRuc, 5, cdsSeteoF)) return;

			if (!Validacion.vbClienteExiste(this.txtRuc, false, 43, 9, cdsSeteoF)) return;
			
			if (!Validacion.vbCampoEnteroVacio(this.txtNumCuotas, "Escriba el número de cuotas", (int)this.txtNumCuotas.MinValue, (int)this.txtNumCuotas.MaxValue)) return;

			#region Fechas de pago
			if (this.cboFechaPrimerPago.Text.ToString() == "")
			{
				MessageBox.Show("Seleccione la fecha de pago.", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				this.cboFechaPrimerPago.Focus();
				return;
			}

			if (!Validacion.bFechas(this.cboFechaPrimerPago.Text))
			{
				MessageBox.Show("La Fecha de Primer Pago no es valida,\n\nUTILICE UNICAMENTE LAS FECHAS DE LA LISTA", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				this.cboFechaPrimerPago.Focus();
				return;
			}

			if (Convert.ToDateTime(this.cboFechaPrimerPago.Text).Day > iDiaMaximoDePago)
			{
				MessageBox.Show(string.Format("El dia maximo de pago es {0},\n\nUTILICE UNICAMENTE LAS FECHAS DE LA LISTA", iDiaMaximoDePago), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				this.cboFechaPrimerPago.Focus();
				return;
			}

			DiasDePago();

			int iDifDias = Funcion.iEscalarSQL(cdsSeteoF, string.Format("Select DATEDIFF(DAY, '{0}', '{1}')", Convert.ToDateTime(this.dtFecha.Value).ToString("yyyyMMdd"), Convert.ToDateTime(this.cboFechaPrimerPago.Text).ToString("yyyyMMdd")));
			if (iDifDias < iMinimoDiasPago || iDifDias > iMaximoDiasPago)
			{
				MessageBox.Show(string.Format("Entre la fecha de Primer Pago y la fecha del credito hay {0} dias,\n\nLa fecha de primer pago debe de estar entre {1} y {2} dias despues de la fecha de la pre-factura,\n\nUTILICE UNICAMENTE LAS FECHAS DE LA LISTA", iDifDias, iMinimoDiasPago, iMaximoDiasPago), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				this.cboFechaPrimerPago.Focus();
				return;
			}
			#endregion Fechas de pago			

			#region Valor de las Cuotas
			if (idInstitucion == 0)
			{
				if ((decimal) this.txtCuotaDisponible.Value < (decimal)this.txtValorCuotas.Value)
				{
					MessageBox.Show("El valor de la cuota es mayor al cupo disponible", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Error);
					return;
				}
			}
			#endregion Valor de las Cuotas

			#region Parametros de crédito
			if (idInstitucion == 0)
			{
				decimal TotalPref = Convert.ToDecimal(this.txtTotal.Value) - Convert.ToDecimal(this.txtCuotaDeEntrada.Value);
				decimal ValCuota = Convert.ToDecimal(this.txtValorCuotas.Value);
					
				string sSQLValCred = string.Format("Exec PointValidaCredito {0}, {1}, {2}, {3}, '', '{4}'", 
					TotalPref, (int)this.txtNumCuotas.Value, ValCuota, (int)this.txtidEntidadFinanciera.Value, Convert.ToDateTime(this.dtFecha.Value).ToString("yyyyMMdd"));
				string sMensaje = Funcion.sEscalarSQL(cdsSeteoF, sSQLValCred);

				if (sMensaje.Length > 0)
				{
					MessageBox.Show(sMensaje, "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
					return;
				}
			}
			#endregion Parametros de crédito			

			#region Valida Pre-Facturas Pendientes de Facturar
			if (Funcion.iEscalarSQL(cdsSeteoF, string.Format("Select Count(*) From Compra Where idTipoFactura = 43 And Borrar = 0 And Estado In (1, 5) And idCliente = {0} And Bodega = {1} And idCompra <> {2}", (int)this.txtNumIdCliente.Value, (int)this.cmbBodega.Value, (int)this.txtNumIdCompra.Value)) > 0)
			{
				MessageBox.Show("No puede Grabar hasta que haya Facturado las Prefacturas Pendientes", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);				
				return;
			}
			#endregion Valida Pre-Facturas Pendientes de Facturar

			#region Carpa
			if (this.chkVerificar.Checked)
			{
				if (this.txtObservaciones.Text.ToString().Trim().Length == 0)
				{
					MessageBox.Show("Para carpa debe escribir una Observación", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
					this.cmbVendedor.Focus();
					return;
				}
			}			
			#endregion Carpa			
						
			if (!Validacion.vbGrids(this.ultraGrid2, "Codigo", 43)) return;			
			#endregion Validaciones

			#region Seguro de desgravamen
			if ((int)this.txtidEntidadFinanciera.Value != 14) 
			{
				if (Funcion.bEscalarSQL(cdsSeteoF, "Select SeguroDesgravamen From SeteoF", true))
				{
					int iAños = Funcion.iEscalarSQL(cdsSeteoF, string.Format("Select DATEDIFF(YEAR, FechaNacimiento, '{0}') From Cre_DatosGenerales Where NumeroIdentificacion = '{1}'", 
						Convert.ToDateTime(this.dtFecha.Value).ToString("yyyyMMdd"), this.txtRuc.Text.ToString()));

					if (iAños < 71)
					{
						int iContSegDesg = 0;
						foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.ultraGrid2.Rows.All)
							if ((int)dr.Cells["idArticulo"].Value == 10191) iContSegDesg++;

						if (iContSegDesg == 0)
						{
							decimal dMonto = Convert.ToDecimal(this.txtTotal.Value) - Convert.ToDecimal(this.txtCuotaDeEntrada.Value);

							string sSQL = string.Format("Exec Cre_CalculaValorSeguroDesgravamen '{0}', {1}, {2}, {3}", 
								Convert.ToDateTime(this.dtFecha.Value).ToString("yyyyMMdd"), (int)this.txtidEntidadFinanciera.Value, dMonto, (int)this.txtNumCuotas.Value);
							decimal dValorSeguro = Funcion.decEscalarSQL(cdsSeteoF, sSQL);

							bActualiza = false;
							FuncionesInventario.FacturacionCargaSeguroDesgravamen(this.ultraGrid2, dValorSeguro, dPorcIva, 10191, false, cdsSeteoF);		
							bActualiza = true;

							cmdTotal();
						}
					}
				}
			} 
			#endregion Seguro de desgravamen
      
			cmdTotal();

			#region Grabar
			this.txtNumIdCompra.Value = 
				Facturacion.GrabaFacturacion((int)this.txtNumIdCompra.Value, 0, (int)this.txtIdSucursal.Value, 
				(int)this.txtIdTipoFactura.Value, (int)this.txtIdComprobante.Value, 
				(int)this.txtNumEstado.Value, false, (int)this.txtNumContadoCredito.Value, (int)this.txtIdFormaPago.Value,
				(int)this.cmbBodega.Value, (int)this.txtNumIdProyecto.Value,
				(int)this.cmbVendedor.Value, this.txtNumero.Text.ToString(), this.txtComprobante.Text.ToString(),
				(int)this.txtidTipoRuc.Value, (int)this.txtNumIdCliente.Value, this.txtRuc.Text.ToString(), this.txtNombre.Text.ToString(),
				(DateTime)this.dtFecha.Value, "", "", DateTime.Today, false, DateTime.Today, DateTime.Today, 0,
				this.txtNotas.Text.ToString(), true, Convert.ToDecimal(this.txtTotalCuotas.Value), (int)this.txtNumCuotas.Value, Convert.ToDecimal(this.txtCuotaDeEntrada.Value), DateTime.Parse(this.cboFechaPrimerPago.Text),
				Convert.ToDecimal(this.txtIva0.Value), Convert.ToDecimal(this.txtDescIva0Total.Value), Convert.ToDecimal(this.txtIva.Value), Convert.ToDecimal(this.txtDescIvaTotal.Value), Convert.ToDecimal(this.txtIvaTotal.Value), Convert.ToDecimal(this.txtTotal.Value),
				true, false, false, false, bAnulado, true, (int)this.cmbBodega1.Value, 0, this.txtObservaciones.Text.ToString().Trim(), 0.00m,
				Convert.ToDecimal(this.txtInteres.Value), Convert.ToDecimal(this.txtCuotaAsignada.Value), Convert.ToDecimal(this.txtCuotaDisponible.Value), Convert.ToDecimal(this.txtCuotaUsada.Value), 
				(bool)this.chkAutorizar.Checked, Convert.ToDecimal(this.txtValorCuotas.Value), (int)this.txtidEntidadFinanciera.Value, 0, 0, 0, 0, "", (bool)this.chkVerificar.Checked,
				1, "", "", 0, "", 0.00m, 0, "", 0.00m,
				bNuevo, true, idBloqueaTransacciones, this.ultraGrid2,
				true, 0, this.grdSerial, bDescuentoActivo);					
			#endregion Grabar
			
			if ((int)this.txtNumIdCompra.Value == 0) return;

			Facturacion.ActualizaEstadoCuponWeb(idCuponWeb, 2, (int)this.txtNumIdCompra.Value, cdsSeteoF);
			idCuponWeb = 0;

			#region Impresion
			if (!(bool) bImpreso)
			{
				decimal dValor = Convert.ToDecimal(this.txtTotal.Value) - Convert.ToDecimal(this.txtCuotaDeEntrada.Value);
				string Fecha = this.cboFechaPrimerPago.Text;
				DocsPrefactura Dctos = new DocsPrefactura ((int)this.txtNumIdCompra.Value, dValor, (int)this.txtNumCuotas.Value, (int)this.txtNumIdCliente.Value, Fecha, 
					(DateTime)this.dtFecha.Value, (int)this.txtidEntidadFinanciera.Value, this.txtRuc.Text.ToString());
				Dctos.ShowDialog();
			}
			#endregion Impresion
					
			#region Controles
			this.cmbBodega.Enabled = false;
			this.cmbVendedor.Enabled = false;
			this.txtNumero.Enabled = false;
			this.txtRuc.Enabled = false;
			this.dtFecha.Enabled = false;
			this.txtNumCuotas.Enabled = false;
			this.txtCuotaDeEntrada.Enabled = false;
			this.cboFechaPrimerPago.Enabled = false;
			this.chkVerificar.Enabled = false;
			this.txtNotas.Enabled = false;
			this.chkAutorizar.Enabled = false;
			this.txtObservaciones.Enabled = false;	
			this.txtSerial.Enabled = false; 
			this.txtIngresoSerial.Enabled = false; 
			
			FuncionesProcedimientos.EstadoGrids(false, this.ultraGrid2);

			bEliminaAlValidar = false;
			bModificaColor = false;

			bEdicion = false;
			bNuevo = false;
					
			bDescuentoActivo = false;
			dClavePorcDscto = 0.00m;

			this.Consulta((int)this.txtNumIdCompra.Value);
			#endregion Controles
		}

		private void btnAprobacion_Click(object sender, System.EventArgs e)
		{
			this.MnuAprobacion.Show(this.btnAprobacion, new Point(1, 20));
		}
				
		private void ultraGrid1_BeforeCellUpdate(object sender, Infragistics.Win.UltraWinGrid.BeforeCellUpdateEventArgs e)
		{		
			if (bEdicion)
			{
				if (e.Cell.Column.ToString() == "Codigo")
				{
					if (!Validacion.bValidaCambioArticuloVenta(e, bModificaColor, false))
					{
						e.Cancel = true;
						return;						
					}
				}
				
				if (e.Cell.Column.ToString() == "Precio")
					if (e.Cell.Row.Cells["Tipo"].Value != System.DBNull.Value)
						if ((int)e.Cell.Row.Cells["Tipo"].Value == 5) e.Cancel = true;				

				if (e.Cell.Column.ToString() == "DescuentoPorc")
					Validacion.ModificaDescuentos(e, false, false);
			}
		}
		
		private void btnClaveDescuento_Click(object sender, System.EventArgs e)
		{
			using (frmVentaClaveDescuento Clave = new frmVentaClaveDescuento())
			{				
				if (DialogResult.OK == Clave.ShowDialog())
				{
					#region Audita Descuento
					bDescuentoActivo = true;

					string sMensaje = "";
					
					int iContador = 1;

					foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in ultraGrid2.Rows.All)
					{
						if (iContador > 1) sMensaje = " - ";

						sMensaje = sMensaje + " Codigo : " + dr.Cells["Codigo"].Value.ToString() + " % Dscto : " + dr.Cells["DescuentoPorc"].Value.ToString();

						iContador++;
					}

					Funcion.EjecutaSQL(cdsSeteoF, string.Format("Exec RegistraAuditoriaClaveDesCuento {0}, 'PREFACTURA', {1}, 1, '{2}'", MenuLatinium.IdUsuario, (int)this.txtNumIdCompra.Value, sMensaje));
					
					dClavePorcDscto = Funcion.decEscalarSQL(cdsSeteoF, string.Format("Exec RetornaPorcentajeClaveDescuento {0}", MenuLatinium.IdUsuario));
					#endregion Audita Descuento
				}
				else dClavePorcDscto = 0.00m;
			}
		}

		private void dtFecha_ValueChanged(object sender, System.EventArgs e)
		{
			if (this.dtFecha.Value != System.DBNull.Value)
				FuncionesProcedimientos.ListaBodegas(this.cmbBodega, MenuLatinium.IdUsuario, 43, (DateTime)this.dtFecha.Value, false, cdsSeteoF);
		}

		private void dtFecha_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Delete) e.Handled = true;
			if (e.KeyCode == Keys.Enter) this.cmbBodega.Focus();
		}

		private void cmbBodega_Validating(object sender, System.ComponentModel.CancelEventArgs e)
		{
			if (bEdicion)
				if (!Validacion.vbComboVacio(this.cmbBodega, "Seleccione un local")) return;
		}
		
		private void txtNumCuotas_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Delete) e.Handled = true;
			if (e.KeyCode == Keys.Enter) this.txtCuotaDeEntrada.Focus();
		}

		private void txtCuotaDeEntrada_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Delete) e.Handled = true;
			if (e.KeyCode == Keys.Enter) this.cboFechaPrimerPago.Focus();
		}

		private void ultraGrid1_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
			if (ultraGrid2.ActiveCell.Column.Key.ToString().ToUpper() == "CANTIDAD")
				Funcion.CamposNumericos(sender, e);
		}
		
		private void txtIngresoSerial_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Enter)
			{
				#region Valida
				if (!Validacion.vbComboVacio(this.cmbBodega1, "Seleccione una bodega")) return;

				if (this.txtRuc.Text.ToString().Length == 0)
				{
					MessageBox.Show("Ingrese una cedula", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Information);
					this.txtRuc.Focus();
					return;
				}

				if (this.txtIngresoSerial.Text.ToString().Length == 0)
				{
					MessageBox.Show("Ingrese una serial", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Information);
					this.txtIngresoSerial.Focus();
					return;
				}
				#endregion Valida

				//FuncionesInventario.FacturacionCargaDYD(this.ultraGrid2, 0.00m, dPorcIva, cdsSeteoF);

				#region Valida serial existente en el grid
				foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in grdSerial.Rows.All)
				{
					string sSerial = dr.Cells["Serial"].Value.ToString();
					if (this.txtIngresoSerial.Text.ToString() == sSerial)
					{
						MessageBox.Show("Ya registro esta serial.", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Error);
						this.txtIngresoSerial.Text = "";
						this.txtIngresoSerial.Focus();
						return;
					}
				}
				#endregion Valida serial existente en el grid
        
				#region Variables
				int IdDetalleSerial = 0;
				bool SerialRepetida = false;
				int CantSerie = 0;
				int IPosicion = this.ultraGrid2.Rows.Count;
				decimal dPrecio = 0.00m;
				#endregion Variables

				#region idArticulo
				int IdArticulo = Funcion.iEscalarSQL(cdsSeteoF, string.Format("Exec RetornaIdArticuloCodigo '{0}', {1}", this.txtIngresoSerial.Text.ToString(), (int)this.cmbBodega1.Value));
				if (IdArticulo == 0)
				{
					MessageBox.Show("La serial o el Producto no existe,\n\nComuniquese con Auditoria", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Error);
					this.txtIngresoSerial.Focus();
					return;		
				}
				#endregion idArticulo

				#region maneja serie
				string stMensaje = string.Format("select isNull(ManejaSerial, 0) from Articulo where idArticulo={0}", IdArticulo);
				bool bManejaSerial = Funcion.bEscalarSQL(cdsSeteoF, stMensaje, true);
				#endregion maneja serie

				if (!bManejaSerial)
				{

					string sCodigo = Funcion.sEscalarSQL(cdsSeteoF, string.Format("Select Codigo From Articulo Where idArticulo = {0}", IdArticulo));
					this.ultraGrid2.Rows.Band.AddNew();
					this.ultraGrid2.Rows[IPosicion].Cells["Codigo"].Value = sCodigo;
					IPosicion++;

					this.cmbBodega.Enabled = false;
					this.cmbVendedor.Enabled = false;
					this.txtNumero.Enabled = false;
					this.txtRuc.Enabled = false;
					this.dtFecha.Enabled = false;
					this.txtNumCuotas.Enabled = false;
					this.txtCuotaDeEntrada.Enabled = false;
					//this.cboFechaPrimerPago.Enabled = false;

					this.txtIngresoSerial.Text = "";
					RegistroIdSerial = 0;
					this.cmbBodega1.Enabled = false;
				}
				else 
				{

				//int IdArticulo = Funcion.iEscalarSQL(cdsSeteoF, string.Format("Exec RetornaIdArticuloSeriales '{0}', {1}", this.txtIngresoSerial.Text.ToString(), (int)this.cmbBodega1.Value));
					string sCodigo = Funcion.sEscalarSQL(cdsSeteoF, string.Format("Select Codigo From Articulo Where idArticulo = {0}", IdArticulo));
								
					string sSQL = string.Format("Exec CuentaSerialArticulo '{0}', {1}, {2}", this.txtIngresoSerial.Text.ToString(), IdArticulo, (int)this.cmbBodega1.Value);
					int iCtaSerial = Funcion.iEscalarSQL(cdsSeteoF, sSQL);
        
					if (iCtaSerial == 1)
					{	
						#region Serie					
						RegistroIdSerial = 0;
						IdDetalleSerial = 0;
						string sSQLIdSerial = string.Format("Exec BuscaIdSerial '{0}', {1}, {2}", this.txtIngresoSerial.Text.ToString(), IdArticulo, (int)this.cmbBodega1.Value);
						SqlDataReader drIdSerial = Funcion.rEscalarSQL(cdsSeteoF, sSQLIdSerial, true);
						drIdSerial.Read();
						if (drIdSerial.HasRows)
						{
							RegistroIdSerial = drIdSerial.GetInt32(0);
							IdDetalleSerial = drIdSerial.GetInt32(1);
						}
						drIdSerial.Close();
						#endregion Serie		
						#region Articulo
						this.ultraGrid2.Rows.Band.AddNew();
						this.ultraGrid2.Rows[IPosicion].Cells["Codigo"].Value = sCodigo;
						IPosicion++;
						#endregion Articulo
						#region Detalle Seriales
						this.grdSerial.Rows.Band.AddNew();
						this.grdSerial.ActiveRow.Cells["idSerial"].Value =(int) RegistroIdSerial;
						this.grdSerial.ActiveRow.Cells["idDetalleSerial"].Value = 0;
						this.grdSerial.ActiveRow.Cells["idOrigen"].Value = IdDetalleSerial;
						this.grdSerial.ActiveRow.Cells["idArticulo"].Value = IdArticulo;
						this.grdSerial.ActiveRow.Cells["Serial"].Value = this.txtIngresoSerial.Text.ToString();	
						this.grdSerial.ActiveRow.Cells["Estado"].Value = 28;	
						this.grdSerial.ActiveRow.Cells["Posicion"].Value = IPosicion;

						this.cmbBodega.Enabled = false;
						this.cmbVendedor.Enabled = false;
						this.txtNumero.Enabled = false;
						this.txtRuc.Enabled = false;
						this.dtFecha.Enabled = false;
						this.txtNumCuotas.Enabled = false;
						this.txtCuotaDeEntrada.Enabled = false;
						//this.cboFechaPrimerPago.Enabled = false;

						this.txtIngresoSerial.Text = "";
						RegistroIdSerial = 0;
						#endregion Detalle Seriales

						this.cmbBodega1.Enabled = false;
					}
					else
					{					
						MessageBox.Show("La serial no existe, Esta en otro Local,\n\nComuniquese con el Administrador", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Error);
						this.txtIngresoSerial.Focus();
						return;				
					}
				}
			}
		
		}
		
		private void grdSerial_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			FuncionesProcedimientos.DesplazamientoGrids(sender, e, this.ultraGrid2);
		}
		
		private void cmbBodega_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Enter) this.cmbBodega1.Focus();
		}

		private void cmbBodega1_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Enter) this.cmbVendedor.Focus();
		}

		private void cboFechaPrimerPago_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Enter) this.txtIngresoSerial.Focus();
		}

		private void grdSerial_BeforeRowsDeleted(object sender, Infragistics.Win.UltraWinGrid.BeforeRowsDeletedEventArgs e)
		{
			e.DisplayPromptMsg = false;
		}

		private void cmbBodega_InitializeLayout(object sender, Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs e)
		{
		
		}

		private void ultraGrid2_DoubleClick(object sender, System.EventArgs e)
		{
			if (bEdicion)
			{		
				if (!FuncionesInventario.CambioColorSerialesIngresadas(this.ultraGrid2, this.grdSerial)) return;

				bModificaColor = true;
				Facturacion.CambiaColorProducto(this.ultraGrid2, 1, (int)this.cmbBodega.Value, cdsSeteoF);
				bModificaColor = false;												
			}
		}

		private void txtIngresoSerial_ValueChanged(object sender, System.EventArgs e)
		{
		
		}

		private void txtRuc_ValueChanged(object sender, System.EventArgs e)
		{
		
		}
	}
}