using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using DevExpress.XtraEditors;
using Infragistics.Win.UltraWinGrid;
using System.Data;
using System.Data.SqlClient;
using System.Globalization;
using CrystalDecisions.Shared;

namespace Latinium
{
	/// <summary>
	/// Summary description for frmDTIngresoPorRetiro.
	/// </summary>
	public class frmDTIngresoPorRetiro : DevExpress.XtraEditors.XtraForm
	{
		public Infragistics.Win.UltraWinGrid.UltraGrid uGridDTIngreso;
		private Infragistics.Win.UltraWinSchedule.UltraCalendarCombo dtFecha;
		private System.Windows.Forms.Label lblEstado;
		private System.Windows.Forms.Label label36;
		private System.Windows.Forms.Label label32;
		private System.Windows.Forms.Label label4;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbBodega;
		private System.Windows.Forms.Label lblNumero;
		private System.Windows.Forms.GroupBox groupBox1;
		private System.Windows.Forms.Label label8;
		private System.Windows.Forms.Label label6;
		private System.Windows.Forms.Label label5;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtObservaciones;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtCodigo;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtNombre;
		private System.Windows.Forms.Label label1;
		private System.Windows.Forms.Label label2;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtEstado;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtBodegaPredef;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtIdIngreso;
		private System.Windows.Forms.Button btnNuevo;
		private System.Windows.Forms.Button btnConsultar;
		private System.Windows.Forms.Button btnEditar;
		private System.Windows.Forms.Button btnImprimir;
		private System.Windows.Forms.Button btnGuardar;
		private System.Windows.Forms.Button btnCancelar;
		private System.Windows.Forms.Button btnAnular;
		private System.Windows.Forms.Button btnSalir;
		private System.Windows.Forms.GroupBox groupBox5;
		private System.Windows.Forms.Label label3;
		private System.Windows.Forms.Label label16;
		private System.Windows.Forms.Label label7;
		private System.Windows.Forms.Label label9;
		private System.Windows.Forms.Label label10;
		private System.Windows.Forms.Label label11;
		private System.Windows.Forms.Label label12;
		private System.Windows.Forms.Label label13;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource2;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtIdCompra;
		private System.Windows.Forms.Label label14;
		private Infragistics.Win.UltraWinSchedule.UltraCalendarCombo dtEmision;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtDiasDeMora;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtSaldo;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtPagado;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtTotalFactura;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtLocal;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtVendedor;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtTotalCredito;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtNumeroFactura;
		private Infragistics.Win.UltraWinSchedule.UltraCalendarCombo dtFechaRetiro;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtEntidadFinanciera;
		private System.Windows.Forms.ToolTip toolTip1;
		private System.Windows.Forms.GroupBox groupBox2;
		private C1.Data.C1DataSet cdsSeteoF;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource1;
		private Infragistics.Win.UltraWinGrid.UltraDropDown cmbEstados;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource3;
		private System.Windows.Forms.Label label15;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtValorNC;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtidTransferencia;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtIdNotaDeCredito;
		private System.Windows.Forms.Label label18;
		private Infragistics.Win.UltraWinEditors.UltraCheckEditor chkFiscal;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbMotivos;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource4;
		private System.Windows.Forms.Label lblSubTotal;
		private System.Windows.Forms.Label lblIva;
		private System.Windows.Forms.Label lblDescuento;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtIvaTotal;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtDescIvaTotal;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtIva;
		private System.Windows.Forms.Label lblVendedor;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbVendedor;
		private System.Windows.Forms.Label label19;
		private System.Windows.Forms.Label label20;
		private System.Windows.Forms.Label label21;
		private System.Windows.Forms.Button btnProcesar;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbRetiradoPor;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtTiempoUso;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtNLote;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtIdTransferenciaDT;
		private System.Windows.Forms.ContextMenu mnuImprimir;
		private System.Windows.Forms.MenuItem menuItem1;
		private System.Windows.Forms.MenuItem menuItem2;
		private System.Windows.Forms.MenuItem menuItem3;
		private System.Windows.Forms.MenuItem menuItem4;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtidTransferenciaRC;
		private System.Windows.Forms.MenuItem menuItem5;
		private System.Windows.Forms.Label label17;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtNumeroNotaDeCredito;
		private System.Windows.Forms.Label label22;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtProntoPago;
		private System.Windows.Forms.Label label23;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtDiferencia;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtIdNotaDeCredito1;
		private System.Windows.Forms.MenuItem menuItem6;
		private System.Windows.Forms.Label label24;
		private Infragistics.Win.UltraWinEditors.UltraCheckEditor chkReservaMercaderia;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtTiempoPago;
		private System.Windows.Forms.Label label25;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbBodegaActual;
		private System.ComponentModel.IContainer components;

		public frmDTIngresoPorRetiro()
		{
			//
			// Required for Windows Form Designer support
			//
			InitializeComponent();

			//
			// TODO: Add any constructor code after InitializeComponent call
			//
		}

		int idIngresoRetiro = 0;

		public frmDTIngresoPorRetiro(int IdIngresoRetiro)
		{
			//
			// Required for Windows Form Designer support
			//
			InitializeComponent();

			idIngresoRetiro = IdIngresoRetiro;

			//
			// TODO: Add any constructor code after InitializeComponent call
			//
		}

		/// <summary>
		/// Clean up any resources being used.
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
		/// Required method for Designer support - do not modify
		/// the contents of this method with the code editor.
		/// </summary>
		private void InitializeComponent()
		{
			this.components = new System.ComponentModel.Container();
			Infragistics.Win.Appearance appearance1 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand1 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn1 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idDetalleIngresoRetiro");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn2 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idIngresoRetiro");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn3 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idDetCompra");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn4 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCompra");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn5 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idSerial");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn6 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Serial");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn7 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idArticulo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn8 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Codigo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn9 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Articulo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn10 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Cantidad");
			Infragistics.Win.Appearance appearance2 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn11 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Precio");
			Infragistics.Win.Appearance appearance3 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn12 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("DescuentoArt");
			Infragistics.Win.Appearance appearance4 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn13 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("DescuentoPorc");
			Infragistics.Win.Appearance appearance5 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn14 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("SubTotal");
			Infragistics.Win.Appearance appearance6 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn15 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idDTEstado", -1, "cmbEstados");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn16 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Observaciones");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn17 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Sel");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn18 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idAccesorio1");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn19 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idAccesorio2");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn20 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idAccesorio3");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn21 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idAccesorio4");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn22 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idAccesorio5");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn23 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idAccesorio6");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn24 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idAccesorio7");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn25 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idAccesorio8");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn26 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idAccesorio9");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn27 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idAccesorio10");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn28 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idAccesorio11");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn29 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idAccesorio12");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn30 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idAccesorio13");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn31 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idAccesorio14");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn32 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idAccesorio15");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn33 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idAccesorio16");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn34 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idAccesorio17");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn35 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idAccesorio18");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn36 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idAccesorio19");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn37 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idAccesorio20");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn38 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idDetalleNotaDeCredito");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn39 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idArticuloMod");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn40 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idArticuloGE");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn41 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("CostoUltimo");
			Infragistics.Win.Appearance appearance7 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn42 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Cobrado");
			Infragistics.Win.Appearance appearance8 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn43 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("CostoFactura");
			Infragistics.Win.Appearance appearance9 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn44 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Cuotas_Pagadas");
			Infragistics.Win.Appearance appearance10 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn45 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Cuotas_Vencidas");
			Infragistics.Win.Appearance appearance11 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn46 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("DiasAtraso");
			Infragistics.Win.Appearance appearance12 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn47 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Bodega");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn48 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("UnboundColumn1", 0);
			Infragistics.Win.Appearance appearance13 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance14 = new Infragistics.Win.Appearance();
			System.Resources.ResourceManager resources = new System.Resources.ResourceManager(typeof(frmDTIngresoPorRetiro));
			Infragistics.Win.Appearance appearance15 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance16 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance17 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance18 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance19 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance20 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn1 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idDetalleIngresoRetiro");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn2 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idIngresoRetiro");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn3 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idDetCompra");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn4 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idCompra");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn5 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idSerial");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn6 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Serial");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn7 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idArticulo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn8 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Codigo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn9 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Articulo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn10 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Cantidad");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn11 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Precio");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn12 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("DescuentoArt");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn13 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("DescuentoPorc");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn14 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("SubTotal");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn15 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idDTEstado");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn16 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Observaciones");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn17 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Sel");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn18 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idAccesorio1");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn19 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idAccesorio2");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn20 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idAccesorio3");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn21 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idAccesorio4");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn22 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idAccesorio5");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn23 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idAccesorio6");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn24 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idAccesorio7");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn25 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idAccesorio8");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn26 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idAccesorio9");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn27 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idAccesorio10");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn28 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idAccesorio11");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn29 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idAccesorio12");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn30 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idAccesorio13");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn31 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idAccesorio14");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn32 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idAccesorio15");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn33 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idAccesorio16");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn34 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idAccesorio17");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn35 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idAccesorio18");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn36 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idAccesorio19");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn37 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idAccesorio20");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn38 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idDetalleNotaDeCredito");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn39 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idArticuloMod");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn40 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idArticuloGE");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn41 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("CostoUltimo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn42 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Cobrado");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn43 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("CostoFactura");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn44 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Cuotas_Pagadas");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn45 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Cuotas_Vencidas");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn46 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("DiasAtraso");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn47 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Bodega");
			Infragistics.Win.Appearance appearance21 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton dateButton1 = new Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton();
			Infragistics.Win.Appearance appearance22 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand2 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn49 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Bodega");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn50 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn51 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Codigo", 0);
			Infragistics.Win.Appearance appearance23 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance24 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance25 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance26 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton dateButton2 = new Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton();
			Infragistics.Win.Appearance appearance27 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance28 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance29 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance30 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance31 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance32 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance33 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton dateButton3 = new Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton();
			Infragistics.Win.Appearance appearance34 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance35 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance36 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance37 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn48 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idEntidadFinanciera");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn49 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Codigo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn50 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Nombre");
			Infragistics.Win.Appearance appearance38 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance39 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance40 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance41 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand3 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn52 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idDTEstado");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn53 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Descripcion");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn51 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idDTEstado");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn52 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Descripcion");
			Infragistics.Win.Appearance appearance42 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance43 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance44 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance45 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance46 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand4 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn54 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idMotivo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn55 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Motivo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn53 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idMotivo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn54 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Motivo");
			Infragistics.Win.Appearance appearance47 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance48 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance49 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance50 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand5 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn56 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idPersonal");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn57 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn58 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Codigo");
			Infragistics.Win.Appearance appearance51 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance52 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand6 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn59 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idPersonal");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn60 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn61 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Codigo");
			Infragistics.Win.Appearance appearance53 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance54 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance55 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance56 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance57 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance58 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance59 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance60 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance61 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand7 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn62 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Bodega");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn63 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn64 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Codigo", 0);
			this.uGridDTIngreso = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.ultraDataSource1 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.dtFecha = new Infragistics.Win.UltraWinSchedule.UltraCalendarCombo();
			this.lblEstado = new System.Windows.Forms.Label();
			this.label36 = new System.Windows.Forms.Label();
			this.label32 = new System.Windows.Forms.Label();
			this.label4 = new System.Windows.Forms.Label();
			this.cmbBodega = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.lblNumero = new System.Windows.Forms.Label();
			this.groupBox1 = new System.Windows.Forms.GroupBox();
			this.label8 = new System.Windows.Forms.Label();
			this.label6 = new System.Windows.Forms.Label();
			this.label5 = new System.Windows.Forms.Label();
			this.txtObservaciones = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.txtCodigo = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.txtNombre = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.dtEmision = new Infragistics.Win.UltraWinSchedule.UltraCalendarCombo();
			this.label1 = new System.Windows.Forms.Label();
			this.label2 = new System.Windows.Forms.Label();
			this.txtNumeroFactura = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.txtEstado = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtBodegaPredef = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtIdIngreso = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.btnNuevo = new System.Windows.Forms.Button();
			this.btnConsultar = new System.Windows.Forms.Button();
			this.btnEditar = new System.Windows.Forms.Button();
			this.btnImprimir = new System.Windows.Forms.Button();
			this.btnGuardar = new System.Windows.Forms.Button();
			this.btnCancelar = new System.Windows.Forms.Button();
			this.btnAnular = new System.Windows.Forms.Button();
			this.btnSalir = new System.Windows.Forms.Button();
			this.groupBox5 = new System.Windows.Forms.GroupBox();
			this.txtDiasDeMora = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.label3 = new System.Windows.Forms.Label();
			this.txtSaldo = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.label16 = new System.Windows.Forms.Label();
			this.dtFechaRetiro = new Infragistics.Win.UltraWinSchedule.UltraCalendarCombo();
			this.label7 = new System.Windows.Forms.Label();
			this.txtPagado = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.label9 = new System.Windows.Forms.Label();
			this.txtTotalFactura = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.label10 = new System.Windows.Forms.Label();
			this.label11 = new System.Windows.Forms.Label();
			this.label12 = new System.Windows.Forms.Label();
			this.txtLocal = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.txtVendedor = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.label13 = new System.Windows.Forms.Label();
			this.ultraDataSource2 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.txtIdCompra = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtTotalCredito = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.label14 = new System.Windows.Forms.Label();
			this.txtEntidadFinanciera = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.toolTip1 = new System.Windows.Forms.ToolTip(this.components);
			this.txtNumeroNotaDeCredito = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.groupBox2 = new System.Windows.Forms.GroupBox();
			this.cdsSeteoF = new C1.Data.C1DataSet();
			this.cmbEstados = new Infragistics.Win.UltraWinGrid.UltraDropDown();
			this.ultraDataSource3 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.txtValorNC = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.label15 = new System.Windows.Forms.Label();
			this.txtidTransferencia = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtIdNotaDeCredito = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.label18 = new System.Windows.Forms.Label();
			this.chkFiscal = new Infragistics.Win.UltraWinEditors.UltraCheckEditor();
			this.cmbMotivos = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.ultraDataSource4 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.lblSubTotal = new System.Windows.Forms.Label();
			this.lblIva = new System.Windows.Forms.Label();
			this.lblDescuento = new System.Windows.Forms.Label();
			this.txtIvaTotal = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtDescIvaTotal = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtIva = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.lblVendedor = new System.Windows.Forms.Label();
			this.cmbVendedor = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.label19 = new System.Windows.Forms.Label();
			this.txtNLote = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.label20 = new System.Windows.Forms.Label();
			this.cmbRetiradoPor = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.label21 = new System.Windows.Forms.Label();
			this.txtTiempoUso = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.btnProcesar = new System.Windows.Forms.Button();
			this.txtIdTransferenciaDT = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.mnuImprimir = new System.Windows.Forms.ContextMenu();
			this.menuItem1 = new System.Windows.Forms.MenuItem();
			this.menuItem6 = new System.Windows.Forms.MenuItem();
			this.menuItem2 = new System.Windows.Forms.MenuItem();
			this.menuItem3 = new System.Windows.Forms.MenuItem();
			this.menuItem5 = new System.Windows.Forms.MenuItem();
			this.menuItem4 = new System.Windows.Forms.MenuItem();
			this.txtidTransferenciaRC = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.label17 = new System.Windows.Forms.Label();
			this.label22 = new System.Windows.Forms.Label();
			this.txtProntoPago = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.label23 = new System.Windows.Forms.Label();
			this.txtDiferencia = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtIdNotaDeCredito1 = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.label24 = new System.Windows.Forms.Label();
			this.chkReservaMercaderia = new Infragistics.Win.UltraWinEditors.UltraCheckEditor();
			this.txtTiempoPago = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.label25 = new System.Windows.Forms.Label();
			this.cmbBodegaActual = new Infragistics.Win.UltraWinGrid.UltraCombo();
			((System.ComponentModel.ISupportInitialize)(this.uGridDTIngreso)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource1)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.dtFecha)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbBodega)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtObservaciones)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtCodigo)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNombre)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.dtEmision)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNumeroFactura)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtEstado)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtBodegaPredef)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIdIngreso)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtDiasDeMora)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtSaldo)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.dtFechaRetiro)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtPagado)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtTotalFactura)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtLocal)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtVendedor)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource2)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIdCompra)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtTotalCredito)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtEntidadFinanciera)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNumeroNotaDeCredito)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbEstados)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource3)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtValorNC)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtidTransferencia)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIdNotaDeCredito)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbMotivos)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource4)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIvaTotal)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtDescIvaTotal)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIva)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbVendedor)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNLote)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbRetiradoPor)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtTiempoUso)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIdTransferenciaDT)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtidTransferenciaRC)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtProntoPago)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtDiferencia)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIdNotaDeCredito1)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtTiempoPago)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbBodegaActual)).BeginInit();
			this.SuspendLayout();
			// 
			// uGridDTIngreso
			// 
			this.uGridDTIngreso.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.uGridDTIngreso.Cursor = System.Windows.Forms.Cursors.Default;
			this.uGridDTIngreso.DataSource = this.ultraDataSource1;
			appearance1.BackColor = System.Drawing.Color.White;
			appearance1.ForeColor = System.Drawing.Color.Black;
			appearance1.ForeColorDisabled = System.Drawing.Color.Black;
			this.uGridDTIngreso.DisplayLayout.Appearance = appearance1;
			this.uGridDTIngreso.DisplayLayout.AutoFitColumns = true;
			ultraGridColumn1.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn1.Header.VisiblePosition = 0;
			ultraGridColumn1.Hidden = true;
			ultraGridColumn2.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn2.Header.VisiblePosition = 1;
			ultraGridColumn2.Hidden = true;
			ultraGridColumn3.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn3.Header.VisiblePosition = 3;
			ultraGridColumn3.Hidden = true;
			ultraGridColumn4.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn4.Header.VisiblePosition = 2;
			ultraGridColumn4.Hidden = true;
			ultraGridColumn5.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn5.Header.VisiblePosition = 4;
			ultraGridColumn5.Hidden = true;
			ultraGridColumn6.Case = Infragistics.Win.UltraWinGrid.Case.Upper;
			ultraGridColumn6.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn6.Header.VisiblePosition = 5;
			ultraGridColumn6.Width = 125;
			ultraGridColumn7.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn7.Header.VisiblePosition = 6;
			ultraGridColumn7.Hidden = true;
			ultraGridColumn8.Case = Infragistics.Win.UltraWinGrid.Case.Upper;
			ultraGridColumn8.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn8.Header.VisiblePosition = 7;
			ultraGridColumn8.Width = 184;
			ultraGridColumn9.Case = Infragistics.Win.UltraWinGrid.Case.Upper;
			ultraGridColumn9.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn9.Header.VisiblePosition = 8;
			ultraGridColumn9.Width = 219;
			ultraGridColumn10.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance2.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn10.CellAppearance = appearance2;
			ultraGridColumn10.Header.Caption = "Cant";
			ultraGridColumn10.Header.VisiblePosition = 10;
			ultraGridColumn10.Width = 35;
			ultraGridColumn11.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance3.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn11.CellAppearance = appearance3;
			ultraGridColumn11.Format = "#,##0.00";
			ultraGridColumn11.Header.VisiblePosition = 11;
			ultraGridColumn11.Width = 70;
			ultraGridColumn12.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance4.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn12.CellAppearance = appearance4;
			ultraGridColumn12.Format = "#,##0.00";
			ultraGridColumn12.Header.VisiblePosition = 12;
			ultraGridColumn12.Hidden = true;
			ultraGridColumn12.Width = 70;
			ultraGridColumn13.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance5.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn13.CellAppearance = appearance5;
			ultraGridColumn13.Format = "#,##0.00";
			ultraGridColumn13.Header.Caption = "% Dscto";
			ultraGridColumn13.Header.VisiblePosition = 13;
			ultraGridColumn13.Width = 49;
			ultraGridColumn14.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance6.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn14.CellAppearance = appearance6;
			ultraGridColumn14.Format = "#,##0.00";
			ultraGridColumn14.Header.Caption = "Subtotal";
			ultraGridColumn14.Header.VisiblePosition = 14;
			ultraGridColumn14.Width = 70;
			ultraGridColumn15.ButtonDisplayStyle = Infragistics.Win.UltraWinGrid.ButtonDisplayStyle.Always;
			ultraGridColumn15.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn15.Header.Caption = "Estado";
			ultraGridColumn15.Header.VisiblePosition = 9;
			ultraGridColumn15.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.DropDown;
			ultraGridColumn15.Width = 79;
			ultraGridColumn16.Case = Infragistics.Win.UltraWinGrid.Case.Upper;
			ultraGridColumn16.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn16.Header.VisiblePosition = 15;
			ultraGridColumn16.Width = 270;
			ultraGridColumn17.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn17.Header.Caption = "DT";
			ultraGridColumn17.Header.VisiblePosition = 16;
			ultraGridColumn17.Hidden = true;
			ultraGridColumn17.Width = 18;
			ultraGridColumn18.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn18.Header.VisiblePosition = 17;
			ultraGridColumn18.Hidden = true;
			ultraGridColumn19.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn19.Header.VisiblePosition = 19;
			ultraGridColumn19.Hidden = true;
			ultraGridColumn20.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn20.Header.VisiblePosition = 20;
			ultraGridColumn20.Hidden = true;
			ultraGridColumn21.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn21.Header.VisiblePosition = 21;
			ultraGridColumn21.Hidden = true;
			ultraGridColumn22.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn22.Header.VisiblePosition = 22;
			ultraGridColumn22.Hidden = true;
			ultraGridColumn23.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn23.Header.VisiblePosition = 23;
			ultraGridColumn23.Hidden = true;
			ultraGridColumn24.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn24.Header.VisiblePosition = 24;
			ultraGridColumn24.Hidden = true;
			ultraGridColumn25.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn25.Header.VisiblePosition = 25;
			ultraGridColumn25.Hidden = true;
			ultraGridColumn26.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn26.Header.VisiblePosition = 26;
			ultraGridColumn26.Hidden = true;
			ultraGridColumn27.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn27.Header.VisiblePosition = 27;
			ultraGridColumn27.Hidden = true;
			ultraGridColumn28.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn28.Header.VisiblePosition = 28;
			ultraGridColumn28.Hidden = true;
			ultraGridColumn29.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn29.Header.VisiblePosition = 29;
			ultraGridColumn29.Hidden = true;
			ultraGridColumn30.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn30.Header.VisiblePosition = 30;
			ultraGridColumn30.Hidden = true;
			ultraGridColumn31.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn31.Header.VisiblePosition = 31;
			ultraGridColumn31.Hidden = true;
			ultraGridColumn32.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn32.Header.VisiblePosition = 32;
			ultraGridColumn32.Hidden = true;
			ultraGridColumn33.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn33.Header.VisiblePosition = 33;
			ultraGridColumn33.Hidden = true;
			ultraGridColumn34.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn34.Header.VisiblePosition = 34;
			ultraGridColumn34.Hidden = true;
			ultraGridColumn35.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn35.Header.VisiblePosition = 35;
			ultraGridColumn35.Hidden = true;
			ultraGridColumn36.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn36.Header.VisiblePosition = 36;
			ultraGridColumn36.Hidden = true;
			ultraGridColumn37.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn37.Header.VisiblePosition = 37;
			ultraGridColumn37.Hidden = true;
			ultraGridColumn38.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn38.Header.VisiblePosition = 38;
			ultraGridColumn38.Hidden = true;
			ultraGridColumn39.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn39.Header.VisiblePosition = 39;
			ultraGridColumn39.Hidden = true;
			ultraGridColumn40.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn40.Header.VisiblePosition = 40;
			ultraGridColumn40.Hidden = true;
			ultraGridColumn41.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance7.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn41.CellAppearance = appearance7;
			ultraGridColumn41.Format = "#,##0.00";
			ultraGridColumn41.Header.VisiblePosition = 41;
			ultraGridColumn41.Hidden = true;
			ultraGridColumn42.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance8.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn42.CellAppearance = appearance8;
			ultraGridColumn42.Format = "#,##0.00";
			ultraGridColumn42.Header.VisiblePosition = 42;
			ultraGridColumn42.Hidden = true;
			ultraGridColumn43.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance9.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn43.CellAppearance = appearance9;
			ultraGridColumn43.Format = "#,##0.00";
			ultraGridColumn43.Header.VisiblePosition = 43;
			ultraGridColumn43.Hidden = true;
			ultraGridColumn44.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance10.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn44.CellAppearance = appearance10;
			ultraGridColumn44.Format = "#,##0";
			ultraGridColumn44.Header.VisiblePosition = 44;
			ultraGridColumn44.Hidden = true;
			ultraGridColumn45.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance11.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn45.CellAppearance = appearance11;
			ultraGridColumn45.Format = "#,##0";
			ultraGridColumn45.Header.VisiblePosition = 45;
			ultraGridColumn45.Hidden = true;
			ultraGridColumn46.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance12.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn46.CellAppearance = appearance12;
			ultraGridColumn46.Format = "#,##0";
			ultraGridColumn46.Header.VisiblePosition = 46;
			ultraGridColumn46.Hidden = true;
			ultraGridColumn47.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn47.Header.VisiblePosition = 47;
			ultraGridColumn47.Hidden = true;
			ultraGridColumn47.Width = 51;
			ultraGridColumn48.ButtonDisplayStyle = Infragistics.Win.UltraWinGrid.ButtonDisplayStyle.Always;
			appearance13.ImageHAlign = Infragistics.Win.HAlign.Center;
			appearance13.ImageVAlign = Infragistics.Win.VAlign.Middle;
			ultraGridColumn48.CellAppearance = appearance13;
			appearance14.Image = ((object)(resources.GetObject("appearance14.Image")));
			appearance14.ImageHAlign = Infragistics.Win.HAlign.Center;
			appearance14.ImageVAlign = Infragistics.Win.VAlign.Middle;
			appearance14.TextHAlign = Infragistics.Win.HAlign.Center;
			appearance14.TextVAlign = Infragistics.Win.VAlign.Middle;
			ultraGridColumn48.CellButtonAppearance = appearance14;
			ultraGridColumn48.Header.Caption = "A";
			ultraGridColumn48.Header.VisiblePosition = 18;
			ultraGridColumn48.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.Button;
			ultraGridColumn48.Width = 28;
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
																										 ultraGridColumn48});
			this.uGridDTIngreso.DisplayLayout.BandsSerializer.Add(ultraGridBand1);
			appearance15.ForeColor = System.Drawing.Color.Black;
			appearance15.ForeColorDisabled = System.Drawing.Color.Black;
			this.uGridDTIngreso.DisplayLayout.Override.ActiveRowAppearance = appearance15;
			this.uGridDTIngreso.DisplayLayout.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.No;
			this.uGridDTIngreso.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.False;
			this.uGridDTIngreso.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.True;
			appearance16.BackColor = System.Drawing.Color.Transparent;
			this.uGridDTIngreso.DisplayLayout.Override.CardAreaAppearance = appearance16;
			appearance17.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance17.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance17.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance17.FontData.BoldAsString = "True";
			appearance17.FontData.Name = "Arial";
			appearance17.FontData.SizeInPoints = 7.75F;
			appearance17.ForeColor = System.Drawing.Color.White;
			appearance17.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.uGridDTIngreso.DisplayLayout.Override.HeaderAppearance = appearance17;
			appearance18.BackColor = System.Drawing.SystemColors.ControlLight;
			appearance18.BackColor2 = System.Drawing.Color.Gainsboro;
			appearance18.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridDTIngreso.DisplayLayout.Override.RowAlternateAppearance = appearance18;
			appearance19.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance19.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance19.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridDTIngreso.DisplayLayout.Override.RowSelectorAppearance = appearance19;
			appearance20.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance20.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance20.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridDTIngreso.DisplayLayout.Override.SelectedRowAppearance = appearance20;
			this.uGridDTIngreso.Font = new System.Drawing.Font("Tahoma", 7F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.uGridDTIngreso.Location = new System.Drawing.Point(8, 232);
			this.uGridDTIngreso.Name = "uGridDTIngreso";
			this.uGridDTIngreso.Size = new System.Drawing.Size(1150, 144);
			this.uGridDTIngreso.TabIndex = 4;
			this.uGridDTIngreso.AfterRowsDeleted += new System.EventHandler(this.uGridDTIngreso_AfterRowsDeleted);
			this.uGridDTIngreso.AfterRowInsert += new Infragistics.Win.UltraWinGrid.RowEventHandler(this.uGridDTIngreso_AfterRowInsert);
			this.uGridDTIngreso.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.uGridDTIngreso_KeyPress);
			this.uGridDTIngreso.BeforeCellUpdate += new Infragistics.Win.UltraWinGrid.BeforeCellUpdateEventHandler(this.uGridDTIngreso_BeforeCellUpdate);
			this.uGridDTIngreso.BeforeRowsDeleted += new Infragistics.Win.UltraWinGrid.BeforeRowsDeletedEventHandler(this.uGridDTIngreso_BeforeRowsDeleted);
			this.uGridDTIngreso.KeyDown += new System.Windows.Forms.KeyEventHandler(this.uGridDTIngreso_KeyDown);
			this.uGridDTIngreso.ClickCellButton += new Infragistics.Win.UltraWinGrid.CellEventHandler(this.uGridDTIngreso_ClickCellButton);
			this.uGridDTIngreso.AfterCellUpdate += new Infragistics.Win.UltraWinGrid.CellEventHandler(this.uGridDTIngreso_AfterCellUpdate);
			this.uGridDTIngreso.InitializeLayout += new Infragistics.Win.UltraWinGrid.InitializeLayoutEventHandler(this.uGridDTIngreso_InitializeLayout);
			// 
			// ultraDataSource1
			// 
			ultraDataColumn1.DataType = typeof(int);
			ultraDataColumn2.DataType = typeof(int);
			ultraDataColumn3.DataType = typeof(int);
			ultraDataColumn4.DataType = typeof(int);
			ultraDataColumn5.DataType = typeof(int);
			ultraDataColumn7.DataType = typeof(int);
			ultraDataColumn10.DataType = typeof(int);
			ultraDataColumn10.DefaultValue = 0;
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
			ultraDataColumn15.DataType = typeof(int);
			ultraDataColumn17.DataType = typeof(bool);
			ultraDataColumn17.DefaultValue = false;
			ultraDataColumn18.DataType = typeof(int);
			ultraDataColumn19.DataType = typeof(int);
			ultraDataColumn20.DataType = typeof(int);
			ultraDataColumn21.DataType = typeof(int);
			ultraDataColumn22.DataType = typeof(int);
			ultraDataColumn23.DataType = typeof(int);
			ultraDataColumn24.DataType = typeof(int);
			ultraDataColumn25.DataType = typeof(int);
			ultraDataColumn26.DataType = typeof(int);
			ultraDataColumn27.DataType = typeof(int);
			ultraDataColumn28.DataType = typeof(int);
			ultraDataColumn29.DataType = typeof(int);
			ultraDataColumn30.DataType = typeof(int);
			ultraDataColumn31.DataType = typeof(int);
			ultraDataColumn32.DataType = typeof(int);
			ultraDataColumn33.DataType = typeof(int);
			ultraDataColumn34.DataType = typeof(int);
			ultraDataColumn35.DataType = typeof(int);
			ultraDataColumn36.DataType = typeof(int);
			ultraDataColumn37.DataType = typeof(int);
			ultraDataColumn38.DataType = typeof(int);
			ultraDataColumn39.DataType = typeof(int);
			ultraDataColumn40.DataType = typeof(int);
			ultraDataColumn41.DataType = typeof(System.Decimal);
			ultraDataColumn41.DefaultValue = new System.Decimal(new int[] {
																																			0,
																																			0,
																																			0,
																																			0});
			ultraDataColumn42.DataType = typeof(System.Decimal);
			ultraDataColumn42.DefaultValue = new System.Decimal(new int[] {
																																			0,
																																			0,
																																			0,
																																			0});
			ultraDataColumn43.DataType = typeof(System.Decimal);
			ultraDataColumn43.DefaultValue = new System.Decimal(new int[] {
																																			0,
																																			0,
																																			0,
																																			0});
			ultraDataColumn44.DataType = typeof(int);
			ultraDataColumn44.DefaultValue = 0;
			ultraDataColumn45.DataType = typeof(int);
			ultraDataColumn45.DefaultValue = 0;
			ultraDataColumn46.DataType = typeof(int);
			ultraDataColumn46.DefaultValue = 0;
			ultraDataColumn47.DataType = typeof(int);
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
																																 ultraDataColumn37,
																																 ultraDataColumn38,
																																 ultraDataColumn39,
																																 ultraDataColumn40,
																																 ultraDataColumn41,
																																 ultraDataColumn42,
																																 ultraDataColumn43,
																																 ultraDataColumn44,
																																 ultraDataColumn45,
																																 ultraDataColumn46,
																																 ultraDataColumn47});
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
			this.dtFecha.Location = new System.Drawing.Point(504, 81);
			this.dtFecha.Name = "dtFecha";
			this.dtFecha.NonAutoSizeHeight = 23;
			this.dtFecha.Size = new System.Drawing.Size(112, 21);
			this.dtFecha.SpinButtonsVisible = true;
			this.dtFecha.TabIndex = 399;
			this.dtFecha.Value = ((object)(resources.GetObject("dtFecha.Value")));
			this.dtFecha.KeyDown += new System.Windows.Forms.KeyEventHandler(this.dtFecha_KeyDown);
			this.dtFecha.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.dtFecha_KeyPress);
			this.dtFecha.Validating += new System.ComponentModel.CancelEventHandler(this.dtFecha_Validating);
			this.dtFecha.Validated += new System.EventHandler(this.dtFecha_Validated);
			this.dtFecha.BeforeDropDown += new System.ComponentModel.CancelEventHandler(this.dtFecha_BeforeDropDown);
			this.dtFecha.ValueChanged += new System.EventHandler(this.dtFecha_ValueChanged);
			// 
			// lblEstado
			// 
			this.lblEstado.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.lblEstado.AutoSize = true;
			this.lblEstado.Font = new System.Drawing.Font("Microsoft Sans Serif", 15F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.lblEstado.ForeColor = System.Drawing.Color.Firebrick;
			this.lblEstado.Location = new System.Drawing.Point(600, 439);
			this.lblEstado.Name = "lblEstado";
			this.lblEstado.Size = new System.Drawing.Size(0, 26);
			this.lblEstado.TabIndex = 398;
			// 
			// label36
			// 
			this.label36.AutoSize = true;
			this.label36.Location = new System.Drawing.Point(80, 83);
			this.label36.Name = "label36";
			this.label36.Size = new System.Drawing.Size(30, 16);
			this.label36.TabIndex = 396;
			this.label36.Text = "Local";
			this.label36.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// label32
			// 
			this.label32.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			this.label32.AutoSize = true;
			this.label32.Location = new System.Drawing.Point(957, 11);
			this.label32.Name = "label32";
			this.label32.Size = new System.Drawing.Size(44, 16);
			this.label32.TabIndex = 395;
			this.label32.Text = "Número";
			this.label32.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// label4
			// 
			this.label4.AutoSize = true;
			this.label4.Location = new System.Drawing.Point(424, 83);
			this.label4.Name = "label4";
			this.label4.Size = new System.Drawing.Size(34, 16);
			this.label4.TabIndex = 393;
			this.label4.Text = "Fecha";
			this.label4.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// cmbBodega
			// 
			appearance22.ForeColor = System.Drawing.Color.Black;
			appearance22.ForeColorDisabled = System.Drawing.Color.Black;
			this.cmbBodega.Appearance = appearance22;
			this.cmbBodega.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbBodega.Cursor = System.Windows.Forms.Cursors.Default;
			ultraGridColumn49.Header.VisiblePosition = 2;
			ultraGridColumn50.Header.VisiblePosition = 1;
			ultraGridColumn50.Width = 180;
			ultraGridColumn51.Header.VisiblePosition = 0;
			ultraGridBand2.Columns.AddRange(new object[] {
																										 ultraGridColumn49,
																										 ultraGridColumn50,
																										 ultraGridColumn51});
			this.cmbBodega.DisplayLayout.BandsSerializer.Add(ultraGridBand2);
			this.cmbBodega.DisplayMember = "Nombre";
			this.cmbBodega.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbBodega.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbBodega.Enabled = false;
			this.cmbBodega.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbBodega.Location = new System.Drawing.Point(144, 80);
			this.cmbBodega.MaxDropDownItems = 30;
			this.cmbBodega.Name = "cmbBodega";
			this.cmbBodega.Size = new System.Drawing.Size(256, 22);
			this.cmbBodega.TabIndex = 397;
			this.cmbBodega.ValueMember = "Bodega";
			this.cmbBodega.KeyDown += new System.Windows.Forms.KeyEventHandler(this.cmbBodega_KeyDown);
			this.cmbBodega.Validating += new System.ComponentModel.CancelEventHandler(this.cmbBodega_Validating);
			this.cmbBodega.InitializeLayout += new Infragistics.Win.UltraWinGrid.InitializeLayoutEventHandler(this.cmbBodega_InitializeLayout);
			// 
			// lblNumero
			// 
			this.lblNumero.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			this.lblNumero.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
			this.lblNumero.Font = new System.Drawing.Font("Tahoma", 10F, System.Drawing.FontStyle.Bold);
			this.lblNumero.Location = new System.Drawing.Point(1029, 8);
			this.lblNumero.Name = "lblNumero";
			this.lblNumero.Size = new System.Drawing.Size(128, 22);
			this.lblNumero.TabIndex = 394;
			this.lblNumero.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// groupBox1
			// 
			this.groupBox1.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.groupBox1.Location = new System.Drawing.Point(0, 64);
			this.groupBox1.Name = "groupBox1";
			this.groupBox1.Size = new System.Drawing.Size(1165, 8);
			this.groupBox1.TabIndex = 400;
			this.groupBox1.TabStop = false;
			// 
			// label8
			// 
			this.label8.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.label8.AutoSize = true;
			this.label8.BackColor = System.Drawing.Color.Transparent;
			this.label8.Location = new System.Drawing.Point(80, 444);
			this.label8.Name = "label8";
			this.label8.Size = new System.Drawing.Size(33, 16);
			this.label8.TabIndex = 406;
			this.label8.Text = "Notas";
			this.label8.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.label8.Click += new System.EventHandler(this.label8_Click);
			// 
			// label6
			// 
			this.label6.AutoSize = true;
			this.label6.BackColor = System.Drawing.Color.Transparent;
			this.label6.Location = new System.Drawing.Point(80, 155);
			this.label6.Name = "label6";
			this.label6.Size = new System.Drawing.Size(38, 16);
			this.label6.TabIndex = 404;
			this.label6.Text = "Cédula";
			this.label6.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// label5
			// 
			this.label5.AutoSize = true;
			this.label5.BackColor = System.Drawing.Color.Transparent;
			this.label5.Location = new System.Drawing.Point(280, 155);
			this.label5.Name = "label5";
			this.label5.Size = new System.Drawing.Size(44, 16);
			this.label5.TabIndex = 402;
			this.label5.Text = "Nombre";
			this.label5.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// txtObservaciones
			// 
			this.txtObservaciones.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			appearance23.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtObservaciones.Appearance = appearance23;
			this.txtObservaciones.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtObservaciones.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtObservaciones.Enabled = false;
			this.txtObservaciones.Location = new System.Drawing.Point(144, 408);
			this.txtObservaciones.MaxLength = 1000;
			this.txtObservaciones.Multiline = true;
			this.txtObservaciones.Name = "txtObservaciones";
			this.txtObservaciones.Size = new System.Drawing.Size(432, 88);
			this.txtObservaciones.TabIndex = 405;
			this.txtObservaciones.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtObservaciones_KeyDown);
			// 
			// txtCodigo
			// 
			appearance24.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtCodigo.Appearance = appearance24;
			this.txtCodigo.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtCodigo.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtCodigo.Enabled = false;
			this.txtCodigo.Location = new System.Drawing.Point(144, 152);
			this.txtCodigo.MaxLength = 10;
			this.txtCodigo.Name = "txtCodigo";
			this.txtCodigo.Size = new System.Drawing.Size(112, 22);
			this.txtCodigo.TabIndex = 403;
			this.txtCodigo.ValueChanged += new System.EventHandler(this.txtCodigo_ValueChanged);
			// 
			// txtNombre
			// 
			appearance25.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtNombre.Appearance = appearance25;
			this.txtNombre.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtNombre.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtNombre.Enabled = false;
			this.txtNombre.Location = new System.Drawing.Point(344, 152);
			this.txtNombre.Name = "txtNombre";
			this.txtNombre.Size = new System.Drawing.Size(384, 22);
			this.txtNombre.TabIndex = 401;
			// 
			// dtEmision
			// 
			appearance26.ForeColorDisabled = System.Drawing.Color.Black;
			this.dtEmision.Appearance = appearance26;
			this.dtEmision.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			dateButton2.Caption = "Today";
			this.dtEmision.DateButtons.Add(dateButton2);
			this.dtEmision.Enabled = false;
			this.dtEmision.Format = "dd/MM/yyyy";
			this.dtEmision.Location = new System.Drawing.Point(144, 177);
			this.dtEmision.Name = "dtEmision";
			this.dtEmision.NonAutoSizeHeight = 23;
			this.dtEmision.Size = new System.Drawing.Size(112, 21);
			this.dtEmision.SpinButtonsVisible = true;
			this.dtEmision.TabIndex = 408;
			this.dtEmision.Value = ((object)(resources.GetObject("dtEmision.Value")));
			// 
			// label1
			// 
			this.label1.AutoSize = true;
			this.label1.Location = new System.Drawing.Point(80, 179);
			this.label1.Name = "label1";
			this.label1.Size = new System.Drawing.Size(43, 16);
			this.label1.TabIndex = 407;
			this.label1.Text = "Emisión";
			this.label1.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// label2
			// 
			this.label2.AutoSize = true;
			this.label2.BackColor = System.Drawing.Color.Transparent;
			this.label2.Location = new System.Drawing.Point(80, 11);
			this.label2.Name = "label2";
			this.label2.TabIndex = 410;
			this.label2.Text = "Numero de Factura";
			this.label2.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// txtNumeroFactura
			// 
			appearance27.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtNumeroFactura.Appearance = appearance27;
			this.txtNumeroFactura.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtNumeroFactura.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtNumeroFactura.Enabled = false;
			this.txtNumeroFactura.Location = new System.Drawing.Point(192, 8);
			this.txtNumeroFactura.MaxLength = 13;
			this.txtNumeroFactura.Name = "txtNumeroFactura";
			this.txtNumeroFactura.Size = new System.Drawing.Size(112, 22);
			this.txtNumeroFactura.TabIndex = 409;
			this.toolTip1.SetToolTip(this.txtNumeroFactura, "Presione F3 para Busca un Documento");
			this.txtNumeroFactura.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtNumeroFactura_KeyDown);
			this.txtNumeroFactura.ValueChanged += new System.EventHandler(this.txtNumeroFactura_ValueChanged);
			// 
			// txtEstado
			// 
			this.txtEstado.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			appearance28.ForeColorDisabled = System.Drawing.Color.Black;
			appearance28.TextHAlign = Infragistics.Win.HAlign.Right;
			appearance28.TextVAlign = Infragistics.Win.VAlign.Middle;
			this.txtEstado.Appearance = appearance28;
			this.txtEstado.Enabled = false;
			this.txtEstado.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.txtEstado.FormatString = "";
			this.txtEstado.Location = new System.Drawing.Point(1141, 528);
			this.txtEstado.Name = "txtEstado";
			this.txtEstado.PromptChar = ' ';
			this.txtEstado.Size = new System.Drawing.Size(16, 21);
			this.txtEstado.TabIndex = 423;
			this.txtEstado.Visible = false;
			// 
			// txtBodegaPredef
			// 
			this.txtBodegaPredef.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			appearance29.ForeColorDisabled = System.Drawing.Color.Black;
			appearance29.TextHAlign = Infragistics.Win.HAlign.Right;
			appearance29.TextVAlign = Infragistics.Win.VAlign.Middle;
			this.txtBodegaPredef.Appearance = appearance29;
			this.txtBodegaPredef.Enabled = false;
			this.txtBodegaPredef.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.txtBodegaPredef.FormatString = "";
			this.txtBodegaPredef.Location = new System.Drawing.Point(1125, 528);
			this.txtBodegaPredef.Name = "txtBodegaPredef";
			this.txtBodegaPredef.PromptChar = ' ';
			this.txtBodegaPredef.Size = new System.Drawing.Size(16, 21);
			this.txtBodegaPredef.TabIndex = 422;
			this.txtBodegaPredef.Visible = false;
			// 
			// txtIdIngreso
			// 
			this.txtIdIngreso.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			appearance30.ForeColorDisabled = System.Drawing.Color.Black;
			appearance30.TextHAlign = Infragistics.Win.HAlign.Right;
			appearance30.TextVAlign = Infragistics.Win.VAlign.Middle;
			this.txtIdIngreso.Appearance = appearance30;
			this.txtIdIngreso.Enabled = false;
			this.txtIdIngreso.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.txtIdIngreso.FormatString = "";
			this.txtIdIngreso.Location = new System.Drawing.Point(1093, 528);
			this.txtIdIngreso.Name = "txtIdIngreso";
			this.txtIdIngreso.PromptChar = ' ';
			this.txtIdIngreso.Size = new System.Drawing.Size(16, 21);
			this.txtIdIngreso.TabIndex = 420;
			this.txtIdIngreso.Visible = false;
			// 
			// btnNuevo
			// 
			this.btnNuevo.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btnNuevo.CausesValidation = false;
			this.btnNuevo.Enabled = false;
			this.btnNuevo.Image = ((System.Drawing.Image)(resources.GetObject("btnNuevo.Image")));
			this.btnNuevo.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnNuevo.Location = new System.Drawing.Point(8, 528);
			this.btnNuevo.Name = "btnNuevo";
			this.btnNuevo.Size = new System.Drawing.Size(72, 23);
			this.btnNuevo.TabIndex = 412;
			this.btnNuevo.Text = "&Nuevo";
			this.btnNuevo.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnNuevo.Click += new System.EventHandler(this.btnNuevo_Click);
			// 
			// btnConsultar
			// 
			this.btnConsultar.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btnConsultar.CausesValidation = false;
			this.btnConsultar.Enabled = false;
			this.btnConsultar.Image = ((System.Drawing.Image)(resources.GetObject("btnConsultar.Image")));
			this.btnConsultar.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnConsultar.Location = new System.Drawing.Point(82, 528);
			this.btnConsultar.Name = "btnConsultar";
			this.btnConsultar.Size = new System.Drawing.Size(78, 23);
			this.btnConsultar.TabIndex = 413;
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
			this.btnEditar.Location = new System.Drawing.Point(162, 528);
			this.btnEditar.Name = "btnEditar";
			this.btnEditar.Size = new System.Drawing.Size(76, 23);
			this.btnEditar.TabIndex = 416;
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
			this.btnImprimir.Location = new System.Drawing.Point(240, 528);
			this.btnImprimir.Name = "btnImprimir";
			this.btnImprimir.Size = new System.Drawing.Size(77, 23);
			this.btnImprimir.TabIndex = 415;
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
			this.btnGuardar.Location = new System.Drawing.Point(319, 528);
			this.btnGuardar.Name = "btnGuardar";
			this.btnGuardar.Size = new System.Drawing.Size(76, 23);
			this.btnGuardar.TabIndex = 414;
			this.btnGuardar.Text = "&Guardar";
			this.btnGuardar.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnGuardar.Click += new System.EventHandler(this.btnGuardar_Click);
			// 
			// btnCancelar
			// 
			this.btnCancelar.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btnCancelar.CausesValidation = false;
			this.btnCancelar.Enabled = false;
			this.btnCancelar.Image = ((System.Drawing.Image)(resources.GetObject("btnCancelar.Image")));
			this.btnCancelar.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnCancelar.Location = new System.Drawing.Point(475, 528);
			this.btnCancelar.Name = "btnCancelar";
			this.btnCancelar.Size = new System.Drawing.Size(72, 23);
			this.btnCancelar.TabIndex = 418;
			this.btnCancelar.Text = "&Cancelar";
			this.btnCancelar.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnCancelar.Click += new System.EventHandler(this.btnCancelar_Click);
			// 
			// btnAnular
			// 
			this.btnAnular.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btnAnular.CausesValidation = false;
			this.btnAnular.Enabled = false;
			this.btnAnular.Image = ((System.Drawing.Image)(resources.GetObject("btnAnular.Image")));
			this.btnAnular.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnAnular.Location = new System.Drawing.Point(397, 528);
			this.btnAnular.Name = "btnAnular";
			this.btnAnular.Size = new System.Drawing.Size(76, 23);
			this.btnAnular.TabIndex = 417;
			this.btnAnular.Text = "&Anular";
			this.btnAnular.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnAnular.Click += new System.EventHandler(this.btnAnular_Click);
			// 
			// btnSalir
			// 
			this.btnSalir.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btnSalir.CausesValidation = false;
			this.btnSalir.Image = ((System.Drawing.Image)(resources.GetObject("btnSalir.Image")));
			this.btnSalir.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnSalir.Location = new System.Drawing.Point(549, 528);
			this.btnSalir.Name = "btnSalir";
			this.btnSalir.TabIndex = 419;
			this.btnSalir.Text = "&Salir";
			this.btnSalir.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnSalir.Click += new System.EventHandler(this.btnSalir_Click);
			// 
			// groupBox5
			// 
			this.groupBox5.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.groupBox5.Location = new System.Drawing.Point(0, 512);
			this.groupBox5.Name = "groupBox5";
			this.groupBox5.Size = new System.Drawing.Size(1165, 8);
			this.groupBox5.TabIndex = 411;
			this.groupBox5.TabStop = false;
			// 
			// txtDiasDeMora
			// 
			appearance31.ForeColor = System.Drawing.Color.Black;
			appearance31.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtDiasDeMora.Appearance = appearance31;
			this.txtDiasDeMora.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtDiasDeMora.Enabled = false;
			this.txtDiasDeMora.Location = new System.Drawing.Point(800, 31);
			this.txtDiasDeMora.MinValue = 0;
			this.txtDiasDeMora.Name = "txtDiasDeMora";
			this.txtDiasDeMora.PromptChar = ' ';
			this.txtDiasDeMora.Size = new System.Drawing.Size(72, 22);
			this.txtDiasDeMora.TabIndex = 425;
			// 
			// label3
			// 
			this.label3.AutoSize = true;
			this.label3.BackColor = System.Drawing.Color.Transparent;
			this.label3.Location = new System.Drawing.Point(704, 34);
			this.label3.Name = "label3";
			this.label3.Size = new System.Drawing.Size(71, 16);
			this.label3.TabIndex = 424;
			this.label3.Text = "Dias De Mora";
			this.label3.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// txtSaldo
			// 
			appearance32.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance32.ForeColor = System.Drawing.Color.Black;
			appearance32.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtSaldo.Appearance = appearance32;
			this.txtSaldo.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			this.txtSaldo.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtSaldo.Enabled = false;
			this.txtSaldo.FormatString = "#,##0.00";
			this.txtSaldo.Location = new System.Drawing.Point(800, 152);
			this.txtSaldo.Name = "txtSaldo";
			this.txtSaldo.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtSaldo.PromptChar = ' ';
			this.txtSaldo.Size = new System.Drawing.Size(88, 22);
			this.txtSaldo.TabIndex = 427;
			// 
			// label16
			// 
			this.label16.AutoSize = true;
			this.label16.BackColor = System.Drawing.Color.Transparent;
			this.label16.Location = new System.Drawing.Point(752, 155);
			this.label16.Name = "label16";
			this.label16.Size = new System.Drawing.Size(31, 16);
			this.label16.TabIndex = 426;
			this.label16.Text = "Saldo";
			this.label16.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// dtFechaRetiro
			// 
			appearance33.ForeColorDisabled = System.Drawing.Color.Black;
			this.dtFechaRetiro.Appearance = appearance33;
			this.dtFechaRetiro.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			dateButton3.Caption = "Today";
			this.dtFechaRetiro.DateButtons.Add(dateButton3);
			this.dtFechaRetiro.Enabled = false;
			this.dtFechaRetiro.Format = "dd/MM/yyyy";
			this.dtFechaRetiro.Location = new System.Drawing.Point(192, 32);
			this.dtFechaRetiro.Name = "dtFechaRetiro";
			this.dtFechaRetiro.NonAutoSizeHeight = 23;
			this.dtFechaRetiro.Size = new System.Drawing.Size(112, 21);
			this.dtFechaRetiro.SpinButtonsVisible = true;
			this.dtFechaRetiro.TabIndex = 429;
			this.dtFechaRetiro.Value = ((object)(resources.GetObject("dtFechaRetiro.Value")));
			this.dtFechaRetiro.KeyDown += new System.Windows.Forms.KeyEventHandler(this.dtFechaRetiro_KeyDown);
			this.dtFechaRetiro.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.dtFechaRetiro_KeyPress);
			this.dtFechaRetiro.Validating += new System.ComponentModel.CancelEventHandler(this.dtFechaRetiro_Validating);
			this.dtFechaRetiro.BeforeDropDown += new System.ComponentModel.CancelEventHandler(this.dtFechaRetiro_BeforeDropDown);
			this.dtFechaRetiro.ValueChanged += new System.EventHandler(this.dtFechaRetiro_ValueChanged);
			// 
			// label7
			// 
			this.label7.AutoSize = true;
			this.label7.Location = new System.Drawing.Point(80, 34);
			this.label7.Name = "label7";
			this.label7.Size = new System.Drawing.Size(82, 16);
			this.label7.TabIndex = 428;
			this.label7.Text = "Fecha de Retiro";
			this.label7.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// txtPagado
			// 
			appearance34.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance34.ForeColor = System.Drawing.Color.Black;
			appearance34.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtPagado.Appearance = appearance34;
			this.txtPagado.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			this.txtPagado.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtPagado.Enabled = false;
			this.txtPagado.FormatString = "#,##0.00";
			this.txtPagado.Location = new System.Drawing.Point(992, 152);
			this.txtPagado.Name = "txtPagado";
			this.txtPagado.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtPagado.PromptChar = ' ';
			this.txtPagado.Size = new System.Drawing.Size(88, 22);
			this.txtPagado.TabIndex = 431;
			// 
			// label9
			// 
			this.label9.AutoSize = true;
			this.label9.BackColor = System.Drawing.Color.Transparent;
			this.label9.Location = new System.Drawing.Point(904, 155);
			this.label9.Name = "label9";
			this.label9.Size = new System.Drawing.Size(41, 16);
			this.label9.TabIndex = 430;
			this.label9.Text = "Pagado";
			this.label9.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// txtTotalFactura
			// 
			appearance35.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance35.ForeColor = System.Drawing.Color.Black;
			appearance35.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtTotalFactura.Appearance = appearance35;
			this.txtTotalFactura.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			this.txtTotalFactura.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtTotalFactura.Enabled = false;
			this.txtTotalFactura.FormatString = "#,##0.00";
			this.txtTotalFactura.Location = new System.Drawing.Point(800, 176);
			this.txtTotalFactura.Name = "txtTotalFactura";
			this.txtTotalFactura.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtTotalFactura.PromptChar = ' ';
			this.txtTotalFactura.Size = new System.Drawing.Size(88, 22);
			this.txtTotalFactura.TabIndex = 433;
			// 
			// label10
			// 
			this.label10.AutoSize = true;
			this.label10.BackColor = System.Drawing.Color.Transparent;
			this.label10.Location = new System.Drawing.Point(752, 179);
			this.label10.Name = "label10";
			this.label10.Size = new System.Drawing.Size(29, 16);
			this.label10.TabIndex = 432;
			this.label10.Text = "Total";
			this.label10.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// label11
			// 
			this.label11.AutoSize = true;
			this.label11.BackColor = System.Drawing.Color.Transparent;
			this.label11.Location = new System.Drawing.Point(80, 203);
			this.label11.Name = "label11";
			this.label11.Size = new System.Drawing.Size(30, 16);
			this.label11.TabIndex = 437;
			this.label11.Text = "Local";
			this.label11.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.label11.Click += new System.EventHandler(this.label11_Click);
			// 
			// label12
			// 
			this.label12.AutoSize = true;
			this.label12.BackColor = System.Drawing.Color.Transparent;
			this.label12.Location = new System.Drawing.Point(424, 203);
			this.label12.Name = "label12";
			this.label12.Size = new System.Drawing.Size(52, 16);
			this.label12.TabIndex = 435;
			this.label12.Text = "Vendedor";
			this.label12.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// txtLocal
			// 
			appearance36.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtLocal.Appearance = appearance36;
			this.txtLocal.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtLocal.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtLocal.Enabled = false;
			this.txtLocal.Location = new System.Drawing.Point(144, 200);
			this.txtLocal.MaxLength = 13;
			this.txtLocal.Name = "txtLocal";
			this.txtLocal.Size = new System.Drawing.Size(232, 22);
			this.txtLocal.TabIndex = 436;
			// 
			// txtVendedor
			// 
			appearance37.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtVendedor.Appearance = appearance37;
			this.txtVendedor.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtVendedor.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtVendedor.Enabled = false;
			this.txtVendedor.Location = new System.Drawing.Point(496, 200);
			this.txtVendedor.Name = "txtVendedor";
			this.txtVendedor.Size = new System.Drawing.Size(232, 22);
			this.txtVendedor.TabIndex = 434;
			// 
			// label13
			// 
			this.label13.AutoSize = true;
			this.label13.Location = new System.Drawing.Point(280, 179);
			this.label13.Name = "label13";
			this.label13.Size = new System.Drawing.Size(34, 16);
			this.label13.TabIndex = 438;
			this.label13.Text = "Banco";
			this.label13.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// ultraDataSource2
			// 
			ultraDataColumn48.DataType = typeof(int);
			this.ultraDataSource2.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn48,
																																 ultraDataColumn49,
																																 ultraDataColumn50});
			// 
			// txtIdCompra
			// 
			this.txtIdCompra.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			appearance38.ForeColorDisabled = System.Drawing.Color.Black;
			appearance38.TextHAlign = Infragistics.Win.HAlign.Right;
			appearance38.TextVAlign = Infragistics.Win.VAlign.Middle;
			this.txtIdCompra.Appearance = appearance38;
			this.txtIdCompra.Enabled = false;
			this.txtIdCompra.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.txtIdCompra.FormatString = "";
			this.txtIdCompra.Location = new System.Drawing.Point(1109, 528);
			this.txtIdCompra.Name = "txtIdCompra";
			this.txtIdCompra.PromptChar = ' ';
			this.txtIdCompra.Size = new System.Drawing.Size(16, 21);
			this.txtIdCompra.TabIndex = 440;
			this.txtIdCompra.Visible = false;
			// 
			// txtTotalCredito
			// 
			appearance39.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance39.ForeColor = System.Drawing.Color.Black;
			appearance39.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtTotalCredito.Appearance = appearance39;
			this.txtTotalCredito.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			this.txtTotalCredito.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtTotalCredito.Enabled = false;
			this.txtTotalCredito.FormatString = "#,##0.00";
			this.txtTotalCredito.Location = new System.Drawing.Point(992, 176);
			this.txtTotalCredito.Name = "txtTotalCredito";
			this.txtTotalCredito.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtTotalCredito.PromptChar = ' ';
			this.txtTotalCredito.Size = new System.Drawing.Size(88, 22);
			this.txtTotalCredito.TabIndex = 442;
			// 
			// label14
			// 
			this.label14.AutoSize = true;
			this.label14.BackColor = System.Drawing.Color.Transparent;
			this.label14.Location = new System.Drawing.Point(904, 179);
			this.label14.Name = "label14";
			this.label14.Size = new System.Drawing.Size(74, 16);
			this.label14.TabIndex = 441;
			this.label14.Text = "Monto Crédito";
			this.label14.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// txtEntidadFinanciera
			// 
			appearance40.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtEntidadFinanciera.Appearance = appearance40;
			this.txtEntidadFinanciera.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtEntidadFinanciera.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtEntidadFinanciera.Enabled = false;
			this.txtEntidadFinanciera.Location = new System.Drawing.Point(344, 176);
			this.txtEntidadFinanciera.MaxLength = 13;
			this.txtEntidadFinanciera.Name = "txtEntidadFinanciera";
			this.txtEntidadFinanciera.Size = new System.Drawing.Size(176, 22);
			this.txtEntidadFinanciera.TabIndex = 443;
			// 
			// txtNumeroNotaDeCredito
			// 
			appearance41.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtNumeroNotaDeCredito.Appearance = appearance41;
			this.txtNumeroNotaDeCredito.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtNumeroNotaDeCredito.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtNumeroNotaDeCredito.Enabled = false;
			this.txtNumeroNotaDeCredito.Location = new System.Drawing.Point(504, 104);
			this.txtNumeroNotaDeCredito.MaxLength = 9;
			this.txtNumeroNotaDeCredito.Name = "txtNumeroNotaDeCredito";
			this.txtNumeroNotaDeCredito.Size = new System.Drawing.Size(112, 22);
			this.txtNumeroNotaDeCredito.TabIndex = 473;
			this.txtNumeroNotaDeCredito.Visible = false;
			this.txtNumeroNotaDeCredito.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtNumeroNotaDeCredito_KeyDown);
			this.txtNumeroNotaDeCredito.Validated += new System.EventHandler(this.txtNumeroNotaDeCredito_Validated);
			// 
			// groupBox2
			// 
			this.groupBox2.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.groupBox2.Location = new System.Drawing.Point(0, 136);
			this.groupBox2.Name = "groupBox2";
			this.groupBox2.Size = new System.Drawing.Size(1165, 8);
			this.groupBox2.TabIndex = 444;
			this.groupBox2.TabStop = false;
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
			// cmbEstados
			// 
			this.cmbEstados.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbEstados.DataSource = this.ultraDataSource3;
			this.cmbEstados.DisplayLayout.AutoFitColumns = true;
			ultraGridColumn52.Header.VisiblePosition = 0;
			ultraGridColumn52.Hidden = true;
			ultraGridColumn52.Width = 78;
			ultraGridColumn53.Header.VisiblePosition = 1;
			ultraGridColumn53.Width = 98;
			ultraGridBand3.Columns.AddRange(new object[] {
																										 ultraGridColumn52,
																										 ultraGridColumn53});
			ultraGridBand3.SummaryFooterCaption = "Grand Summaries";
			this.cmbEstados.DisplayLayout.BandsSerializer.Add(ultraGridBand3);
			this.cmbEstados.DisplayMember = "Descripcion";
			this.cmbEstados.Font = new System.Drawing.Font("Tahoma", 8.25F);
			this.cmbEstados.Location = new System.Drawing.Point(456, 240);
			this.cmbEstados.Name = "cmbEstados";
			this.cmbEstados.Size = new System.Drawing.Size(100, 104);
			this.cmbEstados.TabIndex = 445;
			this.cmbEstados.ValueMember = "idDTEstado";
			this.cmbEstados.Visible = false;
			// 
			// ultraDataSource3
			// 
			ultraDataColumn51.DataType = typeof(int);
			this.ultraDataSource3.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn51,
																																 ultraDataColumn52});
			// 
			// txtValorNC
			// 
			this.txtValorNC.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			appearance42.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance42.ForeColor = System.Drawing.Color.Black;
			appearance42.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtValorNC.Appearance = appearance42;
			this.txtValorNC.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			this.txtValorNC.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtValorNC.Enabled = false;
			this.txtValorNC.FormatString = "#,##0.00";
			this.txtValorNC.Location = new System.Drawing.Point(1068, 456);
			this.txtValorNC.Name = "txtValorNC";
			this.txtValorNC.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtValorNC.PromptChar = ' ';
			this.txtValorNC.Size = new System.Drawing.Size(88, 22);
			this.txtValorNC.TabIndex = 449;
			// 
			// label15
			// 
			this.label15.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			this.label15.AutoSize = true;
			this.label15.BackColor = System.Drawing.Color.Transparent;
			this.label15.Location = new System.Drawing.Point(940, 459);
			this.label15.Name = "label15";
			this.label15.Size = new System.Drawing.Size(29, 16);
			this.label15.TabIndex = 448;
			this.label15.Text = "Total";
			this.label15.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// txtidTransferencia
			// 
			this.txtidTransferencia.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			appearance43.ForeColorDisabled = System.Drawing.Color.Black;
			appearance43.TextHAlign = Infragistics.Win.HAlign.Right;
			appearance43.TextVAlign = Infragistics.Win.VAlign.Middle;
			this.txtidTransferencia.Appearance = appearance43;
			this.txtidTransferencia.Enabled = false;
			this.txtidTransferencia.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.txtidTransferencia.FormatString = "";
			this.txtidTransferencia.Location = new System.Drawing.Point(1077, 528);
			this.txtidTransferencia.Name = "txtidTransferencia";
			this.txtidTransferencia.PromptChar = ' ';
			this.txtidTransferencia.Size = new System.Drawing.Size(16, 21);
			this.txtidTransferencia.TabIndex = 451;
			this.txtidTransferencia.Visible = false;
			// 
			// txtIdNotaDeCredito
			// 
			this.txtIdNotaDeCredito.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			appearance44.ForeColorDisabled = System.Drawing.Color.Black;
			appearance44.TextHAlign = Infragistics.Win.HAlign.Right;
			appearance44.TextVAlign = Infragistics.Win.VAlign.Middle;
			this.txtIdNotaDeCredito.Appearance = appearance44;
			this.txtIdNotaDeCredito.Enabled = false;
			this.txtIdNotaDeCredito.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.txtIdNotaDeCredito.FormatString = "";
			this.txtIdNotaDeCredito.Location = new System.Drawing.Point(1061, 528);
			this.txtIdNotaDeCredito.Name = "txtIdNotaDeCredito";
			this.txtIdNotaDeCredito.PromptChar = ' ';
			this.txtIdNotaDeCredito.Size = new System.Drawing.Size(16, 21);
			this.txtIdNotaDeCredito.TabIndex = 450;
			this.txtIdNotaDeCredito.Visible = false;
			// 
			// label18
			// 
			this.label18.AutoSize = true;
			this.label18.Location = new System.Drawing.Point(344, 34);
			this.label18.Name = "label18";
			this.label18.Size = new System.Drawing.Size(37, 16);
			this.label18.TabIndex = 453;
			this.label18.Text = "Mótivo";
			this.label18.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// chkFiscal
			// 
			appearance45.ForeColorDisabled = System.Drawing.Color.Black;
			this.chkFiscal.Appearance = appearance45;
			this.chkFiscal.BackColor = System.Drawing.Color.Transparent;
			this.chkFiscal.CausesValidation = false;
			this.chkFiscal.Enabled = false;
			this.chkFiscal.Location = new System.Drawing.Point(957, 32);
			this.chkFiscal.Name = "chkFiscal";
			this.chkFiscal.Size = new System.Drawing.Size(48, 22);
			this.chkFiscal.TabIndex = 452;
			this.chkFiscal.Text = "Fiscal";
			this.chkFiscal.CheckedChanged += new System.EventHandler(this.chkFiscal_CheckedChanged);
			// 
			// cmbMotivos
			// 
			appearance46.ForeColorDisabled = System.Drawing.Color.Black;
			this.cmbMotivos.Appearance = appearance46;
			this.cmbMotivos.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbMotivos.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbMotivos.DataSource = this.ultraDataSource4;
			ultraGridColumn54.Header.VisiblePosition = 0;
			ultraGridColumn54.Hidden = true;
			ultraGridColumn55.Header.VisiblePosition = 1;
			ultraGridColumn55.Width = 344;
			ultraGridBand4.Columns.AddRange(new object[] {
																										 ultraGridColumn54,
																										 ultraGridColumn55});
			this.cmbMotivos.DisplayLayout.BandsSerializer.Add(ultraGridBand4);
			this.cmbMotivos.DisplayMember = "Motivo";
			this.cmbMotivos.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbMotivos.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbMotivos.Enabled = false;
			this.cmbMotivos.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbMotivos.Location = new System.Drawing.Point(424, 32);
			this.cmbMotivos.Name = "cmbMotivos";
			this.cmbMotivos.Size = new System.Drawing.Size(256, 22);
			this.cmbMotivos.TabIndex = 455;
			this.cmbMotivos.ValueMember = "idMotivo";
			this.cmbMotivos.KeyDown += new System.Windows.Forms.KeyEventHandler(this.cmbMotivos_KeyDown);
			this.cmbMotivos.ValueChanged += new System.EventHandler(this.cmbMotivos_ValueChanged);
			this.cmbMotivos.InitializeLayout += new Infragistics.Win.UltraWinGrid.InitializeLayoutEventHandler(this.cmbMotivos_InitializeLayout);
			// 
			// ultraDataSource4
			// 
			ultraDataColumn53.DataType = typeof(int);
			this.ultraDataSource4.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn53,
																																 ultraDataColumn54});
			// 
			// lblSubTotal
			// 
			this.lblSubTotal.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			this.lblSubTotal.AutoSize = true;
			this.lblSubTotal.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.lblSubTotal.Location = new System.Drawing.Point(940, 387);
			this.lblSubTotal.Name = "lblSubTotal";
			this.lblSubTotal.Size = new System.Drawing.Size(46, 16);
			this.lblSubTotal.TabIndex = 458;
			this.lblSubTotal.Text = "Subtotal";
			// 
			// lblIva
			// 
			this.lblIva.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			this.lblIva.AutoSize = true;
			this.lblIva.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.lblIva.Location = new System.Drawing.Point(940, 435);
			this.lblIva.Name = "lblIva";
			this.lblIva.Size = new System.Drawing.Size(22, 16);
			this.lblIva.TabIndex = 457;
			this.lblIva.Text = "IVA";
			// 
			// lblDescuento
			// 
			this.lblDescuento.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			this.lblDescuento.AutoSize = true;
			this.lblDescuento.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.lblDescuento.Location = new System.Drawing.Point(940, 411);
			this.lblDescuento.Name = "lblDescuento";
			this.lblDescuento.Size = new System.Drawing.Size(58, 16);
			this.lblDescuento.TabIndex = 456;
			this.lblDescuento.Text = "Descuento";
			// 
			// txtIvaTotal
			// 
			this.txtIvaTotal.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			appearance47.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtIvaTotal.Appearance = appearance47;
			this.txtIvaTotal.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtIvaTotal.Enabled = false;
			this.txtIvaTotal.FormatString = "#,##0.00";
			this.txtIvaTotal.Location = new System.Drawing.Point(1068, 432);
			this.txtIvaTotal.Name = "txtIvaTotal";
			this.txtIvaTotal.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtIvaTotal.PromptChar = ' ';
			this.txtIvaTotal.Size = new System.Drawing.Size(88, 22);
			this.txtIvaTotal.TabIndex = 461;
			// 
			// txtDescIvaTotal
			// 
			this.txtDescIvaTotal.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			appearance48.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtDescIvaTotal.Appearance = appearance48;
			this.txtDescIvaTotal.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtDescIvaTotal.Enabled = false;
			this.txtDescIvaTotal.FormatString = "#,##0.00";
			this.txtDescIvaTotal.Location = new System.Drawing.Point(1068, 408);
			this.txtDescIvaTotal.Name = "txtDescIvaTotal";
			this.txtDescIvaTotal.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtDescIvaTotal.PromptChar = ' ';
			this.txtDescIvaTotal.Size = new System.Drawing.Size(88, 22);
			this.txtDescIvaTotal.TabIndex = 460;
			// 
			// txtIva
			// 
			this.txtIva.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			appearance49.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtIva.Appearance = appearance49;
			this.txtIva.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtIva.Enabled = false;
			this.txtIva.FormatString = "#,##0.00";
			this.txtIva.Location = new System.Drawing.Point(1068, 384);
			this.txtIva.Name = "txtIva";
			this.txtIva.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtIva.PromptChar = ' ';
			this.txtIva.Size = new System.Drawing.Size(88, 22);
			this.txtIva.TabIndex = 459;
			// 
			// lblVendedor
			// 
			this.lblVendedor.AutoSize = true;
			this.lblVendedor.Location = new System.Drawing.Point(80, 107);
			this.lblVendedor.Name = "lblVendedor";
			this.lblVendedor.Size = new System.Drawing.Size(52, 16);
			this.lblVendedor.TabIndex = 463;
			this.lblVendedor.Text = "Vendedor";
			// 
			// cmbVendedor
			// 
			appearance50.ForeColor = System.Drawing.Color.Black;
			appearance50.ForeColorDisabled = System.Drawing.Color.Black;
			this.cmbVendedor.Appearance = appearance50;
			this.cmbVendedor.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbVendedor.Cursor = System.Windows.Forms.Cursors.Default;
			ultraGridColumn56.Header.VisiblePosition = 0;
			ultraGridColumn56.Hidden = true;
			ultraGridColumn57.Header.VisiblePosition = 1;
			ultraGridColumn57.Width = 335;
			ultraGridColumn58.Header.VisiblePosition = 2;
			ultraGridBand5.Columns.AddRange(new object[] {
																										 ultraGridColumn56,
																										 ultraGridColumn57,
																										 ultraGridColumn58});
			this.cmbVendedor.DisplayLayout.BandsSerializer.Add(ultraGridBand5);
			this.cmbVendedor.DisplayMember = "Nombre";
			this.cmbVendedor.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbVendedor.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbVendedor.Enabled = false;
			this.cmbVendedor.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbVendedor.Location = new System.Drawing.Point(144, 104);
			this.cmbVendedor.Name = "cmbVendedor";
			this.cmbVendedor.Size = new System.Drawing.Size(256, 22);
			this.cmbVendedor.TabIndex = 462;
			this.cmbVendedor.ValueMember = "idPersonal";
			this.cmbVendedor.KeyDown += new System.Windows.Forms.KeyEventHandler(this.cmbVendedor_KeyDown);
			// 
			// label19
			// 
			this.label19.AutoSize = true;
			this.label19.BackColor = System.Drawing.Color.Transparent;
			this.label19.Location = new System.Drawing.Point(616, 179);
			this.label19.Name = "label19";
			this.label19.Size = new System.Drawing.Size(40, 16);
			this.label19.TabIndex = 464;
			this.label19.Text = "Lote N.";
			this.label19.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// txtNLote
			// 
			appearance51.ForeColor = System.Drawing.Color.Black;
			appearance51.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtNLote.Appearance = appearance51;
			this.txtNLote.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtNLote.Enabled = false;
			this.txtNLote.Location = new System.Drawing.Point(672, 176);
			this.txtNLote.MinValue = 0;
			this.txtNLote.Name = "txtNLote";
			this.txtNLote.PromptChar = ' ';
			this.txtNLote.Size = new System.Drawing.Size(56, 22);
			this.txtNLote.TabIndex = 465;
			// 
			// label20
			// 
			this.label20.AutoSize = true;
			this.label20.Location = new System.Drawing.Point(344, 11);
			this.label20.Name = "label20";
			this.label20.Size = new System.Drawing.Size(67, 16);
			this.label20.TabIndex = 467;
			this.label20.Text = "Responsable";
			// 
			// cmbRetiradoPor
			// 
			appearance52.ForeColorDisabled = System.Drawing.Color.Black;
			this.cmbRetiradoPor.Appearance = appearance52;
			this.cmbRetiradoPor.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbRetiradoPor.Cursor = System.Windows.Forms.Cursors.Default;
			ultraGridColumn59.Header.VisiblePosition = 0;
			ultraGridColumn59.Hidden = true;
			ultraGridColumn60.Header.VisiblePosition = 1;
			ultraGridColumn60.Width = 335;
			ultraGridColumn61.Header.VisiblePosition = 2;
			ultraGridBand6.Columns.AddRange(new object[] {
																										 ultraGridColumn59,
																										 ultraGridColumn60,
																										 ultraGridColumn61});
			this.cmbRetiradoPor.DisplayLayout.BandsSerializer.Add(ultraGridBand6);
			this.cmbRetiradoPor.DisplayMember = "Nombre";
			this.cmbRetiradoPor.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbRetiradoPor.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbRetiradoPor.Enabled = false;
			this.cmbRetiradoPor.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbRetiradoPor.Location = new System.Drawing.Point(424, 9);
			this.cmbRetiradoPor.Name = "cmbRetiradoPor";
			this.cmbRetiradoPor.Size = new System.Drawing.Size(256, 22);
			this.cmbRetiradoPor.TabIndex = 466;
			this.cmbRetiradoPor.ValueMember = "idPersonal";
			this.cmbRetiradoPor.KeyDown += new System.Windows.Forms.KeyEventHandler(this.cmbRetiradoPor_KeyDown);
			this.cmbRetiradoPor.ValueChanged += new System.EventHandler(this.cmbRetiradoPor_ValueChanged);
			this.cmbRetiradoPor.InitializeLayout += new Infragistics.Win.UltraWinGrid.InitializeLayoutEventHandler(this.cmbRetiradoPor_InitializeLayout);
			// 
			// label21
			// 
			this.label21.AutoSize = true;
			this.label21.BackColor = System.Drawing.Color.Transparent;
			this.label21.Location = new System.Drawing.Point(704, 11);
			this.label21.Name = "label21";
			this.label21.Size = new System.Drawing.Size(79, 16);
			this.label21.TabIndex = 468;
			this.label21.Text = "Tiempo de Uso";
			this.label21.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// txtTiempoUso
			// 
			appearance53.ForeColor = System.Drawing.Color.Black;
			appearance53.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtTiempoUso.Appearance = appearance53;
			this.txtTiempoUso.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtTiempoUso.Enabled = false;
			this.txtTiempoUso.Location = new System.Drawing.Point(800, 8);
			this.txtTiempoUso.MinValue = 0;
			this.txtTiempoUso.Name = "txtTiempoUso";
			this.txtTiempoUso.PromptChar = ' ';
			this.txtTiempoUso.Size = new System.Drawing.Size(72, 22);
			this.txtTiempoUso.TabIndex = 469;
			this.txtTiempoUso.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtTiempoUso_KeyDown);
			this.txtTiempoUso.Click += new System.EventHandler(this.txtTiempoUso_Click);
			this.txtTiempoUso.Validated += new System.EventHandler(this.txtTiempoUso_Validated);
			this.txtTiempoUso.ValueChanged += new System.EventHandler(this.txtTiempoUso_ValueChanged);
			this.txtTiempoUso.Enter += new System.EventHandler(this.txtTiempoUso_Enter);
			// 
			// btnProcesar
			// 
			this.btnProcesar.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btnProcesar.Enabled = false;
			this.btnProcesar.Image = ((System.Drawing.Image)(resources.GetObject("btnProcesar.Image")));
			this.btnProcesar.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnProcesar.Location = new System.Drawing.Point(936, 528);
			this.btnProcesar.Name = "btnProcesar";
			this.btnProcesar.Size = new System.Drawing.Size(76, 23);
			this.btnProcesar.TabIndex = 470;
			this.btnProcesar.Text = "&Procesar";
			this.btnProcesar.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnProcesar.Visible = false;
			this.btnProcesar.Click += new System.EventHandler(this.btnProcesar_Click);
			// 
			// txtIdTransferenciaDT
			// 
			this.txtIdTransferenciaDT.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			appearance54.ForeColorDisabled = System.Drawing.Color.Black;
			appearance54.TextHAlign = Infragistics.Win.HAlign.Right;
			appearance54.TextVAlign = Infragistics.Win.VAlign.Middle;
			this.txtIdTransferenciaDT.Appearance = appearance54;
			this.txtIdTransferenciaDT.Enabled = false;
			this.txtIdTransferenciaDT.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.txtIdTransferenciaDT.FormatString = "";
			this.txtIdTransferenciaDT.Location = new System.Drawing.Point(1044, 528);
			this.txtIdTransferenciaDT.Name = "txtIdTransferenciaDT";
			this.txtIdTransferenciaDT.PromptChar = ' ';
			this.txtIdTransferenciaDT.Size = new System.Drawing.Size(16, 21);
			this.txtIdTransferenciaDT.TabIndex = 471;
			this.txtIdTransferenciaDT.Visible = false;
			// 
			// mnuImprimir
			// 
			this.mnuImprimir.MenuItems.AddRange(new System.Windows.Forms.MenuItem[] {
																																								this.menuItem1,
																																								this.menuItem6,
																																								this.menuItem2,
																																								this.menuItem3,
																																								this.menuItem5,
																																								this.menuItem4});
			// 
			// menuItem1
			// 
			this.menuItem1.Index = 0;
			this.menuItem1.Text = "Nota De Credito";
			this.menuItem1.Click += new System.EventHandler(this.menuItem1_Click);
			// 
			// menuItem6
			// 
			this.menuItem6.Index = 1;
			this.menuItem6.Text = "Nota De Credito Diferencias";
			this.menuItem6.Click += new System.EventHandler(this.menuItem6_Click);
			// 
			// menuItem2
			// 
			this.menuItem2.Index = 2;
			this.menuItem2.Text = "Transferencia BI";
			this.menuItem2.Click += new System.EventHandler(this.menuItem2_Click);
			// 
			// menuItem3
			// 
			this.menuItem3.Index = 3;
			this.menuItem3.Text = "Transferencia DT";
			this.menuItem3.Click += new System.EventHandler(this.menuItem3_Click);
			// 
			// menuItem5
			// 
			this.menuItem5.Index = 4;
			this.menuItem5.Text = "Transferencia CR";
			this.menuItem5.Visible = false;
			this.menuItem5.Click += new System.EventHandler(this.menuItem5_Click);
			// 
			// menuItem4
			// 
			this.menuItem4.Index = 5;
			this.menuItem4.Text = "Nota de Entrega";
			this.menuItem4.Click += new System.EventHandler(this.menuItem4_Click);
			// 
			// txtidTransferenciaRC
			// 
			this.txtidTransferenciaRC.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			appearance55.ForeColorDisabled = System.Drawing.Color.Black;
			appearance55.TextHAlign = Infragistics.Win.HAlign.Right;
			appearance55.TextVAlign = Infragistics.Win.VAlign.Middle;
			this.txtidTransferenciaRC.Appearance = appearance55;
			this.txtidTransferenciaRC.Enabled = false;
			this.txtidTransferenciaRC.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.txtidTransferenciaRC.FormatString = "";
			this.txtidTransferenciaRC.Location = new System.Drawing.Point(1028, 528);
			this.txtidTransferenciaRC.Name = "txtidTransferenciaRC";
			this.txtidTransferenciaRC.PromptChar = ' ';
			this.txtidTransferenciaRC.Size = new System.Drawing.Size(16, 21);
			this.txtidTransferenciaRC.TabIndex = 472;
			this.txtidTransferenciaRC.Visible = false;
			// 
			// label17
			// 
			this.label17.AutoSize = true;
			this.label17.BackColor = System.Drawing.Color.Transparent;
			this.label17.Location = new System.Drawing.Point(424, 107);
			this.label17.Name = "label17";
			this.label17.Size = new System.Drawing.Size(44, 16);
			this.label17.TabIndex = 474;
			this.label17.Text = "Número";
			this.label17.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.label17.Visible = false;
			// 
			// label22
			// 
			this.label22.AutoSize = true;
			this.label22.BackColor = System.Drawing.Color.Transparent;
			this.label22.Location = new System.Drawing.Point(904, 203);
			this.label22.Name = "label22";
			this.label22.Size = new System.Drawing.Size(74, 16);
			this.label22.TabIndex = 475;
			this.label22.Text = "Re liquidación";
			this.label22.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// txtProntoPago
			// 
			appearance56.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance56.ForeColor = System.Drawing.Color.Black;
			appearance56.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtProntoPago.Appearance = appearance56;
			this.txtProntoPago.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			this.txtProntoPago.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtProntoPago.Enabled = false;
			this.txtProntoPago.FormatString = "#,##0.00";
			this.txtProntoPago.Location = new System.Drawing.Point(992, 200);
			this.txtProntoPago.Name = "txtProntoPago";
			this.txtProntoPago.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtProntoPago.PromptChar = ' ';
			this.txtProntoPago.Size = new System.Drawing.Size(88, 22);
			this.txtProntoPago.TabIndex = 476;
			// 
			// label23
			// 
			this.label23.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			this.label23.AutoSize = true;
			this.label23.BackColor = System.Drawing.Color.Transparent;
			this.label23.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Bold);
			this.label23.ForeColor = System.Drawing.Color.Red;
			this.label23.Location = new System.Drawing.Point(941, 483);
			this.label23.Name = "label23";
			this.label23.Size = new System.Drawing.Size(62, 17);
			this.label23.TabIndex = 477;
			this.label23.Text = "Diferencia";
			this.label23.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.label23.Visible = false;
			// 
			// txtDiferencia
			// 
			this.txtDiferencia.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			appearance57.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance57.ForeColor = System.Drawing.Color.Red;
			appearance57.ForeColorDisabled = System.Drawing.Color.Red;
			this.txtDiferencia.Appearance = appearance57;
			this.txtDiferencia.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			this.txtDiferencia.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtDiferencia.Enabled = false;
			this.txtDiferencia.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Bold);
			this.txtDiferencia.FormatString = "#,##0.00";
			this.txtDiferencia.Location = new System.Drawing.Point(1068, 480);
			this.txtDiferencia.Name = "txtDiferencia";
			this.txtDiferencia.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtDiferencia.PromptChar = ' ';
			this.txtDiferencia.Size = new System.Drawing.Size(88, 22);
			this.txtDiferencia.TabIndex = 478;
			this.txtDiferencia.Visible = false;
			// 
			// txtIdNotaDeCredito1
			// 
			this.txtIdNotaDeCredito1.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			appearance58.ForeColorDisabled = System.Drawing.Color.Black;
			appearance58.TextHAlign = Infragistics.Win.HAlign.Right;
			appearance58.TextVAlign = Infragistics.Win.VAlign.Middle;
			this.txtIdNotaDeCredito1.Appearance = appearance58;
			this.txtIdNotaDeCredito1.Enabled = false;
			this.txtIdNotaDeCredito1.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.txtIdNotaDeCredito1.FormatString = "";
			this.txtIdNotaDeCredito1.Location = new System.Drawing.Point(1013, 528);
			this.txtIdNotaDeCredito1.Name = "txtIdNotaDeCredito1";
			this.txtIdNotaDeCredito1.PromptChar = ' ';
			this.txtIdNotaDeCredito1.Size = new System.Drawing.Size(16, 21);
			this.txtIdNotaDeCredito1.TabIndex = 479;
			this.txtIdNotaDeCredito1.Visible = false;
			// 
			// label24
			// 
			this.label24.AutoSize = true;
			this.label24.BackColor = System.Drawing.Color.Transparent;
			this.label24.Location = new System.Drawing.Point(944, 107);
			this.label24.Name = "label24";
			this.label24.Size = new System.Drawing.Size(140, 16);
			this.label24.TabIndex = 481;
			this.label24.Text = "Tiempo de espera del pago";
			this.label24.Visible = false;
			this.label24.Click += new System.EventHandler(this.label24_Click);
			// 
			// chkReservaMercaderia
			// 
			appearance59.ForeColorDisabled = System.Drawing.Color.Black;
			this.chkReservaMercaderia.Appearance = appearance59;
			this.chkReservaMercaderia.BackColor = System.Drawing.Color.Transparent;
			this.chkReservaMercaderia.Enabled = false;
			this.chkReservaMercaderia.Location = new System.Drawing.Point(768, 104);
			this.chkReservaMercaderia.Name = "chkReservaMercaderia";
			this.chkReservaMercaderia.Size = new System.Drawing.Size(152, 22);
			this.chkReservaMercaderia.TabIndex = 482;
			this.chkReservaMercaderia.Text = "En Garantia De Pago";
			this.chkReservaMercaderia.Visible = false;
			this.chkReservaMercaderia.CheckedChanged += new System.EventHandler(this.chkVerificar_CheckedChanged);
			// 
			// txtTiempoPago
			// 
			appearance60.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtTiempoPago.Appearance = appearance60;
			this.txtTiempoPago.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtTiempoPago.Enabled = false;
			this.txtTiempoPago.FormatString = "";
			this.txtTiempoPago.Location = new System.Drawing.Point(1096, 104);
			this.txtTiempoPago.MaskInput = "nn";
			this.txtTiempoPago.MaxValue = 365;
			this.txtTiempoPago.MinValue = 0;
			this.txtTiempoPago.Name = "txtTiempoPago";
			this.txtTiempoPago.PromptChar = ' ';
			this.txtTiempoPago.Size = new System.Drawing.Size(72, 22);
			this.txtTiempoPago.TabIndex = 480;
			this.txtTiempoPago.Visible = false;
			this.txtTiempoPago.Click += new System.EventHandler(this.txtTiempoPago_Click);
			// 
			// label25
			// 
			this.label25.AutoSize = true;
			this.label25.Location = new System.Drawing.Point(704, 83);
			this.label25.Name = "label25";
			this.label25.Size = new System.Drawing.Size(76, 16);
			this.label25.TabIndex = 483;
			this.label25.Text = "Bodega Actual";
			this.label25.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.label25.Visible = false;
			// 
			// cmbBodegaActual
			// 
			appearance61.ForeColor = System.Drawing.Color.Black;
			appearance61.ForeColorDisabled = System.Drawing.Color.Black;
			this.cmbBodegaActual.Appearance = appearance61;
			this.cmbBodegaActual.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbBodegaActual.Cursor = System.Windows.Forms.Cursors.Default;
			ultraGridColumn62.Header.VisiblePosition = 2;
			ultraGridColumn63.Header.VisiblePosition = 1;
			ultraGridColumn63.Width = 180;
			ultraGridColumn64.Header.VisiblePosition = 0;
			ultraGridBand7.Columns.AddRange(new object[] {
																										 ultraGridColumn62,
																										 ultraGridColumn63,
																										 ultraGridColumn64});
			this.cmbBodegaActual.DisplayLayout.BandsSerializer.Add(ultraGridBand7);
			this.cmbBodegaActual.DisplayMember = "Nombre";
			this.cmbBodegaActual.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbBodegaActual.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbBodegaActual.Enabled = false;
			this.cmbBodegaActual.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbBodegaActual.Location = new System.Drawing.Point(800, 80);
			this.cmbBodegaActual.MaxDropDownItems = 30;
			this.cmbBodegaActual.Name = "cmbBodegaActual";
			this.cmbBodegaActual.Size = new System.Drawing.Size(256, 22);
			this.cmbBodegaActual.TabIndex = 484;
			this.cmbBodegaActual.ValueMember = "Bodega";
			this.cmbBodegaActual.Visible = false;
			this.cmbBodegaActual.KeyDown += new System.Windows.Forms.KeyEventHandler(this.cmbBodegaActual_KeyDown);
			// 
			// frmDTIngresoPorRetiro
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 14);
			this.ClientSize = new System.Drawing.Size(1165, 558);
			this.Controls.Add(this.label25);
			this.Controls.Add(this.label24);
			this.Controls.Add(this.label23);
			this.Controls.Add(this.label22);
			this.Controls.Add(this.label17);
			this.Controls.Add(this.label21);
			this.Controls.Add(this.label20);
			this.Controls.Add(this.label19);
			this.Controls.Add(this.lblVendedor);
			this.Controls.Add(this.lblSubTotal);
			this.Controls.Add(this.lblIva);
			this.Controls.Add(this.lblDescuento);
			this.Controls.Add(this.label18);
			this.Controls.Add(this.label15);
			this.Controls.Add(this.label14);
			this.Controls.Add(this.label13);
			this.Controls.Add(this.label11);
			this.Controls.Add(this.label12);
			this.Controls.Add(this.label10);
			this.Controls.Add(this.label9);
			this.Controls.Add(this.label7);
			this.Controls.Add(this.label16);
			this.Controls.Add(this.label3);
			this.Controls.Add(this.label2);
			this.Controls.Add(this.label1);
			this.Controls.Add(this.label8);
			this.Controls.Add(this.label6);
			this.Controls.Add(this.label5);
			this.Controls.Add(this.label36);
			this.Controls.Add(this.label32);
			this.Controls.Add(this.label4);
			this.Controls.Add(this.cmbBodegaActual);
			this.Controls.Add(this.chkReservaMercaderia);
			this.Controls.Add(this.txtTiempoPago);
			this.Controls.Add(this.txtIdNotaDeCredito1);
			this.Controls.Add(this.txtDiferencia);
			this.Controls.Add(this.txtProntoPago);
			this.Controls.Add(this.txtNumeroNotaDeCredito);
			this.Controls.Add(this.txtidTransferenciaRC);
			this.Controls.Add(this.txtIdTransferenciaDT);
			this.Controls.Add(this.btnProcesar);
			this.Controls.Add(this.txtTiempoUso);
			this.Controls.Add(this.cmbRetiradoPor);
			this.Controls.Add(this.txtNLote);
			this.Controls.Add(this.cmbVendedor);
			this.Controls.Add(this.txtIvaTotal);
			this.Controls.Add(this.txtDescIvaTotal);
			this.Controls.Add(this.txtIva);
			this.Controls.Add(this.cmbMotivos);
			this.Controls.Add(this.chkFiscal);
			this.Controls.Add(this.txtidTransferencia);
			this.Controls.Add(this.txtIdNotaDeCredito);
			this.Controls.Add(this.txtValorNC);
			this.Controls.Add(this.cmbEstados);
			this.Controls.Add(this.groupBox2);
			this.Controls.Add(this.txtEntidadFinanciera);
			this.Controls.Add(this.txtTotalCredito);
			this.Controls.Add(this.txtIdCompra);
			this.Controls.Add(this.txtLocal);
			this.Controls.Add(this.txtVendedor);
			this.Controls.Add(this.txtTotalFactura);
			this.Controls.Add(this.txtPagado);
			this.Controls.Add(this.dtFechaRetiro);
			this.Controls.Add(this.txtSaldo);
			this.Controls.Add(this.txtDiasDeMora);
			this.Controls.Add(this.txtEstado);
			this.Controls.Add(this.txtBodegaPredef);
			this.Controls.Add(this.txtIdIngreso);
			this.Controls.Add(this.btnNuevo);
			this.Controls.Add(this.btnConsultar);
			this.Controls.Add(this.btnEditar);
			this.Controls.Add(this.btnImprimir);
			this.Controls.Add(this.btnGuardar);
			this.Controls.Add(this.btnCancelar);
			this.Controls.Add(this.btnAnular);
			this.Controls.Add(this.btnSalir);
			this.Controls.Add(this.groupBox5);
			this.Controls.Add(this.txtNumeroFactura);
			this.Controls.Add(this.dtEmision);
			this.Controls.Add(this.txtObservaciones);
			this.Controls.Add(this.txtCodigo);
			this.Controls.Add(this.txtNombre);
			this.Controls.Add(this.groupBox1);
			this.Controls.Add(this.dtFecha);
			this.Controls.Add(this.lblEstado);
			this.Controls.Add(this.cmbBodega);
			this.Controls.Add(this.lblNumero);
			this.Controls.Add(this.uGridDTIngreso);
			this.KeyPreview = true;
			this.Name = "frmDTIngresoPorRetiro";
			this.Text = "Ingreso Mercaderia de Retiro";
			this.KeyDown += new System.Windows.Forms.KeyEventHandler(this.frmDTIngresoPorRetiro_KeyDown);
			this.Closing += new System.ComponentModel.CancelEventHandler(this.frmDTIngresoPorRetiro_Closing);
			this.Load += new System.EventHandler(this.frmDTIngresoPorRetiro_Load);
			((System.ComponentModel.ISupportInitialize)(this.uGridDTIngreso)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource1)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.dtFecha)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbBodega)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtObservaciones)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtCodigo)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNombre)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.dtEmision)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNumeroFactura)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtEstado)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtBodegaPredef)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIdIngreso)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtDiasDeMora)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtSaldo)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.dtFechaRetiro)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtPagado)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtTotalFactura)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtLocal)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtVendedor)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource2)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIdCompra)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtTotalCredito)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtEntidadFinanciera)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNumeroNotaDeCredito)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbEstados)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource3)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtValorNC)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtidTransferencia)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIdNotaDeCredito)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbMotivos)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource4)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIvaTotal)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtDescIvaTotal)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIva)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbVendedor)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNLote)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbRetiradoPor)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtTiempoUso)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIdTransferenciaDT)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtidTransferenciaRC)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtProntoPago)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtDiferencia)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIdNotaDeCredito1)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtTiempoPago)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbBodegaActual)).EndInit();
			this.ResumeLayout(false);

		}
		#endregion

		bool bNuevo = false;
		bool bEdicion = false;
		private Acceso miAcceso;
		int idEntidadFinanciera = 0;

		private void btnSalir_Click(object sender, System.EventArgs e)
		{
			this.Close();
		}

		private void frmDTIngresoPorRetiro_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (!bEdicion)
			{
				if (e.KeyCode == Keys.N) if (miAcceso.Nuevo)	if (this.btnNuevo.Enabled) this.btnNuevo_Click(sender, e);
				if (e.KeyCode == Keys.B) if (miAcceso.Buscar) if (this.btnConsultar.Enabled) this.btnConsultar_Click(sender, e);
				if (e.KeyCode == Keys.E) if (miAcceso.Editar) if ((int)this.txtEstado.Value == 0) this.btnEditar_Click(sender, e); 																											
				if (e.KeyCode == Keys.I) if (miAcceso.Imprimir) if (this.btnImprimir.Enabled) this.btnImprimir_Click(sender, e);
				if (e.KeyCode == Keys.A) if (miAcceso.Anular) if ((int)this.txtEstado.Value == 0) this.btnAnular_Click(sender, e); 																																	 
				if (e.KeyCode == Keys.F5)
				{
					miAcceso = new Acceso(cdsSeteoF, "1004");

					this.Consulta((int)this.txtIdIngreso.Value);					
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

		private void btnCancelar_Click(object sender, System.EventArgs e)
		{
			bNuevo = false;
			bEdicion = false;

			this.txtIdIngreso.Value = 0;
			this.txtIdCompra.Value = 0;
			this.txtBodegaPredef.Value = 0;
			this.txtEstado.Value = 0;
			this.txtIdNotaDeCredito.Value = 0;
			this.txtIdNotaDeCredito1.Value = 0;
			this.txtIdTransferenciaDT.Value = 0;
			this.txtidTransferenciaRC.Value = 0;
			this.txtidTransferencia.Value = 0;
			this.txtNumeroNotaDeCredito.Text = "";

			this.cmbBodega.Value = System.DBNull.Value;
			this.cmbBodegaActual.Value = System.DBNull.Value;
			this.dtFecha.Value = System.DBNull.Value;
			this.cmbVendedor.Value = System.DBNull.Value;
			this.cmbRetiradoPor.Value = System.DBNull.Value;
			this.lblNumero.Text = "";

			this.txtNumeroFactura.Text = "";
			this.dtFechaRetiro.Value = System.DBNull.Value;
			this.chkReservaMercaderia.Checked = false;
			this.txtTiempoPago.Value = 0;
			
			this.chkFiscal.Checked = false;
			this.cmbMotivos.Value = System.DBNull.Value;
			this.txtCodigo.Text = "";
			this.txtNombre.Text = "";
			this.txtEntidadFinanciera.Text = "";
			this.dtEmision.Value = System.DBNull.Value;
			this.txtLocal.Text = "";
			this.txtVendedor.Text = "";
			
			this.txtDiasDeMora.Value = 0;
			this.txtSaldo.Value = 0;
			this.txtPagado.Value = 0;
			this.txtNLote.Value = 0;
			this.txtTiempoUso.Value = 0;
			this.txtTotalFactura.Value = 0.00m;
			this.txtTotalCredito.Value = 0.00m;
			this.txtDescIvaTotal.Value = 0.00m;
			this.txtValorNC.Value = 0.00m;
			this.txtProntoPago.Value = 0.00m;
			this.txtIva.Value = 0.00m;
			this.txtIvaTotal.Value = 0.00m;
			
			this.txtObservaciones.Text = "";
			this.lblEstado.Text = "";
			
			this.cmbRetiradoPor.Enabled = false;
			this.cmbBodega.Enabled = false;
			this.cmbBodegaActual.Enabled = false;
			this.dtFecha.Enabled = false;
			this.txtNumeroFactura.Enabled = false;
			this.dtFechaRetiro.Enabled = false;
			this.chkReservaMercaderia.Enabled = false;
			this.txtTiempoPago.Enabled = false;							
			this.chkFiscal.Enabled = false;
			this.cmbMotivos.Enabled = false;
			this.txtDescIvaTotal.Enabled = false;
			this.txtObservaciones.Enabled = false;
			this.txtValorNC.Enabled = false;
			this.txtTiempoUso.Enabled = false;
			this.cmbVendedor.Enabled = false;
			this.cmbVendedor.Enabled = false;
			this.txtNumeroNotaDeCredito.Enabled = false;

			FuncionesProcedimientos.EstadoGrids(false, this.uGridDTIngreso);
			this.uGridDTIngreso.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "Exec DTConsultaDetalleFacturasDeRetiro 0");

			if (miAcceso.Nuevo) this.btnNuevo.Enabled = true;
			if (miAcceso.Buscar) this.btnConsultar.Enabled = true;
			this.btnEditar.Enabled = false;
			this.btnImprimir.Enabled = false;
			this.btnGuardar.Enabled = false;
			this.btnProcesar.Enabled = false;
			this.btnAnular.Enabled = false;
			this.btnCancelar.Enabled = false;
			
			idEntidadFinanciera = 0;
		}

		private void btnNuevo_Click(object sender, System.EventArgs e)
		{
			this.btnCancelar_Click(sender, e);

			using (frmDTBuscaFacturaRetiro Busqueda = new frmDTBuscaFacturaRetiro())
			{
				if (DialogResult.OK == Busqueda.ShowDialog())
				{
					this.txtIdCompra.Value = (int)Busqueda.uGridFacturas.ActiveRow.Cells["idCompra"].Value;
					this.txtNumeroFactura.Text = Busqueda.uGridFacturas.ActiveRow.Cells["Numero"].Value.ToString();
					this.txtCodigo.Text = Busqueda.uGridFacturas.ActiveRow.Cells["RUC"].Value.ToString();
					this.txtNombre.Text = Busqueda.uGridFacturas.ActiveRow.Cells["Nombre"].Value.ToString();
					idEntidadFinanciera = (int)Busqueda.uGridFacturas.ActiveRow.Cells["idEntidadFinanciera"].Value;
					this.txtEntidadFinanciera.Text = Busqueda.uGridFacturas.ActiveRow.Cells["EntidadFinanciera"].Value.ToString();
					this.dtEmision.Value = (DateTime)Busqueda.uGridFacturas.ActiveRow.Cells["Fecha"].Value;
					this.txtLocal.Text = Busqueda.uGridFacturas.ActiveRow.Cells["Local"].Value.ToString();
					this.txtVendedor.Text = Busqueda.uGridFacturas.ActiveRow.Cells["Vendedor"].Value.ToString();
					this.txtNLote.Value = (int)Busqueda.uGridFacturas.ActiveRow.Cells["NumeroLote"].Value;
					this.txtDiasDeMora.Value = (int)Busqueda.uGridFacturas.ActiveRow.Cells["DiasDeMora"].Value;
					this.txtSaldo.Value = (decimal)Busqueda.uGridFacturas.ActiveRow.Cells["SaldoAnticipo"].Value;
					this.txtPagado.Value = (decimal)Busqueda.uGridFacturas.ActiveRow.Cells["Pagado"].Value;
					this.txtTotalFactura.Value = (decimal)Busqueda.uGridFacturas.ActiveRow.Cells["Total"].Value;
					this.txtTotalCredito.Value = (decimal)Busqueda.uGridFacturas.ActiveRow.Cells["TotalCuotas"].Value;

					this.cmbMotivos.Value = (int)Busqueda.cmbMotivos.Value;

					#region Valida Vendedor Bodega
					if ((int)Busqueda.cmbMotivos.Value == 4)
					{
						if (Funcion.iEscalarSQL(cdsSeteoF, string.Format("Select DATEDIFF(DAY, '{0}', '{1}')", Convert.ToDateTime(this.dtEmision.Value).ToString("yyyyMMdd"), DateTime.Today.ToString("yyyyMMdd"))) < 181)
						{
							this.cmbBodega.Value = (int)Busqueda.uGridFacturas.ActiveRow.Cells["Bodega"].Value;
							this.cmbVendedor.Value = (int)Busqueda.uGridFacturas.ActiveRow.Cells["idVendedor"].Value;
						}
						else
						{
							this.cmbBodega.Value = 18;
							this.cmbVendedor.Value = 205;
						}
					}
					else
					{
						this.cmbBodega.Value = (int)Busqueda.uGridFacturas.ActiveRow.Cells["Bodega"].Value;
						this.cmbVendedor.Value = (int)Busqueda.uGridFacturas.ActiveRow.Cells["idVendedor"].Value;
					}
					#endregion Valida Vendedor Bodega	

					bNuevo = true;
					bEdicion = true;

					//this.cmbBodegaActual.Enabled = true;
					this.chkFiscal.Checked = true;
					this.dtFechaRetiro.Value = DateTime.Today;
					this.dtFecha.Value = DateTime.Today;

					this.uGridDTIngreso.DataSource = Funcion.dvProcedimiento(cdsSeteoF, string.Format("Exec DTConsultaDetalleFacturasDeRetiro {0}", (int)this.txtIdCompra.Value));
						
					this.dtFechaRetiro.CalendarInfo.MinDate = (DateTime)this.dtEmision.Value;
					
					Total();
										
					this.dtFechaRetiro.Enabled = true;
					this.cmbRetiradoPor.Enabled = true;			
													
					//if (miAcceso.CambiarFecha) this.dtFecha.Enabled = true;
					//this.chkReservaMercaderia.Enabled = true;					
					this.txtObservaciones.Enabled = true;

					FuncionesProcedimientos.EstadoGrids(true, this.uGridDTIngreso);

					this.btnNuevo.Enabled = false;
					this.btnConsultar.Enabled = false;
					this.btnGuardar.Enabled = true;			
					this.btnCancelar.Enabled = true;													

					#region Valor re liquidacion
					if ((int)this.cmbMotivos.Value == 4)
					{
						if (idEntidadFinanciera != 2)
						{
							string sSQLRel = string.Format("Exec ConsultaValorReliquidacion {0}, '{1}', {2}, {3}, 2", 
								(int)this.txtIdCompra.Value, Convert.ToDateTime(this.dtFecha.Value).ToString("yyyyMMdd"), 0.00m, 0.00m);
							SqlDataReader drR = Funcion.rEscalarSQL(cdsSeteoF, sSQLRel, true);
							drR.Read();
							if(drR.HasRows)
							{
								this.txtProntoPago.Value = drR.GetDecimal(0);								
							}
							drR.Close();

							CalculaDiferencia();
						}						
					}
					#endregion Valor re liquidacion
					
					this.dtFechaRetiro.Focus();
				}
			}
		}

		private void frmDTIngresoPorRetiro_Load(object sender, System.EventArgs e)
		{
			miAcceso = new Acceso(cdsSeteoF, "1004");

			if (miAcceso.Nuevo) this.btnNuevo.Enabled = true;
			if (miAcceso.Buscar) this.btnConsultar.Enabled = true;

			this.dtFecha.CalendarInfo.MaxDate = DateTime.Today;
			this.dtFechaRetiro.CalendarInfo.MaxDate = DateTime.Today;

			FuncionesProcedimientos.EstadoGrids(false, this.uGridDTIngreso);

			this.cmbBodega.DataSource = Funcion.dvProcedimiento(cdsSeteoF, string.Format("Exec ListaBodegas {0}, 54", MenuLatinium.IdUsuario));

			FuncionesProcedimientos.ListaBodegas(this.cmbBodegaActual, MenuLatinium.IdUsuario, 47, DateTime.Today, false, cdsSeteoF);

			this.cmbMotivos.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "Select idMotivo, Motivo From MotivosDevolucion Where idTipoFactura = 5 And idMotivo In (4, 18) Order By Motivo");

			this.cmbVendedor.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "Exec PersonalVendedor 0");
			this.cmbRetiradoPor.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "Select idPersonal, Codigo, Nombre From Personal Where idCargo In (2, 36, 37, 38) And FechaSalida IS NULL Order By Nombre");

			string sSQL = string.Format("Exec BodegaAsignadaPorUsuario {0}", MenuLatinium.IdUsuario);		
			this.txtBodegaPredef.Value = Funcion.iEscalarSQL(cdsSeteoF, sSQL);

			this.cmbEstados.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "SELECT idDTEstado, Descripcion FROM DTEstados");

			if (idIngresoRetiro > 0) this.Consulta(idIngresoRetiro);
		}

		private void Consulta(int idIngreso)
		{
			this.btnNuevo.Enabled = false;
			this.btnConsultar.Enabled = false;
			this.btnEditar.Enabled = false;
			this.btnImprimir.Enabled = false;								
			this.btnAnular.Enabled = false;																
			this.btnProcesar.Enabled = false;

			#region Informacion General
			string sSQL = string.Format("Exec DTConsultaRetiroIndividual {0}", idIngreso);
			SqlDataReader dr = Funcion.rEscalarSQL(cdsSeteoF, sSQL, true);
			dr.Read();
			if (dr.HasRows)
			{
				this.txtIdIngreso.Value = idIngreso;
				this.txtIdCompra.Value = dr.GetInt32(1);
				this.txtEstado.Value = dr.GetInt32(2);
				this.cmbBodega.Value = dr.GetInt32(3);
				this.lblNumero.Text = dr.GetString(4);
				this.dtFecha.Value = dr.GetDateTime(5);
				this.txtNumeroFactura.Text = dr.GetString(6); 
				this.dtFechaRetiro.Value = dr.GetDateTime(7);
				this.txtCodigo.Text = dr.GetString(8);
				this.txtNombre.Text = dr.GetString(9);
				this.txtEntidadFinanciera.Text = dr.GetString(10);
				this.dtEmision.Value = dr.GetDateTime(11);
				this.txtLocal.Text = dr.GetString(12);
				this.txtVendedor.Text = dr.GetString(13);
				this.txtDiasDeMora.Value = dr.GetInt32(14);
				this.txtSaldo.Value = Convert.ToDecimal(dr.GetValue(15));
				this.txtPagado.Value = Convert.ToDecimal(dr.GetValue(16));
				this.txtTotalCredito.Value = Convert.ToDecimal(dr.GetValue(17));
				this.txtTotalFactura.Value = Convert.ToDecimal(dr.GetValue(18));
				this.txtDescIvaTotal.Value = Convert.ToDecimal(dr.GetValue(19));
				this.txtValorNC.Value = Convert.ToDecimal(dr.GetValue(20));
				this.txtObservaciones.Text = dr.GetValue(21).ToString();
				this.txtIdNotaDeCredito.Value = dr.GetInt32(22);
				this.txtidTransferencia.Value = dr.GetInt32(23);
				this.txtIdTransferenciaDT.Value = dr.GetInt32(24);
				this.cmbVendedor.Value = dr.GetInt32(25);
				this.txtTiempoUso.Value = dr.GetInt32(26);
				this.cmbRetiradoPor.Value = dr.GetInt32(27);
				this.cmbMotivos.Value = dr.GetInt32(28);
				this.txtNLote.Value = dr.GetInt32(29);
				this.txtidTransferenciaRC.Value = dr.GetInt32(30);
				this.txtNumeroNotaDeCredito.Value = dr.GetValue(31).ToString();
				this.chkFiscal.Checked = dr.GetBoolean(32);
				this.txtIdNotaDeCredito1.Value = dr.GetInt32(33);
				this.txtProntoPago.Value = Convert.ToDecimal(dr.GetValue(34));
				this.txtDiferencia.Value = Convert.ToDecimal(dr.GetValue(35));
				this.chkReservaMercaderia.Checked = dr.GetBoolean(36);
				this.txtDiferencia.Value = dr.GetInt32(37);
				this.cmbBodegaActual.Value = dr.GetInt32(38);
			}
			dr.Close();
			#endregion Informacion General

			this.uGridDTIngreso.DataSource = Funcion.dvProcedimiento(cdsSeteoF, string.Format("Exec DTConsultaDetalleRetiro {0}", idIngreso));

			//this.Total();

			if ((int)this.txtEstado.Value == 0) this.lblEstado.Text = "PENDIENTE";
			if ((int)this.txtEstado.Value == 1) this.lblEstado.Text = "PROCESADO";
			if ((int)this.txtEstado.Value == 2) this.lblEstado.Text = "ANULADO";

			#region Controles
			if (miAcceso.Nuevo) this.btnNuevo.Enabled = true;
			if (miAcceso.Buscar) this.btnConsultar.Enabled = true;

			if (miAcceso.Editar) if ((int)this.txtEstado.Value == 0) this.btnEditar.Enabled = true;
			if (miAcceso.Imprimir) if ((int)this.txtEstado.Value == 1) this.btnImprimir.Enabled = true;
			if (miAcceso.BProcesar) if ((int)this.txtEstado.Value == 0) this.btnProcesar.Enabled = true;
			if (miAcceso.Anular) if ((int)this.txtEstado.Value == 0) this.btnAnular.Enabled = true;

			this.btnCancelar.Enabled = true;			
			#endregion Controles
		}

		private void btnConsultar_Click(object sender, System.EventArgs e)
		{
			this.btnCancelar_Click(sender, e);

			using (frmConsultaRetiros Busqueda = new frmConsultaRetiros())
			{
				if (DialogResult.OK == Busqueda.ShowDialog())
				{
					this.Consulta((int)Busqueda.ultraGrid1.ActiveRow.Cells["idIngresoRetiro"].Value);
				}
			}
		}

		private void btnEditar_Click(object sender, System.EventArgs e)
		{			
			this.Consulta((int)this.txtIdIngreso.Value);

			if ((int)this.txtEstado.Value == 1)
			{
				MessageBox.Show("El documento ya fue procesado", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				return;
			}
			if ((int)this.txtEstado.Value == 2)
			{
				MessageBox.Show("El documento esta anulado", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				return;
			}

			if (miAcceso.CambiarFecha) this.dtFecha.Enabled = true;
		
			this.dtFechaRetiro.Enabled = true;
			this.cmbRetiradoPor.Enabled = true;
			this.cmbMotivos.Enabled = true;
			this.chkReservaMercaderia.Enabled = true;			
			this.txtObservaciones.Enabled = true;
			
			if (!this.chkFiscal.Checked)
			{
				this.cmbBodega.Enabled = true;
				this.cmbVendedor.Enabled = true;
			}

			this.cmbBodegaActual.Enabled = true;

			if (this.chkFiscal.Checked) if ((int)this.txtEstado.Value == 0) this.txtNumeroNotaDeCredito.Enabled = true;
			
			FuncionesProcedimientos.EstadoGrids(true, this.uGridDTIngreso);

			this.btnNuevo.Enabled = false;
			this.btnConsultar.Enabled = false;
			this.btnEditar.Enabled = false;
			this.btnImprimir.Enabled = false;
			this.btnGuardar.Enabled = true;	
			this.btnProcesar.Enabled = false;
			this.btnAnular.Enabled = false;
			this.btnCancelar.Enabled = true;
			
			bNuevo = false;
			bEdicion = true;

			this.dtFechaRetiro.Focus();
		}

		private void txtNumeroFactura_ValueChanged(object sender, System.EventArgs e)
		{
		
		}

		private void CalculaDiferencia()
		{
			this.txtDiferencia.Value = 0.00m;

			if (idEntidadFinanciera != 2)
			{				
				if (this.cmbMotivos.ActiveRow != null)
				{
					if ((int)this.cmbMotivos.Value != 40)
					{
						decimal dDiferencia = Convert.ToDecimal(this.txtProntoPago.Value) - Convert.ToDecimal(this.txtTotalFactura.Value);

						if (decimal.Round(dDiferencia, 2) > 0.00m)
							this.txtDiferencia.Value = decimal.Round(dDiferencia, 2);
					}
				}
			}
		}

		private void txtNumeroFactura_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			
		}

		private void cdsSeteoF_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			cdsSeteoF.Schema.Connections[0].ConnectionString = miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);	
		}

		private void dtFechaRetiro_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Delete) e.Handled = true;
			if (e.KeyCode == Keys.Enter) this.cmbRetiradoPor.Focus();
		}

		private void dtFechaRetiro_Validating(object sender, System.ComponentModel.CancelEventArgs e)
		{
			if (bEdicion)
			{
				if (this.dtFechaRetiro.Value == System.DBNull.Value)
				{
					MessageBox.Show("Seleccione la Fecha de Retiro", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
					this.dtFechaRetiro.Focus();
					e.Cancel = true; 
				}
			}
		}

		private void uGridDTIngreso_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (bEdicion)
			{
				if (e.KeyCode == Keys.F3)				
					if (this.uGridDTIngreso.ActiveCell.Column.Key.ToString().ToUpper() == "CODIGO")
						if ((int)this.uGridDTIngreso.ActiveRow.Cells["idArticuloGE"].Value == 0)
							FuncionesInventario.AbreBusquedaArticulos(this.uGridDTIngreso, (int)this.cmbBodega.Value, 54, 9, Convert.ToDateTime(this.dtFecha.Value), false);

				FuncionesProcedimientos.DesplazamientoGrids(sender, e, this.uGridDTIngreso);	
			}
		}
		
		private void btnGuardar_Click(object sender, System.EventArgs e)
		{	
			if (!Validacion.vbFechaEnDocumentos(this.dtFecha, "Seleccione la Fecha de Retiro", false, "Venta", cdsSeteoF)) return;

			if (!Validacion.vbComparaFechas(Convert.ToDateTime(this.dtFechaRetiro.Value), Convert.ToDateTime(this.dtFecha.Value), this.dtFechaRetiro, "La fecha de retiro no puede ser mayor a la fecha del documento", false)) return;
												
//			if (this.chkFiscal.Checked)
//				if (!Validacion.vbNumeracion(false, (int)this.txtIdNotaDeCredito.Value, 5, (int)this.cmbBodega.Value, this.txtNumeroNotaDeCredito, cdsSeteoF)) return;
			
			if (!Validacion.vbComboVacio(this.cmbRetiradoPor, "Seleccione un Verificador")) return;
			if (!Validacion.vbComboVacio(this.cmbMotivos, "Seleccione un Motivo")) return;
			
			if (!this.chkFiscal.Checked)
			{
				if (!Validacion.vbComboVacio(this.cmbBodega, "Seleccione un Local")) return;
				if (!Validacion.vbComboVacio(this.cmbVendedor, "Seleccione un Vendedor")) return;				
			}

			if (!Validacion.vbComboVacio(this.cmbBodegaActual, "Seleccione la bodega actual del inventario")) return;

			if (this.chkReservaMercaderia.Checked)
				if (!Validacion.vbCampoEnteroVacio(this.txtTiempoPago, "Si el Cliente reserva la mercaderia el tiempo de pago en dias no puede ser igual cero", 1, 365)) return;

			if (!Validacion.ValidaTotales((decimal)this.txtValorNC.Value)) return;

			if (!Validacion.vbTexto(this.txtObservaciones, 20, 200, "Notas")) return;

			if (!Validacion.vbGrids(this.uGridDTIngreso, "Codigo", 54)) return;

			#region Observaciones
			bool bMensaje = false;
			foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.uGridDTIngreso.Rows.All)
			{	
				if (dr.Cells["idArticulo"].Value != System.DBNull.Value)
				{
					if (dr.Cells["Observaciones"].Value.ToString().Length == 0)
						{
							this.uGridDTIngreso.ActiveRow = dr;
							this.uGridDTIngreso.ActiveRow.Selected = true;
							bMensaje = true;
							continue;
						}
				}
			}

			if (bMensaje)
			{
				MessageBox.Show("Ingrese una observación en los articulos resaltados", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				return;
			}
			#endregion Observaciones

			if (this.chkFiscal.Checked && !Validacion.vbValidaValorDctoOrigen((int)this.txtIdCompra.Value, Convert.ToDecimal(this.txtValorNC.Value), cdsSeteoF)) return;
			
			#region Seriales
			bool bSeriales = false;

			foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.uGridDTIngreso.Rows.All)
			{
				if (dr.Cells["idArticulo"].Value != System.DBNull.Value)
				{
					if (Funcion.bEscalarSQL(cdsSeteoF, string.Format("Select ManejaSerial From Articulo Where idArticulo = {0}", (int)dr.Cells["idArticulo"].Value), true))
					{
						if (dr.Cells["Serial"].Value.ToString().Length == 0)
						{
							this.uGridDTIngreso.ActiveRow = dr;
							this.uGridDTIngreso.ActiveRow.Selected = true;
							bSeriales = true;
							continue;
						}
					}
				}
			}

			if (bSeriales)
			{
				MessageBox.Show("Ingrese la serial de los articulos resaltados", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				return;
			}
			#endregion Seriales

			#region Estado
			bool bEstado = false;

			foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.uGridDTIngreso.Rows.All)
			{
				if (dr.Cells["idArticulo"].Value != System.DBNull.Value)
				{
					int iEstado = 0;
					if (dr.Cells["idDTEstado"].Value != System.DBNull.Value) iEstado = (int) dr.Cells["idDTEstado"].Value; 

					if (iEstado == 0) 
					{							
						this.uGridDTIngreso.ActiveRow = dr;
						this.uGridDTIngreso.ActiveRow.Selected = true;
						bEstado = true;
						continue;									
					}						
				}
			}
			#endregion Estado

			if (bEstado)
			{
				MessageBox.Show("Seleccione el Estado de los Articulos Resaltados", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				return;
			}
			
			using (SqlConnection oConexion = new SqlConnection("Password=Bl45o6$9;Persist Security Info=True;User ID= " + MenuLatinium.stUsuario + ";Initial Catalog=" + MenuLatinium.stDirFacturacion + ";Data Source=" + MenuLatinium.stDirServidor))
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
					if (bNuevo)
					{
						string sSQLNumero = string.Format("Exec NumeracionLocales 54, {0}, 0", (int)this.cmbBodega.Value);
						oCmdActualiza.CommandText = sSQLNumero;
						this.lblNumero.Text = oCmdActualiza.ExecuteScalar().ToString();
					}	
					#endregion Numeracion
			
					#region Graba Registro
					#region Maestro
					DateTime dtFecha = (DateTime)this.dtFecha.Value;
					DateTime dtFechaRetiro = (DateTime)this.dtFechaRetiro.Value;

					string sSQLGuarda = string.Format("Exec DTGuardaIngresoPorRetiro {0}, {1}, {2}, {3}, '{4}', '{5}', '{6}', '{7}', {8}, {9}, {10}, {11}, {12}, {13}, {14}, '{15}', {16}, {17}, {18}, {19}, {20}", 
						(int) this.txtIdIngreso.Value, (int) this.txtIdCompra.Value, (int)this.txtEstado.Value, (int)this.cmbBodega.Value, dtFecha.ToString("yyyyMMdd"), this.lblNumero.Text.ToString(),
						dtFechaRetiro.ToString("yyyyMMdd"), this.txtObservaciones.Text.ToString(), Convert.ToDecimal(this.txtDescIvaTotal.Value), Convert.ToDecimal(this.txtValorNC.Value),
						(bool)this.chkFiscal.Checked, (int)this.cmbMotivos.Value, (int)this.cmbVendedor.Value, (int)this.cmbRetiradoPor.Value, (int)this.txtTiempoUso.Value, 
						this.txtNumeroNotaDeCredito.Text.ToString(), Convert.ToDecimal(this.txtProntoPago.Value), Convert.ToDecimal(this.txtDiferencia.Value),
						(bool)this.chkReservaMercaderia.Checked, (int)this.txtTiempoPago.Value, (int)this.cmbBodegaActual.Value);
					oCmdActualiza.CommandText = sSQLGuarda;
					this.txtIdIngreso.Value = (int)oCmdActualiza.ExecuteScalar();
					#endregion Maestro

					#region Detalle
					foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.uGridDTIngreso.Rows.All)
					{							
						string sSQLDetalle = string.Format("Exec DTGuardaDetalleRetiro {0}, {1}, {2}, {3}, '{4}', {5}, {6}, {7}, {8}, {9}, {10}, {11}, {12}, {13}, {14}, {15}, {16}, {17}, {18}, {19}, {20}, {21}, {22}, {23}, {24}, {25}, {26}, {27}, {28}, {29}, {30}, {31}, {32}, '{33}', {34}, {35}, {36}, {37}, {38}, {39}, {40}",
							(int)dr.Cells["idDetalleIngresoRetiro"].Value, (int)this.txtIdIngreso.Value, (int)dr.Cells["idDetCompra"].Value, 
							(int)dr.Cells["idDTEstado"].Value, dr.Cells["Observaciones"].Value.ToString(), (bool)dr.Cells["Sel"].Value,
							(int)dr.Cells["idArticulo"].Value, (int)dr.Cells["idSerial"].Value, Convert.ToDecimal(dr.Cells["Subtotal"].Value),
							(int)dr.Cells["idAccesorio1"].Value, (int)dr.Cells["idAccesorio2"].Value, (int)dr.Cells["idAccesorio3"].Value, (int)dr.Cells["idAccesorio4"].Value,
							(int)dr.Cells["idAccesorio5"].Value, (int)dr.Cells["idAccesorio6"].Value, (int)dr.Cells["idAccesorio7"].Value, (int)dr.Cells["idAccesorio8"].Value,
							(int)dr.Cells["idAccesorio9"].Value, (int)dr.Cells["idAccesorio10"].Value, (int)dr.Cells["idAccesorio11"].Value, (int)dr.Cells["idAccesorio12"].Value,
							(int)dr.Cells["idAccesorio13"].Value, (int)dr.Cells["idAccesorio14"].Value, (int)dr.Cells["idAccesorio15"].Value, (int)dr.Cells["idAccesorio16"].Value,
							(int)dr.Cells["idAccesorio17"].Value, (int)dr.Cells["idAccesorio18"].Value, (int)dr.Cells["idAccesorio19"].Value, (int)dr.Cells["idAccesorio20"].Value,
							Convert.ToInt32(dr.Cells["Cantidad"].Value), Convert.ToDecimal(dr.Cells["Precio"].Value), 
							Convert.ToDecimal(dr.Cells["DescuentoArt"].Value), Convert.ToDecimal(dr.Cells["DescuentoPorc"].Value),
							dr.Cells["Serial"].Value.ToString(), (int)dr.Cells["idArticuloGE"].Value,
							Convert.ToDecimal(dr.Cells["CostoUltimo"].Value), Convert.ToDecimal(dr.Cells["Cobrado"].Value), Convert.ToDecimal(dr.Cells["CostoFactura"].Value),
							Convert.ToInt32(dr.Cells["Cuotas_Pagadas"].Value), Convert.ToInt32(dr.Cells["Cuotas_Vencidas"].Value), Convert.ToInt32(dr.Cells["DiasAtraso"].Value));
								
						oCmdActualiza.CommandText = sSQLDetalle;
						oCmdActualiza.ExecuteNonQuery();
					}
					#endregion Detalle	

					string sSQLActDetalle = string.Format("Exec DTConsultaDetalleRetiro {0}", (int)this.txtIdIngreso.Value);
					oCmdActualiza.CommandText = sSQLActDetalle;
					SqlDataAdapter oDADetalle = new SqlDataAdapter(oCmdActualiza);
					DataTable oDTDetalle = new DataTable();
					oDADetalle.Fill(oDTDetalle);
					this.uGridDTIngreso.DataSource = oDTDetalle;							
					#endregion Graba Maestro
							
					oTransaction.Commit();	
								
					#region Controles
					this.cmbBodega.Enabled = false;
					this.cmbBodegaActual.Enabled = false;
					this.dtFecha.Enabled = false;
					this.txtNumeroFactura.Enabled = false;
					this.dtFechaRetiro.Enabled = false;	
					this.chkReservaMercaderia.Enabled = false;
					this.txtTiempoPago.Enabled = false;							
					this.txtObservaciones.Enabled = false;

					if (miAcceso.Nuevo) this.btnNuevo.Enabled = true;
					if (miAcceso.Buscar) this.btnConsultar.Enabled = true;
					if (miAcceso.Editar) this.btnEditar.Enabled = true;
					if (miAcceso.Imprimir) this.btnImprimir.Enabled = true;								
					if (miAcceso.Anular) this.btnAnular.Enabled = true;																
					if (miAcceso.BProcesar) if ((int)this.txtEstado.Value == 0) this.btnProcesar.Enabled = true;
							
					this.chkFiscal.Enabled = false;
					this.cmbMotivos.Enabled = false;
					this.btnGuardar.Enabled = false;
					this.btnCancelar.Enabled = true;
					this.cmbVendedor.Enabled = false;
					this.cmbRetiradoPor.Enabled = false;
					this.txtTiempoUso.Enabled = false;
							
					this.lblEstado.Text = "PENDIENTE";

					FuncionesProcedimientos.EstadoGrids(false, this.uGridDTIngreso);

					this.txtNumeroNotaDeCredito.Enabled = false;

					bNuevo = false;
					bEdicion = false;
					#endregion Controles
				}
				catch (Exception ex)
				{
					oTransaction.Rollback();

					MessageBox.Show(string.Format("Commit Exception Type : {0} {1}", ex.GetType(), ex.Message), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Error);					
				}
				finally
				{
					oConexion.Close();
				}
			}						
		}
		
		private void txtObservaciones_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Enter) this.btnGuardar.Focus();
		}

		private void uGridDTIngreso_ClickCellButton(object sender, Infragistics.Win.UltraWinGrid.CellEventArgs e)
		{
			if (bEdicion) 
			{
				string sCol = "";
				string sFiltro = "";
				if (e.Cell.Column.ToString() == "UnboundColumn1")
				{
					if ((int)e.Cell.Row.Cells["idArticuloGE"].Value > 0) 
					{
						MessageBox.Show("Garantia extendida no tiene accesorios", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
						return;
					}					

					for (int j = 17; j <= 36; j++)
					{
						sCol = this.uGridDTIngreso.DisplayLayout.Bands[0].Columns[j].Header.Caption;

						if (sFiltro.Length == 0) sFiltro = e.Cell.Row.Cells[sCol].Value.ToString();
						if (sFiltro.Length > 0) sFiltro = sFiltro + "," + e.Cell.Row.Cells[sCol].Value.ToString();
					}
				}

				using (frmAccesorios Accesorios = new frmAccesorios (sFiltro))
				{
					if (DialogResult.OK == Accesorios.ShowDialog())
					{
						int iPos = 1;						
						foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in Accesorios.ultraGrid1.Rows.All)
						{
							string sColAsigna = "idAccesorio";
							if ((bool)dr.Cells["Sel"].Value && (int)dr.Cells["idAccesorio"].Value > 0)
							{
								sColAsigna = sColAsigna + iPos.ToString();
								e.Cell.Row.Cells[sColAsigna].Value = (int)dr.Cells["idAccesorio"].Value;
								iPos++;				
							}
						}

						int i = 16;
						i = i + iPos;
						for (int k = i ; k <= 36; k++)
						{
							sCol = this.uGridDTIngreso.DisplayLayout.Bands[0].Columns[k].Header.Caption;
							e.Cell.Row.Cells[sCol].Value = 0;
						}
					}
				}
			}
		}

		private void Total()
		{
			#region Creacion de Variables
			decimal dDescuentoLineal = 0.00m;
			decimal dIvaLineal = 0.00m;
			#endregion Creacion de Variables
			
			this.txtIva.Value = 0.00m;			
			this.txtDescIvaTotal.Value = 0.00m;			
			this.txtIvaTotal.Value = 0.00m;
			this.txtValorNC.Value = 0.00m;
			this.txtDiferencia.Value = 0.00m;

			#region Desplazamiento de grilla
			foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.uGridDTIngreso.Rows.All)
			{
				dDescuentoLineal = 0.00m;
				dIvaLineal = 0.00m;

				if (dr.Cells["Cantidad"].Value != DBNull.Value && dr.Cells["Precio"].Value != DBNull.Value && dr.Cells["DescuentoPorc"].Value != DBNull.Value)
				{
					dr.Cells["Subtotal"].Value = Math.Round(Convert.ToDecimal(dr.Cells["Cantidad"].Value) * Convert.ToDecimal(dr.Cells["Precio"].Value), 2);
					dDescuentoLineal = Math.Round((Convert.ToDecimal(dr.Cells["Subtotal"].Value) * Convert.ToDecimal(dr.Cells["DescuentoPorc"].Value))/100, 2); //Math.Round((Convert.ToDecimal( dr.Cells["Cantidad"].Value * Convert.ToDecimal( dr.Cells["Precio"].Value) * Convert.ToDecimal( dr.Cells["DescuentoPorc"].Value/100, 2);

					dr.Cells["DescuentoArt"].Value = dDescuentoLineal;

					this.txtIva.Value = Math.Round(Convert.ToDecimal(this.txtIva.Value) + Convert.ToDecimal(dr.Cells["Subtotal"].Value), 2);
					this.txtDescIvaTotal.Value = Math.Round(Convert.ToDecimal(this.txtDescIvaTotal.Value) + Convert.ToDecimal(dr.Cells["DescuentoArt"].Value), 2);															
				}
								
				if (this.chkFiscal.Checked) 
				{
					decimal dPorcIva = Facturacion.dPorcentajeIVAVigente(this.chkFiscal.Checked, true, (int)this.txtIdCompra.Value, (DateTime)this.dtFecha.Value, cdsSeteoF);
					this.txtIvaTotal.Value = (Math.Round(Convert.ToDecimal(this.txtIva.Value), 2) - Math.Round(Convert.ToDecimal(this.txtDescIvaTotal.Value), 2)) * dPorcIva /100;
				}

				this.txtValorNC.Value = Convert.ToDecimal(this.txtIva.Value) 					
					- Convert.ToDecimal(this.txtDescIvaTotal.Value)
					+ Convert.ToDecimal(this.txtIvaTotal.Value);
			} 
			#endregion Desplazamiento de grilla

			//CalculaDiferencia();
		}

		private void uGridDTIngreso_InitializeLayout(object sender, Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs e)
		{
			FuncionesProcedimientos.FormatoGrid(e, "Precio", 2);
			FuncionesProcedimientos.FormatoGrid(e, "DescuentoPorc", 2);
			FuncionesProcedimientos.FormatoGrid(e, "Subtotal", 2);			
		}

		private void cmbBodega_Validating(object sender, System.ComponentModel.CancelEventArgs e)
		{
			
		}

		private void dtFecha_Validating(object sender, System.ComponentModel.CancelEventArgs e)
		{
			if (bEdicion)
			{
				if (this.dtFecha.Value == System.DBNull.Value)
				{
					MessageBox.Show("Seleccione la Fecha", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
					this.dtFecha.Focus();
					e.Cancel = true; 
				}
			}
		}

		private void dtFechaRetiro_BeforeDropDown(object sender, System.ComponentModel.CancelEventArgs e)
		{
		
		}

		private void uGridDTIngreso_AfterRowInsert(object sender, Infragistics.Win.UltraWinGrid.RowEventArgs e)
		{
			if (bEdicion)
			{
				e.Row.Cells["idDetalleIngresoRetiro"].Value = 0; 
				e.Row.Cells["idDetCompra"].Value = 0; 
				e.Row.Cells["idDTEstado"].Value = 0; 
				e.Row.Cells["Cantidad"].Value = 1;
				e.Row.Cells["DescuentoPorc"].Value = 0.00m;
				e.Row.Cells["DescuentoArt"].Value = 0.00m;				
				e.Row.Cells["Precio"].Value = 0.00m;
				e.Row.Cells["Subtotal"].Value = 0.00m;
				e.Row.Cells["Sel"].Value = false; 
				e.Row.Cells["idArticulo"].Value = 0; 
				e.Row.Cells["idSerial"].Value = 0; 								
				e.Row.Cells["idAccesorio1"].Value = 0; 
				e.Row.Cells["idAccesorio2"].Value = 0; 
				e.Row.Cells["idAccesorio3"].Value = 0; 
				e.Row.Cells["idAccesorio4"].Value = 0;
				e.Row.Cells["idAccesorio5"].Value = 0; 
				e.Row.Cells["idAccesorio6"].Value = 0; 
				e.Row.Cells["idAccesorio7"].Value = 0; 
				e.Row.Cells["idAccesorio8"].Value = 0;
				e.Row.Cells["idAccesorio9"].Value = 0; 
				e.Row.Cells["idAccesorio10"].Value = 0; 
				e.Row.Cells["idAccesorio11"].Value = 0; 
				e.Row.Cells["idAccesorio12"].Value = 0;
				e.Row.Cells["idAccesorio13"].Value = 0; 
				e.Row.Cells["idAccesorio14"].Value = 0; 
				e.Row.Cells["idAccesorio15"].Value = 0; 
				e.Row.Cells["idAccesorio16"].Value = 0;
				e.Row.Cells["idAccesorio17"].Value = 0; 
				e.Row.Cells["idAccesorio18"].Value = 0; 
				e.Row.Cells["idAccesorio19"].Value = 0; 
				e.Row.Cells["idAccesorio20"].Value = 0;
				e.Row.Cells["idArticuloGE"].Value = 0;
			}
		}

		private void uGridDTIngreso_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
			if (bEdicion)
				if (this.uGridDTIngreso.ActiveCell.Column.Key.ToString().ToUpper() == "CANTIDAD")
					Funcion.CamposNumericos(sender, e);
		}

		bool bActualiza = true;

		private void uGridDTIngreso_AfterCellUpdate(object sender, Infragistics.Win.UltraWinGrid.CellEventArgs e)
		{
			if (bEdicion)
			{
				if (!bActualiza) return;

				#region Codigo o Articulo
				if (e.Cell.Column.ToString() == "Codigo")
				{
					bActualiza = false;
					FuncionesInventario.FacturacionCargaArticulos(sender, e, this.uGridDTIngreso, 9, 4, 
						(int)this.cmbBodega.Value, 0, 0, 0, 0, 0, Convert.ToDateTime(this.dtFecha.Value), 54, this.chkFiscal.Checked, false, false, false, 0, idEntidadFinanciera, 0, 0, 0, false, false, 0, cdsSeteoF);
					Total();
					bActualiza = true;					
				} 
				#endregion Codigo o articulo

				if (e.Cell.Column.ToString() == "Serial")
				{
					string sSQLSerial = string.Format("Declare @idSerial Int = 0 If Exists(Select idSerial From Seriales Where Serial = '{0}') Select @idSerial = idSerial From Seriales Where Serial = '{0}' Else Set @idSerial = 0 Select @idSerial", 
						e.Cell.Row.Cells["Serial"].Value.ToString());
					int idSerial = Funcion.iEscalarSQL(cdsSeteoF, sSQLSerial);
					e.Cell.Row.Cells["idSerial"].Value = idSerial;
				}

				if (e.Cell.Column.ToString() == "Precio")
				{
					if (!FuncionesInventario.ValidaCamposGridDecimal(sender, e, "Precio"))
					{							
						Total();
						bActualiza = false;
						return;
					}

					if (!Validacion.vbValidaValoresNotasDeCredito(sender, e, 1, cdsSeteoF))
					{
						Total();
						bActualiza = false;
//						return;
					}

					Total();
				}

				if (e.Cell.Column.ToString() == "DescuentoArt")
				{
					if (!FuncionesInventario.ValidaCamposGridDecimal(sender, e, "DescuentoArt"))
					{							
						Total();
						bActualiza = false;
						return;
					}

					Total();
				}

				if (e.Cell.Column.ToString() == "DescuentoPorc")
				{
					if (!FuncionesInventario.ValidaCamposGridDecimal(sender, e, "DescuentoPorc"))
					{							
						Total();
						bActualiza = false;
						return;
					}

					if (!Validacion.vbValidaValoresNotasDeCredito(sender, e, 2, cdsSeteoF))
					{
						Total();
						bActualiza = false;
//						return;
					}

					Total();
				}

				bActualiza = true;
			}
		}

		private void uGridDTIngreso_BeforeCellUpdate(object sender, Infragistics.Win.UltraWinGrid.BeforeCellUpdateEventArgs e)
		{
			if (bEdicion)
			{
				if (e.Cell.Column.ToString() == "Sel")
					if (e.Cell.Row.Cells["Sel"].Value != System.DBNull.Value)
						if ((bool)e.Cell.Row.Cells["Sel"].Value)
							if (e.Cell.Row.Cells["idDetCompra"].Value != System.DBNull.Value)
								if ((int)e.Cell.Row.Cells["idDetCompra"].Value == 0)
									e.Cancel = true;
				
				if (e.Cell.Column.ToString() == "Serial")
					if ((int)e.Cell.Row.Cells["idArticuloGE"].Value > 0) e.Cancel = true;

				if (e.Cell.Column.ToString() == "Codigo")
					if ((int)e.Cell.Row.Cells["idArticuloGE"].Value > 0) e.Cancel = true;

				if (e.Cell.Column.ToString() == "idDTEstado")
					if ((int)e.Cell.Row.Cells["idArticuloGE"].Value > 0) e.Cancel = true;

				if (e.Cell.Column.ToString() == "Sel")
					if ((int)e.Cell.Row.Cells["idArticuloGE"].Value > 0) e.Cancel = true;

				if (e.Cell.Column.ToString() == "Precio")
					if ((int)e.Cell.Row.Cells["idArticuloGE"].Value > 0) e.Cancel = true;

				if (e.Cell.Column.ToString() == "DescuentoPorc")
				{
					if ((int)e.Cell.Row.Cells["idArticuloGE"].Value > 0) e.Cancel = true;
					if (this.chkFiscal.Checked) e.Cancel = true;
				}

			}
		}

		private void txtCodigo_ValueChanged(object sender, System.EventArgs e)
		{
		
		}

		private void txtDescuento_ValueChanged(object sender, System.EventArgs e)
		{
		
		}

		private void btnImprimir_Click(object sender, System.EventArgs e)
		{
			this.mnuImprimir.Show(this.btnImprimir, new Point(1, 25));
		}

		private void dtFecha_BeforeDropDown(object sender, System.ComponentModel.CancelEventArgs e)
		{
		
		}

		private void dtFecha_ValueChanged(object sender, System.EventArgs e)
		{
			
		}

		private void btnAnular_Click(object sender, System.EventArgs e)
		{
			this.Consulta((int)this.txtIdIngreso.Value);

			if ((int)this.txtEstado.Value == 1)
			{
				MessageBox.Show("El documento ya fue procesado", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				return;
			}

			if ((int)this.txtEstado.Value == 2)
			{
				MessageBox.Show("El documento esta anulado", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				return;
			}

			Funcion.EjecutaSQL(cdsSeteoF, string.Format("Update DTIngresoRetiros Set Estado = 2 Where idIngresoRetiro = {0}", (int)this.txtIdIngreso.Value));

			this.Consulta((int)this.txtIdIngreso.Value);
		}

		private void chkFiscal_CheckedChanged(object sender, System.EventArgs e)
		{
			if (bEdicion) 
			{
				this.Total();
				
				this.txtNumeroNotaDeCredito.Enabled = this.chkFiscal.Checked;				
				this.cmbBodega.Enabled = !this.chkFiscal.Checked;
				this.cmbVendedor.Enabled = !this.chkFiscal.Checked;			
			}
		}

		private void ProcesoEnGarantia()
		{
			using (SqlConnection oConexion = new SqlConnection("Password=Bl45o6$9;Persist Security Info=True;User ID= " + MenuLatinium.stUsuario + ";Initial Catalog=" + MenuLatinium.stDirFacturacion + ";Data Source=" + MenuLatinium.stDirServidor + ";Max Pool Size = 1024;"))
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
								
					#region Nota De Credito
					string sSQL = string.Format("Exec CreaNotaDeCreditoPorGarantiaDePago {0}, '{1}', {2}, {3}, '{4}', {5}, {6}",
						(int)this.txtIdCompra.Value, Convert.ToDateTime(this.dtFecha.Value).ToString("yyyyMMdd"), (int)this.cmbMotivos.Value, (int)this.txtIdIngreso.Value, this.txtObservaciones.Text.ToString(), 
						(int)this.txtTiempoPago.Value, Convert.ToDecimal(this.txtValorNC.Value));
					oCmdActualiza.CommandText = sSQL;
					this.txtIdTransferenciaDT.Value = (int)oCmdActualiza.ExecuteScalar();					
					#endregion Nota De Credito
            
					string sSQLActDetalle = string.Format("Exec DTConsultaDetalleRetiro {0}", (int)this.txtIdIngreso.Value);						
					oCmdActualiza.CommandText = sSQLActDetalle;
					SqlDataAdapter oDADetalle1 = new SqlDataAdapter(oCmdActualiza);
					DataTable oDTDetalle1 = new DataTable();
					oDADetalle1.Fill(oDTDetalle1);
					this.uGridDTIngreso.DataSource = oDTDetalle1;

					#region Seriales
					foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.uGridDTIngreso.Rows.All)
					{	
						if ((int)dr.Cells["idSerial"].Value > 0)
						{
							string sSQLActDispSerial = string.Format("Update DetalleSeriales Set Disponible = 0 Where idSerial = {0}", (int)dr.Cells["idSerial"].Value);								
							oCmdActualiza.CommandText = sSQLActDispSerial;
							oCmdActualiza.ExecuteNonQuery();

							string sSQLDS = string.Format("Exec GuardaDetalleSeriales {0}, {1}, {2}, {3}, '', 1, 'REGISTRADO DESDE RETIRO DE MERCADERIA', 1, 0, 0, 0",
								(int)dr.Cells["idSerial"].Value, (int)dr.Cells["idDetalleNotaDeCredito"].Value, 0, (int)dr.Cells["Bodega"].Value);								
							oCmdActualiza.CommandText = sSQLDS;
							oCmdActualiza.ExecuteNonQuery();
								
							string sSQLActSerial = string.Format("Update Seriales Set Estado = 17, idArticulo = {0}, Bodega = {1} Where idSerial = {2}",
								(int)dr.Cells["idArticulo"].Value, (int)dr.Cells["Bodega"].Value, (int)dr.Cells["idSerial"].Value);								
							oCmdActualiza.CommandText = sSQLActSerial;
							oCmdActualiza.ExecuteNonQuery();
						}
						else if ((int)dr.Cells["idSerial"].Value == 0)
						{
							string sSQLSeriales = string.Format("Exec GuardaSeriales {0}, {1}, '{2}', {3}, {4}",
								(int)dr.Cells["idSerial"].Value, (int)dr.Cells["idArticulo"].Value, dr.Cells["Serial"].Value.ToString(), 17, (int)dr.Cells["Bodega"].Value);								
							oCmdActualiza.CommandText = sSQLSeriales;
							int idSerial = (int)oCmdActualiza.ExecuteScalar();								

							string sSQLDS = string.Format("Exec GuardaDetalleSeriales {0}, {1}, {2}, {3}, '', 1, 'REGISTRADO DESDE RETIRO DE MERCADERIA', 1, 0, 0, 0",
								idSerial, (int)dr.Cells["idDetalleNotaDeCredito"].Value, 0, (int)dr.Cells["Bodega"].Value);								
							oCmdActualiza.CommandText = sSQLDS;
							oCmdActualiza.ExecuteNonQuery();

							string sSQLActSer = string.Format("Update DTDetalleIngresoRetiros Set idSerial = {0} Where idDetalleIngresoRetiro = {1}", 
								idSerial, (int)dr.Cells["idDetalleIngresoRetiro"].Value);								
							oCmdActualiza.CommandText = sSQLActSer;
							oCmdActualiza.ExecuteNonQuery();
						}
					}
					#endregion Seriales
																							
					string sSQLEstado = string.Format("Update DTIngresoRetiros Set Estado = 1 Where idIngresoRetiro = {0}", (int)this.txtIdIngreso.Value);						
					oCmdActualiza.CommandText = sSQLEstado;
					oCmdActualiza.ExecuteNonQuery();	

					oTransaction.Commit();	

					this.Consulta((int)this.txtIdIngreso.Value);
				}
				catch (Exception ex)
				{
					oTransaction.Rollback();
				
					MessageBox.Show(string.Format("Commit Exception Type : {0} {1}", ex.GetType(), ex.Message), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Error);					
				}
				finally
				{
					oConexion.Close();
				}
			}	
		}

		#region Proceso en dacion de pago
		private void ProcesoDacionPago()
		{
			using (SqlConnection oConexion = new SqlConnection("Password=Bl45o6$9;Persist Security Info=True;User ID= " + MenuLatinium.stUsuario + ";Initial Catalog=" + MenuLatinium.stDirFacturacion + ";Data Source=" + MenuLatinium.stDirServidor + ";Max Pool Size = 1024;"))
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
			
					DateTime dtFecha = (DateTime)this.dtFecha.Value;

					#region Articulo Nuevo
					foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.uGridDTIngreso.Rows.All)
					{
						if ((int)dr.Cells["idDTEstado"].Value > 1)
						{
							string sCodigo = "RECU-" + dr.Cells["Codigo"].Value.ToString();

							string sSQLGuardaArticulo = string.Format("Exec DTCreaArticuloRecuperacion {0}, {1}, '{2}'", (int)dr.Cells["idDetalleIngresoRetiro"].Value, (int)dr.Cells["idArticulo"].Value, sCodigo);							
							oCmdActualiza.CommandText = sSQLGuardaArticulo;
							oCmdActualiza.ExecuteNonQuery();																				
						}
						else
						{
							string sSQLActArt = string.Format("Update DTDetalleIngresoRetiros Set idArticuloMod = {0} Where idDetalleIngresoRetiro = {1}", 
								(int)dr.Cells["idArticulo"].Value, (int)dr.Cells["idDetalleIngresoRetiro"].Value);									
							oCmdActualiza.CommandText = sSQLActArt;
							oCmdActualiza.ExecuteNonQuery();
						}
					}					
					#endregion Articulo Nuevo

					string sSQLActDetalle = string.Format("Exec DTConsultaDetalleRetiro {0}", (int)this.txtIdIngreso.Value);						
					oCmdActualiza.CommandText = sSQLActDetalle;
					SqlDataAdapter oDADetalle = new SqlDataAdapter(oCmdActualiza);
					DataTable oDTDetalle = new DataTable();
					oDADetalle.Fill(oDTDetalle);
					this.uGridDTIngreso.DataSource = oDTDetalle;		
						
					#region Nota De Credito
					oCmdActualiza.CommandText = "Exec NumeracionLocales 5, 1, 1";
					this.txtNumeroNotaDeCredito.Text = oCmdActualiza.ExecuteScalar().ToString(); 
					
					int iBodegaDestino = 18;					

					string sSQLNC = string.Format("Exec CreaNotaDeCreditoPorRetiroMercaderia {0}, {1}, '{2}', {3}, {4}, {5}, {6}, '{7}', {8}, '{9}', {10}",
						(int)this.txtIdNotaDeCredito.Value, (int)this.txtIdCompra.Value, dtFecha.ToString("yyyyMMdd"), (int)this.cmbBodega.Value, this.chkFiscal.Checked, 
						(int)this.cmbMotivos.Value, (int)this.txtIdIngreso.Value, this.txtObservaciones.Text.ToString(), (int)this.cmbVendedor.Value, this.txtNumeroNotaDeCredito.Text.ToString(), iBodegaDestino);
					oCmdActualiza.CommandText = sSQLNC;
					this.txtIdNotaDeCredito.Value = (int)oCmdActualiza.ExecuteScalar();
					
					if (Convert.ToDecimal(this.txtDiferencia.Value) > 0.00m)
					{
						string sSQLNDDiferencia = string.Format("Exec DTCreaNotaDeCreditoDiferenciasRetiro {0}, {1}", 
							(int)this.txtIdIngreso.Value, Convert.ToDecimal(this.txtDiferencia.Value));							
						oCmdActualiza.CommandText = sSQLNDDiferencia;
						this.txtIdNotaDeCredito1.Value = (int)oCmdActualiza.ExecuteScalar();					
					}
					#endregion Nota De Credito
            
					sSQLActDetalle = string.Format("Exec DTConsultaDetalleRetiro {0}", (int)this.txtIdIngreso.Value);						
					oCmdActualiza.CommandText = sSQLActDetalle;
					SqlDataAdapter oDADetalle1 = new SqlDataAdapter(oCmdActualiza);
					DataTable oDTDetalle1 = new DataTable();
					oDADetalle1.Fill(oDTDetalle1);
					this.uGridDTIngreso.DataSource = oDTDetalle1;

					#region Seriales
					foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.uGridDTIngreso.Rows.All)
					{	
						if ((int)dr.Cells["idSerial"].Value > 0)
						{
							string sSQLActDispSerial = string.Format("Update DetalleSeriales Set Disponible = 0 Where idSerial = {0}", (int)dr.Cells["idSerial"].Value);								
							oCmdActualiza.CommandText = sSQLActDispSerial;
							oCmdActualiza.ExecuteNonQuery();

							string sSQLDS = string.Format("Exec GuardaDetalleSeriales {0}, {1}, {2}, {3}, '', 1, 'REGISTRADO DESDE RETIRO DE MERCADERIA', 1, 0, 0, 0",
								(int)dr.Cells["idSerial"].Value, (int)dr.Cells["idDetalleNotaDeCredito"].Value, 0, iBodegaDestino);								
							oCmdActualiza.CommandText = sSQLDS;
							oCmdActualiza.ExecuteNonQuery();
								
							string sSQLActSerial = string.Format("Update Seriales Set Estado = 17, idArticulo = {0}, Bodega = {1} Where idSerial = {2}",
								(int)dr.Cells["idArticuloMod"].Value, iBodegaDestino, (int)dr.Cells["idSerial"].Value);								
							oCmdActualiza.CommandText = sSQLActSerial;
							oCmdActualiza.ExecuteNonQuery();
						}
						else if ((int)dr.Cells["idSerial"].Value == 0)
						{
							string sSQLSeriales = string.Format("Exec GuardaSeriales {0}, {1}, '{2}', {3}, {4}",
								(int)dr.Cells["idSerial"].Value, (int)dr.Cells["idArticuloMod"].Value, dr.Cells["Serial"].Value.ToString(), 17, iBodegaDestino);								
							oCmdActualiza.CommandText = sSQLSeriales;
							int idSerial = (int)oCmdActualiza.ExecuteScalar();								

							string sSQLDS = string.Format("Exec GuardaDetalleSeriales {0}, {1}, {2}, {3}, '', 1, 'REGISTRADO DESDE RETIRO DE MERCADERIA', 1, 0, 0, 0",
								idSerial, (int)dr.Cells["idDetalleNotaDeCredito"].Value, 0, iBodegaDestino);								
							oCmdActualiza.CommandText = sSQLDS;
							oCmdActualiza.ExecuteNonQuery();

							string sSQLActSer = string.Format("Update DTDetalleIngresoRetiros Set idSerial = {0} Where idDetalleIngresoRetiro = {1}", 
								idSerial, (int)dr.Cells["idDetalleIngresoRetiro"].Value);								
							oCmdActualiza.CommandText = sSQLActSer;
							oCmdActualiza.ExecuteNonQuery();
						}
					}
					#endregion Seriales
						
					sSQLActDetalle = string.Format("Exec DTConsultaDetalleRetiro {0}", (int)this.txtIdIngreso.Value);						
					oCmdActualiza.CommandText = sSQLActDetalle;
					SqlDataAdapter oDADetalle2 = new SqlDataAdapter(oCmdActualiza);
					DataTable oDTDetalle2 = new DataTable();
					oDADetalle2.Fill(oDTDetalle2);
					this.uGridDTIngreso.DataSource = oDTDetalle2;
			
					#region Transferencia DT
					string sSQLTransf = string.Format("Exec CreaTransferenciaRetiro {0}, {1}, '{2}', {3}, {4}",
						(int)this.txtidTransferencia.Value, iBodegaDestino, dtFecha.ToString("yyyyMMdd"), (int)this.txtIdIngreso.Value, (int)this.txtIdCompra.Value);
					oCmdActualiza.CommandText = sSQLTransf;
					oCmdActualiza.ExecuteNonQuery();		
					#endregion Transferencia DT

					#region id Transferencias
					string sSQLIdTransfBI = string.Format("Select ISNULL(idTransferencia, 0) From DTIngresoRetiros Where idIngresoRetiro = {0}", (int)this.txtIdIngreso.Value);							
					oCmdActualiza.CommandText = sSQLIdTransfBI;
					this.txtidTransferencia.Value = (int)oCmdActualiza.ExecuteScalar();

					string sSQLIdTransfDT = string.Format("Select ISNULL(idTransferenciaDT, 0) From DTIngresoRetiros Where idIngresoRetiro = {0}", (int)this.txtIdIngreso.Value);							
					oCmdActualiza.CommandText = sSQLIdTransfDT;
					this.txtIdTransferenciaDT.Value = (int)oCmdActualiza.ExecuteScalar();												
					#endregion id Transferencias

					string sSQLEstado = string.Format("Update DTIngresoRetiros Set Estado = 1 Where idIngresoRetiro = {0}", (int)this.txtIdIngreso.Value);						
					oCmdActualiza.CommandText = sSQLEstado;
					oCmdActualiza.ExecuteNonQuery();	

					oTransaction.Commit();	

					this.Consulta((int)this.txtIdIngreso.Value);
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
		#endregion Proceso en dacion de pago

		private void btnProcesar_Click(object sender, System.EventArgs e)
		{
			this.Consulta((int)this.txtIdIngreso.Value);

			if ((int)this.txtEstado.Value == 1)
			{
				MessageBox.Show("El documento ya fue procesado", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				return;
			}
			if ((int)this.txtEstado.Value == 2)
			{
				MessageBox.Show("El documento esta anulado", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				return;
			}

			if (DialogResult.Yes == MessageBox.Show("Antes de Procesar. \n\n VERIQUE QUE LOS DATOS REGISTRADOS SON LOS CORRECTOS \n\n INFORMACION DEL RETIRO \n\n INFORMACIÓN DEL CLIENTE \n PRODUCTOS \n PRECIOS CANTIDADES DESCUENTOS \n\n DESTINO DE LA MERCADERIA \n\n\n RECUERDE QUE ESTE PROCESO NO PUEDE SER REVERTIDO \n\n\n\n ¿Esta seguro de Procesar este retiro de mercaderia?", "Point Technology", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2)) 
			{				
				if ((int)this.cmbMotivos.Value != 40) ProcesoDacionPago();
				else ProcesoEnGarantia();
			}
		}

		private void uGridDTIngreso_BeforeRowsDeleted(object sender, Infragistics.Win.UltraWinGrid.BeforeRowsDeletedEventArgs e)
		{
			if (bEdicion)
			{
				if ((int)this.cmbMotivos.Value != 40)
				{
					for (int i=0; i<e.Rows.Length; i++)
					{
						if ((int)e.Rows[i].Cells["idArticuloGE"].Value > 0) 
						{
							e.Cancel = true;
							return;
						}
					}
				}

				if (DialogResult.Yes ==	MessageBox.Show("¿Esta Seguro de Borrar los registros seleccionados?", "Point Technology", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2))
				{
					if (!bNuevo)
					{						
						for (int i=0; i<e.Rows.Length; i++)
						{							
							e.DisplayPromptMsg = false;

							string sSQLElimina = string.Format("Delete From DTDetalleIngresoRetiros Where idDetalleIngresoRetiro = {0}", 
								(int)e.Rows[i].Cells["idDetalleIngresoRetiro"].Value);
							Funcion.EjecutaSQL(cdsSeteoF, sSQLElimina, true);						
						}
					}
					else e.DisplayPromptMsg = false;
				}
			}
			else e.Cancel = true;
		}

		private void uGridDTIngreso_AfterRowsDeleted(object sender, System.EventArgs e)
		{
			if (bEdicion) this.Total();
		}

		private void cmbRetiradoPor_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Enter) this.cmbMotivos.Focus();
		}

		private void txtTiempoUso_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			
		}

		private void cmbMotivos_InitializeLayout(object sender, Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs e)
		{
		
		}

		private void txtTiempoUso_ValueChanged(object sender, System.EventArgs e)
		{
		
		}

		private void cmbMotivos_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Enter) 
			{
				if (!this.chkFiscal.Checked) this.cmbBodega.Focus();
				else if (miAcceso.CambiarFecha) this.dtFecha.Focus();
				else this.txtNumeroNotaDeCredito.Focus();
			}
		}

		private void txtTiempoUso_Enter(object sender, System.EventArgs e)
		{
			this.txtTiempoUso.SelectAll();
		}

		private void txtTiempoUso_Click(object sender, System.EventArgs e)
		{
			this.txtTiempoUso.SelectAll();
		}

		private void menuItem2_Click(object sender, System.EventArgs e)
		{
			#region Imprime Transferencia	BI
			if ((int)this.txtidTransferencia.Value == 0)
			{
				MessageBox.Show("No existe transferencia a BI de este documento", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				return;
			}

			int idCompra = Funcion.iEscalarSQL(cdsSeteoF, string.Format("Select idCompra From CompraTransf Where idCompraTransf = {0}", (int)this.txtidTransferencia.Value));

			ParameterFields paramFields = new ParameterFields ();		
	    
			ParameterField paramFieldCompra = new ParameterField ();
			ParameterDiscreteValue ValIdCompra = new ParameterDiscreteValue ();
			paramFieldCompra.ParameterFieldName = "@idCompra";
			ValIdCompra.Value = idCompra;
			paramFieldCompra.CurrentValues.Add (ValIdCompra);
			paramFields.Add (paramFieldCompra);

			string sReporte = "";
			string sTitulo = "";
			
			sReporte = "RepTransfSalida.rpt";
			sTitulo = "Resumen de Envio de Transferencia";
			
			Reporte miRep = new Reporte(sReporte, "");
			miRep.MdiParent = this.MdiParent;
			miRep.Titulo(sTitulo);
			miRep.crVista.ParameterFieldInfo = paramFields;
			miRep.Show();
			#endregion Imprime Transferencia BI
		}

		private void menuItem3_Click(object sender, System.EventArgs e)
		{
			#region Imprime Transferencia	DT
			if ((int)this.txtIdTransferenciaDT.Value == 0)
			{
				MessageBox.Show("No existe transferencia al Departamento Tecnico de este documento", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				return;
			}
							
			int idCompraDT = Funcion.iEscalarSQL(cdsSeteoF, string.Format("Select idCompra From CompraTransf Where idCompraTransf = {0}", (int)this.txtIdTransferenciaDT.Value));

			ParameterFields paramFieldsDT = new ParameterFields ();
			ParameterField paramFieldCompraDT = new ParameterField ();
			ParameterDiscreteValue ValIdCompraDT = new ParameterDiscreteValue ();
			paramFieldCompraDT.ParameterFieldName = "@idCompra";
			ValIdCompraDT.Value = idCompraDT;
			paramFieldCompraDT.CurrentValues.Add (ValIdCompraDT);
			paramFieldsDT.Add (paramFieldCompraDT);
					
			string sReporte = "RepTransfSalida.rpt";
						
			Reporte Rep = new Reporte(sReporte, "");
			Rep.MdiParent = this.MdiParent;
			Rep.Titulo("Transferencia DT");
			Rep.crVista.ParameterFieldInfo = paramFieldsDT;
			Rep.Show();
			#endregion Imprime Transferencia DT
		}

		private void menuItem1_Click(object sender, System.EventArgs e)
		{
			this.Consulta((int)this.txtIdIngreso.Value);
			
			bool bAnulado = false;
			if ((int)this.txtEstado.Value == 2) bAnulado = true;

			bool bImpreso = false;

			if (this.chkFiscal.Checked)
				bImpreso = Funcion.bEscalarSQL(cdsSeteoF, string.Format("Select Impreso From Compra Where idCompra = {0}", (int)this.txtIdNotaDeCredito.Value), true);

			Facturacion.Impresion((int)this.txtIdNotaDeCredito.Value, bAnulado, 8, miAcceso.Imprimir, false, this.chkFiscal.Checked, bImpreso, 5, 0.00m, 0, 0, "", DateTime.Today, 0, "", cdsSeteoF);						
		}

		private void menuItem4_Click(object sender, System.EventArgs e)
		{
			#region Nota de Entrega
			ParameterFields paramFieldsDT = new ParameterFields ();
			
			ParameterField paramFieldCompraDT = new ParameterField ();
			ParameterDiscreteValue ValIdCompraDT = new ParameterDiscreteValue ();
			paramFieldCompraDT.ParameterFieldName = "@idIngresoRetiro";
			ValIdCompraDT.Value = (int)this.txtIdIngreso.Value;
			paramFieldCompraDT.CurrentValues.Add (ValIdCompraDT);
			paramFieldsDT.Add (paramFieldCompraDT);
					
			Reporte Rep = new Reporte("RMProductosEntregados.rpt", "");
			Rep.MdiParent = this.MdiParent;			
			Rep.crVista.ParameterFieldInfo = paramFieldsDT;
			Rep.Show();
			#endregion Imprime Transferencia DT
		}

		private void txtTiempoUso_Validated(object sender, System.EventArgs e)
		{
			if (this.txtTiempoUso.Value == System.DBNull.Value) this.txtTiempoUso.Value = 0;
		}

		private void menuItem5_Click(object sender, System.EventArgs e)
		{
			#region Imprime Transferencia	RC
			ParameterFields paramFieldsDT = new ParameterFields ();
				
			int idCompraRC = Funcion.iEscalarSQL(cdsSeteoF, string.Format("Select idCompra From CompraTransf Where idCompraTransf = {0}", (int)this.txtidTransferenciaRC.Value));

			ParameterField paramFieldCompraDT = new ParameterField ();
			ParameterDiscreteValue ValIdCompraDT = new ParameterDiscreteValue ();
			paramFieldCompraDT.ParameterFieldName = "@idCompra";
			ValIdCompraDT.Value = idCompraRC;
			paramFieldCompraDT.CurrentValues.Add (ValIdCompraDT);
			paramFieldsDT.Add (paramFieldCompraDT);
					
			string sReporte = "RepTransfSalida.rpt";
						
			Reporte Rep = new Reporte(sReporte, "");
			Rep.MdiParent = this.MdiParent;
			Rep.Titulo("Transferencia DT");
			Rep.crVista.ParameterFieldInfo = paramFieldsDT;
			Rep.Show();
			#endregion Imprime Transferencia RC
		}

		private void dtFecha_Validated(object sender, System.EventArgs e)
		{
			if (bEdicion)
			{
				this.dtFechaRetiro.CalendarInfo.MaxDate = (DateTime)this.dtFecha.Value;
			}
		}

		private void cmbBodega_InitializeLayout(object sender, Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs e)
		{
		
		}

		private void txtNumeroNotaDeCredito_Validated(object sender, System.EventArgs e)
		{
			if (this.txtNumeroNotaDeCredito.Text.ToString().Length > 0)
			{
				char Pad = '0';
				this.txtNumeroNotaDeCredito.Text = this.txtNumeroNotaDeCredito.Value.ToString().PadLeft(9, Pad);
			}
		}

		private void label11_Click(object sender, System.EventArgs e)
		{
		
		}

		private void label8_Click(object sender, System.EventArgs e)
		{
		
		}

		private void cmbBodega_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Enter) this.cmbVendedor.Focus();
		}

		private void cmbVendedor_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Enter) if (miAcceso.CambiarFecha) this.dtFecha.Focus(); else this.txtNumeroNotaDeCredito.Focus();
		}

		private void dtFecha_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Delete) e.Handled = true;
			if (e.KeyCode == Keys.Enter) this.txtNumeroNotaDeCredito.Focus();
		}

		private void dtFecha_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
			if (e.KeyChar == 8) e.Handled = true;
		}

		private void dtFechaRetiro_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
			if (e.KeyChar == 8) e.Handled = true;
		}

		private void txtNumeroNotaDeCredito_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Enter) this.cmbBodegaActual.Focus();
		}

		private void dtFechaRetiro_ValueChanged(object sender, System.EventArgs e)
		{
			if (bEdicion)
			{
				this.txtTiempoUso.Value = Funcion.iEscalarSQL(cdsSeteoF, string.Format("Select DATEDIFF(DAY, '{0}', '{1}')", Convert.ToDateTime(this.dtEmision.Value).ToString("yyyyMMdd"), Convert.ToDateTime(this.dtFechaRetiro.Value).ToString("yyyyMMdd")));
			}
		}

		private void frmDTIngresoPorRetiro_Closing(object sender, System.ComponentModel.CancelEventArgs e)
		{
			if (bEdicion)
			{
				MessageBox.Show("No puede salir mientras edita un Retiro", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				e.Cancel = true;
				return;
			}
		}

		private void cmbRetiradoPor_InitializeLayout(object sender, Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs e)
		{
		
		}

		private void menuItem6_Click(object sender, System.EventArgs e)
		{
			this.Consulta((int)this.txtIdIngreso.Value);
						
			if ((int)this.txtIdNotaDeCredito1.Value == 0) 
			{
				MessageBox.Show("No existe nota de credito por diferencia de saldos", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				return;
			}

			bool bAnulado = false;
			if ((int)this.txtEstado.Value == 2) bAnulado = true;

			Facturacion.Impresion((int)this.txtIdNotaDeCredito1.Value, bAnulado, 8, miAcceso.Imprimir, true, false, false, 5, 0.00m, 0, 0, "", DateTime.Today, 0, "", cdsSeteoF);
		}

		private void label24_Click(object sender, System.EventArgs e)
		{
		
		}

		private void chkVerificar_CheckedChanged(object sender, System.EventArgs e)
		{
			if (bEdicion)
				this.txtTiempoPago.Enabled = this.chkReservaMercaderia.Checked;
		}

		private void txtTiempoPago_Click(object sender, System.EventArgs e)
		{
			this.txtTiempoPago.SelectAll();
		}

		private void cmbBodegaActual_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Enter) this.txtObservaciones.Focus();
		}

		private void cmbRetiradoPor_ValueChanged(object sender, System.EventArgs e)
		{
			
		}

		private void cmbMotivos_ValueChanged(object sender, System.EventArgs e)
		{
//			if (bEdicion)
//			{
//				if (this.cmbMotivos.ActiveRow != null)
//				{
//					this.chkFiscal.Checked = true;
//					this.txtTiempoPago.Value = 0;
//					this.txtTiempoPago.Enabled = false;
//					this.chkReservaMercaderia.Checked = false;
//
//					if ((int)this.cmbMotivos.Value == 40)
//					{
//						this.chkFiscal.Checked = false;
//						this.txtTiempoPago.Enabled = true;
//						this.txtTiempoPago.Value = 30;
//						this.chkReservaMercaderia.Checked = true;
//					}
//				}
//			}
		}
	}
}

