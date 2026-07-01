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
using Latinium.Services;
using Latinium.Models;

namespace Latinium
{
	/// <summary>
	/// Summary description for frmPrefactura.
	/// </summary>
	public class frmPrefactura : DevExpress.XtraEditors.XtraForm
	{
		private System.Windows.Forms.Label label21;
		private System.Windows.Forms.Label label61;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtNumCuotas;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtCuotaDeEntrada;
		private System.Windows.Forms.Label lblVendedor;
		private System.Windows.Forms.Label label6;
		private System.Windows.Forms.Label label62;
		private System.Windows.Forms.ComboBox cboFechaPrimerPago;
		private System.Windows.Forms.Label lblIdentificacion;
		private System.Windows.Forms.Label lblCliente;
		private System.Windows.Forms.Label lblBodega;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtRuc;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbBodega;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtNombre;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtNotas;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtCuotaAsignada;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtCuotaDisponible;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtCuotaUsada;
		private System.Windows.Forms.Label label63;
		private System.Windows.Forms.Label label64;
		private System.Windows.Forms.Label label65;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtValorCuotas;
		private System.Windows.Forms.Label label23;
		private System.Windows.Forms.Label lblEstado;
		private System.Windows.Forms.Label lblUsuario;
		private System.Windows.Forms.Label lblIva;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtTotal;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtIvaTotal;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtDescIvaTotal;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtIva;
		private System.Windows.Forms.Label label1;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtNumero;
		private System.Windows.Forms.Label label2;
		private System.Windows.Forms.ContextMenu MnuAprobacion;
		private System.Windows.Forms.MenuItem mnuAprobacionAprobar;
		private System.Windows.Forms.MenuItem mnuAprobacionRechazar;
		private C1.Data.C1DataSet cdsSeteoF;
		private System.Windows.Forms.GroupBox gbBotones;
		private System.Windows.Forms.Label label3;
		private System.Windows.Forms.Button btnNuevo;
		private System.Windows.Forms.Button btnConsultar;
		private System.Windows.Forms.Button btnEditar;
		private System.Windows.Forms.Button btnImprimir;
		private System.Windows.Forms.Button btnGuardar;
		private System.Windows.Forms.Button btnAnular;
		private System.Windows.Forms.Button btnCancelar;
		private System.Windows.Forms.Button btnSalir;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtNumIdProyecto;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtIdComprobante;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtNumEstado;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtIdTipoFactura;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtIdSucursal;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtNumContadoCredito;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtNumIdCompra;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtNumIdCliente;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtIdFormaPago;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtidTipoRuc;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtComprobante;
		private System.Windows.Forms.Label lblComprobante;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtTotalCuotas;
		private System.Windows.Forms.Label label4;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource2;
		private System.Windows.Forms.Button btnAprobacion;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtInteres;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource1;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtidEntidadFinanciera;
		private System.Windows.Forms.Button btnClaveDescuento;
		/// <summary>
		/// Required designer variable.
		/// </summary>
		private System.ComponentModel.Container components = null;
		
		private Infragistics.Win.UltraWinEditors.UltraCheckEditor chkVerificar;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtObservaciones;
		private System.Windows.Forms.Label label5;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtBodegaPredef;
		private Infragistics.Win.UltraWinGrid.UltraGrid ultraGrid2;
		private System.Windows.Forms.Label lblIva0;
		private System.Windows.Forms.Label label55;
		private System.Windows.Forms.Label label8;
		private System.Windows.Forms.Label label9;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtDescIva0Total;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtIva0;
		private System.Windows.Forms.Label label15;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbBodega1;
		private System.Windows.Forms.Button btnValidarMail;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtProteccionDD;
		private Infragistics.Win.UltraWinSchedule.UltraCalendarCombo dtFecha;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbVendedor;
		private System.Windows.Forms.Label lblClienteEntidad;
		private Infragistics.Win.UltraWinEditors.UltraCheckEditor chkAnclar;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtSerieNuovo;
		private System.Windows.Forms.Label label7;
		private System.Windows.Forms.Label lblColateral;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtSerieNuovo2;
		private System.Windows.Forms.Label label10;
		private Infragistics.Win.UltraWinEditors.UltraCheckEditor chkAutorizar;
		private C1.Data.C1DataSet cdsSeteo;
		private Infragistics.Win.UltraWinEditors.UltraCheckEditor chkFreelance;
		private System.Windows.Forms.Label lblClienteReiterativo;
		private System.Windows.Forms.Button btnCrediPoint;
		private System.Windows.Forms.Label lblRecurrente;
		bool bFundas=false;
		
		public frmPrefactura()
		{
			//
			// Required for Windows Form Designer support
			//
			InitializeComponent();

			//
			// TODO: Add any constructor code after InitializeComponent call
			//
		}

		int getIdCompra = 0;
		bool getbInactivo = false;

		public frmPrefactura(int idCompra, bool bInactivo)
		{
			//
			// Required for Windows Form Designer support
			//
			InitializeComponent();
			getIdCompra = idCompra;
			getbInactivo = bInactivo;
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
			Infragistics.Win.Appearance appearance1 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance2 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance3 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand1 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn1 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idPersonal");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn2 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn3 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Codigo");
			Infragistics.Win.Appearance appearance4 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand2 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn4 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Bodega");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn5 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn6 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Codigo", 0);
			Infragistics.Win.Appearance appearance5 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance6 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand3 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn7 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idDetCompra");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn8 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idOrigen");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn9 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCompra");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn10 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idArticulo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn11 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idUnidad");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn12 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Codigo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn13 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Articulo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn14 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Cantidad");
			Infragistics.Win.Appearance appearance7 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn15 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Precio");
			Infragistics.Win.Appearance appearance8 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn16 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Impuesto");
			Infragistics.Win.Appearance appearance9 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn17 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("DescuentoArt");
			Infragistics.Win.Appearance appearance10 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn18 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("DescuentoPorc");
			Infragistics.Win.Appearance appearance11 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn19 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("SubTotal");
			Infragistics.Win.Appearance appearance12 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn20 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("EstadoSerial");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn21 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Tipo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn22 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Entregado");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn23 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Posicion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn24 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Principal");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn25 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Bloqueado");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn26 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCombo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn27 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idArticuloGE");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn28 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Jornadas");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn29 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Aporte");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn30 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idSubProyecto");
			Infragistics.Win.Appearance appearance13 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance14 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance15 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance16 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance17 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance18 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance19 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn1 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idDetCompra");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn2 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idOrigen");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn3 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idCompra");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn4 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idArticulo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn5 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idUnidad");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn6 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Codigo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn7 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Articulo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn8 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Cantidad");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn9 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Precio");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn10 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Impuesto");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn11 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("DescuentoArt");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn12 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("DescuentoPorc");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn13 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("SubTotal");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn14 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("EstadoSerial");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn15 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Tipo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn16 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Entregado");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn17 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Posicion");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn18 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Principal");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn19 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Bloqueado");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn20 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idCombo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn21 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idArticuloGE");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn22 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Jornadas");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn23 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Aporte");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn24 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idSubProyecto");
			Infragistics.Win.Appearance appearance20 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance21 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance22 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance23 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance24 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance25 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance26 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance27 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance28 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance29 = new Infragistics.Win.Appearance();
			System.Resources.ResourceManager resources = new System.Resources.ResourceManager(typeof(frmPrefactura));
			Infragistics.Win.Appearance appearance30 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn25 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idEntidadFinanciera");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn26 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Nombre");
			Infragistics.Win.Appearance appearance31 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance32 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance33 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance34 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance35 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance36 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand4 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn31 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Bodega");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn32 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn33 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Codigo", 0);
			Infragistics.Win.Appearance appearance37 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance38 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton dateButton1 = new Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton();
			Infragistics.Win.Appearance appearance39 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance40 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance41 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance42 = new Infragistics.Win.Appearance();
			this.label21 = new System.Windows.Forms.Label();
			this.label61 = new System.Windows.Forms.Label();
			this.txtNumCuotas = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtCuotaDeEntrada = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.lblVendedor = new System.Windows.Forms.Label();
			this.label6 = new System.Windows.Forms.Label();
			this.label62 = new System.Windows.Forms.Label();
			this.cmbVendedor = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.cboFechaPrimerPago = new System.Windows.Forms.ComboBox();
			this.lblIdentificacion = new System.Windows.Forms.Label();
			this.lblCliente = new System.Windows.Forms.Label();
			this.lblBodega = new System.Windows.Forms.Label();
			this.txtRuc = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.cmbBodega = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.txtNombre = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.ultraGrid2 = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.ultraDataSource2 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.txtNotas = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.txtCuotaAsignada = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtCuotaDisponible = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtCuotaUsada = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.label63 = new System.Windows.Forms.Label();
			this.label64 = new System.Windows.Forms.Label();
			this.label65 = new System.Windows.Forms.Label();
			this.txtValorCuotas = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.label23 = new System.Windows.Forms.Label();
			this.lblEstado = new System.Windows.Forms.Label();
			this.lblUsuario = new System.Windows.Forms.Label();
			this.lblIva = new System.Windows.Forms.Label();
			this.txtTotal = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtIvaTotal = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtDescIvaTotal = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtIva = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.label1 = new System.Windows.Forms.Label();
			this.txtNumero = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.label2 = new System.Windows.Forms.Label();
			this.MnuAprobacion = new System.Windows.Forms.ContextMenu();
			this.mnuAprobacionAprobar = new System.Windows.Forms.MenuItem();
			this.mnuAprobacionRechazar = new System.Windows.Forms.MenuItem();
			this.cdsSeteoF = new C1.Data.C1DataSet();
			this.gbBotones = new System.Windows.Forms.GroupBox();
			this.label3 = new System.Windows.Forms.Label();
			this.btnNuevo = new System.Windows.Forms.Button();
			this.btnConsultar = new System.Windows.Forms.Button();
			this.btnEditar = new System.Windows.Forms.Button();
			this.btnImprimir = new System.Windows.Forms.Button();
			this.btnGuardar = new System.Windows.Forms.Button();
			this.btnAnular = new System.Windows.Forms.Button();
			this.btnCancelar = new System.Windows.Forms.Button();
			this.btnSalir = new System.Windows.Forms.Button();
			this.txtNumIdProyecto = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtIdComprobante = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtNumEstado = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtIdTipoFactura = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtIdSucursal = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtNumContadoCredito = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtNumIdCompra = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtNumIdCliente = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtIdFormaPago = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtidTipoRuc = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtComprobante = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.lblComprobante = new System.Windows.Forms.Label();
			this.txtTotalCuotas = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.label4 = new System.Windows.Forms.Label();
			this.btnAprobacion = new System.Windows.Forms.Button();
			this.txtInteres = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.ultraDataSource1 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.txtidEntidadFinanciera = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.btnClaveDescuento = new System.Windows.Forms.Button();
			this.chkVerificar = new Infragistics.Win.UltraWinEditors.UltraCheckEditor();
			this.txtObservaciones = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.label5 = new System.Windows.Forms.Label();
			this.chkAutorizar = new Infragistics.Win.UltraWinEditors.UltraCheckEditor();
			this.txtBodegaPredef = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.lblIva0 = new System.Windows.Forms.Label();
			this.label55 = new System.Windows.Forms.Label();
			this.label8 = new System.Windows.Forms.Label();
			this.label9 = new System.Windows.Forms.Label();
			this.txtDescIva0Total = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtIva0 = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.label15 = new System.Windows.Forms.Label();
			this.cmbBodega1 = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.btnValidarMail = new System.Windows.Forms.Button();
			this.txtProteccionDD = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.lblClienteReiterativo = new System.Windows.Forms.Label();
			this.dtFecha = new Infragistics.Win.UltraWinSchedule.UltraCalendarCombo();
			this.lblClienteEntidad = new System.Windows.Forms.Label();
			this.chkAnclar = new Infragistics.Win.UltraWinEditors.UltraCheckEditor();
			this.txtSerieNuovo = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.label7 = new System.Windows.Forms.Label();
			this.lblColateral = new System.Windows.Forms.Label();
			this.txtSerieNuovo2 = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.label10 = new System.Windows.Forms.Label();
			this.cdsSeteo = new C1.Data.C1DataSet();
			this.chkFreelance = new Infragistics.Win.UltraWinEditors.UltraCheckEditor();
			this.btnCrediPoint = new System.Windows.Forms.Button();
			this.lblRecurrente = new System.Windows.Forms.Label();
			((System.ComponentModel.ISupportInitialize)(this.txtNumCuotas)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtCuotaDeEntrada)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbVendedor)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtRuc)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbBodega)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNombre)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraGrid2)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource2)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNotas)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtCuotaAsignada)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtCuotaDisponible)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtCuotaUsada)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtValorCuotas)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtTotal)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIvaTotal)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtDescIvaTotal)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIva)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNumero)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNumIdProyecto)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIdComprobante)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNumEstado)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIdTipoFactura)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIdSucursal)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNumContadoCredito)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNumIdCompra)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNumIdCliente)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIdFormaPago)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtidTipoRuc)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtComprobante)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtTotalCuotas)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtInteres)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource1)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtidEntidadFinanciera)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtObservaciones)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtBodegaPredef)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtDescIva0Total)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIva0)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbBodega1)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtProteccionDD)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.dtFecha)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtSerieNuovo)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtSerieNuovo2)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteo)).BeginInit();
			this.SuspendLayout();
			// 
			// label21
			// 
			this.label21.AutoSize = true;
			this.label21.BackColor = System.Drawing.Color.Transparent;
			this.label21.Location = new System.Drawing.Point(8, 75);
			this.label21.Name = "label21";
			this.label21.Size = new System.Drawing.Size(38, 16);
			this.label21.TabIndex = 127;
			this.label21.Text = "Cuotas";
			// 
			// label61
			// 
			this.label61.AutoSize = true;
			this.label61.BackColor = System.Drawing.Color.Transparent;
			this.label61.Location = new System.Drawing.Point(136, 75);
			this.label61.Name = "label61";
			this.label61.Size = new System.Drawing.Size(43, 16);
			this.label61.TabIndex = 129;
			this.label61.Text = "Entrada";
			// 
			// txtNumCuotas
			// 
			appearance1.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtNumCuotas.Appearance = appearance1;
			this.txtNumCuotas.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtNumCuotas.Enabled = false;
			this.txtNumCuotas.FormatString = "";
			this.txtNumCuotas.Location = new System.Drawing.Point(56, 72);
			this.txtNumCuotas.MaskInput = "nn";
			this.txtNumCuotas.Name = "txtNumCuotas";
			this.txtNumCuotas.PromptChar = ' ';
			this.txtNumCuotas.Size = new System.Drawing.Size(64, 22);
			this.txtNumCuotas.SpinButtonDisplayStyle = Infragistics.Win.ButtonDisplayStyle.Always;
			this.txtNumCuotas.TabIndex = 5;
			this.txtNumCuotas.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtNumCuotas_KeyDown);
			this.txtNumCuotas.Click += new System.EventHandler(this.txtNumCuotas_Click);
			this.txtNumCuotas.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.txtNumCuotas_KeyPress);
			this.txtNumCuotas.Validating += new System.ComponentModel.CancelEventHandler(this.txtNumCuotas_Validating);
			this.txtNumCuotas.Validated += new System.EventHandler(this.txtNumCuotas_Validated);
			this.txtNumCuotas.ValueChanged += new System.EventHandler(this.txtNumCuotas_ValueChanged);
			this.txtNumCuotas.Enter += new System.EventHandler(this.txtNumCuotas_Enter);
			// 
			// txtCuotaDeEntrada
			// 
			appearance2.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtCuotaDeEntrada.Appearance = appearance2;
			this.txtCuotaDeEntrada.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtCuotaDeEntrada.Enabled = false;
			this.txtCuotaDeEntrada.FormatString = "#,##0.00";
			this.txtCuotaDeEntrada.Location = new System.Drawing.Point(192, 72);
			this.txtCuotaDeEntrada.Name = "txtCuotaDeEntrada";
			this.txtCuotaDeEntrada.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtCuotaDeEntrada.PromptChar = ' ';
			this.txtCuotaDeEntrada.Size = new System.Drawing.Size(80, 22);
			this.txtCuotaDeEntrada.TabIndex = 6;
			this.txtCuotaDeEntrada.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtCuotaDeEntrada_KeyDown);
			this.txtCuotaDeEntrada.Click += new System.EventHandler(this.txtCuotaDeEntrada_Click);
			this.txtCuotaDeEntrada.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.txtCuotaDeEntrada_KeyPress);
			this.txtCuotaDeEntrada.Validated += new System.EventHandler(this.txtCuotaDeEntrada_Validated);
			this.txtCuotaDeEntrada.ValueChanged += new System.EventHandler(this.txtCuotaDeEntrada_ValueChanged);
			this.txtCuotaDeEntrada.Enter += new System.EventHandler(this.txtCuotaDeEntrada_Enter);
			// 
			// lblVendedor
			// 
			this.lblVendedor.AutoSize = true;
			this.lblVendedor.Location = new System.Drawing.Point(688, 11);
			this.lblVendedor.Name = "lblVendedor";
			this.lblVendedor.Size = new System.Drawing.Size(52, 16);
			this.lblVendedor.TabIndex = 157;
			this.lblVendedor.Text = "Vendedor";
			// 
			// label6
			// 
			this.label6.AutoSize = true;
			this.label6.Location = new System.Drawing.Point(8, 11);
			this.label6.Name = "label6";
			this.label6.Size = new System.Drawing.Size(34, 16);
			this.label6.TabIndex = 153;
			this.label6.Text = "Fecha";
			// 
			// label62
			// 
			this.label62.AutoSize = true;
			this.label62.BackColor = System.Drawing.Color.Transparent;
			this.label62.Location = new System.Drawing.Point(288, 75);
			this.label62.Name = "label62";
			this.label62.Size = new System.Drawing.Size(67, 16);
			this.label62.TabIndex = 155;
			this.label62.Text = "D?a de pago";
			// 
			// cmbVendedor
			// 
			appearance3.ForeColorDisabled = System.Drawing.Color.Black;
			this.cmbVendedor.Appearance = appearance3;
			this.cmbVendedor.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbVendedor.Cursor = System.Windows.Forms.Cursors.Default;
			ultraGridColumn1.Header.VisiblePosition = 0;
			ultraGridColumn1.Hidden = true;
			ultraGridColumn2.Header.VisiblePosition = 1;
			ultraGridColumn2.Width = 335;
			ultraGridColumn3.Header.VisiblePosition = 2;
			ultraGridBand1.Columns.AddRange(new object[] {
																										 ultraGridColumn1,
																										 ultraGridColumn2,
																										 ultraGridColumn3});
			this.cmbVendedor.DisplayLayout.BandsSerializer.Add(ultraGridBand1);
			this.cmbVendedor.DisplayMember = "Nombre";
			this.cmbVendedor.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbVendedor.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbVendedor.Enabled = false;
			this.cmbVendedor.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbVendedor.Location = new System.Drawing.Point(752, 8);
			this.cmbVendedor.Name = "cmbVendedor";
			this.cmbVendedor.Size = new System.Drawing.Size(256, 22);
			this.cmbVendedor.TabIndex = 3;
			this.cmbVendedor.ValueMember = "idPersonal";
			this.cmbVendedor.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.cmbVendedor_KeyPress);
			this.cmbVendedor.Validating += new System.ComponentModel.CancelEventHandler(this.cmbVendedor_Validating);
			this.cmbVendedor.ValueChanged += new System.EventHandler(this.cmbVendedor_ValueChanged);
			// 
			// cboFechaPrimerPago
			// 
			this.cboFechaPrimerPago.Enabled = false;
			this.cboFechaPrimerPago.Location = new System.Drawing.Point(360, 73);
			this.cboFechaPrimerPago.MaxDropDownItems = 15;
			this.cboFechaPrimerPago.Name = "cboFechaPrimerPago";
			this.cboFechaPrimerPago.Size = new System.Drawing.Size(100, 21);
			this.cboFechaPrimerPago.TabIndex = 7;
			this.cboFechaPrimerPago.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.cboFechaPrimerPago_KeyPress);
			this.cboFechaPrimerPago.Validating += new System.ComponentModel.CancelEventHandler(this.cboFechaPrimerPago_Validating);
			this.cboFechaPrimerPago.SelectedIndexChanged += new System.EventHandler(this.cboFechaPrimerPago_SelectedIndexChanged);
			// 
			// lblIdentificacion
			// 
			this.lblIdentificacion.AutoSize = true;
			this.lblIdentificacion.Location = new System.Drawing.Point(8, 43);
			this.lblIdentificacion.Name = "lblIdentificacion";
			this.lblIdentificacion.Size = new System.Drawing.Size(38, 16);
			this.lblIdentificacion.TabIndex = 163;
			this.lblIdentificacion.Text = "Cédula";
			this.lblIdentificacion.Click += new System.EventHandler(this.lblIdentificacion_Click);
			// 
			// lblCliente
			// 
			this.lblCliente.AutoSize = true;
			this.lblCliente.Location = new System.Drawing.Point(176, 43);
			this.lblCliente.Name = "lblCliente";
			this.lblCliente.Size = new System.Drawing.Size(44, 16);
			this.lblCliente.TabIndex = 161;
			this.lblCliente.Text = "Nombre";
			// 
			// lblBodega
			// 
			this.lblBodega.AutoSize = true;
			this.lblBodega.Location = new System.Drawing.Point(176, 11);
			this.lblBodega.Name = "lblBodega";
			this.lblBodega.Size = new System.Drawing.Size(30, 16);
			this.lblBodega.TabIndex = 160;
			this.lblBodega.Text = "Local";
			this.lblBodega.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
			// 
			// txtRuc
			// 
			appearance4.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtRuc.Appearance = appearance4;
			this.txtRuc.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtRuc.Enabled = false;
			this.txtRuc.Location = new System.Drawing.Point(56, 40);
			this.txtRuc.MaxLength = 10;
			this.txtRuc.Name = "txtRuc";
			this.txtRuc.Size = new System.Drawing.Size(112, 22);
			this.txtRuc.TabIndex = 4;
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
			ultraGridColumn4.Header.VisiblePosition = 2;
			ultraGridColumn5.Header.VisiblePosition = 1;
			ultraGridColumn5.Width = 180;
			ultraGridColumn6.Header.VisiblePosition = 0;
			ultraGridBand2.Columns.AddRange(new object[] {
																										 ultraGridColumn4,
																										 ultraGridColumn5,
																										 ultraGridColumn6});
			this.cmbBodega.DisplayLayout.BandsSerializer.Add(ultraGridBand2);
			this.cmbBodega.DisplayMember = "Nombre";
			this.cmbBodega.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbBodega.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbBodega.Enabled = false;
			this.cmbBodega.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbBodega.Location = new System.Drawing.Point(216, 8);
			this.cmbBodega.MaxDropDownItems = 30;
			this.cmbBodega.Name = "cmbBodega";
			this.cmbBodega.Size = new System.Drawing.Size(180, 22);
			this.cmbBodega.TabIndex = 2;
			this.cmbBodega.ValueMember = "Bodega";
			this.cmbBodega.Validating += new System.ComponentModel.CancelEventHandler(this.cmbBodega_Validating);
			this.cmbBodega.ValueChanged += new System.EventHandler(this.cmbBodega_ValueChanged);
			this.cmbBodega.InitializeLayout += new Infragistics.Win.UltraWinGrid.InitializeLayoutEventHandler(this.cmbBodega_InitializeLayout);
			this.cmbBodega.KeyUp += new System.Windows.Forms.KeyEventHandler(this.cmbBodega_KeyUp);
			// 
			// txtNombre
			// 
			appearance5.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtNombre.Appearance = appearance5;
			this.txtNombre.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtNombre.Enabled = false;
			this.txtNombre.Location = new System.Drawing.Point(232, 40);
			this.txtNombre.Name = "txtNombre";
			this.txtNombre.Size = new System.Drawing.Size(264, 22);
			this.txtNombre.TabIndex = 159;
			// 
			// ultraGrid2
			// 
			this.ultraGrid2.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.ultraGrid2.Cursor = System.Windows.Forms.Cursors.Default;
			this.ultraGrid2.DataSource = this.ultraDataSource2;
			appearance6.BackColor = System.Drawing.Color.White;
			this.ultraGrid2.DisplayLayout.Appearance = appearance6;
			this.ultraGrid2.DisplayLayout.AutoFitColumns = true;
			ultraGridBand3.AddButtonCaption = "_DetCompra";
			ultraGridColumn7.Header.VisiblePosition = 10;
			ultraGridColumn7.Hidden = true;
			ultraGridColumn8.Header.VisiblePosition = 13;
			ultraGridColumn8.Hidden = true;
			ultraGridColumn9.Header.VisiblePosition = 0;
			ultraGridColumn9.Hidden = true;
			ultraGridColumn10.Header.VisiblePosition = 1;
			ultraGridColumn10.Hidden = true;
			ultraGridColumn10.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.DropDownList;
			ultraGridColumn10.Width = 194;
			ultraGridColumn11.Header.Caption = "Unidad";
			ultraGridColumn11.Header.VisiblePosition = 5;
			ultraGridColumn11.Hidden = true;
			ultraGridColumn11.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.DropDownValidate;
			ultraGridColumn11.Width = 53;
			ultraGridColumn12.Case = Infragistics.Win.UltraWinGrid.Case.Upper;
			ultraGridColumn12.Header.VisiblePosition = 2;
			ultraGridColumn12.Width = 260;
			ultraGridColumn13.Case = Infragistics.Win.UltraWinGrid.Case.Upper;
			ultraGridColumn13.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn13.Header.VisiblePosition = 3;
			ultraGridColumn13.Width = 348;
			ultraGridColumn14.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance7.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn14.CellAppearance = appearance7;
			ultraGridColumn14.Format = "";
			ultraGridColumn14.Header.VisiblePosition = 4;
			ultraGridColumn14.MaskInput = "";
			ultraGridColumn14.NullText = "0";
			ultraGridColumn14.PromptChar = ' ';
			ultraGridColumn14.Width = 75;
			appearance8.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn15.CellAppearance = appearance8;
			ultraGridColumn15.Format = "#,##0.00";
			ultraGridColumn15.Header.VisiblePosition = 6;
			ultraGridColumn15.MaskInput = "{LOC}nn,nnn,nnn.nn";
			ultraGridColumn15.PromptChar = ' ';
			ultraGridColumn15.Width = 85;
			appearance9.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn16.CellAppearance = appearance9;
			ultraGridColumn16.Format = "#,##0.00";
			ultraGridColumn16.Header.Caption = "Iva";
			ultraGridColumn16.Header.VisiblePosition = 7;
			ultraGridColumn16.MaskInput = "{LOC}nnnnnnn.nn";
			ultraGridColumn16.PromptChar = ' ';
			ultraGridColumn16.Width = 60;
			ultraGridColumn17.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance10.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn17.CellAppearance = appearance10;
			ultraGridColumn17.Header.Caption = "Descuento";
			ultraGridColumn17.Header.VisiblePosition = 8;
			ultraGridColumn17.Hidden = true;
			ultraGridColumn17.MaskInput = "{LOC}nn,nnn,nnn.nn";
			ultraGridColumn17.PromptChar = ' ';
			ultraGridColumn17.Width = 61;
			appearance11.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn18.CellAppearance = appearance11;
			ultraGridColumn18.Format = "#,##0.0000";
			ultraGridColumn18.Header.Caption = "% Desc.";
			ultraGridColumn18.Header.VisiblePosition = 9;
			ultraGridColumn18.MaskInput = "";
			ultraGridColumn18.PromptChar = ' ';
			ultraGridColumn18.Width = 60;
			ultraGridColumn19.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance12.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn19.CellAppearance = appearance12;
			ultraGridColumn19.Format = "#,##0.00";
			ultraGridColumn19.Header.VisiblePosition = 11;
			ultraGridColumn19.PromptChar = ' ';
			ultraGridColumn19.Width = 95;
			ultraGridColumn20.Header.VisiblePosition = 12;
			ultraGridColumn20.Hidden = true;
			ultraGridColumn21.Header.VisiblePosition = 14;
			ultraGridColumn21.Hidden = true;
			ultraGridColumn21.Width = 41;
			ultraGridColumn22.Header.VisiblePosition = 15;
			ultraGridColumn22.Hidden = true;
			ultraGridColumn22.Width = 47;
			ultraGridColumn23.Header.VisiblePosition = 16;
			ultraGridColumn23.Hidden = true;
			ultraGridColumn23.Width = 41;
			ultraGridColumn24.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn24.Header.VisiblePosition = 17;
			ultraGridColumn24.Hidden = true;
			ultraGridColumn24.Width = 38;
			ultraGridColumn25.Header.VisiblePosition = 18;
			ultraGridColumn25.Hidden = true;
			ultraGridColumn25.Width = 47;
			ultraGridColumn26.Header.VisiblePosition = 19;
			ultraGridColumn26.Hidden = true;
			ultraGridColumn26.Width = 42;
			ultraGridColumn27.Header.VisiblePosition = 20;
			ultraGridColumn27.Hidden = true;
			ultraGridColumn27.Width = 54;
			ultraGridColumn28.Header.VisiblePosition = 21;
			ultraGridColumn28.Hidden = true;
			ultraGridColumn28.Width = 64;
			ultraGridColumn29.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			ultraGridColumn29.Format = "#,##0.00";
			ultraGridColumn29.Header.VisiblePosition = 22;
			ultraGridColumn29.Hidden = true;
			ultraGridColumn29.Width = 95;
			ultraGridColumn30.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			ultraGridColumn30.Header.VisiblePosition = 23;
			ultraGridColumn30.Hidden = true;
			ultraGridColumn30.Width = 83;
			ultraGridBand3.Columns.AddRange(new object[] {
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
			ultraGridBand3.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.TemplateOnBottom;
			this.ultraGrid2.DisplayLayout.BandsSerializer.Add(ultraGridBand3);
			appearance13.ForeColor = System.Drawing.Color.Black;
			appearance13.ForeColorDisabled = System.Drawing.Color.Black;
			this.ultraGrid2.DisplayLayout.Override.ActiveRowAppearance = appearance13;
			this.ultraGrid2.DisplayLayout.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.TemplateOnBottom;
			this.ultraGrid2.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.True;
			this.ultraGrid2.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.True;
			appearance14.BackColor = System.Drawing.Color.Transparent;
			this.ultraGrid2.DisplayLayout.Override.CardAreaAppearance = appearance14;
			appearance15.ForeColor = System.Drawing.Color.Black;
			appearance15.ForeColorDisabled = System.Drawing.Color.Black;
			this.ultraGrid2.DisplayLayout.Override.CellAppearance = appearance15;
			appearance16.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance16.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance16.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance16.FontData.BoldAsString = "True";
			appearance16.FontData.Name = "Arial";
			appearance16.FontData.SizeInPoints = 10F;
			appearance16.ForeColor = System.Drawing.Color.White;
			appearance16.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.ultraGrid2.DisplayLayout.Override.HeaderAppearance = appearance16;
			appearance17.BackColor = System.Drawing.SystemColors.ControlLight;
			appearance17.BackColor2 = System.Drawing.Color.Gainsboro;
			appearance17.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ultraGrid2.DisplayLayout.Override.RowAlternateAppearance = appearance17;
			appearance18.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance18.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance18.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ultraGrid2.DisplayLayout.Override.RowSelectorAppearance = appearance18;
			appearance19.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance19.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance19.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ultraGrid2.DisplayLayout.Override.SelectedRowAppearance = appearance19;
			this.ultraGrid2.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.ultraGrid2.Location = new System.Drawing.Point(6, 136);
			this.ultraGrid2.Name = "ultraGrid2";
			this.ultraGrid2.Size = new System.Drawing.Size(1004, 184);
			this.ultraGrid2.TabIndex = 10;
			this.ultraGrid2.Error += new Infragistics.Win.UltraWinGrid.ErrorEventHandler(this.ultraGrid1_Error);
			this.ultraGrid2.DoubleClick += new System.EventHandler(this.ultraGrid1_DoubleClick);
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
			ultraDataColumn8.DataType = typeof(int);
			ultraDataColumn8.DefaultValue = 0;
			ultraDataColumn9.DataType = typeof(System.Decimal);
			ultraDataColumn9.DefaultValue = new System.Decimal(new int[] {
																																		 0,
																																		 0,
																																		 0,
																																		 0});
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
			ultraDataColumn14.DataType = typeof(bool);
			ultraDataColumn14.DefaultValue = false;
			ultraDataColumn15.DataType = typeof(int);
			ultraDataColumn15.DefaultValue = 0;
			ultraDataColumn16.DataType = typeof(bool);
			ultraDataColumn16.DefaultValue = false;
			ultraDataColumn17.DataType = typeof(int);
			ultraDataColumn17.DefaultValue = 0;
			ultraDataColumn18.DataType = typeof(bool);
			ultraDataColumn18.DefaultValue = false;
			ultraDataColumn19.DataType = typeof(bool);
			ultraDataColumn19.DefaultValue = false;
			ultraDataColumn20.DataType = typeof(int);
			ultraDataColumn20.DefaultValue = 0;
			ultraDataColumn21.DataType = typeof(int);
			ultraDataColumn21.DefaultValue = 0;
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
			this.ultraDataSource2.Band.Columns.AddRange(new object[] {
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
			// txtNotas
			// 
			this.txtNotas.AcceptsReturn = true;
			this.txtNotas.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			appearance20.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtNotas.Appearance = appearance20;
			this.txtNotas.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtNotas.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtNotas.Enabled = false;
			this.txtNotas.Location = new System.Drawing.Point(56, 336);
			this.txtNotas.Multiline = true;
			this.txtNotas.Name = "txtNotas";
			this.txtNotas.Size = new System.Drawing.Size(408, 63);
			this.txtNotas.TabIndex = 11;
			this.txtNotas.Leave += new System.EventHandler(this.txtNotas_Leave);
			this.txtNotas.ValueChanged += new System.EventHandler(this.txtNotas_ValueChanged);
			// 
			// txtCuotaAsignada
			// 
			appearance21.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtCuotaAsignada.Appearance = appearance21;
			this.txtCuotaAsignada.BorderStyle = Infragistics.Win.UIElementBorderStyle.Solid;
			this.txtCuotaAsignada.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtCuotaAsignada.Enabled = false;
			this.txtCuotaAsignada.FormatString = "#,##0.00";
			this.txtCuotaAsignada.Location = new System.Drawing.Point(600, 41);
			this.txtCuotaAsignada.Name = "txtCuotaAsignada";
			this.txtCuotaAsignada.NullText = "0.00";
			this.txtCuotaAsignada.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtCuotaAsignada.PromptChar = ' ';
			this.txtCuotaAsignada.Size = new System.Drawing.Size(80, 20);
			this.txtCuotaAsignada.TabIndex = 177;
			// 
			// txtCuotaDisponible
			// 
			appearance22.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtCuotaDisponible.Appearance = appearance22;
			this.txtCuotaDisponible.BorderStyle = Infragistics.Win.UIElementBorderStyle.Solid;
			this.txtCuotaDisponible.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtCuotaDisponible.Enabled = false;
			this.txtCuotaDisponible.FormatString = "#,##0.00";
			this.txtCuotaDisponible.Location = new System.Drawing.Point(784, 41);
			this.txtCuotaDisponible.Name = "txtCuotaDisponible";
			this.txtCuotaDisponible.NullText = "0.00";
			this.txtCuotaDisponible.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtCuotaDisponible.PromptChar = ' ';
			this.txtCuotaDisponible.Size = new System.Drawing.Size(80, 20);
			this.txtCuotaDisponible.TabIndex = 176;
			// 
			// txtCuotaUsada
			// 
			appearance23.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtCuotaUsada.Appearance = appearance23;
			this.txtCuotaUsada.BorderStyle = Infragistics.Win.UIElementBorderStyle.Solid;
			this.txtCuotaUsada.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtCuotaUsada.Enabled = false;
			this.txtCuotaUsada.FormatString = "#,##0.00";
			this.txtCuotaUsada.Location = new System.Drawing.Point(600, 73);
			this.txtCuotaUsada.Name = "txtCuotaUsada";
			this.txtCuotaUsada.NullText = "0.00";
			this.txtCuotaUsada.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtCuotaUsada.PromptChar = ' ';
			this.txtCuotaUsada.Size = new System.Drawing.Size(80, 20);
			this.txtCuotaUsada.TabIndex = 175;
			// 
			// label63
			// 
			this.label63.AutoSize = true;
			this.label63.BackColor = System.Drawing.Color.Transparent;
			this.label63.Location = new System.Drawing.Point(688, 43);
			this.label63.Name = "label63";
			this.label63.Size = new System.Drawing.Size(87, 16);
			this.label63.TabIndex = 174;
			this.label63.Text = "Cuota disponible";
			// 
			// label64
			// 
			this.label64.AutoSize = true;
			this.label64.BackColor = System.Drawing.Color.Transparent;
			this.label64.Location = new System.Drawing.Point(504, 75);
			this.label64.Name = "label64";
			this.label64.Size = new System.Drawing.Size(66, 16);
			this.label64.TabIndex = 173;
			this.label64.Text = "Cuota usada";
			// 
			// label65
			// 
			this.label65.AutoSize = true;
			this.label65.BackColor = System.Drawing.Color.Transparent;
			this.label65.Location = new System.Drawing.Point(504, 43);
			this.label65.Name = "label65";
			this.label65.Size = new System.Drawing.Size(81, 16);
			this.label65.TabIndex = 172;
			this.label65.Text = "Cuota asignada";
			// 
			// txtValorCuotas
			// 
			appearance24.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtValorCuotas.Appearance = appearance24;
			this.txtValorCuotas.BorderStyle = Infragistics.Win.UIElementBorderStyle.Solid;
			this.txtValorCuotas.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtValorCuotas.Enabled = false;
			this.txtValorCuotas.FormatString = "#,##0.00";
			this.txtValorCuotas.Location = new System.Drawing.Point(784, 73);
			this.txtValorCuotas.Name = "txtValorCuotas";
			this.txtValorCuotas.NullText = "0.00";
			this.txtValorCuotas.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtValorCuotas.PromptChar = ' ';
			this.txtValorCuotas.Size = new System.Drawing.Size(80, 20);
			this.txtValorCuotas.TabIndex = 170;
			this.txtValorCuotas.ValueChanged += new System.EventHandler(this.txtValorCuotas_ValueChanged);
			// 
			// label23
			// 
			this.label23.AutoSize = true;
			this.label23.BackColor = System.Drawing.Color.Transparent;
			this.label23.Location = new System.Drawing.Point(688, 75);
			this.label23.Name = "label23";
			this.label23.Size = new System.Drawing.Size(62, 16);
			this.label23.TabIndex = 168;
			this.label23.Text = "Valor Cuota";
			// 
			// lblEstado
			// 
			this.lblEstado.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			this.lblEstado.AutoSize = true;
			this.lblEstado.Font = new System.Drawing.Font("Microsoft Sans Serif", 15F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.lblEstado.ForeColor = System.Drawing.Color.Firebrick;
			this.lblEstado.Location = new System.Drawing.Point(784, 96);
			this.lblEstado.Name = "lblEstado";
			this.lblEstado.Size = new System.Drawing.Size(0, 26);
			this.lblEstado.TabIndex = 179;
			this.lblEstado.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
			// 
			// lblUsuario
			// 
			this.lblUsuario.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.lblUsuario.AutoSize = true;
			this.lblUsuario.BackColor = System.Drawing.Color.Transparent;
			this.lblUsuario.Font = new System.Drawing.Font("Tahoma", 10F, ((System.Drawing.FontStyle)((System.Drawing.FontStyle.Bold | System.Drawing.FontStyle.Underline))));
			this.lblUsuario.Location = new System.Drawing.Point(480, 328);
			this.lblUsuario.Name = "lblUsuario";
			this.lblUsuario.Size = new System.Drawing.Size(0, 20);
			this.lblUsuario.TabIndex = 180;
			this.lblUsuario.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
			// 
			// lblIva
			// 
			this.lblIva.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			this.lblIva.AutoSize = true;
			this.lblIva.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.lblIva.Location = new System.Drawing.Point(736, 411);
			this.lblIva.Name = "lblIva";
			this.lblIva.Size = new System.Drawing.Size(22, 16);
			this.lblIva.TabIndex = 184;
			this.lblIva.Text = "IVA";
			// 
			// txtTotal
			// 
			this.txtTotal.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			appearance25.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtTotal.Appearance = appearance25;
			this.txtTotal.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtTotal.Enabled = false;
			this.txtTotal.Font = new System.Drawing.Font("Microsoft Sans Serif", 15.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.txtTotal.FormatString = "#,##0.00";
			this.txtTotal.Location = new System.Drawing.Point(824, 432);
			this.txtTotal.Name = "txtTotal";
			this.txtTotal.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtTotal.PromptChar = ' ';
			this.txtTotal.Size = new System.Drawing.Size(184, 34);
			this.txtTotal.TabIndex = 191;
			// 
			// txtIvaTotal
			// 
			this.txtIvaTotal.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			appearance26.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtIvaTotal.Appearance = appearance26;
			this.txtIvaTotal.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtIvaTotal.Enabled = false;
			this.txtIvaTotal.FormatString = "#,##0.00";
			this.txtIvaTotal.Location = new System.Drawing.Point(928, 408);
			this.txtIvaTotal.Name = "txtIvaTotal";
			this.txtIvaTotal.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtIvaTotal.PromptChar = ' ';
			this.txtIvaTotal.Size = new System.Drawing.Size(80, 22);
			this.txtIvaTotal.TabIndex = 190;
			// 
			// txtDescIvaTotal
			// 
			this.txtDescIvaTotal.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			appearance27.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtDescIvaTotal.Appearance = appearance27;
			this.txtDescIvaTotal.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtDescIvaTotal.Enabled = false;
			this.txtDescIvaTotal.FormatString = "#,##0.00";
			this.txtDescIvaTotal.Location = new System.Drawing.Point(928, 384);
			this.txtDescIvaTotal.Name = "txtDescIvaTotal";
			this.txtDescIvaTotal.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtDescIvaTotal.PromptChar = ' ';
			this.txtDescIvaTotal.Size = new System.Drawing.Size(80, 22);
			this.txtDescIvaTotal.TabIndex = 188;
			// 
			// txtIva
			// 
			this.txtIva.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			appearance28.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtIva.Appearance = appearance28;
			this.txtIva.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtIva.Enabled = false;
			this.txtIva.FormatString = "#,##0.00";
			this.txtIva.Location = new System.Drawing.Point(928, 360);
			this.txtIva.Name = "txtIva";
			this.txtIva.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtIva.PromptChar = ' ';
			this.txtIva.Size = new System.Drawing.Size(80, 22);
			this.txtIva.TabIndex = 187;
			// 
			// label1
			// 
			this.label1.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			this.label1.AutoSize = true;
			this.label1.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label1.Location = new System.Drawing.Point(736, 441);
			this.label1.Name = "label1";
			this.label1.Size = new System.Drawing.Size(29, 16);
			this.label1.TabIndex = 192;
			this.label1.Text = "Total";
			// 
			// txtNumero
			// 
			appearance29.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtNumero.Appearance = appearance29;
			this.txtNumero.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtNumero.Enabled = false;
			this.txtNumero.Location = new System.Drawing.Point(64, 104);
			this.txtNumero.Name = "txtNumero";
			this.txtNumero.Size = new System.Drawing.Size(112, 22);
			this.txtNumero.TabIndex = 193;
			// 
			// label2
			// 
			this.label2.AutoSize = true;
			this.label2.BackColor = System.Drawing.Color.Transparent;
			this.label2.Location = new System.Drawing.Point(8, 107);
			this.label2.Name = "label2";
			this.label2.Size = new System.Drawing.Size(44, 16);
			this.label2.TabIndex = 194;
			this.label2.Text = "Número";
			// 
			// MnuAprobacion
			// 
			this.MnuAprobacion.MenuItems.AddRange(new System.Windows.Forms.MenuItem[] {
																																									this.mnuAprobacionAprobar,
																																									this.mnuAprobacionRechazar});
			this.MnuAprobacion.Popup += new System.EventHandler(this.MnuAprobacion_Popup);
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
			// gbBotones
			// 
			this.gbBotones.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.gbBotones.Location = new System.Drawing.Point(0, 472);
			this.gbBotones.Name = "gbBotones";
			this.gbBotones.Size = new System.Drawing.Size(1016, 8);
			this.gbBotones.TabIndex = 290;
			this.gbBotones.TabStop = false;
			// 
			// label3
			// 
			this.label3.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			this.label3.AutoSize = true;
			this.label3.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.5F, ((System.Drawing.FontStyle)((System.Drawing.FontStyle.Bold | System.Drawing.FontStyle.Underline))), System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label3.ForeColor = System.Drawing.Color.Firebrick;
			this.label3.Location = new System.Drawing.Point(656, 496);
			this.label3.Name = "label3";
			this.label3.Size = new System.Drawing.Size(0, 18);
			this.label3.TabIndex = 285;
			this.label3.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
			// 
			// btnNuevo
			// 
			this.btnNuevo.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btnNuevo.CausesValidation = false;
			this.btnNuevo.Image = ((System.Drawing.Image)(resources.GetObject("btnNuevo.Image")));
			this.btnNuevo.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnNuevo.Location = new System.Drawing.Point(8, 488);
			this.btnNuevo.Name = "btnNuevo";
			this.btnNuevo.Size = new System.Drawing.Size(72, 23);
			this.btnNuevo.TabIndex = 0;
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
			this.btnConsultar.Location = new System.Drawing.Point(82, 488);
			this.btnConsultar.Name = "btnConsultar";
			this.btnConsultar.Size = new System.Drawing.Size(78, 23);
			this.btnConsultar.TabIndex = 15;
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
			this.btnEditar.Location = new System.Drawing.Point(162, 488);
			this.btnEditar.Name = "btnEditar";
			this.btnEditar.Size = new System.Drawing.Size(76, 23);
			this.btnEditar.TabIndex = 14;
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
			this.btnImprimir.Location = new System.Drawing.Point(240, 488);
			this.btnImprimir.Name = "btnImprimir";
			this.btnImprimir.Size = new System.Drawing.Size(77, 23);
			this.btnImprimir.TabIndex = 13;
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
			this.btnGuardar.Location = new System.Drawing.Point(319, 488);
			this.btnGuardar.Name = "btnGuardar";
			this.btnGuardar.Size = new System.Drawing.Size(76, 23);
			this.btnGuardar.TabIndex = 13;
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
			this.btnAnular.Location = new System.Drawing.Point(482, 488);
			this.btnAnular.Name = "btnAnular";
			this.btnAnular.Size = new System.Drawing.Size(76, 23);
			this.btnAnular.TabIndex = 17;
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
			this.btnCancelar.Location = new System.Drawing.Point(560, 488);
			this.btnCancelar.Name = "btnCancelar";
			this.btnCancelar.Size = new System.Drawing.Size(72, 23);
			this.btnCancelar.TabIndex = 18;
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
			this.btnSalir.Location = new System.Drawing.Point(634, 488);
			this.btnSalir.Name = "btnSalir";
			this.btnSalir.TabIndex = 19;
			this.btnSalir.Text = "&Salir";
			this.btnSalir.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnSalir.Click += new System.EventHandler(this.btnSalir_Click);
			// 
			// txtNumIdProyecto
			// 
			this.txtNumIdProyecto.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			this.txtNumIdProyecto.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtNumIdProyecto.Enabled = false;
			this.txtNumIdProyecto.Location = new System.Drawing.Point(888, 496);
			this.txtNumIdProyecto.Name = "txtNumIdProyecto";
			this.txtNumIdProyecto.PromptChar = ' ';
			this.txtNumIdProyecto.Size = new System.Drawing.Size(16, 22);
			this.txtNumIdProyecto.TabIndex = 276;
			this.txtNumIdProyecto.Visible = false;
			// 
			// txtIdComprobante
			// 
			this.txtIdComprobante.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			this.txtIdComprobante.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtIdComprobante.Enabled = false;
			this.txtIdComprobante.Location = new System.Drawing.Point(968, 496);
			this.txtIdComprobante.Name = "txtIdComprobante";
			this.txtIdComprobante.PromptChar = ' ';
			this.txtIdComprobante.Size = new System.Drawing.Size(16, 22);
			this.txtIdComprobante.TabIndex = 275;
			this.txtIdComprobante.Visible = false;
			// 
			// txtNumEstado
			// 
			this.txtNumEstado.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			this.txtNumEstado.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtNumEstado.Enabled = false;
			this.txtNumEstado.Location = new System.Drawing.Point(952, 496);
			this.txtNumEstado.Name = "txtNumEstado";
			this.txtNumEstado.PromptChar = ' ';
			this.txtNumEstado.Size = new System.Drawing.Size(16, 22);
			this.txtNumEstado.TabIndex = 274;
			this.txtNumEstado.Visible = false;
			// 
			// txtIdTipoFactura
			// 
			this.txtIdTipoFactura.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			this.txtIdTipoFactura.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtIdTipoFactura.Enabled = false;
			this.txtIdTipoFactura.Location = new System.Drawing.Point(936, 496);
			this.txtIdTipoFactura.Name = "txtIdTipoFactura";
			this.txtIdTipoFactura.PromptChar = ' ';
			this.txtIdTipoFactura.Size = new System.Drawing.Size(16, 22);
			this.txtIdTipoFactura.TabIndex = 273;
			this.txtIdTipoFactura.Visible = false;
			// 
			// txtIdSucursal
			// 
			this.txtIdSucursal.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			this.txtIdSucursal.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtIdSucursal.Enabled = false;
			this.txtIdSucursal.Location = new System.Drawing.Point(904, 496);
			this.txtIdSucursal.Name = "txtIdSucursal";
			this.txtIdSucursal.PromptChar = ' ';
			this.txtIdSucursal.Size = new System.Drawing.Size(16, 22);
			this.txtIdSucursal.TabIndex = 271;
			this.txtIdSucursal.Visible = false;
			// 
			// txtNumContadoCredito
			// 
			this.txtNumContadoCredito.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			this.txtNumContadoCredito.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtNumContadoCredito.Enabled = false;
			this.txtNumContadoCredito.Location = new System.Drawing.Point(872, 496);
			this.txtNumContadoCredito.Name = "txtNumContadoCredito";
			this.txtNumContadoCredito.PromptChar = ' ';
			this.txtNumContadoCredito.Size = new System.Drawing.Size(16, 22);
			this.txtNumContadoCredito.TabIndex = 268;
			this.txtNumContadoCredito.Visible = false;
			// 
			// txtNumIdCompra
			// 
			this.txtNumIdCompra.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			this.txtNumIdCompra.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtNumIdCompra.Enabled = false;
			this.txtNumIdCompra.Location = new System.Drawing.Point(840, 496);
			this.txtNumIdCompra.Name = "txtNumIdCompra";
			this.txtNumIdCompra.PromptChar = ' ';
			this.txtNumIdCompra.Size = new System.Drawing.Size(16, 22);
			this.txtNumIdCompra.TabIndex = 266;
			this.txtNumIdCompra.Visible = false;
			// 
			// txtNumIdCliente
			// 
			this.txtNumIdCliente.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			this.txtNumIdCliente.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtNumIdCliente.Enabled = false;
			this.txtNumIdCliente.Location = new System.Drawing.Point(856, 496);
			this.txtNumIdCliente.Name = "txtNumIdCliente";
			this.txtNumIdCliente.PromptChar = ' ';
			this.txtNumIdCliente.Size = new System.Drawing.Size(16, 22);
			this.txtNumIdCliente.TabIndex = 265;
			this.txtNumIdCliente.Visible = false;
			// 
			// txtIdFormaPago
			// 
			this.txtIdFormaPago.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			this.txtIdFormaPago.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtIdFormaPago.Enabled = false;
			this.txtIdFormaPago.Location = new System.Drawing.Point(984, 496);
			this.txtIdFormaPago.Name = "txtIdFormaPago";
			this.txtIdFormaPago.PromptChar = ' ';
			this.txtIdFormaPago.Size = new System.Drawing.Size(16, 22);
			this.txtIdFormaPago.TabIndex = 291;
			this.txtIdFormaPago.Visible = false;
			// 
			// txtidTipoRuc
			// 
			this.txtidTipoRuc.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			this.txtidTipoRuc.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtidTipoRuc.Enabled = false;
			this.txtidTipoRuc.Location = new System.Drawing.Point(1000, 496);
			this.txtidTipoRuc.Name = "txtidTipoRuc";
			this.txtidTipoRuc.PromptChar = ' ';
			this.txtidTipoRuc.Size = new System.Drawing.Size(16, 22);
			this.txtidTipoRuc.TabIndex = 292;
			this.txtidTipoRuc.Visible = false;
			// 
			// txtComprobante
			// 
			appearance30.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtComprobante.Appearance = appearance30;
			this.txtComprobante.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtComprobante.Enabled = false;
			this.txtComprobante.Location = new System.Drawing.Point(288, 104);
			this.txtComprobante.Name = "txtComprobante";
			this.txtComprobante.Size = new System.Drawing.Size(112, 22);
			this.txtComprobante.TabIndex = 294;
			// 
			// lblComprobante
			// 
			this.lblComprobante.AutoSize = true;
			this.lblComprobante.BackColor = System.Drawing.Color.Transparent;
			this.lblComprobante.Location = new System.Drawing.Point(200, 107);
			this.lblComprobante.Name = "lblComprobante";
			this.lblComprobante.Size = new System.Drawing.Size(71, 16);
			this.lblComprobante.TabIndex = 293;
			this.lblComprobante.Text = "Comprobante";
			// 
			// txtTotalCuotas
			// 
			this.txtTotalCuotas.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			this.txtTotalCuotas.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtTotalCuotas.Enabled = false;
			this.txtTotalCuotas.FormatString = "#,##0.00";
			this.txtTotalCuotas.Location = new System.Drawing.Point(824, 496);
			this.txtTotalCuotas.Name = "txtTotalCuotas";
			this.txtTotalCuotas.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtTotalCuotas.PromptChar = ' ';
			this.txtTotalCuotas.Size = new System.Drawing.Size(16, 22);
			this.txtTotalCuotas.TabIndex = 295;
			this.txtTotalCuotas.Visible = false;
			// 
			// label4
			// 
			this.label4.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.label4.AutoSize = true;
			this.label4.BackColor = System.Drawing.Color.Transparent;
			this.label4.Location = new System.Drawing.Point(8, 359);
			this.label4.Name = "label4";
			this.label4.Size = new System.Drawing.Size(33, 16);
			this.label4.TabIndex = 296;
			this.label4.Text = "Notas";
			// 
			// btnAprobacion
			// 
			this.btnAprobacion.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btnAprobacion.CausesValidation = false;
			this.btnAprobacion.Enabled = false;
			this.btnAprobacion.Image = ((System.Drawing.Image)(resources.GetObject("btnAprobacion.Image")));
			this.btnAprobacion.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnAprobacion.Location = new System.Drawing.Point(397, 488);
			this.btnAprobacion.Name = "btnAprobacion";
			this.btnAprobacion.Size = new System.Drawing.Size(83, 23);
			this.btnAprobacion.TabIndex = 16;
			this.btnAprobacion.Text = "&Aprobación";
			this.btnAprobacion.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnAprobacion.Click += new System.EventHandler(this.btnAprobacion_Click);
			// 
			// txtInteres
			// 
			this.txtInteres.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			this.txtInteres.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtInteres.Enabled = false;
			this.txtInteres.FormatString = "#,##0.00";
			this.txtInteres.Location = new System.Drawing.Point(808, 496);
			this.txtInteres.Name = "txtInteres";
			this.txtInteres.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtInteres.PromptChar = ' ';
			this.txtInteres.Size = new System.Drawing.Size(16, 22);
			this.txtInteres.TabIndex = 298;
			this.txtInteres.Visible = false;
			// 
			// ultraDataSource1
			// 
			ultraDataColumn25.DataType = typeof(int);
			this.ultraDataSource1.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn25,
																																 ultraDataColumn26});
			// 
			// txtidEntidadFinanciera
			// 
			this.txtidEntidadFinanciera.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			this.txtidEntidadFinanciera.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtidEntidadFinanciera.Enabled = false;
			this.txtidEntidadFinanciera.FormatString = "#,##0.00";
			this.txtidEntidadFinanciera.Location = new System.Drawing.Point(792, 496);
			this.txtidEntidadFinanciera.Name = "txtidEntidadFinanciera";
			this.txtidEntidadFinanciera.PromptChar = ' ';
			this.txtidEntidadFinanciera.Size = new System.Drawing.Size(16, 22);
			this.txtidEntidadFinanciera.TabIndex = 300;
			this.txtidEntidadFinanciera.Visible = false;
			// 
			// btnClaveDescuento
			// 
			this.btnClaveDescuento.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btnClaveDescuento.CausesValidation = false;
			this.btnClaveDescuento.Enabled = false;
			this.btnClaveDescuento.Location = new System.Drawing.Point(712, 488);
			this.btnClaveDescuento.Name = "btnClaveDescuento";
			this.btnClaveDescuento.Size = new System.Drawing.Size(24, 22);
			this.btnClaveDescuento.TabIndex = 301;
			this.btnClaveDescuento.Text = "D";
			this.btnClaveDescuento.Visible = false;
			this.btnClaveDescuento.Click += new System.EventHandler(this.btnClaveDescuento_Click);
			// 
			// chkVerificar
			// 
			this.chkVerificar.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			appearance31.ForeColorDisabled = System.Drawing.Color.Black;
			this.chkVerificar.Appearance = appearance31;
			this.chkVerificar.BackColor = System.Drawing.Color.Transparent;
			this.chkVerificar.Enabled = false;
			this.chkVerificar.Location = new System.Drawing.Point(480, 352);
			this.chkVerificar.Name = "chkVerificar";
			this.chkVerificar.Size = new System.Drawing.Size(56, 22);
			this.chkVerificar.TabIndex = 302;
			this.chkVerificar.Text = "Carpa";
			// 
			// txtObservaciones
			// 
			this.txtObservaciones.AcceptsReturn = true;
			this.txtObservaciones.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			appearance32.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtObservaciones.Appearance = appearance32;
			this.txtObservaciones.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtObservaciones.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtObservaciones.Enabled = false;
			this.txtObservaciones.Location = new System.Drawing.Point(96, 408);
			this.txtObservaciones.Multiline = true;
			this.txtObservaciones.Name = "txtObservaciones";
			this.txtObservaciones.Size = new System.Drawing.Size(368, 63);
			this.txtObservaciones.TabIndex = 12;
			this.txtObservaciones.ValueChanged += new System.EventHandler(this.txtObservaciones_ValueChanged);
			// 
			// label5
			// 
			this.label5.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.label5.AutoSize = true;
			this.label5.BackColor = System.Drawing.Color.Transparent;
			this.label5.Location = new System.Drawing.Point(8, 431);
			this.label5.Name = "label5";
			this.label5.Size = new System.Drawing.Size(76, 16);
			this.label5.TabIndex = 304;
			this.label5.Text = "Observaciones";
			// 
			// chkAutorizar
			// 
			this.chkAutorizar.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			appearance33.ForeColorDisabled = System.Drawing.Color.Black;
			this.chkAutorizar.Appearance = appearance33;
			this.chkAutorizar.BackColor = System.Drawing.Color.Transparent;
			this.chkAutorizar.Enabled = false;
			this.chkAutorizar.Location = new System.Drawing.Point(536, 352);
			this.chkAutorizar.Name = "chkAutorizar";
			this.chkAutorizar.Size = new System.Drawing.Size(88, 22);
			this.chkAutorizar.TabIndex = 305;
			this.chkAutorizar.Text = "Aut. Entrada";
			this.chkAutorizar.Visible = false;
			// 
			// txtBodegaPredef
			// 
			this.txtBodegaPredef.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			appearance34.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(230)), ((System.Byte)(230)), ((System.Byte)(230)));
			appearance34.ForeColor = System.Drawing.Color.FromArgb(((System.Byte)(50)), ((System.Byte)(50)), ((System.Byte)(50)));
			this.txtBodegaPredef.Appearance = appearance34;
			this.txtBodegaPredef.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(230)), ((System.Byte)(230)), ((System.Byte)(230)));
			this.txtBodegaPredef.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtBodegaPredef.Enabled = false;
			this.txtBodegaPredef.Font = new System.Drawing.Font("Tahoma", 8.25F);
			this.txtBodegaPredef.ForeColor = System.Drawing.Color.FromArgb(((System.Byte)(50)), ((System.Byte)(50)), ((System.Byte)(50)));
			this.txtBodegaPredef.Location = new System.Drawing.Point(920, 496);
			this.txtBodegaPredef.Name = "txtBodegaPredef";
			this.txtBodegaPredef.PromptChar = ' ';
			this.txtBodegaPredef.Size = new System.Drawing.Size(16, 22);
			this.txtBodegaPredef.TabIndex = 272;
			this.txtBodegaPredef.Visible = false;
			// 
			// lblIva0
			// 
			this.lblIva0.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			this.lblIva0.AutoSize = true;
			this.lblIva0.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.lblIva0.Location = new System.Drawing.Point(824, 339);
			this.lblIva0.Name = "lblIva0";
			this.lblIva0.Size = new System.Drawing.Size(32, 16);
			this.lblIva0.TabIndex = 441;
			this.lblIva0.Text = "IVA 0";
			this.lblIva0.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// label55
			// 
			this.label55.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			this.label55.AutoSize = true;
			this.label55.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label55.Location = new System.Drawing.Point(928, 339);
			this.label55.Name = "label55";
			this.label55.Size = new System.Drawing.Size(22, 16);
			this.label55.TabIndex = 442;
			this.label55.Text = "IVA";
			this.label55.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// label8
			// 
			this.label8.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			this.label8.AutoSize = true;
			this.label8.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label8.Location = new System.Drawing.Point(736, 363);
			this.label8.Name = "label8";
			this.label8.Size = new System.Drawing.Size(46, 16);
			this.label8.TabIndex = 438;
			this.label8.Text = "Subtotal";
			this.label8.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// label9
			// 
			this.label9.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			this.label9.AutoSize = true;
			this.label9.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label9.Location = new System.Drawing.Point(736, 387);
			this.label9.Name = "label9";
			this.label9.Size = new System.Drawing.Size(58, 16);
			this.label9.TabIndex = 437;
			this.label9.Text = "Descuento";
			this.label9.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// txtDescIva0Total
			// 
			this.txtDescIva0Total.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			appearance35.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtDescIva0Total.Appearance = appearance35;
			this.txtDescIva0Total.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtDescIva0Total.Enabled = false;
			this.txtDescIva0Total.FormatString = "#,##0.00";
			this.txtDescIva0Total.Location = new System.Drawing.Point(824, 384);
			this.txtDescIva0Total.Name = "txtDescIva0Total";
			this.txtDescIva0Total.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtDescIva0Total.PromptChar = ' ';
			this.txtDescIva0Total.Size = new System.Drawing.Size(80, 22);
			this.txtDescIva0Total.TabIndex = 440;
			// 
			// txtIva0
			// 
			this.txtIva0.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			appearance36.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtIva0.Appearance = appearance36;
			this.txtIva0.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtIva0.Enabled = false;
			this.txtIva0.FormatString = "#,##0.00";
			this.txtIva0.Location = new System.Drawing.Point(824, 360);
			this.txtIva0.Name = "txtIva0";
			this.txtIva0.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtIva0.PromptChar = ' ';
			this.txtIva0.Size = new System.Drawing.Size(80, 22);
			this.txtIva0.TabIndex = 439;
			// 
			// label15
			// 
			this.label15.AutoSize = true;
			this.label15.Location = new System.Drawing.Point(408, 11);
			this.label15.Name = "label15";
			this.label15.Size = new System.Drawing.Size(41, 16);
			this.label15.TabIndex = 444;
			this.label15.Text = "Bodega";
			this.label15.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// cmbBodega1
			// 
			this.cmbBodega1.CausesValidation = false;
			this.cmbBodega1.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbBodega1.Cursor = System.Windows.Forms.Cursors.Default;
			ultraGridColumn31.Header.VisiblePosition = 2;
			ultraGridColumn32.Header.VisiblePosition = 1;
			ultraGridColumn32.Width = 180;
			ultraGridColumn33.Header.VisiblePosition = 0;
			ultraGridBand4.Columns.AddRange(new object[] {
																										 ultraGridColumn31,
																										 ultraGridColumn32,
																										 ultraGridColumn33});
			this.cmbBodega1.DisplayLayout.BandsSerializer.Add(ultraGridBand4);
			this.cmbBodega1.DisplayMember = "Nombre";
			this.cmbBodega1.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbBodega1.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbBodega1.Enabled = false;
			this.cmbBodega1.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbBodega1.Location = new System.Drawing.Point(464, 8);
			this.cmbBodega1.MaxDropDownItems = 30;
			this.cmbBodega1.Name = "cmbBodega1";
			this.cmbBodega1.Size = new System.Drawing.Size(180, 22);
			this.cmbBodega1.TabIndex = 443;
			this.cmbBodega1.ValueMember = "Bodega";
			this.cmbBodega1.Validating += new System.ComponentModel.CancelEventHandler(this.cmbBodega1_Validating);
			this.cmbBodega1.Validated += new System.EventHandler(this.cmbBodega1_Validated);
			this.cmbBodega1.ValueChanged += new System.EventHandler(this.cmbBodega1_ValueChanged);
			this.cmbBodega1.InitializeLayout += new Infragistics.Win.UltraWinGrid.InitializeLayoutEventHandler(this.cmbBodega1_InitializeLayout);
			// 
			// btnValidarMail
			// 
			this.btnValidarMail.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			this.btnValidarMail.Enabled = false;
			this.btnValidarMail.Location = new System.Drawing.Point(872, 40);
			this.btnValidarMail.Name = "btnValidarMail";
			this.btnValidarMail.Size = new System.Drawing.Size(136, 22);
			this.btnValidarMail.TabIndex = 445;
			this.btnValidarMail.Text = "Validar Mail Promocional";
			this.btnValidarMail.Click += new System.EventHandler(this.btnValidarMail_Click);
			// 
			// txtProteccionDD
			// 
			appearance37.FontData.SizeInPoints = 14F;
			appearance37.ForeColor = System.Drawing.Color.Firebrick;
			appearance37.ForeColorDisabled = System.Drawing.Color.Firebrick;
			this.txtProteccionDD.Appearance = appearance37;
			this.txtProteccionDD.BorderStyle = Infragistics.Win.UIElementBorderStyle.Solid;
			this.txtProteccionDD.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtProteccionDD.Enabled = false;
			this.txtProteccionDD.FormatString = "#,##0.00";
			this.txtProteccionDD.Location = new System.Drawing.Point(872, 64);
			this.txtProteccionDD.Name = "txtProteccionDD";
			this.txtProteccionDD.NullText = "0.00";
			this.txtProteccionDD.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtProteccionDD.PromptChar = ' ';
			this.txtProteccionDD.Size = new System.Drawing.Size(8, 30);
			this.txtProteccionDD.TabIndex = 447;
			this.txtProteccionDD.Visible = false;
			// 
			// lblClienteReiterativo
			// 
			this.lblClienteReiterativo.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			this.lblClienteReiterativo.AutoSize = true;
			this.lblClienteReiterativo.Font = new System.Drawing.Font("Tahoma", 11F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.lblClienteReiterativo.ForeColor = System.Drawing.Color.Firebrick;
			this.lblClienteReiterativo.Location = new System.Drawing.Point(416, 104);
			this.lblClienteReiterativo.Name = "lblClienteReiterativo";
			this.lblClienteReiterativo.Size = new System.Drawing.Size(0, 21);
			this.lblClienteReiterativo.TabIndex = 448;
			this.lblClienteReiterativo.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
			// 
			// dtFecha
			// 
			appearance38.ForeColorDisabled = System.Drawing.Color.Black;
			this.dtFecha.Appearance = appearance38;
			this.dtFecha.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			dateButton1.Caption = "Today";
			this.dtFecha.DateButtons.Add(dateButton1);
			this.dtFecha.Enabled = false;
			this.dtFecha.Format = "dd/MM/yyyy";
			this.dtFecha.Location = new System.Drawing.Point(56, 8);
			this.dtFecha.Name = "dtFecha";
			this.dtFecha.NonAutoSizeHeight = 23;
			this.dtFecha.Size = new System.Drawing.Size(112, 21);
			this.dtFecha.SpinButtonsVisible = true;
			this.dtFecha.TabIndex = 449;
			this.dtFecha.Value = ((object)(resources.GetObject("dtFecha.Value")));
			this.dtFecha.KeyDown += new System.Windows.Forms.KeyEventHandler(this.dtFecha_KeyDown_1);
			this.dtFecha.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.dtFecha_KeyPress_1);
			this.dtFecha.Validating += new System.ComponentModel.CancelEventHandler(this.dtFecha_Validating_1);
			this.dtFecha.Validated += new System.EventHandler(this.dtFecha_Validated_1);
			this.dtFecha.BeforeDropDown += new System.ComponentModel.CancelEventHandler(this.dtFecha_BeforeDropDown_1);
			this.dtFecha.ValueChanged += new System.EventHandler(this.dtFecha_ValueChanged_1);
			// 
			// lblClienteEntidad
			// 
			this.lblClienteEntidad.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			this.lblClienteEntidad.AutoSize = true;
			this.lblClienteEntidad.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.lblClienteEntidad.ForeColor = System.Drawing.Color.Firebrick;
			this.lblClienteEntidad.Location = new System.Drawing.Point(624, 440);
			this.lblClienteEntidad.Name = "lblClienteEntidad";
			this.lblClienteEntidad.Size = new System.Drawing.Size(0, 22);
			this.lblClienteEntidad.TabIndex = 179;
			this.lblClienteEntidad.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
			this.lblClienteEntidad.Visible = false;
			// 
			// chkAnclar
			// 
			this.chkAnclar.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			appearance39.ForeColorDisabled = System.Drawing.Color.Black;
			this.chkAnclar.Appearance = appearance39;
			this.chkAnclar.BackColor = System.Drawing.Color.Transparent;
			this.chkAnclar.Enabled = false;
			this.chkAnclar.Location = new System.Drawing.Point(480, 384);
			this.chkAnclar.Name = "chkAnclar";
			this.chkAnclar.Size = new System.Drawing.Size(88, 22);
			this.chkAnclar.TabIndex = 450;
			this.chkAnclar.Text = "Colateral";
			this.chkAnclar.Visible = false;
			this.chkAnclar.CheckedChanged += new System.EventHandler(this.chkAnclar_CheckedChanged);
			// 
			// txtSerieNuovo
			// 
			appearance40.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtSerieNuovo.Appearance = appearance40;
			this.txtSerieNuovo.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtSerieNuovo.Enabled = false;
			this.txtSerieNuovo.Location = new System.Drawing.Point(480, 440);
			this.txtSerieNuovo.MaxLength = 10;
			this.txtSerieNuovo.Name = "txtSerieNuovo";
			this.txtSerieNuovo.Size = new System.Drawing.Size(104, 22);
			this.txtSerieNuovo.TabIndex = 451;
			this.txtSerieNuovo.Visible = false;
			this.txtSerieNuovo.Validated += new System.EventHandler(this.txtSerieNuovo_Validated);
			// 
			// label7
			// 
			this.label7.AutoSize = true;
			this.label7.BackColor = System.Drawing.Color.Transparent;
			this.label7.Location = new System.Drawing.Point(480, 416);
			this.label7.Name = "label7";
			this.label7.Size = new System.Drawing.Size(74, 16);
			this.label7.TabIndex = 452;
			this.label7.Text = "Codigo Nuovo";
			this.label7.Visible = false;
			// 
			// lblColateral
			// 
			this.lblColateral.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.lblColateral.AutoSize = true;
			this.lblColateral.BackColor = System.Drawing.Color.Transparent;
			this.lblColateral.Font = new System.Drawing.Font("Tahoma", 7F, ((System.Drawing.FontStyle)((System.Drawing.FontStyle.Bold | System.Drawing.FontStyle.Underline))));
			this.lblColateral.Location = new System.Drawing.Point(544, 328);
			this.lblColateral.Name = "lblColateral";
			this.lblColateral.Size = new System.Drawing.Size(0, 15);
			this.lblColateral.TabIndex = 453;
			this.lblColateral.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
			// 
			// txtSerieNuovo2
			// 
			appearance41.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtSerieNuovo2.Appearance = appearance41;
			this.txtSerieNuovo2.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtSerieNuovo2.Enabled = false;
			this.txtSerieNuovo2.Location = new System.Drawing.Point(592, 440);
			this.txtSerieNuovo2.MaxLength = 10;
			this.txtSerieNuovo2.Name = "txtSerieNuovo2";
			this.txtSerieNuovo2.Size = new System.Drawing.Size(104, 22);
			this.txtSerieNuovo2.TabIndex = 454;
			this.txtSerieNuovo2.Visible = false;
			this.txtSerieNuovo2.Validated += new System.EventHandler(this.txtSerieNuovo2_Validated);
			// 
			// label10
			// 
			this.label10.AutoSize = true;
			this.label10.BackColor = System.Drawing.Color.Transparent;
			this.label10.Location = new System.Drawing.Point(592, 416);
			this.label10.Name = "label10";
			this.label10.Size = new System.Drawing.Size(83, 16);
			this.label10.TabIndex = 455;
			this.label10.Text = "Codigo Nuovo 2";
			this.label10.Visible = false;
			// 
			// cdsSeteo
			// 
			this.cdsSeteo.BindingContextCtrl = this;
			this.cdsSeteo.DataLibrary = "LibContabilidad";
			this.cdsSeteo.DataLibraryUrl = "";
			this.cdsSeteo.DataSetDef = "dsSeteo";
			this.cdsSeteo.Locale = new System.Globalization.CultureInfo("es-EC");
			this.cdsSeteo.Name = "cdsSeteo";
			this.cdsSeteo.SchemaClassName = "LibContabilidad.DataClass";
			this.cdsSeteo.SchemaDef = null;
			this.cdsSeteo.BeforeFill += new C1.Data.FillEventHandler(this.cdsSeteo_BeforeFill);
			// 
			// chkFreelance
			// 
			this.chkFreelance.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			appearance42.ForeColorDisabled = System.Drawing.Color.Black;
			this.chkFreelance.Appearance = appearance42;
			this.chkFreelance.BackColor = System.Drawing.Color.Transparent;
			this.chkFreelance.Enabled = false;
			this.chkFreelance.Location = new System.Drawing.Point(568, 384);
			this.chkFreelance.Name = "chkFreelance";
			this.chkFreelance.Size = new System.Drawing.Size(88, 22);
			this.chkFreelance.TabIndex = 456;
			this.chkFreelance.Text = "Freelance";
			this.chkFreelance.Visible = false;
			// 
			// btnCrediPoint
			// 
			this.btnCrediPoint.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btnCrediPoint.CausesValidation = false;
			this.btnCrediPoint.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnCrediPoint.ImageIndex = 67;
			this.btnCrediPoint.Location = new System.Drawing.Point(744, 488);
			this.btnCrediPoint.Name = "btnCrediPoint";
			this.btnCrediPoint.Size = new System.Drawing.Size(80, 23);
			this.btnCrediPoint.TabIndex = 457;
			this.btnCrediPoint.Text = "Credi-POINT";
			this.btnCrediPoint.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnCrediPoint.Visible = false;
			this.btnCrediPoint.Click += new System.EventHandler(this.btnCrediPoint_Click);
			// 
			// lblRecurrente
			// 
			this.lblRecurrente.AutoSize = true;
			this.lblRecurrente.BackColor = System.Drawing.Color.Transparent;
			this.lblRecurrente.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
			this.lblRecurrente.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Bold);
			this.lblRecurrente.ForeColor = System.Drawing.Color.Firebrick;
			this.lblRecurrente.Location = new System.Drawing.Point(867, 70);
			this.lblRecurrente.Name = "lblRecurrente";
			this.lblRecurrente.Size = new System.Drawing.Size(0, 20);
			this.lblRecurrente.TabIndex = 714;
			this.lblRecurrente.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
			// 
			// frmPrefactura
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 14);
			this.CausesValidation = false;
			this.ClientSize = new System.Drawing.Size(1014, 518);
			this.Controls.Add(this.lblRecurrente);
			this.Controls.Add(this.btnCrediPoint);
			this.Controls.Add(this.chkFreelance);
			this.Controls.Add(this.label10);
			this.Controls.Add(this.label7);
			this.Controls.Add(this.label15);
			this.Controls.Add(this.lblIva0);
			this.Controls.Add(this.label55);
			this.Controls.Add(this.label8);
			this.Controls.Add(this.label9);
			this.Controls.Add(this.label5);
			this.Controls.Add(this.label4);
			this.Controls.Add(this.lblComprobante);
			this.Controls.Add(this.label2);
			this.Controls.Add(this.label1);
			this.Controls.Add(this.lblIva);
			this.Controls.Add(this.label63);
			this.Controls.Add(this.label64);
			this.Controls.Add(this.label65);
			this.Controls.Add(this.label23);
			this.Controls.Add(this.lblIdentificacion);
			this.Controls.Add(this.lblCliente);
			this.Controls.Add(this.lblBodega);
			this.Controls.Add(this.lblVendedor);
			this.Controls.Add(this.label6);
			this.Controls.Add(this.label62);
			this.Controls.Add(this.label21);
			this.Controls.Add(this.label61);
			this.Controls.Add(this.txtSerieNuovo2);
			this.Controls.Add(this.lblColateral);
			this.Controls.Add(this.txtSerieNuovo);
			this.Controls.Add(this.chkAnclar);
			this.Controls.Add(this.lblClienteEntidad);
			this.Controls.Add(this.dtFecha);
			this.Controls.Add(this.lblClienteReiterativo);
			this.Controls.Add(this.txtProteccionDD);
			this.Controls.Add(this.btnValidarMail);
			this.Controls.Add(this.cmbBodega1);
			this.Controls.Add(this.txtDescIva0Total);
			this.Controls.Add(this.txtIva0);
			this.Controls.Add(this.chkAutorizar);
			this.Controls.Add(this.txtObservaciones);
			this.Controls.Add(this.chkVerificar);
			this.Controls.Add(this.btnClaveDescuento);
			this.Controls.Add(this.txtidEntidadFinanciera);
			this.Controls.Add(this.txtInteres);
			this.Controls.Add(this.btnAprobacion);
			this.Controls.Add(this.txtTotalCuotas);
			this.Controls.Add(this.txtComprobante);
			this.Controls.Add(this.txtidTipoRuc);
			this.Controls.Add(this.txtIdFormaPago);
			this.Controls.Add(this.gbBotones);
			this.Controls.Add(this.label3);
			this.Controls.Add(this.btnNuevo);
			this.Controls.Add(this.btnConsultar);
			this.Controls.Add(this.btnEditar);
			this.Controls.Add(this.btnImprimir);
			this.Controls.Add(this.btnGuardar);
			this.Controls.Add(this.btnAnular);
			this.Controls.Add(this.btnCancelar);
			this.Controls.Add(this.btnSalir);
			this.Controls.Add(this.txtNumIdProyecto);
			this.Controls.Add(this.txtIdComprobante);
			this.Controls.Add(this.txtNumEstado);
			this.Controls.Add(this.txtIdTipoFactura);
			this.Controls.Add(this.txtBodegaPredef);
			this.Controls.Add(this.txtIdSucursal);
			this.Controls.Add(this.txtNumContadoCredito);
			this.Controls.Add(this.txtNumIdCompra);
			this.Controls.Add(this.txtNumIdCliente);
			this.Controls.Add(this.txtNumero);
			this.Controls.Add(this.txtTotal);
			this.Controls.Add(this.txtIvaTotal);
			this.Controls.Add(this.txtDescIvaTotal);
			this.Controls.Add(this.txtIva);
			this.Controls.Add(this.lblUsuario);
			this.Controls.Add(this.lblEstado);
			this.Controls.Add(this.txtCuotaAsignada);
			this.Controls.Add(this.txtCuotaDisponible);
			this.Controls.Add(this.txtCuotaUsada);
			this.Controls.Add(this.txtValorCuotas);
			this.Controls.Add(this.txtNotas);
			this.Controls.Add(this.ultraGrid2);
			this.Controls.Add(this.txtRuc);
			this.Controls.Add(this.cmbBodega);
			this.Controls.Add(this.txtNombre);
			this.Controls.Add(this.cmbVendedor);
			this.Controls.Add(this.cboFechaPrimerPago);
			this.Controls.Add(this.txtNumCuotas);
			this.Controls.Add(this.txtCuotaDeEntrada);
			this.KeyPreview = true;
			this.Name = "frmPrefactura";
			this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
			this.Text = "Prefactura";
			this.KeyDown += new System.Windows.Forms.KeyEventHandler(this.frmPrefactura_KeyDown);
			this.Resize += new System.EventHandler(this.frmPrefactura_Resize);
			this.Closing += new System.ComponentModel.CancelEventHandler(this.frmPrefactura_Closing);
			this.Load += new System.EventHandler(this.frmPrefactura_Load);
			((System.ComponentModel.ISupportInitialize)(this.txtNumCuotas)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtCuotaDeEntrada)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbVendedor)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtRuc)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbBodega)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNombre)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraGrid2)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource2)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNotas)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtCuotaAsignada)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtCuotaDisponible)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtCuotaUsada)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtValorCuotas)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtTotal)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIvaTotal)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtDescIvaTotal)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIva)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNumero)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNumIdProyecto)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIdComprobante)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNumEstado)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIdTipoFactura)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIdSucursal)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNumContadoCredito)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNumIdCompra)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNumIdCliente)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIdFormaPago)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtidTipoRuc)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtComprobante)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtTotalCuotas)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtInteres)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource1)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtidEntidadFinanciera)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtObservaciones)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtBodegaPredef)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtDescIva0Total)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIva0)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbBodega1)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtProteccionDD)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.dtFecha)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtSerieNuovo)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtSerieNuovo2)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteo)).EndInit();
			this.ResumeLayout(false);

		}
		#endregion

		#region Variables Globales
		decimal dPorcEnt = 0.00m; 
		decimal dPorcEntMin = 0.00m;
		decimal dEntrada = 0.00m;
		decimal dEntMin = 0.00m;
		int idPromocionCombo = 0;
		int IdArtProm = 0;
		bool bEliminaAlValidar = false;
		private Acceso miAcceso;
		private C1DataRow drSeteoF;
		private C1DataRow drSeteoC;
		private C1DataRow drSeteo;
		private int IdSigno = 2;
		private bool bIncluyeIva = false;
		string controlW="";
		string controlW1="";		
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
		decimal dCuotaAsignada = 0.00m;
		decimal dInteres = 0.00m;
		decimal dCuotaDisponible = 0.00m;
		decimal dCuotaUsada = 0.00m;
		bool AutorizaEntrada = false;
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
		int idEntidadFinanciera = 0;
		int idArtGrup = 0;
		int idPlazoD=0;
		bool bCuponWeb = false;
		decimal dDsctoReiterativo = 0.00m;
		decimal dDsctoReiterativoVIP = 0.00m;
		int iColateral = 0;
		int iNoColateral = 0;
		int iEstadoSol= 0;
		bool bNuovoHabil = false;
		bool bEmpleado = false;
		decimal dCuotaEntrada = 0.00m;
		bool bGuarda = false;
    decimal PorcentajeEntrada = 0.00m;
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
			this.txtNumCuotas.MaxValue = iMaximoNumeroCuotas;
			this.txtNumCuotas.Value = iMaximoNumeroCuotas;
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

			if((int)this.txtidEntidadFinanciera.Value !=14)
			{
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
			}

			DiaIni = fechaIni.Date.Day;
			
			this.cboFechaPrimerPago.Items.Add(fechaIni.ToString("dd/MM/yyyy"));

			for (int i = iMinimoDiasPago + 1; i <= iMaximoDiasPago; i++)
			{
				DateTime Fecha = fechaIni.Date.AddDays(1);

				int Dia = Fecha.Date.Day;

				if (i <= iMaximoDiasPago)//  46 || i == iMaximoDiasPago)//60)	//60
					if (Dia <= iDiaMaximoDePago) 
						this.cboFechaPrimerPago.Items.Add(Fecha.ToString("dd/MM/yyyy"));
												
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
			this.txtProteccionDD.Value = 0.00m;

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

			//this.txtProteccionDD.Value = FuncionesInventario.FacturacionValorDYD(this.ultraGrid2, 9, (int)this.txtNumCuotas.Value);

			#region Calculo Cuotas
			decimal dMontoCredito = Math.Round(Convert.ToDecimal(this.txtTotal.Value), 2) - Math.Round(Convert.ToDecimal(this.txtCuotaDeEntrada.Value), 2);

			if (dMontoCredito > 0.00m)
			{
				if ((int)this.txtidEntidadFinanciera.Value != 6)
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
				else
				{
					this.txtValorCuotas.Value = Math.Round(dMontoCredito / (int)this.txtNumCuotas.Value, 2);
				}
			}
			#endregion Calculo Cuotas

			#region Calculo Cuotas Entrada
			dMontoCredito = Math.Round(Convert.ToDecimal(this.txtTotal.Value), 2) - Math.Round(Convert.ToDecimal(this.txtCuotaDeEntrada.Value), 2);
	
			if (dMontoCredito > 0.00m)
			{
				if ((int)this.txtidEntidadFinanciera.Value != 6)
				{		
					#region CuotaEntrada
					//
					////					 this.txtCuotaDeEntrada.Value = 0.00m;
					////					 dCuotaEntrada = 0.00m;
					//
					////					string sSQLCe= string.Format("Exec CalculaCuotaInicial {0}, {1}, {2}, '{3}', {4}", 
					////						dMontoCredito, (int)this.txtNumCuotas.Value, Convert.ToInt32(this.txtNumIdCliente.Value), Convert.ToDateTime(this.dtFecha.Value).ToString("yyyyMMdd"), Convert.ToInt32(this.txtidEntidadFinanciera.Value));
					////					SqlDataReader dReaderC =	Funcion.rEscalarSQL(cdsSeteoF, sSQLCe, true);
					////
					////					dReaderC.Read();
					//////					this.txtCuotaDeEntrada.Value = Convert.ToDecimal(dReaderC.GetValue(2));
					////					dCuotaEntrada  = Convert.ToDecimal(dReaderC.GetValue(2));
					////					dReaderC.Close();	
					//   this.txtCuotaDeEntrada.Value = dCuotaEntrada; 	         
					//dMontoCredito = Math.Round(Convert.ToDecimal(this.txtTotal.Value), 2) - Math.Round(Convert.ToDecimal(this.txtCuotaDeEntrada.Value), 2);      
					#endregion CuotaEntrada 
						
					string sSQL1 = string.Format("Exec CalculaCuotaInicial {0}, {1}, {2}, '{3}', {4}", 
						dMontoCredito, (int)this.txtNumCuotas.Value, Convert.ToInt32(this.txtNumIdCliente.Value), Convert.ToDateTime(this.dtFecha.Value).ToString("yyyyMMdd"), Convert.ToInt32(this.txtidEntidadFinanciera.Value));
					SqlDataReader dReader =	Funcion.rEscalarSQL(cdsSeteoF, sSQL1, true);
	
					dReader.Read();
					decimal dValorc = Convert.ToDecimal(dReader.GetValue(2));
					this.txtValorCuotas.Value = Convert.ToDecimal(dReader.GetValue(2));
					dReader.Close();				
				
					this.txtTotalCuotas.Value = Math.Round(Convert.ToDecimal(this.txtValorCuotas.Value), 2) * (int)this.txtNumCuotas.Value;			
				}
				else
				{
					this.txtValorCuotas.Value = Math.Round(dMontoCredito / (int)this.txtNumCuotas.Value, 2);
				}
			}
			#endregion Calculo Cuotas Entrada 
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
				//				if (this.miAcceso.CambiarDescuento) this.btnClaveDescuento.Visible = true;			
								
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

			//Inactivar consulta
			PrefacturaControl();
		}

		private void PrefacturaControl()
		{
			if(getIdCompra > 0)
			{
				this.Consulta(getIdCompra);
			}
			if(getbInactivo)
			{
				this.btnNuevo.Enabled = false;
				this.btnConsultar.Enabled = false;
				this.btnEditar.Enabled = false;
				this.btnCancelar.Enabled = false;
				this.btnAprobacion.Enabled = false;
				this.btnAnular.Enabled = false;
				this.btnClaveDescuento.Enabled = false;
				this.btnGuardar.Enabled = false;
			}
		}
 
		#region Ultragrid
		private void ultraGrid1_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (!bEdicion) return;

			if (e.KeyCode == Keys.F3)
			{
				if (this.ultraGrid2.ActiveCell.Column.Key.ToString().ToUpper() == "CODIGO")
				{
					if (!Validacion.bValidaAbreListaArticulos(this.ultraGrid2)) return;

					FuncionesInventario.AbreBusquedaArticulos(this.ultraGrid2, (int)this.cmbBodega1.Value, 43, 9, Convert.ToDateTime(this.dtFecha.Value), false);
				}
			}

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

		private void ultraGrid1_Error(object sender, Infragistics.Win.UltraWinGrid.ErrorEventArgs e)
		{
			//			if (e.MaskErrorInfo == null)
			//			{
			//				e.Cancel = true;
			//				string stError = e.DataErrorInfo.InvalidValue.ToString();
			//				string stCol = e.DataErrorInfo.Cell.Column.ToString();
			//				string stMensaje = "Error Grilla: PREFACTURA";			
			//
			//				string stAudita = "Exec AuditaCrear 53, 7, '" + stMensaje + "'"; // Se graba como salida de pantalla
			//				Funcion.EjecutaSQL(cdsSeteoF, stAudita, true);
			//
			//				MessageBox.Show("Valor: " + stError + " no valido en columna: " + stCol + "\n1.- GRABE LOS DATOS. \n2.- SALGA DE LA PANTALLA.\n3.- VUELVA A INGRESAR", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Error);
			//			}
		}

		bool bEliminando = false;
		private void ultraGrid1_BeforeRowsDeleted(object sender, Infragistics.Win.UltraWinGrid.BeforeRowsDeletedEventArgs e)
		{
			if (bEdicion)
			{			
				try
				{
					if (!bEliminaAlValidar)
					{
						if (!bEliminando) 
						{
							bEliminando = true;				
							if (!Facturacion.FacturacionEliminaFilas(this.ultraGrid2, this.ultraGrid2, e, this.btnCancelar, 43, bNuevo, bImpreso, miAcceso.BEliArtCmbProm, miAcceso.BEliArt, (int)this.cmbBodega.Value, idTipoPrecio, 9, (decimal)this.txtTotal.Value,
								0, 0, 0, 0, (int)this.txtNumCuotas.Value, bCuponWeb, false, cdsSeteoF)) e.Cancel = true;
							e.DisplayPromptMsg = false;
							bEliminando = false;						
						}
						else e.DisplayPromptMsg = false;
					}
					else e.DisplayPromptMsg = false;
				}
				catch (Exception ex2)
				{
					MessageBox.Show(string.Format("Error al eliminar filas\n\nSi la pre factura ya fue guardada se anulara automaticamente\n\nLa Pantalla se cerrara vuelva a abrirla para continuar", ex2.GetType(), ex2.Message), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);													

					if ((int)this.txtNumIdCompra.Value > 0) 
						Funcion.EjecutaSQL(cdsSeteoF, string.Format("Exec AnulacionCompraVenta {0}, 43, 'ANULACION AUTOMATICA POR ERROR EN ELIMINACION DE PRODUCTOS'", (int)this.txtNumIdCompra.Value));

					this.btnCancelar_Click(sender, e);
					this.Close();
					e.DisplayPromptMsg = false;
				}
			}
		}

		private void ultraGrid1_AfterRowInsert(object sender, Infragistics.Win.UltraWinGrid.RowEventArgs e)
		{
			if (bEdicion)Facturacion.IniciaValoresColumnas(sender, e, this.ultraGrid2, 43, (int)this.cmbBodega.Value);
			//			foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.ultraGrid2.Rows.All)	
			//			{
			//				if (dr.Cells["idArticulo"].Value != System.DBNull.Value)
			//				{
			//					if ((int)dr.Cells["idArticulo"].Value >0)
			//					{
			//						idArtGrup = Funcion.iEscalarSQL(cdsSeteoF, string.Format("select idGrupoArticulo from Articulo  where idArticulo= {0}",(int)dr.Cells["idArticulo"].Value));
			//						if (idEntidadFinanciera== 14)
			//						{
			//							if (idArtGrup !=8 || idArtGrup !=57)
			//							{
			//								MessageBox.Show("Para celulares la entidad  financiera  debe ser  DATACULTR", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Information);
			//								//btnCancelar_Click(sender, e);
			//								Validacion.vEliminaFilasVaciasF(this.ultraGrid2, "idArticulo",(int)dr.Cells["idArticulo"].Value);
			//								//return;
			//							}
			//						}
			//						
			//					}
			//				}
			//			}
		}	
		
		private void ultraGrid1_AfterCellUpdate(object sender, Infragistics.Win.UltraWinGrid.CellEventArgs e)
		{		
			if (bEdicion)
				if (bFundas)return;
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
							if((int)this.txtidEntidadFinanciera.Value ==14)
							{
								idPlazoD = 0;
								if ((int)this.txtNumCuotas.Value == 13 && !estadoFF.EsFF){idPlazoD=1; idTipoPrecio=5; this.txtNumCuotas.Enabled=false;}
								if ((int)this.txtNumCuotas.Value == 26 && !estadoFF.EsFF){idPlazoD=1; idTipoPrecio=6; this.txtNumCuotas.Enabled=false;}
								if ((int)this.txtNumCuotas.Value == 39 && !estadoFF.EsFF){idPlazoD=1; idTipoPrecio=7; this.txtNumCuotas.Enabled=false;}
								if ((int)this.txtNumCuotas.Value == 52 && !estadoFF.EsFF){idPlazoD=1; idTipoPrecio=8; this.txtNumCuotas.Enabled=false;}
								if(idPlazoD == 0 )
								{
									MessageBox.Show("Para venta de equipos Point Movil el numero de pagos debe ser  13,26,39,52  semanas ", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
									this.txtNumCuotas.Focus();
									return;
								}
							}

						FuncionesInventario.FacturacionCargaArticulos(sender, e, this.ultraGrid2, 9, idTipoPrecio, (int)this.cmbBodega.Value, (int)this.cmbBodega1.Value, 0, 0, 0, 0, Convert.ToDateTime(this.dtFecha.Value), 43, true, true, true, false, (int)this.txtNumCuotas.Value, (int)this.txtidEntidadFinanciera.Value, (int)this.txtNumIdCliente.Value, 0, 0, bCuponWeb, false, 0, cdsSeteoF, false,(int)this.txtNumCuotas.Value, false);
						this.txtNumCuotas.Enabled = false;						
						if (this.ultraGrid2.Rows.Count > 0) this.ultraGrid2.Rows.Band.AddNew();

						bActualiza = true;
						this.cmdTotal();

						foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.ultraGrid2.Rows.All)
							if ((int)dr.Cells["idCombo"].Value > 0 && ((int)dr.Cells["Tipo"].Value == 6))
								if (Funcion.bEscalarSQL(cdsSeteoF, string.Format("Exec PromocionRetornaSiAplicaCuotas {0}", (int)dr.Cells["idCombo"].Value), true)) 
									this.txtNumCuotas.Enabled = false;
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
						#region Cliente VIP
						if(bVIP)
						{
							dDsctoReiterativo = dDsctoReiterativoVIP;
							cmdTotal();
							bActualiza = true;
							return;
						}
						#endregion Cliente VIP

						if (!Facturacion.bValidaDescuentos(e, this.ultraGrid2, (int)this.cmbBodega.Value, bDescuentoActivo, dClavePorcDscto, 43, miAcceso.CambiarDescuento, idTipoPrecio, 0, 0, 0, 0, dPorcentajeCuponWeb, dDsctoReiterativo, bCuponWeb, idMail, bEmpleado, cdsSeteoF))
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
					
					if (e.Cell.Row.Cells["idArticulo"].Value != System.DBNull.Value)
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

		bool bTieneBonoEntrada = false;
		private void CargaCliente()
		{
			string[] sParametros = new string[11];

			#region CrediDigital
			if (Funcion.iEscalarSQL(cdsSeteoF, string.Format("Exec ValidaBodegaCrediPointdigital {0}", (int) this.cmbBodega.Value)) == 0)
			{
				sParametros = FuncionesProcedimientos.RetornaParametrosClientesProveedores(false, 43, this.txtRuc.Text.ToString(), cdsSeteoF);
			}
			else
			{
				sParametros = FuncionesProcedimientos.RetornaParametrosClientesProveedoresCrediDigital(false, 43, this.txtRuc.Text.ToString(), cdsSeteoF, 1);
			}
			#endregion CrediDigital

			if (sParametros[0] != null) this.txtNumIdCliente.Value = int.Parse(sParametros[0]);
			if (sParametros[1] != null) this.txtidTipoRuc.Value = int.Parse(sParametros[1]);
			if (sParametros[3] != null) this.txtRuc.Text = sParametros[3];
			if (sParametros[4] != null) this.txtNombre.Text = sParametros[4];	
			if (sParametros[11] != null) idInstitucion = int.Parse(sParametros[11]);

			int idBonoEntradaTipo = 0;
			//Creo que debo cargar la tabla de Bono entrada tipo e ir realizando un for por cada bono y si se encuentra en vigencia
			DataView dvResultado = Funcion.dvProcedimiento(cdsSeteoF, "Exec BonoEntradaTipoCon");
			if (dvResultado != null)
			{
				foreach (DataRowView rowView in dvResultado)
				{
					DataRow row = rowView.Row;
					idBonoEntradaTipo = Convert.ToInt32(row["idBonoEntradaTipo"]);
					//Validar que tiene bono de entrada, y que se encuentra activo, y si es as? no permita entrar a esta funci?n
					string squery = String.Format("Exec BonoEntradaTieneBono {0}, {1}", (int)this.txtNumIdCliente.Value, idBonoEntradaTipo);
					bTieneBonoEntrada = Funcion.bEjecutaSQL(cdsSeteo, squery);
				}
			}

			#region Cliente VIP Bloqueado
//			if(!bTieneBonoEntrada)
//			{			
//				//PREGUNTAR PORQUE O PARA QUE E.C.	
//				int iExistePromocion = Funcion.iEscalarSQL(cdsSeteoF, string.Format("EXEC ConsultaPromocionVIP '{0}' ", (int)this.cmbBodega.Value));
//
//				decimal dReiterativo = Funcion.decEscalarSQL(cdsSeteoF, string.Format("Select dbo.ValidaDescuentoGestionVentas({0})", (int)this.txtNumIdCliente.Value));
//				
//				dDsctoReiterativo = Funcion.decEscalarSQL(cdsSeteoF, string.Format("Select dbo.ValidaDescuentoGestionVentasVIP('{0}')", this.txtRuc.Text.ToString()));//5  
//				
//				//Funcion.decEscalarSQL(cdsSeteoF, string.Format("Exec Cre_ValidaPromocionReiterativos {0}", (int)this.txtNumIdCliente.Value));//
//				
//				if(dReiterativo > dDsctoReiterativo) 
//				{
//					this.lblClienteReiterativo.Text = string.Format("CLIENTE REITERATIVO TIENE DESCUENTO DEL {0} %", dReiterativo);
//					dDsctoReiterativo = dReiterativo;
//				}
//				else if (dDsctoReiterativo > 0)
//				{
//					//						if( iExistePromocion == 0)
//					//						{
//					/* si es vip y no existe una promocion toma el valor d ela tabla*/
//
//					this.lblClienteReiterativo.Text = string.Format("CLIENTE VIP TIENE UN DESCUENTO DEL {0} %", );
//					dDsctoReiterativoVIP = dDsctoReiterativo;
//					//dDsctoReiterativo = 0;
//					//						}
//					//						else
//					//						{
//					//							/*si existe una promocion toma solo el 3 */
//					//							this.lblClienteReiterativo.Text = string.Format("CLIENTE TIENE DESCUENTO DE {0} %", dReiterativo);
//					//							dDsctoReiterativo= dReiterativo;
//					//						}
//				}
//			}
			#endregion Cliente VIP Bloqueado

			if ((int)this.txtNumIdCliente.Value > 0) this.txtRuc.Enabled = false;	
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
			estadoFF.EsFF = false;
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

				bool ControlCrediDigital = true;
				#region CrediDigital
				if (Funcion.iEscalarSQL(cdsSeteoF, string.Format("Exec ValidaBodegaCrediPointdigital {0}", (int) this.cmbBodega.Value)) == 0)
				{
					if (Funcion.iEscalarSQL(cdsSeteoF, string.Format("Exec RetornaIdEFSolCredito '{0}'", this.txtRuc.Value.ToString())) == 0)
					{
						MessageBox.Show("Cliente le falta un Parametro\n\nComuniquese con Credito", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Information);
						this.txtRuc.Focus();
						ControlCrediDigital = false;
						e.Cancel = true;
					}
				}
				else
				{
					if (Funcion.iEscalarSQL(cdsSeteoF, string.Format("Exec RetornaIdEFSolCreditoDigital '{0}',{1}", this.txtRuc.Value.ToString(), (int) this.cmbBodega.Value)) == 0)
					{
						MessageBox.Show("Cliente CREDIPOINT-DIGITAL le falta un Parametro\n\nComuniquese con Credito", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Information);
						this.txtRuc.Focus();
						ControlCrediDigital = false;
						e.Cancel = true;
					}
				}
				#endregion CrediDigital

				#region FamilyFriends
				if(ControlCrediDigital)
				{
					VerificarFamilyFriends();
				}
				#endregion FamilyFriends

			}
		}

		// Request que recuerda la última cédula y forma de pago confirmadas
		private FamilyFriendsRequestModel reqFF = new FamilyFriendsRequestModel();
		// Solo una variable para todo el estado FF
		private FamilyFriendsEstadoModel estadoFF = new FamilyFriendsEstadoModel();

		private void VerificarFamilyFriends()
		{
			FamilyFriendsService service =
				new FamilyFriendsService(frmRequerimiento.sconexionPoint);

			// Reutilizamos el mismo request global
			reqFF.ValorFormaPago = 9;
			reqFF.RucCliente = this.txtRuc.Text;
			reqFF.Owner = this;

			// Ejecuta la verificación
			FamilyFriendsVerificacionModel r = service.VerificarFF(reqFF);

			// Guardar en el modelo
			estadoFF.EsFF = r.EsFF;
			estadoFF.IdCliente = r.IdCliente;
			estadoFF.IdWEB_FF_Invitado = r.IdWEB_FF_Invitado;
			estadoFF.RechazoPromocion = r.RechazoPromocion;

			// ==============================
			// Manejo de estados FF
			// ==============================
			if (estadoFF.EsFF)
			{
				idTipoPrecio = r.IdTipoPrecio;
				this.lblRecurrente.Text = "CLIENTE FAMILY & FRIENDS";
			}
			else
			{
				estadoFF.IdCliente = 0;

				if (estadoFF.RechazoPromocion)
				{
					this.lblRecurrente.Text = "PROMOCIÓN FF RECHAZADA";
				}
				else
				{
					this.lblRecurrente.Text = "";
				}
			}

			// Mostrar mensaje del SP si existe
			if (r.MensajeSP != null && r.MensajeSP.Length > 0)
			{
				MessageBox.Show(
					r.MensajeSP,
					"Family & Friends",
					MessageBoxButtons.OK,
					MessageBoxIcon.Information
					);
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
				this.lblClienteReiterativo.Text = "";
				this.lblColateral.Text ="";
				dDsctoReiterativo = 0.00m;
				dDsctoReiterativoVIP = 0.00m;
				iColateral = 0;
				iNoColateral = 0;

				if (!Validacion.vbComboVacio(this.cmbBodega, "Seleccione el Local"))
				{
					this.txtRuc.Text ="";
					this.cmbBodega.Focus();		
					return;
				}

				CargaCliente();

				#region Parametros de Credito
				if ((int)this.txtNumIdCliente.Value > 0)
				{
					DateTime dtFecha = Convert.ToDateTime(this.dtFecha.Value);
					string sSQLParametros ="";
					
					#region CrediDigital
					if (Funcion.iEscalarSQL(cdsSeteoF, string.Format("Exec ValidaBodegaCrediPointdigital {0}", (int) this.cmbBodega.Value)) == 0)
					{					
						sSQLParametros = string.Format("Exec Cre_RetornaParametros {0}, '{1}'", (int)this.txtNumIdCliente.Value, dtFecha.ToString("yyyyMMdd"));
					}
					else
					{
						sSQLParametros = string.Format("Exec Cre_RetornaParametrosCreditoDigital {0}, '{1}', {2}", (int)this.txtNumIdCliente.Value, dtFecha.ToString("yyyyMMdd"),(int) this.cmbBodega.Value);				
					}
					#endregion CrediDigital

					SqlDataReader dr = Funcion.rEscalarSQL(cdsSeteoF, sSQLParametros, true);
					dr.Read();
					if (dr.HasRows)
					{
						if (bNuevo)
						{
							this.txtidEntidadFinanciera.Value = dr.GetInt32(0);
							idEntidadFinanciera= dr.GetInt32(0);
							if (idEntidadFinanciera == 14)
								this.lblClienteEntidad.Text = "Point Movil";
						}

						this.txtCuotaAsignada.Value = dr.GetDecimal(1);
						this.txtCuotaUsada.Value = dr.GetDecimal(2);
						this.txtCuotaDisponible.Value = dr.GetDecimal(3);
						this.txtInteres.Value = dr.GetDecimal(4);		
						iColateral = dr.GetInt32(11);
						iNoColateral = dr.GetInt32(12);
						iEstadoSol = dr.GetInt32(13);
						PorcentajeEntrada = dr.GetDecimal(14);
					}
					dr.Close();
			
					ValidaNumeroCuotas();

					CargaFecha();

					if (idInstitucion > 0 && !estadoFF.EsFF) idTipoPrecio = 2; 
					else if (idInstitucion == 0) 
					{
						if (idEntidadFinanciera==14 && !estadoFF.EsFF) 
						{
							if ((int)this.txtNumCuotas.Value == 13 && !estadoFF.EsFF)idTipoPrecio=5;
							if ((int)this.txtNumCuotas.Value == 26 && !estadoFF.EsFF)idTipoPrecio=6;
							if ((int)this.txtNumCuotas.Value == 39 && !estadoFF.EsFF)idTipoPrecio=7;
							if ((int)this.txtNumCuotas.Value == 52 && !estadoFF.EsFF)idTipoPrecio=8;
						}
						else if(!estadoFF.EsFF) idTipoPrecio = 4;
					}

					if (!bNuovoHabil)
					{
						this.chkAnclar.Visible = false;
						this.txtSerieNuovo.Visible = false;
					}

					if (bNuovoHabil && iNoColateral == 0 )
					{
						if (iColateral > 0)
						{
							this.chkAnclar.Checked = true;
							this.chkAnclar.Enabled = false; 
							this.lblColateral.Text ="CLIENTE APLICA COLATERAL";
						}
					}
				}
				#endregion Parametros de Credito			

				bEmpleado = Funcion.bEscalarSQL(cdsSeteoF, string.Format("Exec FacturacionValidaEmpleado '{0}'", this.txtRuc.Text.ToString()), true);
				if (bEmpleado) this.lblClienteReiterativo.Text = "Cliente aplica a promocion empleados";

				#region CrediDigital
				int iYear = 0;
				if (Funcion.iEscalarSQL(cdsSeteoF, string.Format("Exec ValidaBodegaCrediPointdigital {0}", (int) this.cmbBodega.Value)) == 0)
				{
					iYear = Funcion.iEscalarSQL(cdsSeteoF, string.Format("Select DATEDIFF(YEAR, FechaNacimiento, '{0}') From Cre_DatosGenerales Where NumeroIdentificacion = '{1}'", 
							 Convert.ToDateTime(this.dtFecha.Value).ToString("yyyyMMdd"), this.txtRuc.Text.ToString()));
				}
				else
				{
					iYear = Funcion.iEscalarSQL(cdsSeteoF, string.Format("Select top 1 DATEDIFF(YEAR, FechaNacimiento, '{0}') From Web_SolicitudGrande Where Cedula = '{1}'", 
							 Convert.ToDateTime(this.dtFecha.Value).ToString("yyyyMMdd"), this.txtRuc.Text.ToString()));
				}
				#endregion CrediDigital

				//				if (iYear < 66)
				//					FuncionesInventario.FacturacionCargaDYD(this.ultraGrid2, Convert.ToDecimal(this.txtProteccionDD.Value), dPorcIva, cdsSeteoF);

				//				if (iYear < 71)
				//				{
				//					decimal dValor = 0.00m;
				//					if (Convert.ToDecimal(this.txtTotal.Value) - Convert.ToDecimal(this.txtCuotaDeEntrada.Value) > 1000.00m) dValor = 0.00m;
				//
				//					FuncionesInventario.FacturacionCargaSeguroDesgravamen(this.ultraGrid2, dValor, dPorcIva, cdsSeteoF);
				//				}

				#region Encuesta
				//				if (bNuevo)
				//				{
				//					if (Funcion.bEscalarSQL(cdsSeteoF, "Select Encuesta From SeteoF", true))
				//					{
				//						using (frmFacturacionEncuesta FE = new frmFacturacionEncuesta())
				//						{
				//							if (DialogResult.OK == FE.ShowDialog())
				//							{
				//								iEncuesta = (int)FE.optEncuesta.Value;
				//								sOtrosEncuesta = FE.txtOtros.Text.ToString();
				//							}
				//							else return;
				//						}
				//					}
				//				}
				#endregion Encuesta
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
				if (!Validacion.vbCampoEnteroVacio(this.txtNumCuotas, "Escriba el n?mero de cuotas", (int)this.txtNumCuotas.MinValue, (int)this.txtNumCuotas.MaxValue)) e.Cancel = true;
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

		private void dtFecha_BeforeDropDown(object sender, System.ComponentModel.CancelEventArgs e)
		{
		
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
			try
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
						if (DialogResult.Yes == MessageBox.Show("Esta seguro de Cancelar la Edici?n", "Point Technology", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2)) 
							this.btnCancelar_Click(sender, e);	

					if (e.KeyCode == Keys.F12) this.btnGuardar_Click(sender, e);
				}
			}
			catch(System.Exception ex)
			{
				MessageBox.Show(ex.Message, "Point Technology");
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
				MessageBox.Show(string.Format("No puede modificar documentos en esta fecha, El per?odo '{0}' esta cerrado,\n\nConsulte al Administrador", Convert.ToDateTime(this.dtFecha.Value).ToString("MMMM").ToUpper()), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);									
				return;
			}

			if ((int) this.txtNumEstado.Value == 4)
			{
				MessageBox.Show("Este documento esta Rechazado", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);				
				return;
			}
			if ((int) this.txtNumEstado.Value == 5)
			{
				MessageBox.Show("Este documento ya fue Aprobado.", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);				
				return;
			}
			if ((int) this.txtNumEstado.Value == 9)
			{
				MessageBox.Show("Este documento ya se Factur?.", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);				
				return;
			}

			#region Validar Que al crear la factura la bodega no sea de remates.
			// Consultar desde la tabla compra
			string squery = String.Format("Select BodegaDestino from Compra where idCompra = {0}", (int)this.txtNumIdCompra.Value);
			int getIdBodega = Funcion.iEscalarSQL(cdsSeteoF, squery,false);
			if(getIdBodega > 0)
			{
				//Consultar si es de remates 
				squery = String.Format("Select Remates from Bodega where Bodega = {0}", getIdBodega);
				bool Remates = Funcion.bEjecutaSQL(cdsSeteoF, squery);
				if(Remates)
				{
					//Coger el mensaje de prefacturas de remates
					squery = String.Format("Select MensajePrefacturaRemate from Seteo");
					string MensajePrefacturaRemate = Funcion.sEscalarSQL(cdsSeteoF, squery, false);
					MessageBox.Show(MensajePrefacturaRemate, "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);				
					return;
				}
			}
			#endregion Validar Que al crear la factura la bodega no sea de remates.

			#region Valida Creditos Anulados
			string sSQLCA = string.Format("Exec Cre_ValidaCreditosAnulados {0}", (int)this.txtNumIdCliente.Value);
			int iAnulados = Funcion.iEscalarSQL(cdsSeteoF, sSQLCA);
			bool bRevisado = false;
			if (iAnulados > 0)
			{
				if (DialogResult.OK == MessageBox.Show(string.Format("El Cliente tiene {0} Pre-Facturas ANULADAS/RECHAZADAS\n\nVerifique que el PAGARE y la TABLA DE AMORTIZACI?N\n\nCORRESPONDAN AL NUEVO CR?DITO", iAnulados), "Point Technology", MessageBoxButtons.OKCancel, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button2))
				{
					bRevisado = true;
				}
				else return;
			}			
			#endregion Valida Creditos Anulados

			#endregion Verificaciones			

			int iCOntAR = 0;

			foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.ultraGrid2.Rows.All)
			{
				if ((int)dr.Cells["idArticulo"].Value == 9104)
				{
					iCOntAR = iCOntAR +1;
				}
			}
				
			using (frmCre_CarpaAlmacen CA = new frmCre_CarpaAlmacen(iCOntAR, (int) this.cmbBodega.Value, (int) this.txtNumIdCompra.Value))
			{
				if (DialogResult.OK == CA.ShowDialog())
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

							#region Crea Factura
							string sSQLTransforma = string.Format("Exec CompraTransformaCopia {0}, 1, {1}, {2}, {3}, {4}, '{5}', {6}",
								(int)this.txtNumIdCompra.Value, (int)this.cmbBodega.Value, (bool)CA.optTipos.Value, bRevisado, false, CA.txtNotas.Text.ToString().Trim(), CA.chkPreventa.Checked);
							oCmdActualiza.CommandText = sSQLTransforma;
							int idCompraDest = (int)oCmdActualiza.ExecuteScalar();						
							#endregion Crea Factura

							if( (bool) this.chkAnclar.Checked)
							{
								string sSQL = string.Format(" update InsertaCodigoNOVO set estado = 2 where SerieNovo = '{0}'",
									(string) this.txtSerieNuovo.Text);						
								oCmdActualiza.CommandText = sSQL;
								oCmdActualiza.ExecuteNonQuery();

								if ( this.txtSerieNuovo2.Text.Length > 3)
								{
									string sSQLa = string.Format(" update InsertaCodigoNOVO set estado = 2 where SerieNovo = '{0}'",
										(string) this.txtSerieNuovo2.Text);						
									oCmdActualiza.CommandText = sSQLa;
									oCmdActualiza.ExecuteNonQuery();
								}
							}
			
							#region Guarda la tabla de amortizaci?n de la factura
							//							if ((int)this.txtidEntidadFinanciera.Value != 6)
							//							{
							//								decimal dValor = Convert.ToDecimal(this.txtTotal.Value) - Convert.ToDecimal(this.txtCuotaDeEntrada.Value);
							//								DateTime Fecha = Convert.ToDateTime(this.cboFechaPrimerPago.Text.ToString());
							//								DateTime dtFecha = (DateTime)this.dtFecha.Value;
							//
							//								string sSQL = string.Format("Exec GuardaTablaDeAmortizacion {0}, {1}, '{2}', '{3}', {4}", 
							//									dValor, (int)this.txtNumCuotas.Value, dtFecha.ToString("yyyyMMdd"), Fecha.ToString("yyyyMMdd"), idCompraDest);						
							//								oCmdActualiza.CommandText = sSQL;
							//								oCmdActualiza.ExecuteNonQuery();
							//							}
							#endregion Guarda la tabla de amortizaci?n de la factura

							
							oTransaction.Commit();	

							int idBonoEntradaTipo = 0;

							//Creo que debo cargar la tabla de Bono entrada tipo e ir realizando un for por cada bono y si se encuentra en vigencia
							DataView dvResultado = Funcion.dvProcedimiento(cdsSeteoF, "Exec BonoEntradaTipoCon");
							if (dvResultado != null)
							{
								foreach (DataRowView rowView in dvResultado)
								{
									bool UsoBono = false;
									DataRow row = rowView.Row;
									idBonoEntradaTipo = Convert.ToInt32(row["idBonoEntradaTipo"]);
									//Actualizar el estado si tiene bono entrada 25
									//Se guarda consulta por el idPrefactura
									#region Bono Entrada 25
									//Consultar si existe un registro que tenga la prefactura guardada
									int iBonoEntradaPrefactura = 0;
									squery = String.Format("Select Count(*) From BonoEntrada Where idPrefactura = {0} And Estado = 1 And idBonoEntradaTipo = {1}", 
										(int)this.txtNumIdCompra.Value,
										idBonoEntradaTipo);
									iBonoEntradaPrefactura = Funcion.iEscalarSQL(cdsSeteoF, squery);
									if(iBonoEntradaPrefactura > 0)
									{
										BonoEntradaInactivar(cdsSeteo,(int)this.txtNumIdCompra.Value, idBonoEntradaTipo, 0);
										UsoBono = true;
									}
									//Validar si  existe un registro en bono entrada el estado estado = 0, existe el valor de la idprefactura, esto demuestra que ya se uso uno de los dos 
									//bonos por lo que se enviar?a a cambiar el estado de los otros bonos que tiene el cliente si estos estan activos
									//se deber?a actualizar por el idCliente, idCompra, Estado = 1
									if(UsoBono)
									{
										BonoEntradaInactivarRelacion(cdsSeteo, (int)this.txtNumIdCompra.Value, idBonoEntradaTipo, 0);
									}
									#endregion Bono Entrada 25
								}
							}
					
							#region FamilyFriends
							FamilyFriendsCompraPrefactura(FamilyFriendsService.EstadoFFCompraTemporal, idCompraDest);
							#endregion FamilyFriends
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
								oTransaction.Rollback();

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
				MessageBox.Show("Este documento ya se Factur?.", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);				
				return;
			}

			#region CrediDigital
			int idCre_SolicitudWeb = 0; 
			if (Funcion.iEscalarSQL(cdsSeteoF, string.Format("Exec ValidaBodegaCrediPointdigital {0}", (int) this.cmbBodega.Value)) > 0)
			{
				string sMensaje =  Funcion.sEscalarSQL(cdsSeteoF, string.Format(" EXEC RetornaIdSolicitudWeb '{0}'", (int) this.txtNumIdCompra.Value));
				if ( sMensaje.Length <= 0)
				{
					MessageBox.Show("El RECHAZO de la PRE-FACTURA no esta atada a  CREDI-DIGITAL. \n Comunicate con el area de DESARROLLO", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);				
					return;

				}

				idCre_SolicitudWeb =	Funcion.iEscalarSQL(cdsSeteoF, string.Format("Exec [RetornaIdSolicitudWebidCre_SolicitudWeb] {0}", (int)this.txtNumIdCompra.Value));
				if( idCre_SolicitudWeb <= 0)
				{
					MessageBox.Show("RECHAZO de la PRE-FACTURA no esta atada a  CREDI-DIGITAL. \n Comunicate con el area de DESARROLLO ", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Information);
					return;
				}
			
			}
			#endregion CrediDigital
			
			if (Funcion.ValidaBloqueo((int)this.txtNumIdCompra.Value, 1, cdsSeteoF)) return;
			
			using (frmJustAnulaVenta Nota = new frmJustAnulaVenta())
			{
				if (DialogResult.OK == Nota.ShowDialog())
				{
					string sSQL = string.Format("Exec ActualizaEdoPedido {0}, 4, '', 0, ''", (int)this.txtNumIdCompra.Value);
					Funcion.EjecutaSQL(cdsSeteoF, sSQL, true);
					if(idCre_SolicitudWeb> 0)
					{
						Funcion.EjecutaSQL(cdsSeteoF, string.Format("Exec [AnulacionRechazoSolicitudesWeb] {0},{1},'{2}'",  idCre_SolicitudWeb, 18, this.txtNumero.Text));
					}
					Funcion.EjecutaSQL(cdsSeteoF, string.Format("Update Compra Set idCre_SolicitudWeb = 0,  Borrar = 1, Notas = '{1}' Where idOrigen = {0} And idTipoFactura = 1", (int)this.txtNumIdCompra.Value, Nota.txtNotas.Text.ToString()));

					Funcion.EjecutaSQL(cdsSeteoF, string.Format("Update Compra Set FechaRechazado = GETDATE(), Notas = '{1}' Where idCompra = {0}", (int)this.txtNumIdCompra.Value, Nota.txtNotas.Text.ToString()));						

					if ((bool) this.chkAnclar.Checked) 
					{
						Funcion.EjecutaSQL(cdsSeteoF, string.Format("Exec [ActualizARechazoPrefactura] '{0}', {1}", (string)this.txtSerieNuovo.Text, (int)this.txtNumIdCompra.Value));
						Funcion.EjecutaSQL(cdsSeteoF, string.Format("Exec [ActualizARechazoPrefactura] '{0}', {1}", (string)this.txtSerieNuovo2.Text, (int)this.txtNumIdCompra.Value));
					
					}

					#region Bono entrada
					int idBonoEntradaTipo = 0;
					//Creo que debo cargar la tabla de Bono entrada tipo e ir realizando un for por cada bono y si se encuentra en vigencia
					DataView dvResultado = Funcion.dvProcedimiento(cdsSeteoF, "Exec BonoEntradaTipoCon");

					if (dvResultado != null)
					{
						foreach (DataRowView rowView in dvResultado)
						{
							DataRow row = rowView.Row;
							idBonoEntradaTipo = Convert.ToInt32(row["idBonoEntradaTipo"]);
							// Si el bono quedó Estado=1 vinculado (no se inactivó antes de rechazo),
							// limpiar idPrefactura para que vuelva a aparecer disponible
							Funcion.EjecutaSQL(cdsSeteoF, String.Format(
								"Update BonoEntrada Set idPrefactura = 0 Where idPrefactura = {0} And Estado = 1 And idBonoEntradaTipo = {1}",
								(int)this.txtNumIdCompra.Value, idBonoEntradaTipo));
							// Si el bono sí llegó a Estado=0 (fue inactivado), reactivar via SP
							BonoEntradaInactivar(cdsSeteoF,(int)this.txtNumIdCompra.Value, idBonoEntradaTipo, 1);
							BonoEntradaInactivarRelacion(cdsSeteoF,(int)this.txtNumIdCompra.Value, idBonoEntradaTipo, 1);
						}
					}
					#endregion Bono entrada

					SetteartxtNumCuotasMinValue();
					this.Consulta((int)this.txtNumIdCompra.Value);
				}
			}			
		}

		public void SetteartxtNumCuotasMinValue()
		{
			//Setterar para que no se caiga la consulta
			this.txtNumCuotas.MinValue = 0;
		}

		public static void BonoEntradaInactivar(C1.Data.C1DataSet ds, int idBonoEntradaPrefactura, int idBonoEntradaTipo25, int EstadoBonoEntrada)
		{
			if(EstadoBonoEntrada == 1)
			{
				//Habilitar el bono seg?n la vigencia del mismo
				#region Bono Entrada 25
				//Tiene ya asignado un bono a la prefacturas
				int iBonoEntradaPrefactura = 0;
				string squery = String.Format("Select Count(*) From BonoEntrada Where idPrefactura = {0} And Estado = 0 And idBonoEntradaTipo = {1}", idBonoEntradaPrefactura, idBonoEntradaTipo25);
				iBonoEntradaPrefactura = Funcion.iEscalarSQL(ds, squery);
				if(iBonoEntradaPrefactura > 0)
				{
					squery = String.Format("Exec BonoEntradaInactivar {0}, {1}, {2}", idBonoEntradaPrefactura, idBonoEntradaTipo25, EstadoBonoEntrada);
					Funcion.EjecutaSQL(ds, squery);
				}
				#endregion Bono Entrada 25
			}
			else
			{
				string squery = String.Format("Exec BonoEntradaInactivar {0}, {1}, {2}", idBonoEntradaPrefactura, idBonoEntradaTipo25, EstadoBonoEntrada);
				Funcion.EjecutaSQL(ds, squery);
			}
		}

		public static void BonoEntradaInactivarRelacion(C1.Data.C1DataSet ds, int idBonoEntradaPrefactura, int idBonoEntradaTipo, int EstadoBonoEntrada)
		{
			string squery = String.Format("Exec BonoEntradaInactivarRelacion {0}, {1}, {2}", idBonoEntradaPrefactura, idBonoEntradaTipo, EstadoBonoEntrada);
			Funcion.EjecutaSQL(ds, squery);
		}

		private void txtCuotaDeEntrada_ValueChanged(object sender, System.EventArgs e)
		{
		
		}

		private void txtNotas_Leave(object sender, System.EventArgs e)
		{
		
		}

		private void cmbVendedor_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
			if (e.KeyChar == 13) this.txtRuc.Focus();			
		}

		bool bElCombo = false;
		int idComboElimina = 0;
		int idArticuloElimina = 0;

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
			if (bEdicion) cmdTotal();
		}

		private void txtNotas_ValueChanged(object sender, System.EventArgs e)
		{
		
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

					sMensaje = sMensaje + " Codigo : " + dr.Cells["Codigo"].Value.ToString() + " % Dscto : " + dr.Cells[""].Value.ToString();

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
						
			this.txtNumero.Text = "";
			this.txtRuc.Text = "";
			this.txtNombre.Text = "";
			this.txtComprobante.Text = "";
			this.lblEstado.Text = "";
			this.lblUsuario.Text = "";
			this.lblClienteReiterativo.Text = "";
			this.lblClienteReiterativo.Text = "";
			this.lblClienteEntidad.Text = "";
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
			this.cmbBodega.Enabled = false;
			this.cmbBodega1.Enabled = false;
			this.cmbVendedor.Enabled = false;			
			this.chkAnclar.Checked = false;
			this.chkAnclar.Enabled = false;
			this.txtSerieNuovo.Enabled = false;
			this.txtSerieNuovo.Text = "";
			this.txtSerieNuovo2.Enabled = false;
			this.txtSerieNuovo2.Text = "";
			this.lblColateral.Text ="";

			this.txtNumero.Enabled = false;
			this.txtRuc.Enabled = false;
			this.dtFecha.Enabled = false;
			this.txtNumCuotas.Enabled = false;
			this.txtCuotaDeEntrada.Enabled = false;
			this.cboFechaPrimerPago.Enabled = false;
			
			this.txtNotas.Enabled = false;
			this.txtObservaciones.Enabled = false;
			this.chkAutorizar.Enabled = false;
			this.chkVerificar.Enabled = false;
			this.btnAprobacion.Enabled = false;
			this.btnValidarMail.Enabled = false;
			bGuarda = false;
		 

			this.chkAnclar.Visible = true;
			this.txtSerieNuovo.Visible = true;
			this.txtSerieNuovo2.Visible = true;

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
			this.chkFreelance.Checked = false;
			this.chkFreelance.Visible = false;
			FuncionesProcedimientos.EstadoGrids(false, this.ultraGrid2);

			this.ultraGrid2.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "Exec ConsultaDetalleFactura 0");

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
			if(!estadoFF.EsFF)idTipoPrecio = 4;
			bArtPromocion = false;
			idBloqueaTransacciones = 0;
			dClavePorcDscto = 0.00m;
			iDiaMaximoDePago = 0;
			iMinimoDiasPago = 0;
			iMaximoDiasPago = 0;
			idCuponWeb = 0;
			bCuponWeb = false;
			dDsctoReiterativo = 0.00m;
			dDsctoReiterativoVIP = 0.00m;
			dPorcentajeCuponWeb = 0.00m;
			idMail = 0;
			bEmpleado = false;
			bNuovoHabil = false;
			iColateral = 0;
			iNoColateral = 0;
			iEstadoSol=0;
			iAgregaColateral = 0;
			PorcentajeEntrada = 0;
			#endregion Variables

			#region FamilyFriends
			SettearFamilyFriends();
			#endregion FamilyFriends
		}

		private void SettearFamilyFriends()
		{
			reqFF = new FamilyFriendsRequestModel();
			this.lblRecurrente.Text = "";
			estadoFF = new FamilyFriendsEstadoModel();
		}


		private void btnAnular_Click(object sender, System.EventArgs e)
		{
			this.Consulta((int)this.txtNumIdCompra.Value);
			string Mensaje = "?Esta seguro de anular esta Prefactura?.";
			int idCre_SolicitudWebReturn = 0;
			#region CrediDigital
			if (Funcion.iEscalarSQL(cdsSeteoF, string.Format("Exec ValidaBodegaCrediPointdigital {0}", (int) this.cmbBodega.Value)) > 0)
			{
				string sMensaje =  Funcion.sEscalarSQL(cdsSeteoF, string.Format(" EXEC RetornaIdSolicitudWeb '{0}'", (int) this.txtNumIdCompra.Value));
				if ( sMensaje.Length > 0)
				{
					Mensaje = Mensaje + "\n Recuerde que esta atada \n CREDI-DIGITAL N? " + sMensaje ;
				}
				else
				{
					MessageBox.Show("RECHAZO de la PRE-FACTURA no esta atada a  CREDI-DIGITAL. \n Comunicate con el area de DESARROLLO ", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Information);
					return;
				}
				idCre_SolicitudWebReturn = Funcion.iEscalarSQL(cdsSeteoF, string.Format("Exec RetornaIdSolicitudWebidCre_SolicitudWeb {0}", (int)this.txtNumIdCompra.Value));
			  
				if( idCre_SolicitudWebReturn <= 0)
				{
					MessageBox.Show("RECHAZO de la PRE-FACTURA no esta atada a  CREDI-DIGITAL. \n Comunicate con el area de DESARROLLO ", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Information);
					return;
				}
			}
			
			#endregion CrediDigital

			if (!Facturacion.ValidaEdicionAnulacion((int)this.txtNumIdCompra.Value, this.ultraGrid2, miAcceso.Anular, (int)this.txtNumEstado.Value, 43, (int)this.cmbBodega.Value, 
				(DateTime)this.dtFecha.Value, true, bAnulado, 0, cdsSeteoF, false, false, false, false, "VENTA", 0, true)) return;
			
			if (DialogResult.Yes == MessageBox.Show(Mensaje, "Point Technology", MessageBoxButtons.YesNo, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button2))
			{
				using (frmJustAnulaVenta Nota = new frmJustAnulaVenta())
				{
					if (DialogResult.OK == Nota.ShowDialog())
					{
						if(idCre_SolicitudWebReturn > 0 )
						{
							Funcion.EjecutaSQL(cdsSeteoF, string.Format("Exec [AnulacionRechazoSolicitudesWeb] {0},{1},'{2}'",  idCre_SolicitudWebReturn, 16, this.txtNumero.Text));
					
						}
					
						Funcion.EjecutaSQL(cdsSeteoF, string.Format("Update Compra Set idCre_SolicitudWeb = 0, FechaAnulacion = GETDATE(), Borrar = 1, Notas = '{1}' Where idCompra = {0}", (int)this.txtNumIdCompra.Value, Nota.txtNotas.Text.ToString()));
						
						//Funcion.EjecutaSQL(cdsSeteoF, string.Format("update RegistroNOVO set borrar = 1 where idCompra =  {0}", (int)this.txtNumIdCompra.Value));
						if ((bool) this.chkAnclar.Checked) 
						{
							Funcion.EjecutaSQL(cdsSeteoF, string.Format("Exec [ActualizARechazoPrefactura] '{0}', {1}", (string)this.txtSerieNuovo.Text, (int)this.txtNumIdCompra.Value));
							Funcion.EjecutaSQL(cdsSeteoF, string.Format("Exec [ActualizARechazoPrefactura] '{0}', {1}", (string)this.txtSerieNuovo2.Text, (int)this.txtNumIdCompra.Value));
						}
						
						//Bono entrada — reactivar para todos los tipos configurados
						DataView dvBonoTiposAnul = Funcion.dvProcedimiento(cdsSeteoF, "Exec BonoEntradaTipoCon");
						if (dvBonoTiposAnul != null)
						{
							foreach (DataRowView rowViewAnul in dvBonoTiposAnul)
							{
								int idBETipoAnul = Convert.ToInt32(rowViewAnul.Row["idBonoEntradaTipo"]);
								// Limpiar idPrefactura si bono quedó Estado=1 vinculado (no fue inactivado)
								Funcion.EjecutaSQL(cdsSeteoF, String.Format(
									"Update BonoEntrada Set idPrefactura = 0 Where idPrefactura = {0} And Estado = 1 And idBonoEntradaTipo = {1}",
									(int)this.txtNumIdCompra.Value, idBETipoAnul));
								// Reactivar si bono llegó a Estado=0
								BonoEntradaInactivar(cdsSeteoF,(int)this.txtNumIdCompra.Value, idBETipoAnul, 1);
								BonoEntradaInactivarRelacion(cdsSeteoF,(int)this.txtNumIdCompra.Value, idBETipoAnul, 1);
							}
						}

						this.Consulta((int)this.txtNumIdCompra.Value);
					}
				}
			}
		}

		private void btnEditar_Click(object sender, System.EventArgs e)
		{
			this.Consulta((int)this.txtNumIdCompra.Value);
			
			if (!Facturacion.ValidaEdicionAnulacion((int)this.txtNumIdCompra.Value, this.ultraGrid2, miAcceso.Editar, (int)this.txtNumEstado.Value, 43, (int)this.cmbBodega.Value, Convert.ToDateTime(this.dtFecha.Value), false, bAnulado, 0, cdsSeteoF, bImpreso, false, false, false, "VENTA", 0, true)) return;
			
			int iContDYD = 0;

			foreach(Infragistics.Win.UltraWinGrid.UltraGridRow drDYD in this.ultraGrid2.Rows.All)
				if ((int)drDYD.Cells["idArticulo"].Value == 8335) iContDYD++;
  	
			if (iContDYD > 0)
			{
				MessageBox.Show("No puede editar una Pre Factura con Proteccion D & D.\n\nsi necesita modificar informaci?n primero anule e ingrese una nueva prefactura", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);	
				return;
			}

			#region Habilita
			this.txtNumCuotas.Enabled = false;
			this.txtCuotaDeEntrada.Enabled = true;

			if (miAcceso.CambiarFecha) this.dtFecha.Enabled = true;
			if (miAcceso.BAutEntrada)
			{
				this.chkAutorizar.Enabled = true;
				//this.chkVerificar.Enabled = true;
				this.txtObservaciones.Enabled = true;
			}

			this.cboFechaPrimerPago.Enabled = true;
			this.cmbVendedor.Enabled = true;

			this.txtNotas.Enabled = true;
			this.cmbVendedor.Enabled = true;

			this.btnNuevo.Enabled = false;
			this.btnConsultar.Enabled = false;
			this.btnEditar.Enabled = false;
			this.btnAnular.Enabled = false;
			this.btnImprimir.Enabled = false;
			this.btnGuardar.Enabled = true;
			this.btnAprobacion.Enabled = false;
			this.btnCancelar.Enabled = true;
			this.btnValidarMail.Enabled = true;

			FuncionesProcedimientos.EstadoGrids(true, this.ultraGrid2);

			bNuevo = false;
			bEdicion = true;	

			//			if ((int)this.txtNumEstado.Value == 1 && this.miAcceso.CambiarDescuento) this.btnClaveDescuento.Enabled = true;

			if ((int)this.txtNumEstado.Value == 1)
			{
				string[] sPrmCW = new string[2];
				sPrmCW = Facturacion.sEditarCuponWeb((int)this.txtNumIdCompra.Value, cdsSeteoF);
				idCuponWeb = int.Parse(sPrmCW[0]);
				
				if (idCuponWeb > 0) 
				{
					bCuponWeb = true;
					dPorcentajeCuponWeb = decimal.Parse(sPrmCW[1]);
				}
			}

			bNuovoHabil = Funcion.bEjecutaSQL(cdsSeteoF, string.Format("Select CONVERT (BIT, iNuovo) From dbo.Bodega Where Bodega = {0}", (int)this.cmbBodega.Value));
					
			if (!bNuovoHabil )
			{
				this.chkAnclar.Visible = false;
				this.txtSerieNuovo.Visible = false;
			}
			else
			{
				this.chkAnclar.Visible = true;
				this.txtSerieNuovo.Visible = true;
				this.chkAnclar.Enabled = true;
				this.txtSerieNuovo.Enabled = true;
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
				//				this.btnEditar.Enabled = false;
				this.btnImprimir.Enabled = false;
				this.btnGuardar.Enabled = false;
				this.btnAnular.Enabled = false;
				this.btnCancelar.Enabled = false;
				this.btnAprobacion.Enabled = false;
				this.btnClaveDescuento.Enabled = false;
				bDescuentoActivo = false;
				this.mnuAprobacionAprobar.Enabled = false;
				this.mnuAprobacionRechazar.Enabled = false;
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
					this.cmbBodega1.Value = drPreFactura.GetInt32(39);
					this.chkAnclar.Checked = drPreFactura.GetBoolean(42);
					if ((bool) this.chkAnclar.Checked)
					{
						this.txtSerieNuovo.Text = drPreFactura.GetString(43);
						this.txtSerieNuovo2.Text = drPreFactura.GetString(44);
					}
					if ((int)this.txtidEntidadFinanciera.Value  == 14)
					{
						this.lblClienteEntidad.Text = "Point Movil";
					}
					PorcentajeEntrada = drPreFactura.GetDecimal(46);
				
					#endregion Maestro 
				}
				drPreFactura.Close();

				this.ultraGrid2.DataSource = Funcion.dvProcedimiento(cdsSeteoF, string.Format("Exec ConsultaDetalleFactura {0}", (int)this.txtNumIdCompra.Value));					
				
				this.cmbBodega1.DataSource = Funcion.dvProcedimiento(cdsSeteoF, string.Format("Exec ListaBodegasPorLocal {0}, 0", (int)this.cmbBodega.Value));
				this.cmbVendedor.DataSource = FuncionesProcedimientos.Lista_Vendedores(Convert.ToDateTime(this.dtFecha.Value), (int)this.cmbBodega.Value, cdsSeteoF);

				if (bAnulado) this.lblEstado.Text = "ANULADO";
				else
				{
					if ((int)this.txtNumEstado.Value == 1) this.lblEstado.Text = "PENDIENTE";
					if ((int)this.txtNumEstado.Value == 2) this.lblEstado.Text = "VERIFICACI?N";
					if ((int)this.txtNumEstado.Value == 4) this.lblEstado.Text = "RECHAZADO";
					if ((int)this.txtNumEstado.Value == 5) this.lblEstado.Text = "APROBADO";
					if ((int)this.txtNumEstado.Value == 9) this.lblEstado.Text = "FACTURADO";
				
					if (miAcceso.Editar && (int)this.txtNumEstado.Value == 1 && !bImpreso) this.btnEditar.Enabled = true;					
					if (!bImpreso) if (miAcceso.Imprimir) this.btnImprimir.Enabled = true;
					if (bImpreso) if (miAcceso.ReImprimir) this.btnImprimir.Enabled = true;
					if (miAcceso.Anular && !bAnulado && ((int)this.txtNumEstado.Value == 1 || (int)this.txtNumEstado.Value == 4)) this.btnAnular.Enabled = true;						
					if ((miAcceso.BAprobarPrefactura || miAcceso.BRechazarPrefactura) && ((int)this.txtNumEstado.Value == 1 || (int)this.txtNumEstado.Value == 5)) this.btnAprobacion.Enabled = true;
					if (miAcceso.BAprobarPrefactura) if ((int)this.txtNumEstado.Value == 1) this.mnuAprobacionAprobar.Enabled = true;
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
			
			using (frmBuscaFacturas Busqueda = new frmBuscaFacturas((int)this.txtIdTipoFactura.Value, (int)this.txtBodegaPredef.Value, miAcceso.BLimiteBusquedaFechas))
			{
				if (DialogResult.OK == Busqueda.ShowDialog())
				{
					this.Consulta((int)Busqueda.grdFacturas.ActiveRow.Cells["idCompra"].Value);

					#region FamilyFriends
					ConsultarEstadoFamilyFriends((int)Busqueda.grdFacturas.ActiveRow.Cells["idCompra"].Value);
					#endregion FamilyFriends
				}
			}
		}

		private void ConsultarEstadoFamilyFriends(int idCompra)
		{
			try
			{
				FamilyFriendsService service = new FamilyFriendsService(frmRequerimiento.sconexionPoint);

				// Obtener temporal para validar SIN tocar lo actual
				FamilyFriendsEstadoModel temp = service.ObtenerEstadoFF(idCompra);

				// Solo actualizar si SÍ hay registro válido
				if (temp.EsFF || temp.RechazoPromocion)
				{
					estadoFF = temp; // Actualiza el modelo principal
					idTipoPrecio = estadoFF.IdTipoPrecio;
					this.lblRecurrente.Text = estadoFF.MensajeEstadoFF;
				}
				else
				{
					// NO tocar estadoFF ni idTipoPrecio
					this.lblRecurrente.Text = "";
				}
			}
			catch
			{
				// No cambiar nada si hubo error
				this.lblRecurrente.Text = "";
			}
		}

		private void btnNuevo_Click(object sender, System.EventArgs e)
		{			
			FuncionesInventario._mostrarAsistenciaDental = true;
			this.btnCancelar_Click(sender, e);									

			#region Habilita Controles
			if (miAcceso.BAutEntrada)
			{
				this.chkAutorizar.Enabled = true;
				this.txtObservaciones.Enabled = true;
				//this.chkVerificar.Enabled = true;
			}
						
			
			//this.chkVerificar.Enabled = true;
		
			this.btnNuevo.Enabled = false;
			this.btnConsultar.Enabled = false;
			this.btnEditar.Enabled = false;
			this.btnImprimir.Enabled = false;
			this.btnGuardar.Enabled = true;
			this.btnCancelar.Enabled = true;
			this.btnValidarMail.Enabled = true;
			this.chkAnclar.Enabled = true;
		
			FuncionesProcedimientos.EstadoGrids(true, this.ultraGrid2);

			bNuevo = true;
			bEdicion = true;

			this.txtNumEstado.Value = 1;
			this.txtidTipoRuc.Value	= 5;
			//FuncionesProcedimientos.FechasDocumentos(this.dtFecha);//this.dtFecha.Value = DateTime.Today;
			FuncionesProcedimientos.RetornaFechaServidor(this.dtFecha, cdsSeteoF, false);

			if (miAcceso.CambiarFecha) this.dtFecha.Enabled = true;

			if ((int)this.txtBodegaPredef.Value == 0) 
			{
				this.cmbBodega.Enabled = true;				
				this.cmbBodega.Value = 12;
			}

			if ((int)this.txtBodegaPredef.Value > 0) this.cmbBodega.Value = (int)this.txtBodegaPredef.Value;

			this.txtRuc.Enabled = true;
			this.txtNumCuotas.Enabled = true;
			this.txtCuotaDeEntrada.Enabled = true;
			this.cboFechaPrimerPago.Enabled = true;
			//this.cmbBodega1.Enabled = true;
			this.cmbVendedor.Enabled = true;
			this.txtNotas.Enabled = true;
			
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
					this.cmbBodega1.DataSource= Funcion.dvProcedimiento(cdsSeteoF, string.Format("Exec ListaBodegasPorLocal {0}, 0", (int)this.cmbBodega.Value));

					this.txtNumIdProyecto.Value = Funcion.iEscalarSQL(cdsSeteoF, string.Format("Select isNull(idProyecto, 0) From Bodega Where Bodega = {0}", (int)this.cmbBodega.Value));
					
					if (this.cmbVendedor.ActiveRow != null) this.cmbVendedor.Value = System.DBNull.Value;
					this.cmbVendedor.DataSource = FuncionesProcedimientos.Lista_Vendedores(Convert.ToDateTime(this.dtFecha.Value), (int)this.cmbBodega.Value, cdsSeteoF);
					
					if (bNuevo)
					{
						this.cmbBodega1.Value = System.DBNull.Value;
						if (this.cmbBodega1.Rows.Count == 1) 
							this.cmbBodega1.Value = Funcion.iEscalarSQL(cdsSeteoF, 
								string.Format("Select a.Bodega From AsignaBodegasLocales a Inner Join Bodega b On b.Bodega = a.Bodega Where Local = {0} And Remates = 0 OR a.Consignacion = 1", 
								(int)this.cmbBodega.Value)); //Convert.ToInt32(this.cmbBodega.ActiveRow.Cells["Bodega"].Value);
						else this.cmbBodega1.Enabled = true;
					}

					//					if (bNuevo)
					//					{
					//						string[] sPrmCW = new string[2];
					//						sPrmCW = Facturacion.sCuponWeb(cdsSeteoF);
					//						idCuponWeb = int.Parse(sPrmCW[0]);
					//						
					//						if (idCuponWeb > 0)
					//						{
					//							bCuponWeb = true;
					//							dPorcentajeCuponWeb = decimal.Parse(sPrmCW[1]);
					//						}
					//					}
					bNuovoHabil = Funcion.bEjecutaSQL(cdsSeteoF, string.Format("Select CONVERT (BIT, iNuovo) From dbo.Bodega Where Bodega = {0}", (int)this.cmbBodega.Value));
					if (!bNuovoHabil )
					{
						this.chkAnclar.Visible = false;
						this.txtSerieNuovo.Visible = false;
						this.txtSerieNuovo2.Visible = false;
					}
					else
					{
						this.chkAnclar.Visible = true;
						this.txtSerieNuovo.Visible = true;
						this.txtSerieNuovo2.Visible = true;
					}
				}
			}
		}

		private void cboFechaPrimerPago_SelectedIndexChanged(object sender, System.EventArgs e)
		{
		
		}

		private void btnImprimir_Click(object sender, System.EventArgs e)
		{							
			this.Consulta(Convert.ToInt32(this.txtNumIdCompra.Value));

			decimal dValor = Convert.ToDecimal(this.txtTotal.Value) - Convert.ToDecimal(this.txtCuotaDeEntrada.Value);
			string Fecha = this.cboFechaPrimerPago.Text;

			Facturacion.Impresion((int)this.txtNumIdCompra.Value, bAnulado, (int)this.txtNumEstado.Value, miAcceso.Imprimir, miAcceso.ReImprimir, true, bImpreso, 43, dValor, 
				(int)this.txtNumCuotas.Value, (int)this.txtNumIdCliente.Value, Fecha, (DateTime)this.dtFecha.Value, (int)this.txtidEntidadFinanciera.Value, this.txtRuc.Text.ToString(), cdsSeteoF);
		}

		bool bVIP = false;
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
			
			if (!Validacion.vbCampoEnteroVacio(this.txtNumCuotas, "Escriba el n?mero de cuotas", (int)this.txtNumCuotas.MinValue, (int)this.txtNumCuotas.MaxValue)) return;

			if((int)this.txtidEntidadFinanciera.Value ==14)
			{
				idPlazoD = 0;
				if ((int)this.txtNumCuotas.Value == 13)idPlazoD=1;
				if ((int)this.txtNumCuotas.Value == 26)idPlazoD=1;
				if ((int)this.txtNumCuotas.Value == 39)idPlazoD=1;
				if ((int)this.txtNumCuotas.Value == 52)idPlazoD=1;
				if(idPlazoD == 0 )
				{
					MessageBox.Show("Para venta de equipos Precio Movil el numero de pagos debe ser  13,26,39,52  semanas ", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
					this.txtNumCuotas.Focus();
					return;
				}
			}

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

			//
			//			int  val = Convert.ToDateTime(this.cboFechaPrimerPago.Text).Day;
			//
			//			if (Convert.ToDateTime(this.cboFechaPrimerPago.Text).Day > iDiaMaximoDePago)
			//			{
			//				MessageBox.Show(string.Format("El dia maximo de pago es {0},\n\nUTILICE UNICAMENTE LAS FECHAS DE LA LISTA", iDiaMaximoDePago), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
			//				this.cboFechaPrimerPago.Focus();
			//				return;
			//			}

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
					MessageBox.Show("Para carpa debe escribir una Observaci?n", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
					this.cmbVendedor.Focus();
					return;
				}
			}			
			#endregion Carpa			
						
			#region D&FI
			//			int iYear = Funcion.iEscalarSQL(cdsSeteoF, string.Format("Select DATEDIFF(YEAR, FechaNacimiento, '{0}') From Cre_DatosGenerales Where NumeroIdentificacion = '{1}'", 
			//				Convert.ToDateTime(this.dtFecha.Value).ToString("yyyyMMdd"), this.txtRuc.Text.ToString()));
			//
			//			
			//				int iCont = 0;
			//				foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.ultraGrid2.Rows.All)				
			//					if (dr.Cells["idArticulo"].Value != System.DBNull.Value)				
			//						if ((int)dr.Cells["idArticulo"].Value == 8683) 
			//							iCont++;				
			//
			//			if (iYear < 66)
			//			{
			//				if (iCont == 0)
			//				{
			//					MessageBox.Show("No esta registrado el D&FI", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Error);							
			//					return;
			//				}			
			//			}
			//			if (iYear > 66)
			//			{
			//				if (iCont > 0)
			//				{
			//					MessageBox.Show("Cliente mayor a 65 a?os no aplica a D&FI", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Error);							
			//					return;
			//				}			
			//			}
			#endregion D&FI

			if (!Validacion.vbGrids(this.ultraGrid2, "Codigo", 43)) return;		

			int iclienteReite = 0;
			iclienteReite = Funcion.iEscalarSQL(cdsSeteoF, string.Format("Exec [ValidaReiterativo] '{0}'", (string) this.txtRuc.Text ));

			if (iNoColateral > 0 ) iclienteReite = 1;
			if (iclienteReite == 0)
			{
				#region Valida Nuovo
				/*si el local esta habilitado y si el cliente cumple con el score menor*/
				if ((bool) this.chkAnclar.Checked && iColateral == 1) 
				{
					//				if( iColateral > 0)
					//				{
					if (!Validacion.vbTexto(this.txtSerieNuovo, 3, 50, "Serie NUOVO")) return;
					int idCodNuovoExiste = 0;
					string sMensajeRepetido= "";
					idCodNuovoExiste = Funcion.iEscalarSQL(cdsSeteoF, string.Format("Exec [ValidaCodigoNuovoExistente] {0},'{1}'", (int)this.txtNumIdCompra.Value, this.txtSerieNuovo.Text));
					sMensajeRepetido = Funcion.sEscalarSQL(cdsSeteoF, string.Format("Exec [ValidaCodigoNuovoRepetido] {0},'{1}'", (int)this.txtNumIdCompra.Value, this.txtSerieNuovo.Text));

					if (idCodNuovoExiste == 0 )
					{
						MessageBox.Show("El Codigo Nuovo no existe", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
						return;
					}
					if (sMensajeRepetido.Length > 0)
					{
						MessageBox.Show(sMensajeRepetido, "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
						return;
					}
					//				}
				}
				else
				{
					if((int) this.txtidEntidadFinanciera.Value != 14 )
					{				
						//					if( iclienteReite == 0)
						//					{
						#region Valida Si es combo y si tiene celular
				
						int iNuovoCel = 0;
						foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.ultraGrid2.Rows.All)	
						{
							int iValCel = 0;
							iValCel = Funcion.iEscalarSQL(cdsSeteoF, string.Format("Exec [ValidaNuovoCombo] {0}, {1}",0,(int) dr.Cells["idArticulo"].Value ));	
							if (iValCel > 0)
							{
								iNuovoCel ++;
							}
						}

						#region si existe un celular
						if (iNuovoCel == 1)
						{
							this.chkAnclar.Checked = true;
							if (!Validacion.vbTexto(this.txtSerieNuovo, 3, 50, "Serie NUOVO")) return;
							int idCodNuovoExiste = 0;
							string sMensajeRepetido= "";
							idCodNuovoExiste = Funcion.iEscalarSQL(cdsSeteoF, string.Format("Exec [ValidaCodigoNuovoExistente] {0},'{1}'", (int)this.txtNumIdCompra.Value, this.txtSerieNuovo.Text));
							sMensajeRepetido = Funcion.sEscalarSQL(cdsSeteoF, string.Format("Exec [ValidaCodigoNuovoRepetido] {0},'{1}'", (int)this.txtNumIdCompra.Value, this.txtSerieNuovo.Text));

							if (idCodNuovoExiste == 0 )
							{
								MessageBox.Show("El Codigo Nuovo no existe", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
								return;
							}
							if (sMensajeRepetido.Length > 0)
							{
								MessageBox.Show(sMensajeRepetido, "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
								return;
							}
						}
						#endregion si existe un celular

						#region si existe 2 celular
						if (iNuovoCel > 1)
						{
							this.chkAnclar.Checked = true;
							this.txtSerieNuovo2.Enabled = true;
							#region Valida 1er celular
							if (!Validacion.vbTexto(this.txtSerieNuovo, 3, 50, "Serie NUOVO")) return;
							int idCodNuovoExiste = 0;
							string sMensajeRepetido= "";
							idCodNuovoExiste = Funcion.iEscalarSQL(cdsSeteoF, string.Format("Exec [ValidaCodigoNuovoExistente] {0},'{1}'", (int)this.txtNumIdCompra.Value, this.txtSerieNuovo.Text));
							sMensajeRepetido = Funcion.sEscalarSQL(cdsSeteoF, string.Format("Exec [ValidaCodigoNuovoRepetido] {0},'{1}'", (int)this.txtNumIdCompra.Value, this.txtSerieNuovo.Text));

							if (idCodNuovoExiste == 0 )
							{
								MessageBox.Show("El Codigo Nuovo no existe", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
								return;
							}
							if (sMensajeRepetido.Length > 0)
							{
								MessageBox.Show(sMensajeRepetido, "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
								return;
							}
							#endregion Valida 1er celular

							#region Valida 2do celular
							if (!Validacion.vbTexto(this.txtSerieNuovo2, 3, 50, "Serie NUOVO 2")) return;
							int idCodNuovoExiste2 = 0;
							string sMensajeRepetido2= "";
							idCodNuovoExiste2 = Funcion.iEscalarSQL(cdsSeteoF, string.Format("Exec [ValidaCodigoNuovoExistente] {0},'{1}'", (int)this.txtNumIdCompra.Value, this.txtSerieNuovo2.Text));
							sMensajeRepetido2 = Funcion.sEscalarSQL(cdsSeteoF, string.Format("Exec [ValidaCodigoNuovoRepetido] {0},'{1}'", (int)this.txtNumIdCompra.Value, this.txtSerieNuovo2.Text));

							if (idCodNuovoExiste2 == 0 )
							{
								MessageBox.Show("El Codigo Nuovo 2 no existe", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
								return;
							}
							if (sMensajeRepetido2.Length > 0)
							{
								MessageBox.Show(sMensajeRepetido, "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
								return;
							}
							#endregion Valida 2do celular

							if ((string) this.txtSerieNuovo.Text == (string) this.txtSerieNuovo2.Text)
							{
								MessageBox.Show("El Codigo Nuovo 1 y 2 no pueden ser iguales", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
								return;
							}
						}
						#endregion si existe 2 celular

						#endregion Valida Si es combo y si tiene celular	
						//					}
					}
				}
				#endregion Vlida Nuovo
			}
			#endregion Validaciones

			#region Valida Bodega solicitud
			bool bBodegaSolictud = Funcion.bEscalarSQL(cdsSeteoF, string.Format("Exec ValidaBodegaSolicitud {0}",(int)this.cmbBodega1.Value), true);
			if (bBodegaSolictud)
			{
				if(iEstadoSol!=1)
				{
					MessageBox.Show("La Solicitud de credito no fue completada", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
					this.cboFechaPrimerPago.Focus();
					return;
				}
			}
			#endregion Valida Bodega solicitud

			#region Valida porcentaje entrada
			decimal dCuotaEntrada = 0;
			decimal dEntradaCuota = 0;
			int iConte = 0;
			foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.ultraGrid2.Rows.All)	
			{

				if (dr.Cells["idArticulo"].Value != System.DBNull.Value)
				{
					if(iConte == 0)
					{
						if ((int)dr.Cells["idArticulo"].Value >0)
						{
							decimal dPromocionAdicional = Funcion.decEscalarSQL(cdsSeteoF, string.Format("Select dbo.ValidaIdPromocion ({0}, {1}, {2}, {3})", 
								(int) this.cmbBodega.Value, Convert.ToInt32( dr.Cells["idArticulo"].Value) , 4, (int) this.txtNumCuotas.Value));

							if (dPromocionAdicional > 0)
							{
								dCuotaEntrada = dPromocionAdicional;
								iConte = iConte + 1;
							}
						}
					}	
				}
			}


			if (dCuotaEntrada > 0)
			{
				dEntradaCuota = ( (Decimal) this.txtTotal.Value * dCuotaEntrada )/100;
			}

			if (dEntradaCuota > 0)
			{
				decimal Dmaximo = (Decimal) this.txtTotal.Value;
				if (!Validacion.vbCampoDecimalVacio(this.txtCuotaDeEntrada, "El Valor minimo es ", dEntradaCuota, Dmaximo)) return; 
			}
			#endregion Valida porcentaje entrada

			#region Valida porcentaje entrada con variable global
			if (PorcentajeEntrada > 0)
			{
				decimal total = (decimal)this.txtTotal.Value;
				decimal valorMinimoEntrada = Math.Round(total * PorcentajeEntrada / 100, 2);
				decimal valorUsuario = Convert.ToDecimal(this.txtCuotaDeEntrada.Value);
				if (valorUsuario < valorMinimoEntrada)
				{
					decimal diferencia = valorMinimoEntrada - valorUsuario;
					MessageBox.Show("El monto mínimo de entrada requerido es: $" + valorMinimoEntrada.ToString("N2") + " (" + PorcentajeEntrada.ToString("N2") + "% del total $" + total.ToString("N2") + ")\nFaltan: $" + diferencia.ToString("N2") + " para cumplir el mínimo.", "Validación de Entrada", MessageBoxButtons.OK, MessageBoxIcon.Warning);
					return;
				}
			}
			#endregion Valida porcentaje entrada con variable global
			
			#region Valida Cantidad Celulares 
			if (Funcion.iEscalarSQL(cdsSeteoF, string.Format("Exec Cre_ValidaClienteRecurrente '{0}'", this.txtRuc.Text.ToString())) == 0)
			{
				int iContCel = 0;
				int idSubGrupo = 0;

				foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.ultraGrid2.Rows.All)	
				{
					idSubGrupo = Funcion.iEscalarSQL(cdsSeteoF, string.Format("SELECT idSubGrupo FROM Articulo WHERE idArticulo = {0}", (int)dr.Cells["idArticulo"].Value));
				
					if (idSubGrupo == 23) iContCel++;
				}

				if (iContCel > 1)
				{
					MessageBox.Show(string.Format("No puede ingresar {0} celulares en una misma factura.", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop));
					return;
				}
			}
			#endregion Valida Cantidad Celulares 

			#region Validar Que exista solo un credidigital

			if (Funcion.iEscalarSQL(cdsSeteoF, string.Format("Exec ValidaBodegaCrediPointdigital {0}", (int) this.cmbBodega.Value)) > 0)
			{
				string sMensajes = Funcion.sEscalarSQL(cdsSeteoF, string.Format("Exec [CountSolicitudesporID] {0},'{1}'", (int) this.cmbBodega.Value, this.txtRuc.Text));
				if( sMensajes.Length > 0)
				{
					MessageBox.Show(sMensajes, "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
					return;
				}
			
			}
			#endregion Validar Que exista solo un credidigital

			#region Celulares con el 10% entrada
			if((int) this.txtidEntidadFinanciera.Value != 14 )
			{
				if( iclienteReite == 0)
				{
					int iNuovoCelEn = 0;
					decimal dPRecioCel = 0.00m;
					decimal dValentCelu = 0.00m;
					foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.ultraGrid2.Rows.All)	
					{
						if ( dr.Cells["Precio"].Value != System.DBNull.Value && dr.Cells["idArticulo"].Value != System.DBNull.Value)
						{
							int iValCel = 0;
			
							iValCel = Funcion.iEscalarSQL(cdsSeteoF, string.Format("Exec ValidaNuovoCombo {0}, {1}",0,(int) dr.Cells["idArticulo"].Value ));	
							if (iValCel > 0)
							{
								dPRecioCel = dPRecioCel + Convert.ToDecimal(dr.Cells["Precio"].Value);
								iNuovoCelEn ++;
							}
						}
					}

					if(iNuovoCelEn > 0 && dPRecioCel > 0)
					{
					
						dValentCelu = Math.Round(dPRecioCel * 10)/100;
						//MessageBox.Show("!!?", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
						//MessageBox.Show(Convert.ToString( dValentCelu), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
						
						if(Convert.ToDecimal( this.txtCuotaDeEntrada.Value ) <  Convert.ToDecimal( dValentCelu))
						{
							//	MessageBox.Show(Convert.ToString( dValentCelu), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
							MessageBox.Show("EL VALOR DE LA ENTRADA NO SE ENCUENTRA EN EL RANGO PERMITIDO SEGUN LA POLITICA DE CREDITO VIGENTE!!!.", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
							return;
						}
					}
				}
			}
			#endregion Celulares con el 10% entrada

			#region Validaciones iPhone y Accesorios

			// Verificar si hay un iPhone en el grid
			bool tieneIPhone = false;
			int cantidadIPhones = 0;
			// Bandera para iPhone 15 (IPHO00002, idArticulo = 13563)
			bool tieneIPhone15ProMax = false;
			// Bandera para iPhone 14 (IPHO00005, idArticulo = 14019)
			bool tieneIPhone14 = false;
			bool tieneHonorMagic = false;

			foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.ultraGrid2.Rows.All)
			{
				if (dr.Cells["idArticulo"].Value != System.DBNull.Value)
				{
					int idArticulo = (int)dr.Cells["idArticulo"].Value;
		
					// Validar si es un iPhone
					int esIPhone = Funcion.iEscalarSQL(cdsSeteoF, 
						string.Format("SELECT COUNT(*) FROM Articulo WHERE idGrupoArticulo = 8 AND idSubGrupo in (401,402) AND idArticulo = {0}", idArticulo));
		
					if (idArticulo == 14067) // HONOR MAGIC 8
					{
						tieneHonorMagic = true;
					}

					if (esIPhone > 0)
					{
						tieneIPhone = true;
			
						// Obtener cantidad del iPhone
						if (dr.Cells["Cantidad"].Value != System.DBNull.Value)
						{
							cantidadIPhones += Convert.ToInt32(dr.Cells["Cantidad"].Value);
						}
						// Detectar específicamente el iPhone 15 (IPHO00002)
						if (idArticulo == 13563 || idArticulo == 13897)
						{
							tieneIPhone15ProMax = true;
						}
						if (idArticulo == 14019)
						{
							tieneIPhone14 = true;
						}
					}
				}
			}

			#endregion Validaciones iPhone y Accesorios

			#region Celulares 12 meses
			int idArticuloMax =0;
			if((int) this.txtNumCuotas.Value > 12)
			{
				decimal maxPrecio = 0;
				foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.ultraGrid2.Rows.All)	
				{
					decimal precio = Convert.ToDecimal(dr.Cells["Precio"].Value);
					if (precio > maxPrecio)
					{
						maxPrecio = precio; 
						idArticuloMax = Convert.ToInt32( dr.Cells["idArticulo"].Value);  // Actualizamos el idArticulo correspondiente
					}
				}

				int idGrupoArticulo = 0;
				idGrupoArticulo = Funcion.iEscalarSQL(cdsSeteoF, string.Format("SELECT idGrupoArticulo FROM Articulo WHERE idArticulo = {0}", idArticuloMax));
			
				if(idGrupoArticulo == 8 && !tieneIPhone)
				{
					MessageBox.Show("La cuota Maxima para Movilidad/Celulares es de 12 Cuotas.", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
					return;
				}
			}
			#endregion Celulares 12 meses

			#region Validaciones Honor Magic
//			if (tieneHonorMagic)
//			{
//				// Validación 1: Número de cuotas entre 3 y 15
//				int numCuotasMagic = (int)this.txtNumCuotas.Value;
//				int maxCuotasMagic = 15;
//
//				if (numCuotasMagic < 3 || numCuotasMagic > maxCuotasMagic)
//				{
//					string mensajeCuotasMagic = "El número de cuotas debe estar entre 3 y 15 para ventas con Honor Magic.";
//
//					MessageBox.Show(mensajeCuotasMagic, "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
//					this.txtNumCuotas.Focus();
//					return;
//				}
//
//					
//				if (numCuotasMagic > 12 && numCuotasMagic <= 15)
//				{
//					// Validar que la cuota de entrada sea al menos el 10% del total del credito 
//					decimal entradaMagic = 0;
//					if (this.txtCuotaDeEntrada.Value != null && this.txtCuotaDeEntrada.Value != System.DBNull.Value)
//					{
//						entradaMagic = Convert.ToDecimal(this.txtCuotaDeEntrada.Value);
//					}
//		
//					decimal minimoEntradaMagic = Math.Round(Convert.ToDecimal(this.txtTotal.Value) * 0.10m, 2);
//		
//					if (entradaMagic < minimoEntradaMagic)
//					{
//						MessageBox.Show(string.Format("Para este teléfono, la cuota de entrada debe ser al menos el 10% del valor total del credito.\n\nMínimo requerido (10%): ${0:N2}", 
//							minimoEntradaMagic), 
//							"Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
//						this.txtCuotaDeEntrada.Focus();
//						return;
//					}
//				}
//			}
			#endregion Validaciones Honor Magic

			#region Comentario
			//			if (bEmpleado)
			//			{
			//				if (Convert.ToDecimal(this.txtCuotaDeEntrada.Value) < 100.00m)
			//				{
			//					MessageBox.Show("El valor de entrada no puede ser menor a 100.00 usd", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
			//					return;
			//				}
			//
			//				if (Convert.ToDecimal(this.txtValorCuotas.Value) > 65.00m)
			//				{
			//					MessageBox.Show("El valor minimo de cuotas no puede ser menor a 65.00 usd", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
			//					return;
			//				}
			//			}
			#endregion Comentario

			#region CREDI-DIGITAL
			int idCre_SolicitudWebReturn = 0;
			if (Funcion.iEscalarSQL(cdsSeteoF, string.Format("Exec ValidaBodegaCrediPointdigital {0}", (int) this.cmbBodega.Value)) > 0)
			{
			 idCre_SolicitudWebReturn = Funcion.iEscalarSQL(cdsSeteoF, string.Format("Exec [RetornaidSolciitud] {0},'{1}'", (int) this.cmbBodega.Value, this.txtRuc.Text));
			if(idCre_SolicitudWebReturn <= 0)
			{
				MessageBox.Show("La PRE-FACTURA no encuentra CREDI-DIGITAL. \n Comunicate con el area de DESARROLLO ", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Information);
				return;
			}
					}

			#endregion region CREDI-DIGITAL	

			cmdTotal();
						
			decimal dBaseImponible = (decimal)this.txtIva.Value + (decimal)this.txtIva0.Value - (decimal)this.txtDescIvaTotal.Value - (decimal)this.txtDescIva0Total.Value;

			#region Seguros
			//			int iContSeg = 0;
			//			foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in uGrid.Rows.All)
			//				if ((int)dr.Cells["idArticulo"].Value == 8683) iContSeg++;
			//
			//			if (Funcion.bEscalarSQL(cdsSeteoF, "Select Seguro From SeteoF", true))
			//			{
			//				if (iContSeg == 0)
			//				{
			//					int iDias = Funcion.iEscalarSQL(cdsSeteoF, string.Format("Select DATEDIFF(DAY, '{0}', '{1}')", 
			//						Convert.ToDateTime(this.dtFecha.Value).ToString("yyyyMMdd"), DateTime.Parse(this.cboFechaPrimerPago.Text.ToString()).ToString("yyyyMMdd")));
			//
			//					if (iDias > 55 || (int)this.txtNumCuotas.Value > 18)
			//					{
			//						MessageBox.Show("Con 60 dias de gracia o cuotas de pago mayor a 18 es obligatorio Protecci?n D & D", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Information);
			//
			//						FuncionesInventario.FacturacionCargaDYD(this.ultraGrid2, Convert.ToDecimal(this.txtProteccionDD.Value), dPorcIva, cdsSeteoF);
			//					}
			//					else if (DialogResult.Yes == MessageBox.Show("?Desea cargar Protecci?n D & D?", "Point Technology", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2)) 
			//						FuncionesInventario.FacturacionCargaDYD(this.ultraGrid2, Convert.ToDecimal(this.txtProteccionDD.Value), dPorcIva, cdsSeteoF);
			//				}
			//			}			
			#endregion Seguros

			#region Seguro de desgravamen
			if ((int)this.txtidEntidadFinanciera.Value != 14) 
			{
				if (Funcion.bEscalarSQL(cdsSeteoF, "Select SeguroDesgravamen From SeteoF", true))
				{
					int iYear = 0;
					#region CrediDigital
					if (Funcion.iEscalarSQL(cdsSeteoF, string.Format("Exec ValidaBodegaCrediPointdigital {0}", (int) this.cmbBodega.Value)) == 0)
					{
						iYear = Funcion.iEscalarSQL(cdsSeteoF, string.Format("Select DATEDIFF(YEAR, FechaNacimiento, '{0}') From Cre_DatosGenerales Where NumeroIdentificacion = '{1}'", 
								 Convert.ToDateTime(this.dtFecha.Value).ToString("yyyyMMdd"), this.txtRuc.Text.ToString()));
					}
					else
					{
						iYear = Funcion.iEscalarSQL(cdsSeteoF, string.Format("Select top 1 DATEDIFF(YEAR, FechaNacimiento, '{0}') From Web_SolicitudGrande Where Cedula = '{1}'", 
								 Convert.ToDateTime(this.dtFecha.Value).ToString("yyyyMMdd"), this.txtRuc.Text.ToString()));
					}

					#endregion CrediDigital

					if (iYear < 71)
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
														FuncionesInventario.FacturacionCargaSeguroDesgravamen(this.ultraGrid2, dValorSeguro, dPorcIva,  10191, false, false, cdsSeteoF);		
														bActualiza = true;

														cmdTotal();
													}
												}
				}
			} 
			#endregion Seguro de desgravamen

			#region POINT RESPALDO
			if (Funcion.bEscalarSQL(cdsSeteoF, "Select PointRespaldo From SeteoF", true))
			{
				// 1) Buscar si ya existe la fila del Point Respaldo (idArticulo = 12876)
				Infragistics.Win.UltraWinGrid.UltraGridRow rowPointResp = null;

				foreach (Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.ultraGrid2.Rows.All)
				{
					if (dr.Cells["idArticulo"].Value != DBNull.Value &&
						Convert.ToInt32(dr.Cells["idArticulo"].Value) == 12876)
					{
						rowPointResp = dr;
						break;
					}
				}

				// 2) Valores de este form
				decimal valorCuota = Convert.ToDecimal(this.txtValorCuotas.Value);
				int numeroCuotas = Convert.ToInt32(this.txtNumCuotas.Value);

				// (si no lo usas, puedes quitarlo)
				decimal dMonto = Convert.ToDecimal(this.txtTotal.Value) - Convert.ToDecimal(this.txtCuotaDeEntrada.Value);

				// 3) Consultar opcionalidad
				string sSQLOpc = string.Format("Exec Cre_OpcionalidadPointRespaldo {0}", valorCuota);
				bool esOpcional = Funcion.bEscalarSQL(cdsSeteoF, sSQLOpc, true);

				bool debeCargar = true; // por defecto SIEMPRE carga

				if (esOpcional)
				{
					DialogResult resp = MessageBox.Show(
						"¿Desea cargar el Point Respaldo?",
						"Point Respaldo",
						MessageBoxButtons.YesNo,
						MessageBoxIcon.Question
						);

					if (resp == DialogResult.No)
					{
						debeCargar = false;
					}
				}

				if (!debeCargar)
				{
					// Si el usuario dijo NO (solo aplica cuando esOpcional == true)
					if (rowPointResp != null)
					{
						rowPointResp.Delete(false);
					}

					cmdTotal();
				}
				else
				{
					// 4) Debe cargar: calcular totalRespaldo y agregar/actualizar
					string sSQL = string.Format(
						"EXEC Cre_CalcularPrecioPointRespaldo {0}, {1}",
						valorCuota,
						numeroCuotas
						);

					decimal totalRespaldo = Funcion.decEscalarSQL(cdsSeteoF, sSQL);

					if (rowPointResp == null)
					{
						bActualiza = false;
						FuncionesInventario.FacturacionCargaSeguroDesgravamen(
							this.ultraGrid2,
							totalRespaldo,
							dPorcIva,
							12876,
							false,   // mantengo tu valor actual
							esOpcional,
							cdsSeteoF
							);
						bActualiza = true;
					}
					else
					{
						rowPointResp.Cells["Precio"].Value = totalRespaldo;
					}

					cmdTotal();
				}
			}
			#endregion POINT RESPALDO
						
			cmdTotal();

			
			#region Validaciones IPhone
			
			// Solo ejecutar validaciones si hay un iPhone en el grid
			if (tieneIPhone)
			{
				decimal cuotaEntrada = 0;

				// Validación 0: Bancarizado C
				int idTipoCliente = Funcion.iEscalarSQL(cdsSeteoF, 
					string.Format("SELECT ISNULL(idTipoCliente, 0) FROM Cre_SolicitudWeb WHERE Estado = 1 AND Cedula = '{0}'", this.txtRuc.Text.ToString().Trim()));
	
				if (idTipoCliente == 8)
				{
					// Obtener el valor del iPhone
					decimal valorIPhone = 0;
		
					foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.ultraGrid2.Rows.All)
					{
						if (dr.Cells["idArticulo"].Value != System.DBNull.Value)
						{
							int idArticulo = (int)dr.Cells["idArticulo"].Value;
				
							// Verificar si es iPhone
							int esIPhone = Funcion.iEscalarSQL(cdsSeteoF, 
								string.Format("SELECT COUNT(*) FROM Articulo WHERE idGrupoArticulo = 8 AND idSubGrupo in (401,402) AND idArticulo = {0}", idArticulo));
				
							if (esIPhone > 0)
							{
								if (dr.Cells["SubTotal"].Value != System.DBNull.Value)
								{
									valorIPhone = Convert.ToDecimal(dr.Cells["SubTotal"].Value);
								}
								break;
							}
						}
					}
		
					// Validar que la cuota de entrada sea al menos el 15% del valor del iPhone
					if (this.txtCuotaDeEntrada.Value != null && this.txtCuotaDeEntrada.Value != System.DBNull.Value)
					{
						cuotaEntrada = Convert.ToDecimal(this.txtCuotaDeEntrada.Value);
					}
		
					decimal minimoEntrada = valorIPhone * 0.15m;
		
					if (cuotaEntrada < minimoEntrada)
					{
						MessageBox.Show(string.Format("Para clientes Bancarizado C, la cuota de entrada debe ser al menos el 15% del valor del iPhone.\n\nValor iPhone: ${0:N2}\nMínimo requerido (15%): ${1:N2}\nCuota de entrada actual: ${2:N2}", 
							valorIPhone, minimoEntrada, cuotaEntrada), 
							"Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
						this.txtCuotaDeEntrada.Focus();
						return;
					}
				}

				
				// Validación de entrada mínima 35 USD para TODAS las ventas con iPhone
				// determinar si es cliente recurrente (9, 10, 11)
				bool esClienteRecurrente = (idTipoCliente == 9 || idTipoCliente == 10 || idTipoCliente == 11);

				// excepto si es cliente recurrente
				if (!esClienteRecurrente)
				{
					if (this.txtCuotaDeEntrada.Value == null || this.txtCuotaDeEntrada.Value == System.DBNull.Value)
					{
						MessageBox.Show(
							"Para ventas con iPhone se requiere una cuota de entrada mínima de $35.00.",
							"Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
						this.txtCuotaDeEntrada.Focus();
						return;
					}

					cuotaEntrada = Convert.ToDecimal(this.txtCuotaDeEntrada.Value);

					if (cuotaEntrada < 35m)
					{
						MessageBox.Show(
							string.Format("Para ventas con iPhone se requiere una cuota de entrada mínima de $35.00.\nCuota de entrada actual: ${0:N2}", cuotaEntrada),
							"Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
						this.txtCuotaDeEntrada.Focus();
						return;
					}
				}

				// Validación 1: Número de cuotas entre 3 y 15
				int numCuotas = (int)this.txtNumCuotas.Value;
				int maxCuotas = tieneIPhone15ProMax ? 18 : 15;

				if (numCuotas < 3 || numCuotas > maxCuotas)
				{
					string mensajeCuotas = tieneIPhone15ProMax
						? "El número de cuotas debe estar entre 3 y 18 para ventas con iPhone 15 Pro Max."
						: "El número de cuotas debe estar entre 3 y 15 para ventas con iPhone.";

					MessageBox.Show(mensajeCuotas, "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
					this.txtNumCuotas.Focus();
					return;
				}

	
				// Validación 2: Solo puede haber cantidad 1 de iPhone
				if (cantidadIPhones != 1)
				{
					MessageBox.Show("Solo se permite la venta de 1 iPhone por prefactura.", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
					return;
				}
	
					// Validación 3: Los demás productos deben pertenecer a grupos 1, 28 o 63 (asistencias/seguros)
				decimal totalAccesorios = 0;
				
				foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.ultraGrid2.Rows.All)
				{
					if (dr.Cells["idArticulo"].Value != System.DBNull.Value)
					{
						int idArticulo = (int)dr.Cells["idArticulo"].Value;
						
						// Verificar si NO es iPhone
						int esIPhone = Funcion.iEscalarSQL(cdsSeteoF, 
							string.Format("SELECT COUNT(*) FROM Articulo WHERE idGrupoArticulo = 8 AND idSubGrupo in (401,402) AND idArticulo = {0}", idArticulo));
						
						if (esIPhone == 0 && (int)dr.Cells["Tipo"].Value == 0)
						{
							// Verificar si es asistencia/seguro (grupo 63) - estos tienen bypass
							int esAsistenciaSeguro = Funcion.iEscalarSQL(cdsSeteoF, 
								string.Format("SELECT COUNT(*) FROM Articulo WHERE idGrupoArticulo = 63 AND idArticulo = {0}", idArticulo));
							
							if (esAsistenciaSeguro == 0)
							{
								// No es asistencia/seguro, validar que pertenezca a grupos 1 o 28
								int esAccesorioValido = Funcion.iEscalarSQL(cdsSeteoF, 
									string.Format("SELECT COUNT(*) FROM Articulo WHERE idGrupoArticulo IN (1,28) AND idArticulo = {0}", idArticulo));
								
								if (esAccesorioValido == 0)
								{
									string nombreArticulo = dr.Cells["Codigo"].Value != System.DBNull.Value ? dr.Cells["Codigo"].Value.ToString() : "Desconocido";
									MessageBox.Show(string.Format("El producto '{0}' no es un accesorio válido para ventas con iPhone.\nSolo se permiten accesorios de los grupos 1 y 28, o asistencias/seguros.", nombreArticulo), 
										"Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
									return;
								}
								
								// Sumar el valor del accesorio (solo grupos 1 y 28, NO grupo 63)
								if (dr.Cells["SubTotal"].Value != System.DBNull.Value)
								{
									totalAccesorios += Convert.ToDecimal(dr.Cells["SubTotal"].Value);
								}
							}
							// Si es grupo 63 (asistencia/seguro), no suma al total de accesorios
						}
					}
				}
	
	
				// Validación 4: La suma de accesorios no debe superar los $100
				if (totalAccesorios > 100)
				{
					MessageBox.Show(string.Format("La suma de los accesorios (${0:N2}) supera el límite permitido de $100.00 para ventas con iPhone.", totalAccesorios), 
						"Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
					return;
				}
	
				// Validación 5: Total - Cuota de Entrada debe ser $1600 o menos
				decimal total = 0;
	
				// Validar y obtener el Total
				if (this.txtTotal.Value != null && this.txtTotal.Value != System.DBNull.Value)
				{
					total = Convert.ToDecimal(this.txtTotal.Value);
				}
	
				// Validar y obtener la Cuota de Entrada
				if (this.txtCuotaDeEntrada.Value != null && this.txtCuotaDeEntrada.Value != System.DBNull.Value)
				{
					cuotaEntrada = Convert.ToDecimal(this.txtCuotaDeEntrada.Value);
				}
	
				decimal diferencia = total - cuotaEntrada;
	
				if (diferencia > 2200 && tieneIPhone15ProMax)
				{
					MessageBox.Show(string.Format("La diferencia entre el Total (${0:N2}) y la Cuota de Entrada (${1:N2}) es de ${2:N2}.\nEsta diferencia no puede superar los $2,200.00 para ventas con iPhone.", 
						total, cuotaEntrada, diferencia), 
						"Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
					this.txtTotal.Focus();
					return;
				}
				else if (diferencia > 1790 && tieneIPhone14)
				{
					MessageBox.Show(string.Format("La diferencia entre el Total (${0:N2}) y la Cuota de Entrada (${1:N2}) es de ${2:N2}.\nEsta diferencia no puede superar los $1,790.00 para ventas con iPhone.", 
						total, cuotaEntrada, diferencia), 
						"Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
					this.txtTotal.Focus();
					return;
				}
				else if (diferencia > 1600 && !tieneIPhone15ProMax && !tieneIPhone14)
				{
					MessageBox.Show(string.Format("La diferencia entre el Total (${0:N2}) y la Cuota de Entrada (${1:N2}) es de ${2:N2}.\nEsta diferencia no puede superar los $1,600.00 para ventas con iPhone.", 
						total, cuotaEntrada, diferencia), 
						"Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
					this.txtTotal.Focus();
					return;
				}
			}
			#endregion Validaciones IPhone
						
			#region Valida Politicas
			if ((int)this.txtidEntidadFinanciera.Value == 14)
				if (!Validacion.vbPointMovil(this.ultraGrid2, cdsSeteoF)) return;

			int idGrupo = 0;
			decimal dPrecioPrd = 0.00m;
			foreach(Infragistics.Win.UltraWinGrid.UltraGridRow drVP in this.ultraGrid2.Rows.All)	
			{
				if (idGrupo == 0)
				{
					idGrupo = Funcion.iEscalarSQL(cdsSeteoF, string.Format("Select idGrupoArticulo From Articulo Where idArticulo = {0}", (int)drVP.Cells["idArticulo"].Value));
					dPrecioPrd = Convert.ToDecimal(drVP.Cells["Precio"].Value);
				}
				else 
				{
					if (dPrecioPrd < Convert.ToDecimal(drVP.Cells["Precio"].Value))
					{
						idGrupo = Funcion.iEscalarSQL(cdsSeteoF, string.Format("Select idGrupoArticulo From Articulo Where idArticulo = {0}", (int)drVP.Cells["idArticulo"].Value));
						dPrecioPrd = Convert.ToDecimal(drVP.Cells["Precio"].Value);
					}
				}
			}
						
			string sMensajeVP = "";
			bool bPrimeraCuota = false;
			string sSQLVPC = string.Format("Exec Cre_ValidaPoliticasCredito '{0}', {1}, {2}, {3}, '{4}', {5}, {6}, {7}", 
				this.txtRuc.Text.ToString(), idGrupo, Convert.ToDecimal(this.txtValorCuotas.Value), 
				(int)this.txtNumCuotas.Value, Convert.ToDateTime(this.dtFecha.Value).ToString("yyyyMMdd"),
				Convert.ToDecimal(this.txtCuotaDeEntrada.Value), Convert.ToDecimal(this.txtTotal.Value), (int)this.cmbBodega.Value);
			SqlDataReader drVPC = Funcion.rEscalarSQL(cdsSeteoF, sSQLVPC, true);
			drVPC.Read();
			if (drVPC.HasRows)
			{
				sMensajeVP = drVPC.GetString(0);
				bPrimeraCuota = drVPC.GetBoolean(1);
			}
			drVPC.Close();

			if (sMensajeVP.Length > 0)
			{	
				MessageBox.Show(sMensajeVP, "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				return;
			}
			#endregion Valida Politicas

			#region Parametros de cr?dito
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
			#endregion Parametros de cr?dito			

			#region Fundas
			if (Convert.ToDecimal(this.txtTotal.Value) > 99.99m)
			{
				int iSumaCantidad=0;
				int iFundaMediana= 10717;
				int iFundaGrande= 10718;
							
				#region Existencia Fundas Medianas
				DateTime dtFechaFM = (DateTime)this.dtFecha.Value;
				int ExistenciaFundasM = Funcion.iEscalarSQL(cdsSeteoF, string.Format("Select dbo.RetornaExistencia('{0}', {1}, {2}, {3},0,0)", 
					dtFechaFM.ToString("yyyyMMdd"), (int)this.txtNumIdCompra.Value, (int)this.cmbBodega.Value, iFundaMediana));
				#endregion Existencia Fundas Medianas

				#region Existencia Fundas Grandes
				DateTime dtFechaFG = (DateTime)this.dtFecha.Value;
				int ExistenciaFundasG = Funcion.iEscalarSQL(cdsSeteoF, string.Format("Select dbo.RetornaExistencia('{0}', {1}, {2}, {3},0,0)", 
					dtFechaFG.ToString("yyyyMMdd"), (int)this.txtNumIdCompra.Value, (int)this.cmbBodega.Value, iFundaGrande));
				#endregion Existencia Fundas Grandes

				iSumaCantidad=ExistenciaFundasM+ExistenciaFundasG;

				if (iSumaCantidad>0 )
				{
					Fundas();
				}
				iSumaCantidad=0;

			}
			#endregion Fundas

			cmdTotal();

			#region Elimina Filas Vacias
			bEliminaAlValidar = true;
			Validacion.vEliminaFilasVacias(this.ultraGrid2, "idArticulo");
			bEliminaAlValidar = false;
			#endregion Elimina Filas Vacias

			int idBonoEntradaTipo = 0;
			bool bActualizarIdPrefactura = false;
			int iBonoEntradaSel = 0;

			//Creo que debo cargar la tabla de Bono entrada tipo e ir realizando un for por cada bono y si se encuentra en vigencia
			DataView dvResultado = Funcion.dvProcedimiento(cdsSeteoF, "Exec BonoEntradaTipoCon");
			
			if (dvResultado != null && bNuevo)
			{
				foreach (DataRowView rowView in dvResultado)
				{
					DataRow row = rowView.Row;
					idBonoEntradaTipo = Convert.ToInt32(row["idBonoEntradaTipo"]);

					#region Bono Entrada 25
					//Suma del bono estrada 25
					//Validar si el cliente tiene el bono por lo que saco el id del cliente, y la fecha actual, para ver si es posible utilizar el bono?,
					//pero lo mejor es validar por el estado ya que va existir un job que desabilite los bonos
					int iBonoEntrada25IdCliente = (int)this.txtNumIdCliente.Value;
					decimal dBonoEntrada25 = 0; 
					int iBonoEntradaCompraCount = 0;
					bool iRegistroBonoEntrada25 = false;
					bool bBonoEntradaCodicion = false;
					decimal ValorCreditoHasta = 0;
					decimal ValorCreditoDesde = 0;
					int idBonoEntradaValoresCredito = 0;
					decimal getValorBono = 0;

					//Creo que debo cargar la tabla de Bono entrada tipo e ir realizando un for por cada bono y si se encuentra en vigencia
					string sconsulta = String.Format("Exec BonoEntradaValoresCreditoConXId {0}", idBonoEntradaTipo);
					DataView dvBonoEntradaValoresCreditoConXId = Funcion.dvProcedimiento(cdsSeteoF, sconsulta);
					if (dvBonoEntradaValoresCreditoConXId != null)
					{
						foreach (DataRowView rowViewBonoEntradaValoresCreditoConXId in dvBonoEntradaValoresCreditoConXId)
						{
							DataRow rowBonoEntradaValoresCreditoConXId = rowViewBonoEntradaValoresCreditoConXId.Row;
							idBonoEntradaValoresCredito = Convert.ToInt32(rowBonoEntradaValoresCreditoConXId["idBonoEntradaValoresCredito"]);
							ValorCreditoDesde = Convert.ToDecimal(rowBonoEntradaValoresCreditoConXId["ValorCreditoDesde"]);
							ValorCreditoHasta = Convert.ToDecimal(rowBonoEntradaValoresCreditoConXId["ValorCreditoHasta"]);
							decimal dBonoEntradaTotal = Convert.ToDecimal(this.txtTotal.Value);
							if(dBonoEntradaTotal >= ValorCreditoDesde && dBonoEntradaTotal <= ValorCreditoHasta)
							{
								getValorBono = Convert.ToDecimal(rowBonoEntradaValoresCreditoConXId["Valor"]);
								bBonoEntradaCodicion = true;
							}
						}
					}
					if(bBonoEntradaCodicion)
					{
						//Elegir de que factura va a coger el bono entrada 25, ya que podr?a tener dos, en el caso que tenga 2, si tiene uno no se mostrar?a la pantalla
						string squery = String.Format("Select Count(*) from BonoEntrada Where idCliente = {0} and idBonoEntradaTipo = {1} And Estado = 1",
							iBonoEntrada25IdCliente,
							idBonoEntradaTipo);
						iBonoEntradaCompraCount = Funcion.iEscalarSQL(cdsSeteoF, squery);
						if(iBonoEntradaCompraCount > 0)
						{
							iRegistroBonoEntrada25 = true;
						}
						if(iRegistroBonoEntrada25)
						{
							if(iBonoEntradaCompraCount > 1)
							{
								//Si tiene m?s de 1 Bono activo
								//iBonoEntrada25IdCompraSel = 
								using (frmPrefacturaBonoEntradaSel frmPBES = new frmPrefacturaBonoEntradaSel(iBonoEntrada25IdCliente, idBonoEntradaTipo))
								{				
									if (DialogResult.OK == frmPBES.ShowDialog())
									{
										iBonoEntradaSel = frmPBES.idBonoEntrada;
									}
								}
							}
							else if(iBonoEntradaCompraCount == 1)
							{
								//Si tiene solo 1 Bono activo
								squery = String.Format("Select top 1 idBonoEntrada From BonoEntrada Where idCliente = {0} and idBonoEntradaTipo = {1} And Estado = 1",
									iBonoEntrada25IdCliente,
									idBonoEntradaTipo);
								iBonoEntradaSel = Funcion.iEscalarSQL(cdsSeteoF,squery,false);
							}
							if(iBonoEntradaSel > 0)
							{
								//Actualizar el ValorBono de el bono seleccionado
								string ActualizarValorBono = String.Format("Update BonoEntrada Set ValorBono = {0} Where idBonoEntrada = {1}", getValorBono, iBonoEntradaSel);
								Funcion.EjecutaSQL(cdsSeteo, ActualizarValorBono,false);

								squery = String.Format("Exec BonoEntradaValor {0}", iBonoEntradaSel);
								dBonoEntrada25 = Funcion.decEscalarSQL(cdsSeteoF, squery);
								if(dBonoEntrada25 > 0)
								{
									this.txtCuotaDeEntrada.Value = Convert.ToDecimal(this.txtCuotaDeEntrada.Value) + dBonoEntrada25;
									bActualizarIdPrefactura = true;
								}
							}
						}
					}
					#endregion Bono Entrada 25
				}
			}

			#region Cliente VIP Bloqueado
//			if (bNuevo)
//			{
//				bGuarda = true;
//				decimal dDescuentoVIP = dDsctoReiterativoVIP;//Funcion.decEscalarSQL(cdsSeteoF, string.Format("SELECT COALESCE( ( SELECT PORCENTAJE  FROM CRE_PARAMETROVIP WHERE {0} BETWEEN mONTODESDE AND mONTOHASTA),0) ", (Decimal)this.txtTotal.Value));
//				decimal dValorCuotasAnterior = 0;
//				if (dDsctoReiterativoVIP > 0)// && dDescuentoVIP > 0))
//				{
//					dValorCuotasAnterior = (decimal) this.txtValorCuotas.Value;
//				
//					foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.ultraGrid2.Rows.All)	
//					{
//						if (dr.Cells["idArticulo"].Value != System.DBNull.Value  && Convert.ToInt32( dr.Cells["idArticulo"].Value) != 10717 && Convert.ToInt32( dr.Cells["idArticulo"].Value) != 10718)
//						{
//							//if (Convert.ToInt32( dr.Cells["Tipo"].Value) == 0 && Convert.ToInt32( dr.Cells["idCombo"].Value) == 0 && Convert.ToInt32( dr.Cells["idArticuloGE"].Value) == 0 && Convert.ToInt32( dr.Cells["Principal"].Value) == 0 )
//							//{
//							bVIP = true;
//							if (Convert.ToDecimal(dr.Cells["DescuentoPorc"].Value) < 95.01m)
//								dr.Cells["DescuentoPorc"].Value = Convert.ToDecimal(dr.Cells["DescuentoPorc"].Value) + dDescuentoVIP;
//							//}
//						}
//					}
//				}
//				bGuarda = false;
//			}
			#endregion Cliente VIP Bloqueado

			cmdTotal();
			
			decimal dBaseImponibleCosto = Convert.ToDecimal(this.txtIva0.Value) + 
					Convert.ToDecimal(this.txtIva.Value) - 
					Convert.ToDecimal(this.txtDescIva0Total.Value) - 
					Convert.ToDecimal(this.txtDescIvaTotal.Value);

			if (!Validacion.vCostoCrediPoint(this.ultraGrid2, dBaseImponibleCosto, cdsSeteoF)) return;

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
				(bool)this.chkAutorizar.Checked, Convert.ToDecimal(this.txtValorCuotas.Value), (int)this.txtidEntidadFinanciera.Value, 0, 0, 0, 0, "", false,//(bool)this.chkVerificar.Checked,
				1, "", "", 0, "", 0.00m, 0, "", 0.00m,
				bNuevo, true, idBloqueaTransacciones, this.ultraGrid2,
				false, 0, this.ultraGrid2, bDescuentoActivo);					
			#endregion Grabar

      string NumSol ="";
			if ((int)this.txtNumIdCompra.Value == 0) return;
			/// freelance3
			Funcion.EjecutaSQL(cdsSeteoF, string.Format("Update Compra Set ifreelance = {1} Where idCompra = {0}", (int)this.txtNumIdCompra.Value,Convert.ToInt32( this.chkFreelance.Checked) ));
			if (Funcion.iEscalarSQL(cdsSeteoF, string.Format("Exec ValidaBodegaCrediPointdigital {0}", (int) this.cmbBodega.Value)) > 0)
			{
				int idCre_SolicitudWeb = Funcion.iEscalarSQL(cdsSeteoF, string.Format("Exec [RetornaidSolciitud] {0},'{1}'", (int) this.cmbBodega.Value, this.txtRuc.Text));
				 NumSol = Funcion.sEscalarSQL(cdsSeteoF, string.Format("select numero from Compra where idCompra =  {0} ", (int)this.txtNumIdCompra.Value));
				
				int idTipoClienteCredito = Funcion.iEscalarSQL(cdsSeteoF, string.Format("Exec [RetornatipoClienteWeb] {0},'{1}'", (int) this.cmbBodega.Value, this.txtRuc.Text));
			
				Funcion.EjecutaSQL(cdsSeteoF, string.Format("Update Compra Set idCre_SolicitudWeb = {1}, idTipoClienteCredito = {2} Where idCompra = {0}", (int)this.txtNumIdCompra.Value, idCre_SolicitudWeb, idTipoClienteCredito));

				if(idCre_SolicitudWeb > 0 && bNuevo)
					Funcion.EjecutaSQL(cdsSeteoF, string.Format("Exec [InsertTiempoSolicitudesWeb] {0}, {1}, {2}, '{3}', '{4}'", 15, idCre_SolicitudWeb,1, MenuLatinium.stUsuario, NumSol));
			}
			if (!bPrimeraCuota) Funcion.EjecutaSQL(cdsSeteoF, string.Format("Update Compra Set AbonaCuota = 0 Where idCompra = {0}", (int)this.txtNumIdCompra.Value));
			if (bPrimeraCuota) Funcion.EjecutaSQL(cdsSeteoF, string.Format("Update Compra Set AbonaCuota = 1 Where idCompra = {0}", (int)this.txtNumIdCompra.Value));
			if ((bool) this.chkAnclar.Checked) Funcion.EjecutaSQL(cdsSeteoF, string.Format("Update Compra Set bColateral = 1 Where idCompra = {0}", (int)this.txtNumIdCompra.Value));
			if ((bool) this.chkAnclar.Checked)
			{
				Funcion.EjecutaSQL(cdsSeteoF, string.Format("Exec [InsertaSerieNuovo] {0}, '{1}'", (int)this.txtNumIdCompra.Value, (string)this.txtSerieNuovo.Text));
				if(this.txtSerieNuovo2.Text.Length > 3)
				{
					Funcion.EjecutaSQL(cdsSeteoF, string.Format("Exec [InsertaSerieNuovo] {0}, '{1}', {2}", (int)this.txtNumIdCompra.Value, (string)this.txtSerieNuovo2.Text, 1));
				}
			}

			#region Bono entrada
			if(bActualizarIdPrefactura)
			{
				//Guardar el id de la prefactura generada en bono entrada
				string squery = String.Format("Update BonoEntrada Set idPrefactura = {0} Where idBonoEntrada = {1}", (int)this.txtNumIdCompra.Value, iBonoEntradaSel);
				Funcion.EjecutaSQL(cdsSeteoF, squery); 
			}

			idBonoEntradaTipo = 0;
			//Creo que debo cargar la tabla de Bono entrada tipo e ir realizando un for por cada bono y si se encuentra en vigencia
			dvResultado = Funcion.dvProcedimiento(cdsSeteoF, "Exec BonoEntradaTipoCon");
			if (dvResultado != null && bNuevo)
			{
				foreach (DataRowView rowView in dvResultado)
				{
					DataRow row = rowView.Row;
					idBonoEntradaTipo = Convert.ToInt32(row["idBonoEntradaTipo"]);
					//Actualizar el estado si tiene bono entrada 25
					//Se guarda consulta por el idPrefactura
					//Consultar si existe un registro que tenga la prefactura guardada
					int iBonoEntradaPrefactura = 0;
					string squery = String.Format("Select Count(*) From BonoEntrada Where idPrefactura = {0} And Estado = 1 And idBonoEntradaTipo = {1}", 
						(int)this.txtNumIdCompra.Value,
						idBonoEntradaTipo);
					iBonoEntradaPrefactura = Funcion.iEscalarSQL(cdsSeteoF, squery);
					if(iBonoEntradaPrefactura > 0)
					{
						//Validar si  existe un registro en bono entrada el estado estado = 0, existe el valor de la idprefactura, esto demuestra que ya se uso uno de los dos 
						//bonos por lo que se enviar?a a cambiar el estado de los otros bonos que tiene el cliente si estos estan activos
						//se deber?a actualizar por el idCliente, idCompra, Estado = 1
						BonoEntradaInactivarRelacion(cdsSeteo, (int)this.txtNumIdCompra.Value, idBonoEntradaTipo, 0);
					}
				}
			}
			#endregion Bono entrada

			#region FamilyFriends
			FamilyFriendsCompraPrefactura(FamilyFriendsService.EstadoFFCompraTemporal, (int)this.txtNumIdCompra.Value);
			#endregion FamilyFriends

			#region comentado

			//			int bContAbonaCuota = 0;
			//
			//			if (dEntradaCuota == 0)
			//			{
			//				foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.ultraGrid2.Rows.All)
			//					if (Funcion.bEscalarSQL(cdsSeteoF, string.Format("Select CobroPrimeraCuota From ArticuloSubGrupo sg Inner Join Articulo a On a.idSubGrupo = sg.idSubGrupo Where idArticulo = {0}", (int)dr.Cells["idArticulo"].Value), true))
			//						bContAbonaCuota++;
			//
			//				if (bContAbonaCuota == 1)
			//				{
			//					foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.ultraGrid2.Rows.All)
			//					{
			//						if (Funcion.bEscalarSQL(cdsSeteoF, string.Format("Select CobroPrimeraCuota From ArticuloSubGrupo sg Inner Join Articulo a On a.idSubGrupo = sg.idSubGrupo Where idArticulo = {0}", (int)dr.Cells["idArticulo"].Value), true))
			//						{
			//							int idSubgrupo = Funcion.iEscalarSQL(cdsSeteoF, string.Format("Select idSubGrupo From Articulo Where idArticulo = {0}", (int)dr.Cells["idArticulo"].Value));
			//							decimal dPrecio = Convert.ToDecimal(dr.Cells["Precio"].Value);
			//							bContAbonaCuota = 1;
			//
			//							foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr1 in this.ultraGrid2.Rows.All)
			//							{								
			//								if ((int)dr.Cells["idArticulo"].Value != (int)dr1.Cells["idArticulo"].Value)
			//								{
			//									int idSubgrupo1 = Funcion.iEscalarSQL(cdsSeteoF, string.Format("Select idSubGrupo From Articulo Where idArticulo = {0}", (int)dr1.Cells["idArticulo"].Value));
			//					
			//									if (idSubgrupo != idSubgrupo1)
			//									{
			//										if (Convert.ToDecimal(dr1.Cells["Precio"].Value) > dPrecio)
			//										{
			//											bContAbonaCuota = 0;									
			//										}
			//									}
			//								}
			//							}
			//						}
			//						continue;
			//					}
			//				}				
			//			}
			//			
			//			if (bContAbonaCuota > 0)
			//				Funcion.EjecutaSQL(cdsSeteoF, string.Format("Update Compra Set AbonaCuota = 1 Where idCompra = {0}", (int)this.txtNumIdCompra.Value));	
			#endregion comentado

			if (idMail > 0)
			{
				string sSQLLM = string.Format("Update listacorreosmarketing Set idCompra = {0} Where idCorreosMarketing = {1}", (int)this.txtNumIdCompra.Value, idMail);
				Funcion.EjecutaSQL(cdsSeteoF, sSQLLM);
			}

			Facturacion.ActualizaEstadoCuponWeb(idCuponWeb, 2, (int)this.txtNumIdCompra.Value, cdsSeteoF);



			#region Cliente VIP Bloqueado
//			if (dDsctoReiterativoVIP > 0 && dDescuentoVIP > 0)
//			{
//				decimal cuotaAnterior = Math.Round(txtValorCuotasAnterior, 2);
//				decimal cuotaActual = Math.Round((decimal)this.txtValorCuotas.Value, 2);
//
//				MessageBox.Show(string.Format("Cuota anterior: {0} \nCuota Actual (VIP): {1}", cuotaAnterior, cuotaActual), 
//					"Point Technology", 
//					MessageBoxButtons.OK, 
//					MessageBoxIcon.Information);
//			}
			#endregion Cliente VIP Bloqueado

			idCuponWeb = 0;
			bCuponWeb = false;
			bNuovoHabil = false;
			bVIP = false;			

			#region Impresion
			if ((int)this.txtidEntidadFinanciera.Value != 6)
			{
				if (!(bool) bImpreso)
				{
					decimal dValor = Convert.ToDecimal(this.txtTotal.Value) - Convert.ToDecimal(this.txtCuotaDeEntrada.Value);
					string Fecha = this.cboFechaPrimerPago.Text;
					DocsPrefactura Dctos = new DocsPrefactura ((int)this.txtNumIdCompra.Value, dValor, (int)this.txtNumCuotas.Value, (int)this.txtNumIdCliente.Value, Fecha, 
						(DateTime)this.dtFecha.Value, (int)this.txtidEntidadFinanciera.Value, this.txtRuc.Text.ToString());
					Dctos.ShowDialog();
				}
			}
			else Funcion.EjecutaSQL(cdsSeteoF, string.Format("Update Compra Set Impreso = 1 Where idCompra = {0}", (int)this.txtNumIdCompra.Value));
			#endregion Impresion
					
			#region Controles
			this.cmbBodega.Enabled = false;
			this.cmbBodega1.Enabled = false;
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
			this.txtSerieNuovo.Enabled = false;	
			this.txtSerieNuovo2.Enabled = false;	
			this.chkAnclar.Enabled = false;	

			this.btnValidarMail.Enabled = false;
			
			FuncionesProcedimientos.EstadoGrids(false, this.ultraGrid2);

			bEliminaAlValidar = false;
			bModificaColor = false;

			bEdicion = false;
			bNuevo = false;
					
			bDescuentoActivo = false;
			dClavePorcDscto = 0.00m;

			this.Consulta((int)this.txtNumIdCompra.Value);
			#endregion Controles

			//Encuesta cliente
			ReporteEncuestasOrigenAct();
		}

		private void FamilyFriendsCompraPrefactura(int Estado, int idCompra)
		{
			string nombreClase = this.GetType().Name;
			string nombreMetodo = "FamilyFriendsCompra_Prefactura";

			bool debeGuardar = true;
			bool estadoNeutro = (!estadoFF.EsFF && !estadoFF.RechazoPromocion);

			// =====================================================
			// 0) VALIDACIÓN PREVIA (ESTADO DE PANTALLA)
			// =====================================================
			if (estadoNeutro)
			{
				Funcion.LogSistema(frmRequerimiento.sconexionInicio,
					nombreClase + ": " + nombreMetodo,
					"Estado neutro FF en pantalla, se reconsulta BD",
					nombreMetodo,
					Funcion.slogtipoinformacion
					);

				// Re-consulta a BD para confirmar estado real
				ConsultarEstadoFamilyFriends(idCompra);

				Funcion.LogSistema(frmRequerimiento.sconexionInicio,
					nombreClase + ": " + nombreMetodo,
					string.Format(
					"Post reconsulta FF | EsFF: {0}, RechazoPromocion: {1}, IdWEB_FF_Invitado: {2}",
					estadoFF.EsFF,
					estadoFF.RechazoPromocion,
					estadoFF.IdWEB_FF_Invitado
					),
					nombreMetodo,
					Funcion.slogtipoinformacion
					);

				// Recalcular estado tras reconsulta
				estadoNeutro = (!estadoFF.EsFF && !estadoFF.RechazoPromocion);
			}

			// =====================================================
			// 1) VALIDACIONES DEFINITIVAS
			// =====================================================
			if (estadoNeutro)
			{
				Funcion.LogSistema(frmRequerimiento.sconexionInicio,
					nombreClase + ": " + nombreMetodo,
					"Estado neutro FF confirmado, no se guarda Prefactura",
					nombreMetodo,
					Funcion.slogtipoinformacion
					);

				debeGuardar = false;
			}

			// =====================================================
			// 2) DECISIÓN FINAL
			// =====================================================
			if (!debeGuardar)
			{
				Funcion.LogSistema(frmRequerimiento.sconexionInicio,
					nombreClase + ": " + nombreMetodo,
					"Fin proceso FF Prefactura sin guardar",
					nombreMetodo,
					Funcion.slogtipoinformacion
					);
				return;
			}

			// =====================================================
			// 3) PREPARAR MODELO
			// =====================================================
			FamilyFriendsService service =
				new FamilyFriendsService(frmRequerimiento.sconexionPoint);

			FamilyFriendsCompraModel model = new FamilyFriendsCompraModel();
			model.IdCompra = idCompra;
			model.IdWEB_FF_Invitado = estadoFF.IdWEB_FF_Invitado;
			model.IdTipoPrecio = idTipoPrecio;
			model.IdFormaPago = 9; // Prefactura siempre forma de pago 9
			model.IdCliente = (int)this.txtNumIdCliente.Value;

			model.Estado = estadoFF.RechazoPromocion
				? FamilyFriendsService.EstadoFFCompraNoAcepta
				: Estado;

			// =====================================================
			// 4) GUARDAR
			// =====================================================
			Guid id = service.GuardarCompra(model);

			Funcion.LogSistema(frmRequerimiento.sconexionInicio,
				nombreClase + ": " + nombreMetodo,
				"GuardarCompra_Prefactura OK | idWEB_FF_Compra: " + id,
				nombreMetodo,
				Funcion.slogtipoinformacion
				);
		}

		#region Encuesta origen cliente
		//Actualizaci?n encuesta cliente
		int iEncuesta = 0;
		string sOtrosEncuesta = "";
		private void ReporteEncuestasOrigenAct()
		{
			string squery = String.Format("EXEC ReporteEncuentasOrigenAct {0},{1}", (int)this.txtNumIdCompra.Value, iEncuesta);
			Funcion.EjecutaSQL(cdsSeteoF, squery);
		}
		#endregion Encuesta origen cliente

		private void btnAprobacion_Click(object sender, System.EventArgs e)
		{
			this.MnuAprobacion.Show(this.btnAprobacion, new Point(1, 20));
		}

		private void cmbBodega_KeyUp(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Enter) this.cmbVendedor.Focus();
		}
		
		private void ultraGrid1_DoubleClick(object sender, System.EventArgs e)
		{
			Infragistics.Win.UltraWinGrid.UltraGridRow row = this.ultraGrid2.ActiveRow;

			if (row != null && row.Cells["idArticulo"] != null && row.Cells["idArticulo"].Value != DBNull.Value)
			{
				int idArticulo = Convert.ToInt32(row.Cells["idArticulo"].Value);
				//if (idArticulo == 12920 || idArticulo == 12921) 
				if (idArticulo == 12929 || idArticulo == 12930) 
				{
					FuncionesInventario.AsistenciaSaludDental(this.ultraGrid2, Convert.ToDateTime(this.dtFecha.Value), this.ultraGrid2.Rows.Count - 1, cdsSeteoF, 9, false);
				}
			}
			if (bEdicion)
			{
				bModificaColor = true;
				Facturacion.CambiaColorProducto(this.ultraGrid2, 43, (int)this.cmbBodega.Value, cdsSeteoF);
				bModificaColor = false;				
			}
		}

		private void ultraGrid1_BeforeCellUpdate(object sender, Infragistics.Win.UltraWinGrid.BeforeCellUpdateEventArgs e)
		{		
			if (bEdicion)
			{
				if (e.Cell.Column.ToString() == "Codigo")
					if (!Validacion.bValidaCambioArticuloVenta(e, bModificaColor, false))//miAcceso.BModArtCmbProm 					
						e.Cancel = true;

				if (e.Cell.Column.ToString() == "Precio")
				{
					if (e.Cell.Row.Cells["Tipo"].Value != System.DBNull.Value)
						if ((int)e.Cell.Row.Cells["Tipo"].Value == 5 || (int)e.Cell.Row.Cells["Tipo"].Value == 9 || (int)e.Cell.Row.Cells["Tipo"].Value == 10) e.Cancel = true;				
				
					if (e.Cell.Row.Cells["idArticulo"].Value != System.DBNull.Value)
						if ((int)e.Cell.Row.Cells["Tipo"].Value == 8) e.Cancel = true;
				}

				if (e.Cell.Column.ToString() == "Cantidad")
				{
					if (e.Cell.Row.Cells["idArticulo"].Value != System.DBNull.Value)
						if ((int)e.Cell.Row.Cells["Tipo"].Value == 8) e.Cancel = true;
				}
				//
				if (e.Cell.Column.ToString() == "Impuesto") 
				{
					if (e.Cell.Row.Cells["idArticulo"].Value != System.DBNull.Value)
						if ((int)e.Cell.Row.Cells["idArticulo"].Value == 8) e.Cancel = true;
				}

				if (e.Cell.Column.ToString() == "DescuentoPorc")
				{
					Validacion.ModificaDescuentos(e, bGuarda, bEmpleado);

					if (e.Cell.Row.Cells["idArticulo"].Value != System.DBNull.Value)
						if ((int)e.Cell.Row.Cells["Tipo"].Value == 5 || (int)e.Cell.Row.Cells["idArticulo"].Value == 8 || (int)e.Cell.Row.Cells["Tipo"].Value == 9 || (int)e.Cell.Row.Cells["Tipo"].Value == 10 || (int)e.Cell.Row.Cells["Tipo"].Value == 11) e.Cancel = true;
				}
			}
		}

		private void chkAutorizar_CheckedChanged(object sender, System.EventArgs e)
		{
		
		}

		private void btnClaveDescuento_Click(object sender, System.EventArgs e)
		{
			return;

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

		private void lblIdentificacion_Click(object sender, System.EventArgs e)
		{
		
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
		}

		private void txtCuotaDeEntrada_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Delete) e.Handled = true;
		}

		private void ultraGrid1_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
			if (ultraGrid2.ActiveCell.Column.Key.ToString().ToUpper() == "CANTIDAD")
				Funcion.CamposNumericos(sender, e);
		}

		private void cmbBodega_InitializeLayout(object sender, Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs e)
		{
		
		}

		int idMail = 0;
		private void btnValidarMail_Click(object sender, System.EventArgs e)
		{
			idMail = Facturacion.iPromocionMail(cdsSeteoF);
		}

		private void dtFecha_KeyDown_1(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Delete) e.Handled = true;
			if (e.KeyCode == Keys.Enter) this.cmbBodega.Focus();if (e.KeyCode == Keys.Delete) e.Handled = true;
			if (e.KeyCode == Keys.Enter) this.cmbBodega.Focus();
		}

		private void dtFecha_KeyPress_1(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
			
			if (e.KeyChar == 8) e.Handled = true;			
	
		}

		private void dtFecha_Validated_1(object sender, System.EventArgs e)
		{
			if (bEdicion)	
			{				
				if (this.cmbVendedor.ActiveRow != null) this.cmbVendedor.Value = System.DBNull.Value;
				if (this.cmbBodega.ActiveRow != null) 
					this.cmbVendedor.DataSource = FuncionesProcedimientos.Lista_Vendedores(Convert.ToDateTime(this.dtFecha.Value), (int)this.cmbBodega.Value, cdsSeteoF);

				CargaFecha();
			}
		}

		private void dtFecha_Validating_1(object sender, System.ComponentModel.CancelEventArgs e)
		{
			
			if (bEdicion)
				if (!Validacion.vbFechaEnDocumentos(this.dtFecha, "Ingrese la Fecha", true, "Venta", cdsSeteoF)) e.Cancel = true;
		

		
		}

		private void dtFecha_ValueChanged_1(object sender, System.EventArgs e)
		{
			if (this.dtFecha.Value != System.DBNull.Value)
				FuncionesProcedimientos.ListaBodegas(this.cmbBodega, MenuLatinium.IdUsuario, 43, (DateTime)this.dtFecha.Value, false, cdsSeteoF);
			 
		}

		private void cmbBodega1_ValueChanged(object sender, System.EventArgs e)
		{
			//			if (bEdicion) 
			//			{
			//				if (this.cmbBodega1.ActiveRow != null)
			//				{
			//					 if ((int)this.cmbBodega1.Value > 0)
			//					{
			//						this.cmbBodega1.Enabled = false;
			//					}
			//				}
			//			}
		}

		private void cmbBodega1_Validating(object sender, System.ComponentModel.CancelEventArgs e)
		{
			if (bEdicion) 
			{
				if (this.cmbBodega1.ActiveRow != null)
				{
					if ((int)this.cmbBodega1.Value > 0)
					{
						this.cmbBodega1.Enabled = false;
					}
				}
			}
		
		}

		private void cmbBodega1_Validated(object sender, System.EventArgs e)
		{
			if (bEdicion) 
			{
				if (this.cmbBodega1.ActiveRow != null)
				{
					if ((int)this.cmbBodega1.Value > 0)
					{
						this.cmbBodega1.Enabled = false;
					}
				}
			}
		}

		private void txtObservaciones_ValueChanged(object sender, System.EventArgs e)
		{
		
		}

		private void cmbBodega1_InitializeLayout(object sender, Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs e)
		{
		
		}

		private void MnuAprobacion_Popup(object sender, System.EventArgs e)
		{
		
		}

		private void dtFecha_BeforeDropDown_1(object sender, System.ComponentModel.CancelEventArgs e)
		{
		
		}
		int iAgregaColateral = 0;
		private void chkAnclar_CheckedChanged(object sender, System.EventArgs e)
		{
			if (!bEdicion) return;

			if ((bool) this.chkAnclar.Checked)
			{ 
				//				if (iColateral <=0 && iAgregaColateral <= 1)
				//				{
				this.chkAnclar.Enabled = false;
				this.txtSerieNuovo.Enabled = true;
				iAgregaColateral ++;
				//				}
			}
			else
			{
				this.txtSerieNuovo.Enabled = false;
				iAgregaColateral --;

			}
		}

		private void txtSerieNuovo_Validated(object sender, System.EventArgs e)
		{
			if (bEdicion)			
			{
				string sSQLLM = string.Format("Exec [GuardaCodigoNOVO]   '{0}'", this.txtSerieNuovo.Text);
				Funcion.EjecutaSQL(cdsSeteoF, sSQLLM);
			}
		}
		private void Fundas()
		{
			string sIdArticulos = "";
			int iFila = 0;
			#region Articulos en el Grid
			foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in ultraGrid2.Rows.All)
			{
				if (dr.Cells["idArticulo"].Value != System.DBNull.Value)
				{
					if (sIdArticulos.Length == 0) sIdArticulos = dr.Cells["idArticulo"].Value.ToString();
					else if (sIdArticulos.Length > 0) sIdArticulos = sIdArticulos + ", " + dr.Cells["idArticulo"].Value.ToString();										
				}
			}
			#endregion Articulos en el Grid
			if (sIdArticulos.Length > 0)
			{
				if (Funcion.iEscalarSQL(cdsSeteoF, string.Format("Exec ValidaArticulosFundas '{0}'", sIdArticulos)) > 0)
				{
					using (frmFundasProductos GE = new frmFundasProductos(this.ultraGrid2))
					{
						if (DialogResult.OK == GE.ShowDialog())
						{
							if (GE.ultraGrid1.Rows.Count > 0)
							{
								bFundas=true;
								#region Carga Productos
								foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in GE.ultraGrid1.Rows.All)
								{
									if ((bool)dr.Cells["Sel"].Value)
									{
										int iArticulo =(int)dr.Cells["idArticulo"].Value ;
										#region Carga Fundas	
										decimal dIVA = Funcion.decEscalarSQL(cdsSeteoF, string.Format("Select dbo.RetornaImpuestoIVAVigente('{0}', 1)", 
											Convert.ToDateTime(this.dtFecha.Value).ToString("yyyyMMdd")));
										decimal SPrecio = Funcion.decEscalarSQL(cdsSeteoF, string.Format("Select Precio3 as Precio From Articulo a Inner Join ArticuloSubGrupo c On c.idSubGrupo = a.idSubGrupo  where  idArticulo= {0}", iArticulo));																			
										this.ultraGrid2.DisplayLayout.Bands[0].AddNew();	
										iFila++;
										#region Asignacion Varibles				
										string sSQLInf = string.Format("Select a.idArticulo,a.Codigo,a.Articulo From Articulo a Where idArticulo = {0}", iArticulo);
										SqlDataReader drA = Funcion.rEscalarSQL(cdsSeteoF, sSQLInf, true);
										drA.Read();
										this.ultraGrid2.ActiveRow.Cells["idDetCompra"].Value = 0;
										this.ultraGrid2.ActiveRow.Cells["idOrigen"].Value = 0;
										this.ultraGrid2.ActiveRow.Cells["idUnidad"].Value = 1;			
										this.ultraGrid2.ActiveRow.Cells["idCombo"].Value = 0;
										this.ultraGrid2.ActiveRow.Cells["Tipo"].Value = 0;
										this.ultraGrid2.ActiveRow.Cells["Cantidad"].Value = 1;
										this.ultraGrid2.ActiveRow.Cells["Precio"].Value = SPrecio;
										this.ultraGrid2.ActiveRow.Cells["Codigo"].Value = drA.GetString(1);		
										this.ultraGrid2.ActiveRow.Cells["Articulo"].Value = drA.GetString(2);
										this.ultraGrid2.ActiveRow.Cells["idArticulo"].Value = drA.GetInt32(0);		
										this.ultraGrid2.ActiveRow.Cells["Impuesto"].Value = dIVA;
										this.ultraGrid2.ActiveRow.Cells["DescuentoArt"].Value = 0.00m;
										this.ultraGrid2.ActiveRow.Cells["DescuentoPorc"].Value = 70.00m;			
										this.ultraGrid2.ActiveRow.Cells["Entregado"].Value = false;
										this.ultraGrid2.ActiveRow.Cells["idArticuloGE"].Value = 0;
										this.ultraGrid2.ActiveRow.Cells["Principal"].Value = false;			
										this.ultraGrid2.ActiveRow.Cells["Bloqueado"].Value = false;
										this.ultraGrid2.ActiveRow.Cells["EstadoSerial"].Value = false;
										this.ultraGrid2.ActiveRow.Cells["Posicion"].Value = this.ultraGrid2.Rows.Count;	
										this.ultraGrid2.ActiveRow.Cells["Jornadas"].Value = 0.00m;
										this.ultraGrid2.ActiveRow.Cells["Aporte"].Value = 0.00m;
										this.ultraGrid2.ActiveRow.Cells["idSubProyecto"].Value = 0;

										drA.Close();	
										#endregion Asignacion Varibles				
										#endregion Carga Articulo Fundas
									}
								}
								bFundas=false;
								#endregion Carga Productos
							}
						}
					}
				}
			}
		}

		private void txtSerieNuovo2_Validated(object sender, System.EventArgs e)
		{
			if (bEdicion)			
			{
				string sSQLLM = string.Format("Exec [GuardaCodigoNOVO]   '{0}'", this.txtSerieNuovo2.Text);
				Funcion.EjecutaSQL(cdsSeteoF, sSQLLM);
			}
		}

		private void cdsSeteo_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			cdsSeteoF.Schema.Connections[0].ConnectionString = miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);		
		}

		private void cmbVendedor_ValueChanged(object sender, System.EventArgs e)
		{
			if (cmbVendedor.ActiveRow != null)
			{
				UltraGridRow activeRow = cmbVendedor.ActiveRow;
				// Obtener el valor de la columna 'codigo' de la fila activa
				string codigo = activeRow.Cells["codigo"].Text;
				if ( codigo == "FREELANCE")
				{
					this.chkFreelance.Visible = true;
					this.chkFreelance.Checked = true;
				}
				else
				{
					this.chkFreelance.Visible = false;
					this.chkFreelance.Checked = false;
				}
			}
		}

		private void btnCrediPoint_Click(object sender, System.EventArgs e)
		{
			btnCancelar_Click(sender, e);
			
			using (frmCrediPoint Busqueda = new frmCrediPoint())
			{
				if (DialogResult.OK == Busqueda.ShowDialog())
				{
					//	this.Consulta((int)Busqueda.grdFacturas.ActiveRow.Cells["idCompra"].Value);
				}
			}
		}

		private void txtValorCuotas_ValueChanged(object sender, System.EventArgs e)
		{
		
		}

		private void txtRuc_ValueChanged(object sender, System.EventArgs e)
		{
		
		}
	}
}