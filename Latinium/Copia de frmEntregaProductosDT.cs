using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using DevExpress.XtraEditors;
using System.Data.SqlClient;
using System.IO;
using System.Data;
using System.Text;
using CrystalDecisions;
using CrystalDecisions.CrystalReports.Engine;
using CrystalDecisions.ReportSource;
using CrystalDecisions.Shared;

namespace Latinium
{
	/// <summary>
	/// Descripción breve de frmEntregaProductosDT.
	/// </summary>
	public class frmEntregaProductosDT : System.Windows.Forms.Form
	{
		private System.Windows.Forms.Label label8;
		private C1.Data.C1DataSet cdsSeteoF;
		private System.Windows.Forms.ToolTip toolTip1;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource1;
		private System.Windows.Forms.Label label31;
		private System.Windows.Forms.Label label36;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbBodega;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource2;
		private System.Windows.Forms.Label label14;
		private System.Windows.Forms.Label label11;
		private System.Windows.Forms.Label label1;
		private System.Windows.Forms.Label label2;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtDaño;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtBuscar;
		private System.Windows.Forms.Button btnBuscar;
		private System.Windows.Forms.Label lblContador;
		private System.Windows.Forms.Label label3;
		private System.Windows.Forms.GroupBox groupBox1;
		private System.Windows.Forms.GroupBox groupBox2;
		private System.Windows.Forms.Button btnSalir;
		private System.Windows.Forms.ImageList imageList1;
		private System.Windows.Forms.Button btnCancelar;
		private System.Windows.Forms.Button btnNuevo;
		private System.Windows.Forms.Button btnGuardar;
		private System.Windows.Forms.GroupBox groupBox3;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtFacturar;
		private System.Windows.Forms.Label label4;
		private Infragistics.Win.UltraWinGrid.UltraGrid uGridIngresos;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbUsuario;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource3;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtEstadoFisico;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtInformeTecnico;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtInformeRMA;
		private System.Windows.Forms.Label lblEstado;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtInformeSalida;
		private System.Windows.Forms.Label lblPagoPendiente;
		private System.Windows.Forms.Label label5;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbFormaPago;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource4;
		private System.Windows.Forms.Button btnImprimir;
		private System.ComponentModel.IContainer components;

		public frmEntregaProductosDT()
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
			this.components = new System.ComponentModel.Container();
			Infragistics.Win.Appearance appearance1 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand1 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn1 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idIngreso");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn2 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("BodegaDestino");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn3 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Bodega");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn4 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Ingreso");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn5 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idTecnico");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn6 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Fecha");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn7 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("FechaReparado");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn8 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idFVenta");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn9 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCliente");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn10 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idTipoRuc");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn11 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Ruc");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn12 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn13 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Codigo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn14 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Articulo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn15 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Serial");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn16 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("EstadoIngreso");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn17 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Estado");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn18 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Daño");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn19 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("EstadoFisico");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn20 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("InformeTecnico");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn21 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("InformeRMA");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn22 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("ValorFactura");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn23 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCompra");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn24 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Factura");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn25 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Column 0");
			Infragistics.Win.Appearance appearance2 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance3 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance4 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance5 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance6 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance7 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance8 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn1 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idIngreso");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn2 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("BodegaDestino");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn3 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Bodega");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn4 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Ingreso");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn5 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idTecnico");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn6 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Fecha");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn7 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("FechaReparado");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn8 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idFVenta");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn9 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idCliente");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn10 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idTipoRuc");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn11 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Ruc");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn12 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Nombre");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn13 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Codigo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn14 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Articulo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn15 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Serial");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn16 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("EstadoIngreso");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn17 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Estado");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn18 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Daño");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn19 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("EstadoFisico");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn20 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("InformeTecnico");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn21 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("InformeRMA");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn22 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("ValorFactura");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn23 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idCompra");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn24 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Factura");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn25 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Column 0");
			Infragistics.Win.Appearance appearance9 = new Infragistics.Win.Appearance();
			System.Resources.ResourceManager resources = new System.Resources.ResourceManager(typeof(frmEntregaProductosDT));
			Infragistics.Win.Appearance appearance10 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand2 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn26 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Bodega");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn27 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn28 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Codigo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn26 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Bodega");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn27 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Nombre");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn28 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Codigo");
			Infragistics.Win.Appearance appearance11 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance12 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance13 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance14 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance15 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand3 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn29 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idPersonal");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn30 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn31 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Codigo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn29 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idPersonal");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn30 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Nombre");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn31 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Codigo");
			Infragistics.Win.Appearance appearance16 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand4 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn32 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idFormaPago");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn33 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("FormaPago");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn32 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idFormaPago");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn33 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("FormaPago");
			this.uGridIngresos = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.ultraDataSource1 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.label8 = new System.Windows.Forms.Label();
			this.txtInformeSalida = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.cdsSeteoF = new C1.Data.C1DataSet();
			this.toolTip1 = new System.Windows.Forms.ToolTip(this.components);
			this.btnBuscar = new System.Windows.Forms.Button();
			this.btnNuevo = new System.Windows.Forms.Button();
			this.imageList1 = new System.Windows.Forms.ImageList(this.components);
			this.btnGuardar = new System.Windows.Forms.Button();
			this.btnImprimir = new System.Windows.Forms.Button();
			this.txtFacturar = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.label31 = new System.Windows.Forms.Label();
			this.label36 = new System.Windows.Forms.Label();
			this.cmbBodega = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.ultraDataSource2 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.label14 = new System.Windows.Forms.Label();
			this.txtInformeRMA = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.label11 = new System.Windows.Forms.Label();
			this.txtEstadoFisico = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.label1 = new System.Windows.Forms.Label();
			this.txtInformeTecnico = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.label2 = new System.Windows.Forms.Label();
			this.txtDaño = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.txtBuscar = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.lblContador = new System.Windows.Forms.Label();
			this.label3 = new System.Windows.Forms.Label();
			this.groupBox1 = new System.Windows.Forms.GroupBox();
			this.groupBox2 = new System.Windows.Forms.GroupBox();
			this.btnSalir = new System.Windows.Forms.Button();
			this.btnCancelar = new System.Windows.Forms.Button();
			this.groupBox3 = new System.Windows.Forms.GroupBox();
			this.label4 = new System.Windows.Forms.Label();
			this.cmbUsuario = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.ultraDataSource3 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.lblEstado = new System.Windows.Forms.Label();
			this.lblPagoPendiente = new System.Windows.Forms.Label();
			this.label5 = new System.Windows.Forms.Label();
			this.cmbFormaPago = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.ultraDataSource4 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			((System.ComponentModel.ISupportInitialize)(this.uGridIngresos)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource1)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtInformeSalida)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtFacturar)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbBodega)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource2)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtInformeRMA)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtEstadoFisico)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtInformeTecnico)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtDaño)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtBuscar)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbUsuario)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource3)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbFormaPago)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource4)).BeginInit();
			this.SuspendLayout();
			// 
			// uGridIngresos
			// 
			this.uGridIngresos.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.uGridIngresos.Cursor = System.Windows.Forms.Cursors.Default;
			this.uGridIngresos.DataSource = this.ultraDataSource1;
			appearance1.BackColor = System.Drawing.Color.White;
			this.uGridIngresos.DisplayLayout.Appearance = appearance1;
			ultraGridColumn1.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn1.Header.VisiblePosition = 0;
			ultraGridColumn1.Hidden = true;
			ultraGridColumn2.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn2.Header.VisiblePosition = 1;
			ultraGridColumn2.Hidden = true;
			ultraGridColumn3.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn3.Header.VisiblePosition = 2;
			ultraGridColumn3.Width = 40;
			ultraGridColumn4.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn4.Header.VisiblePosition = 3;
			ultraGridColumn4.Width = 100;
			ultraGridColumn5.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn5.Header.VisiblePosition = 4;
			ultraGridColumn5.Hidden = true;
			ultraGridColumn6.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn6.Header.Caption = "Ingresado";
			ultraGridColumn6.Header.VisiblePosition = 5;
			ultraGridColumn6.Width = 70;
			ultraGridColumn7.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn7.Header.Caption = "Reparado";
			ultraGridColumn7.Header.VisiblePosition = 6;
			ultraGridColumn7.Width = 70;
			ultraGridColumn8.Header.VisiblePosition = 24;
			ultraGridColumn8.Hidden = true;
			ultraGridColumn9.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn9.Header.VisiblePosition = 7;
			ultraGridColumn9.Hidden = true;
			ultraGridColumn10.Header.VisiblePosition = 23;
			ultraGridColumn10.Hidden = true;
			ultraGridColumn11.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn11.Header.VisiblePosition = 8;
			ultraGridColumn11.Hidden = true;
			ultraGridColumn12.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn12.Header.Caption = "Nombre Cliente";
			ultraGridColumn12.Header.VisiblePosition = 9;
			ultraGridColumn12.Width = 250;
			ultraGridColumn13.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn13.Header.VisiblePosition = 10;
			ultraGridColumn13.Width = 180;
			ultraGridColumn14.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn14.Header.VisiblePosition = 11;
			ultraGridColumn14.Hidden = true;
			ultraGridColumn14.Width = 270;
			ultraGridColumn15.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn15.Header.VisiblePosition = 12;
			ultraGridColumn15.Width = 150;
			ultraGridColumn16.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn16.Header.VisiblePosition = 13;
			ultraGridColumn16.Hidden = true;
			ultraGridColumn17.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn17.Header.VisiblePosition = 14;
			ultraGridColumn17.Width = 180;
			ultraGridColumn18.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn18.Header.VisiblePosition = 15;
			ultraGridColumn18.Hidden = true;
			ultraGridColumn18.Width = 250;
			ultraGridColumn19.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn19.Header.VisiblePosition = 16;
			ultraGridColumn19.Hidden = true;
			ultraGridColumn20.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn20.Header.VisiblePosition = 17;
			ultraGridColumn20.Hidden = true;
			ultraGridColumn21.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn21.Header.VisiblePosition = 18;
			ultraGridColumn21.Hidden = true;
			ultraGridColumn22.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn22.Header.VisiblePosition = 19;
			ultraGridColumn22.Hidden = true;
			ultraGridColumn23.Header.VisiblePosition = 20;
			ultraGridColumn23.Hidden = true;
			ultraGridColumn24.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn24.Header.VisiblePosition = 21;
			ultraGridColumn24.Hidden = true;
			ultraGridColumn25.Header.VisiblePosition = 22;
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
																										 ultraGridColumn25});
			this.uGridIngresos.DisplayLayout.BandsSerializer.Add(ultraGridBand1);
			appearance2.ForeColor = System.Drawing.Color.Black;
			appearance2.ForeColorDisabled = System.Drawing.Color.Black;
			this.uGridIngresos.DisplayLayout.Override.ActiveRowAppearance = appearance2;
			this.uGridIngresos.DisplayLayout.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.No;
			this.uGridIngresos.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.False;
			this.uGridIngresos.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.False;
			appearance3.BackColor = System.Drawing.Color.Transparent;
			this.uGridIngresos.DisplayLayout.Override.CardAreaAppearance = appearance3;
			appearance4.ForeColor = System.Drawing.Color.Black;
			appearance4.ForeColorDisabled = System.Drawing.Color.Black;
			this.uGridIngresos.DisplayLayout.Override.GroupByRowAppearance = appearance4;
			appearance5.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance5.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance5.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance5.FontData.BoldAsString = "True";
			appearance5.FontData.Name = "Arial";
			appearance5.FontData.SizeInPoints = 8.5F;
			appearance5.ForeColor = System.Drawing.Color.White;
			appearance5.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.uGridIngresos.DisplayLayout.Override.HeaderAppearance = appearance5;
			this.uGridIngresos.DisplayLayout.Override.HeaderClickAction = Infragistics.Win.UltraWinGrid.HeaderClickAction.SortSingle;
			appearance6.BackColor = System.Drawing.SystemColors.ControlLight;
			appearance6.BackColor2 = System.Drawing.Color.Gainsboro;
			appearance6.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridIngresos.DisplayLayout.Override.RowAlternateAppearance = appearance6;
			appearance7.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance7.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance7.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridIngresos.DisplayLayout.Override.RowSelectorAppearance = appearance7;
			appearance8.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance8.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance8.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridIngresos.DisplayLayout.Override.SelectedRowAppearance = appearance8;
			this.uGridIngresos.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.uGridIngresos.Location = new System.Drawing.Point(8, 48);
			this.uGridIngresos.Name = "uGridIngresos";
			this.uGridIngresos.Size = new System.Drawing.Size(1024, 152);
			this.uGridIngresos.TabIndex = 658;
			this.uGridIngresos.BeforeSelectChange += new Infragistics.Win.UltraWinGrid.BeforeSelectChangeEventHandler(this.uGridIngresos_BeforeSelectChange);
			this.uGridIngresos.InitializeLayout += new Infragistics.Win.UltraWinGrid.InitializeLayoutEventHandler(this.uGridIngresos_InitializeLayout);
			// 
			// ultraDataSource1
			// 
			ultraDataColumn1.DataType = typeof(int);
			ultraDataColumn2.DataType = typeof(int);
			ultraDataColumn5.DataType = typeof(int);
			ultraDataColumn6.DataType = typeof(System.DateTime);
			ultraDataColumn6.DefaultValue = new System.DateTime(((long)(0)));
			ultraDataColumn7.DataType = typeof(System.DateTime);
			ultraDataColumn7.DefaultValue = new System.DateTime(((long)(0)));
			ultraDataColumn8.DataType = typeof(int);
			ultraDataColumn9.DataType = typeof(int);
			ultraDataColumn10.DataType = typeof(int);
			ultraDataColumn16.DataType = typeof(int);
			ultraDataColumn22.DataType = typeof(System.Decimal);
			ultraDataColumn23.DataType = typeof(int);
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
																																 ultraDataColumn25});
			// 
			// label8
			// 
			this.label8.AutoSize = true;
			this.label8.BackColor = System.Drawing.Color.Transparent;
			this.label8.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label8.Location = new System.Drawing.Point(8, 406);
			this.label8.Name = "label8";
			this.label8.Size = new System.Drawing.Size(77, 17);
			this.label8.TabIndex = 681;
			this.label8.Text = "Informe Salida";
			this.label8.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// txtInformeSalida
			// 
			appearance9.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtInformeSalida.Appearance = appearance9;
			this.txtInformeSalida.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtInformeSalida.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtInformeSalida.Enabled = false;
			this.txtInformeSalida.Font = new System.Drawing.Font("Tahoma", 8.25F);
			this.txtInformeSalida.Location = new System.Drawing.Point(104, 392);
			this.txtInformeSalida.Multiline = true;
			this.txtInformeSalida.Name = "txtInformeSalida";
			this.txtInformeSalida.Scrollbars = System.Windows.Forms.ScrollBars.Vertical;
			this.txtInformeSalida.Size = new System.Drawing.Size(416, 44);
			this.txtInformeSalida.TabIndex = 4;
			this.txtInformeSalida.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtInformeSalida_KeyDown);
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
			// btnBuscar
			// 
			this.btnBuscar.CausesValidation = false;
			this.btnBuscar.Image = ((System.Drawing.Image)(resources.GetObject("btnBuscar.Image")));
			this.btnBuscar.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnBuscar.Location = new System.Drawing.Point(568, 7);
			this.btnBuscar.Name = "btnBuscar";
			this.btnBuscar.Size = new System.Drawing.Size(80, 23);
			this.btnBuscar.TabIndex = 2;
			this.btnBuscar.Text = "Buscar";
			this.btnBuscar.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.toolTip1.SetToolTip(this.btnBuscar, "Actualizar información");
			this.btnBuscar.Click += new System.EventHandler(this.btnBuscar_Click);
			// 
			// btnNuevo
			// 
			this.btnNuevo.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btnNuevo.CausesValidation = false;
			this.btnNuevo.Enabled = false;
			this.btnNuevo.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnNuevo.ImageIndex = 50;
			this.btnNuevo.ImageList = this.imageList1;
			this.btnNuevo.Location = new System.Drawing.Point(8, 474);
			this.btnNuevo.Name = "btnNuevo";
			this.btnNuevo.TabIndex = 718;
			this.btnNuevo.Text = "Nuevo";
			this.btnNuevo.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.toolTip1.SetToolTip(this.btnNuevo, "Nuevo informe técnico");
			this.btnNuevo.Click += new System.EventHandler(this.btnNuevo_Click);
			// 
			// imageList1
			// 
			this.imageList1.ImageSize = new System.Drawing.Size(16, 16);
			this.imageList1.ImageStream = ((System.Windows.Forms.ImageListStreamer)(resources.GetObject("imageList1.ImageStream")));
			this.imageList1.TransparentColor = System.Drawing.Color.Transparent;
			// 
			// btnGuardar
			// 
			this.btnGuardar.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btnGuardar.CausesValidation = false;
			this.btnGuardar.Enabled = false;
			this.btnGuardar.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnGuardar.ImageIndex = 9;
			this.btnGuardar.ImageList = this.imageList1;
			this.btnGuardar.Location = new System.Drawing.Point(87, 474);
			this.btnGuardar.Name = "btnGuardar";
			this.btnGuardar.TabIndex = 716;
			this.btnGuardar.Text = "Guardar";
			this.btnGuardar.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.toolTip1.SetToolTip(this.btnGuardar, "Grabar Información");
			this.btnGuardar.Click += new System.EventHandler(this.btnGuardar_Click);
			// 
			// btnImprimir
			// 
			this.btnImprimir.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btnImprimir.CausesValidation = false;
			this.btnImprimir.Enabled = false;
			this.btnImprimir.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnImprimir.ImageIndex = 6;
			this.btnImprimir.ImageList = this.imageList1;
			this.btnImprimir.Location = new System.Drawing.Point(166, 474);
			this.btnImprimir.Name = "btnImprimir";
			this.btnImprimir.Size = new System.Drawing.Size(77, 23);
			this.btnImprimir.TabIndex = 729;
			this.btnImprimir.Text = "Imprimir";
			this.btnImprimir.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.toolTip1.SetToolTip(this.btnImprimir, "Reeimprimir N/E Proveedores...");
			this.btnImprimir.Click += new System.EventHandler(this.btnImprimir_Click);
			// 
			// txtFacturar
			// 
			appearance10.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtFacturar.Appearance = appearance10;
			this.txtFacturar.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtFacturar.Enabled = false;
			this.txtFacturar.Font = new System.Drawing.Font("Tahoma", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.txtFacturar.FormatString = "#,##0.00";
			this.txtFacturar.Location = new System.Drawing.Point(912, 368);
			this.txtFacturar.Name = "txtFacturar";
			this.txtFacturar.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtFacturar.PromptChar = ' ';
			this.txtFacturar.Size = new System.Drawing.Size(120, 29);
			this.txtFacturar.TabIndex = 690;
			// 
			// label31
			// 
			this.label31.AutoSize = true;
			this.label31.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label31.Location = new System.Drawing.Point(8, 370);
			this.label31.Name = "label31";
			this.label31.Size = new System.Drawing.Size(43, 17);
			this.label31.TabIndex = 693;
			this.label31.Text = "Entrega";
			this.label31.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// label36
			// 
			this.label36.AutoSize = true;
			this.label36.Location = new System.Drawing.Point(8, 10);
			this.label36.Name = "label36";
			this.label36.Size = new System.Drawing.Size(43, 16);
			this.label36.TabIndex = 696;
			this.label36.Text = "Bodega";
			this.label36.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
			// 
			// cmbBodega
			// 
			this.cmbBodega.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbBodega.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbBodega.DataSource = this.ultraDataSource2;
			ultraGridColumn26.Header.VisiblePosition = 0;
			ultraGridColumn26.Hidden = true;
			ultraGridColumn27.Header.VisiblePosition = 1;
			ultraGridColumn27.Width = 250;
			ultraGridColumn28.Header.VisiblePosition = 2;
			ultraGridColumn28.Hidden = true;
			ultraGridBand2.Columns.AddRange(new object[] {
																										 ultraGridColumn26,
																										 ultraGridColumn27,
																										 ultraGridColumn28});
			this.cmbBodega.DisplayLayout.BandsSerializer.Add(ultraGridBand2);
			this.cmbBodega.DisplayMember = "Nombre";
			this.cmbBodega.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbBodega.Enabled = false;
			this.cmbBodega.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbBodega.Location = new System.Drawing.Point(72, 8);
			this.cmbBodega.MaxDropDownItems = 30;
			this.cmbBodega.Name = "cmbBodega";
			this.cmbBodega.Size = new System.Drawing.Size(200, 21);
			this.cmbBodega.TabIndex = 0;
			this.cmbBodega.ValueMember = "Bodega";
			this.cmbBodega.KeyDown += new System.Windows.Forms.KeyEventHandler(this.cmbBodega_KeyDown);
			this.cmbBodega.ValueChanged += new System.EventHandler(this.cmbBodega_ValueChanged);
			this.cmbBodega.InitializeLayout += new Infragistics.Win.UltraWinGrid.InitializeLayoutEventHandler(this.cmbBodega_InitializeLayout);
			// 
			// ultraDataSource2
			// 
			this.ultraDataSource2.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn26,
																																 ultraDataColumn27,
																																 ultraDataColumn28});
			// 
			// label14
			// 
			this.label14.AutoSize = true;
			this.label14.BackColor = System.Drawing.Color.Transparent;
			this.label14.Font = new System.Drawing.Font("Tahoma", 8.25F);
			this.label14.Location = new System.Drawing.Point(536, 305);
			this.label14.Name = "label14";
			this.label14.Size = new System.Drawing.Size(70, 17);
			this.label14.TabIndex = 709;
			this.label14.Text = "Informe RMA";
			this.label14.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// txtInformeRMA
			// 
			appearance11.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtInformeRMA.Appearance = appearance11;
			this.txtInformeRMA.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtInformeRMA.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtInformeRMA.Enabled = false;
			this.txtInformeRMA.Font = new System.Drawing.Font("Tahoma", 8.25F);
			this.txtInformeRMA.Location = new System.Drawing.Point(616, 280);
			this.txtInformeRMA.Multiline = true;
			this.txtInformeRMA.Name = "txtInformeRMA";
			this.txtInformeRMA.Scrollbars = System.Windows.Forms.ScrollBars.Vertical;
			this.txtInformeRMA.Size = new System.Drawing.Size(416, 66);
			this.txtInformeRMA.TabIndex = 708;
			// 
			// label11
			// 
			this.label11.AutoSize = true;
			this.label11.BackColor = System.Drawing.Color.Transparent;
			this.label11.Font = new System.Drawing.Font("Tahoma", 8.25F);
			this.label11.Location = new System.Drawing.Point(536, 233);
			this.label11.Name = "label11";
			this.label11.Size = new System.Drawing.Size(69, 17);
			this.label11.TabIndex = 707;
			this.label11.Text = "Estado Fisico";
			this.label11.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// txtEstadoFisico
			// 
			appearance12.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtEstadoFisico.Appearance = appearance12;
			this.txtEstadoFisico.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtEstadoFisico.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtEstadoFisico.Enabled = false;
			this.txtEstadoFisico.Font = new System.Drawing.Font("Tahoma", 8.25F);
			this.txtEstadoFisico.Location = new System.Drawing.Point(616, 208);
			this.txtEstadoFisico.Multiline = true;
			this.txtEstadoFisico.Name = "txtEstadoFisico";
			this.txtEstadoFisico.Scrollbars = System.Windows.Forms.ScrollBars.Vertical;
			this.txtEstadoFisico.Size = new System.Drawing.Size(416, 66);
			this.txtEstadoFisico.TabIndex = 706;
			// 
			// label1
			// 
			this.label1.AutoSize = true;
			this.label1.BackColor = System.Drawing.Color.Transparent;
			this.label1.Font = new System.Drawing.Font("Tahoma", 8.25F);
			this.label1.Location = new System.Drawing.Point(8, 305);
			this.label1.Name = "label1";
			this.label1.Size = new System.Drawing.Size(86, 17);
			this.label1.TabIndex = 705;
			this.label1.Text = "Informe Tecnico";
			this.label1.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// txtInformeTecnico
			// 
			appearance13.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtInformeTecnico.Appearance = appearance13;
			this.txtInformeTecnico.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtInformeTecnico.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtInformeTecnico.Enabled = false;
			this.txtInformeTecnico.Font = new System.Drawing.Font("Tahoma", 8.25F);
			this.txtInformeTecnico.Location = new System.Drawing.Point(104, 280);
			this.txtInformeTecnico.Multiline = true;
			this.txtInformeTecnico.Name = "txtInformeTecnico";
			this.txtInformeTecnico.Scrollbars = System.Windows.Forms.ScrollBars.Vertical;
			this.txtInformeTecnico.Size = new System.Drawing.Size(416, 66);
			this.txtInformeTecnico.TabIndex = 704;
			// 
			// label2
			// 
			this.label2.AutoSize = true;
			this.label2.BackColor = System.Drawing.Color.Transparent;
			this.label2.Font = new System.Drawing.Font("Tahoma", 8.25F);
			this.label2.Location = new System.Drawing.Point(8, 233);
			this.label2.Name = "label2";
			this.label2.Size = new System.Drawing.Size(85, 17);
			this.label2.TabIndex = 703;
			this.label2.Text = "Daño Reportado";
			this.label2.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// txtDaño
			// 
			appearance14.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtDaño.Appearance = appearance14;
			this.txtDaño.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtDaño.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtDaño.Enabled = false;
			this.txtDaño.Font = new System.Drawing.Font("Tahoma", 8.25F);
			this.txtDaño.Location = new System.Drawing.Point(104, 208);
			this.txtDaño.Multiline = true;
			this.txtDaño.Name = "txtDaño";
			this.txtDaño.Scrollbars = System.Windows.Forms.ScrollBars.Vertical;
			this.txtDaño.Size = new System.Drawing.Size(416, 66);
			this.txtDaño.TabIndex = 702;
			// 
			// txtBuscar
			// 
			appearance15.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtBuscar.Appearance = appearance15;
			this.txtBuscar.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtBuscar.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtBuscar.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.txtBuscar.Location = new System.Drawing.Point(352, 7);
			this.txtBuscar.MaxLength = 100;
			this.txtBuscar.Name = "txtBuscar";
			this.txtBuscar.Size = new System.Drawing.Size(200, 22);
			this.txtBuscar.TabIndex = 1;
			this.txtBuscar.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtBuscar_KeyDown);
			// 
			// lblContador
			// 
			this.lblContador.AutoSize = true;
			this.lblContador.Font = new System.Drawing.Font("Tahoma", 9F);
			this.lblContador.ForeColor = System.Drawing.Color.Firebrick;
			this.lblContador.Location = new System.Drawing.Point(680, 9);
			this.lblContador.Name = "lblContador";
			this.lblContador.Size = new System.Drawing.Size(0, 18);
			this.lblContador.TabIndex = 713;
			this.lblContador.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
			// 
			// label3
			// 
			this.label3.AutoSize = true;
			this.label3.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label3.Location = new System.Drawing.Point(296, 10);
			this.label3.Name = "label3";
			this.label3.Size = new System.Drawing.Size(37, 17);
			this.label3.TabIndex = 711;
			this.label3.Text = "Buscar";
			this.label3.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// groupBox1
			// 
			this.groupBox1.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.groupBox1.Location = new System.Drawing.Point(0, 32);
			this.groupBox1.Name = "groupBox1";
			this.groupBox1.Size = new System.Drawing.Size(1040, 8);
			this.groupBox1.TabIndex = 714;
			this.groupBox1.TabStop = false;
			// 
			// groupBox2
			// 
			this.groupBox2.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.groupBox2.Location = new System.Drawing.Point(0, 352);
			this.groupBox2.Name = "groupBox2";
			this.groupBox2.Size = new System.Drawing.Size(1040, 8);
			this.groupBox2.TabIndex = 715;
			this.groupBox2.TabStop = false;
			// 
			// btnSalir
			// 
			this.btnSalir.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btnSalir.CausesValidation = false;
			this.btnSalir.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnSalir.ImageIndex = 62;
			this.btnSalir.ImageList = this.imageList1;
			this.btnSalir.Location = new System.Drawing.Point(326, 474);
			this.btnSalir.Name = "btnSalir";
			this.btnSalir.TabIndex = 720;
			this.btnSalir.Text = "&Salir";
			this.btnSalir.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnSalir.Click += new System.EventHandler(this.btnSalir_Click);
			// 
			// btnCancelar
			// 
			this.btnCancelar.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btnCancelar.CausesValidation = false;
			this.btnCancelar.Enabled = false;
			this.btnCancelar.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnCancelar.ImageIndex = 10;
			this.btnCancelar.ImageList = this.imageList1;
			this.btnCancelar.Location = new System.Drawing.Point(247, 474);
			this.btnCancelar.Name = "btnCancelar";
			this.btnCancelar.TabIndex = 719;
			this.btnCancelar.Text = "&Cancelar";
			this.btnCancelar.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnCancelar.Click += new System.EventHandler(this.btnCancelar_Click);
			// 
			// groupBox3
			// 
			this.groupBox3.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.groupBox3.Location = new System.Drawing.Point(0, 464);
			this.groupBox3.Name = "groupBox3";
			this.groupBox3.Size = new System.Drawing.Size(1040, 8);
			this.groupBox3.TabIndex = 721;
			this.groupBox3.TabStop = false;
			// 
			// label4
			// 
			this.label4.AutoSize = true;
			this.label4.Location = new System.Drawing.Point(776, 374);
			this.label4.Name = "label4";
			this.label4.Size = new System.Drawing.Size(122, 16);
			this.label4.TabIndex = 722;
			this.label4.Text = "Valor Factura Sin I.V.A.";
			this.label4.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// cmbUsuario
			// 
			this.cmbUsuario.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbUsuario.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbUsuario.DataSource = this.ultraDataSource3;
			ultraGridColumn29.Header.VisiblePosition = 0;
			ultraGridColumn29.Hidden = true;
			ultraGridColumn30.Header.VisiblePosition = 1;
			ultraGridColumn30.Width = 250;
			ultraGridColumn31.Header.VisiblePosition = 2;
			ultraGridColumn31.Hidden = true;
			ultraGridBand3.Columns.AddRange(new object[] {
																										 ultraGridColumn29,
																										 ultraGridColumn30,
																										 ultraGridColumn31});
			this.cmbUsuario.DisplayLayout.BandsSerializer.Add(ultraGridBand3);
			this.cmbUsuario.DisplayMember = "Nombre";
			this.cmbUsuario.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbUsuario.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbUsuario.Enabled = false;
			this.cmbUsuario.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbUsuario.Location = new System.Drawing.Point(104, 368);
			this.cmbUsuario.MaxDropDownItems = 30;
			this.cmbUsuario.Name = "cmbUsuario";
			this.cmbUsuario.Size = new System.Drawing.Size(250, 21);
			this.cmbUsuario.TabIndex = 3;
			this.cmbUsuario.ValueMember = "idPersonal";
			this.cmbUsuario.KeyDown += new System.Windows.Forms.KeyEventHandler(this.cmbUsuario_KeyDown);
			// 
			// ultraDataSource3
			// 
			ultraDataColumn29.DataType = typeof(int);
			this.ultraDataSource3.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn29,
																																 ultraDataColumn30,
																																 ultraDataColumn31});
			// 
			// lblEstado
			// 
			this.lblEstado.AutoSize = true;
			this.lblEstado.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.lblEstado.ForeColor = System.Drawing.Color.Firebrick;
			this.lblEstado.Location = new System.Drawing.Point(536, 442);
			this.lblEstado.Name = "lblEstado";
			this.lblEstado.Size = new System.Drawing.Size(0, 19);
			this.lblEstado.TabIndex = 724;
			this.lblEstado.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
			// 
			// lblPagoPendiente
			// 
			this.lblPagoPendiente.AutoSize = true;
			this.lblPagoPendiente.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.lblPagoPendiente.ForeColor = System.Drawing.Color.Firebrick;
			this.lblPagoPendiente.Location = new System.Drawing.Point(528, 417);
			this.lblPagoPendiente.Name = "lblPagoPendiente";
			this.lblPagoPendiente.Size = new System.Drawing.Size(0, 19);
			this.lblPagoPendiente.TabIndex = 725;
			this.lblPagoPendiente.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
			// 
			// label5
			// 
			this.label5.AutoSize = true;
			this.label5.Location = new System.Drawing.Point(8, 443);
			this.label5.Name = "label5";
			this.label5.Size = new System.Drawing.Size(82, 16);
			this.label5.TabIndex = 727;
			this.label5.Text = "Forma de Pago";
			this.label5.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// cmbFormaPago
			// 
			appearance16.FontData.Name = "Tahoma";
			appearance16.ForeColorDisabled = System.Drawing.Color.Black;
			this.cmbFormaPago.Appearance = appearance16;
			this.cmbFormaPago.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbFormaPago.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbFormaPago.DataSource = this.ultraDataSource4;
			ultraGridColumn32.Header.VisiblePosition = 0;
			ultraGridColumn32.Hidden = true;
			ultraGridColumn33.Header.VisiblePosition = 1;
			ultraGridColumn33.Width = 248;
			ultraGridBand4.Columns.AddRange(new object[] {
																										 ultraGridColumn32,
																										 ultraGridColumn33});
			this.cmbFormaPago.DisplayLayout.BandsSerializer.Add(ultraGridBand4);
			this.cmbFormaPago.DisplayMember = "FormaPago";
			this.cmbFormaPago.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbFormaPago.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbFormaPago.Enabled = false;
			this.cmbFormaPago.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbFormaPago.Location = new System.Drawing.Point(104, 440);
			this.cmbFormaPago.Name = "cmbFormaPago";
			this.cmbFormaPago.Size = new System.Drawing.Size(248, 22);
			this.cmbFormaPago.TabIndex = 728;
			this.cmbFormaPago.ValueMember = "idFormaPago";
			this.cmbFormaPago.KeyDown += new System.Windows.Forms.KeyEventHandler(this.cmbFormaPago_KeyDown);
			// 
			// ultraDataSource4
			// 
			ultraDataColumn32.DataType = typeof(int);
			this.ultraDataSource4.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn32,
																																 ultraDataColumn33});
			// 
			// frmEntregaProductosDT
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 13);
			this.ClientSize = new System.Drawing.Size(1040, 504);
			this.Controls.Add(this.btnImprimir);
			this.Controls.Add(this.cmbFormaPago);
			this.Controls.Add(this.label5);
			this.Controls.Add(this.label4);
			this.Controls.Add(this.label3);
			this.Controls.Add(this.label14);
			this.Controls.Add(this.label11);
			this.Controls.Add(this.label1);
			this.Controls.Add(this.label2);
			this.Controls.Add(this.label36);
			this.Controls.Add(this.label31);
			this.Controls.Add(this.label8);
			this.Controls.Add(this.lblPagoPendiente);
			this.Controls.Add(this.lblEstado);
			this.Controls.Add(this.cmbUsuario);
			this.Controls.Add(this.groupBox3);
			this.Controls.Add(this.btnCancelar);
			this.Controls.Add(this.btnNuevo);
			this.Controls.Add(this.btnGuardar);
			this.Controls.Add(this.groupBox2);
			this.Controls.Add(this.groupBox1);
			this.Controls.Add(this.txtBuscar);
			this.Controls.Add(this.btnBuscar);
			this.Controls.Add(this.lblContador);
			this.Controls.Add(this.txtInformeRMA);
			this.Controls.Add(this.txtEstadoFisico);
			this.Controls.Add(this.txtInformeTecnico);
			this.Controls.Add(this.txtDaño);
			this.Controls.Add(this.cmbBodega);
			this.Controls.Add(this.txtFacturar);
			this.Controls.Add(this.txtInformeSalida);
			this.Controls.Add(this.uGridIngresos);
			this.Controls.Add(this.btnSalir);
			this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog;
			this.KeyPreview = true;
			this.MaximizeBox = false;
			this.Name = "frmEntregaProductosDT";
			this.Text = "Entrega de Articulos Servicio Tecnico";
			this.KeyDown += new System.Windows.Forms.KeyEventHandler(this.frmEntregaProductosDT_KeyDown);
			this.Closing += new System.ComponentModel.CancelEventHandler(this.frmEntregaProductosDT_Closing);
			this.Load += new System.EventHandler(this.frmEntregaProductosDT_Load);
			((System.ComponentModel.ISupportInitialize)(this.uGridIngresos)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource1)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtInformeSalida)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtFacturar)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbBodega)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource2)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtInformeRMA)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtEstadoFisico)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtInformeTecnico)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtDaño)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtBuscar)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbUsuario)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource3)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbFormaPago)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource4)).EndInit();
			this.ResumeLayout(false);

		}
		#endregion
			
		int iBodegaPredefinida = 0;
		int idIngreso = 0;
		int idPersonal = 0;
		int idCliente = 0;
		int idCompra = 0;
		string sRUC = "";
		string sCliente = "";
		bool bEdicion = false;
		
		private void cdsSeteoF_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			cdsSeteoF.Schema.Connections[0].ConnectionString = miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);
		}

		private void frmEntregaProductosDT_Load(object sender, System.EventArgs e)
		{		
			this.cmbBodega.DataSource = Funcion.dvProcedimiento(cdsSeteoF, string.Format("Exec ListaBodegas {0}, 50", MenuLatinium.IdUsuario));

			this.cmbFormaPago.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "Select idFormaPago, FormaPago From CompraFormaPago Where idFormaPago In (1, 3, 11)");
			idPersonal = Funcion.iEscalarSQL(cdsSeteoF, string.Format("If Exists(Select idPersonal From Nomina Where Codigo In (Select Nombre From Usuario Where idUsuario = {0})) Select idPersonal From Nomina Where Codigo In (Select Nombre From Usuario Where idUsuario = {0}) Else Select 0", MenuLatinium.IdUsuario));

			string sSQL = string.Format("Exec BodegaAsignadaPorUsuario {0}", MenuLatinium.IdUsuario);		
			iBodegaPredefinida = Funcion.iEscalarSQL(cdsSeteoF, sSQL);

			if (iBodegaPredefinida > 0) this.cmbBodega.Value = iBodegaPredefinida;
			else 
			{
				this.cmbBodega.Value = 20;			
				this.cmbBodega.Enabled = true;
			}
		}

		private bool CreaFactura()
		{			
			#region Variables
			string sAutorizacion = "";			
			string sSerie = "";
			DateTime dtFechaCaducidad = DateTime.Now;			
						
			int idCliente = (int)this.uGridIngresos.ActiveRow.Cells["idCliente"].Value;
			string sRUC = this.uGridIngresos.ActiveRow.Cells["Ruc"].Value.ToString().Trim();
			string sNombre = this.uGridIngresos.ActiveRow.Cells["Nombre"].Value.ToString().Trim();
			
			DateTime dtFecha = DateTime.Now;
			int iBodega = 1;
			if ((int)this.cmbBodega.Value != 20) iBodega = (int)this.cmbBodega.Value;
			int idVendedor = (int)this.uGridIngresos.ActiveRow.Cells["idTecnico"].Value;
			
			int idProyecto = Funcion.iEscalarSQL(cdsSeteoF, string.Format("Select isNull(idProyecto, 0) From Bodega Where Bodega = {0}", iBodega));			
			int idTipoRuc = Funcion.iEscalarSQL(cdsSeteoF, string.Format("Select isNull(idTiporuc, 0) From Cliente Where idCliente = {0}", idCliente));

			decimal dPorcIva = Funcion.decEscalarSQL(cdsSeteoF, string.Format("Select dbo.RetornaImpuestoIVAVigente('{0}', 1)", dtFecha.ToString("yyyyMMdd")));

			SqlDataReader dr = Funcion.rEscalarSQL(cdsSeteoF, string.Format("Select Serie, Autorizacion, Caducidad From Bodega Where Bodega = {0}", iBodega), true);	
			dr.Read();
			if (dr.HasRows)
			{
				sSerie = dr.GetValue(0).ToString();
				sAutorizacion = dr.GetValue(1).ToString();
				dtFechaCaducidad = dr.GetDateTime(2);
			}
			dr.Close();	
			#endregion Variables
			
			#region Fechas
			DateTime dtFechaCad = DateTime.Today;
			DateTime dtFechaEnt = DateTime.Today;
			DateTime dtHoraEntrega = DateTime.Today;
			DateTime dtFechPrimPago = DateTime.Today;			
			#endregion Fechas

			#region Graba Factura
			using (SqlConnection oConexion = new SqlConnection("Password=Bl45o6$9;Persist Security Info=True;User ID= " + MenuLatinium.stUsuario + ";Initial Catalog=" + MenuLatinium.stDirFacturacion + ";Data Source=" + MenuLatinium.stDirServidor))
			{
				oConexion.Open();

				SqlTransaction oTransaction = null;
				SqlCommand oCmdActualiza = oConexion.CreateCommand();
				oCmdActualiza.Connection = oConexion;
				oCmdActualiza.CommandType = CommandType.Text;
				oCmdActualiza.CommandTimeout = 0;
						
				try
				{							
					oTransaction = oConexion.BeginTransaction();
					oCmdActualiza.Transaction = oTransaction;

					#region Maestro
					string sSQL = string.Format("Exec GrabaMaestroFactura {0}, {1}, {2}, {3}, {4}, {5}, {6}, {7}, {8}, {9}, {10}, {11}, '{12}', '{13}', {14}, {15}, '{16}', '{17}', '{18}', '{19}', '{20}', '{21}', {22}, '{23}', '{24}', {25}, '{26}', {27}, {28}, {29}, {30}, '{31}', {32}, {33}, {34}, {35}, {36}, {37}, {38}, {39}, {40}, {41}, {42}, {43}, {44}, {45}, '{46}', {47}, {48}, {49}, {50}, {51}, {52}, {53}, {54}, {55}, {56}, {57}, {58}, '{59}', {60}, {61}, '{62}', '{63}', {64}, '{65}', {66}, {67}, '{68}', {69}", 
						idCompra, 0, 0, 1, 15, 6, false, 1, (int)this.cmbFormaPago.Value, iBodega, idProyecto, idVendedor, 
						"", this.uGridIngresos.ActiveRow.Cells["Ingreso"].Value.ToString(), idTipoRuc, idCliente, sRUC, sNombre, dtFecha.ToString("yyyyMMdd"), 
						sSerie, sAutorizacion, dtFechaCaducidad.ToString("yyyyMMdd"), false, dtFechaEnt.ToString("yyyyMMdd"), dtHoraEntrega.ToString("yyyyMMdd hh:mm"), 0, "FACTURA GENERADA POR SERVICIO TECNICO",
						false, 0, 0, 0, dtFechPrimPago.ToString("yyyyMMdd"), 
						0, 0, (decimal)this.txtFacturar.Value, 0, ((decimal)this.txtFacturar.Value * dPorcIva)/100.00m, (decimal)this.txtFacturar.Value + ((decimal)this.txtFacturar.Value * dPorcIva)/100.00m,	
						false, false, false, false, false, true, iBodega, 0, "", 0, 0, 0, 0, 0, false, 0, 0, 0, 0, 0, 0, "", false, 0, "", "", 0, "", 0, 0, "", 0, 0, "");
					oCmdActualiza.CommandText = sSQL;
					idCompra = (int)oCmdActualiza.ExecuteScalar();
					#endregion Maestro

					#region Detalle
					string sSQLDetalle = string.Format("Exec GuardaDetalleFactura {0}, {1}, {2}, {3}, {4}, '{5}', '{6}', {7}, {8}, {9}, {10}, {11}, {12}, {13}, {14}, {15}, {16}, {17}, {18}, {19}, {20}, {21}",
						0, 0, idCompra, 1884, 0, "SERVICIO TECNICO", "MANO DE OBRA SERVICIO TECNICO", 1, (decimal)this.txtFacturar.Value, dPorcIva, 0, 0, (decimal)this.txtFacturar.Value,
						false, 1, false, 1, true, true, 0, 0, 0);
					oCmdActualiza.CommandText = sSQLDetalle;
					oCmdActualiza.ExecuteNonQuery();
					#endregion Detalle	
						
					oTransaction.Commit();
					
					return true;
				}
				catch (Exception ex)
				{
					oTransaction.Rollback();
					MessageBox.Show(string.Format("Commit Exception Type : {0} {1}", ex.GetType(), ex.Message), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);						
					return false;
				}
				finally
				{
					oConexion.Close();					

					if ((int)this.cmbBodega.Value != 20)
					{
						frmVenta Factura = new frmVenta(idCompra);
						Factura.Show();
					}
					else MessageBox.Show("Factura generada correctamente en Estado Temporal", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Information);					
				}
			}			
			#endregion Graba Factura				
		}

		private void cmbBodega_ValueChanged(object sender, System.EventArgs e)
		{
			int iBodega = 0;

			if (this.cmbBodega.ActiveRow != null) iBodega = (int)this.cmbBodega.Value;

			this.cmbUsuario.DataSource = Funcion.dvProcedimiento(cdsSeteoF, string.Format("Exec ListaUsuarios {0}", iBodega));

			this.Consulta();
		}

		private void Consulta()
		{
			int iBodega = 0;
			if (this.cmbBodega.ActiveRow != null) iBodega = (int)this.cmbBodega.Value;

			string sSQL = string.Format("Exec DT_ProductosReparados {0}, '{1}'",
				iBodega, this.txtBuscar.Text.ToString().Trim());
			FuncionesProcedimientos.dsGeneral(sSQL, this.uGridIngresos);

			this.lblContador.Text = this.uGridIngresos.Rows.Count + " REGISTROS ENCONTRADOS";
		}

		private void btnBuscar_Click(object sender, System.EventArgs e)
		{
			this.Consulta();			
		}

		private void cmbBodega_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Enter) this.txtBuscar.Focus();
		}

		private void txtBuscar_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Enter) this.btnBuscar_Click(sender, e);
		}

		private void uGridIngresos_BeforeSelectChange(object sender, Infragistics.Win.UltraWinGrid.BeforeSelectChangeEventArgs e)
		{
			idIngreso = 0;
			this.txtDaño.Text = "";
			this.txtEstadoFisico.Text = "";
			this.txtInformeTecnico.Text = "";
			this.txtInformeRMA.Text = "";
			idCliente = 0;
			idCompra = 0;
			sRUC = "";
			sCliente = "";
			this.cmbUsuario.Value = System.DBNull.Value;
			this.txtInformeSalida.Text = "";			
			this.txtFacturar.Value = 0.00m;
			this.lblEstado.Text = "";
			this.lblPagoPendiente.Text = "";
			this.btnNuevo.Enabled = false;
			this.btnImprimir.Enabled = false;

			idIngreso = (int)this.uGridIngresos.ActiveRow.Cells["idIngreso"].Value;

			this.txtDaño.Value = this.uGridIngresos.ActiveRow.Cells["Daño"].Value;
			this.txtEstadoFisico.Value = this.uGridIngresos.ActiveRow.Cells["EstadoFisico"].Value.ToString();
			this.txtInformeTecnico.Value = this.uGridIngresos.ActiveRow.Cells["InformeTecnico"].Value;
			this.txtInformeRMA.Value = this.uGridIngresos.ActiveRow.Cells["InformeRMA"].Value;
			
			if (Convert.ToDecimal(this.uGridIngresos.ActiveRow.Cells["ValorFactura"].Value) > 0.00m)
			{		
				idCliente = (int)this.uGridIngresos.ActiveRow.Cells["idCliente"].Value;
				sRUC = this.uGridIngresos.ActiveRow.Cells["Ruc"].Value.ToString();
				sCliente = this.uGridIngresos.ActiveRow.Cells["Nombre"].Value.ToString();															
				this.txtFacturar.Value = (decimal)this.uGridIngresos.ActiveRow.Cells["ValorFactura"].Value;
			}

			this.lblPagoPendiente.Text = Funcion.sEscalarSQL(cdsSeteoF, string.Format("Exec VerifkCliente {0}", Convert.ToInt32(this.uGridIngresos.ActiveRow.Cells["idIngreso"].Value)));

			this.lblEstado.Text = this.uGridIngresos.ActiveRow.Cells["Estado"].Value.ToString();			
		
			if ((int)this.uGridIngresos.ActiveRow.Cells["EstadoIngreso"].Value == 13 || (int)this.uGridIngresos.ActiveRow.Cells["EstadoIngreso"].Value == 31)
				this.btnImprimir.Enabled = true;

			this.btnNuevo.Enabled = true;
		}

		private void btnSalir_Click(object sender, System.EventArgs e)
		{
			this.Close();
		}

		private void btnCancelar_Click(object sender, System.EventArgs e)
		{
			idIngreso = 0;
			this.txtDaño.Text = "";
			this.txtEstadoFisico.Text = "";
			this.txtInformeTecnico.Text = "";
			this.txtInformeRMA.Text = "";
			this.cmbFormaPago.Value = System.DBNull.Value;

			idCliente = 0;
			idCompra = 0;
			sRUC = "";
			sCliente = "";
			this.cmbUsuario.Value = System.DBNull.Value;
			this.txtInformeSalida.Text = "";			
			this.txtFacturar.Value = 0.00m;
			this.lblEstado.Text = "";
			this.lblPagoPendiente.Text = "";
			
			if (iBodegaPredefinida == 0) this.cmbBodega.Enabled = true;
			this.txtBuscar.Enabled = true;
			this.btnBuscar.Enabled = true;
			this.uGridIngresos.Enabled = true;

			this.cmbUsuario.Enabled = false;
			this.txtInformeSalida.Enabled = false;
			this.cmbFormaPago.Enabled = false;

			this.btnNuevo.Enabled = false;
			this.btnGuardar.Enabled = false;
			this.btnCancelar.Enabled = false;

			bEdicion = false;

			Consulta();
		}		
		
		private void Imprimir()
		{
			#region Imprime el reporte
			Cursor = Cursors.WaitCursor;					
						
			ParameterFields paramFields = new ParameterFields ();
				
			ParameterField pfRet = new ParameterField ();
			ParameterDiscreteValue pidIngreso = new ParameterDiscreteValue ();
			pfRet.ParameterFieldName = "@idIngreso";
			pidIngreso.Value = idIngreso;
			pfRet.CurrentValues.Add (pidIngreso);
			paramFields.Add (pfRet);
				
			Reporte miRepor = new Reporte("SalidaProductoDt.rpt", "");
			miRepor.MdiParent = this.MdiParent;
			miRepor.Titulo("Ingreso DT");
			miRepor.crVista.ParameterFieldInfo = paramFields;			
			miRepor.Show();	
			
			Cursor = Cursors.Default;
			#endregion Imprime el reporte
		}

		private void GuardaEntrega(int idEstado)
		{
			string sNumero = Funcion.sEscalarSQL(cdsSeteoF, string.Format("Exec NumeracionLocales 58, {0}, 0", 1));

			string sSQL = string.Format("Exec GuardaDT_Procesos {0}, {1}, {2}, {3}, '{4}', 0, '', 0, 0, 0, 0, '', '19000101', {5}",
				idIngreso, (int)this.cmbBodega.Value, idEstado, (int)this.cmbUsuario.Value, this.txtInformeSalida.Text.ToString().Trim(), idCompra);
			Funcion.EjecutaSQL(cdsSeteoF, sSQL);
			
			if (idEstado == 28) Imprimir();			
		}

		private void btnGuardar_Click(object sender, System.EventArgs e)
		{
			if(!Validacion.vbComboVacio(this.cmbUsuario, "Seleccione el nombre de la persona que entrega")) return;

			if (!Validacion.vbTexto(this.txtInformeSalida, 20, 1000, "Informe de salida")) return;			
				
			if (Convert.ToDecimal(this.txtFacturar.Value) > 0.00m) 
				if(!Validacion.vbComboVacio(this.cmbFormaPago, "Seleccione la forma de pago")) return;
						
			#region Facturacion
			if (Convert.ToDecimal(this.txtFacturar.Value) > 0.00m)
				if (!this.CreaFactura()) return;

			GuardaEntrega(28);

			this.btnCancelar_Click(sender, e);
			#endregion Facturacion
		}

		private void btnNuevo_Click(object sender, System.EventArgs e)
		{
			if (!Validacion.vbComboVacio(this.cmbBodega, "Seleccione una bodega")) return;
			
			if (this.lblPagoPendiente.Text.ToString().Length > 0)
			{
				MessageBox.Show(this.lblPagoPendiente.Text.ToString(), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				return;
			}

			this.cmbBodega.Enabled = false;
			this.txtBuscar.Enabled = false;
			this.btnBuscar.Enabled = false;
			this.uGridIngresos.Enabled = false;

			if (idPersonal == 0) this.cmbUsuario.Enabled = true;
			else this.cmbUsuario.Value = idPersonal;	
			this.cmbUsuario.Enabled = true;			
			this.txtInformeSalida.Enabled = true;

			if (Convert.ToDecimal(this.txtFacturar.Value) > 0.00m) this.cmbFormaPago.Enabled = true;

			this.btnNuevo.Enabled = false;
			this.btnGuardar.Enabled = true;
			this.btnCancelar.Enabled = true;

			bEdicion = true;

			if (this.cmbUsuario.Enabled) this.cmbUsuario.Focus(); else this.txtInformeSalida.Focus();
		}

		private void frmEntregaProductosDT_Closing(object sender, System.ComponentModel.CancelEventArgs e)
		{
			if (bEdicion)
			{
				MessageBox.Show("No puede salir mientras edita una informe", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				e.Cancel = true;
				return;
			}
		}

		private void frmEntregaProductosDT_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (!bEdicion)
			{
				if (e.KeyCode == Keys.N) if (this.btnNuevo.Enabled) this.btnNuevo_Click(sender, e);				
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

		private void cmbUsuario_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Enter) this.txtInformeSalida.Focus();
		}

		private void txtInformeSalida_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Enter) this.cmbFormaPago.Focus();
		}

		private void cmbBodega_InitializeLayout(object sender, Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs e)
		{
		
		}

		private void cmbFormaPago_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Enter) this.btnGuardar.Focus();
		}

		private void btnImprimir_Click(object sender, System.EventArgs e)
		{
			ParameterFields paramFields = new ParameterFields ();
				
			ParameterField pfRet = new ParameterField ();
			ParameterDiscreteValue pIdIngreso = new ParameterDiscreteValue ();
			pfRet.ParameterFieldName = "@IdIngreso";
			pIdIngreso.Value = (int)this.uGridIngresos.ActiveRow.Cells["IdIngreso"].Value;
			pfRet.CurrentValues.Add (pIdIngreso);
			paramFields.Add (pfRet);
				
			Reporte Reporte = new Reporte("AutorizacionNC.rpt", "");
			Reporte.MdiParent = this.MdiParent;
			Reporte.Titulo("Ingreso DT");
			Reporte.crVista.ParameterFieldInfo = paramFields;			
			Reporte.Show();

			//this.Imprimir();
		}

		private void uGridIngresos_InitializeLayout(object sender, Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs e)
		{
		
		}
	}
}
