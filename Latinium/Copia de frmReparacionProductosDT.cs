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
	/// Descripción breve de frmReparacionProductosDT.
	/// </summary>
	public class frmReparacionProductosDT : System.Windows.Forms.Form
	{
		private Infragistics.Win.UltraWinGrid.UltraGrid ugIngresosDT;
		private System.Windows.Forms.ToolTip toolTip1;
		private System.Windows.Forms.Label label4;
		private C1.Data.C1DataSet cdsSeteoF;
		private System.Windows.Forms.GroupBox groupBox2;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource1;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource2;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtInformeTecnico;
		private Infragistics.Win.UltraWinGrid.UltraGrid ugInformesDT;
		private System.Windows.Forms.Label label36;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource3;
		private System.Windows.Forms.GroupBox groupBox3;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtEmpqsObservaciones;
		private System.Windows.Forms.CheckBox chkEmpqs;
		private System.Windows.Forms.GroupBox groupBox4;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtAccesorios;
		private System.Windows.Forms.CheckBox chkAccesorios;
		private System.Windows.Forms.GroupBox groupBox5;
		private Infragistics.Win.UltraWinEditors.UltraOptionSet rbSerial;
		private Infragistics.Win.UltraWinEditors.UltraOptionSet rbBotones;
		private Infragistics.Win.UltraWinEditors.UltraOptionSet rbPines;
		private Infragistics.Win.UltraWinEditors.UltraOptionSet rbSupPosterior;
		private Infragistics.Win.UltraWinEditors.UltraOptionSet rbSupDelantera;
		private Infragistics.Win.UltraWinEditors.UltraOptionSet rbBordes;
		private System.Windows.Forms.CheckBox chkSerial;
		private System.Windows.Forms.CheckBox chkBotones;
		private System.Windows.Forms.CheckBox chkPines;
		private System.Windows.Forms.CheckBox chkSupPosterior;
		private System.Windows.Forms.CheckBox chkSupDelantera;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtEdoFiscoProducto;
		private System.Windows.Forms.Label label14;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtDaño;
		private System.Windows.Forms.CheckBox chkBordes;
		private System.Windows.Forms.Label lblSubTotal;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtRepuesto;
		private System.Windows.Forms.Label label2;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtAFacturar;
		private System.Windows.Forms.Label label6;
		private System.Windows.Forms.Label label7;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtBuscar;
		private System.Windows.Forms.GroupBox groupBox1;
		private System.Windows.Forms.Label lblContador;
		private Infragistics.Win.UltraWinEditors.UltraOptionSet rbEmpaques;
		private Infragistics.Win.UltraWinEditors.UltraOptionSet rbAccesorios;
		private System.Windows.Forms.ComboBox cmbGarantia;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbEstadosBusqueda;
		private System.Windows.Forms.Button btnCancelar;
		private System.Windows.Forms.ImageList imageList1;
		private System.Windows.Forms.Button btnSalir;
		private System.Windows.Forms.Button btnBuscar;
		private System.Windows.Forms.Button btnNuevo;
		private System.Windows.Forms.Button btnGuardar;
		private System.Windows.Forms.Button btnImprimir;
		private System.Windows.Forms.Label label1;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbBodega;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbEstados;
		private System.ComponentModel.IContainer components;

		public frmReparacionProductosDT()
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
			Infragistics.Win.Appearance appearance2 = new Infragistics.Win.Appearance();
			Infragistics.Win.ValueListItem valueListItem1 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.ValueListItem valueListItem2 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.ValueListItem valueListItem3 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.ValueListItem valueListItem4 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.ValueListItem valueListItem5 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.Appearance appearance3 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance4 = new Infragistics.Win.Appearance();
			Infragistics.Win.ValueListItem valueListItem6 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.ValueListItem valueListItem7 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.ValueListItem valueListItem8 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.ValueListItem valueListItem9 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.ValueListItem valueListItem10 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.Appearance appearance5 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance6 = new Infragistics.Win.Appearance();
			Infragistics.Win.ValueListItem valueListItem11 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.ValueListItem valueListItem12 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.ValueListItem valueListItem13 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.ValueListItem valueListItem14 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.ValueListItem valueListItem15 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.Appearance appearance7 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance8 = new Infragistics.Win.Appearance();
			Infragistics.Win.ValueListItem valueListItem16 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.ValueListItem valueListItem17 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.ValueListItem valueListItem18 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.ValueListItem valueListItem19 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.ValueListItem valueListItem20 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.Appearance appearance9 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance10 = new Infragistics.Win.Appearance();
			Infragistics.Win.ValueListItem valueListItem21 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.ValueListItem valueListItem22 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.ValueListItem valueListItem23 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.ValueListItem valueListItem24 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.ValueListItem valueListItem25 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.Appearance appearance11 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance12 = new Infragistics.Win.Appearance();
			Infragistics.Win.ValueListItem valueListItem26 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.ValueListItem valueListItem27 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.ValueListItem valueListItem28 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.ValueListItem valueListItem29 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.ValueListItem valueListItem30 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.Appearance appearance13 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance14 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance15 = new Infragistics.Win.Appearance();
			Infragistics.Win.ValueListItem valueListItem31 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.ValueListItem valueListItem32 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.ValueListItem valueListItem33 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.Appearance appearance16 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance17 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance18 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance19 = new Infragistics.Win.Appearance();
			Infragistics.Win.ValueListItem valueListItem34 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.ValueListItem valueListItem35 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.ValueListItem valueListItem36 = new Infragistics.Win.ValueListItem();
			System.Resources.ResourceManager resources = new System.Resources.ResourceManager(typeof(frmReparacionProductosDT));
			Infragistics.Win.Appearance appearance20 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand1 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn1 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idIngreso");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn2 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idPersonal");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn3 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("InternoExterno");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn4 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("PointCliente");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn5 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Bodega");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn6 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Ingreso");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn7 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Fecha");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn8 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn9 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Articulo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn10 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Serial");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn11 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Ingresos");
			Infragistics.Win.Appearance appearance21 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn12 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Garantia");
			Infragistics.Win.Appearance appearance22 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn13 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("EstadoIngreso");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn14 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Estado");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn15 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Daño");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn16 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("ValorRepuesto");
			Infragistics.Win.Appearance appearance23 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn17 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("ValorFactura");
			Infragistics.Win.Appearance appearance24 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn18 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("NumeroNC");
			Infragistics.Win.Appearance appearance25 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance26 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance27 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance28 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance29 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance30 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance31 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn1 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idIngreso");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn2 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idPersonal");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn3 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("InternoExterno");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn4 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("PointCliente");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn5 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Bodega");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn6 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Ingreso");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn7 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Fecha");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn8 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Nombre");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn9 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Articulo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn10 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Serial");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn11 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Ingresos");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn12 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Garantia");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn13 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("EstadoIngreso");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn14 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Estado");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn15 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Daño");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn16 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("ValorRepuesto");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn17 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("ValorFactura");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn18 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("NumeroNC");
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand2 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn19 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Bodega");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn20 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn21 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Codigo", 0);
			Infragistics.Win.Appearance appearance32 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance33 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand3 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn22 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idEstado");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn23 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn19 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idEstado");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn20 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Nombre");
			Infragistics.Win.Appearance appearance34 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance35 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand4 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn24 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Fecha");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn25 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Informe");
			Infragistics.Win.Appearance appearance36 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance37 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance38 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance39 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance40 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance41 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance42 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn21 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Fecha");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn22 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Informe");
			Infragistics.Win.Appearance appearance43 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand5 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn26 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idEstado");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn27 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.Appearance appearance44 = new Infragistics.Win.Appearance();
			this.groupBox5 = new System.Windows.Forms.GroupBox();
			this.rbSerial = new Infragistics.Win.UltraWinEditors.UltraOptionSet();
			this.rbBotones = new Infragistics.Win.UltraWinEditors.UltraOptionSet();
			this.rbPines = new Infragistics.Win.UltraWinEditors.UltraOptionSet();
			this.rbSupPosterior = new Infragistics.Win.UltraWinEditors.UltraOptionSet();
			this.rbSupDelantera = new Infragistics.Win.UltraWinEditors.UltraOptionSet();
			this.rbBordes = new Infragistics.Win.UltraWinEditors.UltraOptionSet();
			this.chkSerial = new System.Windows.Forms.CheckBox();
			this.chkBotones = new System.Windows.Forms.CheckBox();
			this.chkPines = new System.Windows.Forms.CheckBox();
			this.chkSupPosterior = new System.Windows.Forms.CheckBox();
			this.chkSupDelantera = new System.Windows.Forms.CheckBox();
			this.txtEdoFiscoProducto = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.chkBordes = new System.Windows.Forms.CheckBox();
			this.groupBox4 = new System.Windows.Forms.GroupBox();
			this.rbAccesorios = new Infragistics.Win.UltraWinEditors.UltraOptionSet();
			this.txtAccesorios = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.chkAccesorios = new System.Windows.Forms.CheckBox();
			this.groupBox3 = new System.Windows.Forms.GroupBox();
			this.txtEmpqsObservaciones = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.chkEmpqs = new System.Windows.Forms.CheckBox();
			this.rbEmpaques = new Infragistics.Win.UltraWinEditors.UltraOptionSet();
			this.label4 = new System.Windows.Forms.Label();
			this.btnBuscar = new System.Windows.Forms.Button();
			this.ugIngresosDT = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.ultraDataSource1 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.toolTip1 = new System.Windows.Forms.ToolTip(this.components);
			this.btnGuardar = new System.Windows.Forms.Button();
			this.imageList1 = new System.Windows.Forms.ImageList(this.components);
			this.btnImprimir = new System.Windows.Forms.Button();
			this.btnNuevo = new System.Windows.Forms.Button();
			this.cdsSeteoF = new C1.Data.C1DataSet();
			this.groupBox2 = new System.Windows.Forms.GroupBox();
			this.label1 = new System.Windows.Forms.Label();
			this.cmbBodega = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.btnSalir = new System.Windows.Forms.Button();
			this.btnCancelar = new System.Windows.Forms.Button();
			this.cmbGarantia = new System.Windows.Forms.ComboBox();
			this.label6 = new System.Windows.Forms.Label();
			this.label2 = new System.Windows.Forms.Label();
			this.txtAFacturar = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.lblSubTotal = new System.Windows.Forms.Label();
			this.txtRepuesto = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.label14 = new System.Windows.Forms.Label();
			this.label36 = new System.Windows.Forms.Label();
			this.cmbEstados = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.ultraDataSource3 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.txtInformeTecnico = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.ugInformesDT = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.ultraDataSource2 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.txtDaño = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.label7 = new System.Windows.Forms.Label();
			this.cmbEstadosBusqueda = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.txtBuscar = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.groupBox1 = new System.Windows.Forms.GroupBox();
			this.lblContador = new System.Windows.Forms.Label();
			this.groupBox5.SuspendLayout();
			((System.ComponentModel.ISupportInitialize)(this.rbSerial)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.rbBotones)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.rbPines)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.rbSupPosterior)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.rbSupDelantera)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.rbBordes)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtEdoFiscoProducto)).BeginInit();
			this.groupBox4.SuspendLayout();
			((System.ComponentModel.ISupportInitialize)(this.rbAccesorios)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtAccesorios)).BeginInit();
			this.groupBox3.SuspendLayout();
			((System.ComponentModel.ISupportInitialize)(this.txtEmpqsObservaciones)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.rbEmpaques)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ugIngresosDT)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource1)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).BeginInit();
			this.groupBox2.SuspendLayout();
			((System.ComponentModel.ISupportInitialize)(this.cmbBodega)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtAFacturar)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtRepuesto)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbEstados)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource3)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtInformeTecnico)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ugInformesDT)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource2)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtDaño)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbEstadosBusqueda)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtBuscar)).BeginInit();
			this.SuspendLayout();
			// 
			// groupBox5
			// 
			this.groupBox5.Controls.Add(this.rbSerial);
			this.groupBox5.Controls.Add(this.rbBotones);
			this.groupBox5.Controls.Add(this.rbPines);
			this.groupBox5.Controls.Add(this.rbSupPosterior);
			this.groupBox5.Controls.Add(this.rbSupDelantera);
			this.groupBox5.Controls.Add(this.rbBordes);
			this.groupBox5.Controls.Add(this.chkSerial);
			this.groupBox5.Controls.Add(this.chkBotones);
			this.groupBox5.Controls.Add(this.chkPines);
			this.groupBox5.Controls.Add(this.chkSupPosterior);
			this.groupBox5.Controls.Add(this.chkSupDelantera);
			this.groupBox5.Controls.Add(this.txtEdoFiscoProducto);
			this.groupBox5.Controls.Add(this.chkBordes);
			this.groupBox5.Font = new System.Drawing.Font("Tahoma", 8F);
			this.groupBox5.Location = new System.Drawing.Point(8, 152);
			this.groupBox5.Name = "groupBox5";
			this.groupBox5.Size = new System.Drawing.Size(616, 160);
			this.groupBox5.TabIndex = 674;
			this.groupBox5.TabStop = false;
			this.groupBox5.Text = "Estado Físico del Producto";
			// 
			// rbSerial
			// 
			appearance1.FontData.SizeInPoints = 7F;
			appearance1.ForeColorDisabled = System.Drawing.Color.Black;
			this.rbSerial.Appearance = appearance1;
			this.rbSerial.BorderStyle = Infragistics.Win.UIElementBorderStyle.Solid;
			this.rbSerial.Enabled = false;
			this.rbSerial.ItemAppearance = appearance2;
			this.rbSerial.ItemOrigin = new System.Drawing.Point(10, 0);
			valueListItem1.DataValue = 1;
			valueListItem1.DisplayText = "B";
			valueListItem2.DataValue = 2;
			valueListItem2.DisplayText = "M";
			valueListItem3.DataValue = 3;
			valueListItem3.DisplayText = "R";
			valueListItem4.DataValue = 4;
			valueListItem4.DisplayText = "ROT";
			valueListItem5.DataValue = 5;
			valueListItem5.DisplayText = "NOV";
			this.rbSerial.Items.Add(valueListItem1);
			this.rbSerial.Items.Add(valueListItem2);
			this.rbSerial.Items.Add(valueListItem3);
			this.rbSerial.Items.Add(valueListItem4);
			this.rbSerial.Items.Add(valueListItem5);
			this.rbSerial.ItemSpacingVertical = 10;
			this.rbSerial.Location = new System.Drawing.Point(392, 64);
			this.rbSerial.Name = "rbSerial";
			this.rbSerial.Size = new System.Drawing.Size(215, 24);
			this.rbSerial.TabIndex = 356;
			// 
			// rbBotones
			// 
			appearance3.FontData.SizeInPoints = 7F;
			appearance3.ForeColorDisabled = System.Drawing.Color.Black;
			this.rbBotones.Appearance = appearance3;
			this.rbBotones.BorderStyle = Infragistics.Win.UIElementBorderStyle.Solid;
			this.rbBotones.Enabled = false;
			this.rbBotones.ItemAppearance = appearance4;
			this.rbBotones.ItemOrigin = new System.Drawing.Point(10, 0);
			valueListItem6.DataValue = 1;
			valueListItem6.DisplayText = "B";
			valueListItem7.DataValue = 2;
			valueListItem7.DisplayText = "M";
			valueListItem8.DataValue = 3;
			valueListItem8.DisplayText = "R";
			valueListItem9.DataValue = 4;
			valueListItem9.DisplayText = "ROT";
			valueListItem10.DataValue = 5;
			valueListItem10.DisplayText = "NOV";
			this.rbBotones.Items.Add(valueListItem6);
			this.rbBotones.Items.Add(valueListItem7);
			this.rbBotones.Items.Add(valueListItem8);
			this.rbBotones.Items.Add(valueListItem9);
			this.rbBotones.Items.Add(valueListItem10);
			this.rbBotones.ItemSpacingVertical = 10;
			this.rbBotones.Location = new System.Drawing.Point(72, 64);
			this.rbBotones.Name = "rbBotones";
			this.rbBotones.Size = new System.Drawing.Size(215, 24);
			this.rbBotones.TabIndex = 355;
			// 
			// rbPines
			// 
			appearance5.FontData.SizeInPoints = 7F;
			appearance5.ForeColorDisabled = System.Drawing.Color.Black;
			this.rbPines.Appearance = appearance5;
			this.rbPines.BorderStyle = Infragistics.Win.UIElementBorderStyle.Solid;
			this.rbPines.Enabled = false;
			this.rbPines.ItemAppearance = appearance6;
			this.rbPines.ItemOrigin = new System.Drawing.Point(10, 0);
			valueListItem11.DataValue = 1;
			valueListItem11.DisplayText = "B";
			valueListItem12.DataValue = 2;
			valueListItem12.DisplayText = "M";
			valueListItem13.DataValue = 3;
			valueListItem13.DisplayText = "R";
			valueListItem14.DataValue = 4;
			valueListItem14.DisplayText = "ROT";
			valueListItem15.DataValue = 5;
			valueListItem15.DisplayText = "NOV";
			this.rbPines.Items.Add(valueListItem11);
			this.rbPines.Items.Add(valueListItem12);
			this.rbPines.Items.Add(valueListItem13);
			this.rbPines.Items.Add(valueListItem14);
			this.rbPines.Items.Add(valueListItem15);
			this.rbPines.ItemSpacingVertical = 10;
			this.rbPines.Location = new System.Drawing.Point(72, 40);
			this.rbPines.Name = "rbPines";
			this.rbPines.Size = new System.Drawing.Size(215, 24);
			this.rbPines.TabIndex = 354;
			// 
			// rbSupPosterior
			// 
			appearance7.FontData.SizeInPoints = 7F;
			appearance7.ForeColorDisabled = System.Drawing.Color.Black;
			this.rbSupPosterior.Appearance = appearance7;
			this.rbSupPosterior.BorderStyle = Infragistics.Win.UIElementBorderStyle.Solid;
			this.rbSupPosterior.Enabled = false;
			this.rbSupPosterior.ItemAppearance = appearance8;
			this.rbSupPosterior.ItemOrigin = new System.Drawing.Point(10, 0);
			valueListItem16.DataValue = 1;
			valueListItem16.DisplayText = "B";
			valueListItem17.DataValue = 2;
			valueListItem17.DisplayText = "M";
			valueListItem18.DataValue = 3;
			valueListItem18.DisplayText = "R";
			valueListItem19.DataValue = 4;
			valueListItem19.DisplayText = "ROT";
			valueListItem20.DataValue = 5;
			valueListItem20.DisplayText = "NOV";
			this.rbSupPosterior.Items.Add(valueListItem16);
			this.rbSupPosterior.Items.Add(valueListItem17);
			this.rbSupPosterior.Items.Add(valueListItem18);
			this.rbSupPosterior.Items.Add(valueListItem19);
			this.rbSupPosterior.Items.Add(valueListItem20);
			this.rbSupPosterior.ItemSpacingVertical = 10;
			this.rbSupPosterior.Location = new System.Drawing.Point(392, 40);
			this.rbSupPosterior.Name = "rbSupPosterior";
			this.rbSupPosterior.Size = new System.Drawing.Size(215, 24);
			this.rbSupPosterior.TabIndex = 353;
			// 
			// rbSupDelantera
			// 
			appearance9.FontData.SizeInPoints = 7F;
			appearance9.ForeColorDisabled = System.Drawing.Color.Black;
			this.rbSupDelantera.Appearance = appearance9;
			this.rbSupDelantera.BorderStyle = Infragistics.Win.UIElementBorderStyle.Solid;
			this.rbSupDelantera.Enabled = false;
			this.rbSupDelantera.ItemAppearance = appearance10;
			this.rbSupDelantera.ItemOrigin = new System.Drawing.Point(10, 0);
			valueListItem21.DataValue = 1;
			valueListItem21.DisplayText = "B";
			valueListItem22.DataValue = 2;
			valueListItem22.DisplayText = "M";
			valueListItem23.DataValue = 3;
			valueListItem23.DisplayText = "R";
			valueListItem24.DataValue = 4;
			valueListItem24.DisplayText = "ROT";
			valueListItem25.DataValue = 5;
			valueListItem25.DisplayText = "NOV";
			this.rbSupDelantera.Items.Add(valueListItem21);
			this.rbSupDelantera.Items.Add(valueListItem22);
			this.rbSupDelantera.Items.Add(valueListItem23);
			this.rbSupDelantera.Items.Add(valueListItem24);
			this.rbSupDelantera.Items.Add(valueListItem25);
			this.rbSupDelantera.ItemSpacingVertical = 10;
			this.rbSupDelantera.Location = new System.Drawing.Point(392, 16);
			this.rbSupDelantera.Name = "rbSupDelantera";
			this.rbSupDelantera.Size = new System.Drawing.Size(215, 24);
			this.rbSupDelantera.TabIndex = 352;
			// 
			// rbBordes
			// 
			appearance11.FontData.SizeInPoints = 7F;
			appearance11.ForeColorDisabled = System.Drawing.Color.Black;
			this.rbBordes.Appearance = appearance11;
			this.rbBordes.BorderStyle = Infragistics.Win.UIElementBorderStyle.Solid;
			this.rbBordes.Enabled = false;
			this.rbBordes.ItemAppearance = appearance12;
			this.rbBordes.ItemOrigin = new System.Drawing.Point(10, 0);
			valueListItem26.DataValue = 1;
			valueListItem26.DisplayText = "B";
			valueListItem27.DataValue = 2;
			valueListItem27.DisplayText = "M";
			valueListItem28.DataValue = 3;
			valueListItem28.DisplayText = "R";
			valueListItem29.DataValue = 4;
			valueListItem29.DisplayText = "ROT";
			valueListItem30.DataValue = 5;
			valueListItem30.DisplayText = "NOV";
			this.rbBordes.Items.Add(valueListItem26);
			this.rbBordes.Items.Add(valueListItem27);
			this.rbBordes.Items.Add(valueListItem28);
			this.rbBordes.Items.Add(valueListItem29);
			this.rbBordes.Items.Add(valueListItem30);
			this.rbBordes.ItemSpacingVertical = 10;
			this.rbBordes.Location = new System.Drawing.Point(72, 16);
			this.rbBordes.Name = "rbBordes";
			this.rbBordes.Size = new System.Drawing.Size(215, 24);
			this.rbBordes.TabIndex = 351;
			// 
			// chkSerial
			// 
			this.chkSerial.Enabled = false;
			this.chkSerial.Font = new System.Drawing.Font("Tahoma", 7.5F);
			this.chkSerial.Location = new System.Drawing.Point(296, 64);
			this.chkSerial.Name = "chkSerial";
			this.chkSerial.Size = new System.Drawing.Size(96, 24);
			this.chkSerial.TabIndex = 349;
			this.chkSerial.Text = "Serial Producto";
			// 
			// chkBotones
			// 
			this.chkBotones.Enabled = false;
			this.chkBotones.Font = new System.Drawing.Font("Tahoma", 7.5F);
			this.chkBotones.Location = new System.Drawing.Point(8, 64);
			this.chkBotones.Name = "chkBotones";
			this.chkBotones.Size = new System.Drawing.Size(64, 24);
			this.chkBotones.TabIndex = 347;
			this.chkBotones.Text = "Botones";
			// 
			// chkPines
			// 
			this.chkPines.Enabled = false;
			this.chkPines.Font = new System.Drawing.Font("Tahoma", 7.5F);
			this.chkPines.Location = new System.Drawing.Point(8, 40);
			this.chkPines.Name = "chkPines";
			this.chkPines.Size = new System.Drawing.Size(64, 24);
			this.chkPines.TabIndex = 345;
			this.chkPines.Text = "Pines";
			// 
			// chkSupPosterior
			// 
			this.chkSupPosterior.Enabled = false;
			this.chkSupPosterior.Font = new System.Drawing.Font("Tahoma", 7.5F);
			this.chkSupPosterior.Location = new System.Drawing.Point(296, 40);
			this.chkSupPosterior.Name = "chkSupPosterior";
			this.chkSupPosterior.Size = new System.Drawing.Size(96, 24);
			this.chkSupPosterior.TabIndex = 343;
			this.chkSupPosterior.Text = "Sup. Posterior";
			// 
			// chkSupDelantera
			// 
			this.chkSupDelantera.Enabled = false;
			this.chkSupDelantera.Font = new System.Drawing.Font("Tahoma", 7.5F);
			this.chkSupDelantera.Location = new System.Drawing.Point(296, 16);
			this.chkSupDelantera.Name = "chkSupDelantera";
			this.chkSupDelantera.Size = new System.Drawing.Size(96, 24);
			this.chkSupDelantera.TabIndex = 341;
			this.chkSupDelantera.Text = "Sup. Delantera";
			// 
			// txtEdoFiscoProducto
			// 
			appearance13.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtEdoFiscoProducto.Appearance = appearance13;
			this.txtEdoFiscoProducto.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtEdoFiscoProducto.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtEdoFiscoProducto.Enabled = false;
			this.txtEdoFiscoProducto.Font = new System.Drawing.Font("Tahoma", 8F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.txtEdoFiscoProducto.Location = new System.Drawing.Point(8, 96);
			this.txtEdoFiscoProducto.Multiline = true;
			this.txtEdoFiscoProducto.Name = "txtEdoFiscoProducto";
			this.txtEdoFiscoProducto.Scrollbars = System.Windows.Forms.ScrollBars.Vertical;
			this.txtEdoFiscoProducto.Size = new System.Drawing.Size(600, 60);
			this.txtEdoFiscoProducto.TabIndex = 339;
			// 
			// chkBordes
			// 
			this.chkBordes.Enabled = false;
			this.chkBordes.Font = new System.Drawing.Font("Tahoma", 7.5F);
			this.chkBordes.Location = new System.Drawing.Point(8, 16);
			this.chkBordes.Name = "chkBordes";
			this.chkBordes.Size = new System.Drawing.Size(64, 24);
			this.chkBordes.TabIndex = 0;
			this.chkBordes.Text = "Bordes";
			// 
			// groupBox4
			// 
			this.groupBox4.Controls.Add(this.rbAccesorios);
			this.groupBox4.Controls.Add(this.txtAccesorios);
			this.groupBox4.Controls.Add(this.chkAccesorios);
			this.groupBox4.Font = new System.Drawing.Font("Tahoma", 8.25F);
			this.groupBox4.Location = new System.Drawing.Point(856, 152);
			this.groupBox4.Name = "groupBox4";
			this.groupBox4.Size = new System.Drawing.Size(232, 160);
			this.groupBox4.TabIndex = 673;
			this.groupBox4.TabStop = false;
			// 
			// rbAccesorios
			// 
			appearance14.FontData.SizeInPoints = 7.5F;
			appearance14.ForeColorDisabled = System.Drawing.Color.Black;
			this.rbAccesorios.Appearance = appearance14;
			this.rbAccesorios.BorderStyle = Infragistics.Win.UIElementBorderStyle.Solid;
			this.rbAccesorios.Enabled = false;
			this.rbAccesorios.ItemAppearance = appearance15;
			this.rbAccesorios.ItemOrigin = new System.Drawing.Point(10, 0);
			valueListItem31.DataValue = 1;
			valueListItem31.DisplayText = "Bueno";
			valueListItem32.DataValue = 2;
			valueListItem32.DisplayText = "Malo";
			valueListItem33.DataValue = 3;
			valueListItem33.DisplayText = "Novedades";
			this.rbAccesorios.Items.Add(valueListItem31);
			this.rbAccesorios.Items.Add(valueListItem32);
			this.rbAccesorios.Items.Add(valueListItem33);
			this.rbAccesorios.ItemSpacingVertical = 10;
			this.rbAccesorios.Location = new System.Drawing.Point(8, 24);
			this.rbAccesorios.Name = "rbAccesorios";
			this.rbAccesorios.Size = new System.Drawing.Size(216, 24);
			this.rbAccesorios.TabIndex = 385;
			// 
			// txtAccesorios
			// 
			appearance16.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtAccesorios.Appearance = appearance16;
			this.txtAccesorios.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtAccesorios.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtAccesorios.Enabled = false;
			this.txtAccesorios.Font = new System.Drawing.Font("Tahoma", 8F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.txtAccesorios.Location = new System.Drawing.Point(8, 56);
			this.txtAccesorios.Multiline = true;
			this.txtAccesorios.Name = "txtAccesorios";
			this.txtAccesorios.Scrollbars = System.Windows.Forms.ScrollBars.Vertical;
			this.txtAccesorios.Size = new System.Drawing.Size(216, 100);
			this.txtAccesorios.TabIndex = 339;
			// 
			// chkAccesorios
			// 
			this.chkAccesorios.Enabled = false;
			this.chkAccesorios.Font = new System.Drawing.Font("Tahoma", 8.25F);
			this.chkAccesorios.Location = new System.Drawing.Point(8, 0);
			this.chkAccesorios.Name = "chkAccesorios";
			this.chkAccesorios.Size = new System.Drawing.Size(80, 24);
			this.chkAccesorios.TabIndex = 0;
			this.chkAccesorios.Text = "Accesorios";
			this.chkAccesorios.CheckedChanged += new System.EventHandler(this.chkAccesorios_CheckedChanged);
			// 
			// groupBox3
			// 
			this.groupBox3.Controls.Add(this.txtEmpqsObservaciones);
			this.groupBox3.Controls.Add(this.chkEmpqs);
			this.groupBox3.Controls.Add(this.rbEmpaques);
			this.groupBox3.Font = new System.Drawing.Font("Tahoma", 7F);
			this.groupBox3.Location = new System.Drawing.Point(624, 152);
			this.groupBox3.Name = "groupBox3";
			this.groupBox3.Size = new System.Drawing.Size(232, 160);
			this.groupBox3.TabIndex = 672;
			this.groupBox3.TabStop = false;
			// 
			// txtEmpqsObservaciones
			// 
			appearance17.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtEmpqsObservaciones.Appearance = appearance17;
			this.txtEmpqsObservaciones.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtEmpqsObservaciones.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtEmpqsObservaciones.Enabled = false;
			this.txtEmpqsObservaciones.Font = new System.Drawing.Font("Tahoma", 8F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.txtEmpqsObservaciones.Location = new System.Drawing.Point(8, 56);
			this.txtEmpqsObservaciones.Multiline = true;
			this.txtEmpqsObservaciones.Name = "txtEmpqsObservaciones";
			this.txtEmpqsObservaciones.Scrollbars = System.Windows.Forms.ScrollBars.Vertical;
			this.txtEmpqsObservaciones.Size = new System.Drawing.Size(216, 100);
			this.txtEmpqsObservaciones.TabIndex = 339;
			// 
			// chkEmpqs
			// 
			this.chkEmpqs.Enabled = false;
			this.chkEmpqs.Font = new System.Drawing.Font("Tahoma", 8.25F);
			this.chkEmpqs.Location = new System.Drawing.Point(8, 0);
			this.chkEmpqs.Name = "chkEmpqs";
			this.chkEmpqs.Size = new System.Drawing.Size(120, 24);
			this.chkEmpqs.TabIndex = 0;
			this.chkEmpqs.Text = "Empaques de Caja";
			// 
			// rbEmpaques
			// 
			appearance18.FontData.SizeInPoints = 7.5F;
			appearance18.ForeColorDisabled = System.Drawing.Color.Black;
			this.rbEmpaques.Appearance = appearance18;
			this.rbEmpaques.BorderStyle = Infragistics.Win.UIElementBorderStyle.Solid;
			this.rbEmpaques.Enabled = false;
			this.rbEmpaques.ItemAppearance = appearance19;
			this.rbEmpaques.ItemOrigin = new System.Drawing.Point(10, 0);
			valueListItem34.DataValue = 1;
			valueListItem34.DisplayText = "Bueno";
			valueListItem35.DataValue = 2;
			valueListItem35.DisplayText = "Malo";
			valueListItem36.DataValue = 3;
			valueListItem36.DisplayText = "Novedades";
			this.rbEmpaques.Items.Add(valueListItem34);
			this.rbEmpaques.Items.Add(valueListItem35);
			this.rbEmpaques.Items.Add(valueListItem36);
			this.rbEmpaques.ItemSpacingVertical = 10;
			this.rbEmpaques.Location = new System.Drawing.Point(8, 24);
			this.rbEmpaques.Name = "rbEmpaques";
			this.rbEmpaques.Size = new System.Drawing.Size(216, 24);
			this.rbEmpaques.TabIndex = 697;
			// 
			// label4
			// 
			this.label4.AutoSize = true;
			this.label4.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label4.Location = new System.Drawing.Point(328, 5);
			this.label4.Name = "label4";
			this.label4.Size = new System.Drawing.Size(37, 17);
			this.label4.TabIndex = 670;
			this.label4.Text = "Buscar";
			this.label4.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// btnBuscar
			// 
			this.btnBuscar.CausesValidation = false;
			this.btnBuscar.Image = ((System.Drawing.Image)(resources.GetObject("btnBuscar.Image")));
			this.btnBuscar.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnBuscar.Location = new System.Drawing.Point(592, 2);
			this.btnBuscar.Name = "btnBuscar";
			this.btnBuscar.Size = new System.Drawing.Size(80, 23);
			this.btnBuscar.TabIndex = 659;
			this.btnBuscar.Text = "Buscar";
			this.btnBuscar.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.toolTip1.SetToolTip(this.btnBuscar, "Actualizar información");
			this.btnBuscar.Click += new System.EventHandler(this.btnBuscar_Click);
			// 
			// ugIngresosDT
			// 
			this.ugIngresosDT.Cursor = System.Windows.Forms.Cursors.Default;
			this.ugIngresosDT.DataSource = this.ultraDataSource1;
			appearance20.BackColor = System.Drawing.Color.White;
			this.ugIngresosDT.DisplayLayout.Appearance = appearance20;
			ultraGridColumn1.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn1.Header.VisiblePosition = 0;
			ultraGridColumn1.Hidden = true;
			ultraGridColumn2.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn2.Header.VisiblePosition = 1;
			ultraGridColumn2.Hidden = true;
			ultraGridColumn3.Header.VisiblePosition = 2;
			ultraGridColumn3.Hidden = true;
			ultraGridColumn4.Header.VisiblePosition = 3;
			ultraGridColumn4.Hidden = true;
			ultraGridColumn5.Header.VisiblePosition = 4;
			ultraGridColumn5.Hidden = true;
			ultraGridColumn6.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn6.Header.VisiblePosition = 5;
			ultraGridColumn6.Width = 100;
			ultraGridColumn7.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn7.Header.VisiblePosition = 7;
			ultraGridColumn7.Width = 70;
			ultraGridColumn8.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn8.Header.Caption = "Cliente";
			ultraGridColumn8.Header.VisiblePosition = 8;
			ultraGridColumn8.Width = 215;
			ultraGridColumn9.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn9.Header.VisiblePosition = 9;
			ultraGridColumn9.Width = 180;
			ultraGridColumn10.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn10.Header.VisiblePosition = 10;
			ultraGridColumn10.Width = 145;
			ultraGridColumn11.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance21.TextHAlign = Infragistics.Win.HAlign.Center;
			ultraGridColumn11.CellAppearance = appearance21;
			ultraGridColumn11.Header.VisiblePosition = 11;
			ultraGridColumn11.Width = 60;
			ultraGridColumn12.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance22.TextHAlign = Infragistics.Win.HAlign.Center;
			ultraGridColumn12.CellAppearance = appearance22;
			ultraGridColumn12.Header.VisiblePosition = 12;
			ultraGridColumn12.Width = 90;
			ultraGridColumn13.Header.VisiblePosition = 13;
			ultraGridColumn13.Hidden = true;
			ultraGridColumn14.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn14.Header.VisiblePosition = 6;
			ultraGridColumn14.Width = 180;
			ultraGridColumn15.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn15.Header.VisiblePosition = 14;
			ultraGridColumn15.Hidden = true;
			ultraGridColumn16.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance23.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn16.CellAppearance = appearance23;
			ultraGridColumn16.Format = "#,##0.00";
			ultraGridColumn16.Header.Caption = "V. Repuesto";
			ultraGridColumn16.Header.VisiblePosition = 15;
			ultraGridColumn16.Width = 80;
			ultraGridColumn17.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance24.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn17.CellAppearance = appearance24;
			ultraGridColumn17.Format = "#,##0.00";
			ultraGridColumn17.Header.Caption = "V. Factura";
			ultraGridColumn17.Header.VisiblePosition = 16;
			ultraGridColumn17.Width = 80;
			ultraGridColumn18.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn18.Format = "#,##0.00";
			ultraGridColumn18.Header.Caption = "Numero N/C";
			ultraGridColumn18.Header.VisiblePosition = 17;
			ultraGridColumn18.Hidden = true;
			ultraGridColumn18.Width = 80;
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
																										 ultraGridColumn18});
			this.ugIngresosDT.DisplayLayout.BandsSerializer.Add(ultraGridBand1);
			appearance25.ForeColor = System.Drawing.Color.Black;
			appearance25.ForeColorDisabled = System.Drawing.Color.Black;
			this.ugIngresosDT.DisplayLayout.Override.ActiveRowAppearance = appearance25;
			this.ugIngresosDT.DisplayLayout.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.No;
			this.ugIngresosDT.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.False;
			this.ugIngresosDT.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.False;
			appearance26.BackColor = System.Drawing.Color.Transparent;
			this.ugIngresosDT.DisplayLayout.Override.CardAreaAppearance = appearance26;
			appearance27.ForeColor = System.Drawing.Color.Black;
			appearance27.ForeColorDisabled = System.Drawing.Color.Black;
			this.ugIngresosDT.DisplayLayout.Override.GroupByRowAppearance = appearance27;
			appearance28.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance28.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance28.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance28.FontData.BoldAsString = "True";
			appearance28.FontData.Name = "Arial";
			appearance28.FontData.SizeInPoints = 8.5F;
			appearance28.ForeColor = System.Drawing.Color.White;
			appearance28.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.ugIngresosDT.DisplayLayout.Override.HeaderAppearance = appearance28;
			this.ugIngresosDT.DisplayLayout.Override.HeaderClickAction = Infragistics.Win.UltraWinGrid.HeaderClickAction.SortSingle;
			appearance29.BackColor = System.Drawing.SystemColors.ControlLight;
			appearance29.BackColor2 = System.Drawing.Color.Gainsboro;
			appearance29.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ugIngresosDT.DisplayLayout.Override.RowAlternateAppearance = appearance29;
			appearance30.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance30.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance30.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ugIngresosDT.DisplayLayout.Override.RowSelectorAppearance = appearance30;
			appearance31.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance31.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance31.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ugIngresosDT.DisplayLayout.Override.SelectedRowAppearance = appearance31;
			this.ugIngresosDT.Font = new System.Drawing.Font("Tahoma", 8F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.ugIngresosDT.Location = new System.Drawing.Point(8, 35);
			this.ugIngresosDT.Name = "ugIngresosDT";
			this.ugIngresosDT.Size = new System.Drawing.Size(1080, 120);
			this.ugIngresosDT.TabIndex = 658;
			this.ugIngresosDT.BeforeSelectChange += new Infragistics.Win.UltraWinGrid.BeforeSelectChangeEventHandler(this.ugIngresosDT_BeforeSelectChange);
			this.ugIngresosDT.InitializeLayout += new Infragistics.Win.UltraWinGrid.InitializeLayoutEventHandler(this.ugIngresosDT_InitializeLayout);
			// 
			// ultraDataSource1
			// 
			ultraDataColumn1.DataType = typeof(int);
			ultraDataColumn2.DataType = typeof(int);
			ultraDataColumn3.DataType = typeof(int);
			ultraDataColumn3.DefaultValue = 0;
			ultraDataColumn4.DataType = typeof(int);
			ultraDataColumn4.DefaultValue = 0;
			ultraDataColumn5.DataType = typeof(int);
			ultraDataColumn7.DataType = typeof(System.DateTime);
			ultraDataColumn11.DataType = typeof(int);
			ultraDataColumn11.DefaultValue = 0;
			ultraDataColumn13.DataType = typeof(int);
			ultraDataColumn16.DataType = typeof(System.Decimal);
			ultraDataColumn16.DefaultValue = new System.Decimal(new int[] {
																																			0,
																																			0,
																																			0,
																																			0});
			ultraDataColumn17.DataType = typeof(System.Decimal);
			ultraDataColumn17.DefaultValue = new System.Decimal(new int[] {
																																			0,
																																			0,
																																			0,
																																			0});
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
																																 ultraDataColumn18});
			// 
			// btnGuardar
			// 
			this.btnGuardar.CausesValidation = false;
			this.btnGuardar.Enabled = false;
			this.btnGuardar.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnGuardar.ImageIndex = 9;
			this.btnGuardar.ImageList = this.imageList1;
			this.btnGuardar.Location = new System.Drawing.Point(744, 213);
			this.btnGuardar.Name = "btnGuardar";
			this.btnGuardar.TabIndex = 672;
			this.btnGuardar.Text = "Guardar";
			this.btnGuardar.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.toolTip1.SetToolTip(this.btnGuardar, "Grabar Información");
			this.btnGuardar.Click += new System.EventHandler(this.btnGuardar_Click);
			// 
			// imageList1
			// 
			this.imageList1.ImageSize = new System.Drawing.Size(16, 16);
			this.imageList1.ImageStream = ((System.Windows.Forms.ImageListStreamer)(resources.GetObject("imageList1.ImageStream")));
			this.imageList1.TransparentColor = System.Drawing.Color.Transparent;
			// 
			// btnImprimir
			// 
			this.btnImprimir.CausesValidation = false;
			this.btnImprimir.Enabled = false;
			this.btnImprimir.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnImprimir.ImageIndex = 6;
			this.btnImprimir.ImageList = this.imageList1;
			this.btnImprimir.Location = new System.Drawing.Point(824, 213);
			this.btnImprimir.Name = "btnImprimir";
			this.btnImprimir.TabIndex = 673;
			this.btnImprimir.Text = "Imprimir";
			this.btnImprimir.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.toolTip1.SetToolTip(this.btnImprimir, "Imprimir Documento Autorizado");
			this.btnImprimir.Click += new System.EventHandler(this.btnImprimir_Click);
			// 
			// btnNuevo
			// 
			this.btnNuevo.CausesValidation = false;
			this.btnNuevo.Enabled = false;
			this.btnNuevo.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnNuevo.ImageIndex = 50;
			this.btnNuevo.ImageList = this.imageList1;
			this.btnNuevo.Location = new System.Drawing.Point(664, 213);
			this.btnNuevo.Name = "btnNuevo";
			this.btnNuevo.TabIndex = 674;
			this.btnNuevo.Text = "Nuevo";
			this.btnNuevo.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.toolTip1.SetToolTip(this.btnNuevo, "Nuevo informe técnico");
			this.btnNuevo.Click += new System.EventHandler(this.btnNuevo_Click);
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
			// groupBox2
			// 
			this.groupBox2.Controls.Add(this.label1);
			this.groupBox2.Controls.Add(this.cmbBodega);
			this.groupBox2.Controls.Add(this.btnSalir);
			this.groupBox2.Controls.Add(this.btnCancelar);
			this.groupBox2.Controls.Add(this.cmbGarantia);
			this.groupBox2.Controls.Add(this.label6);
			this.groupBox2.Controls.Add(this.label2);
			this.groupBox2.Controls.Add(this.txtAFacturar);
			this.groupBox2.Controls.Add(this.lblSubTotal);
			this.groupBox2.Controls.Add(this.txtRepuesto);
			this.groupBox2.Controls.Add(this.label14);
			this.groupBox2.Controls.Add(this.label36);
			this.groupBox2.Controls.Add(this.cmbEstados);
			this.groupBox2.Controls.Add(this.btnNuevo);
			this.groupBox2.Controls.Add(this.btnImprimir);
			this.groupBox2.Controls.Add(this.txtInformeTecnico);
			this.groupBox2.Controls.Add(this.btnGuardar);
			this.groupBox2.Controls.Add(this.ugInformesDT);
			this.groupBox2.Controls.Add(this.txtDaño);
			this.groupBox2.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.groupBox2.Location = new System.Drawing.Point(8, 312);
			this.groupBox2.Name = "groupBox2";
			this.groupBox2.Size = new System.Drawing.Size(1080, 240);
			this.groupBox2.TabIndex = 1;
			this.groupBox2.TabStop = false;
			this.groupBox2.Text = "Informe Tecnico";
			// 
			// label1
			// 
			this.label1.AutoSize = true;
			this.label1.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label1.Location = new System.Drawing.Point(664, 107);
			this.label1.Name = "label1";
			this.label1.Size = new System.Drawing.Size(83, 17);
			this.label1.TabIndex = 698;
			this.label1.Text = "Bodega Destino";
			this.label1.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// cmbBodega
			// 
			this.cmbBodega.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbBodega.Cursor = System.Windows.Forms.Cursors.Default;
			ultraGridColumn19.Header.VisiblePosition = 2;
			ultraGridColumn20.Header.VisiblePosition = 1;
			ultraGridColumn20.Width = 180;
			ultraGridColumn21.Header.VisiblePosition = 0;
			ultraGridBand2.Columns.AddRange(new object[] {
																										 ultraGridColumn19,
																										 ultraGridColumn20,
																										 ultraGridColumn21});
			this.cmbBodega.DisplayLayout.BandsSerializer.Add(ultraGridBand2);
			this.cmbBodega.DisplayMember = "Nombre";
			this.cmbBodega.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbBodega.Enabled = false;
			this.cmbBodega.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbBodega.Location = new System.Drawing.Point(760, 104);
			this.cmbBodega.MaxDropDownItems = 30;
			this.cmbBodega.Name = "cmbBodega";
			this.cmbBodega.Size = new System.Drawing.Size(280, 22);
			this.cmbBodega.TabIndex = 697;
			this.cmbBodega.ValueMember = "Bodega";
			this.cmbBodega.KeyDown += new System.Windows.Forms.KeyEventHandler(this.cmbBodega_KeyDown);
			// 
			// btnSalir
			// 
			this.btnSalir.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btnSalir.CausesValidation = false;
			this.btnSalir.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnSalir.ImageIndex = 62;
			this.btnSalir.ImageList = this.imageList1;
			this.btnSalir.Location = new System.Drawing.Point(984, 213);
			this.btnSalir.Name = "btnSalir";
			this.btnSalir.TabIndex = 696;
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
			this.btnCancelar.Location = new System.Drawing.Point(904, 213);
			this.btnCancelar.Name = "btnCancelar";
			this.btnCancelar.TabIndex = 695;
			this.btnCancelar.Text = "&Cancelar";
			this.btnCancelar.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnCancelar.Click += new System.EventHandler(this.btnCancelar_Click);
			// 
			// cmbGarantia
			// 
			this.cmbGarantia.Enabled = false;
			this.cmbGarantia.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbGarantia.Items.AddRange(new object[] {
																										 "SI",
																										 "NO",
																										 "POR VERIFICAR"});
			this.cmbGarantia.Location = new System.Drawing.Point(760, 32);
			this.cmbGarantia.Name = "cmbGarantia";
			this.cmbGarantia.Size = new System.Drawing.Size(128, 21);
			this.cmbGarantia.TabIndex = 694;
			this.cmbGarantia.KeyDown += new System.Windows.Forms.KeyEventHandler(this.cmbGarantia_KeyDown);
			this.cmbGarantia.SelectedValueChanged += new System.EventHandler(this.cmbGarantia_SelectedValueChanged);
			// 
			// label6
			// 
			this.label6.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label6.Location = new System.Drawing.Point(664, 34);
			this.label6.Name = "label6";
			this.label6.Size = new System.Drawing.Size(48, 16);
			this.label6.TabIndex = 693;
			this.label6.Text = "Garantia";
			this.label6.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// label2
			// 
			this.label2.AutoSize = true;
			this.label2.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label2.Location = new System.Drawing.Point(880, 83);
			this.label2.Name = "label2";
			this.label2.Size = new System.Drawing.Size(51, 17);
			this.label2.TabIndex = 686;
			this.label2.Text = "Factura $";
			this.label2.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// txtAFacturar
			// 
			appearance32.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtAFacturar.Appearance = appearance32;
			this.txtAFacturar.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtAFacturar.Enabled = false;
			this.txtAFacturar.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.txtAFacturar.FormatString = "#,##0.00";
			this.txtAFacturar.Location = new System.Drawing.Point(952, 80);
			this.txtAFacturar.Name = "txtAFacturar";
			this.txtAFacturar.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtAFacturar.PromptChar = ' ';
			this.txtAFacturar.Size = new System.Drawing.Size(88, 22);
			this.txtAFacturar.TabIndex = 687;
			this.txtAFacturar.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtAFacturar_KeyDown);
			this.txtAFacturar.Click += new System.EventHandler(this.txtAFacturar_Click);
			this.txtAFacturar.Enter += new System.EventHandler(this.txtAFacturar_Enter);
			// 
			// lblSubTotal
			// 
			this.lblSubTotal.AutoSize = true;
			this.lblSubTotal.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.lblSubTotal.Location = new System.Drawing.Point(664, 83);
			this.lblSubTotal.Name = "lblSubTotal";
			this.lblSubTotal.Size = new System.Drawing.Size(60, 17);
			this.lblSubTotal.TabIndex = 684;
			this.lblSubTotal.Text = "Repuesto $";
			this.lblSubTotal.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// txtRepuesto
			// 
			appearance33.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtRepuesto.Appearance = appearance33;
			this.txtRepuesto.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtRepuesto.Enabled = false;
			this.txtRepuesto.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.txtRepuesto.FormatString = "#,##0.00";
			this.txtRepuesto.Location = new System.Drawing.Point(760, 80);
			this.txtRepuesto.Name = "txtRepuesto";
			this.txtRepuesto.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtRepuesto.PromptChar = ' ';
			this.txtRepuesto.Size = new System.Drawing.Size(88, 22);
			this.txtRepuesto.TabIndex = 685;
			this.txtRepuesto.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtRepuesto_KeyDown);
			this.txtRepuesto.Click += new System.EventHandler(this.txtRepuesto_Click);
			this.txtRepuesto.Enter += new System.EventHandler(this.txtRepuesto_Enter);
			// 
			// label14
			// 
			this.label14.AutoSize = true;
			this.label14.BackColor = System.Drawing.Color.Transparent;
			this.label14.Location = new System.Drawing.Point(8, 16);
			this.label14.Name = "label14";
			this.label14.Size = new System.Drawing.Size(85, 17);
			this.label14.TabIndex = 679;
			this.label14.Text = "Daño Reportado";
			// 
			// label36
			// 
			this.label36.AutoSize = true;
			this.label36.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label36.Location = new System.Drawing.Point(664, 59);
			this.label36.Name = "label36";
			this.label36.Size = new System.Drawing.Size(38, 17);
			this.label36.TabIndex = 676;
			this.label36.Text = "Estado";
			this.label36.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// cmbEstados
			// 
			this.cmbEstados.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbEstados.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbEstados.DataSource = this.ultraDataSource3;
			ultraGridColumn22.Header.VisiblePosition = 0;
			ultraGridColumn22.Hidden = true;
			ultraGridColumn23.Header.VisiblePosition = 1;
			ultraGridColumn23.Width = 360;
			ultraGridBand3.Columns.AddRange(new object[] {
																										 ultraGridColumn22,
																										 ultraGridColumn23});
			this.cmbEstados.DisplayLayout.BandsSerializer.Add(ultraGridBand3);
			this.cmbEstados.DisplayMember = "Nombre";
			this.cmbEstados.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbEstados.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbEstados.Enabled = false;
			this.cmbEstados.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbEstados.Location = new System.Drawing.Point(760, 56);
			this.cmbEstados.MaxDropDownItems = 30;
			this.cmbEstados.Name = "cmbEstados";
			this.cmbEstados.Size = new System.Drawing.Size(280, 22);
			this.cmbEstados.TabIndex = 677;
			this.cmbEstados.ValueMember = "idEstado";
			this.cmbEstados.KeyDown += new System.Windows.Forms.KeyEventHandler(this.cmbEstados_KeyDown);
			this.cmbEstados.ValueChanged += new System.EventHandler(this.cmbEstados_ValueChanged);
			this.cmbEstados.InitializeLayout += new Infragistics.Win.UltraWinGrid.InitializeLayoutEventHandler(this.cmbEstados_InitializeLayout);
			// 
			// ultraDataSource3
			// 
			ultraDataColumn19.DataType = typeof(int);
			ultraDataColumn19.ReadOnly = Infragistics.Win.DefaultableBoolean.False;
			this.ultraDataSource3.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn19,
																																 ultraDataColumn20});
			// 
			// txtInformeTecnico
			// 
			appearance34.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtInformeTecnico.Appearance = appearance34;
			this.txtInformeTecnico.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtInformeTecnico.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtInformeTecnico.Enabled = false;
			this.txtInformeTecnico.Font = new System.Drawing.Font("Tahoma", 8F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.txtInformeTecnico.Location = new System.Drawing.Point(664, 128);
			this.txtInformeTecnico.Multiline = true;
			this.txtInformeTecnico.Name = "txtInformeTecnico";
			this.txtInformeTecnico.Scrollbars = System.Windows.Forms.ScrollBars.Vertical;
			this.txtInformeTecnico.Size = new System.Drawing.Size(408, 80);
			this.txtInformeTecnico.TabIndex = 343;
			this.txtInformeTecnico.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtInformeTecnico_KeyDown);
			// 
			// ugInformesDT
			// 
			this.ugInformesDT.Cursor = System.Windows.Forms.Cursors.Default;
			this.ugInformesDT.DataSource = this.ultraDataSource2;
			appearance35.BackColor = System.Drawing.Color.White;
			this.ugInformesDT.DisplayLayout.Appearance = appearance35;
			this.ugInformesDT.DisplayLayout.AutoFitColumns = true;
			ultraGridColumn24.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn24.Format = "dd/MM/yyyy HH:mm";
			ultraGridColumn24.Header.VisiblePosition = 0;
			ultraGridColumn24.Width = 100;
			ultraGridColumn25.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn25.Header.Caption = "Informe Técnico";
			ultraGridColumn25.Header.VisiblePosition = 1;
			ultraGridColumn25.Width = 529;
			ultraGridBand4.Columns.AddRange(new object[] {
																										 ultraGridColumn24,
																										 ultraGridColumn25});
			this.ugInformesDT.DisplayLayout.BandsSerializer.Add(ultraGridBand4);
			appearance36.ForeColor = System.Drawing.Color.Black;
			appearance36.ForeColorDisabled = System.Drawing.Color.Black;
			this.ugInformesDT.DisplayLayout.Override.ActiveRowAppearance = appearance36;
			this.ugInformesDT.DisplayLayout.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.No;
			this.ugInformesDT.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.False;
			this.ugInformesDT.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.False;
			appearance37.BackColor = System.Drawing.Color.Transparent;
			this.ugInformesDT.DisplayLayout.Override.CardAreaAppearance = appearance37;
			appearance38.ForeColor = System.Drawing.Color.Black;
			appearance38.ForeColorDisabled = System.Drawing.Color.Black;
			this.ugInformesDT.DisplayLayout.Override.GroupByRowAppearance = appearance38;
			appearance39.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance39.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance39.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance39.FontData.BoldAsString = "True";
			appearance39.FontData.Name = "Arial";
			appearance39.FontData.SizeInPoints = 8.5F;
			appearance39.ForeColor = System.Drawing.Color.White;
			appearance39.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.ugInformesDT.DisplayLayout.Override.HeaderAppearance = appearance39;
			this.ugInformesDT.DisplayLayout.Override.HeaderClickAction = Infragistics.Win.UltraWinGrid.HeaderClickAction.SortSingle;
			appearance40.BackColor = System.Drawing.SystemColors.ControlLight;
			appearance40.BackColor2 = System.Drawing.Color.Gainsboro;
			appearance40.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ugInformesDT.DisplayLayout.Override.RowAlternateAppearance = appearance40;
			appearance41.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance41.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance41.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ugInformesDT.DisplayLayout.Override.RowSelectorAppearance = appearance41;
			appearance42.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance42.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance42.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ugInformesDT.DisplayLayout.Override.SelectedRowAppearance = appearance42;
			this.ugInformesDT.Font = new System.Drawing.Font("Tahoma", 8F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.ugInformesDT.Location = new System.Drawing.Point(8, 84);
			this.ugInformesDT.Name = "ugInformesDT";
			this.ugInformesDT.Size = new System.Drawing.Size(650, 152);
			this.ugInformesDT.TabIndex = 674;
			// 
			// ultraDataSource2
			// 
			ultraDataColumn21.DataType = typeof(System.DateTime);
			this.ultraDataSource2.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn21,
																																 ultraDataColumn22});
			// 
			// txtDaño
			// 
			appearance43.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtDaño.Appearance = appearance43;
			this.txtDaño.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtDaño.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtDaño.Enabled = false;
			this.txtDaño.Font = new System.Drawing.Font("Tahoma", 8F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.txtDaño.Location = new System.Drawing.Point(8, 35);
			this.txtDaño.Multiline = true;
			this.txtDaño.Name = "txtDaño";
			this.txtDaño.Scrollbars = System.Windows.Forms.ScrollBars.Vertical;
			this.txtDaño.Size = new System.Drawing.Size(648, 40);
			this.txtDaño.TabIndex = 678;
			// 
			// label7
			// 
			this.label7.AutoSize = true;
			this.label7.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label7.Location = new System.Drawing.Point(8, 5);
			this.label7.Name = "label7";
			this.label7.Size = new System.Drawing.Size(38, 17);
			this.label7.TabIndex = 678;
			this.label7.Text = "Estado";
			this.label7.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// cmbEstadosBusqueda
			// 
			this.cmbEstadosBusqueda.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbEstadosBusqueda.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbEstadosBusqueda.DataSource = this.ultraDataSource3;
			ultraGridColumn26.Header.VisiblePosition = 0;
			ultraGridColumn26.Hidden = true;
			ultraGridColumn27.Header.VisiblePosition = 1;
			ultraGridColumn27.Width = 250;
			ultraGridBand5.Columns.AddRange(new object[] {
																										 ultraGridColumn26,
																										 ultraGridColumn27});
			this.cmbEstadosBusqueda.DisplayLayout.BandsSerializer.Add(ultraGridBand5);
			this.cmbEstadosBusqueda.DisplayMember = "Nombre";
			this.cmbEstadosBusqueda.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbEstadosBusqueda.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbEstadosBusqueda.Location = new System.Drawing.Point(56, 3);
			this.cmbEstadosBusqueda.MaxDropDownItems = 30;
			this.cmbEstadosBusqueda.Name = "cmbEstadosBusqueda";
			this.cmbEstadosBusqueda.Size = new System.Drawing.Size(250, 22);
			this.cmbEstadosBusqueda.TabIndex = 679;
			this.cmbEstadosBusqueda.ValueMember = "idEstado";
			this.cmbEstadosBusqueda.KeyDown += new System.Windows.Forms.KeyEventHandler(this.cmbEstadosBusqueda_KeyDown);
			this.cmbEstadosBusqueda.ValueChanged += new System.EventHandler(this.cmbEstadosBusqueda_ValueChanged);
			// 
			// txtBuscar
			// 
			appearance44.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtBuscar.Appearance = appearance44;
			this.txtBuscar.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtBuscar.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtBuscar.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.txtBuscar.Location = new System.Drawing.Point(376, 3);
			this.txtBuscar.MaxLength = 100;
			this.txtBuscar.Name = "txtBuscar";
			this.txtBuscar.Size = new System.Drawing.Size(200, 22);
			this.txtBuscar.TabIndex = 694;
			this.txtBuscar.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtBuscar_KeyDown);
			// 
			// groupBox1
			// 
			this.groupBox1.Location = new System.Drawing.Point(0, 24);
			this.groupBox1.Name = "groupBox1";
			this.groupBox1.Size = new System.Drawing.Size(1096, 8);
			this.groupBox1.TabIndex = 695;
			this.groupBox1.TabStop = false;
			// 
			// lblContador
			// 
			this.lblContador.AutoSize = true;
			this.lblContador.Font = new System.Drawing.Font("Tahoma", 9F);
			this.lblContador.ForeColor = System.Drawing.Color.Firebrick;
			this.lblContador.Location = new System.Drawing.Point(688, 4);
			this.lblContador.Name = "lblContador";
			this.lblContador.Size = new System.Drawing.Size(0, 18);
			this.lblContador.TabIndex = 696;
			this.lblContador.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
			// 
			// frmReparacionProductosDT
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 13);
			this.ClientSize = new System.Drawing.Size(1096, 558);
			this.Controls.Add(this.txtBuscar);
			this.Controls.Add(this.cmbEstadosBusqueda);
			this.Controls.Add(this.ugIngresosDT);
			this.Controls.Add(this.btnBuscar);
			this.Controls.Add(this.lblContador);
			this.Controls.Add(this.groupBox1);
			this.Controls.Add(this.label7);
			this.Controls.Add(this.label4);
			this.Controls.Add(this.groupBox2);
			this.Controls.Add(this.groupBox3);
			this.Controls.Add(this.groupBox4);
			this.Controls.Add(this.groupBox5);
			this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle;
			this.KeyPreview = true;
			this.MaximizeBox = false;
			this.Name = "frmReparacionProductosDT";
			this.ShowInTaskbar = false;
			this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
			this.Text = "Informes de Reparación";
			this.KeyDown += new System.Windows.Forms.KeyEventHandler(this.frmReparacionProductosDT_KeyDown);
			this.Closing += new System.ComponentModel.CancelEventHandler(this.frmReparacionProductosDT_Closing);
			this.Load += new System.EventHandler(this.frmReparacionProductosDT_Load);
			this.groupBox5.ResumeLayout(false);
			((System.ComponentModel.ISupportInitialize)(this.rbSerial)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.rbBotones)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.rbPines)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.rbSupPosterior)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.rbSupDelantera)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.rbBordes)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtEdoFiscoProducto)).EndInit();
			this.groupBox4.ResumeLayout(false);
			((System.ComponentModel.ISupportInitialize)(this.rbAccesorios)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtAccesorios)).EndInit();
			this.groupBox3.ResumeLayout(false);
			((System.ComponentModel.ISupportInitialize)(this.txtEmpqsObservaciones)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.rbEmpaques)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ugIngresosDT)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource1)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).EndInit();
			this.groupBox2.ResumeLayout(false);
			((System.ComponentModel.ISupportInitialize)(this.cmbBodega)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtAFacturar)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtRepuesto)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbEstados)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource3)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtInformeTecnico)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ugInformesDT)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource2)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtDaño)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbEstadosBusqueda)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtBuscar)).EndInit();
			this.ResumeLayout(false);

		}
		#endregion

		/* Variables de la pantalla */
		int idIngreso = 0;
		int idTecnico = 0;
		bool bEdicion = false;
		int iTipo = 0;

		private void cdsSeteoF_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			cdsSeteoF.Schema.Connections[0].ConnectionString = miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);
		}

		private void frmReparacionProductosDT_Load(object sender, System.EventArgs e)
		{			
			this.cmbEstadosBusqueda.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "Exec DTListaEstados 1");
			this.cmbEstados.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "Exec DTListaEstados 2");
			this.cmbBodega.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "Exec DTListaBodegasSalida 0");

			this.ListaIngresos();
		}

		public void VisulizaInfoProducto()
		{						
			if(this.ugIngresosDT.Rows.Count == 0)
			{
				MessageBox.Show("No existe información para visualizar", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Information);
				return;
			}
	
			BDatosSolicitud((int)this.ugIngresosDT.ActiveRow.Cells["idIngreso"].Value);		

			idIngreso = (int)this.ugIngresosDT.ActiveRow.Cells["idIngreso"].Value;
			idTecnico = (int)this.ugIngresosDT.ActiveRow.Cells["idPersonal"].Value;
			this.txtDaño.Value = this.ugIngresosDT.ActiveRow.Cells["Daño"].Value;
			this.cmbGarantia.Text = this.ugIngresosDT.ActiveRow.Cells["Garantia"].Value.ToString().Trim();
			this.txtRepuesto.Value = (decimal)this.ugIngresosDT.ActiveRow.Cells["ValorRepuesto"].Value;
			this.txtAFacturar.Value = (decimal)this.ugIngresosDT.ActiveRow.Cells["ValorFactura"].Value;				
			
			if ((int)this.ugIngresosDT.ActiveRow.Cells["EstadoIngreso"].Value == 8 || 
				(int)this.ugIngresosDT.ActiveRow.Cells["EstadoIngreso"].Value == 9 ||
				(int)this.ugIngresosDT.ActiveRow.Cells["EstadoIngreso"].Value == 12) this.btnNuevo.Enabled = false; else this.btnNuevo.Enabled = true;

			if ((int)this.ugIngresosDT.ActiveRow.Cells["EstadoIngreso"].Value == 13 || (int)this.ugIngresosDT.ActiveRow.Cells["EstadoIngreso"].Value == 30 ||
				(int)this.ugIngresosDT.ActiveRow.Cells["EstadoIngreso"].Value == 14 || (int)this.ugIngresosDT.ActiveRow.Cells["EstadoIngreso"].Value == 29) this.btnImprimir.Enabled = true; else this.btnImprimir.Enabled = false;				
			
			this.ListaInformes((int)this.ugIngresosDT.ActiveRow.Cells["idIngreso"].Value);
		}

		private void BDatosSolicitud(int idIngreso)
		{
			string sSQL = string.Format("Exec DTBuskIngreso {0}", idIngreso);
			SqlDataReader dr = Funcion.rEscalarSQL(cdsSeteoF, sSQL, true);
			dr.Read();
			if (dr.HasRows)
			{				
				this.chkBordes.Checked = dr.GetBoolean(44);
				this.rbBordes.Value = dr.GetInt32(35);

				this.chkPines.Checked = dr.GetBoolean(47);
				this.rbPines.Value = dr.GetInt32(37);

				this.chkBotones.Checked = dr.GetBoolean(48);
				this.rbBotones.Value = dr.GetInt32(36);

				this.chkSupDelantera.Checked = dr.GetBoolean(45);
				this.rbSupDelantera.Value = dr.GetInt32(39);

				this.chkSupPosterior.Checked = dr.GetBoolean(46);
				this.rbSupPosterior.Value = dr.GetInt32(40);
				
				this.chkSerial.Checked = dr.GetBoolean(49);
				this.rbSerial.Value = dr.GetInt32(38);

				this.txtEdoFiscoProducto.Value = dr.GetString(41).Trim();

				this.chkEmpqs.Checked = dr.GetBoolean(29);
				this.rbEmpaques.Value = dr.GetInt32(28);
				this.txtEmpqsObservaciones.Value = dr.GetString(34).Trim();

				this.chkAccesorios.Checked = dr.GetBoolean(24);
				this.rbAccesorios.Value = dr.GetInt32(27);
				this.txtAccesorios.Value = dr.GetString(25).Trim();								
			}
			dr.Close();
		}

		private void ListaInformes(int idIngreso)
		{
			string sSQL = string.Format("Exec DT_InformesDT {0}", idIngreso);
			FuncionesProcedimientos.dsGeneral(sSQL, this.ugInformesDT);
		}

		private void ListaIngresos()
		{			
			ListaInformes(0);

			int idEstado = 0;
			if (this.cmbEstadosBusqueda.ActiveRow != null) idEstado = (int)this.cmbEstadosBusqueda.Value;

			string sSQL = string.Format("Exec DTListaIngresosTecnicos {0}, {1}, '{2}'", MenuLatinium.IdUsuario, idEstado, this.txtBuscar.Text.ToString());
			FuncionesProcedimientos.dsGeneral(sSQL, this.ugIngresosDT);

			this.lblContador.Text = this.ugIngresosDT.Rows.Count + " REGISTROS ENCONTRADOS";
		}

		private void txtRepuesto_Enter(object sender, System.EventArgs e)
		{
			this.txtRepuesto.SelectAll();	
		}

		private void txtRepuesto_Click(object sender, System.EventArgs e)
		{
			this.txtRepuesto.SelectAll();			
		}

		private void txtAFacturar_Click(object sender, System.EventArgs e)
		{
			this.txtAFacturar.SelectAll();	
		}

		private void txtAFacturar_Enter(object sender, System.EventArgs e)
		{
			this.txtAFacturar.SelectAll();		
		}
		
		private void ugIngresosDT_BeforeSelectChange(object sender, Infragistics.Win.UltraWinGrid.BeforeSelectChangeEventArgs e)
		{
			VisulizaInfoProducto();		
		}

		private void chkAccesorios_CheckedChanged(object sender, System.EventArgs e)
		{
		
		}

		private void btnBuscar_Click(object sender, System.EventArgs e)
		{
			this.ListaIngresos();

			if (this.ugIngresosDT.Rows.Count > 0) this.ugIngresosDT.Focus();
		}

		private void btnSalir_Click(object sender, System.EventArgs e)
		{
			this.Close();
		}

		private void txtBuscar_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Enter) this.ListaIngresos();
		}

		private void cmbEstadosBusqueda_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Enter) this.txtBuscar.Focus();
		}

		private void cmbEstadosBusqueda_ValueChanged(object sender, System.EventArgs e)
		{
			this.ListaIngresos();
		}

		private void btnNuevo_Click(object sender, System.EventArgs e)
		{			
			this.cmbEstadosBusqueda.Enabled = false;
			this.txtBuscar.Enabled = false;
			this.btnBuscar.Enabled = false;
			this.ugIngresosDT.Enabled = false;

			this.cmbEstados.Enabled = true;
									
			if ((int)this.ugIngresosDT.ActiveRow.Cells["PointCliente"].Value == 1)
			{
				if (this.cmbGarantia.Text == "SI") iTipo = 2;
				if (this.cmbGarantia.Text == "NO") iTipo = 3;				

				if (this.cmbGarantia.Text == "POR VERIFICAR" || this.cmbGarantia.Text == "SI") this.cmbGarantia.Enabled = true;
			}
			if ((int)this.ugIngresosDT.ActiveRow.Cells["PointCliente"].Value == 2)
			{
				if (this.cmbGarantia.Text == "SI") iTipo = 4;
				if (this.cmbGarantia.Text == "NO") iTipo = 5;
				if (this.cmbGarantia.Text == "POR VERIFICAR") iTipo = 6;

				if (this.cmbGarantia.Text == "POR VERIFICAR" || this.cmbGarantia.Text == "SI") this.cmbGarantia.Enabled = true;
			}

//			if ((int)this.ugIngresosDT.ActiveRow.Cells["PointCliente"].Value == 1)
//			{				
//				if (this.cmbGarantia.Text == "NO") iTipo = 3;				
//			}

			this.cmbEstados.DataSource = Funcion.dvProcedimiento(cdsSeteoF, string.Format("Exec DTListaEstados {0}", iTipo));

			this.txtInformeTecnico.Enabled = true;

			this.btnNuevo.Enabled = false;
			this.btnGuardar.Enabled = true;
			this.btnImprimir.Enabled = false;
			this.btnCancelar.Enabled = true;

			bEdicion = true;

			if (this.cmbGarantia.Enabled) this.cmbGarantia.Focus(); else this.cmbEstados.Focus();
		}

		private void btnGuardar_Click(object sender, System.EventArgs e)
		{	
			if (this.cmbGarantia.Text.ToString().Trim().Length == 0)
			{
				MessageBox.Show("Seleccione un estado de garantia", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Information);
				this.cmbGarantia.Focus();
				return;
			}

			if (!Validacion.vbComboVacio(this.cmbEstados, "Seleccione un estado")) return;
													
			if (this.cmbGarantia.Text.Trim().Equals("POR VERIFICAR") && (int)this.cmbEstados.Value != 6 && (int)this.cmbEstados.Value != 5)
			{
				MessageBox.Show("Para grabar debe confirmar si el producto tiene o no garantia", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Information);
				this.cmbGarantia.Focus();
				return;
			}

			if ((int)this.cmbEstados.Value == 9 && this.cmbGarantia.Text.Trim().Equals("POR VERIFICAR"))
			{
				MessageBox.Show("Articulo sin verificacion de garantia no puede solicitar nota de credito", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Information);
				this.cmbEstados.Focus();
				return;	
			}

			if ((int)this.cmbEstados.Value == 9 && (this.cmbGarantia.Text.Trim().Equals("NO") || this.cmbGarantia.Text.Trim().Equals("POR VERIFICAR")))
			{
				MessageBox.Show("Articulo sin garantia no puede solicitar nota de credito", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Information);
				this.cmbEstados.Focus();
				return;	
			}
							
			if ((int)this.cmbEstados.Value == 7 || (int)this.cmbEstados.Value == 11)
			{
				if (iTipo == 5 && !Funcion.bEscalarSQL(cdsSeteoF, string.Format("Exec DT_ValidaProcesosPrevios {0}, {1}", idIngreso, (int)this.cmbEstados.Value), true))
				{
					MessageBox.Show("No se ha registrado la autorizacion del cliente para generar factura", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Information);
					this.cmbEstados.Focus();
					return;	
				}

				if ((int)this.ugIngresosDT.ActiveRow.Cells["PointCliente"].Value == 2 && this.cmbGarantia.Text == "NO" && Convert.ToDecimal(this.txtAFacturar.Value) == 0.00m)
				{
					MessageBox.Show("Articulo sin Garantia requiere valor a facturar", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Information);
					this.txtAFacturar.Focus();
					return;	
				}

				if ((int)this.ugIngresosDT.ActiveRow.Cells["InternoExterno"].Value == 2 && Convert.ToDecimal(this.txtAFacturar.Value) == 0.00m)
				{
					MessageBox.Show("Articulo externo requiere valor a facturar", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Information);
					this.txtAFacturar.Focus();
					return;	
				}

				if (Convert.ToDecimal(this.txtRepuesto.Value) > 0.00m)
				{
					if (Convert.ToDecimal(this.txtRepuesto.Value) >= Convert.ToDecimal(this.txtAFacturar.Value))
					{
						MessageBox.Show("El valor del repuesto no puede ser mayor al valor a facturar", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Information);
						this.txtAFacturar.Focus();
						return;	
					}
				}
			}

			if (!Validacion.vbTexto(this.txtInformeTecnico, 20, 1000, "Informe Tecnico")) return;

			int iBodega = 0;
			if (this.cmbBodega.ActiveRow != null) iBodega = (int)this.cmbBodega.Value;

			string sSQL = string.Format("Exec DT_RegistrarInforme {0}, '{1}', {2}, {3}, {4}, {5}, '{6}', '{7}', {8}", 
				idIngreso, this.txtInformeTecnico.Value.ToString().Trim(), 
				(int)this.cmbEstados.Value, (int)this.ugIngresosDT.ActiveRow.Cells["idPersonal"].Value,
				this.txtRepuesto.Value, this.txtAFacturar.Value, this.cmbGarantia.Text.ToString().Trim(),
				this.ugIngresosDT.ActiveRow.Cells["NumeroNC"].Value.ToString(), iBodega);					
			Funcion.EjecutaSQL(cdsSeteoF, sSQL);
			
			this.ListaIngresos();

			#region Controles
			this.cmbEstadosBusqueda.Enabled = true;
			this.txtBuscar.Enabled = true;
			this.btnBuscar.Enabled = true;
			this.ugIngresosDT.Enabled = true;

			this.cmbEstados.Value = System.DBNull.Value;
			this.cmbGarantia.Text = "";
			this.txtRepuesto.Value = 0.00m;
			this.txtAFacturar.Value = 0.00m;
			this.cmbBodega.Value = System.DBNull.Value;
			this.txtInformeTecnico.Text = "";

			this.cmbEstados.Enabled = false;
			this.cmbGarantia.Enabled = false;
			this.txtRepuesto.Enabled = false;
			this.txtAFacturar.Enabled = false;			
			this.cmbBodega.Enabled = false;
			this.txtInformeTecnico.Enabled = false;

			this.btnGuardar.Enabled = false;
			
			bEdicion = false;
			#endregion Controles					

			this.btnBuscar_Click(sender, e);			
		}

		private void btnImprimir_Click(object sender, System.EventArgs e)
		{
			#region Imprime el reporte
			ParameterFields paramFields = new ParameterFields ();
				
			ParameterField pfRet = new ParameterField ();
			ParameterDiscreteValue idIngresos = new ParameterDiscreteValue ();
			pfRet.ParameterFieldName = "@idIngreso";
			idIngresos.Value = idIngreso;
			pfRet.CurrentValues.Add (idIngresos);
			paramFields.Add (pfRet);
				
			Reporte miRepor = new Reporte("AutorizacionProdct.rpt", "");
			miRepor.MdiParent = this.MdiParent;
			miRepor.Titulo("Ingreso DT");
			miRepor.crVista.ParameterFieldInfo = paramFields;			
			miRepor.Show();
			#endregion Imprime el reporte
		}

		private void btnCancelar_Click(object sender, System.EventArgs e)
		{
			this.cmbEstadosBusqueda.Enabled = true;
			this.txtBuscar.Enabled = true;
			this.btnBuscar.Enabled = true;
			this.ugIngresosDT.Enabled = true;
			
			this.cmbEstados.Value = System.DBNull.Value;
			this.cmbGarantia.Text = "";
			this.txtRepuesto.Value = 0.00m;
			this.txtAFacturar.Value = 0.00m;
			this.cmbBodega.Value = System.DBNull.Value;
			this.txtInformeTecnico.Text = "";

			this.cmbEstados.Enabled = false;
			this.cmbGarantia.Enabled = false;
			this.txtAFacturar.Enabled = false;
			this.txtRepuesto.Enabled = false;
			this.cmbBodega.Enabled = false;
			this.txtInformeTecnico.Enabled = false;

			this.btnNuevo.Enabled = false;
			this.btnGuardar.Enabled = false;
			this.btnImprimir.Enabled = false;
			this.btnCancelar.Enabled = false;
			
			idIngreso = 0;
			bEdicion = false;
			this.ListaIngresos();
		}

		private void frmReparacionProductosDT_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
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

		private void frmReparacionProductosDT_Closing(object sender, System.ComponentModel.CancelEventArgs e)
		{
			if (bEdicion)
			{
				MessageBox.Show("No puede salir mientras edita una informe", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				e.Cancel = true;
				return;
			}
		}

		private void cmbGarantia_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Enter) this.cmbEstados.Focus();
		}

		private void txtRepuesto_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Enter)
			{
				if (this.cmbBodega.Enabled) this.cmbBodega.Focus();
				else if (this.txtAFacturar.Enabled) this.txtAFacturar.Focus();
				else if (this.txtInformeTecnico.Enabled) this.txtInformeTecnico.Focus();
			}
		}

		private void cmbBodega_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Enter)
			{
				if (this.txtAFacturar.Enabled) this.txtAFacturar.Focus();				
				else if (this.txtInformeTecnico.Enabled) this.txtInformeTecnico.Focus();
			}
		}

		private void txtAFacturar_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Enter) this.txtInformeTecnico.Focus();
		}

		private void txtInformeTecnico_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Enter) this.btnGuardar.Focus();
		}

		private void cmbGarantia_SelectedValueChanged(object sender, System.EventArgs e)
		{
			if (!bEdicion) return;

			if ((int)this.ugIngresosDT.ActiveRow.Cells["PointCliente"].Value == 1)
			{
				if (this.cmbGarantia.Text == "SI") iTipo = 2;
				if (this.cmbGarantia.Text == "NO") iTipo = 3;				
			}
			if ((int)this.ugIngresosDT.ActiveRow.Cells["PointCliente"].Value == 2)
			{
				if (this.cmbGarantia.Text == "SI") iTipo = 4;
				if (this.cmbGarantia.Text == "NO") iTipo = 5;
				if (this.cmbGarantia.Text == "POR VERIFICAR") iTipo = 6;
			}

			this.cmbEstados.DataSource = Funcion.dvProcedimiento(cdsSeteoF, string.Format("Exec DTListaEstados {0}", iTipo));
		}

		private void cmbEstados_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Enter)
			{
				if (this.txtRepuesto.Enabled) this.txtRepuesto.Focus();
				else if (this.cmbBodega.Enabled) this.cmbBodega.Focus();
				else if (this.txtAFacturar.Enabled) this.txtAFacturar.Focus();				
				else if (this.txtInformeTecnico.Enabled) this.txtInformeTecnico.Focus();
			}
		}

		private void cmbEstados_ValueChanged(object sender, System.EventArgs e)
		{
			if (!bEdicion) return;
						
			if (this.cmbEstados.ActiveRow == null) return;

			if ((int)this.cmbEstados.Value == 5 || (int)this.cmbEstados.Value == 6 ||
				(int)this.cmbEstados.Value == 7 || (int)this.cmbEstados.Value == 9 ||
				(int)this.cmbEstados.Value == 10 || (int)this.cmbEstados.Value == 24)
			{
				if ((int)this.cmbEstados.Value != 10) this.txtRepuesto.Value = 0.00m;
				this.txtAFacturar.Value = 0.00m;				
			}

			this.txtRepuesto.Enabled = false;
			this.txtAFacturar.Enabled = false;
			this.cmbBodega.Enabled = false;

			if ((int)this.cmbEstados.Value == 7)
			{				
				this.txtRepuesto.Enabled = true;
				/*VALIDAR CONDICION PARA ACTIVAR FACTURAS CON CLIENTE INTERNO SIN GARANTIA*/
//(int)this.ugIngresosDT.ActiveRow.Cells["PointCliente"].Value == 2 && 
				if (this.cmbGarantia.Text == "NO") this.txtAFacturar.Enabled = true;
				if ((int)this.ugIngresosDT.ActiveRow.Cells["InternoExterno"].Value == 2) this.txtAFacturar.Enabled = true;
			}
			
			if ((int)this.cmbEstados.Value == 10 || (int)this.cmbEstados.Value == 11 || (int)this.cmbEstados.Value == 24)
			{
				this.cmbBodega.Enabled = true; 		
	
				if ((int)this.cmbEstados.Value == 10) this.cmbBodega.Value = 20;
				else this.cmbBodega.Value = (int)this.ugIngresosDT.ActiveRow.Cells["Bodega"].Value;
			}
		}

		private void ugIngresosDT_InitializeLayout(object sender, Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs e)
		{
		
		}

		private void cmbEstados_InitializeLayout(object sender, Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs e)
		{
		
		}
	}
}
