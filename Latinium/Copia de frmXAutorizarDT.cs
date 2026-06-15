using System;
using System.Data;
using System.Data.SqlClient;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;

namespace Latinium
{
	/// <summary>
	/// Descripción breve de frmXAutorizarDT.
	/// </summary>
	public class frmXAutorizarDT : System.Windows.Forms.Form
	{
		private System.Windows.Forms.Label label4;
		private System.Windows.Forms.ToolTip toolTip1;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource1;
		private Infragistics.Win.UltraWinGrid.UltraGrid ugProdAutorizar;
		private C1.Data.C1DataSet cdsSeteoF;
		private System.Windows.Forms.ImageList imageList1;
		private System.Windows.Forms.Button btnVer;
		private System.Windows.Forms.Label lblContador;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtBuscar;
		private System.Windows.Forms.GroupBox groupBox2;
		private System.Windows.Forms.Label label9;
		private System.Windows.Forms.Label label7;
		private System.Windows.Forms.Label label25;
		private System.Windows.Forms.Label label20;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtNotas;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtDepreciacion;
		private System.Windows.Forms.Label label19;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtSolicita;
		private System.Windows.Forms.Label label22;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtMovitoNC;
		private System.Windows.Forms.Label label15;
		private System.Windows.Forms.Label label14;
		private System.Windows.Forms.Label label13;
		private System.Windows.Forms.Label label12;
		private System.Windows.Forms.Label label10;
		private System.Windows.Forms.Label label8;
		private System.Windows.Forms.Label label6;
		private System.Windows.Forms.Label label5;
		private System.Windows.Forms.Label label3;
		private System.Windows.Forms.Label label2;
		private System.Windows.Forms.Label label1;
		private System.Windows.Forms.Label lblComprobante;
		private System.Windows.Forms.GroupBox groupBox1;
		private System.Windows.Forms.GroupBox groupBox3;
		private System.Windows.Forms.GroupBox groupBox4;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtRepuesto;
		private System.Windows.Forms.Label label16;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtAFacturar;
		private System.Windows.Forms.Label label17;
		private System.Windows.Forms.Label label36;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbEstadosDT;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtBodega;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtLocal;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtInformeTecnico;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtEstadoFisico;
		private Infragistics.Win.UltraWinSchedule.UltraCalendarCombo dtFecha;
		private Infragistics.Win.UltraWinSchedule.UltraCalendarCombo dtFechaFactura;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtSerial;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtDescripcion;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtCodigo;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtFactura;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtIdentificacion;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtCliente;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtIngreso;
		private Infragistics.Win.UltraWinSchedule.UltraCalendarCombo dtFechaIDT;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtGarantia;
		private System.Windows.Forms.Label label18;
		private System.Windows.Forms.Label lblEstado;
		private System.Windows.Forms.Button btnSalir;
		private System.Windows.Forms.Button btnCancelar;
		private System.Windows.Forms.Button btnNuevo;
		private System.Windows.Forms.Button btnGuardar;
		private System.Windows.Forms.GroupBox groupBox5;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtVendedor;
		private System.Windows.Forms.Label label11;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtPrecio;
		private System.Windows.Forms.Label label21;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource2;
		private System.Windows.Forms.Label label23;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbMotivos;
		private System.Windows.Forms.Label label24;
		private Infragistics.Win.UltraWinEditors.UltraOptionSet optTipo;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtDescuento;
		private System.Windows.Forms.Label label26;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtNumero;
		private System.ComponentModel.IContainer components;

		public frmXAutorizarDT()
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
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn2 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Ingreso");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn3 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Bodega");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn4 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Fecha");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn5 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("FechaRecepcion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn6 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idFVenta");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn7 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Numero");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn8 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Local");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn9 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("FechaFactura");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn10 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCliente");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn11 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Ruc");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn12 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Cliente");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn13 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Vendedor");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn14 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idArticulo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn15 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Codigo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn16 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Articulo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn17 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Serial");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn18 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Precio");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn19 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Garantia");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn20 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idPersonal");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn21 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Solicita");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn22 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("ValorRepuesto");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn23 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("ValorFactura");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn24 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("EstadoFisico");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn25 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("InformeTecnico");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn26 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("MotivoNC");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn27 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("EstadoIngreso");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn28 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("EstadoM");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn29 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idFormaPago");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn30 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("LocalFactura");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn31 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idProyecto");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn32 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idEntidadFinanciera");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn33 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idVendedor");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn34 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idSerial");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn35 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idTipoRuc");
			Infragistics.Win.Appearance appearance2 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance3 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance4 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance5 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance6 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance7 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance8 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn1 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idIngreso");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn2 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Ingreso");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn3 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Bodega");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn4 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Fecha");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn5 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("FechaRecepcion");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn6 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idFVenta");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn7 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Numero");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn8 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Local");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn9 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("FechaFactura");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn10 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idCliente");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn11 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Ruc");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn12 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Cliente");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn13 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Vendedor");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn14 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idArticulo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn15 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Codigo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn16 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Articulo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn17 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Serial");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn18 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Precio");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn19 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Garantia");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn20 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idPersonal");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn21 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Solicita");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn22 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("ValorRepuesto");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn23 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("ValorFactura");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn24 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("EstadoFisico");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn25 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("InformeTecnico");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn26 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("MotivoNC");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn27 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("EstadoIngreso");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn28 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("EstadoM");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn29 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idFormaPago");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn30 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("LocalFactura");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn31 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idProyecto");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn32 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idEntidadFinanciera");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn33 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idVendedor");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn34 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idSerial");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn35 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idTipoRuc");
			System.Resources.ResourceManager resources = new System.Resources.ResourceManager(typeof(frmXAutorizarDT));
			Infragistics.Win.Appearance appearance9 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance10 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton dateButton1 = new Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton();
			Infragistics.Win.Appearance appearance11 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance12 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance13 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance14 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance15 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance16 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance17 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance18 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton dateButton2 = new Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton();
			Infragistics.Win.Appearance appearance19 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton dateButton3 = new Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton();
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
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand2 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn36 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idEstado");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn37 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.Appearance appearance30 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance31 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance32 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn36 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idMotivo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn37 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Motivo");
			Infragistics.Win.Appearance appearance33 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand3 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn38 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idMotivo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn39 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Motivo");
			Infragistics.Win.Appearance appearance34 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance35 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance36 = new Infragistics.Win.Appearance();
			Infragistics.Win.ValueListItem valueListItem1 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.ValueListItem valueListItem2 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.Appearance appearance37 = new Infragistics.Win.Appearance();
			this.ugProdAutorizar = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.ultraDataSource1 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.label4 = new System.Windows.Forms.Label();
			this.btnVer = new System.Windows.Forms.Button();
			this.imageList1 = new System.Windows.Forms.ImageList(this.components);
			this.toolTip1 = new System.Windows.Forms.ToolTip(this.components);
			this.btnNuevo = new System.Windows.Forms.Button();
			this.btnGuardar = new System.Windows.Forms.Button();
			this.cdsSeteoF = new C1.Data.C1DataSet();
			this.lblContador = new System.Windows.Forms.Label();
			this.txtBuscar = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.groupBox2 = new System.Windows.Forms.GroupBox();
			this.dtFechaIDT = new Infragistics.Win.UltraWinSchedule.UltraCalendarCombo();
			this.label9 = new System.Windows.Forms.Label();
			this.txtBodega = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.label7 = new System.Windows.Forms.Label();
			this.txtLocal = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.label25 = new System.Windows.Forms.Label();
			this.label20 = new System.Windows.Forms.Label();
			this.txtNotas = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.txtDepreciacion = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.label19 = new System.Windows.Forms.Label();
			this.txtSolicita = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.label22 = new System.Windows.Forms.Label();
			this.txtInformeTecnico = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.txtEstadoFisico = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.dtFecha = new Infragistics.Win.UltraWinSchedule.UltraCalendarCombo();
			this.dtFechaFactura = new Infragistics.Win.UltraWinSchedule.UltraCalendarCombo();
			this.txtMovitoNC = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.txtSerial = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.txtDescripcion = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.txtCodigo = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.txtFactura = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.txtIdentificacion = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.txtCliente = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.label15 = new System.Windows.Forms.Label();
			this.label14 = new System.Windows.Forms.Label();
			this.label13 = new System.Windows.Forms.Label();
			this.label12 = new System.Windows.Forms.Label();
			this.label10 = new System.Windows.Forms.Label();
			this.label8 = new System.Windows.Forms.Label();
			this.label6 = new System.Windows.Forms.Label();
			this.label5 = new System.Windows.Forms.Label();
			this.label3 = new System.Windows.Forms.Label();
			this.label2 = new System.Windows.Forms.Label();
			this.label1 = new System.Windows.Forms.Label();
			this.lblComprobante = new System.Windows.Forms.Label();
			this.txtIngreso = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.groupBox1 = new System.Windows.Forms.GroupBox();
			this.groupBox3 = new System.Windows.Forms.GroupBox();
			this.groupBox4 = new System.Windows.Forms.GroupBox();
			this.txtRepuesto = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.label16 = new System.Windows.Forms.Label();
			this.txtAFacturar = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.label17 = new System.Windows.Forms.Label();
			this.label36 = new System.Windows.Forms.Label();
			this.cmbEstadosDT = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.txtGarantia = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.label18 = new System.Windows.Forms.Label();
			this.lblEstado = new System.Windows.Forms.Label();
			this.btnSalir = new System.Windows.Forms.Button();
			this.btnCancelar = new System.Windows.Forms.Button();
			this.groupBox5 = new System.Windows.Forms.GroupBox();
			this.txtVendedor = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.label11 = new System.Windows.Forms.Label();
			this.txtPrecio = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.label21 = new System.Windows.Forms.Label();
			this.ultraDataSource2 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.label23 = new System.Windows.Forms.Label();
			this.cmbMotivos = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.label24 = new System.Windows.Forms.Label();
			this.txtDescuento = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.optTipo = new Infragistics.Win.UltraWinEditors.UltraOptionSet();
			this.label26 = new System.Windows.Forms.Label();
			this.txtNumero = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			((System.ComponentModel.ISupportInitialize)(this.ugProdAutorizar)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource1)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtBuscar)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.dtFechaIDT)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtBodega)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtLocal)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNotas)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtDepreciacion)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtSolicita)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtInformeTecnico)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtEstadoFisico)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.dtFecha)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.dtFechaFactura)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtMovitoNC)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtSerial)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtDescripcion)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtCodigo)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtFactura)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIdentificacion)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtCliente)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIngreso)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtRepuesto)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtAFacturar)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbEstadosDT)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtGarantia)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtVendedor)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtPrecio)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource2)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbMotivos)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtDescuento)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.optTipo)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNumero)).BeginInit();
			this.SuspendLayout();
			// 
			// ugProdAutorizar
			// 
			this.ugProdAutorizar.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.ugProdAutorizar.Cursor = System.Windows.Forms.Cursors.Default;
			this.ugProdAutorizar.DataSource = this.ultraDataSource1;
			appearance1.BackColor = System.Drawing.Color.White;
			this.ugProdAutorizar.DisplayLayout.Appearance = appearance1;
			ultraGridColumn1.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn1.Header.VisiblePosition = 0;
			ultraGridColumn1.Hidden = true;
			ultraGridColumn2.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn2.Header.VisiblePosition = 2;
			ultraGridColumn2.Width = 100;
			ultraGridColumn3.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn3.Header.VisiblePosition = 4;
			ultraGridColumn3.Hidden = true;
			ultraGridColumn3.Width = 40;
			ultraGridColumn4.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn4.Header.VisiblePosition = 3;
			ultraGridColumn4.Width = 70;
			ultraGridColumn5.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn5.Header.VisiblePosition = 5;
			ultraGridColumn5.Hidden = true;
			ultraGridColumn6.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn6.Header.VisiblePosition = 7;
			ultraGridColumn6.Hidden = true;
			ultraGridColumn7.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn7.Header.VisiblePosition = 18;
			ultraGridColumn7.Hidden = true;
			ultraGridColumn8.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn8.Header.VisiblePosition = 10;
			ultraGridColumn8.Hidden = true;
			ultraGridColumn9.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn9.Header.VisiblePosition = 12;
			ultraGridColumn9.Hidden = true;
			ultraGridColumn10.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn10.Header.VisiblePosition = 14;
			ultraGridColumn10.Hidden = true;
			ultraGridColumn11.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn11.Header.Caption = "Identificacion";
			ultraGridColumn11.Header.VisiblePosition = 6;
			ultraGridColumn11.Width = 90;
			ultraGridColumn12.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn12.Header.VisiblePosition = 8;
			ultraGridColumn12.Width = 220;
			ultraGridColumn13.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn13.Header.VisiblePosition = 26;
			ultraGridColumn13.Hidden = true;
			ultraGridColumn14.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn14.Header.VisiblePosition = 16;
			ultraGridColumn14.Hidden = true;
			ultraGridColumn15.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn15.Header.Caption = "Código";
			ultraGridColumn15.Header.VisiblePosition = 9;
			ultraGridColumn15.Width = 180;
			ultraGridColumn16.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn16.Header.Caption = "Descripcion";
			ultraGridColumn16.Header.VisiblePosition = 11;
			ultraGridColumn16.Width = 220;
			ultraGridColumn17.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn17.Header.VisiblePosition = 13;
			ultraGridColumn17.Width = 150;
			ultraGridColumn18.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn18.Header.VisiblePosition = 27;
			ultraGridColumn18.Hidden = true;
			ultraGridColumn19.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn19.Header.VisiblePosition = 23;
			ultraGridColumn19.Hidden = true;
			ultraGridColumn20.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn20.Header.VisiblePosition = 1;
			ultraGridColumn20.Hidden = true;
			ultraGridColumn21.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn21.Header.VisiblePosition = 24;
			ultraGridColumn21.Hidden = true;
			ultraGridColumn22.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn22.Header.VisiblePosition = 21;
			ultraGridColumn22.Hidden = true;
			ultraGridColumn23.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn23.Header.VisiblePosition = 22;
			ultraGridColumn23.Hidden = true;
			ultraGridColumn24.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn24.Header.VisiblePosition = 19;
			ultraGridColumn24.Hidden = true;
			ultraGridColumn25.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn25.Header.VisiblePosition = 25;
			ultraGridColumn25.Hidden = true;
			ultraGridColumn26.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn26.Header.VisiblePosition = 17;
			ultraGridColumn26.Hidden = true;
			ultraGridColumn27.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn27.Header.Caption = "Estado";
			ultraGridColumn27.Header.VisiblePosition = 15;
			ultraGridColumn27.Hidden = true;
			ultraGridColumn28.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn28.Header.Caption = "Estado";
			ultraGridColumn28.Header.VisiblePosition = 20;
			ultraGridColumn28.Width = 200;
			ultraGridColumn29.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn29.Header.VisiblePosition = 28;
			ultraGridColumn29.Hidden = true;
			ultraGridColumn30.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn30.Header.VisiblePosition = 29;
			ultraGridColumn30.Hidden = true;
			ultraGridColumn31.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn31.Header.VisiblePosition = 30;
			ultraGridColumn31.Hidden = true;
			ultraGridColumn32.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn32.Header.VisiblePosition = 31;
			ultraGridColumn32.Hidden = true;
			ultraGridColumn33.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn33.Header.VisiblePosition = 32;
			ultraGridColumn33.Hidden = true;
			ultraGridColumn34.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn34.Header.VisiblePosition = 33;
			ultraGridColumn34.Hidden = true;
			ultraGridColumn35.Header.VisiblePosition = 34;
			ultraGridColumn35.Hidden = true;
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
																										 ultraGridColumn35});
			this.ugProdAutorizar.DisplayLayout.BandsSerializer.Add(ultraGridBand1);
			appearance2.ForeColor = System.Drawing.Color.Black;
			appearance2.ForeColorDisabled = System.Drawing.Color.Black;
			this.ugProdAutorizar.DisplayLayout.Override.ActiveRowAppearance = appearance2;
			this.ugProdAutorizar.DisplayLayout.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.No;
			this.ugProdAutorizar.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.False;
			this.ugProdAutorizar.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.False;
			appearance3.BackColor = System.Drawing.Color.Transparent;
			this.ugProdAutorizar.DisplayLayout.Override.CardAreaAppearance = appearance3;
			appearance4.ForeColor = System.Drawing.Color.Black;
			appearance4.ForeColorDisabled = System.Drawing.Color.Black;
			this.ugProdAutorizar.DisplayLayout.Override.GroupByRowAppearance = appearance4;
			appearance5.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance5.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance5.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance5.FontData.BoldAsString = "True";
			appearance5.FontData.Name = "Arial";
			appearance5.FontData.SizeInPoints = 8.5F;
			appearance5.ForeColor = System.Drawing.Color.White;
			appearance5.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.ugProdAutorizar.DisplayLayout.Override.HeaderAppearance = appearance5;
			this.ugProdAutorizar.DisplayLayout.Override.HeaderClickAction = Infragistics.Win.UltraWinGrid.HeaderClickAction.SortSingle;
			appearance6.BackColor = System.Drawing.SystemColors.ControlLight;
			appearance6.BackColor2 = System.Drawing.Color.Gainsboro;
			appearance6.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ugProdAutorizar.DisplayLayout.Override.RowAlternateAppearance = appearance6;
			appearance7.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance7.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance7.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ugProdAutorizar.DisplayLayout.Override.RowSelectorAppearance = appearance7;
			appearance8.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance8.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance8.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ugProdAutorizar.DisplayLayout.Override.SelectedRowAppearance = appearance8;
			this.ugProdAutorizar.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.ugProdAutorizar.Location = new System.Drawing.Point(8, 48);
			this.ugProdAutorizar.Name = "ugProdAutorizar";
			this.ugProdAutorizar.Size = new System.Drawing.Size(1080, 136);
			this.ugProdAutorizar.TabIndex = 679;
			this.ugProdAutorizar.BeforeSelectChange += new Infragistics.Win.UltraWinGrid.BeforeSelectChangeEventHandler(this.ugProdAutorizar_BeforeSelectChange);
			this.ugProdAutorizar.InitializeLayout += new Infragistics.Win.UltraWinGrid.InitializeLayoutEventHandler(this.ugProdAutorizar_InitializeLayout);
			// 
			// ultraDataSource1
			// 
			ultraDataColumn1.DataType = typeof(int);
			ultraDataColumn4.DataType = typeof(System.DateTime);
			ultraDataColumn5.DataType = typeof(System.DateTime);
			ultraDataColumn6.DataType = typeof(int);
			ultraDataColumn9.DataType = typeof(System.DateTime);
			ultraDataColumn10.DataType = typeof(int);
			ultraDataColumn14.DataType = typeof(int);
			ultraDataColumn18.DataType = typeof(System.Decimal);
			ultraDataColumn20.DataType = typeof(int);
			ultraDataColumn22.DataType = typeof(System.Decimal);
			ultraDataColumn23.DataType = typeof(System.Decimal);
			ultraDataColumn27.DataType = typeof(int);
			ultraDataColumn29.DataType = typeof(int);
			ultraDataColumn30.DataType = typeof(int);
			ultraDataColumn31.DataType = typeof(int);
			ultraDataColumn32.DataType = typeof(int);
			ultraDataColumn33.DataType = typeof(int);
			ultraDataColumn34.DataType = typeof(int);
			ultraDataColumn35.DataType = typeof(int);
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
																																 ultraDataColumn35});
			// 
			// label4
			// 
			this.label4.AutoSize = true;
			this.label4.Font = new System.Drawing.Font("Tahoma", 8F);
			this.label4.Location = new System.Drawing.Point(8, 11);
			this.label4.Name = "label4";
			this.label4.Size = new System.Drawing.Size(36, 16);
			this.label4.TabIndex = 677;
			this.label4.Text = "Buscar";
			this.label4.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// btnVer
			// 
			this.btnVer.CausesValidation = false;
			this.btnVer.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnVer.ImageIndex = 2;
			this.btnVer.ImageList = this.imageList1;
			this.btnVer.Location = new System.Drawing.Point(280, 8);
			this.btnVer.Name = "btnVer";
			this.btnVer.Size = new System.Drawing.Size(77, 23);
			this.btnVer.TabIndex = 1;
			this.btnVer.Text = "Ver";
			this.btnVer.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.toolTip1.SetToolTip(this.btnVer, "Busca información...");
			this.btnVer.Click += new System.EventHandler(this.btnVer_Click);
			// 
			// imageList1
			// 
			this.imageList1.ImageSize = new System.Drawing.Size(16, 16);
			this.imageList1.ImageStream = ((System.Windows.Forms.ImageListStreamer)(resources.GetObject("imageList1.ImageStream")));
			this.imageList1.TransparentColor = System.Drawing.Color.Transparent;
			// 
			// btnNuevo
			// 
			this.btnNuevo.CausesValidation = false;
			this.btnNuevo.Enabled = false;
			this.btnNuevo.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnNuevo.ImageIndex = 50;
			this.btnNuevo.ImageList = this.imageList1;
			this.btnNuevo.Location = new System.Drawing.Point(8, 592);
			this.btnNuevo.Name = "btnNuevo";
			this.btnNuevo.TabIndex = 808;
			this.btnNuevo.Text = "Nuevo";
			this.btnNuevo.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.toolTip1.SetToolTip(this.btnNuevo, "Nuevo informe técnico");
			this.btnNuevo.Click += new System.EventHandler(this.btnNuevo_Click);
			// 
			// btnGuardar
			// 
			this.btnGuardar.CausesValidation = false;
			this.btnGuardar.Enabled = false;
			this.btnGuardar.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnGuardar.ImageIndex = 9;
			this.btnGuardar.ImageList = this.imageList1;
			this.btnGuardar.Location = new System.Drawing.Point(88, 592);
			this.btnGuardar.Name = "btnGuardar";
			this.btnGuardar.TabIndex = 807;
			this.btnGuardar.Text = "Guardar";
			this.btnGuardar.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.toolTip1.SetToolTip(this.btnGuardar, "Grabar Información");
			this.btnGuardar.Click += new System.EventHandler(this.btnGuardar_Click);
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
			// lblContador
			// 
			this.lblContador.AutoSize = true;
			this.lblContador.Font = new System.Drawing.Font("Tahoma", 10F);
			this.lblContador.ForeColor = System.Drawing.Color.Firebrick;
			this.lblContador.Location = new System.Drawing.Point(584, 9);
			this.lblContador.Name = "lblContador";
			this.lblContador.Size = new System.Drawing.Size(0, 20);
			this.lblContador.TabIndex = 745;
			this.lblContador.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
			// 
			// txtBuscar
			// 
			appearance9.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtBuscar.Appearance = appearance9;
			this.txtBuscar.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtBuscar.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtBuscar.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.txtBuscar.Location = new System.Drawing.Point(64, 8);
			this.txtBuscar.MaxLength = 100;
			this.txtBuscar.Name = "txtBuscar";
			this.txtBuscar.Size = new System.Drawing.Size(200, 22);
			this.txtBuscar.TabIndex = 0;
			this.txtBuscar.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtBuscar_KeyDown);
			// 
			// groupBox2
			// 
			this.groupBox2.Location = new System.Drawing.Point(0, 32);
			this.groupBox2.Name = "groupBox2";
			this.groupBox2.Size = new System.Drawing.Size(1096, 8);
			this.groupBox2.TabIndex = 747;
			this.groupBox2.TabStop = false;
			// 
			// dtFechaIDT
			// 
			appearance10.ForeColorDisabled = System.Drawing.Color.Black;
			this.dtFechaIDT.Appearance = appearance10;
			this.dtFechaIDT.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			dateButton1.Caption = "Today";
			this.dtFechaIDT.DateButtons.Add(dateButton1);
			this.dtFechaIDT.Enabled = false;
			this.dtFechaIDT.Format = "dd/MM/yyyy";
			this.dtFechaIDT.Location = new System.Drawing.Point(816, 193);
			this.dtFechaIDT.Name = "dtFechaIDT";
			this.dtFechaIDT.NonAutoSizeHeight = 23;
			this.dtFechaIDT.Size = new System.Drawing.Size(112, 21);
			this.dtFechaIDT.SpinButtonsVisible = true;
			this.dtFechaIDT.TabIndex = 794;
			this.dtFechaIDT.Value = ((object)(resources.GetObject("dtFechaIDT.Value")));
			// 
			// label9
			// 
			this.label9.AutoSize = true;
			this.label9.BackColor = System.Drawing.Color.Transparent;
			this.label9.Location = new System.Drawing.Point(640, 195);
			this.label9.Name = "label9";
			this.label9.Size = new System.Drawing.Size(159, 17);
			this.label9.TabIndex = 793;
			this.label9.Text = "Ingreso Departamento Tecnico";
			this.label9.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// txtBodega
			// 
			this.txtBodega.AcceptsTab = true;
			appearance11.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtBodega.Appearance = appearance11;
			this.txtBodega.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtBodega.Enabled = false;
			this.txtBodega.Location = new System.Drawing.Point(264, 192);
			this.txtBodega.Name = "txtBodega";
			this.txtBodega.Size = new System.Drawing.Size(160, 22);
			this.txtBodega.TabIndex = 792;
			// 
			// label7
			// 
			this.label7.AutoSize = true;
			this.label7.BackColor = System.Drawing.Color.Transparent;
			this.label7.Location = new System.Drawing.Point(200, 195);
			this.label7.Name = "label7";
			this.label7.Size = new System.Drawing.Size(41, 17);
			this.label7.TabIndex = 791;
			this.label7.Text = "Bodega";
			this.label7.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// txtLocal
			// 
			this.txtLocal.AcceptsTab = true;
			appearance12.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtLocal.Appearance = appearance12;
			this.txtLocal.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtLocal.Enabled = false;
			this.txtLocal.Location = new System.Drawing.Point(264, 232);
			this.txtLocal.Name = "txtLocal";
			this.txtLocal.Size = new System.Drawing.Size(160, 22);
			this.txtLocal.TabIndex = 790;
			// 
			// label25
			// 
			this.label25.AutoSize = true;
			this.label25.BackColor = System.Drawing.Color.Transparent;
			this.label25.Location = new System.Drawing.Point(200, 235);
			this.label25.Name = "label25";
			this.label25.Size = new System.Drawing.Size(30, 17);
			this.label25.TabIndex = 789;
			this.label25.Text = "Local";
			this.label25.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// label20
			// 
			this.label20.AutoSize = true;
			this.label20.BackColor = System.Drawing.Color.Transparent;
			this.label20.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label20.Location = new System.Drawing.Point(8, 520);
			this.label20.Name = "label20";
			this.label20.Size = new System.Drawing.Size(33, 17);
			this.label20.TabIndex = 788;
			this.label20.Text = "Notas";
			this.label20.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
			// 
			// txtNotas
			// 
			this.txtNotas.AcceptsTab = true;
			appearance13.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtNotas.Appearance = appearance13;
			this.txtNotas.AutoSize = false;
			this.txtNotas.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtNotas.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtNotas.Enabled = false;
			this.txtNotas.Location = new System.Drawing.Point(72, 504);
			this.txtNotas.Multiline = true;
			this.txtNotas.Name = "txtNotas";
			this.txtNotas.Scrollbars = System.Windows.Forms.ScrollBars.Vertical;
			this.txtNotas.Size = new System.Drawing.Size(440, 66);
			this.txtNotas.TabIndex = 787;
			this.txtNotas.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtNotas_KeyDown);
			// 
			// txtDepreciacion
			// 
			appearance14.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtDepreciacion.Appearance = appearance14;
			this.txtDepreciacion.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtDepreciacion.Enabled = false;
			this.txtDepreciacion.FormatString = "#,##0.00";
			this.txtDepreciacion.Location = new System.Drawing.Point(800, 552);
			this.txtDepreciacion.MaxValue = 100;
			this.txtDepreciacion.MinValue = 0;
			this.txtDepreciacion.Name = "txtDepreciacion";
			this.txtDepreciacion.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtDepreciacion.PromptChar = ' ';
			this.txtDepreciacion.Size = new System.Drawing.Size(80, 22);
			this.txtDepreciacion.TabIndex = 786;
			this.txtDepreciacion.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtDepreciacion_KeyDown);
			this.txtDepreciacion.Click += new System.EventHandler(this.txtDepreciacion_Click);
			this.txtDepreciacion.Validated += new System.EventHandler(this.txtDepreciacion_Validated);
			this.txtDepreciacion.Enter += new System.EventHandler(this.txtDepreciacion_Enter);
			// 
			// label19
			// 
			this.label19.AutoSize = true;
			this.label19.BackColor = System.Drawing.Color.Transparent;
			this.label19.Location = new System.Drawing.Point(704, 555);
			this.label19.Name = "label19";
			this.label19.Size = new System.Drawing.Size(83, 17);
			this.label19.TabIndex = 785;
			this.label19.Text = "% Depreciacion";
			this.label19.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// txtSolicita
			// 
			this.txtSolicita.AcceptsTab = true;
			appearance15.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtSolicita.Appearance = appearance15;
			this.txtSolicita.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtSolicita.Enabled = false;
			this.txtSolicita.Location = new System.Drawing.Point(72, 320);
			this.txtSolicita.Name = "txtSolicita";
			this.txtSolicita.Size = new System.Drawing.Size(325, 22);
			this.txtSolicita.TabIndex = 778;
			// 
			// label22
			// 
			this.label22.AutoSize = true;
			this.label22.BackColor = System.Drawing.Color.Transparent;
			this.label22.Location = new System.Drawing.Point(8, 406);
			this.label22.Name = "label22";
			this.label22.Size = new System.Drawing.Size(86, 17);
			this.label22.TabIndex = 777;
			this.label22.Text = "Informe Tecnico";
			this.label22.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// txtInformeTecnico
			// 
			this.txtInformeTecnico.AcceptsTab = true;
			appearance16.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtInformeTecnico.Appearance = appearance16;
			this.txtInformeTecnico.AutoSize = false;
			this.txtInformeTecnico.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtInformeTecnico.Enabled = false;
			this.txtInformeTecnico.Location = new System.Drawing.Point(160, 392);
			this.txtInformeTecnico.Multiline = true;
			this.txtInformeTecnico.Name = "txtInformeTecnico";
			this.txtInformeTecnico.Scrollbars = System.Windows.Forms.ScrollBars.Vertical;
			this.txtInformeTecnico.Size = new System.Drawing.Size(800, 44);
			this.txtInformeTecnico.TabIndex = 776;
			// 
			// txtEstadoFisico
			// 
			this.txtEstadoFisico.AcceptsTab = true;
			appearance17.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtEstadoFisico.Appearance = appearance17;
			this.txtEstadoFisico.AutoSize = false;
			this.txtEstadoFisico.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtEstadoFisico.Enabled = false;
			this.txtEstadoFisico.Location = new System.Drawing.Point(160, 344);
			this.txtEstadoFisico.Multiline = true;
			this.txtEstadoFisico.Name = "txtEstadoFisico";
			this.txtEstadoFisico.Scrollbars = System.Windows.Forms.ScrollBars.Vertical;
			this.txtEstadoFisico.Size = new System.Drawing.Size(800, 44);
			this.txtEstadoFisico.TabIndex = 773;
			// 
			// dtFecha
			// 
			appearance18.ForeColorDisabled = System.Drawing.Color.Black;
			this.dtFecha.Appearance = appearance18;
			this.dtFecha.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			dateButton2.Caption = "Today";
			this.dtFecha.DateButtons.Add(dateButton2);
			this.dtFecha.Enabled = false;
			this.dtFecha.Format = "dd/MM/yyyy";
			this.dtFecha.Location = new System.Drawing.Point(504, 193);
			this.dtFecha.Name = "dtFecha";
			this.dtFecha.NonAutoSizeHeight = 23;
			this.dtFecha.Size = new System.Drawing.Size(112, 21);
			this.dtFecha.SpinButtonsVisible = true;
			this.dtFecha.TabIndex = 772;
			this.dtFecha.Value = ((object)(resources.GetObject("dtFecha.Value")));
			// 
			// dtFechaFactura
			// 
			appearance19.ForeColorDisabled = System.Drawing.Color.Black;
			this.dtFechaFactura.Appearance = appearance19;
			this.dtFechaFactura.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			dateButton3.Caption = "Today";
			this.dtFechaFactura.DateButtons.Add(dateButton3);
			this.dtFechaFactura.Enabled = false;
			this.dtFechaFactura.Format = "dd/MM/yyyy";
			this.dtFechaFactura.Location = new System.Drawing.Point(504, 233);
			this.dtFechaFactura.Name = "dtFechaFactura";
			this.dtFechaFactura.NonAutoSizeHeight = 23;
			this.dtFechaFactura.Size = new System.Drawing.Size(112, 21);
			this.dtFechaFactura.SpinButtonsVisible = true;
			this.dtFechaFactura.TabIndex = 771;
			this.dtFechaFactura.Value = ((object)(resources.GetObject("dtFechaFactura.Value")));
			// 
			// txtMovitoNC
			// 
			this.txtMovitoNC.AcceptsTab = true;
			appearance20.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtMovitoNC.Appearance = appearance20;
			this.txtMovitoNC.AutoSize = false;
			this.txtMovitoNC.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtMovitoNC.Enabled = false;
			this.txtMovitoNC.Location = new System.Drawing.Point(160, 440);
			this.txtMovitoNC.Multiline = true;
			this.txtMovitoNC.Name = "txtMovitoNC";
			this.txtMovitoNC.Scrollbars = System.Windows.Forms.ScrollBars.Vertical;
			this.txtMovitoNC.Size = new System.Drawing.Size(800, 44);
			this.txtMovitoNC.TabIndex = 770;
			// 
			// txtSerial
			// 
			this.txtSerial.AcceptsTab = true;
			appearance21.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtSerial.Appearance = appearance21;
			this.txtSerial.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtSerial.Enabled = false;
			this.txtSerial.Location = new System.Drawing.Point(744, 280);
			this.txtSerial.Name = "txtSerial";
			this.txtSerial.Size = new System.Drawing.Size(184, 22);
			this.txtSerial.TabIndex = 769;
			// 
			// txtDescripcion
			// 
			this.txtDescripcion.AcceptsTab = true;
			appearance22.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtDescripcion.Appearance = appearance22;
			this.txtDescripcion.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtDescripcion.Enabled = false;
			this.txtDescripcion.Location = new System.Drawing.Point(360, 280);
			this.txtDescripcion.Name = "txtDescripcion";
			this.txtDescripcion.Size = new System.Drawing.Size(300, 22);
			this.txtDescripcion.TabIndex = 768;
			// 
			// txtCodigo
			// 
			this.txtCodigo.AcceptsTab = true;
			appearance23.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtCodigo.Appearance = appearance23;
			this.txtCodigo.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtCodigo.Enabled = false;
			this.txtCodigo.Location = new System.Drawing.Point(72, 280);
			this.txtCodigo.Name = "txtCodigo";
			this.txtCodigo.Size = new System.Drawing.Size(200, 22);
			this.txtCodigo.TabIndex = 767;
			// 
			// txtFactura
			// 
			this.txtFactura.AcceptsTab = true;
			appearance24.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtFactura.Appearance = appearance24;
			this.txtFactura.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtFactura.Enabled = false;
			this.txtFactura.Location = new System.Drawing.Point(72, 232);
			this.txtFactura.Name = "txtFactura";
			this.txtFactura.Size = new System.Drawing.Size(100, 22);
			this.txtFactura.TabIndex = 766;
			// 
			// txtIdentificacion
			// 
			this.txtIdentificacion.AcceptsTab = true;
			appearance25.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtIdentificacion.Appearance = appearance25;
			this.txtIdentificacion.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtIdentificacion.Enabled = false;
			this.txtIdentificacion.Location = new System.Drawing.Point(72, 256);
			this.txtIdentificacion.Name = "txtIdentificacion";
			this.txtIdentificacion.Size = new System.Drawing.Size(100, 22);
			this.txtIdentificacion.TabIndex = 765;
			// 
			// txtCliente
			// 
			this.txtCliente.AcceptsTab = true;
			appearance26.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtCliente.Appearance = appearance26;
			this.txtCliente.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtCliente.Enabled = false;
			this.txtCliente.Location = new System.Drawing.Point(360, 256);
			this.txtCliente.Name = "txtCliente";
			this.txtCliente.Size = new System.Drawing.Size(300, 22);
			this.txtCliente.TabIndex = 764;
			// 
			// label15
			// 
			this.label15.AutoSize = true;
			this.label15.BackColor = System.Drawing.Color.Transparent;
			this.label15.Location = new System.Drawing.Point(8, 323);
			this.label15.Name = "label15";
			this.label15.Size = new System.Drawing.Size(42, 17);
			this.label15.TabIndex = 763;
			this.label15.Text = "Tecnico";
			this.label15.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// label14
			// 
			this.label14.AutoSize = true;
			this.label14.BackColor = System.Drawing.Color.Transparent;
			this.label14.Location = new System.Drawing.Point(8, 358);
			this.label14.Name = "label14";
			this.label14.Size = new System.Drawing.Size(135, 17);
			this.label14.TabIndex = 762;
			this.label14.Text = "Estado Fisico del Producto";
			this.label14.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// label13
			// 
			this.label13.AutoSize = true;
			this.label13.BackColor = System.Drawing.Color.Transparent;
			this.label13.Location = new System.Drawing.Point(8, 195);
			this.label13.Name = "label13";
			this.label13.Size = new System.Drawing.Size(42, 17);
			this.label13.TabIndex = 761;
			this.label13.Text = "Ingreso";
			this.label13.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// label12
			// 
			this.label12.AutoSize = true;
			this.label12.BackColor = System.Drawing.Color.Transparent;
			this.label12.Location = new System.Drawing.Point(448, 195);
			this.label12.Name = "label12";
			this.label12.Size = new System.Drawing.Size(34, 17);
			this.label12.TabIndex = 760;
			this.label12.Text = "Fecha";
			this.label12.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// label10
			// 
			this.label10.AutoSize = true;
			this.label10.BackColor = System.Drawing.Color.Transparent;
			this.label10.Location = new System.Drawing.Point(448, 235);
			this.label10.Name = "label10";
			this.label10.Size = new System.Drawing.Size(44, 17);
			this.label10.TabIndex = 759;
			this.label10.Text = "Vendido";
			this.label10.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// label8
			// 
			this.label8.AutoSize = true;
			this.label8.BackColor = System.Drawing.Color.Transparent;
			this.label8.Location = new System.Drawing.Point(8, 454);
			this.label8.Name = "label8";
			this.label8.Size = new System.Drawing.Size(121, 17);
			this.label8.TabIndex = 758;
			this.label8.Text = "Motivo Nota De Credito";
			this.label8.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// label6
			// 
			this.label6.AutoSize = true;
			this.label6.BackColor = System.Drawing.Color.Transparent;
			this.label6.Location = new System.Drawing.Point(680, 283);
			this.label6.Name = "label6";
			this.label6.Size = new System.Drawing.Size(32, 17);
			this.label6.TabIndex = 757;
			this.label6.Text = "Serial";
			this.label6.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// label5
			// 
			this.label5.AutoSize = true;
			this.label5.BackColor = System.Drawing.Color.Transparent;
			this.label5.Location = new System.Drawing.Point(288, 283);
			this.label5.Name = "label5";
			this.label5.Size = new System.Drawing.Size(61, 17);
			this.label5.TabIndex = 756;
			this.label5.Text = "Descripción";
			this.label5.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// label3
			// 
			this.label3.AutoSize = true;
			this.label3.BackColor = System.Drawing.Color.Transparent;
			this.label3.Location = new System.Drawing.Point(8, 235);
			this.label3.Name = "label3";
			this.label3.Size = new System.Drawing.Size(41, 17);
			this.label3.TabIndex = 755;
			this.label3.Text = "Factura";
			this.label3.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// label2
			// 
			this.label2.AutoSize = true;
			this.label2.BackColor = System.Drawing.Color.Transparent;
			this.label2.Location = new System.Drawing.Point(8, 283);
			this.label2.Name = "label2";
			this.label2.Size = new System.Drawing.Size(38, 17);
			this.label2.TabIndex = 753;
			this.label2.Text = "Codigo";
			this.label2.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// label1
			// 
			this.label1.AutoSize = true;
			this.label1.BackColor = System.Drawing.Color.Transparent;
			this.label1.Location = new System.Drawing.Point(8, 259);
			this.label1.Name = "label1";
			this.label1.Size = new System.Drawing.Size(38, 17);
			this.label1.TabIndex = 752;
			this.label1.Text = "Cedula";
			this.label1.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// lblComprobante
			// 
			this.lblComprobante.AutoSize = true;
			this.lblComprobante.BackColor = System.Drawing.Color.Transparent;
			this.lblComprobante.Location = new System.Drawing.Point(288, 259);
			this.lblComprobante.Name = "lblComprobante";
			this.lblComprobante.Size = new System.Drawing.Size(38, 17);
			this.lblComprobante.TabIndex = 751;
			this.lblComprobante.Text = "Cliente";
			this.lblComprobante.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// txtIngreso
			// 
			this.txtIngreso.AcceptsTab = true;
			appearance27.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtIngreso.Appearance = appearance27;
			this.txtIngreso.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtIngreso.Enabled = false;
			this.txtIngreso.Location = new System.Drawing.Point(72, 192);
			this.txtIngreso.Name = "txtIngreso";
			this.txtIngreso.Size = new System.Drawing.Size(100, 22);
			this.txtIngreso.TabIndex = 754;
			// 
			// groupBox1
			// 
			this.groupBox1.Location = new System.Drawing.Point(0, 216);
			this.groupBox1.Name = "groupBox1";
			this.groupBox1.Size = new System.Drawing.Size(1096, 8);
			this.groupBox1.TabIndex = 795;
			this.groupBox1.TabStop = false;
			// 
			// groupBox3
			// 
			this.groupBox3.Location = new System.Drawing.Point(0, 304);
			this.groupBox3.Name = "groupBox3";
			this.groupBox3.Size = new System.Drawing.Size(1096, 8);
			this.groupBox3.TabIndex = 796;
			this.groupBox3.TabStop = false;
			// 
			// groupBox4
			// 
			this.groupBox4.Location = new System.Drawing.Point(0, 488);
			this.groupBox4.Name = "groupBox4";
			this.groupBox4.Size = new System.Drawing.Size(1096, 8);
			this.groupBox4.TabIndex = 797;
			this.groupBox4.TabStop = false;
			// 
			// txtRepuesto
			// 
			appearance28.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtRepuesto.Appearance = appearance28;
			this.txtRepuesto.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtRepuesto.Enabled = false;
			this.txtRepuesto.FormatString = "#,##0.00";
			this.txtRepuesto.Location = new System.Drawing.Point(672, 320);
			this.txtRepuesto.MaxValue = 100000;
			this.txtRepuesto.MinValue = 0;
			this.txtRepuesto.Name = "txtRepuesto";
			this.txtRepuesto.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtRepuesto.PromptChar = ' ';
			this.txtRepuesto.Size = new System.Drawing.Size(96, 22);
			this.txtRepuesto.TabIndex = 799;
			// 
			// label16
			// 
			this.label16.AutoSize = true;
			this.label16.BackColor = System.Drawing.Color.Transparent;
			this.label16.Location = new System.Drawing.Point(600, 323);
			this.label16.Name = "label16";
			this.label16.Size = new System.Drawing.Size(51, 17);
			this.label16.TabIndex = 798;
			this.label16.Text = "Repuesto";
			this.label16.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// txtAFacturar
			// 
			appearance29.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtAFacturar.Appearance = appearance29;
			this.txtAFacturar.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtAFacturar.Enabled = false;
			this.txtAFacturar.FormatString = "#,##0.00";
			this.txtAFacturar.Location = new System.Drawing.Point(864, 320);
			this.txtAFacturar.MaxValue = 100000;
			this.txtAFacturar.MinValue = 0;
			this.txtAFacturar.Name = "txtAFacturar";
			this.txtAFacturar.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtAFacturar.PromptChar = ' ';
			this.txtAFacturar.Size = new System.Drawing.Size(96, 22);
			this.txtAFacturar.TabIndex = 801;
			// 
			// label17
			// 
			this.label17.AutoSize = true;
			this.label17.BackColor = System.Drawing.Color.Transparent;
			this.label17.Location = new System.Drawing.Point(792, 323);
			this.label17.Name = "label17";
			this.label17.Size = new System.Drawing.Size(56, 17);
			this.label17.TabIndex = 800;
			this.label17.Text = "A Facturar";
			this.label17.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// label36
			// 
			this.label36.AutoSize = true;
			this.label36.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label36.Location = new System.Drawing.Point(528, 507);
			this.label36.Name = "label36";
			this.label36.Size = new System.Drawing.Size(38, 17);
			this.label36.TabIndex = 802;
			this.label36.Text = "Estado";
			this.label36.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// cmbEstadosDT
			// 
			this.cmbEstadosDT.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbEstadosDT.Cursor = System.Windows.Forms.Cursors.Default;
			ultraGridColumn36.Header.VisiblePosition = 0;
			ultraGridColumn36.Hidden = true;
			ultraGridColumn37.Header.VisiblePosition = 1;
			ultraGridColumn37.Width = 272;
			ultraGridBand2.Columns.AddRange(new object[] {
																										 ultraGridColumn36,
																										 ultraGridColumn37});
			this.cmbEstadosDT.DisplayLayout.BandsSerializer.Add(ultraGridBand2);
			this.cmbEstadosDT.DisplayMember = "Nombre";
			this.cmbEstadosDT.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbEstadosDT.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbEstadosDT.Enabled = false;
			this.cmbEstadosDT.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbEstadosDT.Location = new System.Drawing.Point(608, 504);
			this.cmbEstadosDT.MaxDropDownItems = 30;
			this.cmbEstadosDT.Name = "cmbEstadosDT";
			this.cmbEstadosDT.Size = new System.Drawing.Size(272, 22);
			this.cmbEstadosDT.TabIndex = 803;
			this.cmbEstadosDT.ValueMember = "idEstado";
			this.cmbEstadosDT.KeyDown += new System.Windows.Forms.KeyEventHandler(this.cmbEstadosDT_KeyDown);
			this.cmbEstadosDT.ValueChanged += new System.EventHandler(this.cmbEstadosDT_ValueChanged);
			// 
			// txtGarantia
			// 
			this.txtGarantia.AcceptsTab = true;
			appearance30.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtGarantia.Appearance = appearance30;
			this.txtGarantia.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtGarantia.Enabled = false;
			this.txtGarantia.Location = new System.Drawing.Point(520, 320);
			this.txtGarantia.Name = "txtGarantia";
			this.txtGarantia.Size = new System.Drawing.Size(56, 22);
			this.txtGarantia.TabIndex = 805;
			// 
			// label18
			// 
			this.label18.AutoSize = true;
			this.label18.BackColor = System.Drawing.Color.Transparent;
			this.label18.Location = new System.Drawing.Point(456, 323);
			this.label18.Name = "label18";
			this.label18.Size = new System.Drawing.Size(46, 17);
			this.label18.TabIndex = 804;
			this.label18.Text = "Garantia";
			this.label18.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// lblEstado
			// 
			this.lblEstado.AutoSize = true;
			this.lblEstado.Font = new System.Drawing.Font("Tahoma", 12F, System.Drawing.FontStyle.Bold);
			this.lblEstado.ForeColor = System.Drawing.Color.Firebrick;
			this.lblEstado.Location = new System.Drawing.Point(680, 232);
			this.lblEstado.Name = "lblEstado";
			this.lblEstado.Size = new System.Drawing.Size(0, 23);
			this.lblEstado.TabIndex = 806;
			this.lblEstado.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
			// 
			// btnSalir
			// 
			this.btnSalir.CausesValidation = false;
			this.btnSalir.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnSalir.ImageIndex = 62;
			this.btnSalir.ImageList = this.imageList1;
			this.btnSalir.Location = new System.Drawing.Point(248, 592);
			this.btnSalir.Name = "btnSalir";
			this.btnSalir.TabIndex = 810;
			this.btnSalir.Text = "&Salir";
			this.btnSalir.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnSalir.Click += new System.EventHandler(this.btnSalir_Click);
			// 
			// btnCancelar
			// 
			this.btnCancelar.CausesValidation = false;
			this.btnCancelar.Enabled = false;
			this.btnCancelar.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnCancelar.ImageIndex = 10;
			this.btnCancelar.ImageList = this.imageList1;
			this.btnCancelar.Location = new System.Drawing.Point(168, 592);
			this.btnCancelar.Name = "btnCancelar";
			this.btnCancelar.TabIndex = 809;
			this.btnCancelar.Text = "&Cancelar";
			this.btnCancelar.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnCancelar.Click += new System.EventHandler(this.btnCancelar_Click);
			// 
			// groupBox5
			// 
			this.groupBox5.Location = new System.Drawing.Point(0, 576);
			this.groupBox5.Name = "groupBox5";
			this.groupBox5.Size = new System.Drawing.Size(1096, 8);
			this.groupBox5.TabIndex = 811;
			this.groupBox5.TabStop = false;
			// 
			// txtVendedor
			// 
			this.txtVendedor.AcceptsTab = true;
			appearance31.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtVendedor.Appearance = appearance31;
			this.txtVendedor.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtVendedor.Enabled = false;
			this.txtVendedor.Location = new System.Drawing.Point(744, 256);
			this.txtVendedor.Name = "txtVendedor";
			this.txtVendedor.Size = new System.Drawing.Size(300, 22);
			this.txtVendedor.TabIndex = 813;
			// 
			// label11
			// 
			this.label11.AutoSize = true;
			this.label11.BackColor = System.Drawing.Color.Transparent;
			this.label11.Location = new System.Drawing.Point(680, 259);
			this.label11.Name = "label11";
			this.label11.Size = new System.Drawing.Size(52, 17);
			this.label11.TabIndex = 812;
			this.label11.Text = "Vendedor";
			this.label11.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// txtPrecio
			// 
			appearance32.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtPrecio.Appearance = appearance32;
			this.txtPrecio.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtPrecio.Enabled = false;
			this.txtPrecio.FormatString = "#,##0.00";
			this.txtPrecio.Location = new System.Drawing.Point(992, 280);
			this.txtPrecio.MaxValue = 100000;
			this.txtPrecio.MinValue = 0;
			this.txtPrecio.Name = "txtPrecio";
			this.txtPrecio.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtPrecio.PromptChar = ' ';
			this.txtPrecio.Size = new System.Drawing.Size(96, 22);
			this.txtPrecio.TabIndex = 815;
			// 
			// label21
			// 
			this.label21.AutoSize = true;
			this.label21.BackColor = System.Drawing.Color.Transparent;
			this.label21.Location = new System.Drawing.Point(944, 283);
			this.label21.Name = "label21";
			this.label21.Size = new System.Drawing.Size(34, 17);
			this.label21.TabIndex = 814;
			this.label21.Text = "Precio";
			this.label21.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// ultraDataSource2
			// 
			ultraDataColumn36.DataType = typeof(int);
			this.ultraDataSource2.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn36,
																																 ultraDataColumn37});
			// 
			// label23
			// 
			this.label23.AutoSize = true;
			this.label23.Location = new System.Drawing.Point(528, 530);
			this.label23.Name = "label23";
			this.label23.Size = new System.Drawing.Size(37, 17);
			this.label23.TabIndex = 816;
			this.label23.Text = "Mótivo";
			this.label23.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// cmbMotivos
			// 
			appearance33.ForeColorDisabled = System.Drawing.Color.Black;
			this.cmbMotivos.Appearance = appearance33;
			this.cmbMotivos.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbMotivos.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbMotivos.DataSource = this.ultraDataSource2;
			ultraGridColumn38.Header.VisiblePosition = 0;
			ultraGridColumn38.Hidden = true;
			ultraGridColumn39.Header.VisiblePosition = 1;
			ultraGridColumn39.Width = 272;
			ultraGridBand3.Columns.AddRange(new object[] {
																										 ultraGridColumn38,
																										 ultraGridColumn39});
			this.cmbMotivos.DisplayLayout.BandsSerializer.Add(ultraGridBand3);
			this.cmbMotivos.DisplayMember = "Motivo";
			this.cmbMotivos.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbMotivos.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbMotivos.Enabled = false;
			this.cmbMotivos.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbMotivos.Location = new System.Drawing.Point(608, 528);
			this.cmbMotivos.Name = "cmbMotivos";
			this.cmbMotivos.Size = new System.Drawing.Size(272, 21);
			this.cmbMotivos.TabIndex = 817;
			this.cmbMotivos.ValueMember = "idMotivo";
			this.cmbMotivos.KeyDown += new System.Windows.Forms.KeyEventHandler(this.cmbMotivos_KeyDown);
			// 
			// label24
			// 
			this.label24.AutoSize = true;
			this.label24.BackColor = System.Drawing.Color.Transparent;
			this.label24.Location = new System.Drawing.Point(528, 555);
			this.label24.Name = "label24";
			this.label24.Size = new System.Drawing.Size(71, 17);
			this.label24.TabIndex = 818;
			this.label24.Text = "% Descuento";
			this.label24.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// txtDescuento
			// 
			appearance34.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtDescuento.Appearance = appearance34;
			this.txtDescuento.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtDescuento.Enabled = false;
			this.txtDescuento.FormatString = "#,##0.00";
			this.txtDescuento.Location = new System.Drawing.Point(608, 552);
			this.txtDescuento.MaxValue = 100;
			this.txtDescuento.MinValue = 0;
			this.txtDescuento.Name = "txtDescuento";
			this.txtDescuento.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtDescuento.PromptChar = ' ';
			this.txtDescuento.Size = new System.Drawing.Size(80, 22);
			this.txtDescuento.TabIndex = 819;
			this.txtDescuento.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtDescuento_KeyDown);
			this.txtDescuento.Click += new System.EventHandler(this.txtDescuento_Click);
			this.txtDescuento.Validated += new System.EventHandler(this.txtDescuento_Validated);
			this.txtDescuento.Enter += new System.EventHandler(this.txtDescuento_Enter);
			// 
			// optTipo
			// 
			appearance35.ForeColorDisabled = System.Drawing.Color.Black;
			this.optTipo.Appearance = appearance35;
			this.optTipo.BorderStyle = Infragistics.Win.UIElementBorderStyle.Solid;
			this.optTipo.Enabled = false;
			this.optTipo.ItemAppearance = appearance36;
			this.optTipo.ItemOrigin = new System.Drawing.Point(10, 0);
			valueListItem1.DataValue = 1;
			valueListItem1.DisplayText = "Inventario";
			valueListItem2.DataValue = 2;
			valueListItem2.DisplayText = "Desc. en Ventas";
			this.optTipo.Items.Add(valueListItem1);
			this.optTipo.Items.Add(valueListItem2);
			this.optTipo.ItemSpacingVertical = 10;
			this.optTipo.Location = new System.Drawing.Point(888, 551);
			this.optTipo.Name = "optTipo";
			this.optTipo.Size = new System.Drawing.Size(200, 24);
			this.optTipo.TabIndex = 820;
			this.optTipo.KeyDown += new System.Windows.Forms.KeyEventHandler(this.optTipo_KeyDown);
			// 
			// label26
			// 
			this.label26.AutoSize = true;
			this.label26.Location = new System.Drawing.Point(888, 530);
			this.label26.Name = "label26";
			this.label26.Size = new System.Drawing.Size(44, 17);
			this.label26.TabIndex = 822;
			this.label26.Text = "Número";
			this.label26.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// txtNumero
			// 
			appearance37.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtNumero.Appearance = appearance37;
			this.txtNumero.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtNumero.Enabled = false;
			this.txtNumero.Location = new System.Drawing.Point(968, 527);
			this.txtNumero.MaxLength = 9;
			this.txtNumero.Name = "txtNumero";
			this.txtNumero.Size = new System.Drawing.Size(120, 22);
			this.txtNumero.TabIndex = 821;
			this.txtNumero.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtNumero_KeyDown);
			this.txtNumero.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.txtNumero_KeyPress);
			// 
			// frmXAutorizarDT
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 14);
			this.ClientSize = new System.Drawing.Size(1096, 624);
			this.Controls.Add(this.label26);
			this.Controls.Add(this.label24);
			this.Controls.Add(this.label21);
			this.Controls.Add(this.label11);
			this.Controls.Add(this.label18);
			this.Controls.Add(this.label36);
			this.Controls.Add(this.label17);
			this.Controls.Add(this.label16);
			this.Controls.Add(this.label9);
			this.Controls.Add(this.label7);
			this.Controls.Add(this.label25);
			this.Controls.Add(this.label20);
			this.Controls.Add(this.label19);
			this.Controls.Add(this.label22);
			this.Controls.Add(this.label15);
			this.Controls.Add(this.label14);
			this.Controls.Add(this.label13);
			this.Controls.Add(this.label12);
			this.Controls.Add(this.label10);
			this.Controls.Add(this.label8);
			this.Controls.Add(this.label6);
			this.Controls.Add(this.label5);
			this.Controls.Add(this.label3);
			this.Controls.Add(this.label2);
			this.Controls.Add(this.label1);
			this.Controls.Add(this.lblComprobante);
			this.Controls.Add(this.label4);
			this.Controls.Add(this.label23);
			this.Controls.Add(this.txtNumero);
			this.Controls.Add(this.optTipo);
			this.Controls.Add(this.txtDescuento);
			this.Controls.Add(this.cmbMotivos);
			this.Controls.Add(this.txtPrecio);
			this.Controls.Add(this.txtVendedor);
			this.Controls.Add(this.groupBox5);
			this.Controls.Add(this.btnSalir);
			this.Controls.Add(this.btnCancelar);
			this.Controls.Add(this.btnNuevo);
			this.Controls.Add(this.btnGuardar);
			this.Controls.Add(this.lblEstado);
			this.Controls.Add(this.txtGarantia);
			this.Controls.Add(this.cmbEstadosDT);
			this.Controls.Add(this.txtAFacturar);
			this.Controls.Add(this.txtRepuesto);
			this.Controls.Add(this.groupBox4);
			this.Controls.Add(this.groupBox3);
			this.Controls.Add(this.groupBox1);
			this.Controls.Add(this.dtFechaIDT);
			this.Controls.Add(this.txtBodega);
			this.Controls.Add(this.txtLocal);
			this.Controls.Add(this.txtNotas);
			this.Controls.Add(this.txtDepreciacion);
			this.Controls.Add(this.txtSolicita);
			this.Controls.Add(this.txtInformeTecnico);
			this.Controls.Add(this.txtEstadoFisico);
			this.Controls.Add(this.dtFecha);
			this.Controls.Add(this.dtFechaFactura);
			this.Controls.Add(this.txtMovitoNC);
			this.Controls.Add(this.txtSerial);
			this.Controls.Add(this.txtDescripcion);
			this.Controls.Add(this.txtCodigo);
			this.Controls.Add(this.txtFactura);
			this.Controls.Add(this.txtIdentificacion);
			this.Controls.Add(this.txtCliente);
			this.Controls.Add(this.txtIngreso);
			this.Controls.Add(this.groupBox2);
			this.Controls.Add(this.txtBuscar);
			this.Controls.Add(this.lblContador);
			this.Controls.Add(this.ugProdAutorizar);
			this.Controls.Add(this.btnVer);
			this.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog;
			this.KeyPreview = true;
			this.MaximizeBox = false;
			this.Name = "frmXAutorizarDT";
			this.Text = "Autorizacion de Repuestos y Notas de Credito";
			this.KeyDown += new System.Windows.Forms.KeyEventHandler(this.frmXAutorizarDT_KeyDown);
			this.Closing += new System.ComponentModel.CancelEventHandler(this.frmXAutorizarDT_Closing);
			this.Load += new System.EventHandler(this.frmXAutorizarDT_Load);
			((System.ComponentModel.ISupportInitialize)(this.ugProdAutorizar)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource1)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtBuscar)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.dtFechaIDT)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtBodega)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtLocal)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNotas)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtDepreciacion)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtSolicita)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtInformeTecnico)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtEstadoFisico)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.dtFecha)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.dtFechaFactura)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtMovitoNC)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtSerial)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtDescripcion)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtCodigo)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtFactura)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIdentificacion)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtCliente)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIngreso)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtRepuesto)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtAFacturar)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbEstadosDT)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtGarantia)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtVendedor)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtPrecio)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource2)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbMotivos)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtDescuento)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.optTipo)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNumero)).EndInit();
			this.ResumeLayout(false);

		}
		#endregion

		private void cdsSeteoF_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			cdsSeteoF.Schema.Connections[0].ConnectionString = miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);
		}

		int idIngreso = 0;
		int iEstado = 0;
		int idPersonal = 0;
		bool bEdicion = false;
		bool bNumeroAutomatico = false;

		private void LimpiaControles()
		{
			idIngreso = 0;
			this.txtIngreso.Text = "";
			this.txtBodega.Text = "";
			this.dtFecha.Value = System.DBNull.Value;
			this.dtFechaIDT.Value = System.DBNull.Value;

			this.txtFactura.Text = "";
			this.txtLocal.Text = "";
			this.dtFechaFactura.Value = System.DBNull.Value;

			this.txtIdentificacion.Text = "";
			this.txtCliente.Text = "";
			this.txtVendedor.Text = "";
			
			this.txtCodigo.Text = "";
			this.txtDescripcion.Text = "";
			this.txtSerial.Text = "";
			this.txtPrecio.Value = 0.00m;
						
			this.txtSolicita.Text = "";
			this.txtGarantia.Text = "";
			this.txtRepuesto.Value = 0.00m;
			this.txtAFacturar.Value = 0.00m;

			this.txtEstadoFisico.Text = "";
			this.txtInformeTecnico.Text = "";
			this.txtMovitoNC.Text = "";
			
			iEstado = 0;
			this.lblEstado.Text = "";
		}

		private void txtDepreciacion_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
			if (e.KeyChar == 13) this.txtNotas.Focus();
		}
		
		private void frmXAutorizarDT_Load(object sender, System.EventArgs e)
		{
			Consulta();

			idPersonal = Funcion.iEscalarSQL(cdsSeteoF, string.Format("If Exists(Select idPersonal From Nomina Where Codigo In (Select Nombre From Usuario Where idUsuario = {0})) Select idPersonal From Nomina Where Codigo In (Select Nombre From Usuario Where idUsuario = {0}) Else Select 0", MenuLatinium.IdUsuario));
			
			this.cmbMotivos.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "Select idMotivo, Motivo From MotivosDevolucion Where idTipoFactura = 5 And idMotivo In (2, 3, 6, 36) Order By Motivo");

			bNumeroAutomatico = Funcion.bEscalarSQL(cdsSeteoF, ("Select Automatico From CompraNumero Where idTipoFactura = 5"), true);
		}

		private void Consulta ()
		{
			string sSQL = string.Format("Exec DT_ProdustosXAutoriza '{0}'", this.txtBuscar.Text.ToString().Trim());
			FuncionesProcedimientos.dsGeneral(sSQL, this.ugProdAutorizar);

			this.lblContador.Text = this.ugProdAutorizar.Rows.Count + " REGISTROS ENCONTRADOS";
		}

		private void btnVer_Click(object sender, System.EventArgs e)
		{
			Consulta();
		}

		private void ugProdAutorizar_BeforeSelectChange(object sender, Infragistics.Win.UltraWinGrid.BeforeSelectChangeEventArgs e)
		{
			LimpiaControles();	
		
			idIngreso = (int)this.ugProdAutorizar.ActiveRow.Cells["idIngreso"].Value;
			this.txtIngreso.Text = this.ugProdAutorizar.ActiveRow.Cells["Ingreso"].Value.ToString().Trim();
			this.txtBodega.Text = this.ugProdAutorizar.ActiveRow.Cells["Bodega"].Value.ToString().Trim();
			if (this.ugProdAutorizar.ActiveRow.Cells["Fecha"].Value != System.DBNull.Value) this.dtFecha.Value = Convert.ToDateTime(this.ugProdAutorizar.ActiveRow.Cells["Fecha"].Value);
			if (this.ugProdAutorizar.ActiveRow.Cells["FechaRecepcion"].Value != System.DBNull.Value) this.dtFechaIDT.Value = Convert.ToDateTime(this.ugProdAutorizar.ActiveRow.Cells["FechaRecepcion"].Value);

			this.txtFactura.Text = this.ugProdAutorizar.ActiveRow.Cells["Numero"].Value.ToString().Trim();
			this.txtLocal.Text = this.ugProdAutorizar.ActiveRow.Cells["Local"].Value.ToString().Trim();
			if (this.ugProdAutorizar.ActiveRow.Cells["FechaFactura"].Value != System.DBNull.Value) this.dtFechaFactura.Value = Convert.ToDateTime(this.ugProdAutorizar.ActiveRow.Cells["FechaFactura"].Value);

			this.txtIdentificacion.Text = this.ugProdAutorizar.ActiveRow.Cells["Ruc"].Value.ToString().Trim();
			this.txtCliente.Text = this.ugProdAutorizar.ActiveRow.Cells["Cliente"].Value.ToString().Trim();
			this.txtVendedor.Text = this.ugProdAutorizar.ActiveRow.Cells["Vendedor"].Value.ToString().Trim();
			
			this.txtCodigo.Text = this.ugProdAutorizar.ActiveRow.Cells["Codigo"].Value.ToString().Trim();
			this.txtDescripcion.Text = this.ugProdAutorizar.ActiveRow.Cells["Articulo"].Value.ToString().Trim();
			this.txtSerial.Text = this.ugProdAutorizar.ActiveRow.Cells["Serial"].Value.ToString().Trim();
			this.txtPrecio.Value = Convert.ToDecimal(this.ugProdAutorizar.ActiveRow.Cells["Precio"].Value);

			this.txtSolicita.Text = this.ugProdAutorizar.ActiveRow.Cells["Solicita"].Value.ToString().Trim();
			this.txtGarantia.Text = this.ugProdAutorizar.ActiveRow.Cells["Garantia"].Value.ToString().Trim();
			this.txtRepuesto.Value = Convert.ToDecimal(this.ugProdAutorizar.ActiveRow.Cells["ValorRepuesto"].Value);
			this.txtAFacturar.Value = Convert.ToDecimal(this.ugProdAutorizar.ActiveRow.Cells["ValorFactura"].Value);

			this.txtEstadoFisico.Text = this.ugProdAutorizar.ActiveRow.Cells["EstadoFisico"].Value.ToString().Trim();
			this.txtInformeTecnico.Text = this.ugProdAutorizar.ActiveRow.Cells["InformeTecnico"].Value.ToString().Trim();			
			this.txtMovitoNC.Text = this.ugProdAutorizar.ActiveRow.Cells["MotivoNC"].Value.ToString().Trim();						
			
			iEstado = (int)this.ugProdAutorizar.ActiveRow.Cells["EstadoIngreso"].Value;
			this.lblEstado.Text = this.ugProdAutorizar.ActiveRow.Cells["EstadoM"].Value.ToString().Trim();			

			this.btnNuevo.Enabled = true;
		}

		private void txtBuscar_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Enter) this.btnVer_Click(sender, e);
		}

		private void btnSalir_Click(object sender, System.EventArgs e)
		{
			this.Close();
		}

		private void btnCancelar_Click(object sender, System.EventArgs e)
		{
			this.LimpiaControles();

			idIngreso = 0;
			iEstado = 0;
			bEdicion = false;

			this.txtBuscar.Enabled = true;
			this.btnVer.Enabled = true;
			this.ugProdAutorizar.Enabled = true;

			this.txtNotas.Value = "";
			this.txtNotas.Enabled = false;
			this.cmbEstadosDT.Value = System.DBNull.Value;
			this.cmbEstadosDT.Enabled = false;
			this.txtDepreciacion.Value = 0.00m;
			this.txtDepreciacion.Enabled = false;
						
			this.btnGuardar.Enabled = false;
			this.btnCancelar.Enabled = false;
			
			this.Consulta();
		}

		private void btnNuevo_Click(object sender, System.EventArgs e)
		{
			this.txtBuscar.Enabled = false;
			this.btnVer.Enabled = false;
			this.ugProdAutorizar.Enabled = false;

			this.txtNotas.Enabled = true;			
			this.cmbEstadosDT.Enabled = true;

			if (iEstado == 12 || iEstado == 18) 
				this.txtDepreciacion.Enabled = true;

			if (iEstado == 8) this.cmbEstadosDT.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "Exec DTListaEstados 7");
			if (iEstado == 12) this.cmbEstadosDT.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "Exec DTListaEstados 8");
			if (iEstado == 18) this.cmbEstadosDT.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "Exec DTListaEstados 11");
			
			bEdicion = true;

			this.btnGuardar.Enabled = true;
			this.btnCancelar.Enabled = true;

			this.txtNotas.Focus();
		}

		private void btnGuardar_Click(object sender, System.EventArgs e)
		{
			if (!Validacion.vbTexto(this.txtNotas, 20, 1000, "Notas")) return;

			if (!Validacion.vbComboVacio(this.cmbEstadosDT, "Seleccione un estado")) return;

			#region Validacion - Numeracion Notas de Credito
			int idNotaCredito = 0;

			if ((int)this.cmbEstadosDT.Value == 13 || (int)this.cmbEstadosDT.Value == 31)
			{
				if (iEstado == 12 || iEstado == 18)
				{
					if (!Validacion.vbComboVacio(this.cmbMotivos, "Seleccione un motivo de nota de credito")) return;
				
					if (!Validacion.vbNumeracion(bNumeroAutomatico, idNotaCredito, 5, 0, this.txtNumero, cdsSeteoF)) return;

					if (this.optTipo.CheckedIndex == - 1)
					{
						MessageBox.Show("Seleccione una opcion", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
						this.optTipo.Focus();
						return;
					}
				}
			}
			
			string sNumero = "";
			if (iEstado == 8) sNumero = Funcion.sEscalarSQL(cdsSeteoF, string.Format("Exec NumeracionLocales 59, {0}, 0", 0));
			if (iEstado == 12 || iEstado == 18) sNumero = Funcion.sEscalarSQL(cdsSeteoF, string.Format("Exec NumeracionLocales 60, {0}, 0", 0));
			#endregion Validacion - Numeracion Notas de Credito

			Funcion.EjecutaSQL(cdsSeteoF, string.Format("Update DTIngreso Set BodegaDestino = Case When EstadoIngreso In (12, 18) Then 20 Else 0 End, PorcnDepreciacion = {0}, NumAutRepuesto = Case EstadoIngreso When 8 Then '{2}' Else '' End, NumAutNC = Case When EstadoIngreso In (12, 18) Then '{2}' Else '' End Where idIngreso = {1}", Convert.ToDecimal(this.txtDepreciacion.Value), idIngreso, sNumero));
			
			if ((int)this.cmbEstadosDT.Value == 13 || (int)this.cmbEstadosDT.Value == 31)			
				if (iEstado == 12 || iEstado == 18) 
					idNotaCredito = this.CreaNotaCredito();

			string sSQL = string.Format("Exec GuardaDT_Procesos {0}, 20, {1}, {2}, '{3}', 0, '', {4}, 0, 0, 0, '{5}', '19000101', 0", 
				idIngreso, (int)this.cmbEstadosDT.Value, idPersonal, this.txtNotas.Text.ToString().Trim(), idNotaCredito, sNumero);		
			Funcion.EjecutaSQL(cdsSeteoF, sSQL);

			if (iEstado == 12 || iEstado == 18) 
			{
				if (idNotaCredito > 0) 
				{
					if (DialogResult.OK == MessageBox.Show(string.Format("Se genero correctamente la nota de credito n.: '{0}',\n\n¿Visualizar el documento?", this.txtNumero.Text.ToString()), "Point Technology", MessageBoxButtons.OKCancel, MessageBoxIcon.Question))
					{
						frmDevolucionVenta NC = new frmDevolucionVenta(idNotaCredito);
						NC.Show();
					}
				}
			}

			this.btnCancelar_Click(sender, e);
		}

		private int CreaNotaCredito()
		{
			int iDiasFactura = Funcion.iEscalarSQL(cdsSeteoF, string.Format("Select DATEDIFF(DAY, '{0}', '{1}')", 
				Convert.ToDateTime(this.dtFechaFactura.Value).ToString("yyyyMMdd"), Convert.ToDateTime(this.dtFecha.Value).ToString("yyyyMMdd")));
			
//			int iBodega = 1;
//			int idProyecto = Funcion.iEscalarSQL(cdsSeteoF, "Select idProyecto From Proyecto Where Matriz = 1");
			//			if (iDiasFactura < 180)
			//			{
			int iBodega = (int)this.ugProdAutorizar.ActiveRow.Cells["LocalFactura"].Value;
			int idProyecto = Funcion.iEscalarSQL(cdsSeteoF, string.Format("Select idProyecto From Bodega Where Bodega = {0}", iBodega));
			//			}
			
			string sSerie = "";
			string sAutorizacion = "";
			DateTime dtFechaCaducidad = DateTime.Today;

			SqlDataReader dr = Funcion.rEscalarSQL(cdsSeteoF, "Select SerieNotaDeCredito, AutNotaDeCredito, CaducaNotaDeCredito From Seteo", true);
			dr.Read();
			if (dr.HasRows)
			{
				if (dr.GetValue(0) != System.DBNull.Value) sSerie = dr.GetString(0);
				if (dr.GetValue(1) != System.DBNull.Value) sAutorizacion = dr.GetString(1);
				if (dr.GetValue(2) != System.DBNull.Value) dtFechaCaducidad = dr.GetDateTime(2);
			}
			dr.Close();

			decimal dSubTotalIva = Convert.ToDecimal(this.txtPrecio.Value);
			decimal dDescuento = Convert.ToDecimal(this.txtDescuento.Value);
			decimal dDescCalculado = (Convert.ToDecimal(this.txtPrecio.Value) * Convert.ToDecimal(this.txtDescuento.Value))/100.00m;
			decimal dBaseImponible = Convert.ToDecimal(this.txtPrecio.Value) + (Convert.ToDecimal(this.txtPrecio.Value) * Convert.ToDecimal(this.txtDescuento.Value))/100.00m;
			decimal dIva = Funcion.decEscalarSQL(cdsSeteoF, string.Format("Select dbo.RetornaImpuestoIVAVigente('{0}', 1)", Convert.ToDateTime(this.dtFechaFactura.Value).ToString("yyyyMMdd")));
			decimal dValorIVA = (dBaseImponible * dIva) / 100.00m;
			decimal dTotal = dBaseImponible + dValorIVA;

			int idArticulo = (int)this.ugProdAutorizar.ActiveRow.Cells["idArticulo"].Value;
			string sCodigo = this.txtCodigo.Value.ToString();
			string sArticulo = this.txtDescripcion.Value.ToString();

			int idSerial = (int)this.ugProdAutorizar.ActiveRow.Cells["idSerial"].Value;
			int idDetCompraN = 0;
			int idDetalleSerialOrigen = 0;

			if ((int)this.optTipo.Value == 1)
			{
				string sSQL = string.Format("Select ds.idDetalleSerial From DetalleSeriales ds Inner Join Seriales s On s.idSerial = ds.idSerial Inner Join DetCompra dc On dc.idDetCompra = ds.idDetCompra Inner Join Compra c On c.idCompra = dc.idCompra Where ds.idSerial = {0} And c.idCompra = {1}",
					idSerial, (int)this.ugProdAutorizar.ActiveRow.Cells["idFVenta"].Value);
				idDetalleSerialOrigen = Funcion.iEscalarSQL(cdsSeteoF, sSQL, true);				
			}

			if ((int)this.optTipo.Value == 2)
			{				
				SqlDataReader drA = Funcion.rEscalarSQL(cdsSeteoF, "Select Codigo, Articulo from Articulo where idArticulo = 2244", true);
				drA.Read();
				if (drA.HasRows)
				{
					idArticulo = 2244;
					sCodigo = drA.GetString(0);
					sArticulo = drA.GetString(1);
				}
			}

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

					if (bNumeroAutomatico)
					{
						oCmdActualiza.CommandText = "Exec NumeracionLocales 5, 1, 1";
						this.txtNumero.Text = oCmdActualiza.ExecuteScalar().ToString();
					}

					#region Guarda Nota de Credito
					string sSQL = string.Format("Exec GrabaMaestroFactura {0}, {1}, {2}, {3}, {4}, {5}, {6}, {7}, {8}, {9}, {10}, {11}, '{12}', '{13}', {14}, {15}, '{16}', '{17}', '{18}', '{19}', '{20}', '{21}', {22}, '{23}', '{24}', {25}, '{26}', {27}, {28}, {29}, {30}, '{31}', {32}, {33}, {34}, {35}, {36}, {37}, {38}, {39}, {40}, {41}, {42}, {43}, {44}, {45}, '{46}', {47}, {48}, {49}, {50}, {51}, {52}, {53}, {54}, {55}, {56}, {57}, {58}, '{59}', {60}, {61}, '{62}', '{63}', {64}, '{65}', {66}, {67}, '{68}', {69}", 
						0, (int)this.ugProdAutorizar.ActiveRow.Cells["idFVenta"].Value, 1, 5, 4, 5, false,
						1, (int)this.ugProdAutorizar.ActiveRow.Cells["idFormaPago"].Value, iBodega, idProyecto, (int)this.ugProdAutorizar.ActiveRow.Cells["idVendedor"].Value, 
						this.txtNumero.Text.ToString(), this.ugProdAutorizar.ActiveRow.Cells["Numero"].Value.ToString(), 
						(int)this.ugProdAutorizar.ActiveRow.Cells["idTipoRuc"].Value, (int)this.ugProdAutorizar.ActiveRow.Cells["idCliente"].Value, 
						this.ugProdAutorizar.ActiveRow.Cells["Ruc"].Value.ToString(), this.ugProdAutorizar.ActiveRow.Cells["Cliente"].Value.ToString(), 
						DateTime.Today.ToString("yyyyMMdd"), sSerie, sAutorizacion, dtFechaCaducidad.ToString("yyyyMMdd"),
						false, DateTime.Today.ToString("yyyyMMdd"), DateTime.Today.ToString("yyyyMMdd HH:mm"), 0, this.txtInformeTecnico.Text.ToString(),
						false, 0.00m, 0, 0.00m, DateTime.Today.ToString("yyyyMMdd"), 
						0.00m, 0.00m, dSubTotalIva, dDescuento, dValorIVA, dTotal, 
						false, false, false, false, false,
						true, 20, (int)this.cmbMotivos.Value, this.txtNotas.Text.ToString(), dTotal, 
						0.00m, 0.00m, 0.00m, 0.00m, 
						false, 0.00m, (int)this.ugProdAutorizar.ActiveRow.Cells["idEntidadFinanciera"].Value, 
						0, 0, 0, 0, "", false,
						1, "", "", 0, "", 0.00m, 0, "", 0.00m);
					oCmdActualiza.CommandText = sSQL;
					int idCompra = (int)oCmdActualiza.ExecuteScalar();
			
					string sSQLDetalle = string.Format("Exec GuardaDetalleFactura {0}, {1}, {2}, {3}, {4}, '{5}', '{6}', {7}, {8}, {9}, {10}, {11}, {12}, {13}, {14}, {15}, {16}, {17}, {18}, {19}, {20}, {21}",
						0, 0, idCompra, idArticulo, 1, sCodigo, sArticulo, 1, Convert.ToDecimal(this.txtPrecio.Value), dIva, dDescCalculado, dDescuento, Convert.ToDecimal(this.txtPrecio.Value),
						true, 1, true, 1, true, false, 0, 0, 0.00m);
					oCmdActualiza.CommandText = sSQLDetalle;
					idDetCompraN = (int)oCmdActualiza.ExecuteScalar();
					#endregion Guarda Nota de Credito

					oCmdActualiza.CommandText = string.Format("Exec FE_ClaveDeAcceso {0}, 1", idCompra);
					string sClaveAcceso = oCmdActualiza.ExecuteScalar().ToString();

					string cadInvertida = Facturacion.invertirCadena(sClaveAcceso.Substring(0, 48));
					string ValM11 = Facturacion.Modulo11(cadInvertida).ToString().Trim();
					string sClaveSRI = sClaveAcceso.ToString().Substring(0, 48) + ValM11;

					oCmdActualiza.CommandText = string.Format("Update Compra Set AutFactura = '{0}' Where idCompra = {1}", sClaveSRI, idCompra);
					oCmdActualiza.ExecuteNonQuery();

					#region Asiento
					string sSQLAsiento = string.Format("Exec CreaAsiDevVenta {0}", idCompra);
					oCmdActualiza.CommandText = sSQLAsiento;
					oCmdActualiza.ExecuteNonQuery();			
					#endregion Asiento

					#region Seriales
					string sSQLGDetalleSeriales = string.Format("Exec GuardaDetalleSeriales {0}, {1}, {2}, 20, '', 1, 'MOVIMIENTO REGISTRADO DESDE SERVICIO TECNICO', 17, 0, 0, 1", 
						idSerial, idDetCompraN, idDetalleSerialOrigen);
					oCmdActualiza.CommandText = sSQLGDetalleSeriales;
					oCmdActualiza.ExecuteNonQuery();
																
					string sSQLActEstSerial = string.Format("Update Seriales Set Estado = 17, Bodega = 20 Where idSerial = {0}", idSerial);
					oCmdActualiza.CommandText = sSQLActEstSerial;																
					oCmdActualiza.ExecuteNonQuery();
					#endregion Seriales

					#region Autorización de Pagos Devolucion de Dinero
					if ((int)this.cmbMotivos.Value == 6)
					{	
						string sSQLAP = string.Format("Exec CreaAutorizacionesDePago 14, {0}, '{1}', '{2}', {3}, 2, '1 DE 1', 'DEVOLUCION DE DINERO CON NOTA DE CREDITO DESDE SERVICIO TECNICO', 0", 
							idCompra, DateTime.Today.ToString("yyyyMMdd"), DateTime.Today.ToString("yyyyMMdd"), Math.Round(dTotal, 2));
						oCmdActualiza.CommandText = sSQLAP;
						oCmdActualiza.ExecuteNonQuery();
					}					
					#endregion Autorización de Pagos Devolucion de Dinero

					oTransaction.Commit();
					
					return idCompra;
				}
				catch (Exception ex)
				{
					oTransaction.Rollback();
					
					MessageBox.Show(string.Format("Commit Exception Type : {0} {1}", ex.GetType(), ex.Message), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);											

					return 0;
				}
				finally
				{
					oConexion.Close();					
				}
			}
		}

		private void txtNotas_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Enter) this.cmbEstadosDT.Focus();
		}

		private void txtDepreciacion_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Enter) this.optTipo.Focus();
		}

		private void cmbEstadosDT_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Enter) if (this.cmbMotivos.Enabled) this.cmbMotivos.Focus(); else this.btnGuardar.Focus();
		}

		private void frmXAutorizarDT_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
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

		private void frmXAutorizarDT_Closing(object sender, System.ComponentModel.CancelEventArgs e)
		{
			if (bEdicion)
			{
				MessageBox.Show("No puede salir mientras edita un informe", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				e.Cancel = true;
				return;
			}
		}

		private void cmbEstadosDT_ValueChanged(object sender, System.EventArgs e)
		{
			if (!bEdicion) return;

			this.cmbMotivos.Value = System.DBNull.Value;
			this.txtNumero.Text = "";
			this.txtDescuento.Value = 0.00m;
			this.txtDepreciacion.Value = 0.00m;
			this.optTipo.CheckedIndex = -1;

			this.cmbMotivos.Enabled = false;
			this.txtNumero.Enabled = false;
			this.txtDescuento.Enabled = false;
			this.txtDepreciacion.Enabled = false;
			this.optTipo.Enabled = false;
			
			if (iEstado == 12 || iEstado == 18)
			{
				this.cmbMotivos.Enabled = true;
				if (!bNumeroAutomatico) 
					this.txtNumero.Enabled = true;
				this.txtDescuento.Enabled = true;
				this.txtDepreciacion.Enabled = true;
				this.optTipo.Enabled = true;
			}
		}

		private void cmbMotivos_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Enter) this.txtNumero.Focus();
		}

		private void txtNumero_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Enter) this.txtDescuento.Focus();
		}

		private void txtDescuento_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Enter) this.txtDepreciacion.Focus();
		}

		private void optTipo_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Enter) this.btnGuardar.Focus();
		}

		private void txtDescuento_Validated(object sender, System.EventArgs e)
		{
			if (this.txtDescuento.Value == System.DBNull.Value) this.txtDescuento.Value = 0.00m;
		}

		private void txtDepreciacion_Validated(object sender, System.EventArgs e)
		{
			if (this.txtDepreciacion.Value == System.DBNull.Value) this.txtDepreciacion.Value = 0.00m;
		}

		private void txtNumero_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
			Funcion.CamposNumericos(sender, e);
		}

		private void txtDescuento_Click(object sender, System.EventArgs e)
		{
			this.txtDescuento.SelectAll();
		}

		private void txtDescuento_Enter(object sender, System.EventArgs e)
		{
			this.txtDescuento.SelectAll();
		}

		private void txtDepreciacion_Click(object sender, System.EventArgs e)
		{
			this.txtDepreciacion.SelectAll();
		}

		private void txtDepreciacion_Enter(object sender, System.EventArgs e)
		{
			this.txtDepreciacion.SelectAll();
		}

		private void ugProdAutorizar_InitializeLayout(object sender, Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs e)
		{
		
		}
	}
}
