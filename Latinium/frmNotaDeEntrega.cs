using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using DevExpress.XtraEditors;
using CrystalDecisions.Shared;
using System.Data;
using System.Data.SqlClient;
using CrystalDecisions;
using CrystalDecisions.CrystalReports.Engine;
using CrystalDecisions.ReportSource;
using System.Text;
using System.Net;
using System.Web;
using System.Xml;
using System.Xml.Schema;
using System.Security.Cryptography.Xml;
using System.Security.Cryptography;
using System.Security.Cryptography.X509Certificates;

namespace Latinium
{
	/// <summary>
	/// Summary description for frmNotaDeEntrega.
	/// </summary>
	public class frmNotaDeEntrega : DevExpress.XtraEditors.XtraForm
	{
		private System.Windows.Forms.Label lblLocal;
		private System.Windows.Forms.Label label1;
		private Infragistics.Win.UltraWinSchedule.UltraCalendarCombo dtFecha;
		private System.Windows.Forms.Label lblIdentificacion;
		private System.Windows.Forms.Label lblCliente;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbTipoRuc;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtRuc;
		private System.Windows.Forms.Label label3;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtNombre;
		private System.Windows.Forms.Label lblVendedor;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbVendedor;
		private System.Windows.Forms.Label label2;
		private Infragistics.Win.UltraWinGrid.UltraGrid ugArticulos;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource3;
		private C1.Data.C1DataSet cdsSeteoF;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtbuscar;
		private System.Windows.Forms.Button btnNuevo;
		private System.Windows.Forms.Button btnConsultar;
		private System.Windows.Forms.Button btnEditar;
		private System.Windows.Forms.Button btnImprimir;
		private System.Windows.Forms.Button btnGuardar;
		private System.Windows.Forms.Button btnAnular;
		private System.Windows.Forms.Button btnCancelar;
		private System.Windows.Forms.Button btnSalir;
		private System.Windows.Forms.GroupBox groupBox1;
		private Infragistics.Win.UltraWinGrid.UltraGrid ugDetalleFactura;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource4;
		public Infragistics.Win.UltraWinEditors.UltraNumericEditor txtidCliente;
		public Infragistics.Win.UltraWinEditors.UltraNumericEditor txtidCompra;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbBodega;
		private System.Windows.Forms.Label label4;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtTotal;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtIvaTotal;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtIva;
		private System.Windows.Forms.Label lblNumero;
		private Infragistics.Win.UltraWinEditors.UltraCheckEditor chkEstado;
		private System.Windows.Forms.GroupBox groupBox2;
		private System.Windows.Forms.Button btnVer;
		private System.Windows.Forms.Label lblEstado;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtNumero;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource1;
		private C1.Data.C1DataSet cdsSeteo;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtEstado;
		private System.Windows.Forms.Label label5;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbGrupo;
		/// <summary>
		/// Required designer variable.
		/// </summary>
		private System.ComponentModel.Container components = null;

		public frmNotaDeEntrega()
		{
			//
			// Required for Windows Form Designer support
			//
			InitializeComponent();

			//
			// TODO: Add any constructor code after InitializeComponent call
			//
		}

		int idCompraV = 0;

		public frmNotaDeEntrega(int IdCompraV)
		{
			InitializeComponent();
			idCompraV = IdCompraV;
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
			Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton dateButton1 = new Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton();
			System.Resources.ResourceManager resources = new System.Resources.ResourceManager(typeof(frmNotaDeEntrega));
			Infragistics.Win.Appearance appearance2 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand1 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn1 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idTipoRuc");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn2 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("TipoRuc");
			Infragistics.Win.Appearance appearance3 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance4 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance5 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand2 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn3 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idPersonal");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn4 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn5 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Codigo");
			Infragistics.Win.Appearance appearance6 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance7 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand3 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn6 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idArticulo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn7 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Codigo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn8 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Articulo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn9 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Precio1");
			Infragistics.Win.Appearance appearance8 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn10 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Precio2");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn11 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("UnboundColumn1", 0);
			Infragistics.Win.Appearance appearance9 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance10 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance11 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance12 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance13 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance14 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance15 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn1 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idArticulo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn2 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Codigo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn3 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Articulo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn4 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Precio1");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn5 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Precio2");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn6 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idArticulo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn7 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idGrupoArticulo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn8 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idSubGrupo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn9 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Articulo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn10 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Precio1");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn11 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Precio2");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn12 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Codigo");
			Infragistics.Win.Appearance appearance16 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand4 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn12 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idDetCompra");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn13 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCompra");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn14 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idArticulo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn15 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Articulo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn16 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Cantidad");
			Infragistics.Win.Appearance appearance17 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn17 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Notas");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn18 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Precio");
			Infragistics.Win.Appearance appearance18 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn19 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Codigo");
			Infragistics.Win.Appearance appearance19 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance20 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance21 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance22 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance23 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance24 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn13 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idDetCompra");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn14 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idCompra");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn15 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idArticulo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn16 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Articulo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn17 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Cantidad");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn18 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Notas");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn19 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Precio");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn20 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Codigo");
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand5 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn20 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idBodega");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn21 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Bodega");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn22 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.Appearance appearance25 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance26 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance27 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance28 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance29 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand6 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn23 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCategoria");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn24 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Categoria");
			this.lblLocal = new System.Windows.Forms.Label();
			this.label1 = new System.Windows.Forms.Label();
			this.dtFecha = new Infragistics.Win.UltraWinSchedule.UltraCalendarCombo();
			this.lblIdentificacion = new System.Windows.Forms.Label();
			this.lblCliente = new System.Windows.Forms.Label();
			this.cmbTipoRuc = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.txtRuc = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.label3 = new System.Windows.Forms.Label();
			this.txtNombre = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.lblVendedor = new System.Windows.Forms.Label();
			this.cmbVendedor = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.label2 = new System.Windows.Forms.Label();
			this.txtbuscar = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.ugArticulos = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.ultraDataSource1 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.ultraDataSource3 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.cdsSeteoF = new C1.Data.C1DataSet();
			this.btnNuevo = new System.Windows.Forms.Button();
			this.btnConsultar = new System.Windows.Forms.Button();
			this.btnEditar = new System.Windows.Forms.Button();
			this.btnImprimir = new System.Windows.Forms.Button();
			this.btnGuardar = new System.Windows.Forms.Button();
			this.btnAnular = new System.Windows.Forms.Button();
			this.btnCancelar = new System.Windows.Forms.Button();
			this.btnSalir = new System.Windows.Forms.Button();
			this.groupBox1 = new System.Windows.Forms.GroupBox();
			this.ugDetalleFactura = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.ultraDataSource4 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.txtidCliente = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtidCompra = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.cmbBodega = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.label4 = new System.Windows.Forms.Label();
			this.txtTotal = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtIvaTotal = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtIva = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.lblNumero = new System.Windows.Forms.Label();
			this.chkEstado = new Infragistics.Win.UltraWinEditors.UltraCheckEditor();
			this.groupBox2 = new System.Windows.Forms.GroupBox();
			this.btnVer = new System.Windows.Forms.Button();
			this.lblEstado = new System.Windows.Forms.Label();
			this.txtNumero = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.cdsSeteo = new C1.Data.C1DataSet();
			this.txtEstado = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.label5 = new System.Windows.Forms.Label();
			this.cmbGrupo = new Infragistics.Win.UltraWinGrid.UltraCombo();
			((System.ComponentModel.ISupportInitialize)(this.dtFecha)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbTipoRuc)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtRuc)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNombre)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbVendedor)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtbuscar)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ugArticulos)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource1)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource3)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ugDetalleFactura)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource4)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtidCliente)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtidCompra)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbBodega)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtTotal)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIvaTotal)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIva)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNumero)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteo)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtEstado)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbGrupo)).BeginInit();
			this.SuspendLayout();
			// 
			// lblLocal
			// 
			this.lblLocal.AutoSize = true;
			this.lblLocal.Location = new System.Drawing.Point(264, 11);
			this.lblLocal.Name = "lblLocal";
			this.lblLocal.Size = new System.Drawing.Size(30, 16);
			this.lblLocal.TabIndex = 821;
			this.lblLocal.Text = "Local";
			this.lblLocal.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// label1
			// 
			this.label1.AutoSize = true;
			this.label1.Location = new System.Drawing.Point(8, 11);
			this.label1.Name = "label1";
			this.label1.Size = new System.Drawing.Size(28, 16);
			this.label1.TabIndex = 823;
			this.label1.Text = "Feha";
			this.label1.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// dtFecha
			// 
			appearance1.ForeColorDisabled = System.Drawing.Color.Black;
			this.dtFecha.Appearance = appearance1;
			this.dtFecha.BorderStyle = Infragistics.Win.UIElementBorderStyle.Solid;
			this.dtFecha.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			dateButton1.Caption = "Today";
			this.dtFecha.DateButtons.Add(dateButton1);
			this.dtFecha.Enabled = false;
			this.dtFecha.Format = "dd/MM/yyyy HH:mm";
			this.dtFecha.Location = new System.Drawing.Point(88, 9);
			this.dtFecha.Name = "dtFecha";
			this.dtFecha.NonAutoSizeHeight = 23;
			this.dtFecha.Size = new System.Drawing.Size(144, 21);
			this.dtFecha.SpinButtonsVisible = true;
			this.dtFecha.TabIndex = 824;
			this.dtFecha.Value = ((object)(resources.GetObject("dtFecha.Value")));
			this.dtFecha.KeyDown += new System.Windows.Forms.KeyEventHandler(this.dtFecha_KeyDown);
			this.dtFecha.Validating += new System.ComponentModel.CancelEventHandler(this.dtFecha_Validating);
			this.dtFecha.ValueChanged += new System.EventHandler(this.dtFecha_ValueChanged);
			// 
			// lblIdentificacion
			// 
			this.lblIdentificacion.AutoSize = true;
			this.lblIdentificacion.Location = new System.Drawing.Point(8, 43);
			this.lblIdentificacion.Name = "lblIdentificacion";
			this.lblIdentificacion.Size = new System.Drawing.Size(71, 16);
			this.lblIdentificacion.TabIndex = 828;
			this.lblIdentificacion.Text = "Identificación";
			this.lblIdentificacion.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// lblCliente
			// 
			this.lblCliente.AutoSize = true;
			this.lblCliente.Location = new System.Drawing.Point(264, 43);
			this.lblCliente.Name = "lblCliente";
			this.lblCliente.Size = new System.Drawing.Size(86, 16);
			this.lblCliente.TabIndex = 827;
			this.lblCliente.Text = "N. Identificación";
			this.lblCliente.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// cmbTipoRuc
			// 
			appearance2.FontData.Name = "Tahoma";
			appearance2.ForeColorDisabled = System.Drawing.Color.Black;
			this.cmbTipoRuc.Appearance = appearance2;
			this.cmbTipoRuc.CausesValidation = false;
			this.cmbTipoRuc.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbTipoRuc.Cursor = System.Windows.Forms.Cursors.Default;
			ultraGridColumn1.Header.VisiblePosition = 0;
			ultraGridColumn1.Hidden = true;
			ultraGridColumn2.Header.VisiblePosition = 1;
			ultraGridColumn2.Width = 233;
			ultraGridBand1.Columns.AddRange(new object[] {
																										 ultraGridColumn1,
																										 ultraGridColumn2});
			this.cmbTipoRuc.DisplayLayout.BandsSerializer.Add(ultraGridBand1);
			this.cmbTipoRuc.DisplayMember = "TipoRuc";
			this.cmbTipoRuc.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbTipoRuc.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbTipoRuc.Enabled = false;
			this.cmbTipoRuc.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbTipoRuc.Location = new System.Drawing.Point(88, 40);
			this.cmbTipoRuc.Name = "cmbTipoRuc";
			this.cmbTipoRuc.Size = new System.Drawing.Size(144, 22);
			this.cmbTipoRuc.TabIndex = 825;
			this.cmbTipoRuc.ValueMember = "idTipoRuc";
			this.cmbTipoRuc.KeyDown += new System.Windows.Forms.KeyEventHandler(this.cmbTipoRuc_KeyDown);
			this.cmbTipoRuc.ValueChanged += new System.EventHandler(this.cmbTipoRuc_ValueChanged);
			this.cmbTipoRuc.InitializeLayout += new Infragistics.Win.UltraWinGrid.InitializeLayoutEventHandler(this.cmbTipoRuc_InitializeLayout);
			// 
			// txtRuc
			// 
			this.txtRuc.AcceptsTab = true;
			appearance3.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtRuc.Appearance = appearance3;
			this.txtRuc.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtRuc.Enabled = false;
			this.txtRuc.Location = new System.Drawing.Point(352, 40);
			this.txtRuc.Name = "txtRuc";
			this.txtRuc.Size = new System.Drawing.Size(128, 22);
			this.txtRuc.TabIndex = 826;
			this.txtRuc.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtRuc_KeyDown);
			this.txtRuc.Validating += new System.ComponentModel.CancelEventHandler(this.txtRuc_Validating);
			this.txtRuc.Validated += new System.EventHandler(this.txtRuc_Validated);
			// 
			// label3
			// 
			this.label3.AutoSize = true;
			this.label3.Location = new System.Drawing.Point(568, 43);
			this.label3.Name = "label3";
			this.label3.Size = new System.Drawing.Size(84, 16);
			this.label3.TabIndex = 830;
			this.label3.Text = "Nombre Factura";
			this.label3.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// txtNombre
			// 
			appearance4.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtNombre.Appearance = appearance4;
			this.txtNombre.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtNombre.Enabled = false;
			this.txtNombre.Location = new System.Drawing.Point(664, 40);
			this.txtNombre.Name = "txtNombre";
			this.txtNombre.Size = new System.Drawing.Size(248, 22);
			this.txtNombre.TabIndex = 829;
			// 
			// lblVendedor
			// 
			this.lblVendedor.AutoSize = true;
			this.lblVendedor.Location = new System.Drawing.Point(568, 11);
			this.lblVendedor.Name = "lblVendedor";
			this.lblVendedor.Size = new System.Drawing.Size(52, 16);
			this.lblVendedor.TabIndex = 832;
			this.lblVendedor.Text = "Vendedor";
			this.lblVendedor.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// cmbVendedor
			// 
			appearance5.FontData.Name = "Tahoma";
			appearance5.ForeColorDisabled = System.Drawing.Color.Black;
			this.cmbVendedor.Appearance = appearance5;
			this.cmbVendedor.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbVendedor.Cursor = System.Windows.Forms.Cursors.Default;
			ultraGridColumn3.Header.VisiblePosition = 0;
			ultraGridColumn3.Hidden = true;
			ultraGridColumn4.Header.VisiblePosition = 1;
			ultraGridColumn4.Width = 335;
			ultraGridColumn5.Header.VisiblePosition = 2;
			ultraGridBand2.Columns.AddRange(new object[] {
																										 ultraGridColumn3,
																										 ultraGridColumn4,
																										 ultraGridColumn5});
			this.cmbVendedor.DisplayLayout.BandsSerializer.Add(ultraGridBand2);
			this.cmbVendedor.DisplayMember = "Nombre";
			this.cmbVendedor.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbVendedor.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbVendedor.Enabled = false;
			this.cmbVendedor.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbVendedor.Location = new System.Drawing.Point(664, 8);
			this.cmbVendedor.Name = "cmbVendedor";
			this.cmbVendedor.Size = new System.Drawing.Size(250, 22);
			this.cmbVendedor.TabIndex = 831;
			this.cmbVendedor.ValueMember = "idPersonal";
			this.cmbVendedor.KeyDown += new System.Windows.Forms.KeyEventHandler(this.cmbVendedor_KeyDown);
			// 
			// label2
			// 
			this.label2.AutoSize = true;
			this.label2.Location = new System.Drawing.Point(680, 91);
			this.label2.Name = "label2";
			this.label2.Size = new System.Drawing.Size(41, 16);
			this.label2.TabIndex = 834;
			this.label2.Text = "Buscar ";
			this.label2.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// txtbuscar
			// 
			appearance6.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtbuscar.Appearance = appearance6;
			this.txtbuscar.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtbuscar.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtbuscar.Enabled = false;
			this.txtbuscar.Location = new System.Drawing.Point(744, 88);
			this.txtbuscar.Name = "txtbuscar";
			this.txtbuscar.Size = new System.Drawing.Size(240, 22);
			this.txtbuscar.TabIndex = 833;
			this.txtbuscar.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtbuscar_KeyDown);
			this.txtbuscar.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.txtbuscar_KeyPress);
			// 
			// ugArticulos
			// 
			this.ugArticulos.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.ugArticulos.Cursor = System.Windows.Forms.Cursors.Default;
			this.ugArticulos.DataSource = this.ultraDataSource1;
			appearance7.BackColor = System.Drawing.Color.White;
			this.ugArticulos.DisplayLayout.Appearance = appearance7;
			this.ugArticulos.DisplayLayout.AutoFitColumns = true;
			ultraGridColumn6.Header.VisiblePosition = 1;
			ultraGridColumn6.Hidden = true;
			ultraGridColumn7.Header.VisiblePosition = 4;
			ultraGridColumn7.Hidden = true;
			ultraGridColumn8.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn8.Header.VisiblePosition = 2;
			ultraGridColumn8.Width = 342;
			ultraGridColumn9.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance8.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn9.CellAppearance = appearance8;
			ultraGridColumn9.Format = "#,##0.00";
			ultraGridColumn9.Header.Caption = "Precio";
			ultraGridColumn9.Header.VisiblePosition = 3;
			ultraGridColumn9.Width = 81;
			ultraGridColumn10.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn10.Format = "#,##0.00";
			ultraGridColumn10.Header.Caption = "Reserva";
			ultraGridColumn10.Header.VisiblePosition = 5;
			ultraGridColumn10.Hidden = true;
			ultraGridColumn10.Width = 80;
			ultraGridColumn11.ButtonDisplayStyle = Infragistics.Win.UltraWinGrid.ButtonDisplayStyle.Always;
			appearance9.Image = ((object)(resources.GetObject("appearance9.Image")));
			ultraGridColumn11.CellButtonAppearance = appearance9;
			ultraGridColumn11.Header.Caption = "....";
			ultraGridColumn11.Header.VisiblePosition = 0;
			ultraGridColumn11.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.Button;
			ultraGridColumn11.Width = 20;
			ultraGridBand3.Columns.AddRange(new object[] {
																										 ultraGridColumn6,
																										 ultraGridColumn7,
																										 ultraGridColumn8,
																										 ultraGridColumn9,
																										 ultraGridColumn10,
																										 ultraGridColumn11});
			this.ugArticulos.DisplayLayout.BandsSerializer.Add(ultraGridBand3);
			appearance10.ForeColor = System.Drawing.Color.Black;
			appearance10.ForeColorDisabled = System.Drawing.Color.Black;
			this.ugArticulos.DisplayLayout.Override.ActiveRowAppearance = appearance10;
			this.ugArticulos.DisplayLayout.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.TemplateOnBottom;
			this.ugArticulos.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.False;
			this.ugArticulos.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.False;
			appearance11.BackColor = System.Drawing.Color.Transparent;
			this.ugArticulos.DisplayLayout.Override.CardAreaAppearance = appearance11;
			appearance12.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance12.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance12.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance12.FontData.BoldAsString = "True";
			appearance12.FontData.Name = "Arial";
			appearance12.FontData.SizeInPoints = 10F;
			appearance12.ForeColor = System.Drawing.Color.White;
			appearance12.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.ugArticulos.DisplayLayout.Override.HeaderAppearance = appearance12;
			appearance13.BackColor = System.Drawing.SystemColors.ControlLight;
			appearance13.BackColor2 = System.Drawing.Color.Gainsboro;
			appearance13.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ugArticulos.DisplayLayout.Override.RowAlternateAppearance = appearance13;
			appearance14.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance14.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance14.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ugArticulos.DisplayLayout.Override.RowSelectorAppearance = appearance14;
			appearance15.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance15.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance15.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ugArticulos.DisplayLayout.Override.SelectedRowAppearance = appearance15;
			this.ugArticulos.Font = new System.Drawing.Font("Arial", 8F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)), true);
			this.ugArticulos.Location = new System.Drawing.Point(680, 120);
			this.ugArticulos.Name = "ugArticulos";
			this.ugArticulos.Size = new System.Drawing.Size(464, 256);
			this.ugArticulos.TabIndex = 836;
			this.ugArticulos.ClickCellButton += new Infragistics.Win.UltraWinGrid.CellEventHandler(this.ugArticulos_ClickCellButton);
			this.ugArticulos.InitializeLayout += new Infragistics.Win.UltraWinGrid.InitializeLayoutEventHandler(this.ugArticulos_InitializeLayout);
			// 
			// ultraDataSource1
			// 
			ultraDataColumn1.DataType = typeof(int);
			ultraDataColumn4.DataType = typeof(System.Decimal);
			ultraDataColumn5.DataType = typeof(System.Decimal);
			this.ultraDataSource1.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn1,
																																 ultraDataColumn2,
																																 ultraDataColumn3,
																																 ultraDataColumn4,
																																 ultraDataColumn5});
			// 
			// ultraDataSource3
			// 
			ultraDataColumn6.DataType = typeof(int);
			ultraDataColumn7.DataType = typeof(int);
			ultraDataColumn8.DataType = typeof(int);
			ultraDataColumn10.DataType = typeof(System.Decimal);
			ultraDataColumn11.DataType = typeof(System.Decimal);
			this.ultraDataSource3.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn6,
																																 ultraDataColumn7,
																																 ultraDataColumn8,
																																 ultraDataColumn9,
																																 ultraDataColumn10,
																																 ultraDataColumn11,
																																 ultraDataColumn12});
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
			// btnNuevo
			// 
			this.btnNuevo.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btnNuevo.BackColor = System.Drawing.SystemColors.Control;
			this.btnNuevo.CausesValidation = false;
			this.btnNuevo.Enabled = false;
			this.btnNuevo.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.btnNuevo.ForeColor = System.Drawing.Color.FromArgb(((System.Byte)(50)), ((System.Byte)(50)), ((System.Byte)(50)));
			this.btnNuevo.Image = ((System.Drawing.Image)(resources.GetObject("btnNuevo.Image")));
			this.btnNuevo.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnNuevo.Location = new System.Drawing.Point(8, 400);
			this.btnNuevo.Name = "btnNuevo";
			this.btnNuevo.Size = new System.Drawing.Size(72, 23);
			this.btnNuevo.TabIndex = 837;
			this.btnNuevo.Text = "Nuevo";
			this.btnNuevo.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnNuevo.Click += new System.EventHandler(this.btnNuevo_Click);
			// 
			// btnConsultar
			// 
			this.btnConsultar.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btnConsultar.BackColor = System.Drawing.SystemColors.Control;
			this.btnConsultar.CausesValidation = false;
			this.btnConsultar.Enabled = false;
			this.btnConsultar.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.btnConsultar.ForeColor = System.Drawing.Color.FromArgb(((System.Byte)(50)), ((System.Byte)(50)), ((System.Byte)(50)));
			this.btnConsultar.Image = ((System.Drawing.Image)(resources.GetObject("btnConsultar.Image")));
			this.btnConsultar.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnConsultar.Location = new System.Drawing.Point(83, 400);
			this.btnConsultar.Name = "btnConsultar";
			this.btnConsultar.Size = new System.Drawing.Size(78, 23);
			this.btnConsultar.TabIndex = 838;
			this.btnConsultar.Text = "Consultar";
			this.btnConsultar.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnConsultar.Click += new System.EventHandler(this.btnConsultar_Click);
			// 
			// btnEditar
			// 
			this.btnEditar.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btnEditar.BackColor = System.Drawing.SystemColors.Control;
			this.btnEditar.CausesValidation = false;
			this.btnEditar.Enabled = false;
			this.btnEditar.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.btnEditar.ForeColor = System.Drawing.Color.FromArgb(((System.Byte)(50)), ((System.Byte)(50)), ((System.Byte)(50)));
			this.btnEditar.Image = ((System.Drawing.Image)(resources.GetObject("btnEditar.Image")));
			this.btnEditar.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnEditar.Location = new System.Drawing.Point(164, 400);
			this.btnEditar.Name = "btnEditar";
			this.btnEditar.Size = new System.Drawing.Size(76, 23);
			this.btnEditar.TabIndex = 841;
			this.btnEditar.Text = "Editar";
			this.btnEditar.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnEditar.Click += new System.EventHandler(this.btnEditar_Click);
			// 
			// btnImprimir
			// 
			this.btnImprimir.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btnImprimir.BackColor = System.Drawing.SystemColors.Control;
			this.btnImprimir.CausesValidation = false;
			this.btnImprimir.Enabled = false;
			this.btnImprimir.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.btnImprimir.ForeColor = System.Drawing.Color.FromArgb(((System.Byte)(50)), ((System.Byte)(50)), ((System.Byte)(50)));
			this.btnImprimir.Image = ((System.Drawing.Image)(resources.GetObject("btnImprimir.Image")));
			this.btnImprimir.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnImprimir.Location = new System.Drawing.Point(243, 400);
			this.btnImprimir.Name = "btnImprimir";
			this.btnImprimir.Size = new System.Drawing.Size(77, 23);
			this.btnImprimir.TabIndex = 840;
			this.btnImprimir.Text = "Imprimir";
			this.btnImprimir.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnImprimir.Click += new System.EventHandler(this.btnImprimir_Click);
			// 
			// btnGuardar
			// 
			this.btnGuardar.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btnGuardar.BackColor = System.Drawing.SystemColors.Control;
			this.btnGuardar.Enabled = false;
			this.btnGuardar.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.btnGuardar.ForeColor = System.Drawing.Color.FromArgb(((System.Byte)(50)), ((System.Byte)(50)), ((System.Byte)(50)));
			this.btnGuardar.Image = ((System.Drawing.Image)(resources.GetObject("btnGuardar.Image")));
			this.btnGuardar.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnGuardar.Location = new System.Drawing.Point(323, 400);
			this.btnGuardar.Name = "btnGuardar";
			this.btnGuardar.Size = new System.Drawing.Size(76, 23);
			this.btnGuardar.TabIndex = 839;
			this.btnGuardar.Text = "Guardar";
			this.btnGuardar.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnGuardar.Click += new System.EventHandler(this.btnGuardar_Click);
			// 
			// btnAnular
			// 
			this.btnAnular.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btnAnular.BackColor = System.Drawing.SystemColors.Control;
			this.btnAnular.CausesValidation = false;
			this.btnAnular.Enabled = false;
			this.btnAnular.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.btnAnular.Image = ((System.Drawing.Image)(resources.GetObject("btnAnular.Image")));
			this.btnAnular.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnAnular.Location = new System.Drawing.Point(402, 400);
			this.btnAnular.Name = "btnAnular";
			this.btnAnular.Size = new System.Drawing.Size(76, 23);
			this.btnAnular.TabIndex = 842;
			this.btnAnular.Text = "Anular";
			this.btnAnular.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnAnular.Click += new System.EventHandler(this.btnAnular_Click);
			// 
			// btnCancelar
			// 
			this.btnCancelar.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btnCancelar.BackColor = System.Drawing.SystemColors.Control;
			this.btnCancelar.CausesValidation = false;
			this.btnCancelar.Enabled = false;
			this.btnCancelar.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.btnCancelar.ForeColor = System.Drawing.Color.FromArgb(((System.Byte)(50)), ((System.Byte)(50)), ((System.Byte)(50)));
			this.btnCancelar.Image = ((System.Drawing.Image)(resources.GetObject("btnCancelar.Image")));
			this.btnCancelar.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnCancelar.Location = new System.Drawing.Point(481, 400);
			this.btnCancelar.Name = "btnCancelar";
			this.btnCancelar.Size = new System.Drawing.Size(72, 23);
			this.btnCancelar.TabIndex = 843;
			this.btnCancelar.Text = "Cancelar";
			this.btnCancelar.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnCancelar.Click += new System.EventHandler(this.btnCancelar_Click);
			// 
			// btnSalir
			// 
			this.btnSalir.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btnSalir.BackColor = System.Drawing.SystemColors.Control;
			this.btnSalir.CausesValidation = false;
			this.btnSalir.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.btnSalir.ForeColor = System.Drawing.Color.FromArgb(((System.Byte)(50)), ((System.Byte)(50)), ((System.Byte)(50)));
			this.btnSalir.Image = ((System.Drawing.Image)(resources.GetObject("btnSalir.Image")));
			this.btnSalir.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnSalir.Location = new System.Drawing.Point(556, 400);
			this.btnSalir.Name = "btnSalir";
			this.btnSalir.TabIndex = 844;
			this.btnSalir.Text = "Salir";
			this.btnSalir.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnSalir.Click += new System.EventHandler(this.btnSalir_Click);
			// 
			// groupBox1
			// 
			this.groupBox1.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.groupBox1.Location = new System.Drawing.Point(-48, 384);
			this.groupBox1.Name = "groupBox1";
			this.groupBox1.Size = new System.Drawing.Size(1200, 8);
			this.groupBox1.TabIndex = 845;
			this.groupBox1.TabStop = false;
			// 
			// ugDetalleFactura
			// 
			this.ugDetalleFactura.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.ugDetalleFactura.Cursor = System.Windows.Forms.Cursors.Default;
			this.ugDetalleFactura.DataSource = this.ultraDataSource4;
			appearance16.BackColor = System.Drawing.Color.White;
			this.ugDetalleFactura.DisplayLayout.Appearance = appearance16;
			ultraGridColumn12.Header.VisiblePosition = 0;
			ultraGridColumn12.Hidden = true;
			ultraGridColumn13.Header.VisiblePosition = 1;
			ultraGridColumn13.Hidden = true;
			ultraGridColumn14.Header.VisiblePosition = 2;
			ultraGridColumn14.Hidden = true;
			ultraGridColumn14.Width = 98;
			ultraGridColumn15.CellActivation = Infragistics.Win.UltraWinGrid.Activation.NoEdit;
			ultraGridColumn15.Header.VisiblePosition = 3;
			ultraGridColumn15.Width = 300;
			appearance17.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn16.CellAppearance = appearance17;
			ultraGridColumn16.Format = "#,##0";
			ultraGridColumn16.Header.VisiblePosition = 4;
			ultraGridColumn16.Width = 40;
			ultraGridColumn17.Case = Infragistics.Win.UltraWinGrid.Case.Upper;
			ultraGridColumn17.Header.VisiblePosition = 6;
			ultraGridColumn17.Width = 400;
			ultraGridColumn18.CellActivation = Infragistics.Win.UltraWinGrid.Activation.NoEdit;
			appearance18.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn18.CellAppearance = appearance18;
			ultraGridColumn18.Format = "#,##0.00";
			ultraGridColumn18.Header.VisiblePosition = 5;
			ultraGridColumn19.Header.VisiblePosition = 7;
			ultraGridColumn19.Hidden = true;
			ultraGridBand4.Columns.AddRange(new object[] {
																										 ultraGridColumn12,
																										 ultraGridColumn13,
																										 ultraGridColumn14,
																										 ultraGridColumn15,
																										 ultraGridColumn16,
																										 ultraGridColumn17,
																										 ultraGridColumn18,
																										 ultraGridColumn19});
			this.ugDetalleFactura.DisplayLayout.BandsSerializer.Add(ultraGridBand4);
			appearance19.ForeColor = System.Drawing.Color.Black;
			appearance19.ForeColorDisabled = System.Drawing.Color.Black;
			this.ugDetalleFactura.DisplayLayout.Override.ActiveRowAppearance = appearance19;
			this.ugDetalleFactura.DisplayLayout.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.TemplateOnBottom;
			this.ugDetalleFactura.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.True;
			this.ugDetalleFactura.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.True;
			appearance20.BackColor = System.Drawing.Color.Transparent;
			this.ugDetalleFactura.DisplayLayout.Override.CardAreaAppearance = appearance20;
			this.ugDetalleFactura.DisplayLayout.Override.CellClickAction = Infragistics.Win.UltraWinGrid.CellClickAction.Edit;
			appearance21.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance21.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance21.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance21.FontData.BoldAsString = "True";
			appearance21.FontData.Name = "Arial";
			appearance21.FontData.SizeInPoints = 10F;
			appearance21.ForeColor = System.Drawing.Color.White;
			appearance21.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.ugDetalleFactura.DisplayLayout.Override.HeaderAppearance = appearance21;
			appearance22.BackColor = System.Drawing.SystemColors.ControlLight;
			appearance22.BackColor2 = System.Drawing.Color.Gainsboro;
			appearance22.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ugDetalleFactura.DisplayLayout.Override.RowAlternateAppearance = appearance22;
			appearance23.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance23.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance23.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ugDetalleFactura.DisplayLayout.Override.RowSelectorAppearance = appearance23;
			appearance24.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance24.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance24.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ugDetalleFactura.DisplayLayout.Override.SelectedRowAppearance = appearance24;
			this.ugDetalleFactura.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.ugDetalleFactura.Location = new System.Drawing.Point(8, 88);
			this.ugDetalleFactura.Name = "ugDetalleFactura";
			this.ugDetalleFactura.Size = new System.Drawing.Size(664, 248);
			this.ugDetalleFactura.TabIndex = 846;
			this.ugDetalleFactura.AfterRowsDeleted += new System.EventHandler(this.ugDetalleFactura_AfterRowsDeleted);
			this.ugDetalleFactura.BeforeRowsDeleted += new Infragistics.Win.UltraWinGrid.BeforeRowsDeletedEventHandler(this.ugDetalleFactura_BeforeRowsDeleted);
			this.ugDetalleFactura.KeyDown += new System.Windows.Forms.KeyEventHandler(this.ugDetalleFactura_KeyDown);
			this.ugDetalleFactura.AfterCellUpdate += new Infragistics.Win.UltraWinGrid.CellEventHandler(this.ugDetalleFactura_AfterCellUpdate);
			// 
			// ultraDataSource4
			// 
			ultraDataColumn13.DataType = typeof(int);
			ultraDataColumn14.DataType = typeof(int);
			ultraDataColumn15.DataType = typeof(int);
			ultraDataColumn17.DataType = typeof(int);
			ultraDataColumn19.DataType = typeof(System.Decimal);
			this.ultraDataSource4.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn13,
																																 ultraDataColumn14,
																																 ultraDataColumn15,
																																 ultraDataColumn16,
																																 ultraDataColumn17,
																																 ultraDataColumn18,
																																 ultraDataColumn19,
																																 ultraDataColumn20});
			// 
			// txtidCliente
			// 
			this.txtidCliente.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtidCliente.Enabled = false;
			this.txtidCliente.Location = new System.Drawing.Point(488, 40);
			this.txtidCliente.Name = "txtidCliente";
			this.txtidCliente.PromptChar = ' ';
			this.txtidCliente.Size = new System.Drawing.Size(72, 22);
			this.txtidCliente.SpinButtonAlignment = Infragistics.Win.ButtonAlignment.Left;
			this.txtidCliente.TabIndex = 847;
			this.txtidCliente.Visible = false;
			// 
			// txtidCompra
			// 
			this.txtidCompra.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtidCompra.Enabled = false;
			this.txtidCompra.Location = new System.Drawing.Point(488, 8);
			this.txtidCompra.Name = "txtidCompra";
			this.txtidCompra.PromptChar = ' ';
			this.txtidCompra.Size = new System.Drawing.Size(72, 22);
			this.txtidCompra.SpinButtonAlignment = Infragistics.Win.ButtonAlignment.Left;
			this.txtidCompra.TabIndex = 848;
			this.txtidCompra.Visible = false;
			// 
			// cmbBodega
			// 
			this.cmbBodega.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbBodega.Cursor = System.Windows.Forms.Cursors.Default;
			ultraGridColumn20.Header.VisiblePosition = 0;
			ultraGridColumn20.Hidden = true;
			ultraGridColumn21.Header.VisiblePosition = 1;
			ultraGridColumn22.Header.VisiblePosition = 2;
			ultraGridBand5.Columns.AddRange(new object[] {
																										 ultraGridColumn20,
																										 ultraGridColumn21,
																										 ultraGridColumn22});
			this.cmbBodega.DisplayLayout.BandsSerializer.Add(ultraGridBand5);
			this.cmbBodega.DisplayMember = "Nombre";
			this.cmbBodega.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbBodega.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbBodega.Enabled = false;
			this.cmbBodega.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbBodega.Location = new System.Drawing.Point(352, 9);
			this.cmbBodega.MaxDropDownItems = 30;
			this.cmbBodega.Name = "cmbBodega";
			this.cmbBodega.Size = new System.Drawing.Size(128, 21);
			this.cmbBodega.TabIndex = 849;
			this.cmbBodega.ValueMember = "Bodega";
			this.cmbBodega.KeyDown += new System.Windows.Forms.KeyEventHandler(this.cmbBodega_KeyDown);
			this.cmbBodega.ValueChanged += new System.EventHandler(this.cmbBodega_ValueChanged_2);
			this.cmbBodega.InitializeLayout += new Infragistics.Win.UltraWinGrid.InitializeLayoutEventHandler(this.cmbBodega_InitializeLayout);
			// 
			// label4
			// 
			this.label4.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.label4.AutoSize = true;
			this.label4.Font = new System.Drawing.Font("Tahoma", 15.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label4.Location = new System.Drawing.Point(424, 344);
			this.label4.Name = "label4";
			this.label4.Size = new System.Drawing.Size(55, 29);
			this.label4.TabIndex = 855;
			this.label4.Text = "Total";
			this.label4.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// txtTotal
			// 
			this.txtTotal.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			appearance25.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtTotal.Appearance = appearance25;
			this.txtTotal.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtTotal.Enabled = false;
			this.txtTotal.Font = new System.Drawing.Font("Tahoma", 15.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.txtTotal.FormatString = "#,##0.00";
			this.txtTotal.Location = new System.Drawing.Point(496, 344);
			this.txtTotal.Name = "txtTotal";
			this.txtTotal.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtTotal.PromptChar = ' ';
			this.txtTotal.Size = new System.Drawing.Size(176, 35);
			this.txtTotal.TabIndex = 854;
			this.txtTotal.ValueChanged += new System.EventHandler(this.txtTotal_ValueChanged);
			// 
			// txtIvaTotal
			// 
			this.txtIvaTotal.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			appearance26.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtIvaTotal.Appearance = appearance26;
			this.txtIvaTotal.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtIvaTotal.Enabled = false;
			this.txtIvaTotal.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.txtIvaTotal.FormatString = "#,##0.00";
			this.txtIvaTotal.Location = new System.Drawing.Point(1120, 392);
			this.txtIvaTotal.Name = "txtIvaTotal";
			this.txtIvaTotal.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtIvaTotal.PromptChar = ' ';
			this.txtIvaTotal.Size = new System.Drawing.Size(8, 22);
			this.txtIvaTotal.TabIndex = 853;
			this.txtIvaTotal.Visible = false;
			// 
			// txtIva
			// 
			this.txtIva.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			appearance27.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtIva.Appearance = appearance27;
			this.txtIva.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtIva.Enabled = false;
			this.txtIva.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.txtIva.FormatString = "#,##0.00";
			this.txtIva.Location = new System.Drawing.Point(1128, 392);
			this.txtIva.Name = "txtIva";
			this.txtIva.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtIva.PromptChar = ' ';
			this.txtIva.Size = new System.Drawing.Size(8, 22);
			this.txtIva.TabIndex = 852;
			this.txtIva.Visible = false;
			// 
			// lblNumero
			// 
			this.lblNumero.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			this.lblNumero.AutoSize = true;
			this.lblNumero.Location = new System.Drawing.Point(928, 11);
			this.lblNumero.Name = "lblNumero";
			this.lblNumero.Size = new System.Drawing.Size(44, 16);
			this.lblNumero.TabIndex = 857;
			this.lblNumero.Text = "Numero";
			this.lblNumero.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// chkEstado
			// 
			appearance28.ForeColorDisabled = System.Drawing.Color.Black;
			this.chkEstado.Appearance = appearance28;
			this.chkEstado.BackColor = System.Drawing.Color.Transparent;
			this.chkEstado.Enabled = false;
			this.chkEstado.Location = new System.Drawing.Point(952, 41);
			this.chkEstado.Name = "chkEstado";
			this.chkEstado.Size = new System.Drawing.Size(56, 21);
			this.chkEstado.TabIndex = 859;
			this.chkEstado.Text = "Activo";
			this.chkEstado.Visible = false;
			// 
			// groupBox2
			// 
			this.groupBox2.Location = new System.Drawing.Point(0, 72);
			this.groupBox2.Name = "groupBox2";
			this.groupBox2.Size = new System.Drawing.Size(1152, 8);
			this.groupBox2.TabIndex = 860;
			this.groupBox2.TabStop = false;
			// 
			// btnVer
			// 
			this.btnVer.Enabled = false;
			this.btnVer.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.btnVer.Image = ((System.Drawing.Image)(resources.GetObject("btnVer.Image")));
			this.btnVer.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnVer.Location = new System.Drawing.Point(1000, 87);
			this.btnVer.Name = "btnVer";
			this.btnVer.Size = new System.Drawing.Size(75, 24);
			this.btnVer.TabIndex = 861;
			this.btnVer.Text = "Ver";
			this.btnVer.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnVer.Click += new System.EventHandler(this.btnVer_Click);
			// 
			// lblEstado
			// 
			this.lblEstado.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			this.lblEstado.AutoSize = true;
			this.lblEstado.Font = new System.Drawing.Font("Microsoft Sans Serif", 15F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.lblEstado.ForeColor = System.Drawing.Color.Firebrick;
			this.lblEstado.Location = new System.Drawing.Point(680, 398);
			this.lblEstado.Name = "lblEstado";
			this.lblEstado.Size = new System.Drawing.Size(0, 26);
			this.lblEstado.TabIndex = 862;
			// 
			// txtNumero
			// 
			appearance29.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtNumero.Appearance = appearance29;
			this.txtNumero.CausesValidation = false;
			this.txtNumero.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtNumero.Enabled = false;
			this.txtNumero.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.txtNumero.Location = new System.Drawing.Point(984, 8);
			this.txtNumero.MaxLength = 9;
			this.txtNumero.Name = "txtNumero";
			this.txtNumero.Size = new System.Drawing.Size(152, 22);
			this.txtNumero.TabIndex = 863;
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
			// txtEstado
			// 
			this.txtEstado.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			this.txtEstado.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtEstado.Location = new System.Drawing.Point(1104, 392);
			this.txtEstado.Name = "txtEstado";
			this.txtEstado.PromptChar = ' ';
			this.txtEstado.Size = new System.Drawing.Size(16, 22);
			this.txtEstado.TabIndex = 864;
			this.txtEstado.Visible = false;
			// 
			// label5
			// 
			this.label5.AutoSize = true;
			this.label5.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label5.Location = new System.Drawing.Point(920, 43);
			this.label5.Name = "label5";
			this.label5.Size = new System.Drawing.Size(53, 16);
			this.label5.TabIndex = 866;
			this.label5.Text = "Categoria";
			this.label5.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// cmbGrupo
			// 
			this.cmbGrupo.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbGrupo.Cursor = System.Windows.Forms.Cursors.Default;
			ultraGridColumn23.Header.VisiblePosition = 0;
			ultraGridColumn23.Hidden = true;
			ultraGridColumn24.Header.VisiblePosition = 1;
			ultraGridColumn24.Width = 180;
			ultraGridBand6.Columns.AddRange(new object[] {
																										 ultraGridColumn23,
																										 ultraGridColumn24});
			this.cmbGrupo.DisplayLayout.BandsSerializer.Add(ultraGridBand6);
			this.cmbGrupo.DisplayMember = "Categoria";
			this.cmbGrupo.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbGrupo.Enabled = false;
			this.cmbGrupo.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbGrupo.Location = new System.Drawing.Point(984, 41);
			this.cmbGrupo.MaxDropDownItems = 30;
			this.cmbGrupo.Name = "cmbGrupo";
			this.cmbGrupo.Size = new System.Drawing.Size(152, 21);
			this.cmbGrupo.TabIndex = 865;
			this.cmbGrupo.ValueMember = "idCategoria";
			this.cmbGrupo.ValueChanged += new System.EventHandler(this.cmbGrupo_ValueChanged);
			// 
			// frmNotaDeEntrega
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 14);
			this.ClientSize = new System.Drawing.Size(1152, 430);
			this.Controls.Add(this.label5);
			this.Controls.Add(this.cmbGrupo);
			this.Controls.Add(this.txtEstado);
			this.Controls.Add(this.txtNumero);
			this.Controls.Add(this.lblEstado);
			this.Controls.Add(this.btnVer);
			this.Controls.Add(this.groupBox2);
			this.Controls.Add(this.chkEstado);
			this.Controls.Add(this.lblNumero);
			this.Controls.Add(this.label4);
			this.Controls.Add(this.txtTotal);
			this.Controls.Add(this.txtIvaTotal);
			this.Controls.Add(this.txtIva);
			this.Controls.Add(this.cmbBodega);
			this.Controls.Add(this.txtidCompra);
			this.Controls.Add(this.txtidCliente);
			this.Controls.Add(this.ugDetalleFactura);
			this.Controls.Add(this.groupBox1);
			this.Controls.Add(this.btnNuevo);
			this.Controls.Add(this.btnConsultar);
			this.Controls.Add(this.btnEditar);
			this.Controls.Add(this.btnImprimir);
			this.Controls.Add(this.btnGuardar);
			this.Controls.Add(this.btnAnular);
			this.Controls.Add(this.btnCancelar);
			this.Controls.Add(this.btnSalir);
			this.Controls.Add(this.ugArticulos);
			this.Controls.Add(this.label2);
			this.Controls.Add(this.txtbuscar);
			this.Controls.Add(this.lblVendedor);
			this.Controls.Add(this.cmbVendedor);
			this.Controls.Add(this.label3);
			this.Controls.Add(this.txtNombre);
			this.Controls.Add(this.lblIdentificacion);
			this.Controls.Add(this.lblCliente);
			this.Controls.Add(this.cmbTipoRuc);
			this.Controls.Add(this.txtRuc);
			this.Controls.Add(this.label1);
			this.Controls.Add(this.dtFecha);
			this.Controls.Add(this.lblLocal);
			this.KeyPreview = true;
			this.Name = "frmNotaDeEntrega";
			this.Text = "Nota de Entrega";
			this.KeyDown += new System.Windows.Forms.KeyEventHandler(this.frmNotaDeEntrega_KeyDown);
			this.Load += new System.EventHandler(this.frmNotaDeEntrega_Load);
			((System.ComponentModel.ISupportInitialize)(this.dtFecha)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbTipoRuc)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtRuc)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNombre)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbVendedor)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtbuscar)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ugArticulos)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource1)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource3)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ugDetalleFactura)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource4)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtidCliente)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtidCompra)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbBodega)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtTotal)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIvaTotal)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIva)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNumero)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteo)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtEstado)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbGrupo)).EndInit();
			this.ResumeLayout(false);

		}
		#endregion

		bool bNuevo = false;
		bool bEdicion = false;
		bool bAnulado = false;
		int idProyct = 0;
		private Acceso miAcceso;
		private bool bActualiza = true;
    int idBodegaPredef = 0;
		int iEstado = 0;

		private void UnloadMe()
		{
			this.Close();
		}

		private void frmNotaDeEntrega_Load(object sender, System.EventArgs e)
		{
			miAcceso = new Acceso(cdsSeteoF, "1122");
			if (!Funcion.Permiso("1491", cdsSeteo))
			{					
				MessageBox.Show("No tiene acceso a Notas de Entrega", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				BeginInvoke(new MethodInvoker(UnloadMe));
				return;
			}

			this.cmbTipoRuc.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "Exec ClienteTipoRucCarga 1");
			
			if (miAcceso.Nuevo) this.btnNuevo.Enabled = true;
			if (miAcceso.Buscar) this.btnConsultar.Enabled = true;
			if (miAcceso.Editar) this.btnEditar.Enabled = true;
			if (miAcceso.Imprimir) this.btnImprimir.Enabled = true;
			if (miAcceso.B15DiasCobro) this.dtFecha.CalendarInfo.MinDate = DateTime.Today.AddDays(-7);	
			
			this.dtFecha.CalendarInfo.MaxDate = DateTime.Today;

			if (this.cmbBodega.ActiveRow != null)
			{
				if (this.cmbVendedor.ActiveRow != null) this.cmbVendedor.Value = System.DBNull.Value;
				this.cmbVendedor.DataSource = Funcion.dvProcedimiento(cdsSeteoF, string.Format("Exec PersonalLotte_Cajero {0}", (int)this.cmbBodega.Value));
			}

			String sSQl = String.Format("Exec BodegaAsignadaPorUsuario {0}",MenuLatinium.IdUsuario);
			idBodegaPredef = Funcion.iEscalarSQL(cdsSeteoF,sSQl);

			if (idCompraV > 0) Consultar(idCompraV);

			string sSQL = string.Format("Select idCategoria, Categoria From ArticuloCategoria Where Activo = 1 Order By Categoria");
			this.cmbGrupo.DataSource = Funcion.dvProcedimiento(cdsSeteo, sSQL);
		}

		private void btBuscar_Click(object sender, System.EventArgs e)
		{
			
		}

		private void cdsSeteoF_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			cdsSeteoF.Schema.Connections[0].ConnectionString = miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);
		}

		private void btnCancelar_Click(object sender, System.EventArgs e)
		{
			bEdicion = false;
			bNuevo = false;
			bAnulado = false;

			#region vaciar campos
			this.dtFecha.Value = System.DBNull.Value;
			this.cmbVendedor.Value = System.DBNull.Value;
			this.cmbBodega.Value = System.DBNull.Value;
			this.cmbTipoRuc.Value = System.DBNull.Value;
			this.cmbGrupo.Value = System.DBNull.Value;
			this.lblEstado.Text ="";
			this.txtNumero.Text="";
			this.txtRuc.Text ="";
			this.txtNombre.Text ="";
			this.txtbuscar.Text = "";
			this.ugArticulos.DataSource = this.ultraDataSource1;

			this.txtIva.Value=0;
			this.txtIvaTotal.Value = 0;
			this.txtTotal.Value = 0;
			this.txtidCliente.Value = 0;
			this.txtidCompra.Value =0;
			this.chkEstado.Checked= false;
			string sSQL = string.Format("Exec ConsultaIndividualDetalleFacturaLotte 0");
			this.ugDetalleFactura.DataSource = Funcion.dvProcedimiento(cdsSeteoF, sSQL);
			#endregion vaciar campos

			#region bloquear campos
			this.dtFecha.Enabled = false;
			this.cmbVendedor.Enabled = false;
			this.cmbBodega.Enabled = false;
			this.cmbTipoRuc.Enabled = false;
			this.cmbGrupo.Enabled = false;
			this.txtRuc.Enabled = false;
			this.txtNombre.Enabled =false;
			this.txtbuscar.Enabled = false;
			this.btnVer.Enabled = false;
			FuncionesProcedimientos.EstadoGrids(false, this.ugDetalleFactura);
			FuncionesProcedimientos.EstadoGrids(false, this.ugArticulos);
			#endregion bloquear campos
			
			if (miAcceso.Nuevo)this.btnNuevo.Enabled = true;
			if (miAcceso.Buscar)this.btnConsultar.Enabled = true;
			this.btnEditar.Enabled = false;
			this.btnImprimir.Enabled = false;
			this.btnGuardar.Enabled = false;			
			this.btnAnular.Enabled = false;
			this.btnCancelar.Enabled = false;
			iEstado = 0;
		}

		private void btnNuevo_Click(object sender, System.EventArgs e)
		{			
			this.btnCancelar_Click(sender, e);

			this.dtFecha.Value = DateTime.Today;

			if (miAcceso.BDia) 
			{
				DateTime dtFecha = DateTime.Today;
				this.dtFecha.CalendarInfo.MinDate = DateTime.Today;
				this.dtFecha.CalendarInfo.MaxDate = dtFecha;
			}
					
			if (miAcceso.CambiarFecha) this.dtFecha.Enabled = true;
	
			#region habilita campos
			
			this.cmbVendedor.Enabled = true;
			this.cmbBodega.Enabled = true;
			//			this.cmbTipoRuc.Enabled = true;
			this.txtRuc.Enabled =true;
			this.txtbuscar.Enabled = true;
			this.btnVer.Enabled = true;
			this.cmbGrupo.Enabled = true;

//			string sSQL = string.Format("Exec FacturaLoteListaInventario '{0}'", this.txtbuscar.Text.ToString());
//			this.ugArticulos.DataSource = Funcion.dvProcedimiento(cdsSeteoF, sSQL);

			FuncionesProcedimientos.EstadoGrids(true, this.ugDetalleFactura);
			FuncionesProcedimientos.EstadoGrids(true, this.ugArticulos);
			#endregion habilita campos
			
			#region controles

			this.btnNuevo.Enabled = false;
			this.btnConsultar.Enabled = false;
			this.btnEditar.Enabled = false;
			this.btnImprimir.Enabled = false;
			this.btnAnular.Enabled = false;
			this.btnCancelar.Enabled = true;	
			this.btnGuardar.Enabled = true;
			#endregion controles

			bNuevo = true;
			bEdicion = true;

			this.cmbTipoRuc.Value = 11;
		}

		private void btnSalir_Click(object sender, System.EventArgs e)
		{
			this.Close();
		}

		private void btnEditar_Click(object sender, System.EventArgs e)
		{
			this.Consultar((int)this.txtidCompra.Value);
			if (!bEdicion) 
			{
				if (this.cmbBodega.ActiveRow != null)
				{
					if (this.cmbVendedor.ActiveRow != null) this.cmbVendedor.Value = System.DBNull.Value;
					this.cmbVendedor.DataSource = Funcion.dvProcedimiento(cdsSeteoF, string.Format("Exec PersonalLotte_Cajero {0}", (int)this.cmbBodega.Value));
				}
			}

			if((bool)this.chkEstado.Checked)
			{
				MessageBox.Show("El documento ya fue anulado ", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				return;
			}		
			
			if (!miAcceso.Editar )
			{
				MessageBox.Show("No tiene acceso a Editar " + Text.Trim(), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);						
				return;
			}

			if((int)this.txtEstado.Value == 10)
			{
				MessageBox.Show("No se puede editar una Nota de Entrega CONTABILIZADA.", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);						
				return;
			}

			#region habilita campos
			if (miAcceso.CambiarFecha) this.dtFecha.Enabled = true;
			if (miAcceso.BEditarNombre)this.cmbVendedor.Enabled = true;
			//			this.cmbBodega.Enabled = true;
			//			this.cmbTipoRuc.Enabled = true;
			//			this.txtRuc.Enabled =true;
			//			this.txtNombre.Enabled =true;
			this.txtbuscar.Enabled = true;
			this.btnVer.Enabled = true;
			FuncionesProcedimientos.EstadoGrids(true, this.ugDetalleFactura);
			FuncionesProcedimientos.EstadoGrids(true, this.ugArticulos);
			#endregion habilita campos

			this.btnNuevo.Enabled = false;
			this.btnConsultar.Enabled = false;
			this.btnEditar.Enabled = false;
			this.btnImprimir.Enabled = false;
			this.btnAnular.Enabled = false;
			this.btnCancelar.Enabled = true;
			this.btnGuardar.Enabled = true;

			bEdicion=true;
		}

		private void dtFecha_ValueChanged(object sender, System.EventArgs e)
		{
			if (this.dtFecha.Value != System.DBNull.Value)
			{
				FuncionesProcedimientos.ListaBodegas(this.cmbBodega, MenuLatinium.IdUsuario, 43, (DateTime)this.dtFecha.Value, false, cdsSeteoF);

				this.dtFecha.CalendarInfo.MinDate = DateTime.Today.AddDays(-7);
			}			
		}

		private void dtFecha_Validating(object sender, System.ComponentModel.CancelEventArgs e)
		{
			if (bEdicion)	
			{				
				if (this.cmbVendedor.ActiveRow != null) 
					this.cmbVendedor.Value = System.DBNull.Value;

				if (this.cmbBodega.ActiveRow != null) 
					this.cmbVendedor.DataSource = Funcion.dvProcedimiento(cdsSeteoF, string.Format("Exec PersonalLotte_Cajero {0}", (int)this.cmbBodega.Value));					
			}
		}	

		private void dtFecha_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if(e.KeyCode == Keys.Enter) this.cmbBodega.Focus();
		}

		private void cmbVendedor_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if(e.KeyCode == Keys.Enter) this.cmbTipoRuc.Focus();
		}

		private void cmbTipoRuc_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if(e.KeyCode == Keys.Enter) this.txtRuc.Focus();
		}

		private void txtRuc_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if(e.KeyCode == Keys.Enter) this.txtbuscar.Focus();
		}

		private void txtbuscar_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if(e.KeyCode == Keys.Enter) this.btnVer.Focus();
		}

		private void txtbuscar_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
			this.btnVer_Click(sender, e);
		}

		private void cmbTipoRuc_ValueChanged(object sender, System.EventArgs e)
		{
			if (bEdicion)
			{
				if (this.cmbTipoRuc.ActiveRow != null) 
				{
					int idCliente = 0;
					this.txtRuc.Text = "";
					this.txtNombre.Text = "";
					//				 this.txtRuc.Enabled = true;

					if ((int)this.cmbTipoRuc.Value == 4 || (int)this.cmbTipoRuc.Value == 6) this.txtRuc.MaxLength = 13;
					if ((int)this.cmbTipoRuc.Value == 5) this.txtRuc.MaxLength = 10;

					#region Consumidor Final
					if ((int)this.cmbTipoRuc.Value == 11) 
					{						
						this.txtRuc.Enabled = false;
						this.txtRuc.MaxLength = 13;
											
						SqlDataReader dr = Funcion.rEscalarSQL(cdsSeteoF, "Select cr.idCliente, cl.Ruc, cl.Nombre From Cliente cl Inner Join ClienteTipoRuc cr On cr.idCliente = cl.idCliente Where cr.idTipoRuc = 11", true);
						dr.Read();
						if (dr.HasRows)
						{
							idCliente = dr.GetInt32(0);
							this.txtRuc.Value = dr.GetString(1);
							this.txtNombre.Value = dr.GetString(2);
						}
						dr.Close();
					}
					#endregion Consumidor Final					
				}
			}
		}

		private void txtRuc_Validating(object sender, System.ComponentModel.CancelEventArgs e)
		{			
			if (!bEdicion) return;
			if (!Validacion.vbIdentificacion(this.txtRuc, (int)this.cmbTipoRuc.Value, cdsSeteoF)) return;
			if (this.txtRuc.Text.ToString().Length == 0) return;
			#region extrae informacion del cliente 
			int idCliente= Funcion.iEscalarSQL(cdsSeteoF,string.Format("if Exists(select idCliente from Cliente where Ruc= '{0}' )select idCliente from Cliente where Ruc= '{0}'   Else Select CONVERT(Int, 0)",this.txtRuc.Text.ToString()));
			if( idCliente > 0)

			{
				this.txtidCliente.Value = idCliente;
				this.txtNombre.Text = Funcion.sEscalarSQL(cdsSeteoF,string.Format("Select Nombre From cliente Where idCliente = {0} ", idCliente));
				this.txtRuc.Text = Funcion.sEscalarSQL(cdsSeteoF,string.Format("Select Ruc From cliente Where idCliente = {0} ", idCliente));
			

			}
			else 
			{
				using (frmClientesProveedores cliente  = new frmClientesProveedores(false,1,idCliente))
				{
					if (DialogResult.OK == cliente.ShowDialog())
					{
						this.txtidCliente.Value = cliente.txtidCliente.Value;
						this.txtNombre.Text = Funcion.sEscalarSQL(cdsSeteoF,string.Format("Select Nombre From cliente Where idCliente = {0} ", (int)this.txtidCliente.Value));
						this.txtRuc.Text = Funcion.sEscalarSQL(cdsSeteoF,string.Format("Select Ruc From cliente Where idCliente = {0} ", (int)this.txtidCliente.Value));
						
					}
				}
			
			}
			#endregion Extrae informacion del cliente 
		}

		private void txtRuc_Validated(object sender, System.EventArgs e)
		{
			
		}

		private void btnGuardar_Click(object sender, System.EventArgs e)
		{
			#region Validacion
			if (!Validacion.vbFechaEnDocumentos(this.dtFecha, "Ingrese la fecha del documento", true, "BODEGA", cdsSeteoF)) return;

			if (!Validacion.vbComboVacio(this.cmbBodega, "Seleccione un local")) return;

			if (!Validacion.vbComboVacio(this.cmbVendedor, "Seleccione un Vendedor")) return;

			if (!Validacion.vbComboVacio(this.cmbTipoRuc, "Seleccione el Tipo de Documento")) return;

			if (!Validacion.vbComboVacio(this.cmbGrupo, "Seleccione una Categoria")) return;

			if (!Validacion.vbTexto(this.txtRuc, 1, 13, "Identificación")) return;
			#endregion Validacion

			#region Valida Grid 
			if (this.ugDetalleFactura.Rows.Count == 0)
			{
				MessageBox.Show("Ingrese el detalle del documento", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Error);
				this.txtbuscar.Focus();
				return;
			}

			int iCantidad =0;
			int iCont =0;
			foreach(Infragistics.Win.UltraWinGrid.UltraGridRow drDTi in this.ugDetalleFactura.Rows.All)
			{
				if((int)drDTi.Cells["Cantidad"].Value == 0)	iCantidad++;
				if(drDTi.Cells["Notas"].Value.ToString().Length < 10)	iCont++;
			}

			if(iCantidad > 0) 
			{
				MessageBox.Show("No Puede Ingresar Articulo en 0 " , "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop); 
				//				ugDetalleFactura.ActiveRow = drDTi;
				ugDetalleFactura.ActiveRow.Selected = true;
				return;
			}

			if(iCont > 0) 
			{
				MessageBox.Show("Debe ingresar nota de minimo 10 caracteres" , "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop); 
				//				ugDetalleFactura.ActiveRow = drDTi;
				ugDetalleFactura.ActiveRow.Selected = true;
				return;
			}

			#region Existencia
			foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.ugDetalleFactura.Rows.All)
			{
				int Existencia = Funcion.iEscalarSQL(cdsSeteoF, string.Format("Exec Existencia_Receta {0}, {1}", 
					(int)dr.Cells["idArticulo"].Value, (int)this.cmbBodega.Value));

				if (Existencia < 1)
				{
					MessageBox.Show("Este Articulo No Tiene Existencia", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop); 
					ugDetalleFactura.ActiveRow = dr;
					ugDetalleFactura .ActiveRow.Selected = true;
					return;
				}

				if ((int)dr.Cells["Cantidad"].Value > Existencia)
				{	
					MessageBox.Show(string.Format("Existencia Actual : {0}", Existencia), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
					ugDetalleFactura.ActiveRow = dr;
					ugDetalleFactura .ActiveRow.Selected = true;
					return;
				}
			}
			#endregion Existencia
			#endregion Valida Grid 

			#region Estado Nota de Entrega
			int iEstadoNE = 0;
			if((int)this.cmbGrupo.Value == 217)
			{
				iEstadoNE = 10;
			}
			else
			{
				iEstadoNE = 1;
			}
			#endregion Estado Nota de Entrega

			if (!Validacion.ValidaTotales((decimal)this.txtTotal.Value)) return;

			if (DialogResult.Yes == MessageBox.Show(string.Format("Esta Seguro de realizar la NOTA DE ENTREGA en el local {0}. \n\n¿Esta seguro de continuar?", (string)this.cmbBodega.Text), "Point Technology", MessageBoxButtons.YesNo, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button2))
			{

				#region Numeracion
				if(bNuevo)
				{
					string sSQL = string.Format("Exec NumeracionLocales 51, {0}, 0", (int)this.cmbBodega.Value);
					this.txtNumero.Text = Funcion.sEscalarSQL(cdsSeteoF, sSQL);		
				}

				#endregion Numeracion

				#region Guardar Maestro
				string  sSQL1 = string.Format("Exec GrabaMaestroFacturaLotte {0}, {1}, {2}, {3}, {4}, {5}, {6}, {7}, {8}, {9}, {10}, {11}, '{12}', '{13}', {14}, {15}, '{16}', '{17}', '{18}', '{19}', '{20}', '{21}', {22}, '{23}', '{24}', {25}, '{26}', {27}, {28}, {29}, {30}, '{31}', {32}, {33}, {34}, {35}, {36}, {37}, {38}, {39}, {40}, {41}, {42}, 1, 0, 0, '', 0.0, {43}, {44}, {45}, {46}, {47}, {48}, {49}, {50}, {51}, {52}, {53},'', 0", 
					(int)this.txtidCompra.Value, /*0*/
					1,/*1*/
					0, /*2*/
					51,/*3*/
					15, /*4*/
					iEstadoNE, /*5*/ //10 Estado
					false, /*6*/
					1, /*7*/
					0, /*8*/
					(int)this.cmbBodega.Value, /*9*/
					idProyct, /*idproyecto*//*10*/
					(int)this.cmbVendedor.Value, 
					this.txtNumero.Text.ToString(), 
					"",
					(int)this.cmbTipoRuc.Value, 
					(int)this.txtidCliente.Value, 
					this.txtRuc.Text.ToString(), 
					this.txtNombre.Text.ToString(), 
					Convert.ToDateTime(this.dtFecha.Value).ToString("yyyyMMdd"), 
					"", 
					"", 
					Convert.ToDateTime(this.dtFecha.Value).ToString("yyyyMMdd"), 
					false, Convert.ToDateTime(this.dtFecha.Value).ToString("yyyyMMdd"), 
					Convert.ToDateTime(this.dtFecha.Value).ToString("yyyyMMdd HH:mm"), 
					0,
					"", /*notas*/
					false,
					0, 
					0, 
					0,/*cuota entrada*/
					Convert.ToDateTime(this.dtFecha.Value).ToString("yyyyMMdd"), 
					0, /* sub total excento*/
					0, /*descuento */
					0, /*sub externo*/
					0, /*descuento*/
					0, /*iva*/
					(decimal)this.txtTotal.Value,	/*total*/
					false, 
					false, 
					false, 
					false, 
					false, 
					0, 
					0, 
					0, 
					0, 
					false, 
					0, 
					0, 
					0,
					0, 
					0, 
					0);
				this.txtidCompra.Value = Funcion.iEscalarSQL(cdsSeteoF,sSQL1);
				#endregion Guardar Maestro

				#region Detalle
				int iPosicion = 0;
				foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.ugDetalleFactura.Rows.All)
				{	
					iPosicion ++;
				
					string sSQLDetalle = string.Format("Exec GuardaDetalleFacturaLotte {0}, {1}, {2}, {3}, '{4}', {5}, {6}, '{7}', '{8}', {9}, {10}, {11}, {12}, {13}, {14}, {15}, {16}, {17}, {18}, {19}, {20}, {21}, {22}, '{23}'",
						(int)dr.Cells["idDetCompra"].Value, 
						0, 
						(int)this.txtidCompra.Value,	
						false, 
						"",
						(int)dr.Cells["idArticulo"].Value, 
						0, 
						dr.Cells["Codigo"].Value.ToString(),
						dr.Cells["Articulo"].Value.ToString(),	
						(int)dr.Cells["Cantidad"].Value, 
						(decimal)dr.Cells["Precio"].Value, 
						0, 
						0, 
						0, 
						0,
						(int)this.cmbBodega.Value, 
						-1,	
						0, 
						false, 
						iPosicion,
						false,
						false, 
						0, 
						dr.Cells["Notas"].Value);
					Funcion.EjecutaSQL(cdsSeteoF, sSQLDetalle);
				}
				#endregion Detalle	

				#region Historial Notas de Entrega
				if(iEstadoNE == 1)
				{
					Funcion.EjecutaSQL(cdsSeteoF, string.Format("Exec GuardaHistorialNotaEntrega {0}, {1}, {2}", (int)this.txtidCompra.Value, 0, (int)this.cmbGrupo.Value));
				}
				#endregion Historial Notas de Entrega

				string sSQL2 = string.Format("Exec ConsultaIndividualDetalleFacturaLotte {0}", (int)this.txtidCompra.Value);
				this.ugDetalleFactura.DataSource = Funcion.dvProcedimiento(cdsSeteoF, sSQL2);

				this.dtFecha.Enabled = false;
				this.cmbVendedor.Enabled = false;
				this.cmbBodega.Enabled = false;
				this.cmbTipoRuc.Enabled = false;
				this.cmbGrupo.Enabled = false;
				this.txtRuc.Enabled =false;
				this.txtNombre.Enabled =false;
				this.txtbuscar.Enabled = false;
				this.btnVer.Enabled = false;
				FuncionesProcedimientos.EstadoGrids(false, this.ugDetalleFactura);
				FuncionesProcedimientos.EstadoGrids(false, this.ugArticulos);
				if (miAcceso.Imprimir) this.btnImprimir.Enabled = true;
				if (miAcceso.Anular)this.btnAnular.Enabled = true;
				if (miAcceso.Nuevo)this.btnNuevo.Enabled = true;
				if (miAcceso.Buscar)this.btnConsultar.Enabled = true;
				if (miAcceso.Editar)this.btnEditar.Enabled = true;
				this.btnGuardar.Enabled = false;
				this.btnCancelar.Enabled = true;
			
				bNuevo = false;
				bEdicion = false;

				this.Consultar((int)this.txtidCompra.Value);
			}
		}

		private void ugArticulos_ClickCellButton(object sender, Infragistics.Win.UltraWinGrid.CellEventArgs e)
		{
			try
			{
				if (e.Cell.Column.ToString() == "UnboundColumn1")
				{
//					foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.ugDetalleFactura.Rows.All)
//					{
//						if ((int)dr.Cells["idArticulo"].Value == (int)e.Cell.Row.Cells["idArticulo"].Value) 
//						{
//							MessageBox.Show("El articulo seleccionado ya esta cargado en la lista de consumos, modifique la Cantidad", "Informacion", MessageBoxButtons.OK, MessageBoxIcon.Stop);
//							ugDetalleFactura.ActiveRow = dr;
//							ugDetalleFactura.ActiveRow.Selected = true;
//							return;
//						}
//					}

					this.CargaArticulos((int)e.Cell.Row.Cells["idArticulo"].Value);
					
				}
				
			}
			catch (Exception ex2) 
			{
				MessageBox.Show(string.Format("Error: {0} {1}", ex2.GetType(), ex2.Message), "Aviso...!!!", MessageBoxButtons.OK, MessageBoxIcon.Error);
			}
		}

		private void Total()
		{		
			#region Creacion de Variables
			
			this.txtTotal.Value = 0.00m;
			#endregion Creacion de Variables

			#region Desplazamiento de grilla
			

			foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.ugDetalleFactura.Rows.All)
			{
				if (dr.Cells["Cantidad"].Value != DBNull.Value && dr.Cells["Precio"].Value != DBNull.Value )
				{
					this.txtTotal.Value = Convert.ToDecimal(this.txtTotal.Value) + Math.Round(Convert.ToDecimal(dr.Cells["Precio"].Value) * (int)dr.Cells["Cantidad"].Value, 6);
				}

				//			this.txtIvaTotal.Value = ((Convert.ToDecimal(this.txtIva.Value) * 12 ) )/100.00m;
			
				this.txtTotal.Value = Math.Round(Convert.ToDecimal(this.txtTotal.Value), 6);
			}
			#endregion Desplazamiento de grilla			
		}

		private void CargaArticulos(int idArticulo)
		{
			#region Carga de Productos
			if (bEdicion)
			{
				#region variables

				int iFila = ugDetalleFactura.Rows.Count - 1;
				iFila++;
				ugDetalleFactura.DisplayLayout.Bands[0].AddNew();	
				#endregion variables

				ugDetalleFactura.Rows[iFila].Cells["IdDetCompra"].Value = 0;  
				ugDetalleFactura.Rows[iFila].Cells["IdCompra"].Value=0;
				ugDetalleFactura.Rows[iFila].Cells["IdArticulo"].Value = (int)this.ugArticulos.ActiveRow.Cells["idArticulo"].Value;
				ugDetalleFactura.Rows[iFila].Cells["Articulo"].Value = this.ugArticulos.ActiveRow.Cells["Articulo"].Value.ToString();
					

				if ( this.ugArticulos.ActiveRow.Cells["Precio1"].Value == System.DBNull.Value)
				{
					ugDetalleFactura.Rows[iFila].Cells["Precio"].Value =  0;
				}
				else 
				{
					ugDetalleFactura.Rows[iFila].Cells["Precio"].Value = this.ugArticulos.ActiveRow.Cells["Precio1"].Value.ToString();
				}

				ugDetalleFactura.Rows[iFila].Cells["Cantidad"].Value = 1;
				ugDetalleFactura.Rows[iFila].Cells["Codigo"].Value = this.ugArticulos.ActiveRow.Cells["Codigo"].Value.ToString();
			}
			#endregion Informacion Productos sin promocion
			
			Total();
		}

		private void cmbBodega_ValueChanged_2(object sender, System.EventArgs e)
		{
			if (bEdicion) 
			{
				if (this.cmbBodega.ActiveRow != null)
				{
				if (this.cmbVendedor.ActiveRow != null) this.cmbVendedor.Value = System.DBNull.Value;
					this.cmbVendedor.DataSource = Funcion.dvProcedimiento(cdsSeteoF, string.Format("Exec PersonalLotte_Cajero {0}", (int)this.cmbBodega.Value));
				}
			}
		}

		private void cmbBodega_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Enter) this.cmbVendedor.Focus();
		}

		private void ugDetalleFactura_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			FuncionesProcedimientos.DesplazamientoGrids(sender, e, this.ugDetalleFactura);	
		}

		private void ugDetalleFactura_AfterCellUpdate(object sender, Infragistics.Win.UltraWinGrid.CellEventArgs e)
		{
			if (!bActualiza) return;

			#region Cantidad

			if (e.Cell.Column.ToString() == "Cantidad")
			{
				bActualiza = false;
				if (e.Cell.Row.Cells["Cantidad"].Value == System.DBNull.Value) e.Cell.Row.Cells["Cantidad"].Value = (int)e.Cell.OriginalValue;
				bActualiza = true;

				Total();
			}
				
			#endregion Cantidad
		}

		private void ugArticulos_InitializeLayout(object sender, Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs e)
		{
		
		}

		private void btnConsultar_Click(object sender, System.EventArgs e)
		{			
			using (frmBuscaFacturas miBusqueda = new frmBuscaFacturas(51, 0, false))
			{
				if (DialogResult.OK == miBusqueda.ShowDialog())
				{
					this.Consultar((int)miBusqueda.grdFacturas.ActiveRow.Cells["idCompra"].Value);
				}
			}
		}

		private void Consultar (int idCompra)
		{			
			this.btnNuevo.Enabled = false;
			this.btnConsultar.Enabled = false;
			this.btnEditar.Enabled = false;
			this.btnImprimir.Enabled =false;
			this.btnGuardar.Enabled = false;
			this.btnCancelar.Enabled = false;

			#region Consulta Individual
			string ssql= string.Format ("Exec ConsultaGeneralFacturaLotte {0}",idCompra); 
			SqlDataReader dr = Funcion.rEscalarSQL(cdsSeteoF,ssql,true);
			dr.Read();
			if (dr.HasRows)
			{
				this.txtidCompra.Value = dr.GetInt32(0);
				if(dr.GetValue(1)!= System.DBNull.Value)this.dtFecha.Value = dr.GetDateTime(1);
				if(dr.GetValue(2)!= System.DBNull.Value)this.cmbBodega.Value = dr.GetInt32(2);
				if(dr.GetValue(3)!= System.DBNull.Value)this.cmbVendedor.Value = dr.GetInt32(3);
				if(dr.GetValue(4)!= System.DBNull.Value)this.txtNumero.Text = dr.GetString(4);
				if(dr.GetValue(5)!= System.DBNull.Value)this.cmbTipoRuc.Value = dr.GetInt32(5);
				if(dr.GetValue(6)!= System.DBNull.Value)this.txtRuc.Text = dr.GetString(6);
				if(dr.GetValue(7)!= System.DBNull.Value)this.txtidCliente.Value = dr.GetInt32(7);
				if(dr.GetValue(8)!= System.DBNull.Value)this.txtNombre.Text = dr.GetString(8);
				if(dr.GetValue(9)!= System.DBNull.Value)this.txtTotal.Value = dr.GetDecimal(9);
				if(dr.GetValue(10)!= System.DBNull.Value)this.chkEstado.Checked = dr.GetBoolean(10);
				if(dr.GetValue(11)!= System.DBNull.Value)this.txtEstado.Value = dr.GetInt32(11);
			}
			dr.Close();
			#endregion Consulta Individual

			string sSQL = string.Format("Exec ConsultaIndividualDetalleFacturaLotte {0}", idCompra);
			this.ugDetalleFactura.DataSource = Funcion.dvProcedimiento(cdsSeteoF, sSQL);
	
			if((bool)this.chkEstado.Checked == false)
			{
				if (miAcceso.Anular)this.btnAnular.Enabled = true;
				if (miAcceso.Editar)this.btnEditar.Enabled = true;
				if (miAcceso.Imprimir) this.btnImprimir.Enabled = true;				
			}
			else
			{
				this.lblEstado.Text = "ANULADO";
				bAnulado = true;
			}

			if(!bAnulado)
			{
				if ((int)this.txtEstado.Value == 1) this.lblEstado.Text = "PENDIENTE";
				if ((int)this.txtEstado.Value == 10) this.lblEstado.Text = "CONTABILIZADO";
			}

			this.cmbVendedor.DataSource = Funcion.dvProcedimiento(cdsSeteoF, string.Format("Exec PersonalLotte_Cajero_Individual {0}", (int)this.cmbVendedor.Value));
			
			if (miAcceso.Nuevo) this.btnNuevo.Enabled = true;
			if (miAcceso.Buscar)this.btnConsultar.Enabled = true;
			this.btnCancelar.Enabled = true;
		}

		private void btnImprimir_Click(object sender, System.EventArgs e)
		{
			Imprimir(0);
		}

		private void Imprimir(int iTipo)
		{
			#region Nombre de Reporte y Filtro
			Cursor = Cursors.WaitCursor;
			string stReporte = "Exec CompraReporteNombre " + (int)this.txtidCompra.Value;
			
			stReporte += ", 1";
			string stFiltro = "";
			stReporte = Funcion.sEscalarSQL(cdsSeteoF, stReporte);
			if (iTipo==0)	stFiltro = "{Compra.idCompra} = " + (int)this.txtidCompra.Value;
			else stFiltro = "{RegistroPromociones.idCompra} = " + (int)this.txtidCompra.Value;
			#endregion Nombre de Reporte y Filtro

			#region Imprime el reporte
			//using (Reporte miReporte = new Reporte(stReporte, stFiltro))
			if (iTipo == 0) 
			{
				#region Impresion
				ParameterFields paramFields = new ParameterFields ();
				
				ParameterField paramFieldId = new ParameterField ();
				ParameterDiscreteValue ValIdId = new ParameterDiscreteValue ();
				paramFieldId.ParameterFieldName = "@idCpa";
				ValIdId.Value = int.Parse(this.txtidCompra.Value.ToString());
				paramFieldId.CurrentValues.Add (ValIdId);
				paramFields.Add (paramFieldId);

				Reporte miRep = new Reporte("FacturaNotaEntrega.rpt", "");
				miRep.MdiParent = this.MdiParent;
				miRep.Titulo("Reserva Habitaciones");
				miRep.crVista.ParameterFieldInfo = paramFields;
				miRep.Show();
				#endregion Impresion
			}
			if (iTipo == 1) 
			{
				using (Reporte miReporte = new Reporte("ImprPromocion", stFiltro))
				{		
					if (iTipo == 1) // Muestra en pantalla
						miReporte.ShowDialog();
					if (iTipo >= 2) // Imprime una o mas copias
					{
						string ImpresoraPredefinida = MenuLatinium.ImpresoraNombre;
						string stImpresora = Funcion.sEscalarSQL(cdsSeteoF, "Exec CompraReporteImpresora '" + stReporte + "'");
						if (stImpresora.Length > 0)	MenuLatinium.ImpresoraNombre = stImpresora;
						miReporte.ImprimeRep(1);
						if (stImpresora.Length > 0)	MenuLatinium.ImpresoraNombre = ImpresoraPredefinida;
					}
					if (iTipo == -1) // Archivo PDF
					{
						miReporte.ImprimeRep(2);
					}
				}
			}
			#endregion Imprime el reporte

			#region Marca como impreso y muestra el tiempo
			string stGrabaImprime = "Update Compra Set Impreso = 1 Where idCompra = " + (int)this.txtidCompra.Value;
			Funcion.EjecutaSQL(cdsSeteoF, stGrabaImprime);
			//			bImpreso = true;
			Cursor = Cursors.Default;
			#endregion Marca como impreso y muestra el tiempo
		}

		private void btnAnular_Click(object sender, System.EventArgs e)
		{
			this.Consultar((int)this.txtidCompra.Value);

			if((bool)this.chkEstado.Checked == true)
			{
				MessageBox.Show("El documento ya fue anulado ", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				return;
			}

			if((int)this.txtEstado.Value == 10)
			{
				MessageBox.Show("No se puede anular una Nota de Entrega CONTABILIZADA.", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);						
				return;
			}
	
			if (!miAcceso.Anular )
			{
				MessageBox.Show("No tiene acceso a Anular " + Text.Trim(), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);						
				return;
			}			
			
			if (DialogResult.Yes == MessageBox.Show("Desea Anular La Nota De Entrega ", "Point Technology", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2))
			{
								
				string sSQLAn = string.Format("Update Compra set Borrar = 1 where idCompra = {0}", (int)this.txtidCompra.Value);
				Funcion.EjecutaSQL(cdsSeteoF, sSQLAn);

				string sSQLAnH = string.Format("Update NotasEntregaLimpieza set Estado = 1 where idCompra = {0}", (int)this.txtidCompra.Value);
				Funcion.EjecutaSQL(cdsSeteoF, sSQLAnH);

				this.Consultar((int)this.txtidCompra.Value);
			}
		}

		private void btnVer_Click(object sender, System.EventArgs e)
		{
			if (!Validacion.vbComboVacio(this.cmbBodega, "Seleccione un local")) return;
			if (this.cmbGrupo.Text.Length == 0)  //this.txtAutFactura.Text.Length
			{
				MessageBox.Show("Seleccione una Categoria.", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);				

				string sSQL = string.Format("Exec ConsultaIndividualDetalleFacturaLotte 0");
				this.ugDetalleFactura.DataSource = Funcion.dvProcedimiento(cdsSeteoF, sSQL);

				this.ugArticulos.DataSource = this.ultraDataSource1;
				this.cmbGrupo.Focus();
				return;
			}

			string sSQLAr = string.Format("Exec FacturaLoteListaInventarioNotaEntrega '{0}', {1}", this.txtbuscar.Text.ToString(), (int)cmbGrupo.Value);
			this.ugArticulos.DataSource = Funcion.dvProcedimiento(cdsSeteoF, sSQLAr);
		}

		private void txtTotal_ValueChanged(object sender, System.EventArgs e)
		{
		
		}

		private void ugDetalleFactura_AfterRowsDeleted(object sender, System.EventArgs e)
		{
			this.Total();
		}

		private void frmNotaDeEntrega_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (!bEdicion)
			{				
				if (e.KeyCode == Keys.N) if (miAcceso.Nuevo)	if (this.btnNuevo.Enabled) this.btnNuevo_Click(sender, e);
				if (e.KeyCode == Keys.B) if (miAcceso.Buscar) if (this.btnConsultar.Enabled) this.btnConsultar_Click(sender, e);
				if ((int)this.txtidCompra.Value > 0)
				{
					if (e.KeyCode == Keys.E) if (miAcceso.Editar) if (this.btnEditar.Enabled) this.btnEditar_Click(sender, e);
					if (e.KeyCode == Keys.I) 
						if (miAcceso.Imprimir) if (this.btnImprimir.Enabled) this.btnImprimir_Click(sender, e);
																		
					if (e.KeyCode == Keys.F5)
					{
						miAcceso = new Acceso(cdsSeteoF, "1122");

						this.Consultar((int)this.txtidCompra.Value );					
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

		private void cmbTipoRuc_InitializeLayout(object sender, Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs e)
		{
		
		}

		private void ugDetalleFactura_BeforeRowsDeleted(object sender, Infragistics.Win.UltraWinGrid.BeforeRowsDeletedEventArgs e)
		{
			if (bEdicion)
			{
				if (DialogResult.Yes ==	MessageBox.Show("¿Esta seguro de Eliminar las filas seleccionadas?", "Point Technology", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2)) 
				{	
					for (int i = 0; i < e.Rows.Length; i++)
					{			
						if ((int)e.Rows[i].Cells["idDetCompra"].Value > 0)
						{	
							Funcion.EjecutaSQL(cdsSeteoF, string.Format("Delete From DetCompra Where idDetCompra = {0}", (int)e.Rows[i].Cells["idDetCompra"].Value));

							e.DisplayPromptMsg = false;
						}
						else e.DisplayPromptMsg = false;
					
						this.btnCancelar.Enabled = false;
					}
				}
				else e.Cancel = true;
			}
			else e.Cancel = true;
		}

		private void cmbBodega_InitializeLayout(object sender, Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs e)
		{
		
		}

		private void cdsSeteo_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			cdsSeteo.Schema.Connections[0].ConnectionString = miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);	
		}

		private void cmbGrupo_ValueChanged(object sender, System.EventArgs e)
		{
			string sSQL = string.Format("Exec ConsultaIndividualDetalleFacturaLotte 0");
			this.ugDetalleFactura.DataSource = Funcion.dvProcedimiento(cdsSeteoF, sSQL);

			this.ugArticulos.DataSource = this.ultraDataSource1;
			this.cmbGrupo.Focus();

			this.txtbuscar.Text = "";			
		}
	}
}

