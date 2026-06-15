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
using System.Text;
using System.Net;
using System.Web;
using System.Xml;
using System.Xml.Schema;
using System.Security.Cryptography.Xml;
using System.Security.Cryptography;
using System.Security.Cryptography.X509Certificates;
using System.Threading;
namespace Latinium
{
	/// <summary>
	/// Descripción breve de frmActaEntregaRecepcion.
	/// </summary>
	public class frmActaEntregaRecepcion : System.Windows.Forms.Form
	{
		public Infragistics.Win.UltraWinEditors.UltraNumericEditor txtidActaEntregaRecepcion;
		private System.Windows.Forms.Label label14;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtNumero;
		private System.Windows.Forms.Label label1;
		private System.Windows.Forms.Label label2;
		private System.Windows.Forms.Label label4;
		private System.Windows.Forms.Label label5;
		private System.Windows.Forms.Label label6;
		private System.Windows.Forms.Label label7;
		private System.Windows.Forms.Label label8;
		private System.Windows.Forms.Label label9;
		private System.Windows.Forms.Label label10;
		private System.Windows.Forms.Label label11;
		private System.Windows.Forms.Label label12;
		private System.Windows.Forms.Label label13;
		private System.Windows.Forms.Label label15;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtFactura;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtVendedor;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtdireccion;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtCliente;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtCel1;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtCel2;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbRealiza;
		private Infragistics.Win.UltraWinSchedule.UltraCalendarCombo dtFecha;
		private Infragistics.Win.UltraWinSchedule.UltraCalendarCombo dtFechaEntrega;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbSector;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtBarrio;
		private System.Windows.Forms.Button btnSalir;
		private System.Windows.Forms.Button btnNuevo;
		private System.Windows.Forms.Button btnConsultar;
		private System.Windows.Forms.Button btnEditar;
		public System.Windows.Forms.Button btnGuardar;
		private System.Windows.Forms.Button btnCancelar;
		private C1.Data.C1DataSet cdsSeteoF;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtBodegaPredef;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtidCompra;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtidEstado;
		private System.Windows.Forms.Label lblEstado;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbCiudad;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource1;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource2;
		public Infragistics.Win.UltraWinGrid.UltraGrid ultgReferidos;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource3;
		public Infragistics.Win.UltraWinEditors.UltraTextEditor txtNota;
		private System.Windows.Forms.Label label16;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbBodegaF;
		private System.Windows.Forms.Label label3;
		private System.Windows.Forms.Label label17;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtRecibe;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtDireccionR;
		private Infragistics.Win.UltraWinGrid.UltraDropDown cmbCentroCosto;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource4;
		/// <summary>
		/// Variable del diseñador requerida.
		/// </summary>
		private System.ComponentModel.Container components = null;

		public frmActaEntregaRecepcion()
		{
			//
			// Necesario para admitir el Diseñador de Windows Forms
			//
			InitializeComponent();

			//
			// TODO: agregar código de constructor después de llamar a InitializeComponent
			//
		}
		bool bVisualizar = false;
		int idVCompra = 0;
		int idVActaEntregaRecepcion = 0;
		public frmActaEntregaRecepcion(bool BVisualizar, int IdCompra, int IdActaEntregaRecepcion)
		{
			//
			// Necesario para admitir el Diseñador de Windows Forms
			//
      bVisualizar = BVisualizar;
      idVCompra = IdCompra;
				idVActaEntregaRecepcion = IdActaEntregaRecepcion;
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
			Infragistics.Win.Appearance appearance2 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance3 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance4 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance5 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance6 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance7 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance8 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand1 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn1 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idPersonal");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn2 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Codigo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn3 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.Appearance appearance9 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton dateButton1 = new Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton();
			System.Resources.ResourceManager resources = new System.Resources.ResourceManager(typeof(frmActaEntregaRecepcion));
			Infragistics.Win.Appearance appearance10 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton dateButton2 = new Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton();
			Infragistics.Win.Appearance appearance11 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand2 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn4 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idSector");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn5 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Sector");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn1 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idSector");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn2 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Sector");
			Infragistics.Win.Appearance appearance12 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance13 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand3 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn6 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCiudad");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn7 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Ciudad");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn3 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idCiudad");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn4 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Ciudad");
			Infragistics.Win.Appearance appearance14 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand4 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn8 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Sel");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn9 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idActaEntregaRecepcion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn10 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idDetalleActaEntregaRecepcion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn11 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idArticulo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn12 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idDetCompra");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn13 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Tipo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn14 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Modelo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn15 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Marca");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn16 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Serial");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn17 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idSerial");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn18 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Bodega", -1, "cmbCentroCosto");
			Infragistics.Win.Appearance appearance15 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance16 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance17 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance18 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn5 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Sel");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn6 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idActaEntregaRecepcion");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn7 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idDetalleActaEntregaRecepcion");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn8 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idArticulo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn9 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idDetCompra");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn10 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Tipo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn11 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Modelo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn12 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Marca");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn13 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Serial");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn14 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idSerial");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn15 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Bodega");
			Infragistics.Win.Appearance appearance19 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance20 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance21 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance22 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand5 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn19 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Bodega");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn20 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn16 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Bodega");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn17 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Nombre");
			this.txtidActaEntregaRecepcion = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.label14 = new System.Windows.Forms.Label();
			this.txtNumero = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.label1 = new System.Windows.Forms.Label();
			this.label2 = new System.Windows.Forms.Label();
			this.label4 = new System.Windows.Forms.Label();
			this.label5 = new System.Windows.Forms.Label();
			this.label6 = new System.Windows.Forms.Label();
			this.label7 = new System.Windows.Forms.Label();
			this.label8 = new System.Windows.Forms.Label();
			this.label9 = new System.Windows.Forms.Label();
			this.label10 = new System.Windows.Forms.Label();
			this.label11 = new System.Windows.Forms.Label();
			this.label12 = new System.Windows.Forms.Label();
			this.label13 = new System.Windows.Forms.Label();
			this.label15 = new System.Windows.Forms.Label();
			this.txtFactura = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.txtVendedor = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.txtdireccion = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.txtCliente = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.txtCel1 = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.txtCel2 = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.cmbRealiza = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.dtFecha = new Infragistics.Win.UltraWinSchedule.UltraCalendarCombo();
			this.dtFechaEntrega = new Infragistics.Win.UltraWinSchedule.UltraCalendarCombo();
			this.cmbSector = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.ultraDataSource2 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.txtBarrio = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.btnSalir = new System.Windows.Forms.Button();
			this.btnNuevo = new System.Windows.Forms.Button();
			this.btnConsultar = new System.Windows.Forms.Button();
			this.btnEditar = new System.Windows.Forms.Button();
			this.btnGuardar = new System.Windows.Forms.Button();
			this.btnCancelar = new System.Windows.Forms.Button();
			this.cdsSeteoF = new C1.Data.C1DataSet();
			this.txtBodegaPredef = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtidCompra = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtidEstado = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.lblEstado = new System.Windows.Forms.Label();
			this.cmbCiudad = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.ultraDataSource1 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.ultgReferidos = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.ultraDataSource3 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.txtNota = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.label16 = new System.Windows.Forms.Label();
			this.cmbBodegaF = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.label3 = new System.Windows.Forms.Label();
			this.txtRecibe = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.txtDireccionR = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.label17 = new System.Windows.Forms.Label();
			this.cmbCentroCosto = new Infragistics.Win.UltraWinGrid.UltraDropDown();
			this.ultraDataSource4 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			((System.ComponentModel.ISupportInitialize)(this.txtidActaEntregaRecepcion)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNumero)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtFactura)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtVendedor)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtdireccion)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtCliente)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtCel1)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtCel2)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbRealiza)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.dtFecha)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.dtFechaEntrega)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbSector)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource2)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtBarrio)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtBodegaPredef)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtidCompra)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtidEstado)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbCiudad)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource1)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultgReferidos)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource3)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNota)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbBodegaF)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtRecibe)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtDireccionR)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbCentroCosto)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource4)).BeginInit();
			this.SuspendLayout();
			// 
			// txtidActaEntregaRecepcion
			// 
			this.txtidActaEntregaRecepcion.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtidActaEntregaRecepcion.Enabled = false;
			this.txtidActaEntregaRecepcion.Location = new System.Drawing.Point(728, 416);
			this.txtidActaEntregaRecepcion.Name = "txtidActaEntregaRecepcion";
			this.txtidActaEntregaRecepcion.PromptChar = ' ';
			this.txtidActaEntregaRecepcion.Size = new System.Drawing.Size(56, 21);
			this.txtidActaEntregaRecepcion.SpinButtonAlignment = Infragistics.Win.ButtonAlignment.Left;
			this.txtidActaEntregaRecepcion.TabIndex = 866;
			this.txtidActaEntregaRecepcion.Visible = false;
			// 
			// label14
			// 
			this.label14.AutoSize = true;
			this.label14.BackColor = System.Drawing.Color.Transparent;
			this.label14.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label14.Location = new System.Drawing.Point(600, 16);
			this.label14.Name = "label14";
			this.label14.Size = new System.Drawing.Size(44, 17);
			this.label14.TabIndex = 985;
			this.label14.Text = "Número";
			this.label14.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// txtNumero
			// 
			appearance1.ForeColor = System.Drawing.Color.Black;
			appearance1.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtNumero.Appearance = appearance1;
			this.txtNumero.Enabled = false;
			this.txtNumero.Font = new System.Drawing.Font("Tahoma", 9F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.txtNumero.Location = new System.Drawing.Point(656, 13);
			this.txtNumero.Name = "txtNumero";
			this.txtNumero.Size = new System.Drawing.Size(144, 23);
			this.txtNumero.TabIndex = 984;
			// 
			// label1
			// 
			this.label1.AutoSize = true;
			this.label1.BackColor = System.Drawing.Color.Transparent;
			this.label1.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label1.Location = new System.Drawing.Point(16, 16);
			this.label1.Name = "label1";
			this.label1.Size = new System.Drawing.Size(150, 17);
			this.label1.TabIndex = 986;
			this.label1.Text = "Almacén que genera la venta";
			this.label1.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// label2
			// 
			this.label2.AutoSize = true;
			this.label2.BackColor = System.Drawing.Color.Transparent;
			this.label2.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label2.Location = new System.Drawing.Point(384, 16);
			this.label2.Name = "label2";
			this.label2.Size = new System.Drawing.Size(65, 17);
			this.label2.TabIndex = 987;
			this.label2.Text = "N°. Factura:";
			this.label2.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// label4
			// 
			this.label4.AutoSize = true;
			this.label4.BackColor = System.Drawing.Color.Transparent;
			this.label4.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label4.Location = new System.Drawing.Point(16, 82);
			this.label4.Name = "label4";
			this.label4.Size = new System.Drawing.Size(110, 17);
			this.label4.TabIndex = 989;
			this.label4.Text = "Nombre Jefe Agencia";
			this.label4.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// label5
			// 
			this.label5.AutoSize = true;
			this.label5.BackColor = System.Drawing.Color.Transparent;
			this.label5.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label5.Location = new System.Drawing.Point(16, 50);
			this.label5.Name = "label5";
			this.label5.Size = new System.Drawing.Size(95, 17);
			this.label5.TabIndex = 990;
			this.label5.Text = "Nombre Vendedor";
			this.label5.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// label6
			// 
			this.label6.AutoSize = true;
			this.label6.BackColor = System.Drawing.Color.Transparent;
			this.label6.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label6.Location = new System.Drawing.Point(360, 50);
			this.label6.Name = "label6";
			this.label6.Size = new System.Drawing.Size(78, 17);
			this.label6.TabIndex = 991;
			this.label6.Text = "Fecha Factura:";
			this.label6.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// label7
			// 
			this.label7.AutoSize = true;
			this.label7.BackColor = System.Drawing.Color.Transparent;
			this.label7.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label7.Location = new System.Drawing.Point(576, 50);
			this.label7.Name = "label7";
			this.label7.Size = new System.Drawing.Size(80, 17);
			this.label7.TabIndex = 992;
			this.label7.Text = "Fecha Entrega:";
			this.label7.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// label8
			// 
			this.label8.AutoSize = true;
			this.label8.BackColor = System.Drawing.Color.Transparent;
			this.label8.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label8.Location = new System.Drawing.Point(376, 82);
			this.label8.Name = "label8";
			this.label8.Size = new System.Drawing.Size(42, 17);
			this.label8.TabIndex = 993;
			this.label8.Text = "Ciudad:";
			this.label8.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// label9
			// 
			this.label9.AutoSize = true;
			this.label9.BackColor = System.Drawing.Color.Transparent;
			this.label9.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label9.Location = new System.Drawing.Point(600, 82);
			this.label9.Name = "label9";
			this.label9.Size = new System.Drawing.Size(40, 17);
			this.label9.TabIndex = 994;
			this.label9.Text = "Sector:";
			this.label9.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// label10
			// 
			this.label10.AutoSize = true;
			this.label10.BackColor = System.Drawing.Color.Transparent;
			this.label10.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label10.Location = new System.Drawing.Point(384, 112);
			this.label10.Name = "label10";
			this.label10.Size = new System.Drawing.Size(38, 17);
			this.label10.TabIndex = 995;
			this.label10.Text = "Barrio:";
			this.label10.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// label11
			// 
			this.label11.AutoSize = true;
			this.label11.BackColor = System.Drawing.Color.Transparent;
			this.label11.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label11.Location = new System.Drawing.Point(384, 144);
			this.label11.Name = "label11";
			this.label11.Size = new System.Drawing.Size(110, 17);
			this.label11.TabIndex = 996;
			this.label11.Text = "Dirección del Cliente:";
			this.label11.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// label12
			// 
			this.label12.AutoSize = true;
			this.label12.BackColor = System.Drawing.Color.Transparent;
			this.label12.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label12.Location = new System.Drawing.Point(16, 144);
			this.label12.Name = "label12";
			this.label12.Size = new System.Drawing.Size(104, 17);
			this.label12.TabIndex = 997;
			this.label12.Text = "Nombre del Cliente:";
			this.label12.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// label13
			// 
			this.label13.AutoSize = true;
			this.label13.BackColor = System.Drawing.Color.Transparent;
			this.label13.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label13.Location = new System.Drawing.Point(432, 178);
			this.label13.Name = "label13";
			this.label13.Size = new System.Drawing.Size(62, 17);
			this.label13.TabIndex = 998;
			this.label13.Text = "Contacto 1:";
			this.label13.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.label13.Click += new System.EventHandler(this.label13_Click);
			// 
			// label15
			// 
			this.label15.AutoSize = true;
			this.label15.BackColor = System.Drawing.Color.Transparent;
			this.label15.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label15.Location = new System.Drawing.Point(616, 178);
			this.label15.Name = "label15";
			this.label15.Size = new System.Drawing.Size(66, 17);
			this.label15.TabIndex = 999;
			this.label15.Text = "Contactot 2:";
			this.label15.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// txtFactura
			// 
			appearance2.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtFactura.Appearance = appearance2;
			this.txtFactura.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtFactura.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtFactura.Enabled = false;
			this.txtFactura.Location = new System.Drawing.Point(464, 14);
			this.txtFactura.Name = "txtFactura";
			this.txtFactura.Size = new System.Drawing.Size(128, 21);
			this.txtFactura.TabIndex = 1001;
			// 
			// txtVendedor
			// 
			appearance3.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtVendedor.Appearance = appearance3;
			this.txtVendedor.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtVendedor.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtVendedor.Enabled = false;
			this.txtVendedor.Location = new System.Drawing.Point(120, 48);
			this.txtVendedor.Name = "txtVendedor";
			this.txtVendedor.Size = new System.Drawing.Size(232, 21);
			this.txtVendedor.TabIndex = 1002;
			// 
			// txtdireccion
			// 
			appearance4.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtdireccion.Appearance = appearance4;
			this.txtdireccion.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtdireccion.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtdireccion.Enabled = false;
			this.txtdireccion.Location = new System.Drawing.Point(496, 144);
			this.txtdireccion.Name = "txtdireccion";
			this.txtdireccion.Size = new System.Drawing.Size(296, 21);
			this.txtdireccion.TabIndex = 1003;
			// 
			// txtCliente
			// 
			appearance5.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtCliente.Appearance = appearance5;
			this.txtCliente.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtCliente.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtCliente.Enabled = false;
			this.txtCliente.Location = new System.Drawing.Point(120, 144);
			this.txtCliente.Name = "txtCliente";
			this.txtCliente.Size = new System.Drawing.Size(256, 21);
			this.txtCliente.TabIndex = 1004;
			// 
			// txtCel1
			// 
			appearance6.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtCel1.Appearance = appearance6;
			this.txtCel1.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtCel1.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtCel1.Enabled = false;
			this.txtCel1.Location = new System.Drawing.Point(496, 176);
			this.txtCel1.Name = "txtCel1";
			this.txtCel1.Size = new System.Drawing.Size(112, 21);
			this.txtCel1.TabIndex = 1005;
			// 
			// txtCel2
			// 
			appearance7.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtCel2.Appearance = appearance7;
			this.txtCel2.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtCel2.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtCel2.Enabled = false;
			this.txtCel2.Location = new System.Drawing.Point(688, 176);
			this.txtCel2.MaxLength = 10;
			this.txtCel2.Name = "txtCel2";
			this.txtCel2.Size = new System.Drawing.Size(112, 21);
			this.txtCel2.TabIndex = 1006;
			// 
			// cmbRealiza
			// 
			appearance8.ForeColor = System.Drawing.Color.Black;
			appearance8.ForeColorDisabled = System.Drawing.Color.Black;
			this.cmbRealiza.Appearance = appearance8;
			this.cmbRealiza.CausesValidation = false;
			this.cmbRealiza.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbRealiza.Cursor = System.Windows.Forms.Cursors.Default;
			ultraGridColumn1.Header.VisiblePosition = 0;
			ultraGridColumn1.Hidden = true;
			ultraGridColumn2.Header.VisiblePosition = 1;
			ultraGridColumn3.Header.VisiblePosition = 2;
			ultraGridColumn3.Width = 300;
			ultraGridBand1.Columns.AddRange(new object[] {
																										 ultraGridColumn1,
																										 ultraGridColumn2,
																										 ultraGridColumn3});
			this.cmbRealiza.DisplayLayout.BandsSerializer.Add(ultraGridBand1);
			this.cmbRealiza.DisplayMember = "Nombre";
			this.cmbRealiza.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbRealiza.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbRealiza.Enabled = false;
			this.cmbRealiza.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbRealiza.Location = new System.Drawing.Point(128, 80);
			this.cmbRealiza.MaxDropDownItems = 30;
			this.cmbRealiza.Name = "cmbRealiza";
			this.cmbRealiza.Size = new System.Drawing.Size(224, 21);
			this.cmbRealiza.TabIndex = 1008;
			this.cmbRealiza.ValueMember = "idPersonal";
			this.cmbRealiza.KeyDown += new System.Windows.Forms.KeyEventHandler(this.cmbRealiza_KeyDown);
			this.cmbRealiza.InitializeLayout += new Infragistics.Win.UltraWinGrid.InitializeLayoutEventHandler(this.cmbRealiza_InitializeLayout);
			// 
			// dtFecha
			// 
			this.dtFecha.AllowMonthSelection = true;
			this.dtFecha.AllowWeekSelection = true;
			appearance9.ForeColor = System.Drawing.Color.Black;
			appearance9.ForeColorDisabled = System.Drawing.Color.Black;
			this.dtFecha.Appearance = appearance9;
			this.dtFecha.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			dateButton1.Caption = "Today";
			this.dtFecha.DateButtons.Add(dateButton1);
			this.dtFecha.Enabled = false;
			this.dtFecha.Format = "dd/MM/yyyy";
			this.dtFecha.Location = new System.Drawing.Point(440, 48);
			this.dtFecha.Name = "dtFecha";
			this.dtFecha.NonAutoSizeHeight = 23;
			this.dtFecha.Size = new System.Drawing.Size(128, 21);
			this.dtFecha.SpinButtonsVisible = true;
			this.dtFecha.TabIndex = 1009;
			this.dtFecha.Value = ((object)(resources.GetObject("dtFecha.Value")));
			this.dtFecha.ValueChanged += new System.EventHandler(this.dtFecha_ValueChanged);
			// 
			// dtFechaEntrega
			// 
			this.dtFechaEntrega.AllowMonthSelection = true;
			this.dtFechaEntrega.AllowWeekSelection = true;
			appearance10.ForeColor = System.Drawing.Color.Black;
			appearance10.ForeColorDisabled = System.Drawing.Color.Black;
			this.dtFechaEntrega.Appearance = appearance10;
			this.dtFechaEntrega.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			dateButton2.Caption = "Today";
			this.dtFechaEntrega.DateButtons.Add(dateButton2);
			this.dtFechaEntrega.Enabled = false;
			this.dtFechaEntrega.Format = "dd/MM/yyyy";
			this.dtFechaEntrega.Location = new System.Drawing.Point(656, 48);
			this.dtFechaEntrega.Name = "dtFechaEntrega";
			this.dtFechaEntrega.NonAutoSizeHeight = 23;
			this.dtFechaEntrega.Size = new System.Drawing.Size(128, 21);
			this.dtFechaEntrega.SpinButtonsVisible = true;
			this.dtFechaEntrega.TabIndex = 1010;
			this.dtFechaEntrega.Value = ((object)(resources.GetObject("dtFechaEntrega.Value")));
			// 
			// cmbSector
			// 
			appearance11.ForeColor = System.Drawing.Color.Black;
			appearance11.ForeColorDisabled = System.Drawing.Color.Black;
			this.cmbSector.Appearance = appearance11;
			this.cmbSector.CausesValidation = false;
			this.cmbSector.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbSector.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbSector.DataSource = this.ultraDataSource2;
			ultraGridColumn4.Header.VisiblePosition = 0;
			ultraGridColumn4.Hidden = true;
			ultraGridColumn5.Header.VisiblePosition = 1;
			ultraGridColumn5.Width = 150;
			ultraGridBand2.Columns.AddRange(new object[] {
																										 ultraGridColumn4,
																										 ultraGridColumn5});
			this.cmbSector.DisplayLayout.BandsSerializer.Add(ultraGridBand2);
			this.cmbSector.DisplayMember = "Sector";
			this.cmbSector.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbSector.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbSector.Enabled = false;
			this.cmbSector.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbSector.Location = new System.Drawing.Point(656, 80);
			this.cmbSector.MaxDropDownItems = 30;
			this.cmbSector.Name = "cmbSector";
			this.cmbSector.Size = new System.Drawing.Size(128, 21);
			this.cmbSector.TabIndex = 1012;
			this.cmbSector.ValueMember = "idSector";
			this.cmbSector.KeyDown += new System.Windows.Forms.KeyEventHandler(this.cmbSector_KeyDown);
			// 
			// ultraDataSource2
			// 
			ultraDataColumn1.DataType = typeof(int);
			this.ultraDataSource2.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn1,
																																 ultraDataColumn2});
			// 
			// txtBarrio
			// 
			appearance12.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtBarrio.Appearance = appearance12;
			this.txtBarrio.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtBarrio.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtBarrio.Enabled = false;
			this.txtBarrio.Location = new System.Drawing.Point(440, 112);
			this.txtBarrio.Name = "txtBarrio";
			this.txtBarrio.Size = new System.Drawing.Size(280, 21);
			this.txtBarrio.TabIndex = 1013;
			this.txtBarrio.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtBarrio_KeyDown);
			this.txtBarrio.ValueChanged += new System.EventHandler(this.ultraTextEditor1_ValueChanged);
			// 
			// btnSalir
			// 
			this.btnSalir.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btnSalir.CausesValidation = false;
			this.btnSalir.Image = ((System.Drawing.Image)(resources.GetObject("btnSalir.Image")));
			this.btnSalir.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnSalir.Location = new System.Drawing.Point(408, 440);
			this.btnSalir.Name = "btnSalir";
			this.btnSalir.TabIndex = 1020;
			this.btnSalir.Text = "&Salir";
			this.btnSalir.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnSalir.Click += new System.EventHandler(this.btnSalir_Click);
			// 
			// btnNuevo
			// 
			this.btnNuevo.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btnNuevo.CausesValidation = false;
			this.btnNuevo.Enabled = false;
			this.btnNuevo.Image = ((System.Drawing.Image)(resources.GetObject("btnNuevo.Image")));
			this.btnNuevo.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnNuevo.Location = new System.Drawing.Point(8, 440);
			this.btnNuevo.Name = "btnNuevo";
			this.btnNuevo.Size = new System.Drawing.Size(72, 24);
			this.btnNuevo.TabIndex = 1018;
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
			this.btnConsultar.Location = new System.Drawing.Point(88, 440);
			this.btnConsultar.Name = "btnConsultar";
			this.btnConsultar.Size = new System.Drawing.Size(78, 24);
			this.btnConsultar.TabIndex = 1017;
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
			this.btnEditar.Location = new System.Drawing.Point(168, 440);
			this.btnEditar.Name = "btnEditar";
			this.btnEditar.Size = new System.Drawing.Size(76, 24);
			this.btnEditar.TabIndex = 1016;
			this.btnEditar.Text = "&Editar";
			this.btnEditar.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnEditar.Click += new System.EventHandler(this.btnEditar_Click);
			// 
			// btnGuardar
			// 
			this.btnGuardar.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btnGuardar.Enabled = false;
			this.btnGuardar.Image = ((System.Drawing.Image)(resources.GetObject("btnGuardar.Image")));
			this.btnGuardar.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnGuardar.Location = new System.Drawing.Point(248, 440);
			this.btnGuardar.Name = "btnGuardar";
			this.btnGuardar.Size = new System.Drawing.Size(76, 24);
			this.btnGuardar.TabIndex = 1015;
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
			this.btnCancelar.Location = new System.Drawing.Point(328, 440);
			this.btnCancelar.Name = "btnCancelar";
			this.btnCancelar.Size = new System.Drawing.Size(72, 24);
			this.btnCancelar.TabIndex = 1014;
			this.btnCancelar.Text = "&Cancelar";
			this.btnCancelar.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnCancelar.Click += new System.EventHandler(this.btnCancelar_Click);
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
			// txtBodegaPredef
			// 
			this.txtBodegaPredef.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			this.txtBodegaPredef.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtBodegaPredef.Enabled = false;
			this.txtBodegaPredef.Location = new System.Drawing.Point(648, 448);
			this.txtBodegaPredef.Name = "txtBodegaPredef";
			this.txtBodegaPredef.PromptChar = ' ';
			this.txtBodegaPredef.Size = new System.Drawing.Size(64, 21);
			this.txtBodegaPredef.TabIndex = 1021;
			this.txtBodegaPredef.Visible = false;
			// 
			// txtidCompra
			// 
			this.txtidCompra.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			this.txtidCompra.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtidCompra.Enabled = false;
			this.txtidCompra.Location = new System.Drawing.Point(648, 416);
			this.txtidCompra.Name = "txtidCompra";
			this.txtidCompra.PromptChar = ' ';
			this.txtidCompra.Size = new System.Drawing.Size(64, 21);
			this.txtidCompra.TabIndex = 1022;
			this.txtidCompra.Visible = false;
			// 
			// txtidEstado
			// 
			this.txtidEstado.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			this.txtidEstado.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtidEstado.Enabled = false;
			this.txtidEstado.Location = new System.Drawing.Point(728, 416);
			this.txtidEstado.Name = "txtidEstado";
			this.txtidEstado.PromptChar = ' ';
			this.txtidEstado.Size = new System.Drawing.Size(64, 21);
			this.txtidEstado.TabIndex = 1023;
			this.txtidEstado.Visible = false;
			// 
			// lblEstado
			// 
			this.lblEstado.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			this.lblEstado.AutoSize = true;
			this.lblEstado.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.lblEstado.ForeColor = System.Drawing.Color.Firebrick;
			this.lblEstado.Location = new System.Drawing.Point(512, 440);
			this.lblEstado.Name = "lblEstado";
			this.lblEstado.Size = new System.Drawing.Size(0, 22);
			this.lblEstado.TabIndex = 1024;
			// 
			// cmbCiudad
			// 
			appearance13.ForeColor = System.Drawing.Color.Black;
			appearance13.ForeColorDisabled = System.Drawing.Color.Black;
			this.cmbCiudad.Appearance = appearance13;
			this.cmbCiudad.CausesValidation = false;
			this.cmbCiudad.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbCiudad.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbCiudad.DataSource = this.ultraDataSource1;
			ultraGridColumn6.Header.VisiblePosition = 0;
			ultraGridColumn6.Hidden = true;
			ultraGridColumn7.Header.VisiblePosition = 1;
			ultraGridColumn7.Width = 150;
			ultraGridBand3.Columns.AddRange(new object[] {
																										 ultraGridColumn6,
																										 ultraGridColumn7});
			this.cmbCiudad.DisplayLayout.BandsSerializer.Add(ultraGridBand3);
			this.cmbCiudad.DisplayMember = "Ciudad";
			this.cmbCiudad.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbCiudad.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbCiudad.Enabled = false;
			this.cmbCiudad.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbCiudad.Location = new System.Drawing.Point(440, 80);
			this.cmbCiudad.MaxDropDownItems = 30;
			this.cmbCiudad.Name = "cmbCiudad";
			this.cmbCiudad.Size = new System.Drawing.Size(128, 21);
			this.cmbCiudad.TabIndex = 1026;
			this.cmbCiudad.ValueMember = "idCiudad";
			this.cmbCiudad.KeyDown += new System.Windows.Forms.KeyEventHandler(this.cmbCiudad_KeyDown_1);
			// 
			// ultraDataSource1
			// 
			ultraDataColumn3.DataType = typeof(int);
			this.ultraDataSource1.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn3,
																																 ultraDataColumn4});
			// 
			// ultgReferidos
			// 
			this.ultgReferidos.Cursor = System.Windows.Forms.Cursors.Default;
			this.ultgReferidos.DataSource = this.ultraDataSource3;
			appearance14.BackColor = System.Drawing.Color.White;
			this.ultgReferidos.DisplayLayout.Appearance = appearance14;
			ultraGridColumn8.Header.Caption = "...";
			ultraGridColumn8.Header.VisiblePosition = 0;
			ultraGridColumn9.Header.VisiblePosition = 1;
			ultraGridColumn9.Hidden = true;
			ultraGridColumn10.Header.VisiblePosition = 9;
			ultraGridColumn10.Hidden = true;
			ultraGridColumn11.Header.VisiblePosition = 2;
			ultraGridColumn11.Hidden = true;
			ultraGridColumn12.Header.VisiblePosition = 3;
			ultraGridColumn12.Hidden = true;
			ultraGridColumn13.Case = Infragistics.Win.UltraWinGrid.Case.Upper;
			ultraGridColumn13.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn13.Header.VisiblePosition = 4;
			ultraGridColumn13.Width = 150;
			ultraGridColumn14.Case = Infragistics.Win.UltraWinGrid.Case.Upper;
			ultraGridColumn14.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn14.Header.VisiblePosition = 5;
			ultraGridColumn14.Width = 140;
			ultraGridColumn15.Case = Infragistics.Win.UltraWinGrid.Case.Upper;
			ultraGridColumn15.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn15.Header.VisiblePosition = 6;
			ultraGridColumn15.Width = 140;
			ultraGridColumn16.Case = Infragistics.Win.UltraWinGrid.Case.Upper;
			ultraGridColumn16.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn16.Header.VisiblePosition = 7;
			ultraGridColumn16.Width = 140;
			ultraGridColumn17.Case = Infragistics.Win.UltraWinGrid.Case.Upper;
			ultraGridColumn17.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn17.Header.VisiblePosition = 8;
			ultraGridColumn17.Hidden = true;
			ultraGridColumn17.Width = 175;
			ultraGridColumn18.ButtonDisplayStyle = Infragistics.Win.UltraWinGrid.ButtonDisplayStyle.Always;
			ultraGridColumn18.Header.VisiblePosition = 10;
			ultraGridColumn18.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.DropDown;
			ultraGridColumn18.Width = 140;
			ultraGridBand4.Columns.AddRange(new object[] {
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
			ultraGridBand4.Header.Caption = "Productos";
			ultraGridBand4.HeaderVisible = true;
			ultraGridBand4.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.False;
			this.ultgReferidos.DisplayLayout.BandsSerializer.Add(ultraGridBand4);
			this.ultgReferidos.DisplayLayout.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.No;
			this.ultgReferidos.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.False;
			this.ultgReferidos.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.True;
			appearance15.BackColor = System.Drawing.Color.Transparent;
			this.ultgReferidos.DisplayLayout.Override.CardAreaAppearance = appearance15;
			appearance16.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance16.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance16.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance16.FontData.BoldAsString = "True";
			appearance16.FontData.Name = "Arial";
			appearance16.FontData.SizeInPoints = 8.5F;
			appearance16.ForeColor = System.Drawing.Color.White;
			appearance16.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.ultgReferidos.DisplayLayout.Override.HeaderAppearance = appearance16;
			appearance17.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance17.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance17.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ultgReferidos.DisplayLayout.Override.RowSelectorAppearance = appearance17;
			appearance18.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance18.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance18.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ultgReferidos.DisplayLayout.Override.SelectedRowAppearance = appearance18;
			this.ultgReferidos.Enabled = false;
			this.ultgReferidos.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.ultgReferidos.Location = new System.Drawing.Point(16, 240);
			this.ultgReferidos.Name = "ultgReferidos";
			this.ultgReferidos.RightToLeft = System.Windows.Forms.RightToLeft.No;
			this.ultgReferidos.Size = new System.Drawing.Size(776, 192);
			this.ultgReferidos.TabIndex = 1027;
			this.ultgReferidos.CellChange += new Infragistics.Win.UltraWinGrid.CellEventHandler(this.ultgReferidos_CellChange);
			this.ultgReferidos.KeyDown += new System.Windows.Forms.KeyEventHandler(this.ultgReferidos_KeyDown);
			// 
			// ultraDataSource3
			// 
			ultraDataColumn5.DataType = typeof(bool);
			ultraDataColumn6.DataType = typeof(int);
			ultraDataColumn7.DataType = typeof(int);
			ultraDataColumn8.DataType = typeof(int);
			ultraDataColumn9.DataType = typeof(int);
			ultraDataColumn14.DataType = typeof(int);
			ultraDataColumn15.DataType = typeof(int);
			this.ultraDataSource3.Band.Columns.AddRange(new object[] {
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
																																 ultraDataColumn15});
			// 
			// txtNota
			// 
			appearance19.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtNota.Appearance = appearance19;
			this.txtNota.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtNota.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtNota.Enabled = false;
			this.txtNota.Location = new System.Drawing.Point(120, 208);
			this.txtNota.Multiline = true;
			this.txtNota.Name = "txtNota";
			this.txtNota.Size = new System.Drawing.Size(664, 24);
			this.txtNota.TabIndex = 1033;
			// 
			// label16
			// 
			this.label16.AutoSize = true;
			this.label16.BackColor = System.Drawing.Color.Transparent;
			this.label16.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label16.Location = new System.Drawing.Point(16, 208);
			this.label16.Name = "label16";
			this.label16.Size = new System.Drawing.Size(80, 17);
			this.label16.TabIndex = 1034;
			this.label16.Text = "Observaciónes:";
			this.label16.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.label16.Click += new System.EventHandler(this.label16_Click);
			// 
			// cmbBodegaF
			// 
			appearance20.ForeColor = System.Drawing.Color.Black;
			appearance20.ForeColorDisabled = System.Drawing.Color.Black;
			this.cmbBodegaF.Appearance = appearance20;
			this.cmbBodegaF.CausesValidation = false;
			this.cmbBodegaF.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbBodegaF.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbBodegaF.DisplayMember = "Nombre";
			this.cmbBodegaF.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbBodegaF.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbBodegaF.Enabled = false;
			this.cmbBodegaF.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbBodegaF.Location = new System.Drawing.Point(168, 16);
			this.cmbBodegaF.MaxDropDownItems = 30;
			this.cmbBodegaF.Name = "cmbBodegaF";
			this.cmbBodegaF.Size = new System.Drawing.Size(208, 21);
			this.cmbBodegaF.TabIndex = 1037;
			this.cmbBodegaF.ValueMember = "Bodega";
			this.cmbBodegaF.ValueChanged += new System.EventHandler(this.cmbBodegaF_ValueChanged);
			// 
			// label3
			// 
			this.label3.AutoSize = true;
			this.label3.BackColor = System.Drawing.Color.Transparent;
			this.label3.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label3.Location = new System.Drawing.Point(16, 112);
			this.label3.Name = "label3";
			this.label3.Size = new System.Drawing.Size(85, 17);
			this.label3.TabIndex = 1038;
			this.label3.Text = "Persona Recibe:";
			this.label3.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// txtRecibe
			// 
			appearance21.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtRecibe.Appearance = appearance21;
			this.txtRecibe.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtRecibe.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtRecibe.Enabled = false;
			this.txtRecibe.Location = new System.Drawing.Point(120, 112);
			this.txtRecibe.Name = "txtRecibe";
			this.txtRecibe.Size = new System.Drawing.Size(256, 21);
			this.txtRecibe.TabIndex = 1039;
			// 
			// txtDireccionR
			// 
			appearance22.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtDireccionR.Appearance = appearance22;
			this.txtDireccionR.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtDireccionR.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtDireccionR.Enabled = false;
			this.txtDireccionR.Location = new System.Drawing.Point(120, 176);
			this.txtDireccionR.Name = "txtDireccionR";
			this.txtDireccionR.Size = new System.Drawing.Size(296, 21);
			this.txtDireccionR.TabIndex = 1040;
			// 
			// label17
			// 
			this.label17.AutoSize = true;
			this.label17.BackColor = System.Drawing.Color.Transparent;
			this.label17.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label17.Location = new System.Drawing.Point(16, 178);
			this.label17.Name = "label17";
			this.label17.Size = new System.Drawing.Size(96, 17);
			this.label17.TabIndex = 1041;
			this.label17.Text = "Direccion Entrega:";
			this.label17.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// cmbCentroCosto
			// 
			this.cmbCentroCosto.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbCentroCosto.DataSource = this.ultraDataSource4;
			ultraGridColumn19.Header.VisiblePosition = 0;
			ultraGridColumn20.Header.VisiblePosition = 1;
			ultraGridColumn20.Width = 150;
			ultraGridBand5.Columns.AddRange(new object[] {
																										 ultraGridColumn19,
																										 ultraGridColumn20});
			this.cmbCentroCosto.DisplayLayout.BandsSerializer.Add(ultraGridBand5);
			this.cmbCentroCosto.DisplayMember = "Nombre";
			this.cmbCentroCosto.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbCentroCosto.Location = new System.Drawing.Point(624, 328);
			this.cmbCentroCosto.Name = "cmbCentroCosto";
			this.cmbCentroCosto.Size = new System.Drawing.Size(160, 56);
			this.cmbCentroCosto.TabIndex = 1042;
			this.cmbCentroCosto.ValueMember = "Bodega";
			this.cmbCentroCosto.Visible = false;
			// 
			// ultraDataSource4
			// 
			ultraDataColumn16.DataType = typeof(int);
			this.ultraDataSource4.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn16,
																																 ultraDataColumn17});
			// 
			// frmActaEntregaRecepcion
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 13);
			this.ClientSize = new System.Drawing.Size(808, 478);
			this.Controls.Add(this.cmbCentroCosto);
			this.Controls.Add(this.label17);
			this.Controls.Add(this.label3);
			this.Controls.Add(this.label16);
			this.Controls.Add(this.label15);
			this.Controls.Add(this.label13);
			this.Controls.Add(this.label12);
			this.Controls.Add(this.label11);
			this.Controls.Add(this.label10);
			this.Controls.Add(this.label9);
			this.Controls.Add(this.label8);
			this.Controls.Add(this.label7);
			this.Controls.Add(this.label6);
			this.Controls.Add(this.label5);
			this.Controls.Add(this.label4);
			this.Controls.Add(this.label2);
			this.Controls.Add(this.label1);
			this.Controls.Add(this.label14);
			this.Controls.Add(this.txtDireccionR);
			this.Controls.Add(this.txtRecibe);
			this.Controls.Add(this.cmbBodegaF);
			this.Controls.Add(this.txtNota);
			this.Controls.Add(this.ultgReferidos);
			this.Controls.Add(this.cmbCiudad);
			this.Controls.Add(this.lblEstado);
			this.Controls.Add(this.txtidEstado);
			this.Controls.Add(this.txtidCompra);
			this.Controls.Add(this.txtBodegaPredef);
			this.Controls.Add(this.btnSalir);
			this.Controls.Add(this.btnNuevo);
			this.Controls.Add(this.btnConsultar);
			this.Controls.Add(this.btnEditar);
			this.Controls.Add(this.btnGuardar);
			this.Controls.Add(this.btnCancelar);
			this.Controls.Add(this.txtBarrio);
			this.Controls.Add(this.cmbSector);
			this.Controls.Add(this.dtFechaEntrega);
			this.Controls.Add(this.dtFecha);
			this.Controls.Add(this.cmbRealiza);
			this.Controls.Add(this.txtCel2);
			this.Controls.Add(this.txtCel1);
			this.Controls.Add(this.txtCliente);
			this.Controls.Add(this.txtdireccion);
			this.Controls.Add(this.txtVendedor);
			this.Controls.Add(this.txtFactura);
			this.Controls.Add(this.txtNumero);
			this.Controls.Add(this.txtidActaEntregaRecepcion);
			this.KeyPreview = true;
			this.Name = "frmActaEntregaRecepcion";
			this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
			this.Text = "ACTA DE ENTREGA-RECEPCIÓN";
			this.KeyDown += new System.Windows.Forms.KeyEventHandler(this.frmActaEntregaRecepcion_KeyDown);
			this.Load += new System.EventHandler(this.frmActaEntregaRecepcion_Load);
			((System.ComponentModel.ISupportInitialize)(this.txtidActaEntregaRecepcion)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNumero)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtFactura)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtVendedor)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtdireccion)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtCliente)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtCel1)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtCel2)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbRealiza)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.dtFecha)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.dtFechaEntrega)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbSector)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource2)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtBarrio)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtBodegaPredef)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtidCompra)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtidEstado)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbCiudad)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource1)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultgReferidos)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource3)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNota)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbBodegaF)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtRecibe)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtDireccionR)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbCentroCosto)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource4)).EndInit();
			this.ResumeLayout(false);

		}
		#endregion

		bool bNuevo = false;
		bool bEdicion = false;
		bool bAnulado = false;

		private Acceso miAcceso;
		private void UnloadMe()
		{
			this.Close();
		}
		bool bEliminaAlValidar = false;
		private void frmActaEntregaRecepcion_Load(object sender, System.EventArgs e)
		{
			miAcceso = new Acceso(cdsSeteoF, "200110");

			if (!Funcion.Permiso("1883", cdsSeteoF))
			{				
				MessageBox.Show("No puede ingresar a Acta Entrega Recepción", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				BeginInvoke(new MethodInvoker(UnloadMe));
				return;
			}

			if (miAcceso.Nuevo) this.btnNuevo.Enabled = true;
			if (miAcceso.Buscar) this.btnConsultar.Enabled = true;
			string sSQL = string.Format("Exec BodegaAsignadaPorUsuario {0}", MenuLatinium.IdUsuario);		
			this.txtBodegaPredef.Value = Funcion.iEscalarSQL(cdsSeteoF, sSQL);

			this.cmbCiudad.DataSource = Funcion.dvProcedimiento(cdsSeteoF,"select idCiudad,Ciudad from dbo.ClienteCiudad where idCiudad in (1,2) order by Ciudad");
			this.cmbSector.DataSource = Funcion.dvProcedimiento(cdsSeteoF,"select idSector, Sector from dbo.ClienteSector");
			this.cmbBodegaF.DataSource = Funcion.dvProcedimiento(cdsSeteoF,"select Bodega, Nombre from dbo.Bodega ");
			this.cmbCentroCosto.DataSource = Funcion.dvProcedimiento(cdsSeteoF,"select Bodega, Nombre from Bodega where Factura = 1 and Inventario = 1 and Activo = 1 and Automatico = 1 and Bodega <> 69 order by Nombre ");

			if (miAcceso.Nuevo) this.btnNuevo.Enabled = true;
			if (miAcceso.Buscar) this.btnConsultar.Enabled = true;
			if( bVisualizar )
			{
				this.ConsultaIndividual(idVCompra, idVActaEntregaRecepcion);
				this.btnEditar.Visible = false;
				this.btnGuardar.Visible = false;
				this.btnNuevo.Visible = false;
				this.btnConsultar.Visible = false;
				this.btnCancelar.Visible = false;
			}
		}

		private void cmbRealiza_InitializeLayout(object sender, Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs e)
		{
		
		}

		private void ultraTextEditor1_ValueChanged(object sender, System.EventArgs e)
		{
		
		}

		private void btnCancelar_Click(object sender, System.EventArgs e)
		{
			bNuevo = false;
			bEdicion = false;

			#region vaciar
			//this.txtAlmacenVenta.Text ="";
			this.txtFactura.Text ="";
			this.txtNumero.Text ="";
			//this.cmbBodegaOrigen.Value = System.DBNull.Value;
			this.txtRecibe.Text = "";
			this.txtDireccionR.Text = "";
			this.cmbRealiza.Value = System.DBNull.Value;
			this.txtVendedor.Text ="";
			this.dtFecha.Value = System.DBNull.Value;
			this.dtFechaEntrega.Value = System.DBNull.Value;
			this.cmbCiudad.Value = System.DBNull.Value;
			this.cmbSector.Value = System.DBNull.Value;
			this.txtBarrio.Text ="";
			this.txtdireccion.Text ="";
			this.txtCliente.Text ="";
			this.txtCel1.Text ="";
			this.txtCel2.Text ="";
			this.txtidActaEntregaRecepcion.Value = 0;
			this.txtidCompra.Value = 0;
			this.txtBodegaPredef.Value = 0;
			this.txtidEstado.Value = 0;
			this.lblEstado.Text ="";
			this.txtNota.Text ="";
			this.cmbBodegaF.Value= System.DBNull.Value;

      #endregion vaciar

			#region bloqueo
			this.cmbRealiza.Enabled = false;
      //this.cmbBodegaOrigen.Enabled = false;
			this.cmbCiudad.Enabled = false;
			this.cmbSector.Enabled = false;
			this.txtBarrio.Enabled = false;
			this.txtCel2.Enabled = false;
			this.txtNota.Enabled = false;
			this.txtRecibe.Enabled = false;
			this.txtDireccionR.Enabled = false;
			#endregion bloqueo

			this.ultgReferidos.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "Exec ConsultaNuevoDetalleActaEntregaRecepcion  0");
			FuncionesProcedimientos.EstadoGrids(false,this.ultgReferidos);


			if (miAcceso.Nuevo) this.btnNuevo.Enabled = true;
			if (miAcceso.Buscar) this.btnConsultar.Enabled = true;
			this.btnEditar.Enabled = false;
			this.btnGuardar.Enabled = false;
			this.btnCancelar.Enabled = false;

		}

		private void btnNuevo_Click(object sender, System.EventArgs e)
		{
			this.btnCancelar_Click(sender, e);

			using (frmBuscaFacturas Busqueda = new frmBuscaFacturas(1, (int)this.txtBodegaPredef.Value, false, false))
			{
				if (DialogResult.OK == Busqueda.ShowDialog())
				{
					int iNumero = Funcion.iEscalarSQL(cdsSeteoF, string.Format(" Select Count(1)  from dbo.ActaEntregaRecepcion where iEstado in (1,2) and idCompra =  {0}",
						(int)Busqueda.grdFacturas.ActiveRow.Cells["idCompra"].Value));	
					if( iNumero > 0)
					{
						string sNumero = Funcion.sEscalarSQL(cdsSeteoF, string.Format(" Select numero  from dbo.ActaEntregaRecepcion where  idCompra =  {0}",
							(int)Busqueda.grdFacturas.ActiveRow.Cells["idCompra"].Value));	

					 MessageBox.Show(string.Format("La Factura Seleccionada ya fue ingresada en el documento N° '{0}'.",  sNumero), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);	
					 return;
					}

					if ((int)Busqueda.grdFacturas.ActiveRow.Cells["Estado"].Value == 10 )
					{
           this.Consultar((int)Busqueda.grdFacturas.ActiveRow.Cells["idCompra"].Value);
						this.btnNuevo.Enabled = false;
						this.btnConsultar.Enabled = false;
						this.btnEditar.Enabled = false;
						this.btnGuardar.Enabled = true;
						this.btnCancelar.Enabled = true;	
			
						bNuevo = true;
						bEdicion = true;
					}
					
				}
			}
			//FuncionesProcedimientos.RetornaFechaServidor(this.dtFecha, cdsSeteoF, true);

			

		}


		private void ConsultaIndividual(int IdCompra, int IdActaEntregaRecepcion)
		{
	  int Bodega = 0;
			try
			{	
				string sSQL = string.Format("Exec ConsultaIndividualActadeEntrega {0}", IdCompra);
				SqlDataReader dr = Funcion.rEscalarSQL(cdsSeteoF, sSQL, true);
				dr.Read();
				if (dr.HasRows)
				{		
					//this.txtAlmacenVenta.Text =  dr.GetString(0);
					Bodega = dr.GetInt32(0);
				
					this.txtFactura.Text =  dr.GetString(1);
					this.txtVendedor.Text =  dr.GetString(2);
					this.dtFecha.Value = dr.GetDateTime(3);
					this.txtdireccion.Text =  dr.GetString(4);
					this.txtCliente.Text =  dr.GetString(5);
					this.txtCel1.Text =  dr.GetString(6);
					this.txtidCompra.Value = dr.GetInt32(7);
		
					
				}
				dr.Close();
        	this.cmbBodegaF.Value = Bodega;
				string sSQLa = string.Format("Exec ConsultaIndidividualActaEntregaRecepcion {0}", IdActaEntregaRecepcion);
				SqlDataReader dre = Funcion.rEscalarSQL(cdsSeteoF, sSQLa, true);
				dre.Read();
				if (dre.HasRows)
				{	
					this.txtidActaEntregaRecepcion.Value = IdActaEntregaRecepcion;
					if (dre.GetValue(1) != System.DBNull.Value)  this.txtNumero.Text = dre.GetString(1); 
					//if (dre.GetValue(2) != System.DBNull.Value)  this.cmbBodegaOrigen.Value = dre.GetInt32(2);
					if (dre.GetValue(3) != System.DBNull.Value)  this.cmbRealiza.Value = dre.GetInt32(3);
					if (dre.GetValue(4) != System.DBNull.Value)  this.dtFechaEntrega.Value = dre.GetDateTime(4);
					if (dre.GetValue(5) != System.DBNull.Value)  this.cmbCiudad.Value = dre.GetInt32(5);
					if (dre.GetValue(6) != System.DBNull.Value)  this.cmbSector.Value = dre.GetInt32(6);
					if (dre.GetValue(7) != System.DBNull.Value)  this.txtBarrio.Text = dre.GetString(7); 
					if (dre.GetValue(8) != System.DBNull.Value)  this.txtCel2.Text = dre.GetString(8); 
					if (dre.GetValue(9) != System.DBNull.Value)  this.txtidEstado.Value = dre.GetInt32(9); 
					if (dre.GetValue(10) != System.DBNull.Value)  this.lblEstado.Text = dre.GetString(10);
					if (dre.GetValue(11) != System.DBNull.Value)  this.txtNota.Text = dre.GetString(11);
					if (dre.GetValue(12) != System.DBNull.Value)  this.txtRecibe.Text = dre.GetString(12);
					if (dre.GetValue(13) != System.DBNull.Value)  this.txtDireccionR.Text = dre.GetString(13);
				}
				dre.Close();

				string sSQLPa = string.Format(" EXEC ConsultaNuevoDetalleActaEntregaRecepcionGuardar   {0}, {1} ", IdActaEntregaRecepcion, IdCompra);
				FuncionesProcedimientos.dsGeneral(sSQLPa, this.ultgReferidos);

				this.btnGuardar.Enabled = false;
				this.btnCancelar.Enabled = true;

				if (miAcceso.Nuevo) this.btnNuevo.Enabled = true;
				if (miAcceso.Buscar) this.btnConsultar.Enabled = true;
				if (miAcceso.Editar)
					if ((int)this.txtidEstado.Value == 1 ) 
						this.btnEditar.Enabled = true;


			}
			catch(Exception Exc)
			{
				MessageBox.Show(string.Format("Error al Consultar: {0}", Exc), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Error);
			}
		}

		private void Consultar(int IdCompra)
		{
	 int Bodega = 0;
			try
			{	
				string sSQL = string.Format("Exec ConsultaIndividualActadeEntrega {0}", IdCompra);
				SqlDataReader dr = Funcion.rEscalarSQL(cdsSeteoF, sSQL, true);
				dr.Read();
				if (dr.HasRows)
				{		
					//this.txtAlmacenVenta.Text =  dr.GetString(0);
					Bodega = dr.GetInt32(0);
					this.txtFactura.Text =  dr.GetString(1);
					this.txtVendedor.Text =  dr.GetString(2);
					this.dtFecha.Value = dr.GetDateTime(3);
					this.txtdireccion.Text =  dr.GetString(4);
					this.txtCliente.Text =  dr.GetString(5);
					this.txtCel1.Text =  dr.GetString(6);
					this.txtidCompra.Value = dr.GetInt32(7);
					this.txtDireccionR.Text =  dr.GetString(4);
					this.txtRecibe.Text =  dr.GetString(5);
          //Bodegas = dr.GetInt32(8);
				}
				dr.Close();
        this.cmbBodegaF.Value = Bodega;
				this.txtidEstado.Value = 1;
				this.lblEstado.Text ="PENDIENTE";

				string sSQLPa = string.Format(" EXEC ConsultaNuevoDetalleActaEntregaRecepcion   {0} ", (int)this.txtidCompra.Value);
				FuncionesProcedimientos.dsGeneral(sSQLPa, this.ultgReferidos);
        
				FuncionesProcedimientos.EstadoGrids(true, this.ultgReferidos);

				#region controles

				#region bloqueo
				this.cmbRealiza.Enabled = true;
				//this.cmbBodegaOrigen.Enabled = true;
				this.cmbCiudad.Enabled = true;
				this.cmbSector.Enabled = true;
				this.txtBarrio.Enabled = true;
				this.txtCel2.Enabled = true;
				this.txtNota.Enabled = true;
				this.txtDireccionR.Enabled = true;
				this.txtRecibe.Enabled = true;
				#endregion bloqueo


				this.btnNuevo.Enabled = false;
				this.btnConsultar.Enabled = false;
				this.btnEditar.Enabled = false;
				this.btnGuardar.Enabled = true;
				this.btnCancelar.Enabled = true;	
				#endregion controles
			
			}
			catch(Exception Exc)
			{
				MessageBox.Show(string.Format("Error al Consultar: {0}", Exc), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Error);
			}
		}

		private void cdsSeteoF_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			cdsSeteoF.Schema.Connections[0].ConnectionString = miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);
		
		}

		private void btnConsultar_Click(object sender, System.EventArgs e)
		{
			btnCancelar_Click(sender, e);

			using (frmBuscaActaEntregaRecepcion Busqueda = new frmBuscaActaEntregaRecepcion( (int)this.txtBodegaPredef.Value))
			{
				if (DialogResult.OK == Busqueda.ShowDialog()) 
				{
					this.ConsultaIndividual((int)Busqueda.grdFacturas.ActiveRow.Cells["idCompra"].Value, (int)Busqueda.grdFacturas.ActiveRow.Cells["idActaEntregaRecepcion"].Value );
				}
			}	

			
		}

		private void dtFecha_ValueChanged(object sender, System.EventArgs e)
		{
//			if (this.dtFecha.Value != System.DBNull.Value)
//			{
//				if (bEdicion ) return;
//				
//				FuncionesProcedimientos.ListaBodegas(this.cmbBodegaOrigen, MenuLatinium.IdUsuario, 1, (DateTime)this.dtFecha.Value, false, cdsSeteoF);
//			
//			}
		}

		private void cmbBodegaOrigen_ValueChanged(object sender, System.EventArgs e)
		{
//			if (this.cmbBodegaOrigen.ActiveRow != null)
//			{				
//
//				if (bNuevo) this.txtNumero.Text = "";
//
//				if (this.cmbRealiza.ActiveRow != null) this.cmbRealiza.Value = System.DBNull.Value;
//
//				if (this.cmbBodegaOrigen.ActiveRow != null && this.dtFecha.Value != System.DBNull.Value)
//				{	
//					//this.cmbRealiza.DataSource = FuncionesProcedimientos.Lista_Vendedores(Convert.ToDateTime(this.dtFecha.Value), (int)this.cmbBodegaOrigen.Value, cdsSeteoF);
//					this.cmbRealiza.DataSource = Funcion.dvProcedimiento(cdsSeteoF, string.Format("EXEC FacturacionListaVendedoresIngreso '{0}', {1} ",Convert.ToDateTime(this.dtFecha.Value).ToString("yyyyMMdd"), (int)this.cmbBodegaOrigen.Value ));
//				}
//			}
		}

		private void cmbBodegaOrigen_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if(e.KeyCode == Keys.Enter) this.cmbRealiza.Focus();
		}

		private void cmbRealiza_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
				if(e.KeyCode == Keys.Enter) this.cmbCiudad.Focus();
		}

		private void cmbCiudad_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if(e.KeyCode == Keys.Enter) this.cmbSector.Focus();
		}

		private void cmbSector_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if(e.KeyCode == Keys.Enter) this.txtBarrio.Focus();
		}

		private void txtBarrio_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if(e.KeyCode == Keys.Enter) this.txtCel2.Focus();
		}

		private void cmbCiudad_KeyDown_1(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if(e.KeyCode == Keys.Enter) this.cmbSector.Focus();
		}

		private void btnGuardar_Click(object sender, System.EventArgs e)
		{
			#region Validacion
			DateTime dtfechacierre = DateTime.Parse("01/01/2000");
			if (this.dtFechaEntrega.Value != System.DBNull.Value)dtfechacierre=(DateTime)this.dtFechaEntrega.Value;
			//if (!Validacion.vbComboVacio(this.cmbBodegaOrigen, "Seleccione la Bodega Origén")) return;
			if (!Validacion.vbComboVacio(this.cmbRealiza, "Selección el Jefe de Agencia")) return;
			if (!Validacion.vbComboVacio(this.cmbCiudad, "Selección la Ciudad")) return;
			if (!Validacion.vbComboVacio(this.cmbSector, "Selección el Sector")) return;
			if (!Validacion.vbTexto(this.txtBarrio, 10, 250, "Barrio")) return;
			if (!Validacion.vbTexto(this.txtRecibe, 10, 299, "Persona Recibe")) return;
			if (!Validacion.vbTexto(this.txtDireccionR, 10, 299, "Dirección Entrega")) return;
			if (this.txtCel1.Text.ToString().Trim().Length == 0   )
			{
				MessageBox.Show(string.Format("Debe Ingresar el Telefono Convencional o el Telefono Celular"), "Point Technology", MessageBoxButtons.OK,MessageBoxIcon.Stop);
				this.txtCel1.Focus();
				return;
				
			}
			if (this.txtCel2.Text.ToString().Trim().Length > 0)
			{
				if (this.txtCel2.Text.ToString().Trim().Length <= 9)
				{
					if (!Validacion.vbValidaTelefono(this.txtCel2, 1, false)) return;
				}
				else
				{
					if (!Validacion.vbValidaTelefono(this.txtCel2, 2, false)) return;
				}
			}

			if (this.txtNota.Text.ToString().Trim().Length > 0)
			{
				if (!Validacion.vbTexto(this.txtNota, 10, 200, "Observación")) return;
			}
			if((int) this.ultgReferidos.Rows.Count == 0)
			{
				MessageBox.Show(string.Format("Seleccione un Articulo."), "Point Technology", MessageBoxButtons.OK,MessageBoxIcon.Stop);
				return;
			}
			int icont = 0;
			foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.ultgReferidos.Rows.All)
			{
				if ((bool)dr.Cells["Sel"].Value == true)
				{
					icont = icont +1;
				}	
			}

			if ( icont == 0 )
			{
				MessageBox.Show(string.Format("Seleccione al menos un Articulo."), "Point Technology", MessageBoxButtons.OK,MessageBoxIcon.Stop);
				return;
			}
			#endregion Validacion
      int intoi = 0;
			foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.ultgReferidos.Rows.All)
			{
				if ((bool)dr.Cells["Sel"].Value == true && (int)dr.Cells["Bodega"].Value == 0 )
				{
					intoi = intoi +1;
				}	
			}
			if ( intoi > 0 )
			{
				MessageBox.Show(string.Format("Seleccione La bodega, en los articulos seleccionados."), "Point Technology", MessageBoxButtons.OK,MessageBoxIcon.Stop);
				return;
			}
			#region Guardar

			//			DateTime dtInicia = DateTime.Today;
			//			DateTime dtRetorna = DateTime.Today;
			//				dtInicia = (DateTime)this.dtInicia.Value;
			//				dtRetorna = (DateTime)this.dtRetorna.Value;

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
					if (bNuevo)
					{
						string ssSQL = string.Format("Exec NumeracionLocales 84, 0, 0", 0);
						oCmdActualiza.CommandText = ssSQL;
						this.txtNumero.Text = oCmdActualiza.ExecuteScalar().ToString();						
					}
					#endregion Numeracion

					#region Graba Maestro
			
//					if (this.cmbCuenta.ActiveRow != null) idCuenta = (int) this.cmbCuenta.Value;
//					if(this.cmbFormaPago.ActiveRow != null) idFormaDePago = (int) this.cmbFormaPago.Value;

					string  sSQL1 = string.Format("Exec GrabaActaEntregaRecepcion {0}, {1}, '{2}', '{3}', {4}, '{5}', {6}, {7}, '{8}', '{9}', {10}, '{11}', '{12}'" ,
						(int)this.txtidActaEntregaRecepcion.Value,
						(int)this.txtidCompra.Value, 
						this.txtNumero.Text,
						this.txtRecibe.Text,
						//(int)this.cmbBodegaOrigen.Value,
						(int)this.cmbRealiza.Value,
						Convert.ToDateTime(dtfechacierre).ToString("yyyyMMdd"), 
						(int)this.cmbCiudad.Value,
						(int)this.cmbSector.Value,
						this.txtBarrio.Text,
						this.txtCel2.Text,
						(int) this.txtidEstado.Value, 
						this.txtNota.Text,
						this.txtDireccionR.Text
						);
					oCmdActualiza.CommandText = sSQL1;
					this.txtidActaEntregaRecepcion.Value = (int)oCmdActualiza.ExecuteScalar();
					#endregion Graba Maestro

					#region Graba Ingreso Gastos
										
					foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.ultgReferidos.Rows.All)
					{
						if ((bool)dr.Cells["Sel"].Value == true)
						{
							string  sSQLDeta = string.Format("Exec GrabaDetalleActaEntregaRecepcion {0}, {1}, {2}, {3}, {4}" ,
								Convert.ToInt32 ( dr.Cells["idDetalleActaEntregaRecepcion"].Value), 
								(int) this.txtidActaEntregaRecepcion.Value, 
								Convert.ToInt32 ( dr.Cells["idDetCompra"].Value), 
								Convert.ToInt32 ( dr.Cells["idSerial"].Value),
								Convert.ToInt32 ( dr.Cells["Bodega"].Value)
								);
							oCmdActualiza.CommandText = sSQLDeta;
							oCmdActualiza.ExecuteNonQuery();	
						}	
					}
				
					#endregion Graba Ingreso Gastos
          
					oTransaction.Commit();

					Cursor = Cursors.Default;	

					MessageBox.Show("Registro Guardado Correctamente", "Información", MessageBoxButtons.OK, MessageBoxIcon.Information);

					#region Controles
					this.cmbRealiza.Enabled = false;
				//	this.cmbBodegaOrigen.Enabled = false;
					this.cmbCiudad.Enabled = false;
					this.cmbSector.Enabled = false;
					this.txtBarrio.Enabled = false;
					this.txtCel2.Enabled = false;
					this.txtNota.Enabled = false;
					this.txtRecibe.Enabled = false;
					this.txtDireccionR.Enabled = false;
				
					FuncionesProcedimientos.EstadoGrids(false,this.ultgReferidos);

				

				//	FuncionesProcedimientos.EstadoGrids(false,this.uGridDetalle);

					if (miAcceso.Nuevo)this.btnNuevo.Enabled = true;
					if (miAcceso.Buscar)this.btnConsultar.Enabled = true;
					if (miAcceso.Editar)this.btnEditar.Enabled = false;

					this.btnGuardar.Enabled = false;
					this.btnCancelar.Enabled = true;

					this.ConsultaIndividual((int)this.txtidCompra.Value,(int) this.txtidActaEntregaRecepcion.Value );
					//this.ConsultaIndividual((int)this.txtidCompra.Value,(int) this.txtidActaEntregaRecepcion.Value );
          
					bNuevo = false;
					bEdicion = false;
					#endregion Controles			
				}
				catch(Exception ex)
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
				
			#endregion Guardar
		}

		private void frmActaEntregaRecepcion_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (!bEdicion)
			{
				if (e.KeyCode == Keys.N) if (this.btnNuevo.Enabled) this.btnNuevo_Click(sender, e);
				if (e.KeyCode == Keys.B) if (this.btnConsultar.Enabled) this.btnConsultar_Click(sender, e);
				if (e.KeyCode == Keys.E) if (this.btnEditar.Enabled) this.btnEditar_Click(sender, e); 																																																										 
				if (e.KeyCode == Keys.F5)	this.ConsultaIndividual((int) this.txtidCompra.Value ,(int)this.txtidActaEntregaRecepcion.Value);
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

		private void btnEditar_Click(object sender, System.EventArgs e)
		{
		
			this.ConsultaIndividual((int)this.txtidCompra.Value,(int) this.txtidActaEntregaRecepcion.Value );
				if((int) this.txtidEstado.Value > 1) return;
			this.cmbRealiza.Enabled = true;
		//	this.cmbBodegaOrigen.Enabled = true;
			this.cmbCiudad.Enabled = true;
			this.cmbSector.Enabled = true;
			this.txtBarrio.Enabled = true;
			this.txtCel2.Enabled = true;
			this.txtNota.Enabled = true;
			this.txtRecibe.Enabled = true;
			this.txtDireccionR.Enabled = true;

			FuncionesProcedimientos.EstadoGrids(true, this.ultgReferidos);
			
			this.btnNuevo.Enabled = false;
			this.btnConsultar.Enabled = false;
			this.btnEditar.Enabled = false;			
			this.btnGuardar.Enabled = true;			
			this.btnCancelar.Enabled = true;
	


			bNuevo = false;
			bEdicion = true;

		}

		private void ultgReferidos_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			FuncionesProcedimientos.DesplazamientoGrids(sender,e,this.ultgReferidos);
		}

		private void ultgReferidos_CellChange(object sender, Infragistics.Win.UltraWinGrid.CellEventArgs e)
		{

				this.ultgReferidos.UpdateData();

				bool bDisponible = false;		
				if ((bool)e.Cell.Row.Cells["Sel"].Value) bDisponible = true;
				else bDisponible = false;
				if (!bDisponible)
				{
				
					#region Ingresa valores
					if((int) e.Cell.Row.Cells["idDetalleActaEntregaRecepcion"].Value > 0 )
					{
						string sSQL = string.Format("Delete From DetalleActaEntregaRecepcion where idDetalleActaEntregaRecepcion = {0} ", (int) e.Cell.Row.Cells["idDetalleActaEntregaRecepcion"].Value);
						Funcion.EjecutaSQL(cdsSeteoF, sSQL, true);

						e.Cell.Row.Cells["idActaEntregaRecepcion"].Value = 0;
						e.Cell.Row.Cells["idDetalleActaEntregaRecepcion"].Value = 0;
					}
					#endregion Ingresa valores
				}
			
		}

		private void btnSalir_Click(object sender, System.EventArgs e)
		{
				this.Close();		
		}

		private void label16_Click(object sender, System.EventArgs e)
		{
		
		}

		private void label13_Click(object sender, System.EventArgs e)
		{
		
		}


		private void cmbBodegaF_ValueChanged(object sender, System.EventArgs e)
		{
						if (this.cmbBodegaF.ActiveRow != null)
						{				
			
							if (bNuevo) this.txtNumero.Text = "";
			
							if (this.cmbRealiza.ActiveRow != null) this.cmbRealiza.Value = System.DBNull.Value;
			
							if (this.cmbBodegaF.ActiveRow != null && this.dtFecha.Value != System.DBNull.Value)
							{	
								//this.cmbRealiza.DataSource = FuncionesProcedimientos.Lista_Vendedores(Convert.ToDateTime(this.dtFecha.Value), (int)this.cmbBodegaOrigen.Value, cdsSeteoF);
								this.cmbRealiza.DataSource = Funcion.dvProcedimiento(cdsSeteoF, string.Format("EXEC FacturacionListaVendedoresIngreso '{0}', {1} ",Convert.ToDateTime(this.dtFecha.Value).ToString("yyyyMMdd"), (int)this.cmbBodegaF.Value ));
							}
						}
		}
	}
}
