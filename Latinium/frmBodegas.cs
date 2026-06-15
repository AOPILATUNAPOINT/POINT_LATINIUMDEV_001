using System.Data;
using System.Data.SqlClient;
using C1.Data;
using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;

namespace Latinium
{
	/// <summary>
	/// Descripción breve de frmBodegas.
	/// </summary>
	public class frmBodegas : System.Windows.Forms.Form
	{
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtIdBodega;
		private System.Windows.Forms.Button btnNuevo;
		private System.Windows.Forms.Button btnConsultar;
		private System.Windows.Forms.Button btnEditar;
		private System.Windows.Forms.Button btnGuardar;
		private System.Windows.Forms.Button btnCancelar;
		private System.Windows.Forms.Button btnSalir;
		private System.Windows.Forms.Label lblCliente;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtCodigo;
		private System.Windows.Forms.Label label3;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtNombre;
		private System.Windows.Forms.Label label1;
		private Infragistics.Win.UltraWinSchedule.UltraCalendarCombo dtFecha;
		private System.Windows.Forms.Label label6;
		private System.Windows.Forms.Label label2;
		private System.Windows.Forms.Label label4;
		private System.Windows.Forms.Label label5;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbCiudad;
		private System.Windows.Forms.Label label7;
		private System.Windows.Forms.Label label8;
		private System.Windows.Forms.Label label9;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtEMail;
		private System.Windows.Forms.Label label11;
		private System.Windows.Forms.Label label10;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbCentroDeCosto;
		private Infragistics.Win.UltraWinEditors.UltraCheckEditor chkActivo;
		private Infragistics.Win.UltraWinEditors.UltraCheckEditor chkTransferencia;
		private Infragistics.Win.UltraWinEditors.UltraCheckEditor chkPrincipal;
		private Infragistics.Win.UltraWinEditors.UltraCheckEditor chkNCP;
		private Infragistics.Win.UltraWinEditors.UltraCheckEditor PTC;
		private Infragistics.Win.UltraWinEditors.UltraCheckEditor chkPromociones;
		private Infragistics.Win.UltraWinEditors.UltraCheckEditor chkCaja;
		private Infragistics.Win.UltraWinEditors.UltraCheckEditor chkPromocionesCombos;
		private Infragistics.Win.UltraWinEditors.UltraCheckEditor chkFS;
		private Infragistics.Win.UltraWinEditors.UltraCheckEditor chkCompras;
		private Infragistics.Win.UltraWinEditors.UltraCheckEditor chkTransferenciaMultiple;
		private Infragistics.Win.UltraWinEditors.UltraCheckEditor chkMatriz;
		private Infragistics.Win.UltraWinEditors.UltraCheckEditor chkDT;
		private Infragistics.Win.UltraWinEditors.UltraCheckEditor chkImprimePrecios;
		private Infragistics.Win.UltraWinEditors.UltraCheckEditor chkCredito;
		private Infragistics.Win.UltraWinEditors.UltraCheckEditor chkFactura;
		private Infragistics.Win.UltraWinEditors.UltraCheckEditor chkCompra;
		private Infragistics.Win.UltraWinEditors.UltraCheckEditor chkInventario;
		private System.Windows.Forms.Label cmbTipoDeAlmacen;
		private System.Windows.Forms.Label label12;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbSectores;
		private System.Windows.Forms.Label label13;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtSerie;
		private System.Windows.Forms.Label label14;
		private System.Windows.Forms.Label label15;
		private System.Windows.Forms.Label label16;
		private System.Windows.Forms.Label label17;
		private C1.Data.C1DataSet cdsSeteoF;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtNumero;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtcprovincia;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource1;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtExtencion;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtdireccion;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtOrdenSpf;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtCreditoGrande;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtContado;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtCreditoMediano;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtCreditoPequeño;
		private Infragistics.Win.UltraWinSchedule.UltraCalendarCombo dtcierre;
		private Infragistics.Win.UltraWinEditors.UltraCheckEditor chkDestinoT;
		private Infragistics.Win.UltraWinEditors.UltraCheckEditor chkCarpas;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtEquipo;
		private System.Windows.Forms.Label label18;
		private System.Windows.Forms.GroupBox groupBox1;
		private System.Windows.Forms.GroupBox groupBox2;
		private System.Windows.Forms.Panel panel1;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbCategorias;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource2;
		private Infragistics.Win.UltraWinEditors.UltraCheckEditor chkNAut;
		private System.Windows.Forms.Label label19;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbRegion;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource3;
		private Infragistics.Win.UltraWinEditors.UltraCheckEditor chkGRAut;
		private Infragistics.Win.UltraWinEditors.UltraCheckEditor chkBuffer2;
		//private Infragistics.Win.UltraWinEditors.UltraCheckEditor chkBuffer22;

		/// <summary>
		/// Variable del diseñador requerida.
		/// </summary>
		private System.ComponentModel.Container components = null;

		public frmBodegas()
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
			System.Resources.ResourceManager resources = new System.Resources.ResourceManager(typeof(frmBodegas));
			Infragistics.Win.Appearance appearance1 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance2 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance3 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton dateButton1 = new Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton();
			Infragistics.Win.Appearance appearance4 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton dateButton2 = new Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton();
			Infragistics.Win.Appearance appearance5 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance6 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand1 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn1 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCiudad");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn2 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Ciudad");
			Infragistics.Win.Appearance appearance7 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance8 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance9 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance10 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance11 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand2 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn3 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idProyecto");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn4 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.Appearance appearance12 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance13 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance14 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance15 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance16 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance17 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance18 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance19 = new Infragistics.Win.Appearance();
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
			Infragistics.Win.Appearance appearance30 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand3 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn5 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idTipoDeAlmacen");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn6 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("TipoDeAlmacen");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn1 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idTipoDeAlmacen");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn2 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("TipoDeAlmacen");
			Infragistics.Win.Appearance appearance31 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand4 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn7 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idSector");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn8 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Sector");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn3 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idSector");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn4 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Sector");
			Infragistics.Win.Appearance appearance32 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance33 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance34 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance35 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance36 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance37 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance38 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance39 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance40 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance41 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance42 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance43 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance44 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand5 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn9 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idRegion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn10 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Region");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn5 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idRegion");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn6 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Region");
			this.txtIdBodega = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.btnNuevo = new System.Windows.Forms.Button();
			this.btnConsultar = new System.Windows.Forms.Button();
			this.btnEditar = new System.Windows.Forms.Button();
			this.btnGuardar = new System.Windows.Forms.Button();
			this.btnCancelar = new System.Windows.Forms.Button();
			this.btnSalir = new System.Windows.Forms.Button();
			this.lblCliente = new System.Windows.Forms.Label();
			this.txtCodigo = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.label3 = new System.Windows.Forms.Label();
			this.txtNombre = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.label1 = new System.Windows.Forms.Label();
			this.dtFecha = new Infragistics.Win.UltraWinSchedule.UltraCalendarCombo();
			this.label6 = new System.Windows.Forms.Label();
			this.dtcierre = new Infragistics.Win.UltraWinSchedule.UltraCalendarCombo();
			this.label2 = new System.Windows.Forms.Label();
			this.txtOrdenSpf = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.label4 = new System.Windows.Forms.Label();
			this.label5 = new System.Windows.Forms.Label();
			this.cmbCiudad = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.label7 = new System.Windows.Forms.Label();
			this.txtcprovincia = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.label8 = new System.Windows.Forms.Label();
			this.txtdireccion = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.label9 = new System.Windows.Forms.Label();
			this.txtExtencion = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.txtEMail = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.label11 = new System.Windows.Forms.Label();
			this.label10 = new System.Windows.Forms.Label();
			this.cmbCentroDeCosto = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.chkActivo = new Infragistics.Win.UltraWinEditors.UltraCheckEditor();
			this.chkTransferencia = new Infragistics.Win.UltraWinEditors.UltraCheckEditor();
			this.chkPrincipal = new Infragistics.Win.UltraWinEditors.UltraCheckEditor();
			this.chkNCP = new Infragistics.Win.UltraWinEditors.UltraCheckEditor();
			this.PTC = new Infragistics.Win.UltraWinEditors.UltraCheckEditor();
			this.chkPromociones = new Infragistics.Win.UltraWinEditors.UltraCheckEditor();
			this.chkCaja = new Infragistics.Win.UltraWinEditors.UltraCheckEditor();
			this.chkPromocionesCombos = new Infragistics.Win.UltraWinEditors.UltraCheckEditor();
			this.chkFS = new Infragistics.Win.UltraWinEditors.UltraCheckEditor();
			this.chkCompras = new Infragistics.Win.UltraWinEditors.UltraCheckEditor();
			this.chkTransferenciaMultiple = new Infragistics.Win.UltraWinEditors.UltraCheckEditor();
			this.chkMatriz = new Infragistics.Win.UltraWinEditors.UltraCheckEditor();
			this.chkDT = new Infragistics.Win.UltraWinEditors.UltraCheckEditor();
			this.chkImprimePrecios = new Infragistics.Win.UltraWinEditors.UltraCheckEditor();
			this.chkCredito = new Infragistics.Win.UltraWinEditors.UltraCheckEditor();
			this.chkFactura = new Infragistics.Win.UltraWinEditors.UltraCheckEditor();
			this.chkCompra = new Infragistics.Win.UltraWinEditors.UltraCheckEditor();
			this.chkInventario = new Infragistics.Win.UltraWinEditors.UltraCheckEditor();
			this.cmbTipoDeAlmacen = new System.Windows.Forms.Label();
			this.cmbCategorias = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.ultraDataSource1 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.label12 = new System.Windows.Forms.Label();
			this.cmbSectores = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.ultraDataSource2 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.label13 = new System.Windows.Forms.Label();
			this.txtSerie = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.label14 = new System.Windows.Forms.Label();
			this.txtCreditoGrande = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.label15 = new System.Windows.Forms.Label();
			this.txtContado = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.label16 = new System.Windows.Forms.Label();
			this.txtCreditoMediano = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.label17 = new System.Windows.Forms.Label();
			this.txtCreditoPequeño = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.cdsSeteoF = new C1.Data.C1DataSet();
			this.txtNumero = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.chkDestinoT = new Infragistics.Win.UltraWinEditors.UltraCheckEditor();
			this.chkCarpas = new Infragistics.Win.UltraWinEditors.UltraCheckEditor();
			this.txtEquipo = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.label18 = new System.Windows.Forms.Label();
			this.groupBox1 = new System.Windows.Forms.GroupBox();
			this.chkGRAut = new Infragistics.Win.UltraWinEditors.UltraCheckEditor();
			this.chkNAut = new Infragistics.Win.UltraWinEditors.UltraCheckEditor();
			this.groupBox2 = new System.Windows.Forms.GroupBox();
			this.chkBuffer2 = new Infragistics.Win.UltraWinEditors.UltraCheckEditor();
			this.panel1 = new System.Windows.Forms.Panel();
			this.label19 = new System.Windows.Forms.Label();
			this.cmbRegion = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.ultraDataSource3 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			((System.ComponentModel.ISupportInitialize)(this.txtIdBodega)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtCodigo)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNombre)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.dtFecha)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.dtcierre)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtOrdenSpf)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbCiudad)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtcprovincia)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtdireccion)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtExtencion)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtEMail)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbCentroDeCosto)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbCategorias)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource1)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbSectores)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource2)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtSerie)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtCreditoGrande)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtContado)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtCreditoMediano)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtCreditoPequeño)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNumero)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtEquipo)).BeginInit();
			this.groupBox1.SuspendLayout();
			this.groupBox2.SuspendLayout();
			this.panel1.SuspendLayout();
			((System.ComponentModel.ISupportInitialize)(this.cmbRegion)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource3)).BeginInit();
			this.SuspendLayout();
			// 
			// txtIdBodega
			// 
			this.txtIdBodega.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.txtIdBodega.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtIdBodega.Enabled = false;
			this.txtIdBodega.Location = new System.Drawing.Point(8, 96);
			this.txtIdBodega.Name = "txtIdBodega";
			this.txtIdBodega.PromptChar = ' ';
			this.txtIdBodega.Size = new System.Drawing.Size(16, 21);
			this.txtIdBodega.SpinButtonAlignment = Infragistics.Win.ButtonAlignment.Left;
			this.txtIdBodega.TabIndex = 279;
			this.txtIdBodega.Visible = false;
			// 
			// btnNuevo
			// 
			this.btnNuevo.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btnNuevo.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(230)), ((System.Byte)(230)), ((System.Byte)(230)));
			this.btnNuevo.CausesValidation = false;
			this.btnNuevo.Enabled = false;
			this.btnNuevo.Font = new System.Drawing.Font("Tahoma", 8.25F);
			this.btnNuevo.ForeColor = System.Drawing.Color.FromArgb(((System.Byte)(50)), ((System.Byte)(50)), ((System.Byte)(50)));
			this.btnNuevo.Image = ((System.Drawing.Image)(resources.GetObject("btnNuevo.Image")));
			this.btnNuevo.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnNuevo.Location = new System.Drawing.Point(8, 480);
			this.btnNuevo.Name = "btnNuevo";
			this.btnNuevo.Size = new System.Drawing.Size(72, 23);
			this.btnNuevo.TabIndex = 273;
			this.btnNuevo.Text = "Nuevo";
			this.btnNuevo.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnNuevo.Click += new System.EventHandler(this.btnNuevo_Click);
			// 
			// btnConsultar
			// 
			this.btnConsultar.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btnConsultar.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(230)), ((System.Byte)(230)), ((System.Byte)(230)));
			this.btnConsultar.CausesValidation = false;
			this.btnConsultar.Enabled = false;
			this.btnConsultar.Font = new System.Drawing.Font("Tahoma", 8.25F);
			this.btnConsultar.ForeColor = System.Drawing.Color.FromArgb(((System.Byte)(50)), ((System.Byte)(50)), ((System.Byte)(50)));
			this.btnConsultar.Image = ((System.Drawing.Image)(resources.GetObject("btnConsultar.Image")));
			this.btnConsultar.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnConsultar.Location = new System.Drawing.Point(94, 480);
			this.btnConsultar.Name = "btnConsultar";
			this.btnConsultar.Size = new System.Drawing.Size(72, 23);
			this.btnConsultar.TabIndex = 274;
			this.btnConsultar.Text = "Consultar";
			this.btnConsultar.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnConsultar.Click += new System.EventHandler(this.btnConsultar_Click);
			// 
			// btnEditar
			// 
			this.btnEditar.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btnEditar.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(230)), ((System.Byte)(230)), ((System.Byte)(230)));
			this.btnEditar.CausesValidation = false;
			this.btnEditar.Enabled = false;
			this.btnEditar.Font = new System.Drawing.Font("Tahoma", 8.25F);
			this.btnEditar.ForeColor = System.Drawing.Color.FromArgb(((System.Byte)(50)), ((System.Byte)(50)), ((System.Byte)(50)));
			this.btnEditar.Image = ((System.Drawing.Image)(resources.GetObject("btnEditar.Image")));
			this.btnEditar.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnEditar.Location = new System.Drawing.Point(180, 480);
			this.btnEditar.Name = "btnEditar";
			this.btnEditar.Size = new System.Drawing.Size(72, 23);
			this.btnEditar.TabIndex = 276;
			this.btnEditar.Text = "Editar";
			this.btnEditar.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnEditar.Click += new System.EventHandler(this.btnEditar_Click);
			// 
			// btnGuardar
			// 
			this.btnGuardar.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btnGuardar.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(230)), ((System.Byte)(230)), ((System.Byte)(230)));
			this.btnGuardar.Font = new System.Drawing.Font("Tahoma", 8.25F);
			this.btnGuardar.ForeColor = System.Drawing.Color.FromArgb(((System.Byte)(50)), ((System.Byte)(50)), ((System.Byte)(50)));
			this.btnGuardar.Image = ((System.Drawing.Image)(resources.GetObject("btnGuardar.Image")));
			this.btnGuardar.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnGuardar.Location = new System.Drawing.Point(266, 480);
			this.btnGuardar.Name = "btnGuardar";
			this.btnGuardar.Size = new System.Drawing.Size(72, 23);
			this.btnGuardar.TabIndex = 275;
			this.btnGuardar.Text = "Guardar";
			this.btnGuardar.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnGuardar.Click += new System.EventHandler(this.btnGuardar_Click);
			// 
			// btnCancelar
			// 
			this.btnCancelar.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btnCancelar.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(230)), ((System.Byte)(230)), ((System.Byte)(230)));
			this.btnCancelar.CausesValidation = false;
			this.btnCancelar.Font = new System.Drawing.Font("Tahoma", 8.25F);
			this.btnCancelar.ForeColor = System.Drawing.Color.FromArgb(((System.Byte)(50)), ((System.Byte)(50)), ((System.Byte)(50)));
			this.btnCancelar.Image = ((System.Drawing.Image)(resources.GetObject("btnCancelar.Image")));
			this.btnCancelar.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnCancelar.Location = new System.Drawing.Point(352, 480);
			this.btnCancelar.Name = "btnCancelar";
			this.btnCancelar.Size = new System.Drawing.Size(72, 23);
			this.btnCancelar.TabIndex = 277;
			this.btnCancelar.Text = "Cancelar";
			this.btnCancelar.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnCancelar.Click += new System.EventHandler(this.btnCancelar_Click);
			// 
			// btnSalir
			// 
			this.btnSalir.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btnSalir.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(230)), ((System.Byte)(230)), ((System.Byte)(230)));
			this.btnSalir.CausesValidation = false;
			this.btnSalir.Font = new System.Drawing.Font("Tahoma", 8.25F);
			this.btnSalir.ForeColor = System.Drawing.Color.FromArgb(((System.Byte)(50)), ((System.Byte)(50)), ((System.Byte)(50)));
			this.btnSalir.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnSalir.ImageIndex = 62;
			this.btnSalir.Location = new System.Drawing.Point(438, 480);
			this.btnSalir.Name = "btnSalir";
			this.btnSalir.Size = new System.Drawing.Size(72, 23);
			this.btnSalir.TabIndex = 278;
			this.btnSalir.Text = "Salir";
			this.btnSalir.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnSalir.Click += new System.EventHandler(this.btnSalir_Click);
			// 
			// lblCliente
			// 
			this.lblCliente.AutoSize = true;
			this.lblCliente.Location = new System.Drawing.Point(208, 8);
			this.lblCliente.Name = "lblCliente";
			this.lblCliente.Size = new System.Drawing.Size(40, 16);
			this.lblCliente.TabIndex = 281;
			this.lblCliente.Text = "Codigo";
			// 
			// txtCodigo
			// 
			this.txtCodigo.AcceptsTab = true;
			appearance1.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtCodigo.Appearance = appearance1;
			this.txtCodigo.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtCodigo.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtCodigo.Enabled = false;
			this.txtCodigo.Location = new System.Drawing.Point(280, 8);
			this.txtCodigo.MaxLength = 2;
			this.txtCodigo.Name = "txtCodigo";
			this.txtCodigo.Size = new System.Drawing.Size(72, 21);
			this.txtCodigo.TabIndex = 280;
			this.txtCodigo.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtCodigo_KeyDown);
			// 
			// label3
			// 
			this.label3.AutoSize = true;
			this.label3.Location = new System.Drawing.Point(400, 8);
			this.label3.Name = "label3";
			this.label3.Size = new System.Drawing.Size(44, 16);
			this.label3.TabIndex = 283;
			this.label3.Text = "Nombre";
			// 
			// txtNombre
			// 
			appearance2.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtNombre.Appearance = appearance2;
			this.txtNombre.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtNombre.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtNombre.Enabled = false;
			this.txtNombre.Location = new System.Drawing.Point(496, 8);
			this.txtNombre.Name = "txtNombre";
			this.txtNombre.Size = new System.Drawing.Size(400, 21);
			this.txtNombre.TabIndex = 282;
			this.txtNombre.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtNombre_KeyDown);
			// 
			// label1
			// 
			this.label1.AutoSize = true;
			this.label1.Location = new System.Drawing.Point(8, 8);
			this.label1.Name = "label1";
			this.label1.Size = new System.Drawing.Size(44, 16);
			this.label1.TabIndex = 284;
			this.label1.Text = "Numero";
			// 
			// dtFecha
			// 
			appearance3.ForeColorDisabled = System.Drawing.Color.Black;
			this.dtFecha.Appearance = appearance3;
			this.dtFecha.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			dateButton1.Caption = "Today";
			this.dtFecha.DateButtons.Add(dateButton1);
			this.dtFecha.Enabled = false;
			this.dtFecha.Format = "dd/MM/yyyy";
			this.dtFecha.Location = new System.Drawing.Point(80, 32);
			this.dtFecha.Name = "dtFecha";
			this.dtFecha.NonAutoSizeHeight = 23;
			this.dtFecha.Size = new System.Drawing.Size(112, 21);
			this.dtFecha.SpinButtonsVisible = true;
			this.dtFecha.TabIndex = 286;
			this.dtFecha.Value = ((object)(resources.GetObject("dtFecha.Value")));
			this.dtFecha.KeyDown += new System.Windows.Forms.KeyEventHandler(this.dtFecha_KeyDown);
			// 
			// label6
			// 
			this.label6.AutoSize = true;
			this.label6.Location = new System.Drawing.Point(8, 32);
			this.label6.Name = "label6";
			this.label6.Size = new System.Drawing.Size(48, 16);
			this.label6.TabIndex = 287;
			this.label6.Text = "Apertura";
			// 
			// dtcierre
			// 
			appearance4.ForeColorDisabled = System.Drawing.Color.Black;
			this.dtcierre.Appearance = appearance4;
			this.dtcierre.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			dateButton2.Caption = "Today";
			this.dtcierre.DateButtons.Add(dateButton2);
			this.dtcierre.Enabled = false;
			this.dtcierre.Format = "dd/MM/yyyy";
			this.dtcierre.Location = new System.Drawing.Point(280, 32);
			this.dtcierre.Name = "dtcierre";
			this.dtcierre.NonAutoSizeHeight = 23;
			this.dtcierre.Size = new System.Drawing.Size(112, 21);
			this.dtcierre.SpinButtonsVisible = true;
			this.dtcierre.TabIndex = 288;
			this.dtcierre.Value = ((object)(resources.GetObject("dtcierre.Value")));
			// 
			// label2
			// 
			this.label2.AutoSize = true;
			this.label2.Location = new System.Drawing.Point(208, 32);
			this.label2.Name = "label2";
			this.label2.Size = new System.Drawing.Size(35, 16);
			this.label2.TabIndex = 289;
			this.label2.Text = "Cierre";
			// 
			// txtOrdenSpf
			// 
			appearance5.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtOrdenSpf.Appearance = appearance5;
			this.txtOrdenSpf.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtOrdenSpf.Enabled = false;
			this.txtOrdenSpf.Location = new System.Drawing.Point(664, 56);
			this.txtOrdenSpf.MaxValue = 999;
			this.txtOrdenSpf.Name = "txtOrdenSpf";
			this.txtOrdenSpf.PromptChar = ' ';
			this.txtOrdenSpf.Size = new System.Drawing.Size(72, 21);
			this.txtOrdenSpf.TabIndex = 291;
			this.txtOrdenSpf.Visible = false;
			this.txtOrdenSpf.KeyDown += new System.Windows.Forms.KeyEventHandler(this.ultraNumericEditor1_KeyDown);
			// 
			// label4
			// 
			this.label4.AutoSize = true;
			this.label4.Location = new System.Drawing.Point(592, 58);
			this.label4.Name = "label4";
			this.label4.Size = new System.Drawing.Size(64, 16);
			this.label4.TabIndex = 290;
			this.label4.Text = "Orden Spiff ";
			// 
			// label5
			// 
			this.label5.AutoSize = true;
			this.label5.BackColor = System.Drawing.Color.Transparent;
			this.label5.Location = new System.Drawing.Point(400, 32);
			this.label5.Name = "label5";
			this.label5.Size = new System.Drawing.Size(43, 16);
			this.label5.TabIndex = 292;
			this.label5.Text = "Ciudad:";
			this.label5.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// cmbCiudad
			// 
			appearance6.ForeColorDisabled = System.Drawing.Color.Black;
			this.cmbCiudad.Appearance = appearance6;
			this.cmbCiudad.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.cmbCiudad.Cursor = System.Windows.Forms.Cursors.Default;
			ultraGridColumn1.Header.VisiblePosition = 0;
			ultraGridColumn1.Hidden = true;
			ultraGridColumn2.Header.VisiblePosition = 1;
			ultraGridColumn2.Width = 276;
			ultraGridBand1.Columns.AddRange(new object[] {
																										 ultraGridColumn1,
																										 ultraGridColumn2});
			this.cmbCiudad.DisplayLayout.BandsSerializer.Add(ultraGridBand1);
			this.cmbCiudad.DisplayMember = "Ciudad";
			this.cmbCiudad.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbCiudad.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbCiudad.Enabled = false;
			this.cmbCiudad.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbCiudad.Location = new System.Drawing.Point(496, 32);
			this.cmbCiudad.Name = "cmbCiudad";
			this.cmbCiudad.Size = new System.Drawing.Size(240, 21);
			this.cmbCiudad.TabIndex = 293;
			this.cmbCiudad.ValueMember = "idCiudad";
			this.cmbCiudad.KeyDown += new System.Windows.Forms.KeyEventHandler(this.cmbCiudad_KeyDown);
			// 
			// label7
			// 
			this.label7.AutoSize = true;
			this.label7.Location = new System.Drawing.Point(752, 32);
			this.label7.Name = "label7";
			this.label7.Size = new System.Drawing.Size(77, 16);
			this.label7.TabIndex = 295;
			this.label7.Text = "Cod Provincial";
			// 
			// txtcprovincia
			// 
			this.txtcprovincia.AcceptsTab = true;
			appearance7.BackColor = System.Drawing.Color.White;
			appearance7.BackColorDisabled = System.Drawing.Color.Transparent;
			appearance7.BackColorDisabled2 = System.Drawing.Color.Transparent;
			appearance7.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtcprovincia.Appearance = appearance7;
			this.txtcprovincia.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtcprovincia.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtcprovincia.Enabled = false;
			this.txtcprovincia.Location = new System.Drawing.Point(840, 32);
			this.txtcprovincia.MaxLength = 2;
			this.txtcprovincia.Name = "txtcprovincia";
			this.txtcprovincia.Size = new System.Drawing.Size(56, 21);
			this.txtcprovincia.TabIndex = 294;
			// 
			// label8
			// 
			this.label8.AutoSize = true;
			this.label8.Location = new System.Drawing.Point(8, 72);
			this.label8.Name = "label8";
			this.label8.Size = new System.Drawing.Size(52, 16);
			this.label8.TabIndex = 297;
			this.label8.Text = "Dirección";
			// 
			// txtdireccion
			// 
			appearance8.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtdireccion.Appearance = appearance8;
			this.txtdireccion.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtdireccion.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtdireccion.Enabled = false;
			this.txtdireccion.Location = new System.Drawing.Point(80, 56);
			this.txtdireccion.Multiline = true;
			this.txtdireccion.Name = "txtdireccion";
			this.txtdireccion.Size = new System.Drawing.Size(312, 63);
			this.txtdireccion.TabIndex = 296;
			this.txtdireccion.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtdireccion_KeyDown);
			// 
			// label9
			// 
			this.label9.AutoSize = true;
			this.label9.BackColor = System.Drawing.Color.Transparent;
			this.label9.Location = new System.Drawing.Point(400, 56);
			this.label9.Name = "label9";
			this.label9.Size = new System.Drawing.Size(57, 16);
			this.label9.TabIndex = 299;
			this.label9.Text = "Extension:";
			this.label9.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// txtExtencion
			// 
			appearance9.BackColorDisabled = System.Drawing.Color.Transparent;
			appearance9.BackColorDisabled2 = System.Drawing.Color.Transparent;
			appearance9.BorderColor = System.Drawing.Color.Black;
			appearance9.BorderColor3DBase = System.Drawing.Color.Black;
			appearance9.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtExtencion.Appearance = appearance9;
			this.txtExtencion.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtExtencion.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtExtencion.Enabled = false;
			this.txtExtencion.Location = new System.Drawing.Point(496, 56);
			this.txtExtencion.MaxLength = 4;
			this.txtExtencion.Name = "txtExtencion";
			this.txtExtencion.Size = new System.Drawing.Size(96, 21);
			this.txtExtencion.TabIndex = 298;
			this.txtExtencion.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtTelefono_KeyDown);
			this.txtExtencion.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.txtExtencion_KeyPress);
			// 
			// txtEMail
			// 
			appearance10.BorderColor = System.Drawing.Color.Black;
			appearance10.BorderColor3DBase = System.Drawing.Color.Black;
			appearance10.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtEMail.Appearance = appearance10;
			this.txtEMail.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtEMail.Enabled = false;
			this.txtEMail.Location = new System.Drawing.Point(688, 56);
			this.txtEMail.Name = "txtEMail";
			this.txtEMail.Size = new System.Drawing.Size(208, 21);
			this.txtEMail.TabIndex = 300;
			this.txtEMail.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtEMail_KeyDown);
			// 
			// label11
			// 
			this.label11.AutoSize = true;
			this.label11.BackColor = System.Drawing.Color.Transparent;
			this.label11.Location = new System.Drawing.Point(632, 56);
			this.label11.Name = "label11";
			this.label11.Size = new System.Drawing.Size(36, 16);
			this.label11.TabIndex = 301;
			this.label11.Text = "Email:";
			this.label11.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// label10
			// 
			this.label10.AutoSize = true;
			this.label10.Location = new System.Drawing.Point(400, 80);
			this.label10.Name = "label10";
			this.label10.Size = new System.Drawing.Size(86, 16);
			this.label10.TabIndex = 303;
			this.label10.Text = "Centro de Costo";
			this.label10.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// cmbCentroDeCosto
			// 
			appearance11.ForeColorDisabled = System.Drawing.Color.Black;
			this.cmbCentroDeCosto.Appearance = appearance11;
			this.cmbCentroDeCosto.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbCentroDeCosto.Cursor = System.Windows.Forms.Cursors.Default;
			ultraGridColumn3.Header.VisiblePosition = 0;
			ultraGridColumn3.Hidden = true;
			ultraGridColumn4.Header.VisiblePosition = 1;
			ultraGridColumn4.Width = 300;
			ultraGridBand2.Columns.AddRange(new object[] {
																										 ultraGridColumn3,
																										 ultraGridColumn4});
			this.cmbCentroDeCosto.DisplayLayout.BandsSerializer.Add(ultraGridBand2);
			this.cmbCentroDeCosto.DisplayMember = "Nombre";
			this.cmbCentroDeCosto.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbCentroDeCosto.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbCentroDeCosto.Enabled = false;
			this.cmbCentroDeCosto.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbCentroDeCosto.Location = new System.Drawing.Point(496, 80);
			this.cmbCentroDeCosto.Name = "cmbCentroDeCosto";
			this.cmbCentroDeCosto.Size = new System.Drawing.Size(300, 21);
			this.cmbCentroDeCosto.TabIndex = 302;
			this.cmbCentroDeCosto.ValueMember = "idProyecto";
			this.cmbCentroDeCosto.KeyDown += new System.Windows.Forms.KeyEventHandler(this.cmbCentroDeCosto_KeyDown);
			// 
			// chkActivo
			// 
			appearance12.ForeColorDisabled = System.Drawing.Color.Black;
			this.chkActivo.Appearance = appearance12;
			this.chkActivo.BackColor = System.Drawing.Color.Transparent;
			this.chkActivo.Enabled = false;
			this.chkActivo.Location = new System.Drawing.Point(712, 104);
			this.chkActivo.Name = "chkActivo";
			this.chkActivo.Size = new System.Drawing.Size(56, 21);
			this.chkActivo.TabIndex = 304;
			this.chkActivo.Text = "Activo";
			// 
			// chkTransferencia
			// 
			appearance13.ForeColorDisabled = System.Drawing.Color.Black;
			this.chkTransferencia.Appearance = appearance13;
			this.chkTransferencia.BackColor = System.Drawing.Color.Transparent;
			this.chkTransferencia.Enabled = false;
			this.chkTransferencia.Location = new System.Drawing.Point(64, 32);
			this.chkTransferencia.Name = "chkTransferencia";
			this.chkTransferencia.Size = new System.Drawing.Size(96, 21);
			this.chkTransferencia.TabIndex = 305;
			this.chkTransferencia.Text = "Transferencia";
			this.chkTransferencia.KeyDown += new System.Windows.Forms.KeyEventHandler(this.chkTransferencia_KeyDown);
			// 
			// chkPrincipal
			// 
			appearance14.ForeColorDisabled = System.Drawing.Color.Black;
			this.chkPrincipal.Appearance = appearance14;
			this.chkPrincipal.BackColor = System.Drawing.Color.Transparent;
			this.chkPrincipal.Enabled = false;
			this.chkPrincipal.Location = new System.Drawing.Point(400, 32);
			this.chkPrincipal.Name = "chkPrincipal";
			this.chkPrincipal.Size = new System.Drawing.Size(72, 21);
			this.chkPrincipal.TabIndex = 307;
			this.chkPrincipal.Text = "Principal";
			this.chkPrincipal.KeyDown += new System.Windows.Forms.KeyEventHandler(this.chkPrincipal_KeyDown);
			this.chkPrincipal.Click += new System.EventHandler(this.chkPrincipal_Click);
			this.chkPrincipal.CheckedChanged += new System.EventHandler(this.chkPrincipal_CheckedChanged);
			// 
			// chkNCP
			// 
			appearance15.ForeColorDisabled = System.Drawing.Color.Black;
			this.chkNCP.Appearance = appearance15;
			this.chkNCP.BackColor = System.Drawing.Color.Transparent;
			this.chkNCP.Enabled = false;
			this.chkNCP.Location = new System.Drawing.Point(184, 56);
			this.chkNCP.Name = "chkNCP";
			this.chkNCP.Size = new System.Drawing.Size(160, 21);
			this.chkNCP.TabIndex = 306;
			this.chkNCP.Text = "Visible en Notas de Crédito";
			this.chkNCP.KeyDown += new System.Windows.Forms.KeyEventHandler(this.chkNCP_KeyDown);
			// 
			// PTC
			// 
			appearance16.ForeColorDisabled = System.Drawing.Color.Black;
			this.PTC.Appearance = appearance16;
			this.PTC.BackColor = System.Drawing.Color.Transparent;
			this.PTC.Enabled = false;
			this.PTC.Location = new System.Drawing.Point(128, 96);
			this.PTC.Name = "PTC";
			this.PTC.Size = new System.Drawing.Size(176, 21);
			this.PTC.TabIndex = 309;
			this.PTC.Text = "Imprime con Precio de Crédito";
			this.PTC.KeyDown += new System.Windows.Forms.KeyEventHandler(this.PTC_KeyDown);
			// 
			// chkPromociones
			// 
			appearance17.ForeColorDisabled = System.Drawing.Color.Black;
			this.chkPromociones.Appearance = appearance17;
			this.chkPromociones.BackColor = System.Drawing.Color.Transparent;
			this.chkPromociones.Enabled = false;
			this.chkPromociones.Location = new System.Drawing.Point(176, 64);
			this.chkPromociones.Name = "chkPromociones";
			this.chkPromociones.Size = new System.Drawing.Size(88, 21);
			this.chkPromociones.TabIndex = 308;
			this.chkPromociones.Text = "Promociones";
			this.chkPromociones.KeyDown += new System.Windows.Forms.KeyEventHandler(this.chkPromociones_KeyDown);
			// 
			// chkCaja
			// 
			appearance18.ForeColorDisabled = System.Drawing.Color.Black;
			this.chkCaja.Appearance = appearance18;
			this.chkCaja.BackColor = System.Drawing.Color.Transparent;
			this.chkCaja.Enabled = false;
			this.chkCaja.Location = new System.Drawing.Point(840, 104);
			this.chkCaja.Name = "chkCaja";
			this.chkCaja.Size = new System.Drawing.Size(48, 21);
			this.chkCaja.TabIndex = 315;
			this.chkCaja.Text = "Caja";
			this.chkCaja.KeyDown += new System.Windows.Forms.KeyEventHandler(this.chkCaja_KeyDown);
			// 
			// chkPromocionesCombos
			// 
			appearance19.ForeColorDisabled = System.Drawing.Color.Black;
			this.chkPromocionesCombos.Appearance = appearance19;
			this.chkPromocionesCombos.BackColor = System.Drawing.Color.Transparent;
			this.chkPromocionesCombos.Enabled = false;
			this.chkPromocionesCombos.Location = new System.Drawing.Point(264, 64);
			this.chkPromocionesCombos.Name = "chkPromocionesCombos";
			this.chkPromocionesCombos.Size = new System.Drawing.Size(152, 21);
			this.chkPromocionesCombos.TabIndex = 314;
			this.chkPromocionesCombos.Text = "Promociones de Combos";
			this.chkPromocionesCombos.KeyDown += new System.Windows.Forms.KeyEventHandler(this.chkPromocionesCombos_KeyDown);
			// 
			// chkFS
			// 
			appearance20.ForeColorDisabled = System.Drawing.Color.Black;
			this.chkFS.Appearance = appearance20;
			this.chkFS.BackColor = System.Drawing.Color.Transparent;
			this.chkFS.Enabled = false;
			this.chkFS.Location = new System.Drawing.Point(416, 64);
			this.chkFS.Name = "chkFS";
			this.chkFS.Size = new System.Drawing.Size(128, 21);
			this.chkFS.TabIndex = 313;
			this.chkFS.Text = "Factura Con Seriales";
			this.chkFS.KeyDown += new System.Windows.Forms.KeyEventHandler(this.chkFS_KeyDown);
			// 
			// chkCompras
			// 
			appearance21.ForeColorDisabled = System.Drawing.Color.Black;
			this.chkCompras.Appearance = appearance21;
			this.chkCompras.BackColor = System.Drawing.Color.Transparent;
			this.chkCompras.Enabled = false;
			this.chkCompras.Location = new System.Drawing.Point(344, 56);
			this.chkCompras.Name = "chkCompras";
			this.chkCompras.Size = new System.Drawing.Size(248, 21);
			this.chkCompras.TabIndex = 312;
			this.chkCompras.Text = "Visible en SPIFF y Actualización de Precios";
			this.chkCompras.KeyDown += new System.Windows.Forms.KeyEventHandler(this.chkCompras_KeyDown);
			this.chkCompras.CheckedChanged += new System.EventHandler(this.chkCompras_CheckedChanged);
			// 
			// chkTransferenciaMultiple
			// 
			appearance22.ForeColorDisabled = System.Drawing.Color.Black;
			this.chkTransferenciaMultiple.Appearance = appearance22;
			this.chkTransferenciaMultiple.BackColor = System.Drawing.Color.Transparent;
			this.chkTransferenciaMultiple.Enabled = false;
			this.chkTransferenciaMultiple.Location = new System.Drawing.Point(184, 32);
			this.chkTransferenciaMultiple.Name = "chkTransferenciaMultiple";
			this.chkTransferenciaMultiple.Size = new System.Drawing.Size(136, 21);
			this.chkTransferenciaMultiple.TabIndex = 311;
			this.chkTransferenciaMultiple.Text = "Transferencia Multiple";
			this.chkTransferenciaMultiple.KeyDown += new System.Windows.Forms.KeyEventHandler(this.chkTransferenciaMultiple_KeyDown);
			// 
			// chkMatriz
			// 
			appearance23.ForeColorDisabled = System.Drawing.Color.Black;
			this.chkMatriz.Appearance = appearance23;
			this.chkMatriz.BackColor = System.Drawing.Color.Transparent;
			this.chkMatriz.Enabled = false;
			this.chkMatriz.Location = new System.Drawing.Point(776, 104);
			this.chkMatriz.Name = "chkMatriz";
			this.chkMatriz.Size = new System.Drawing.Size(56, 21);
			this.chkMatriz.TabIndex = 310;
			this.chkMatriz.Text = "Matriz";
			this.chkMatriz.KeyDown += new System.Windows.Forms.KeyEventHandler(this.chkMatriz_KeyDown);
			// 
			// chkDT
			// 
			appearance24.ForeColorDisabled = System.Drawing.Color.Black;
			this.chkDT.Appearance = appearance24;
			this.chkDT.BackColor = System.Drawing.Color.Transparent;
			this.chkDT.Enabled = false;
			this.chkDT.Location = new System.Drawing.Point(64, 80);
			this.chkDT.Name = "chkDT";
			this.chkDT.Size = new System.Drawing.Size(184, 21);
			this.chkDT.TabIndex = 321;
			this.chkDT.Text = "Ingresos Departamento Técnico";
			this.chkDT.KeyDown += new System.Windows.Forms.KeyEventHandler(this.chkDT_KeyDown);
			this.chkDT.CheckedChanged += new System.EventHandler(this.chkDT_CheckedChanged);
			// 
			// chkImprimePrecios
			// 
			appearance25.ForeColorDisabled = System.Drawing.Color.Black;
			this.chkImprimePrecios.Appearance = appearance25;
			this.chkImprimePrecios.BackColor = System.Drawing.Color.Transparent;
			this.chkImprimePrecios.Enabled = false;
			this.chkImprimePrecios.Location = new System.Drawing.Point(16, 96);
			this.chkImprimePrecios.Name = "chkImprimePrecios";
			this.chkImprimePrecios.Size = new System.Drawing.Size(104, 21);
			this.chkImprimePrecios.TabIndex = 320;
			this.chkImprimePrecios.Text = "Imprime Precios";
			this.chkImprimePrecios.KeyDown += new System.Windows.Forms.KeyEventHandler(this.chkImprimePrecios_KeyDown);
			this.chkImprimePrecios.CheckedChanged += new System.EventHandler(this.chkImprimePrecios_CheckedChanged);
			// 
			// chkCredito
			// 
			appearance26.ForeColorDisabled = System.Drawing.Color.Black;
			this.chkCredito.Appearance = appearance26;
			this.chkCredito.BackColor = System.Drawing.Color.Transparent;
			this.chkCredito.Enabled = false;
			this.chkCredito.Location = new System.Drawing.Point(40, 64);
			this.chkCredito.Name = "chkCredito";
			this.chkCredito.Size = new System.Drawing.Size(64, 21);
			this.chkCredito.TabIndex = 319;
			this.chkCredito.Text = "Crédito";
			this.chkCredito.KeyDown += new System.Windows.Forms.KeyEventHandler(this.chkCredito_KeyDown);
			// 
			// chkFactura
			// 
			appearance27.ForeColorDisabled = System.Drawing.Color.Black;
			this.chkFactura.Appearance = appearance27;
			this.chkFactura.BackColor = System.Drawing.Color.Transparent;
			this.chkFactura.Enabled = false;
			this.chkFactura.Location = new System.Drawing.Point(16, 16);
			this.chkFactura.Name = "chkFactura";
			this.chkFactura.Size = new System.Drawing.Size(16, 21);
			this.chkFactura.TabIndex = 318;
			this.chkFactura.KeyDown += new System.Windows.Forms.KeyEventHandler(this.chkFactura_KeyDown);
			this.chkFactura.CheckedChanged += new System.EventHandler(this.chkFactura_CheckedChanged);
			// 
			// chkCompra
			// 
			appearance28.ForeColorDisabled = System.Drawing.Color.Black;
			this.chkCompra.Appearance = appearance28;
			this.chkCompra.BackColor = System.Drawing.Color.Transparent;
			this.chkCompra.Enabled = false;
			this.chkCompra.Location = new System.Drawing.Point(72, 56);
			this.chkCompra.Name = "chkCompra";
			this.chkCompra.Size = new System.Drawing.Size(120, 21);
			this.chkCompra.TabIndex = 317;
			this.chkCompra.Text = "Visible en Compras";
			this.chkCompra.KeyDown += new System.Windows.Forms.KeyEventHandler(this.chkCompra_KeyDown);
			this.chkCompra.CheckedChanged += new System.EventHandler(this.chkCompra_CheckedChanged);
			// 
			// chkInventario
			// 
			appearance29.ForeColorDisabled = System.Drawing.Color.Black;
			this.chkInventario.Appearance = appearance29;
			this.chkInventario.BackColor = System.Drawing.Color.Transparent;
			this.chkInventario.Enabled = false;
			this.chkInventario.Location = new System.Drawing.Point(16, 16);
			this.chkInventario.Name = "chkInventario";
			this.chkInventario.Size = new System.Drawing.Size(16, 21);
			this.chkInventario.TabIndex = 316;
			this.chkInventario.KeyDown += new System.Windows.Forms.KeyEventHandler(this.chkInventario_KeyDown);
			this.chkInventario.CheckedChanged += new System.EventHandler(this.chkInventario_CheckedChanged);
			// 
			// cmbTipoDeAlmacen
			// 
			this.cmbTipoDeAlmacen.AutoSize = true;
			this.cmbTipoDeAlmacen.Location = new System.Drawing.Point(40, 34);
			this.cmbTipoDeAlmacen.Name = "cmbTipoDeAlmacen";
			this.cmbTipoDeAlmacen.Size = new System.Drawing.Size(53, 16);
			this.cmbTipoDeAlmacen.TabIndex = 323;
			this.cmbTipoDeAlmacen.Text = "Categoria";
			this.cmbTipoDeAlmacen.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// cmbCategorias
			// 
			appearance30.ForeColorDisabled = System.Drawing.Color.Black;
			this.cmbCategorias.Appearance = appearance30;
			this.cmbCategorias.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbCategorias.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbCategorias.DataSource = this.ultraDataSource1;
			ultraGridColumn5.Header.VisiblePosition = 0;
			ultraGridColumn5.Hidden = true;
			ultraGridColumn6.Header.Caption = "Tipo De Almacen";
			ultraGridColumn6.Header.VisiblePosition = 1;
			ultraGridColumn6.Width = 75;
			ultraGridBand3.Columns.AddRange(new object[] {
																										 ultraGridColumn5,
																										 ultraGridColumn6});
			this.cmbCategorias.DisplayLayout.BandsSerializer.Add(ultraGridBand3);
			this.cmbCategorias.DisplayMember = "TipoDeAlmacen";
			this.cmbCategorias.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbCategorias.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbCategorias.Enabled = false;
			this.cmbCategorias.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbCategorias.Location = new System.Drawing.Point(104, 32);
			this.cmbCategorias.Name = "cmbCategorias";
			this.cmbCategorias.Size = new System.Drawing.Size(75, 21);
			this.cmbCategorias.TabIndex = 322;
			this.cmbCategorias.ValueMember = "idTipoDeAlmacen";
			this.cmbCategorias.KeyDown += new System.Windows.Forms.KeyEventHandler(this.cmbCategorias_KeyDown);
			// 
			// ultraDataSource1
			// 
			ultraDataColumn1.DataType = typeof(int);
			this.ultraDataSource1.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn1,
																																 ultraDataColumn2});
			// 
			// label12
			// 
			this.label12.AutoSize = true;
			this.label12.Location = new System.Drawing.Point(192, 34);
			this.label12.Name = "label12";
			this.label12.Size = new System.Drawing.Size(37, 16);
			this.label12.TabIndex = 325;
			this.label12.Text = "Sector";
			this.label12.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// cmbSectores
			// 
			appearance31.ForeColorDisabled = System.Drawing.Color.Black;
			this.cmbSectores.Appearance = appearance31;
			this.cmbSectores.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbSectores.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbSectores.DataSource = this.ultraDataSource2;
			ultraGridColumn7.Header.VisiblePosition = 0;
			ultraGridColumn7.Hidden = true;
			ultraGridColumn8.Header.VisiblePosition = 1;
			ultraGridColumn8.Width = 140;
			ultraGridBand4.Columns.AddRange(new object[] {
																										 ultraGridColumn7,
																										 ultraGridColumn8});
			this.cmbSectores.DisplayLayout.BandsSerializer.Add(ultraGridBand4);
			this.cmbSectores.DisplayMember = "Sector";
			this.cmbSectores.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbSectores.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbSectores.Enabled = false;
			this.cmbSectores.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbSectores.Location = new System.Drawing.Point(240, 32);
			this.cmbSectores.Name = "cmbSectores";
			this.cmbSectores.Size = new System.Drawing.Size(140, 21);
			this.cmbSectores.TabIndex = 324;
			this.cmbSectores.ValueMember = "idSector";
			this.cmbSectores.KeyDown += new System.Windows.Forms.KeyEventHandler(this.cmbSectores_KeyDown);
			// 
			// ultraDataSource2
			// 
			ultraDataColumn3.DataType = typeof(int);
			this.ultraDataSource2.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn3,
																																 ultraDataColumn4});
			// 
			// label13
			// 
			this.label13.AutoSize = true;
			this.label13.Location = new System.Drawing.Point(480, 34);
			this.label13.Name = "label13";
			this.label13.Size = new System.Drawing.Size(31, 16);
			this.label13.TabIndex = 327;
			this.label13.Text = "Serie";
			// 
			// txtSerie
			// 
			this.txtSerie.AcceptsTab = true;
			appearance32.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtSerie.Appearance = appearance32;
			this.txtSerie.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtSerie.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtSerie.Enabled = false;
			this.txtSerie.Location = new System.Drawing.Point(528, 32);
			this.txtSerie.MaxLength = 6;
			this.txtSerie.Name = "txtSerie";
			this.txtSerie.Size = new System.Drawing.Size(72, 21);
			this.txtSerie.TabIndex = 326;
			this.txtSerie.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtSerie_KeyDown);
			this.txtSerie.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.txtSerie_KeyPress);
			// 
			// label14
			// 
			this.label14.AutoSize = true;
			this.label14.Location = new System.Drawing.Point(32, 128);
			this.label14.Name = "label14";
			this.label14.Size = new System.Drawing.Size(82, 16);
			this.label14.TabIndex = 329;
			this.label14.Text = "Credito Grande";
			// 
			// txtCreditoGrande
			// 
			appearance33.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtCreditoGrande.Appearance = appearance33;
			this.txtCreditoGrande.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtCreditoGrande.Enabled = false;
			this.txtCreditoGrande.Location = new System.Drawing.Point(136, 128);
			this.txtCreditoGrande.Name = "txtCreditoGrande";
			this.txtCreditoGrande.Size = new System.Drawing.Size(300, 21);
			this.txtCreditoGrande.TabIndex = 328;
			// 
			// label15
			// 
			this.label15.AutoSize = true;
			this.label15.Location = new System.Drawing.Point(456, 128);
			this.label15.Name = "label15";
			this.label15.Size = new System.Drawing.Size(90, 16);
			this.label15.TabIndex = 331;
			this.label15.Text = "Credito Mediano ";
			// 
			// txtContado
			// 
			appearance34.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtContado.Appearance = appearance34;
			this.txtContado.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtContado.Enabled = false;
			this.txtContado.Location = new System.Drawing.Point(568, 152);
			this.txtContado.Name = "txtContado";
			this.txtContado.Size = new System.Drawing.Size(300, 21);
			this.txtContado.TabIndex = 330;
			// 
			// label16
			// 
			this.label16.AutoSize = true;
			this.label16.Location = new System.Drawing.Point(32, 152);
			this.label16.Name = "label16";
			this.label16.Size = new System.Drawing.Size(89, 16);
			this.label16.TabIndex = 333;
			this.label16.Text = "Credito Pequeño";
			// 
			// txtCreditoMediano
			// 
			appearance35.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtCreditoMediano.Appearance = appearance35;
			this.txtCreditoMediano.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtCreditoMediano.Enabled = false;
			this.txtCreditoMediano.Location = new System.Drawing.Point(568, 128);
			this.txtCreditoMediano.Name = "txtCreditoMediano";
			this.txtCreditoMediano.Size = new System.Drawing.Size(300, 21);
			this.txtCreditoMediano.TabIndex = 332;
			// 
			// label17
			// 
			this.label17.AutoSize = true;
			this.label17.Location = new System.Drawing.Point(456, 152);
			this.label17.Name = "label17";
			this.label17.Size = new System.Drawing.Size(47, 16);
			this.label17.TabIndex = 335;
			this.label17.Text = "Contado";
			// 
			// txtCreditoPequeño
			// 
			appearance36.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtCreditoPequeño.Appearance = appearance36;
			this.txtCreditoPequeño.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtCreditoPequeño.Enabled = false;
			this.txtCreditoPequeño.Location = new System.Drawing.Point(136, 152);
			this.txtCreditoPequeño.Name = "txtCreditoPequeño";
			this.txtCreditoPequeño.Size = new System.Drawing.Size(300, 21);
			this.txtCreditoPequeño.TabIndex = 334;
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
			// txtNumero
			// 
			appearance37.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtNumero.Appearance = appearance37;
			this.txtNumero.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtNumero.Enabled = false;
			this.txtNumero.Location = new System.Drawing.Point(80, 8);
			this.txtNumero.Name = "txtNumero";
			this.txtNumero.PromptChar = ' ';
			this.txtNumero.Size = new System.Drawing.Size(80, 21);
			this.txtNumero.TabIndex = 285;
			// 
			// chkDestinoT
			// 
			appearance38.ForeColorDisabled = System.Drawing.Color.Black;
			this.chkDestinoT.Appearance = appearance38;
			this.chkDestinoT.BackColor = System.Drawing.Color.Transparent;
			this.chkDestinoT.Enabled = false;
			this.chkDestinoT.Location = new System.Drawing.Point(344, 32);
			this.chkDestinoT.Name = "chkDestinoT";
			this.chkDestinoT.Size = new System.Drawing.Size(152, 21);
			this.chkDestinoT.TabIndex = 338;
			this.chkDestinoT.Text = "Destino de Transferencia";
			this.chkDestinoT.KeyDown += new System.Windows.Forms.KeyEventHandler(this.chkDestinoT_KeyDown_1);
			// 
			// chkCarpas
			// 
			appearance39.ForeColorDisabled = System.Drawing.Color.Black;
			this.chkCarpas.Appearance = appearance39;
			this.chkCarpas.BackColor = System.Drawing.Color.Transparent;
			this.chkCarpas.Enabled = false;
			this.chkCarpas.Location = new System.Drawing.Point(112, 64);
			this.chkCarpas.Name = "chkCarpas";
			this.chkCarpas.Size = new System.Drawing.Size(64, 21);
			this.chkCarpas.TabIndex = 339;
			this.chkCarpas.Text = "Carpas";
			this.chkCarpas.KeyDown += new System.Windows.Forms.KeyEventHandler(this.chkCarpas_KeyDown);
			// 
			// txtEquipo
			// 
			appearance40.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtEquipo.Appearance = appearance40;
			this.txtEquipo.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtEquipo.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtEquipo.Enabled = false;
			this.txtEquipo.Location = new System.Drawing.Point(672, 32);
			this.txtEquipo.Name = "txtEquipo";
			this.txtEquipo.Size = new System.Drawing.Size(120, 21);
			this.txtEquipo.TabIndex = 340;
			this.txtEquipo.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtEquipo_KeyDown);
			// 
			// label18
			// 
			this.label18.AutoSize = true;
			this.label18.Location = new System.Drawing.Point(616, 34);
			this.label18.Name = "label18";
			this.label18.Size = new System.Drawing.Size(40, 16);
			this.label18.TabIndex = 341;
			this.label18.Text = "Equipo";
			// 
			// groupBox1
			// 
			this.groupBox1.Controls.Add(this.chkGRAut);
			this.groupBox1.Controls.Add(this.chkNAut);
			this.groupBox1.Controls.Add(this.cmbCategorias);
			this.groupBox1.Controls.Add(this.label17);
			this.groupBox1.Controls.Add(this.txtCreditoPequeño);
			this.groupBox1.Controls.Add(this.chkImprimePrecios);
			this.groupBox1.Controls.Add(this.chkCredito);
			this.groupBox1.Controls.Add(this.chkCarpas);
			this.groupBox1.Controls.Add(this.txtEquipo);
			this.groupBox1.Controls.Add(this.label18);
			this.groupBox1.Controls.Add(this.label12);
			this.groupBox1.Controls.Add(this.chkPrincipal);
			this.groupBox1.Controls.Add(this.chkPromociones);
			this.groupBox1.Controls.Add(this.chkPromocionesCombos);
			this.groupBox1.Controls.Add(this.cmbSectores);
			this.groupBox1.Controls.Add(this.label13);
			this.groupBox1.Controls.Add(this.chkFS);
			this.groupBox1.Controls.Add(this.txtSerie);
			this.groupBox1.Controls.Add(this.label14);
			this.groupBox1.Controls.Add(this.txtCreditoGrande);
			this.groupBox1.Controls.Add(this.label15);
			this.groupBox1.Controls.Add(this.PTC);
			this.groupBox1.Controls.Add(this.txtContado);
			this.groupBox1.Controls.Add(this.label16);
			this.groupBox1.Controls.Add(this.cmbTipoDeAlmacen);
			this.groupBox1.Controls.Add(this.txtCreditoMediano);
			this.groupBox1.Controls.Add(this.chkFactura);
			this.groupBox1.Location = new System.Drawing.Point(5, 272);
			this.groupBox1.Name = "groupBox1";
			this.groupBox1.Size = new System.Drawing.Size(904, 184);
			this.groupBox1.TabIndex = 342;
			this.groupBox1.TabStop = false;
			this.groupBox1.Text = "Factura";
			this.groupBox1.Enter += new System.EventHandler(this.groupBox1_Enter);
			// 
			// chkGRAut
			// 
			appearance41.ForeColorDisabled = System.Drawing.Color.Black;
			this.chkGRAut.Appearance = appearance41;
			this.chkGRAut.BackColor = System.Drawing.Color.Transparent;
			this.chkGRAut.Enabled = false;
			this.chkGRAut.Location = new System.Drawing.Point(672, 64);
			this.chkGRAut.Name = "chkGRAut";
			this.chkGRAut.Size = new System.Drawing.Size(120, 21);
			this.chkGRAut.TabIndex = 343;
			this.chkGRAut.Text = "GuiaR. Automatica";
			// 
			// chkNAut
			// 
			appearance42.ForeColorDisabled = System.Drawing.Color.Black;
			this.chkNAut.Appearance = appearance42;
			this.chkNAut.BackColor = System.Drawing.Color.Transparent;
			this.chkNAut.Enabled = false;
			this.chkNAut.Location = new System.Drawing.Point(552, 64);
			this.chkNAut.Name = "chkNAut";
			this.chkNAut.Size = new System.Drawing.Size(112, 21);
			this.chkNAut.TabIndex = 342;
			this.chkNAut.Text = "Num. Automatica";
			this.chkNAut.KeyDown += new System.Windows.Forms.KeyEventHandler(this.chkNAut_KeyDown);
			// 
			// groupBox2
			// 
			this.groupBox2.Controls.Add(this.chkBuffer2);
			this.groupBox2.Controls.Add(this.label4);
			this.groupBox2.Controls.Add(this.chkDestinoT);
			this.groupBox2.Controls.Add(this.txtOrdenSpf);
			this.groupBox2.Controls.Add(this.chkTransferencia);
			this.groupBox2.Controls.Add(this.chkNCP);
			this.groupBox2.Controls.Add(this.chkCompras);
			this.groupBox2.Controls.Add(this.chkTransferenciaMultiple);
			this.groupBox2.Controls.Add(this.chkDT);
			this.groupBox2.Controls.Add(this.chkCompra);
			this.groupBox2.Controls.Add(this.chkInventario);
			this.groupBox2.Location = new System.Drawing.Point(5, 152);
			this.groupBox2.Name = "groupBox2";
			this.groupBox2.Size = new System.Drawing.Size(904, 120);
			this.groupBox2.TabIndex = 343;
			this.groupBox2.TabStop = false;
			this.groupBox2.Text = "Inventario";
			this.groupBox2.Enter += new System.EventHandler(this.groupBox2_Enter);
			// 
			// chkBuffer2
			// 
			appearance43.ForeColorDisabled = System.Drawing.Color.Black;
			this.chkBuffer2.Appearance = appearance43;
			this.chkBuffer2.BackColor = System.Drawing.Color.Transparent;
			this.chkBuffer2.Enabled = false;
			this.chkBuffer2.Location = new System.Drawing.Point(256, 80);
			this.chkBuffer2.Name = "chkBuffer2";
			this.chkBuffer2.Size = new System.Drawing.Size(120, 21);
			this.chkBuffer2.TabIndex = 344;
			this.chkBuffer2.Text = "Visible en Buffer";
			// 
			// panel1
			// 
			this.panel1.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
			this.panel1.Controls.Add(this.label19);
			this.panel1.Controls.Add(this.cmbRegion);
			this.panel1.Controls.Add(this.label5);
			this.panel1.Controls.Add(this.dtFecha);
			this.panel1.Controls.Add(this.txtNumero);
			this.panel1.Controls.Add(this.label6);
			this.panel1.Controls.Add(this.dtcierre);
			this.panel1.Controls.Add(this.label2);
			this.panel1.Controls.Add(this.cmbCiudad);
			this.panel1.Controls.Add(this.label7);
			this.panel1.Controls.Add(this.txtcprovincia);
			this.panel1.Controls.Add(this.label8);
			this.panel1.Controls.Add(this.txtdireccion);
			this.panel1.Controls.Add(this.label9);
			this.panel1.Controls.Add(this.txtExtencion);
			this.panel1.Controls.Add(this.txtEMail);
			this.panel1.Controls.Add(this.label11);
			this.panel1.Controls.Add(this.label10);
			this.panel1.Controls.Add(this.cmbCentroDeCosto);
			this.panel1.Controls.Add(this.chkActivo);
			this.panel1.Controls.Add(this.chkCaja);
			this.panel1.Controls.Add(this.chkMatriz);
			this.panel1.Controls.Add(this.lblCliente);
			this.panel1.Controls.Add(this.txtCodigo);
			this.panel1.Controls.Add(this.label3);
			this.panel1.Controls.Add(this.txtNombre);
			this.panel1.Controls.Add(this.label1);
			this.panel1.Controls.Add(this.txtIdBodega);
			this.panel1.Location = new System.Drawing.Point(5, 8);
			this.panel1.Name = "panel1";
			this.panel1.Size = new System.Drawing.Size(904, 136);
			this.panel1.TabIndex = 344;
			this.panel1.Paint += new System.Windows.Forms.PaintEventHandler(this.panel1_Paint);
			// 
			// label19
			// 
			this.label19.AutoSize = true;
			this.label19.Location = new System.Drawing.Point(400, 104);
			this.label19.Name = "label19";
			this.label19.Size = new System.Drawing.Size(40, 16);
			this.label19.TabIndex = 317;
			this.label19.Text = "Region";
			this.label19.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// cmbRegion
			// 
			appearance44.ForeColorDisabled = System.Drawing.Color.Black;
			this.cmbRegion.Appearance = appearance44;
			this.cmbRegion.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbRegion.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbRegion.DataSource = this.ultraDataSource3;
			ultraGridColumn9.Header.VisiblePosition = 0;
			ultraGridColumn9.Hidden = true;
			ultraGridColumn10.Header.VisiblePosition = 1;
			ultraGridColumn10.Width = 140;
			ultraGridBand5.Columns.AddRange(new object[] {
																										 ultraGridColumn9,
																										 ultraGridColumn10});
			this.cmbRegion.DisplayLayout.BandsSerializer.Add(ultraGridBand5);
			this.cmbRegion.DisplayMember = "Region";
			this.cmbRegion.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbRegion.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbRegion.Enabled = false;
			this.cmbRegion.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbRegion.Location = new System.Drawing.Point(496, 104);
			this.cmbRegion.Name = "cmbRegion";
			this.cmbRegion.Size = new System.Drawing.Size(144, 21);
			this.cmbRegion.TabIndex = 316;
			this.cmbRegion.ValueMember = "idRegion";
			// 
			// ultraDataSource3
			// 
			ultraDataColumn6.DataType = typeof(int);
			this.ultraDataSource3.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn5,
																																 ultraDataColumn6});
			// 
			// frmBodegas
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 13);
			this.ClientSize = new System.Drawing.Size(914, 512);
			this.Controls.Add(this.panel1);
			this.Controls.Add(this.groupBox2);
			this.Controls.Add(this.groupBox1);
			this.Controls.Add(this.btnEditar);
			this.Controls.Add(this.btnGuardar);
			this.Controls.Add(this.btnCancelar);
			this.Controls.Add(this.btnSalir);
			this.Controls.Add(this.btnNuevo);
			this.Controls.Add(this.btnConsultar);
			this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog;
			this.KeyPreview = true;
			this.MaximizeBox = false;
			this.Name = "frmBodegas";
			this.ShowInTaskbar = false;
			this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
			this.Text = "Bodega";
			this.KeyDown += new System.Windows.Forms.KeyEventHandler(this.frmBodegas_KeyDown);
			this.Closing += new System.ComponentModel.CancelEventHandler(this.frmBodegas_Closing);
			this.SizeChanged += new System.EventHandler(this.frmBodegas_Load);
			this.Load += new System.EventHandler(this.frmBodegas_Load);
			((System.ComponentModel.ISupportInitialize)(this.txtIdBodega)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtCodigo)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNombre)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.dtFecha)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.dtcierre)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtOrdenSpf)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbCiudad)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtcprovincia)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtdireccion)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtExtencion)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtEMail)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbCentroDeCosto)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbCategorias)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource1)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbSectores)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource2)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtSerie)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtCreditoGrande)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtContado)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtCreditoMediano)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtCreditoPequeño)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNumero)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtEquipo)).EndInit();
			this.groupBox1.ResumeLayout(false);
			this.groupBox2.ResumeLayout(false);
			this.panel1.ResumeLayout(false);
			((System.ComponentModel.ISupportInitialize)(this.cmbRegion)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource3)).EndInit();
			this.ResumeLayout(false);

		}
		#endregion
		
		bool bNuevo = false;
		bool bEdicion = false;
		private Acceso miAcceso;

		private void UnloadMe()
		{
			this.Close();
		}

		private void frmBodegas_Load(object sender, System.EventArgs e)
		{		
			miAcceso = new Acceso(cdsSeteoF, "0308");
			
			if (!Funcion.Permiso("42", cdsSeteoF))
			{				
				MessageBox.Show("No puede ingresar a Bodegas", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				BeginInvoke(new MethodInvoker(UnloadMe));				
			}	

			if (miAcceso.Nuevo) this.btnNuevo.Enabled = true;
			if (miAcceso.Buscar) this.btnConsultar.Enabled = true;

			//carga de datos a los combobox ciudad 			
			this.cmbCiudad.DataSource = Funcion.dvProcedimiento(cdsSeteoF,"select idCiudad,Ciudad from ClienteCiudad order by Ciudad");
			//carga de datos a los combobox centro de costo 
			this.cmbCentroDeCosto.DataSource = Funcion.dvProcedimiento(cdsSeteoF,"select idProyecto,Nombre from Proyecto order by Nombre");			

			this.cmbSectores.DataSource = Funcion.dvProcedimiento(cdsSeteoF,"select idSector, Sector from ClienteSector");
			
			this.cmbCategorias.DataSource = Funcion.dvProcedimiento(cdsSeteoF,"select idTipoDeAlmacen, TipoDeAlmacen From TipoDeAlmacen Order By idTipoDeAlmacen");

			this.cmbRegion.DataSource = Funcion.dvProcedimiento(cdsSeteoF,"select idRegion, Region from Region ");
		}

		private void cdsSeteoF_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			cdsSeteoF.Schema.Connections[0].ConnectionString = miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);		
		}

		//BOTON NUEVO
		private void btnNuevo_Click(object sender, System.EventArgs e)
		{	
			this.btnCancelar_Click(sender, e);

			HabilitaChecksInv(this.chkInventario.Checked);
			HabilitarCheckFactura(this.chkFactura.Checked);
			
			int inferior = DateTime.Today.Date.Year;
			this.dtFecha.CalendarInfo.MinDate = new DateTime (inferior,1,1);
			this.dtFecha.CalendarInfo.MaxDate = DateTime.Today;
			this.dtFecha.Value = DateTime.Today;
		
			//SCRIPTS PARA RELACIONAR REGISTROS CON OBJETOS
			this.txtNumero.Value=Funcion.iEscalarSQL(cdsSeteoF,"select top 1 Bodega + 1 from Bodega order by Bodega desc");
			this.txtOrdenSpf.Value=Funcion.iEscalarSQL(cdsSeteoF,"select top 1 OrdenTM + 1 from Bodega order by OrdenTM desc");
			this.dtFecha.Value = DateTime.Today;

			//HABILITACION DE LOS OBJETOS
			this.chkFactura.Enabled= true;
			this.txtNombre.Enabled= true;
			this.dtFecha.Enabled = true;
			this.chkActivo.Checked = true;
			//	this.txtcprovincia.Enabled = true;
			this.txtExtencion.Enabled = true;
			this.txtEMail.Enabled = true;
			this.txtCodigo.Enabled = true;
			if (MenuLatinium.stIdDB != 1 )this.cmbCentroDeCosto.Enabled = true;
			this.cmbRegion.Enabled = true;
			this.cmbCiudad.Enabled = true;
			this.txtdireccion.Enabled=true;
			//ESTABLECER LONGITUD
			this.txtExtencion.MaxLength = 4;			
			
			// habilitacion de checkbox		
			this.chkMatriz.Enabled = true;
			this.chkCaja.Enabled = true;
			this.chkInventario.Enabled = true;
		    
			this.btnNuevo.Enabled = false;
			this.btnConsultar.Enabled = false;
			this.btnEditar.Enabled = false;
			this.btnGuardar.Enabled = true;
			this.btnCancelar.Enabled = true;		
			
			//vaeriables de validacion inicualizadas
			bNuevo = true;
			bEdicion = true;
			//salto de objeto a
			
			this.txtCodigo.Focus();			
		}
		
		// obtencion del codigo de provincia para el texbox codigo provincia
		private void cmbCiudad_ValueChanged(object sender, System.EventArgs e)
		{
			if (bEdicion)
				if(this.cmbCiudad.ActiveRow !=null)
					this.txtcprovincia.Value= Funcion.sEscalarSQL(cdsSeteoF, string.Format("select isnull (Codigo ,'') from ClienteCiudad where idCiudad = {0}", (int)cmbCiudad.Value));
			
		}

		// salto de objeto de codigo a nombre
		private void txtCodigo_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if(e.KeyCode == Keys.Enter) this.txtNombre.Focus();
		}
		
		private void frmBodegas_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (!bEdicion)
			{
				if (e.KeyCode == Keys.N) this.btnNuevo_Click(sender, e);
				if (e.KeyCode == Keys.B) this.btnConsultar_Click(sender, e);
				if ((int)this.txtIdBodega.Value > 0)
				{
					if (e.KeyCode == Keys.E) if (this.btnEditar.Enabled) this.btnEditar_Click(sender, e);					
					if (e.KeyCode == Keys.F5) this.Consultar((int)this.txtIdBodega.Value);					
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

		private void ultraNumericEditor1_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if(e.KeyCode == Keys.Enter) this.chkFactura.Focus();
		}

		private void txtNombre_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if(e.KeyCode == Keys.Enter) this.dtFecha.Focus();
		}

		private void txtTelefono_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if(e.KeyCode == Keys.Enter) this.txtEMail.Focus();
		}
		private void AgregaBodega()
		{
			string sSQL = string.Format("Exec AgregaBodegasBuffer {0},'{1}'",(int)this.txtNumero.Value,this.txtCodigo.Text.ToString());
			Funcion.EjecutaSQL(cdsSeteoF, sSQL);
		}

		private void InsertaAtributos()
		{
			string sSQL = string.Format("Exec InsertaBodegasCaterorias {0}",(int)this.txtNumero.Value);
			Funcion.EjecutaSQL(cdsSeteoF, sSQL);
		}
		private void btnGuardar_Click(object sender, System.EventArgs e)
		{
			#region Validacion
			#region Datos Generales
			if (!Validacion.vbTexto(this.txtCodigo, 2, 3, "Codigo")) return;
			
			if (Funcion.bEscalarSQL(cdsSeteoF, string.Format("Exec ValidaBodegaExistente {0}, '{1}'", (int)this.txtIdBodega.Value, this.txtCodigo.Text.ToString()), true))
			{
				MessageBox.Show("Codigo de bodega existente", "Point Tehnology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				this.txtCodigo.Focus();
				return;
			}

			if (!Validacion.vbTexto(this.txtNombre, 5, 60, "Nombre Bodega")) return;
      
			if (!Validacion.vbFechaEnDocumentos(this.dtFecha, "Seleccione la fecha de apertura", false, "", cdsSeteoF)) return;

			if (this.dtcierre.Value != System.DBNull.Value)
				if(!Validacion.vbComparaFechas((DateTime)this.dtFecha.Value,(DateTime)this.dtcierre.Value,this.dtFecha,"la fecha de cierre no puedes ser mayor al dia actual ni menor a la fecha ingreso",true))return;

			if (!Validacion.vbComboVacio(this.cmbCiudad, "Seleccione la ciudad de ubicacion de la bodega")) return;

			if (!Validacion.vbTexto(this.txtdireccion, 10, 200, "Direccion")) return;

			if (!Validacion.vbTexto(this.txtExtencion, 3, 4, "Extensión")) return;

			if (!Validacion.vbTexto(this.txtEMail, 0, 60, "E-Mail")) return;
			
			if (!Validacion.bEMail(this.txtEMail.Text.ToString())) return;
			if (MenuLatinium.stIdDB != 1 )
			{
				if (!Validacion.vbComboVacio(this.cmbCentroDeCosto, "Seleccione el centro de costo asociado a esta bodega")) return;
			}
			#endregion Datos Generales

			#region Validacion si la Bodega Factura
			if (this.chkFactura.Checked)
			{
				if (!Validacion.vbComboVacio(this.cmbCategorias, "Seleccione la categoria del Almacen")) return;

				if (!Validacion.vbComboVacio(this.cmbSectores, "Seleccione el sector de ubicacion del Almacen")) return;

				if (!Validacion.vbTexto(this.txtSerie, 6, 6, "Serie")) return;
							
				int iRes = Funcion.iEscalarSQL(cdsSeteoF, string.Format("Exec ValidaSerieBodega '{0}', {1}, {2}", 
					this.txtSerie.Text.ToString(), (int)this.txtNumero.Value, (bool)this.chkPrincipal.Checked));
				if (iRes == 2)
				{
					MessageBox.Show("El numero de Serie NO Existe", "Point Tehnology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
					this.txtSerie.Focus();
					return;
				}
				if (iRes == 1)
				{
					MessageBox.Show("La serie esta registrada con otro almacen", "Point Tehnology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
					this.txtSerie.Focus();
					return;
				}

				if (!Validacion.vbTexto(this.txtEquipo, 10, 20, "Equipo")) return;
			}
			#endregion Validacion si la Bodega Factura

			#region Valida al cerrar bodega si tiene inventario o no
			if (this.dtcierre.Value != System.DBNull.Value)				
			{
				string sSQLVI = string.Format("select dbo.CalculaExistenciasInventario('{0}', 0, {1})",
					Convert.ToDateTime(this.dtcierre.Value).ToString("yyyyMMdd"), (int)this.txtNumero.Value );
				int iInventario = Funcion.iEscalarSQL(cdsSeteoF, sSQLVI);

				if(iInventario > 0)
				{
					MessageBox.Show("La bodega aun tiene inventario no puede cerrar", "Point Tehnology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
					return;
				}
			}
			#endregion Valida al cerrar bodega si tiene inventario o no
			#endregion Validacion
			
			int Tipoalmacen = 0;
			int Idsector = 0;
			int IdCentroDeCosto = 0;
			if ( this.cmbCategorias.ActiveRow != null) Tipoalmacen  = ( int) this.cmbCategorias.Value;
			if ( this.cmbSectores.ActiveRow != null) Idsector  = ( int) this.cmbSectores.Value;
			if ( this.cmbCentroDeCosto.ActiveRow != null) IdCentroDeCosto  = ( int) this.cmbCentroDeCosto.Value;

			DateTime dtfechacierre = DateTime.Parse("01/01/2000");
			if (this.dtcierre.Value != System.DBNull.Value)dtfechacierre=(DateTime)this.dtcierre.Value;
			#region centro de costo
			if (MenuLatinium.stIdDB == 1 )
			{
				string sSQLI = string.Format("Exec InsertOrUpdateProyecto {0},'{1}','{2}'",
					IdCentroDeCosto,  this.txtNombre.Text.ToString(), 
					Convert.ToDateTime(this.dtFecha.Value).ToString("yyyyMMdd")					
					);
				IdCentroDeCosto=Funcion.iEscalarSQL(cdsSeteoF,sSQLI);
			}
			#endregion centro de costo
			
			#region Guarda Bodega
			string sSQL = string.Format("Exec GuardaBodega {0},{1},'{2}','{3}','{4}','{5}','{6}',{7},'{8}','{9}','{10}',{11},{12},{13},{14},{15},{16},{17},{18},{19},{20},{21},{22},{23},{24},{25},{26},{27},'{28}',{29},{30},{31},{32},{33},'{34}',{35},{36},'{37}','{38}','{39}','{40}',{41}, {42}, {43},{44}",
				(int)this.txtIdBodega.Value, (int)this.txtNumero.Value, this.txtCodigo.Text.ToString(), this.txtNombre.Text.ToString(), Convert.ToDateTime(this.dtFecha.Value).ToString("yyyyMMdd"), dtfechacierre.ToString("yyyyMMdd"),					

				this.txtcprovincia.Text.ToString(), (int)this.cmbCiudad.Value, this.txtdireccion.Text.ToString(), this.txtExtencion.Text.ToString(), this.txtEMail.Text.ToString(),

				IdCentroDeCosto,

				(bool)this.chkActivo.Checked, (bool)this.chkMatriz.Checked, (bool)this.chkCaja.Checked, 

				(bool)this.chkInventario.Checked, (bool)this.chkTransferencia.Checked, (bool)this.chkTransferenciaMultiple.Checked, (bool)this.chkCompra.Checked, (bool)this.chkNCP.Checked, 
				(bool)this.chkDT.Checked, (bool)this.chkDestinoT.Checked, (bool)this.chkCompras.Checked, (int) this.txtOrdenSpf.Value,

				(bool)this.chkFactura.Checked, Tipoalmacen, Idsector,
				(bool)this.chkPrincipal.Checked, this.txtSerie.Text.ToString(), (bool)this.chkCredito.Checked, (bool)this.chkFS.Checked,
				(bool)this.chkPromociones.Checked, (bool)this.chkPromocionesCombos.Checked, (bool) this.chkCarpas.Checked, this.txtEquipo.Text.ToString(), 
				
				(bool) this.chkImprimePrecios.Checked, (bool) this.PTC.Checked, 
				this.txtCreditoGrande.Text.ToString(), this.txtCreditoMediano.Text.ToString(), this.txtCreditoPequeño.Text.ToString(), this.txtContado.Text.ToString(), (bool)this.chkNAut.Checked,
				this.cmbRegion.Value, (bool) this.chkGRAut.Checked, (bool)this.chkBuffer2.Checked);
			this.txtIdBodega.Value=Funcion.iEscalarSQL(cdsSeteoF,sSQL);

			if( this.chkBuffer2.Checked)
			{
				AgregaBodega();
			}
			if ( this.chkBuffer2.Checked)
			{
				InsertaAtributos();		
			}
			
			//	if( this.chkBuffer2.Checked)
			//{
			
			//	}
		
			#endregion Guarda Bodega

			#region Controles
			this.cmbCentroDeCosto.DataSource = Funcion.dvProcedimiento(cdsSeteoF,"select idProyecto,Nombre from Proyecto order by Nombre");			
			this.Consultar((int)this.txtIdBodega.Value);
			Controles();

	
			#endregion Controles			
		}

		private void Controles()
		{
			#region Controles
			this.btnNuevo.Enabled= true;
			this.txtIdBodega.Enabled=false;
			this.txtNumero.Enabled=false;
			this.txtCodigo.Enabled=false;
			this.txtNombre.Enabled=false;
			this.dtFecha.Enabled=false;
			this.dtFecha.Enabled=false;
			this.dtcierre.Enabled=false;
			this.txtcprovincia.Enabled=false;
			this.cmbCiudad.Enabled=false;
			this.cmbSectores.Enabled=false;
			this.cmbCategorias.Enabled=false;
			this.dtFecha.Enabled=false;
			this.txtdireccion.Enabled=false;
			this.txtExtencion.Enabled=false;
			this.txtEMail.Enabled=false;
			this.cmbCentroDeCosto.Enabled=false;
			this.cmbRegion.Enabled=false;
			this.chkMatriz.Enabled=false;
			this.chkCaja.Enabled=false;
			this.chkInventario.Enabled=false;
			this.chkTransferencia.Enabled=false;
			this.chkTransferenciaMultiple.Enabled=false;
			this.chkCompra.Enabled=false;
			this.chkNCP.Enabled=false;
			this.chkDT.Enabled=false;
			this.chkCompras.Enabled=false;
			this.chkFactura.Enabled=false;
			this.chkPrincipal.Enabled=false;
			this.txtSerie.Enabled=false;
			this.chkCredito.Enabled=false;
			this.chkFS.Enabled=false;
			this.chkPromociones.Enabled=false;
			this.chkPromocionesCombos.Enabled=false;
			this.chkImprimePrecios.Enabled=false;
			this.chkDestinoT.Enabled=false;
			this.PTC.Enabled=false;
			this.chkCarpas.Enabled=false;
			this.txtEquipo.Enabled=false;
			this.chkNAut.Enabled = false;
			this.chkBuffer2.Enabled=false;

			if (miAcceso.Nuevo) this.btnNuevo.Enabled = true;
			if (miAcceso.Buscar) this.btnConsultar.Enabled = true;
			if (miAcceso.Editar) this.btnEditar.Enabled = true;
			this.btnGuardar.Enabled = false;
			this.btnCancelar.Enabled = true;
			
			bNuevo = false;
			bEdicion = false;

			#endregion Controles	
		
		}

		private void txtExtencion_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
			validar.Solonumeros(e);
		}

		// salto de ciudad a direccion
		private void cmbCiudad_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if(e.KeyCode == Keys.Enter) this.txtdireccion.Focus();
		
		}
		// salto de email ac entro de costo
		private void txtEMail_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if(e.KeyCode == Keys.Enter) this.cmbCentroDeCosto.Focus();
	
		}
		// salto de fecha a ciudad
		private void dtFecha_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if(e.KeyCode == Keys.Enter) this.cmbCiudad.Focus();	
		}

		// salto de direccion a extension

		private void txtdireccion_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if(e.KeyCode == Keys.Enter) this.txtExtencion.Focus();	
		}
		
		//salto de compras a orden spiff
		private void chkCompras_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if(e.KeyCode == Keys.Enter) this.chkFactura.Focus();	
		}

		private void chkFactura_CheckedChanged(object sender, System.EventArgs e)
		{
			
			if (!bEdicion) return;
			HabilitarCheckFactura(this.chkFactura.Checked);
		}

		private void chkCompras_CheckedChanged(object sender, System.EventArgs e)
		{
			//HABILITACION Y DESHABILITACION DE OBJETOS CHECKBOX Y TEXTBOX
			if( this.chkCompras.Checked)
			{				
				this.txtOrdenSpf.Visible = true;				
			}
			else
			{							
				this.txtOrdenSpf.Visible = false;				
			}
		}

		private void txtSerie_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
			validar.Solonumeros(e);
		}

		// habilitacion y deshabiltacion de objetos chekbox
		private void chkPrincipal_Click(object sender, System.EventArgs e)
		{
			if( this.chkPrincipal.Checked)
			{
				this.txtSerie.Text="";		
			}
			else
			{		
				this.chkTransferencia.Enabled = true;
				this.chkTransferenciaMultiple.Enabled = true;
				this.chkNCP.Enabled = true;
				this.chkDT.Enabled = true;
				this.chkCompra.Enabled = true ;
				this.chkBuffer2.Enabled = true ;
				this.chkCompras.Enabled = true ;
			}			
		}

		private void txtSerie_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if(e.KeyCode == Keys.Enter) this.txtEquipo.Focus();
		}
		
		// salto  de centro de costo a chkmatriz
		private void cmbCentroDeCosto_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if(e.KeyCode == Keys.Enter) this.chkMatriz.Focus();
		}
		
		// salto de  de factufa a ulktacombo
		private void chkFactura_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if(this.chkFactura.Checked)
			{				
				if(e.KeyCode == Keys.Enter)this.cmbCategorias.Focus();
			}
			else
			{
				if(e.KeyCode == Keys.Enter) this.btnGuardar.Focus();
			}			
		}

		// salto de secttores a principsl
		private void cmbSectores_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if(e.KeyCode == Keys.Enter) this.chkCredito.Focus();
		}

		//salto de principla  a serie
		private void chkPrincipal_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if(e.KeyCode == Keys.Enter) this.txtSerie.Focus();
		}
		
		//SALTO DE MATRIZ A CAJA
		private void chkMatriz_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if(e.KeyCode == Keys.Enter) this.chkCaja.Focus();
		}
		// SALTO DE CAJA A INVENTARIO
		private void chkCaja_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if(e.KeyCode == Keys.Enter) this.chkInventario.Focus();
		}

		private void chkInventario_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			//key down
			
			if(this.chkInventario.Checked)
			{
				if(e.KeyCode == Keys.Enter) this.chkTransferencia.Focus();
			}
			else
			{
				if(e.KeyCode == Keys.Enter)this.chkFactura.Focus();
			}			
		}

		private void chkTransferencia_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if(e.KeyCode == Keys.Enter)this.chkTransferenciaMultiple.Focus();
		}

		private void chkTransferenciaMultiple_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if(e.KeyCode == Keys.Enter)this.chkCompra.Focus();
		}

		private void chkCompra_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if(e.KeyCode == Keys.Enter)this.chkNCP.Focus();
		}

		private void chkNCP_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if(e.KeyCode == Keys.Enter)this.chkDT.Focus();
		}

		private void chkDT_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if(e.KeyCode == Keys.Enter)this.chkDestinoT.Focus();
		}

		private void chkCredito_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if(e.KeyCode == Keys.Enter) this.chkPrincipal.Focus();
		}

		private void chkFS_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if(e.KeyCode == Keys.Enter) this.chkNAut.Focus();
		}

		private void chkPromociones_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if(e.KeyCode == Keys.Enter) this.chkCarpas.Focus();
		}

		private void chkPromocionesCombos_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if(e.KeyCode == Keys.Enter) this.btnGuardar.Focus();
		}

		private void chkImprimePrecios_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if(e.KeyCode == Keys.Enter) this.PTC.Focus();
		}

		private void btnSalir_Click(object sender, System.EventArgs e)
		{
			this.Close();			
		}
				
		// BOTON CANCELAR
		private void btnCancelar_Click(object sender, System.EventArgs e)
		{		  
			bNuevo = false;
			bEdicion = false;

			this.txtIdBodega.Value=0;
			this.txtOrdenSpf.Value=0;
			this.txtNumero.Value = 0;
			this.txtNumero.Enabled=false;
			this.txtCodigo.Text= "";
			this.txtCodigo.Enabled=false;
			this.txtNombre.Text= "";
			this.txtNombre.Enabled=false;
			this.dtFecha.Value = System.DBNull.Value;
			this.dtFecha.Enabled=false;			
			this.dtcierre.Value = System.DBNull.Value;
			this.dtcierre.Enabled=false;
			this.txtcprovincia.Text="";
			this.txtcprovincia.Enabled=false;
			this.cmbCiudad.Value = System.DBNull.Value;
			this.cmbCiudad.Enabled=false;
			this.cmbSectores.Value = System.DBNull.Value;
			this.cmbSectores.Enabled=false;
			this.cmbCategorias.Value = System.DBNull.Value;
			this.cmbCategorias.Enabled=false;			
			this.txtdireccion.Text= "";
			this.txtdireccion.Enabled=false;
			this.cmbRegion.Value = System.DBNull.Value;
			this.cmbRegion.Enabled=false;	

			this.txtExtencion.Text= "";
			this.txtExtencion.Enabled=false;
			this.txtEMail.Text= "";
			this.txtEMail.Enabled=false;
			this.cmbCentroDeCosto.Value = System.DBNull.Value;
			this.cmbCentroDeCosto.Enabled=false;
			this.chkMatriz.Enabled=false;
			this.chkMatriz.Checked=false;
			this.chkCaja.Enabled=false;
			this.chkCaja.Checked=false;
			this.chkInventario.Enabled=false;
			this.chkInventario.Checked=false;

			this.chkTransferencia.Enabled=false;
			this.chkTransferencia.Checked=false;

			this.chkTransferenciaMultiple.Enabled=false;
			this.chkTransferenciaMultiple.Checked=false;

			this.chkCompra.Enabled=false;
			this.chkCompra.Checked=false;
			
			this.chkBuffer2.Enabled=false;

			this.chkNCP.Enabled=false;
			this.chkNCP.Checked=false;
			this.chkDT.Enabled=false;
			this.chkDT.Checked=false;
			this.chkCompras.Enabled=false;
			this.chkCompras.Checked=false;


			this.chkFactura.Enabled=false;
			this.chkFactura.Checked=false;


			this.chkPrincipal.Enabled=false;
			this.chkPrincipal.Checked=false;


			this.txtSerie.Enabled=false;
			
			this.txtSerie.Text="";

			this.chkCredito.Enabled=false;
			this.chkCredito.Checked=false;

			this.chkFS.Enabled=false;
			this.chkFS.Checked=false;

			this.chkPromociones.Enabled=false;
			this.chkPromociones.Checked=false;
			this.chkPromocionesCombos.Enabled=false;
			this.chkPromocionesCombos.Checked=false;

			this.chkImprimePrecios.Enabled=false;
			this.chkImprimePrecios.Checked=false;

			this.chkDestinoT.Enabled=false;
			this.chkDestinoT.Checked=false;

			this.chkNAut.Enabled=false;
			this.chkNAut.Checked=false;

			this.chkGRAut.Enabled=false;
			this.chkGRAut.Checked=false;

			this.PTC.Enabled=false;
			this.PTC.Checked=false;
			this.chkActivo.Checked=false;
			this.txtCreditoGrande.Text="";
			this.txtCreditoMediano.Text="";
			this.txtCreditoPequeño.Text="";
			this.txtContado.Text="";

			this.chkCarpas.Enabled=false;
			this.txtEquipo.Enabled=false;

			this.chkCarpas.Checked=false;
			this.txtEquipo.Text="";



			this.btnNuevo.Enabled = true;
			this.btnConsultar.Enabled = true;
			this.btnEditar.Enabled = false;
			this.btnGuardar.Enabled = false;
			this.btnCancelar.Enabled = false;
			
		}	

		private void HabilitaChecksInv (bool bEstado)
		{
			this.chkTransferencia.Enabled = bEstado;
			this.chkTransferenciaMultiple.Enabled = bEstado;
			this.chkNCP.Enabled = bEstado;
			this.chkDT.Enabled = bEstado;
			this.chkBuffer2.Enabled = bEstado;
			this.chkCompra.Enabled = bEstado ;
			this.chkCompras.Enabled = bEstado ;
			this.chkDestinoT.Enabled = bEstado ;
		
			if (!bEstado)
			{
				this.chkCompra.Checked = false;
				this.chkTransferencia.Checked = false;
				this.chkTransferenciaMultiple.Checked = false;
				this.chkNCP.Checked = false;
				this.chkCompras.Checked = false;
				this.chkDT.Checked = false;
				this.chkBuffer2.Checked = false;
				this.chkDestinoT.Checked= false; 
				this.txtOrdenSpf.Visible = false;
				this.txtOrdenSpf.Enabled = false;				
			}
		}

		private void HabilitarCheckFactura(bool bEstado)
		{							
			this.chkPrincipal.Enabled= bEstado;
			this.cmbSectores.Enabled=bEstado;
			this.chkPromocionesCombos.Enabled=bEstado;
				
			this.cmbCategorias.Enabled = bEstado;
			this.txtSerie.Enabled=bEstado;
			this.chkCredito.Enabled=bEstado;
				
			this.chkFS.Enabled=bEstado;
			
			this.chkPromociones.Enabled=bEstado;
		
			this.PTC.Enabled= bEstado;
			this.chkImprimePrecios.Enabled=bEstado;

			this.chkCarpas.Enabled=bEstado;
			this.txtEquipo.Enabled=bEstado;

			this.chkNAut.Enabled = bEstado;
			this.chkGRAut.Enabled = false;
			this.chkGRAut.Checked = true;
			
			if(!bEstado)
			{
				this.chkPrincipal.Enabled= false;
				this.chkPrincipal.Checked= false;
				this.chkPromocionesCombos.Enabled=false;
				this.chkPromocionesCombos.Checked=false;
				this.cmbCategorias.Enabled = false;
				this.txtSerie.Enabled=false;
				this.chkCredito.Enabled=false;
				this.cmbSectores.Enabled=false;
				this.chkCredito.Checked=false;
				this.chkFS.Enabled=false;
				this.chkFS.Checked=false;
				this.chkPromociones.Enabled=false;
				this.chkPromociones.Checked=false;
				this.chkImprimePrecios.Checked=false;
				this.chkImprimePrecios.Enabled=false;
				this.chkNAut.Enabled= false;
				this.chkNAut.Checked = false;
				this.chkGRAut.Enabled= false;
				this.chkGRAut.Checked = false;
				
				this.PTC.Enabled=false;
				this.PTC.Checked = false;
				this.cmbSectores.Text= "";
				this.txtSerie.Text="";
			
				this.cmbCategorias.Text = "";

				this.chkCarpas.Enabled=false;
				this.txtEquipo.Enabled=false;

				this.chkCarpas.Checked=false;
				this.txtEquipo.Text="";			
			}			
		}


		private void chkInventario_CheckedChanged(object sender, System.EventArgs e)
		{
			if (!bEdicion) return;
			HabilitaChecksInv(this.chkInventario.Checked);						
		}

		private void chkImprimePrecios_CheckedChanged(object sender, System.EventArgs e)
		{
			if (!bEdicion) return;
			if( this.chkImprimePrecios.Checked)
			{
				this.txtCreditoGrande.Value = @"C:\WINDOWS\programas\preciadores\CFPCG.lnk";
				this.txtCreditoMediano.Value = @"C:\WINDOWS\programas\preciadores\CFPCM.lnk";
				this.txtCreditoPequeño.Value =@"C:\WINDOWS\programas\preciadores\CFPCP.lnk";
				this.txtContado.Value=@"C:\WINDOWS\programas\preciadores\CFPC.lnk";
				//			this.txtCreditoMediano.Value=Funcion.dvProcedimiento(cdsSeteoF,"select top 1 DirectorioPC from Bodega Order By DirectorioPC desc ");
				//				this.txtCreditoPequeño.Value=Funcion.dvProcedimiento(cdsSeteoF,"select top 1 DirectorioPCP from Bodega Order By DirectorioPCP desc");
				//			this.txtContado.Value=Funcion.dvProcedimiento(cdsSeteoF,"select top 1 DirectorioPCM from Bodega  Order By DirectorioPCM desc ");
				//	
				
			}
			else
			{
				// limpiar valores
				this.txtCreditoGrande.Value="";
				this.txtCreditoMediano.Value="";
				this.txtCreditoPequeño.Value="";
				this.txtContado.Value="";
			}
		}

		private void chkPrincipal_CheckedChanged(object sender, System.EventArgs e)
		{
			if( this.chkPrincipal.Checked)
			{
				this.txtSerie.Text="";
			}
			else 
			{
				this.txtSerie.Focus();
			}
		}

		private void chkDestinoT_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if(e.KeyCode == Keys.Enter)this.chkCompras.Focus();
		}

		private void chkDestinoT_KeyDown_1(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if(e.KeyCode == Keys.Enter)this.chkCompras.Focus();
		}

		private void txtEquipo_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if(e.KeyCode == Keys.Enter)this.chkFS.Focus();
		}

		private void chkCarpas_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if(e.KeyCode == Keys.Enter)this.chkImprimePrecios.Focus();
		}

		private void PTC_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{		
			if(e.KeyCode == Keys.Enter)this.chkPromocionesCombos.Focus();
		}

		private void Consultar (int IdBodega )
		{
			#region Consulta Individual
			string ssql= string.Format ("Exec ConsultaIndividualBodega {0}",IdBodega); 
			SqlDataReader dr = Funcion.rEscalarSQL(cdsSeteoF,ssql,true);
			dr.Read();
			this.txtIdBodega.Value= dr.GetInt32 (0);
			this.txtNumero.Value = dr.GetInt32 (1);
			this.txtCodigo.Text = dr.GetString (2);
			this.txtNombre.Text = dr.GetString (3);
			this.dtFecha.Value= dr.GetDateTime (4);
			if (dr.GetValue(5) != System.DBNull.Value) this.dtcierre.Value = dr.GetDateTime (5);
			if (dr.GetValue(6) != System.DBNull.Value)this.txtcprovincia.Text = dr.GetString (6);
			this.cmbCiudad.Value = dr.GetInt32(7);
			if (dr.GetValue(8) != System.DBNull.Value)this.txtdireccion.Text=dr.GetString(8); 
			if (dr.GetValue(9) != System.DBNull.Value)this.txtExtencion.Text= dr.GetString(9);
			if (dr.GetValue(10) != System.DBNull.Value)this.txtEMail.Text = dr.GetString (10);
			if (dr.GetValue(11) != System.DBNull.Value)this.cmbCentroDeCosto.Value = dr.GetInt32(11);
			this.chkActivo.Checked = dr.GetBoolean(12);
			this.chkMatriz.Checked = dr.GetBoolean(13);
			this.chkCaja.Checked= dr.GetBoolean(14);
			this.chkInventario.Checked= dr.GetBoolean(15);
			this.chkTransferencia.Checked= dr.GetBoolean(16);
			this.chkTransferenciaMultiple.Checked= dr.GetBoolean(17);
			this.chkCompra.Checked= dr.GetBoolean(18);
			this.chkNCP.Checked= dr.GetBoolean(19);
			this.chkDT.Checked= dr.GetBoolean(20);
			this.chkDestinoT.Checked= dr.GetBoolean(21);
			this.chkCompras.Checked= dr.GetBoolean(22);
			this.txtOrdenSpf.Value= dr.GetInt32 (23);
			this.chkFactura.Checked= dr.GetBoolean(24);
			this.cmbCategorias.Value= dr.GetInt32 (25);
			this.cmbSectores.Value= dr.GetInt32 (26);	
			this.chkCredito.Checked= dr.GetBoolean(27);
			this.chkPrincipal.Checked= dr.GetBoolean(28);
			if (dr.GetValue(29) != System.DBNull.Value)this.txtSerie.Text = dr.GetString(29);
			if (dr.GetValue(30) != System.DBNull.Value)this.txtEquipo.Text = dr.GetString (30);
			this.chkFS.Checked= dr.GetBoolean(31);
			this.chkPromociones.Checked= dr.GetBoolean(32);
			this.chkCarpas.Checked= dr.GetBoolean(33);
			this.chkImprimePrecios.Checked= dr.GetBoolean(34);
			this.PTC.Checked= dr.GetBoolean(35);
			this.chkPromocionesCombos.Checked= dr.GetBoolean(36);
			this.txtCreditoGrande.Text = dr.GetString (37);
			this.txtCreditoMediano.Text = dr.GetString (38);
			this.txtCreditoPequeño.Text = dr.GetString (39);
			if (dr.GetValue(40) != System.DBNull.Value)this.txtContado.Text = dr.GetString (40);
			if (dr.GetValue(41) != System.DBNull.Value)this.chkNAut.Checked= dr.GetBoolean(41);
			if (dr.GetValue(42) != System.DBNull.Value)this.cmbRegion.Value = dr.GetInt32(42);
			if (dr.GetValue(43) != System.DBNull.Value)	this.chkGRAut.Checked= dr.GetBoolean(43);
			if (dr.GetValue(44) != System.DBNull.Value)this.chkBuffer2.Checked= dr.GetBoolean(44);
			dr.Close();
			#endregion Consulta Individual

			this.btnNuevo.Enabled = true;
			this.btnConsultar.Enabled = true;
			this.btnEditar.Enabled = true;
			this.btnGuardar.Enabled = false;
			this.btnCancelar.Enabled = true;
		}

		private void btnConsultar_Click(object sender, System.EventArgs e)
		{		
			this.btnCancelar_Click(sender, e);

			using (frmConsultaBodega consultab = new frmConsultaBodega())
			{
				if (DialogResult.OK == consultab.ShowDialog())
				{
					this.Consultar((int)consultab.uGrid.ActiveRow.Cells["IdBodega"].Value);
				}				
			}		
		}

		private void btnEditar_Click(object sender, System.EventArgs e)
		{
			// habilitar del formulario
			
			if (!miAcceso.Editar)

			{
				MessageBox.Show("No tiene Acceso a Editar", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				Controles();
			}
			else
			{
				this.dtcierre.CalendarInfo.MaxDate=DateTime.Today;
				this.dtcierre.CalendarInfo.MinDate= (DateTime)this.dtFecha.Value;

				this.txtNombre.Enabled= true;
				this.dtcierre.Enabled = true;
				this.txtExtencion.Enabled = true;
				this.txtEMail.Enabled = true;
				this.cmbCiudad.Enabled = true;
				this.cmbRegion.Enabled = true;
				this.txtdireccion.Enabled=true;
				if (MenuLatinium.stIdDB != 1 )this.cmbCentroDeCosto.Enabled = true;
		
				this.chkMatriz.Enabled = true;
				this.chkCaja.Enabled = true;
				this.chkInventario.Enabled = true;
				this.chkFactura.Enabled = true;
		
				this.btnNuevo.Enabled = false;
				this.btnConsultar.Enabled = false;
				this.btnEditar.Enabled = false;
				this.btnGuardar.Enabled = true;
				this.btnCancelar.Enabled = true;
				HabilitaChecksInv(this.chkInventario.Checked);
				HabilitarCheckFactura(this.chkFactura.Checked);
				//		  this.chkInventario.CheckedChanged(sender, e);
				bEdicion=true;
				this.txtNombre.Focus();

			}


	
		}

		private void cmbCategorias_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Enter) this.cmbSectores.Focus();
		}

		private void frmBodegas_Closing(object sender, System.ComponentModel.CancelEventArgs e)
		{
			if (bEdicion)
			{
				MessageBox.Show("No puede salir mientras edita un registro", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				e.Cancel = true;
				return;
			}
		}

		private void chkNAut_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if(e.KeyCode == Keys.Enter) this.chkPromociones.Focus();
		}

		private void panel1_Paint(object sender, System.Windows.Forms.PaintEventArgs e)
		{
		
		}

		private void groupBox1_Enter(object sender, System.EventArgs e)
		{
		
		}

		private void groupBox2_Enter(object sender, System.EventArgs e)
		{
		
		}

		private void chkCompra_CheckedChanged(object sender, System.EventArgs e)
		{
		
		}

		private void ultraCheckEditor1_CheckedChanged(object sender, System.EventArgs e)
		{
		
		}

		private void chkDT_CheckedChanged(object sender, System.EventArgs e)
		{
		
		}
	}
	// activacion
		
	// validamos solo numeros en extencion
	class validar 
	{
		public static void Solonumeros (KeyPressEventArgs Nu)
		{
			if(char.IsDigit(Nu.KeyChar))
			{
			
				Nu.Handled = false;
			
			}
			else if (char.IsControl(Nu.KeyChar))
			{
			
				Nu.Handled = false;			} 
			else
			{
				Nu.Handled= true;
			}
		}
	}	
}

// cerramos validacion
