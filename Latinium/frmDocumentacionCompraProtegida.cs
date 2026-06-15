using System;
using System.IO;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using System.Data;
using System.Data.SqlClient;

namespace Latinium
{
	/// <summary>
	/// Descripción breve de frmDocumentacionCompraProtegida.
	/// </summary>
	public class frmDocumentacionCompraProtegida : System.Windows.Forms.Form
	{
		public Infragistics.Win.UltraWinEditors.UltraCheckEditor chkFormulario;
		/// <summary>
		/// Variable del diseñador requerida.
		/// </summary>
		private System.ComponentModel.Container components = null;
		public Infragistics.Win.UltraWinEditors.UltraCheckEditor chkCertificado;
		public Infragistics.Win.UltraWinEditors.UltraCheckEditor chkCedula;
		public Infragistics.Win.UltraWinEditors.UltraCheckEditor chkFactura;
		public Infragistics.Win.UltraWinEditors.UltraCheckEditor chkProforma;
		private System.Windows.Forms.Button btnFormulario;
		private System.Windows.Forms.Button btnCedula;
		private System.Windows.Forms.Button btnCertificado;
		private System.Windows.Forms.Button btnFactura;
		private System.Windows.Forms.Button btnProforma;
		private System.Windows.Forms.Button btnPST;
		private System.Windows.Forms.Button btnInformeTecnico;
		private System.Windows.Forms.Button btnDenuncia;
		private System.Windows.Forms.Button btnIMEI;
		public Infragistics.Win.UltraWinEditors.UltraCheckEditor chkPST;
		public Infragistics.Win.UltraWinEditors.UltraCheckEditor chkInformeTecnico;
		public Infragistics.Win.UltraWinEditors.UltraCheckEditor chkCBIMEI;
		public Infragistics.Win.UltraWinEditors.UltraCheckEditor chkDenuncia;
		private System.Windows.Forms.Button btnCancelar;
		private System.Windows.Forms.Button btnAceptar;
		private C1.Data.C1DataSet cdsSeteoF;
		public Infragistics.Win.UltraWinEditors.UltraCheckEditor chkCartaAseguradora;
		private System.Windows.Forms.Button btnCartaAseguradora;

		int iTipo = 0;
		private System.Windows.Forms.Button btnCartaBroker;
		public Infragistics.Win.UltraWinEditors.UltraCheckEditor chkCartaBroker;
		int idDetCompra = 0;
		public Infragistics.Win.UltraWinEditors.UltraCheckEditor ChkCobroAseguradora;
		public Infragistics.Win.UltraWinEditors.UltraCheckEditor chkFacturaAseguradora;
		string sCarpetaDocs = "";

		public frmDocumentacionCompraProtegida(int ITipo, int IdDetCompra, string SCarpetaDocs)
		{
			//
			// Necesario para admitir el Diseñador de Windows Forms
			//
			iTipo = ITipo;
			idDetCompra = IdDetCompra;
			sCarpetaDocs = SCarpetaDocs;

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
			Infragistics.Win.Appearance appearance9 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance10 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance11 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance12 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance13 = new Infragistics.Win.Appearance();
			this.chkFormulario = new Infragistics.Win.UltraWinEditors.UltraCheckEditor();
			this.chkCertificado = new Infragistics.Win.UltraWinEditors.UltraCheckEditor();
			this.chkCedula = new Infragistics.Win.UltraWinEditors.UltraCheckEditor();
			this.chkFactura = new Infragistics.Win.UltraWinEditors.UltraCheckEditor();
			this.chkProforma = new Infragistics.Win.UltraWinEditors.UltraCheckEditor();
			this.btnFormulario = new System.Windows.Forms.Button();
			this.btnCedula = new System.Windows.Forms.Button();
			this.btnCertificado = new System.Windows.Forms.Button();
			this.btnFactura = new System.Windows.Forms.Button();
			this.btnProforma = new System.Windows.Forms.Button();
			this.btnPST = new System.Windows.Forms.Button();
			this.btnInformeTecnico = new System.Windows.Forms.Button();
			this.btnDenuncia = new System.Windows.Forms.Button();
			this.btnIMEI = new System.Windows.Forms.Button();
			this.chkPST = new Infragistics.Win.UltraWinEditors.UltraCheckEditor();
			this.chkInformeTecnico = new Infragistics.Win.UltraWinEditors.UltraCheckEditor();
			this.chkCBIMEI = new Infragistics.Win.UltraWinEditors.UltraCheckEditor();
			this.chkDenuncia = new Infragistics.Win.UltraWinEditors.UltraCheckEditor();
			this.btnCartaBroker = new System.Windows.Forms.Button();
			this.btnCartaAseguradora = new System.Windows.Forms.Button();
			this.chkCartaBroker = new Infragistics.Win.UltraWinEditors.UltraCheckEditor();
			this.chkCartaAseguradora = new Infragistics.Win.UltraWinEditors.UltraCheckEditor();
			this.btnAceptar = new System.Windows.Forms.Button();
			this.btnCancelar = new System.Windows.Forms.Button();
			this.cdsSeteoF = new C1.Data.C1DataSet();
			this.ChkCobroAseguradora = new Infragistics.Win.UltraWinEditors.UltraCheckEditor();
			this.chkFacturaAseguradora = new Infragistics.Win.UltraWinEditors.UltraCheckEditor();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).BeginInit();
			this.SuspendLayout();
			// 
			// chkFormulario
			// 
			appearance1.ForeColorDisabled = System.Drawing.Color.Black;
			this.chkFormulario.Appearance = appearance1;
			this.chkFormulario.BackColor = System.Drawing.Color.Transparent;
			this.chkFormulario.Location = new System.Drawing.Point(16, 9);
			this.chkFormulario.Name = "chkFormulario";
			this.chkFormulario.Size = new System.Drawing.Size(192, 22);
			this.chkFormulario.TabIndex = 27;
			this.chkFormulario.Text = "Formulario";
			// 
			// chkCertificado
			// 
			appearance2.ForeColorDisabled = System.Drawing.Color.Black;
			this.chkCertificado.Appearance = appearance2;
			this.chkCertificado.BackColor = System.Drawing.Color.Transparent;
			this.chkCertificado.Location = new System.Drawing.Point(16, 41);
			this.chkCertificado.Name = "chkCertificado";
			this.chkCertificado.Size = new System.Drawing.Size(192, 22);
			this.chkCertificado.TabIndex = 28;
			this.chkCertificado.Text = "Certificado Compra Protegida";
			// 
			// chkCedula
			// 
			appearance3.ForeColorDisabled = System.Drawing.Color.Black;
			this.chkCedula.Appearance = appearance3;
			this.chkCedula.BackColor = System.Drawing.Color.Transparent;
			this.chkCedula.Location = new System.Drawing.Point(16, 73);
			this.chkCedula.Name = "chkCedula";
			this.chkCedula.Size = new System.Drawing.Size(192, 22);
			this.chkCedula.TabIndex = 29;
			this.chkCedula.Text = "Cedula";
			// 
			// chkFactura
			// 
			appearance4.ForeColorDisabled = System.Drawing.Color.Black;
			this.chkFactura.Appearance = appearance4;
			this.chkFactura.BackColor = System.Drawing.Color.Transparent;
			this.chkFactura.Location = new System.Drawing.Point(16, 105);
			this.chkFactura.Name = "chkFactura";
			this.chkFactura.Size = new System.Drawing.Size(192, 22);
			this.chkFactura.TabIndex = 30;
			this.chkFactura.Text = "Factura";
			// 
			// chkProforma
			// 
			appearance5.ForeColorDisabled = System.Drawing.Color.Black;
			this.chkProforma.Appearance = appearance5;
			this.chkProforma.BackColor = System.Drawing.Color.Transparent;
			this.chkProforma.Location = new System.Drawing.Point(16, 137);
			this.chkProforma.Name = "chkProforma";
			this.chkProforma.Size = new System.Drawing.Size(192, 22);
			this.chkProforma.TabIndex = 31;
			this.chkProforma.Text = "Proforma Bien Adquirido";
			// 
			// btnFormulario
			// 
			this.btnFormulario.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnFormulario.Location = new System.Drawing.Point(16, 456);
			this.btnFormulario.Name = "btnFormulario";
			this.btnFormulario.Size = new System.Drawing.Size(168, 24);
			this.btnFormulario.TabIndex = 977;
			this.btnFormulario.Text = "Guardar Documentos";
			this.btnFormulario.Click += new System.EventHandler(this.btnFormulario_Click);
			// 
			// btnCedula
			// 
			this.btnCedula.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnCedula.Location = new System.Drawing.Point(192, 72);
			this.btnCedula.Name = "btnCedula";
			this.btnCedula.Size = new System.Drawing.Size(30, 24);
			this.btnCedula.TabIndex = 978;
			this.btnCedula.Text = "...";
			this.btnCedula.Visible = false;
			this.btnCedula.Click += new System.EventHandler(this.btnCedula_Click);
			// 
			// btnCertificado
			// 
			this.btnCertificado.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnCertificado.Location = new System.Drawing.Point(192, 40);
			this.btnCertificado.Name = "btnCertificado";
			this.btnCertificado.Size = new System.Drawing.Size(30, 24);
			this.btnCertificado.TabIndex = 979;
			this.btnCertificado.Text = "...";
			this.btnCertificado.Visible = false;
			this.btnCertificado.Click += new System.EventHandler(this.btnCertificado_Click);
			// 
			// btnFactura
			// 
			this.btnFactura.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnFactura.Location = new System.Drawing.Point(192, 104);
			this.btnFactura.Name = "btnFactura";
			this.btnFactura.Size = new System.Drawing.Size(30, 24);
			this.btnFactura.TabIndex = 980;
			this.btnFactura.Text = "...";
			this.btnFactura.Visible = false;
			this.btnFactura.Click += new System.EventHandler(this.btnFactura_Click);
			// 
			// btnProforma
			// 
			this.btnProforma.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnProforma.Location = new System.Drawing.Point(192, 136);
			this.btnProforma.Name = "btnProforma";
			this.btnProforma.Size = new System.Drawing.Size(30, 24);
			this.btnProforma.TabIndex = 981;
			this.btnProforma.Text = "...";
			this.btnProforma.Visible = false;
			this.btnProforma.Click += new System.EventHandler(this.btnProforma_Click);
			// 
			// btnPST
			// 
			this.btnPST.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnPST.Location = new System.Drawing.Point(192, 264);
			this.btnPST.Name = "btnPST";
			this.btnPST.Size = new System.Drawing.Size(30, 24);
			this.btnPST.TabIndex = 989;
			this.btnPST.Text = "...";
			this.btnPST.Visible = false;
			this.btnPST.Click += new System.EventHandler(this.btnPST_Click);
			// 
			// btnInformeTecnico
			// 
			this.btnInformeTecnico.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnInformeTecnico.Location = new System.Drawing.Point(192, 232);
			this.btnInformeTecnico.Name = "btnInformeTecnico";
			this.btnInformeTecnico.Size = new System.Drawing.Size(30, 24);
			this.btnInformeTecnico.TabIndex = 988;
			this.btnInformeTecnico.Text = "...";
			this.btnInformeTecnico.Visible = false;
			this.btnInformeTecnico.Click += new System.EventHandler(this.btnInformeTecnico_Click);
			// 
			// btnDenuncia
			// 
			this.btnDenuncia.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnDenuncia.Location = new System.Drawing.Point(192, 168);
			this.btnDenuncia.Name = "btnDenuncia";
			this.btnDenuncia.Size = new System.Drawing.Size(30, 24);
			this.btnDenuncia.TabIndex = 987;
			this.btnDenuncia.Text = "...";
			this.btnDenuncia.Visible = false;
			this.btnDenuncia.Click += new System.EventHandler(this.btnDenuncia_Click);
			// 
			// btnIMEI
			// 
			this.btnIMEI.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnIMEI.Location = new System.Drawing.Point(192, 200);
			this.btnIMEI.Name = "btnIMEI";
			this.btnIMEI.Size = new System.Drawing.Size(30, 24);
			this.btnIMEI.TabIndex = 986;
			this.btnIMEI.Text = "...";
			this.btnIMEI.Visible = false;
			this.btnIMEI.Click += new System.EventHandler(this.btnIMEI_Click);
			// 
			// chkPST
			// 
			appearance6.ForeColorDisabled = System.Drawing.Color.Black;
			this.chkPST.Appearance = appearance6;
			this.chkPST.BackColor = System.Drawing.Color.Transparent;
			this.chkPST.Location = new System.Drawing.Point(16, 264);
			this.chkPST.Name = "chkPST";
			this.chkPST.Size = new System.Drawing.Size(192, 22);
			this.chkPST.TabIndex = 985;
			this.chkPST.Text = "Proforma Servicio Tecnico";
			// 
			// chkInformeTecnico
			// 
			appearance7.ForeColorDisabled = System.Drawing.Color.Black;
			this.chkInformeTecnico.Appearance = appearance7;
			this.chkInformeTecnico.BackColor = System.Drawing.Color.Transparent;
			this.chkInformeTecnico.Location = new System.Drawing.Point(16, 232);
			this.chkInformeTecnico.Name = "chkInformeTecnico";
			this.chkInformeTecnico.Size = new System.Drawing.Size(192, 22);
			this.chkInformeTecnico.TabIndex = 984;
			this.chkInformeTecnico.Text = "Informe Tecnico";
			// 
			// chkCBIMEI
			// 
			appearance8.ForeColorDisabled = System.Drawing.Color.Black;
			this.chkCBIMEI.Appearance = appearance8;
			this.chkCBIMEI.BackColor = System.Drawing.Color.Transparent;
			this.chkCBIMEI.Location = new System.Drawing.Point(16, 200);
			this.chkCBIMEI.Name = "chkCBIMEI";
			this.chkCBIMEI.Size = new System.Drawing.Size(192, 22);
			this.chkCBIMEI.TabIndex = 983;
			this.chkCBIMEI.Text = "Certificado Bloqueo IMEI";
			// 
			// chkDenuncia
			// 
			appearance9.ForeColorDisabled = System.Drawing.Color.Black;
			this.chkDenuncia.Appearance = appearance9;
			this.chkDenuncia.BackColor = System.Drawing.Color.Transparent;
			this.chkDenuncia.Location = new System.Drawing.Point(16, 168);
			this.chkDenuncia.Name = "chkDenuncia";
			this.chkDenuncia.Size = new System.Drawing.Size(192, 22);
			this.chkDenuncia.TabIndex = 982;
			this.chkDenuncia.Text = "Denuncia Original";
			// 
			// btnCartaBroker
			// 
			this.btnCartaBroker.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnCartaBroker.Location = new System.Drawing.Point(192, 328);
			this.btnCartaBroker.Name = "btnCartaBroker";
			this.btnCartaBroker.Size = new System.Drawing.Size(30, 24);
			this.btnCartaBroker.TabIndex = 993;
			this.btnCartaBroker.Text = "...";
			this.btnCartaBroker.Visible = false;
			this.btnCartaBroker.Click += new System.EventHandler(this.btnCartaBroker_Click);
			// 
			// btnCartaAseguradora
			// 
			this.btnCartaAseguradora.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnCartaAseguradora.Location = new System.Drawing.Point(192, 296);
			this.btnCartaAseguradora.Name = "btnCartaAseguradora";
			this.btnCartaAseguradora.Size = new System.Drawing.Size(30, 24);
			this.btnCartaAseguradora.TabIndex = 992;
			this.btnCartaAseguradora.Text = "...";
			this.btnCartaAseguradora.Visible = false;
			this.btnCartaAseguradora.Click += new System.EventHandler(this.btnCartaAseguradora_Click);
			// 
			// chkCartaBroker
			// 
			appearance10.ForeColorDisabled = System.Drawing.Color.Black;
			this.chkCartaBroker.Appearance = appearance10;
			this.chkCartaBroker.BackColor = System.Drawing.Color.Transparent;
			this.chkCartaBroker.Location = new System.Drawing.Point(16, 328);
			this.chkCartaBroker.Name = "chkCartaBroker";
			this.chkCartaBroker.Size = new System.Drawing.Size(192, 22);
			this.chkCartaBroker.TabIndex = 991;
			this.chkCartaBroker.Text = "Carta Broker";
			// 
			// chkCartaAseguradora
			// 
			appearance11.ForeColorDisabled = System.Drawing.Color.Black;
			this.chkCartaAseguradora.Appearance = appearance11;
			this.chkCartaAseguradora.BackColor = System.Drawing.Color.Transparent;
			this.chkCartaAseguradora.Location = new System.Drawing.Point(16, 296);
			this.chkCartaAseguradora.Name = "chkCartaAseguradora";
			this.chkCartaAseguradora.Size = new System.Drawing.Size(192, 22);
			this.chkCartaAseguradora.TabIndex = 990;
			this.chkCartaAseguradora.Text = "Carta Aseguradora";
			// 
			// btnAceptar
			// 
			this.btnAceptar.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnAceptar.Location = new System.Drawing.Point(16, 488);
			this.btnAceptar.Name = "btnAceptar";
			this.btnAceptar.Size = new System.Drawing.Size(80, 24);
			this.btnAceptar.TabIndex = 994;
			this.btnAceptar.Text = "Aceptar";
			this.btnAceptar.Click += new System.EventHandler(this.btnAceptar_Click);
			// 
			// btnCancelar
			// 
			this.btnCancelar.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnCancelar.Location = new System.Drawing.Point(104, 488);
			this.btnCancelar.Name = "btnCancelar";
			this.btnCancelar.Size = new System.Drawing.Size(80, 24);
			this.btnCancelar.TabIndex = 995;
			this.btnCancelar.Text = "Cancelar";
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
			// ChkCobroAseguradora
			// 
			appearance12.ForeColorDisabled = System.Drawing.Color.Black;
			this.ChkCobroAseguradora.Appearance = appearance12;
			this.ChkCobroAseguradora.BackColor = System.Drawing.Color.Transparent;
			this.ChkCobroAseguradora.Location = new System.Drawing.Point(16, 360);
			this.ChkCobroAseguradora.Name = "ChkCobroAseguradora";
			this.ChkCobroAseguradora.Size = new System.Drawing.Size(192, 32);
			this.ChkCobroAseguradora.TabIndex = 996;
			this.ChkCobroAseguradora.Text = "¿ Cobro de la Aseguradora fue Ingresada?";
			// 
			// chkFacturaAseguradora
			// 
			appearance13.ForeColorDisabled = System.Drawing.Color.Black;
			this.chkFacturaAseguradora.Appearance = appearance13;
			this.chkFacturaAseguradora.BackColor = System.Drawing.Color.Transparent;
			this.chkFacturaAseguradora.Location = new System.Drawing.Point(16, 408);
			this.chkFacturaAseguradora.Name = "chkFacturaAseguradora";
			this.chkFacturaAseguradora.Size = new System.Drawing.Size(192, 32);
			this.chkFacturaAseguradora.TabIndex = 997;
			this.chkFacturaAseguradora.Text = "¿ Factura de la Aseguradora fue Ingresada?";
			this.chkFacturaAseguradora.CheckedChanged += new System.EventHandler(this.ultraCheckEditor1_CheckedChanged);
			// 
			// frmDocumentacionCompraProtegida
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 13);
			this.ClientSize = new System.Drawing.Size(234, 528);
			this.ControlBox = false;
			this.Controls.Add(this.chkFacturaAseguradora);
			this.Controls.Add(this.ChkCobroAseguradora);
			this.Controls.Add(this.btnCancelar);
			this.Controls.Add(this.btnAceptar);
			this.Controls.Add(this.btnCartaBroker);
			this.Controls.Add(this.btnCartaAseguradora);
			this.Controls.Add(this.chkCartaBroker);
			this.Controls.Add(this.chkCartaAseguradora);
			this.Controls.Add(this.btnPST);
			this.Controls.Add(this.btnInformeTecnico);
			this.Controls.Add(this.btnDenuncia);
			this.Controls.Add(this.btnIMEI);
			this.Controls.Add(this.chkPST);
			this.Controls.Add(this.chkInformeTecnico);
			this.Controls.Add(this.chkCBIMEI);
			this.Controls.Add(this.chkDenuncia);
			this.Controls.Add(this.btnProforma);
			this.Controls.Add(this.btnFactura);
			this.Controls.Add(this.btnCertificado);
			this.Controls.Add(this.btnCedula);
			this.Controls.Add(this.btnFormulario);
			this.Controls.Add(this.chkProforma);
			this.Controls.Add(this.chkFactura);
			this.Controls.Add(this.chkCedula);
			this.Controls.Add(this.chkCertificado);
			this.Controls.Add(this.chkFormulario);
			this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog;
			this.KeyPreview = true;
			this.MaximizeBox = false;
			this.MinimizeBox = false;
			this.Name = "frmDocumentacionCompraProtegida";
			this.ShowInTaskbar = false;
			this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
			this.Text = "Documentación Compra Protegida";
			this.Load += new System.EventHandler(this.frmDocumentacionCompraProtegida_Load);
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).EndInit();
			this.ResumeLayout(false);

		}
		#endregion


		public string sFormulario = "";
		public string sCertificado = "";
		public string sCedula = "";
		public string sFactura = "";
		public string sProforma = "";
		public string sDenuncia =  "";
		public string sIMEI = "";
		public string sInformeTecnico = "";
		public string sPST = "";
		public string sAseguradora = "";
		public string sBroker = "";
		public int iEstado = 3;

		private void frmDocumentacionCompraProtegida_Load(object sender, System.EventArgs e)
		{
			if (iTipo == 1)
			{
				this.chkInformeTecnico.Enabled = false;
				this.chkPST.Enabled = false;
			}

			if (iTipo == 2)
			{
				this.chkCBIMEI.Enabled = false;
				this.chkDenuncia.Enabled = false;
			}

			/*
			Formulario, Certificado, CedulaSeguro, Factura, Proforma, Denuncia, 5
			BloqueoIMEI, InformeTecnico, ProformaServicioTecnico, CartaAseg, CartaBroker, 10
			DirFormulario, DirCertificado, DirCedulaSeguro, DirFactura, DirProforma, 
			DirDenuncia, DirBloqueoIMEI, DirInformeTecnico, DirProformaServicioTecnico, 
			DirCartaAseg, DirCartaBroker
			*/
			string sSQL = string.Format("Select ISNULL(Formulario, 0), ISNULL(Certificado, 0), ISNULL(CedulaSeguro, 0), ISNULL(Factura, 0), ISNULL(Proforma, 0), ISNULL(Denuncia, 0), ISNULL(BloqueoIMEI, 0), ISNULL(InformeTecnico, 0), ISNULL(ProformaServicioTecnico, 0), ISNULL(CartaAseg, 0), ISNULL(CartaBroker, 0), 	ISNULL(DirFormulario, ''), ISNULL(DirCertificado, ''), ISNULL(DirCedulaSeguro, ''), ISNULL(DirFactura, ''), ISNULL(DirProforma, ''), ISNULL(DirDenuncia, ''), ISNULL(DirBloqueoIMEI, ''), ISNULL(DirInformeTecnico, ''), ISNULL(DirProformaServicioTecnico, ''), ISNULL(DirCartaAseg, ''), ISNULL(DirCartaBroker, '') From DetCompra Where idDetCompra = {0}", idDetCompra);
			SqlDataReader dr = Funcion.rEscalarSQL(cdsSeteoF, sSQL, true);
			dr.Read();
			if (dr.HasRows)
			{
				this.chkFormulario.Checked = dr.GetBoolean(0);
				this.chkCertificado.Checked = dr.GetBoolean(1);
				this.chkCedula.Checked = dr.GetBoolean(2);
				this.chkFactura.Checked = dr.GetBoolean(3);
				this.chkProforma.Checked = dr.GetBoolean(4);
				this.chkDenuncia.Checked = dr.GetBoolean(5);
				this.chkCBIMEI.Checked = dr.GetBoolean(6);
				this.chkInformeTecnico.Checked = dr.GetBoolean(7);
				this.chkPST.Checked = dr.GetBoolean(8);
				this.chkCartaAseguradora.Checked = dr.GetBoolean(9);
				this.chkCartaBroker.Checked = dr.GetBoolean(10);
				sFormulario = dr.GetString(11);
				sCertificado = dr.GetString(12);
				sCedula = dr.GetString(13);
				sFactura = dr.GetString(14);
				sProforma = dr.GetString(15);
				sDenuncia = dr.GetString(16);
				sIMEI = dr.GetString(17);
				sInformeTecnico = dr.GetString(18);
				sPST = dr.GetString(19);
				sAseguradora = dr.GetString(20);
				sBroker = dr.GetString(21);
			}
			dr.Close();
		}

		private void btnCancelar_Click(object sender, System.EventArgs e)
		{
			this.DialogResult = DialogResult.No;
		}

		private void btnAceptar_Click(object sender, System.EventArgs e)
		{
			int iCont = 0;

			if (this.chkFormulario.Checked) iCont++;
			if (this.chkCertificado.Checked) iCont++;
			if (this.chkCedula.Checked) iCont++;
			if (this.chkFactura.Checked) iCont++;
			if (this.chkProforma.Checked) iCont++;
			if (this.chkCartaAseguradora.Checked) iCont++;
			if (this.chkCartaBroker.Checked) iCont++;

			if (this.chkDenuncia.Checked) iCont++;
			if (this.chkCBIMEI.Checked) iCont++;
			if (this.chkInformeTecnico.Checked) iCont++;
			if (this.chkPST.Checked) iCont++;
			
			//if (sFormulario.Length > 0) iEstado = 5;
			if (iCont >= 7) iEstado = 5;

			if ((bool)this.ChkCobroAseguradora.Checked == true && (bool)this.chkFacturaAseguradora.Checked == true) iEstado = 6;

			this.DialogResult = DialogResult.OK;
		}

		private void cdsSeteoF_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			cdsSeteoF.Schema.Connections[0].ConnectionString = miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);		
		}

		private string Documentos(string sCarpeta, string sArchivo)
		{
			string sOriginal = "";
			string sDocumento = "";

			OpenFileDialog ofd = new OpenFileDialog();
			ofd.InitialDirectory = @"C:\";
			ofd.Filter = "PDF: (*.pdf)(*.pdf)|*.pdf";
			if(ofd.ShowDialog() == DialogResult.OK)
			{
				sOriginal = ofd.FileName;

				sDocumento = sCarpeta + sArchivo + ".pdf";

				File.Copy(sOriginal, sDocumento);
			}

			return sDocumento;
		}

		private void btnFormulario_Click(object sender, System.EventArgs e)
		{
			sFormulario = Documentos(sCarpetaDocs, "Documentos");

			if (sFormulario.Length > 0) this.chkFormulario.Checked = true; else this.chkFormulario.Checked = false;
		}

		private void btnCertificado_Click(object sender, System.EventArgs e)
		{
			sCertificado = Documentos(sCarpetaDocs, "Certificado");

			if (sCertificado.Length > 0) this.chkCertificado.Checked = true; else this.chkCertificado.Checked = false;
		}

		private void btnCedula_Click(object sender, System.EventArgs e)
		{
			sCedula = Documentos(sCarpetaDocs, "Cedula");

			if (sCedula.Length > 0) this.chkCedula.Checked = true; else this.chkCedula.Checked = false;
		}

		private void btnFactura_Click(object sender, System.EventArgs e)
		{
			sFactura = Documentos(sCarpetaDocs, "Factura");

			if (sFactura.Length > 0) this.chkFactura.Checked = true; else this.chkFactura.Checked = false;
		}

		private void btnProforma_Click(object sender, System.EventArgs e)
		{
			sProforma = Documentos(sCarpetaDocs, "Proforma");

			if (sProforma.Length > 0) this.chkProforma.Checked = true; else this.chkProforma.Checked = false;
		}

		private void btnDenuncia_Click(object sender, System.EventArgs e)
		{
			sDenuncia = Documentos(sCarpetaDocs, "Denuncia");

			if (sDenuncia.Length > 0) this.chkDenuncia.Checked = true; else this.chkDenuncia.Checked = false;
		}

		private void btnIMEI_Click(object sender, System.EventArgs e)
		{
			sIMEI = Documentos(sCarpetaDocs, "IMEI");

			if (sIMEI.Length > 0) this.chkCBIMEI.Checked = true; else this.chkCBIMEI.Checked = false;
		}

		private void btnInformeTecnico_Click(object sender, System.EventArgs e)
		{
			sInformeTecnico = Documentos(sCarpetaDocs, "Informe_Tecnico");

			if (sInformeTecnico.Length > 0) this.chkInformeTecnico.Checked = true; else this.chkInformeTecnico.Checked = false;
		}

		private void btnPST_Click(object sender, System.EventArgs e)
		{
			sPST = Documentos(sCarpetaDocs, "ProformaServicioTecnico");

			if (sPST.Length > 0) this.chkPST.Checked = true; else this.chkPST.Checked = false;
		}

		private void btnCartaAseguradora_Click(object sender, System.EventArgs e)
		{
			sAseguradora = Documentos(sCarpetaDocs, "Carta_Aseguradora");

			if (sAseguradora.Length > 0) this.chkCartaAseguradora.Checked = true; else this.chkCartaAseguradora.Checked = false;
		}

		private void btnCartaBroker_Click(object sender, System.EventArgs e)
		{
			sBroker = Documentos(sCarpetaDocs, "Carta_Broker");

			if (sBroker.Length > 0) this.chkCartaBroker.Checked = true; else this.chkCartaBroker.Checked = false;
		}

		private void ultraCheckEditor1_CheckedChanged(object sender, System.EventArgs e)
		{
		
		}
	}
}