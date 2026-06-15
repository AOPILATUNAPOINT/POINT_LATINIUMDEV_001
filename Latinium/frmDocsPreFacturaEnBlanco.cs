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
using System.Net;
using System.Web;
using System.Xml;
using System.Xml.Schema;
using CrystalDecisions.Shared;

namespace Latinium
{
	/// <summary>
	/// Descripción breve de frmDocsPreFacturaEnBlanco.
	/// </summary>
	public class frmDocsPreFacturaEnBlanco : System.Windows.Forms.Form
	{
		private System.Windows.Forms.GroupBox groupBox1;
		private Infragistics.Win.UltraWinEditors.UltraCheckEditor chkCompromisoDePago;
		public DevExpress.XtraEditors.SimpleButton btCancelar;
		public DevExpress.XtraEditors.SimpleButton btGrabar;
		private C1.Data.C1DataSet cdsSeteoF;
		private System.Windows.Forms.ContextMenu contextMenu1;
		private System.Windows.Forms.MenuItem mnuImpresora;
		private System.Windows.Forms.MenuItem mnuPantalla;
		private Infragistics.Win.UltraWinEditors.UltraCheckEditor chkPagareB;
		private Infragistics.Win.UltraWinEditors.UltraCheckEditor chkPagareD;
		private Infragistics.Win.UltraWinEditors.UltraCheckEditor chkGastosDeCobranzaB;
		private Infragistics.Win.UltraWinEditors.UltraCheckEditor chkAmortizacionesB;
		private Infragistics.Win.UltraWinEditors.UltraCheckEditor chkAmortizacionesD;
		private Infragistics.Win.UltraWinEditors.UltraCheckEditor chkGastosDeCobranzaD;
		private Infragistics.Win.UltraWinEditors.UltraCheckEditor chkSeguroB;
		private Infragistics.Win.UltraWinEditors.UltraCheckEditor chkSeguroD;
		/// <summary>
		/// Variable del diseñador requerida.
		/// </summary>
		private System.ComponentModel.Container components = null;
		int IdCompra = 0;
		decimal DCapital = 0;
		int ICuotas = 0;
		int IdCliente = 0;
		int IdEntFinanciera = 0;
//		public DevExpress.XtraEditors.SimpleButton btCancelar;
		string Fecha = "";
//		private C1.Data.C1DataSet cdsSeteoF;
//		private System.Windows.Forms.ContextMenu contextMenu1;
//		private System.Windows.Forms.MenuItem mnuImpresora;
//		private System.Windows.Forms.MenuItem mnuPantalla;
		DateTime FHasta;
//		private System.Windows.Forms.GroupBox groupBox1;
//		private Infragistics.Win.UltraWinEditors.UltraCheckEditor chkCompromisoDePago;
		string sRUC = "";

		public frmDocsPreFacturaEnBlanco()
		{
			//
			// Necesario para admitir el Diseñador de Windows Forms
			//
			InitializeComponent();

			//
			// TODO: agregar código de constructor después de llamar a InitializeComponent
			//
		}
    
		public frmDocsPreFacturaEnBlanco(int idCompra, decimal dCapital, int iCuotas, int idCliente, string fecha, DateTime fHasta, int idEntFinanciera, string SRUC)
		{
			//
			// Required for Windows Form Designer support
			//
			InitializeComponent();

			IdCompra = idCompra;
			DCapital = dCapital;
			ICuotas = iCuotas;
			IdCliente = idCliente;
			Fecha = fecha;
			FHasta = fHasta;
			IdEntFinanciera = idEntFinanciera;
			sRUC = SRUC;
			//
			// TODO: Add any constructor code after InitializeComponent call
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
			this.groupBox1 = new System.Windows.Forms.GroupBox();
			this.chkCompromisoDePago = new Infragistics.Win.UltraWinEditors.UltraCheckEditor();
			this.btCancelar = new DevExpress.XtraEditors.SimpleButton();
			this.btGrabar = new DevExpress.XtraEditors.SimpleButton();
			this.chkGastosDeCobranzaB = new Infragistics.Win.UltraWinEditors.UltraCheckEditor();
			this.chkAmortizacionesB = new Infragistics.Win.UltraWinEditors.UltraCheckEditor();
			this.chkPagareB = new Infragistics.Win.UltraWinEditors.UltraCheckEditor();
			this.chkPagareD = new Infragistics.Win.UltraWinEditors.UltraCheckEditor();
			this.chkAmortizacionesD = new Infragistics.Win.UltraWinEditors.UltraCheckEditor();
			this.chkGastosDeCobranzaD = new Infragistics.Win.UltraWinEditors.UltraCheckEditor();
			this.cdsSeteoF = new C1.Data.C1DataSet();
			this.contextMenu1 = new System.Windows.Forms.ContextMenu();
			this.mnuImpresora = new System.Windows.Forms.MenuItem();
			this.mnuPantalla = new System.Windows.Forms.MenuItem();
			this.chkSeguroB = new Infragistics.Win.UltraWinEditors.UltraCheckEditor();
			this.chkSeguroD = new Infragistics.Win.UltraWinEditors.UltraCheckEditor();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).BeginInit();
			this.SuspendLayout();
			// 
			// groupBox1
			// 
			this.groupBox1.Location = new System.Drawing.Point(0, 208);
			this.groupBox1.Name = "groupBox1";
			this.groupBox1.Size = new System.Drawing.Size(320, 8);
			this.groupBox1.TabIndex = 178;
			this.groupBox1.TabStop = false;
			// 
			// chkCompromisoDePago
			// 
			appearance1.ForeColorDisabled = System.Drawing.Color.Black;
			this.chkCompromisoDePago.Appearance = appearance1;
			this.chkCompromisoDePago.BackColor = System.Drawing.Color.Transparent;
			this.chkCompromisoDePago.Checked = true;
			this.chkCompromisoDePago.CheckState = System.Windows.Forms.CheckState.Checked;
			this.chkCompromisoDePago.Location = new System.Drawing.Point(40, 168);
			this.chkCompromisoDePago.Name = "chkCompromisoDePago";
			this.chkCompromisoDePago.Size = new System.Drawing.Size(192, 32);
			this.chkCompromisoDePago.TabIndex = 177;
			this.chkCompromisoDePago.Text = "Compromiso Lugar de Pago";
			// 
			// btCancelar
			// 
			this.btCancelar.Anchor = System.Windows.Forms.AnchorStyles.None;
			this.btCancelar.Location = new System.Drawing.Point(152, 224);
			this.btCancelar.Name = "btCancelar";
			this.btCancelar.Size = new System.Drawing.Size(88, 24);
			this.btCancelar.TabIndex = 176;
			this.btCancelar.Text = "&Cancelar";
			this.btCancelar.Click += new System.EventHandler(this.btCancelar_Click);
			// 
			// btGrabar
			// 
			this.btGrabar.Anchor = System.Windows.Forms.AnchorStyles.None;
			this.btGrabar.Location = new System.Drawing.Point(48, 224);
			this.btGrabar.Name = "btGrabar";
			this.btGrabar.Size = new System.Drawing.Size(88, 24);
			this.btGrabar.TabIndex = 175;
			this.btGrabar.Text = "&Imprimir";
			this.btGrabar.Click += new System.EventHandler(this.btGrabar_Click);
			// 
			// chkGastosDeCobranzaB
			// 
			appearance2.ForeColorDisabled = System.Drawing.Color.Black;
			this.chkGastosDeCobranzaB.Appearance = appearance2;
			this.chkGastosDeCobranzaB.BackColor = System.Drawing.Color.Transparent;
			this.chkGastosDeCobranzaB.Location = new System.Drawing.Point(16, 96);
			this.chkGastosDeCobranzaB.Name = "chkGastosDeCobranzaB";
			this.chkGastosDeCobranzaB.Size = new System.Drawing.Size(16, 22);
			this.chkGastosDeCobranzaB.TabIndex = 174;
			this.chkGastosDeCobranzaB.Text = "Gastos de cobranza en Blanco";
			this.chkGastosDeCobranzaB.Visible = false;
			// 
			// chkAmortizacionesB
			// 
			appearance3.ForeColorDisabled = System.Drawing.Color.Black;
			this.chkAmortizacionesB.Appearance = appearance3;
			this.chkAmortizacionesB.BackColor = System.Drawing.Color.Transparent;
			this.chkAmortizacionesB.Location = new System.Drawing.Point(16, 40);
			this.chkAmortizacionesB.Name = "chkAmortizacionesB";
			this.chkAmortizacionesB.Size = new System.Drawing.Size(16, 48);
			this.chkAmortizacionesB.TabIndex = 173;
			this.chkAmortizacionesB.Text = "Tabla De Amortización En Blanco";
			this.chkAmortizacionesB.Visible = false;
			// 
			// chkPagareB
			// 
			appearance4.ForeColorDisabled = System.Drawing.Color.Black;
			this.chkPagareB.Appearance = appearance4;
			this.chkPagareB.BackColor = System.Drawing.Color.Transparent;
			this.chkPagareB.Location = new System.Drawing.Point(16, 8);
			this.chkPagareB.Name = "chkPagareB";
			this.chkPagareB.Size = new System.Drawing.Size(16, 22);
			this.chkPagareB.TabIndex = 172;
			this.chkPagareB.Text = "Pagaré  en Blanco";
			this.chkPagareB.Visible = false;
			// 
			// chkPagareD
			// 
			appearance5.ForeColorDisabled = System.Drawing.Color.Black;
			this.chkPagareD.Appearance = appearance5;
			this.chkPagareD.BackColor = System.Drawing.Color.Transparent;
			this.chkPagareD.Checked = true;
			this.chkPagareD.CheckState = System.Windows.Forms.CheckState.Checked;
			this.chkPagareD.Location = new System.Drawing.Point(40, 8);
			this.chkPagareD.Name = "chkPagareD";
			this.chkPagareD.Size = new System.Drawing.Size(128, 22);
			this.chkPagareD.TabIndex = 179;
			this.chkPagareD.Text = "Pagaré  Con Datos";
			// 
			// chkAmortizacionesD
			// 
			appearance6.ForeColorDisabled = System.Drawing.Color.Black;
			this.chkAmortizacionesD.Appearance = appearance6;
			this.chkAmortizacionesD.BackColor = System.Drawing.Color.Transparent;
			this.chkAmortizacionesD.Checked = true;
			this.chkAmortizacionesD.CheckState = System.Windows.Forms.CheckState.Checked;
			this.chkAmortizacionesD.Location = new System.Drawing.Point(40, 40);
			this.chkAmortizacionesD.Name = "chkAmortizacionesD";
			this.chkAmortizacionesD.Size = new System.Drawing.Size(144, 48);
			this.chkAmortizacionesD.TabIndex = 180;
			this.chkAmortizacionesD.Text = "Tabla De Amortizacón Con Datos";
			// 
			// chkGastosDeCobranzaD
			// 
			appearance7.ForeColorDisabled = System.Drawing.Color.Black;
			this.chkGastosDeCobranzaD.Appearance = appearance7;
			this.chkGastosDeCobranzaD.BackColor = System.Drawing.Color.Transparent;
			this.chkGastosDeCobranzaD.Checked = true;
			this.chkGastosDeCobranzaD.CheckState = System.Windows.Forms.CheckState.Checked;
			this.chkGastosDeCobranzaD.Location = new System.Drawing.Point(40, 96);
			this.chkGastosDeCobranzaD.Name = "chkGastosDeCobranzaD";
			this.chkGastosDeCobranzaD.Size = new System.Drawing.Size(128, 22);
			this.chkGastosDeCobranzaD.TabIndex = 181;
			this.chkGastosDeCobranzaD.Text = "Gastos de cobranza con Datos";
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
			// contextMenu1
			// 
			this.contextMenu1.MenuItems.AddRange(new System.Windows.Forms.MenuItem[] {
																																								 this.mnuImpresora,
																																								 this.mnuPantalla});
			// 
			// mnuImpresora
			// 
			this.mnuImpresora.Index = 0;
			this.mnuImpresora.Text = "Impresora";
			// 
			// mnuPantalla
			// 
			this.mnuPantalla.Index = 1;
			this.mnuPantalla.Text = "Pantalla";
			this.mnuPantalla.Click += new System.EventHandler(this.mnuPantalla_Click);
			// 
			// chkSeguroB
			// 
			appearance8.ForeColorDisabled = System.Drawing.Color.Black;
			this.chkSeguroB.Appearance = appearance8;
			this.chkSeguroB.BackColor = System.Drawing.Color.Transparent;
			this.chkSeguroB.Location = new System.Drawing.Point(16, 128);
			this.chkSeguroB.Name = "chkSeguroB";
			this.chkSeguroB.Size = new System.Drawing.Size(16, 32);
			this.chkSeguroB.TabIndex = 182;
			this.chkSeguroB.Text = "Seguro D&D en Blanco";
			this.chkSeguroB.Visible = false;
			// 
			// chkSeguroD
			// 
			appearance9.ForeColorDisabled = System.Drawing.Color.Black;
			this.chkSeguroD.Appearance = appearance9;
			this.chkSeguroD.BackColor = System.Drawing.Color.Transparent;
			this.chkSeguroD.Checked = true;
			this.chkSeguroD.CheckState = System.Windows.Forms.CheckState.Checked;
			this.chkSeguroD.Location = new System.Drawing.Point(40, 128);
			this.chkSeguroD.Name = "chkSeguroD";
			this.chkSeguroD.Size = new System.Drawing.Size(144, 32);
			this.chkSeguroD.TabIndex = 183;
			this.chkSeguroD.Text = "Seguro D&D Con Datos";
			// 
			// frmDocsPreFacturaEnBlanco
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 13);
			this.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(247)), ((System.Byte)(245)), ((System.Byte)(241)));
			this.ClientSize = new System.Drawing.Size(320, 254);
			this.Controls.Add(this.chkSeguroD);
			this.Controls.Add(this.chkSeguroB);
			this.Controls.Add(this.chkGastosDeCobranzaD);
			this.Controls.Add(this.chkAmortizacionesD);
			this.Controls.Add(this.chkPagareD);
			this.Controls.Add(this.groupBox1);
			this.Controls.Add(this.chkCompromisoDePago);
			this.Controls.Add(this.btCancelar);
			this.Controls.Add(this.btGrabar);
			this.Controls.Add(this.chkGastosDeCobranzaB);
			this.Controls.Add(this.chkAmortizacionesB);
			this.Controls.Add(this.chkPagareB);
			this.MaximizeBox = false;
			this.MinimizeBox = false;
			this.Name = "frmDocsPreFacturaEnBlanco";
			this.Text = "Documentos Pre Factura";
			this.KeyDown += new System.Windows.Forms.KeyEventHandler(this.frmDocsPreFacturaEnBlanco_KeyDown);
			this.Load += new System.EventHandler(this.frmDocsPreFacturaEnBlanco_Load);
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).EndInit();
			this.ResumeLayout(false);

		}
		#endregion

		private void btGrabar_Click(object sender, System.EventArgs e)
		{
			using (frmAvisoFactura Gestion = new frmAvisoFactura( IdCompra ))
			{
				if (DialogResult.OK == Gestion.ShowDialog( ))
				{
					this.contextMenu1.Show(this.btGrabar, new Point(1, 25));
				}
			}
		}
    int iDyD = 0;
		private void frmDocsPreFacturaEnBlanco_Load(object sender, System.EventArgs e)
		{
			iDyD = Funcion.iEscalarSQL(cdsSeteoF, string.Format("select Count ( *) from DetCompra where idArticulo = 8683 and idcompra = {0}", IdCompra));	
			if ( iDyD == 0)
			{
				this.chkSeguroB.Checked = false;
				this.chkSeguroD.Checked = false;
				this.chkSeguroB.Enabled = false;
				this.chkSeguroD.Enabled = false;
			}
			else
			{
				this.chkSeguroB.Checked = true;
				this.chkSeguroD.Checked = true;
				this.chkSeguroB.Enabled = true;
				this.chkSeguroD.Enabled = true;
			}

			string carpeta_1 ="";
			carpeta_1 = @"C:\Latinium\Reportes\PagareConDatos_2022.rpt";
		}

		private void cdsSeteoF_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			cdsSeteoF.Schema.Connections[0].ConnectionString = miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);	
		}

		private void btCancelar_Click(object sender, System.EventArgs e)
		{
			this.Close();
		}

		private void mnuPantalla_Click(object sender, System.EventArgs e)
		{
			this.Cursor = Cursors.WaitCursor;

			int iReporte = IdEntFinanciera;
			#region Pagare
			if (this.chkPagareB.Checked)
			{				
				string sPagare = "PagareConDatos_2022.rpt";
				if (iReporte == 2) sPagare = "PagareConDatos_2022.rpt";

				ParameterFields paramFields = new ParameterFields ();
				
				ParameterField paramFieldCompra = new ParameterField ();
				ParameterDiscreteValue ValIdCompra = new ParameterDiscreteValue ();
				paramFieldCompra.ParameterFieldName = "@idCompra";
				ValIdCompra.Value = int.Parse(IdCompra.ToString());
				paramFieldCompra.CurrentValues.Add (ValIdCompra);
				paramFields.Add (paramFieldCompra);

				ParameterField paramFieldFecha = new ParameterField ();
				ParameterDiscreteValue ValIdFecha = new ParameterDiscreteValue ();
				paramFieldFecha.ParameterFieldName = "@FechaGenera";
				ValIdFecha.Value = FHasta;
				paramFieldFecha.CurrentValues.Add (ValIdFecha);
				paramFields.Add (paramFieldFecha);
				
				Reporte miRep = new Reporte(sPagare, "");
				miRep.MdiParent = this.MdiParent;
				miRep.Titulo("Págare");
				miRep.crVista.ParameterFieldInfo = paramFields;
				miRep.ShowDialog();
			}

			if (this.chkPagareD.Checked)
			{				
				string sPagare = "PagareEnBlanco.rpt";
				if (iReporte == 2) sPagare = "C.rpt";

				ParameterFields paramFields = new ParameterFields ();
				
				ParameterField paramFieldCompra = new ParameterField ();
				ParameterDiscreteValue ValIdCompra = new ParameterDiscreteValue ();
				paramFieldCompra.ParameterFieldName = "@idCompra";
				ValIdCompra.Value = int.Parse(IdCompra.ToString());
				paramFieldCompra.CurrentValues.Add (ValIdCompra);
				paramFields.Add (paramFieldCompra);

				ParameterField paramFieldFecha = new ParameterField ();
				ParameterDiscreteValue ValIdFecha = new ParameterDiscreteValue ();
				paramFieldFecha.ParameterFieldName = "@FechaGenera";
				ValIdFecha.Value = FHasta;
				paramFieldFecha.CurrentValues.Add (ValIdFecha);
				paramFields.Add (paramFieldFecha);
				
				Reporte miRep = new Reporte(sPagare, "");
				miRep.MdiParent = this.MdiParent;
				miRep.Titulo("Págare");
				miRep.crVista.ParameterFieldInfo = paramFields;
				miRep.ShowDialog();
			}

			
   #endregion Pagare

			#region Amortizacion
			if (this.chkAmortizacionesB.Checked)
			{
				bool bAbonaCuota = Funcion.bEscalarSQL(cdsSeteoF, string.Format("Select AbonaCuota From Compra Where idCompra = {0}", IdCompra), true);
				ParameterFields paramFields = new ParameterFields ();


//				
//				ParameterField pfCapital = new ParameterField ();
//				ParameterDiscreteValue ValCapital = new ParameterDiscreteValue ();
//				pfCapital.ParameterFieldName = "@Capital";
//				ValCapital.Value = DCapital;
//				pfCapital.CurrentValues.Add (ValCapital);
//				paramFields.Add (pfCapital);
//				
//				ParameterField pfCuotas = new ParameterField ();
//				ParameterDiscreteValue ValCuotas = new ParameterDiscreteValue ();
//				pfCuotas.ParameterFieldName = "@Cuotas";
//				ValCuotas.Value = ICuotas;
//				pfCuotas.CurrentValues.Add (ValCuotas);
//				paramFields.Add (pfCuotas);
//				
//				ParameterField pfIdCliente = new ParameterField ();
//				ParameterDiscreteValue ValIdCliente = new ParameterDiscreteValue ();
//				pfIdCliente.ParameterFieldName = "@IdCliente";
//				ValIdCliente.Value = IdCliente;
//				pfIdCliente.CurrentValues.Add (ValIdCliente);
//				paramFields.Add (pfIdCliente);
//				
//				ParameterField pfFecha = new ParameterField ();
//				ParameterDiscreteValue ValFecha = new ParameterDiscreteValue ();
//				pfFecha.ParameterFieldName = "@Fecha";
//				ValFecha.Value = Fecha;
//				pfFecha.CurrentValues.Add (ValFecha);
//				paramFields.Add (pfFecha);
//				
//				ParameterField pfFechaGenera = new ParameterField ();
//				ParameterDiscreteValue ValFechaGenera = new ParameterDiscreteValue ();
//				pfFechaGenera.ParameterFieldName = "@FechaGenera";
//				ValFechaGenera.Value = FHasta;
//				pfFechaGenera.CurrentValues.Add (ValFechaGenera);
//				paramFields.Add (pfFechaGenera);
//
//				ParameterField pfEF = new ParameterField ();
//				ParameterDiscreteValue ValEF = new ParameterDiscreteValue ();
//				pfEF.ParameterFieldName = "@idEntidadFinanciera";
//				ValEF.Value = IdEntFinanciera;
//				pfEF.CurrentValues.Add (ValEF);
//				paramFields.Add (pfEF);


				
				Reporte miRepor = new Reporte("TablaAmortizacionDatos.rpt", "");
				miRepor.MdiParent = this.MdiParent;
				miRepor.Titulo("Págare");
				miRepor.crVista.ParameterFieldInfo = paramFields;
				miRepor.ShowDialog();
			}
    

			if (this.chkAmortizacionesD.Checked)
			{
				bool bAbonaCuota = Funcion.bEscalarSQL(cdsSeteoF, string.Format("Select AbonaCuota From Compra Where idCompra = {0}", IdCompra), true);

				ParameterFields paramFields = new ParameterFields ();
				
				ParameterField pfCapital = new ParameterField ();
				ParameterDiscreteValue ValCapital = new ParameterDiscreteValue ();
				pfCapital.ParameterFieldName = "@Capital";
				ValCapital.Value = DCapital;
				pfCapital.CurrentValues.Add (ValCapital);
				paramFields.Add (pfCapital);
				
				ParameterField pfCuotas = new ParameterField ();
				ParameterDiscreteValue ValCuotas = new ParameterDiscreteValue ();
				pfCuotas.ParameterFieldName = "@Cuotas";
				ValCuotas.Value = ICuotas;
				pfCuotas.CurrentValues.Add (ValCuotas);
				paramFields.Add (pfCuotas);
				
				ParameterField pfIdCliente = new ParameterField ();
				ParameterDiscreteValue ValIdCliente = new ParameterDiscreteValue ();
				pfIdCliente.ParameterFieldName = "@IdCliente";
				ValIdCliente.Value = IdCliente;
				pfIdCliente.CurrentValues.Add (ValIdCliente);
				paramFields.Add (pfIdCliente);
				
				ParameterField pfFecha = new ParameterField ();
				ParameterDiscreteValue ValFecha = new ParameterDiscreteValue ();
				pfFecha.ParameterFieldName = "@Fecha";
				ValFecha.Value = Fecha;
				pfFecha.CurrentValues.Add (ValFecha);
				paramFields.Add (pfFecha);
				
				ParameterField pfFechaGenera = new ParameterField ();
				ParameterDiscreteValue ValFechaGenera = new ParameterDiscreteValue ();
				pfFechaGenera.ParameterFieldName = "@FechaGenera";
				ValFechaGenera.Value = FHasta;
				pfFechaGenera.CurrentValues.Add (ValFechaGenera);
				paramFields.Add (pfFechaGenera);

				ParameterField pfEF = new ParameterField ();
				ParameterDiscreteValue ValEF = new ParameterDiscreteValue ();
				pfEF.ParameterFieldName = "@idEntidadFinanciera";
				ValEF.Value = IdEntFinanciera;
				pfEF.CurrentValues.Add (ValEF);
				paramFields.Add (pfEF);

				ParameterField pfAC = new ParameterField ();
				ParameterDiscreteValue ValAC = new ParameterDiscreteValue ();
				pfAC.ParameterFieldName = "@AbonaCuota";
				ValAC.Value = bAbonaCuota;
				pfAC.CurrentValues.Add (ValAC);
				paramFields.Add (pfAC);
				
				Reporte miRepor = new Reporte("TablaAmortizacionBlanco.rpt", "");
				miRepor.MdiParent = this.MdiParent;
				miRepor.Titulo("Págare");
				miRepor.crVista.ParameterFieldInfo = paramFields;
				miRepor.ShowDialog();
			}
     #endregion Amortizacion

			#region Gastos De cobranza

			if (this.chkGastosDeCobranzaB.Checked)
			{
				string sReporte = "";
				
				// Funcion.iEscalarSQL(cdsSeteoF, string.Format("Select dbo.TipoDeEntidad({0})", IdCompra));

				if (iReporte == 1) sReporte = "GestiondeCobranzaDatos.rpt";
				if (iReporte == 4) sReporte = "GestiondeCobranzaDatos.rpt";
				else if (iReporte == 2) sReporte = "GestiondeCobranzaSolidarioDatos.rpt";

				ParameterFields paramFields = new ParameterFields ();
				
				ParameterField paramFieldCompra = new ParameterField ();
				ParameterDiscreteValue ValIdCompra = new ParameterDiscreteValue ();
				paramFieldCompra.ParameterFieldName = "@idCompra";
				ValIdCompra.Value = int.Parse(IdCompra.ToString());
				paramFieldCompra.CurrentValues.Add (ValIdCompra);
				paramFields.Add (paramFieldCompra);

				int bRef = 0;
				ParameterField paramFieldRef = new ParameterField ();
				ParameterDiscreteValue ValIdRef = new ParameterDiscreteValue ();
				paramFieldRef.ParameterFieldName = "@Refinanciacion";
				ValIdRef.Value = int.Parse(bRef.ToString());;
				paramFieldRef.CurrentValues.Add (ValIdRef);
				paramFields.Add (paramFieldRef);

				Reporte miReporte = new Reporte(sReporte, "");
				miReporte.MdiParent = MdiParent;
				miReporte.Titulo("Gastos de Cobranza");
				miReporte.crVista.ParameterFieldInfo = paramFields;
				miReporte.ShowDialog();
			}	

			if (this.chkGastosDeCobranzaD.Checked)
			{
				string sReporte = "";
				
				// Funcion.iEscalarSQL(cdsSeteoF, string.Format("Select dbo.TipoDeEntidad({0})", IdCompra));

				if (iReporte == 1) sReporte = "GestiondeCobranzaBlanco.rpt";
				if (iReporte == 4) sReporte = "GestiondeCobranzaBlanco.rpt";
				else if (iReporte == 2) sReporte = "GestiondeCobranzaSolidarioBlanco.rpt";

				ParameterFields paramFields = new ParameterFields ();
				
				ParameterField paramFieldCompra = new ParameterField ();
				ParameterDiscreteValue ValIdCompra = new ParameterDiscreteValue ();
				paramFieldCompra.ParameterFieldName = "@idCompra";
				ValIdCompra.Value = int.Parse(IdCompra.ToString());
				paramFieldCompra.CurrentValues.Add (ValIdCompra);
				paramFields.Add (paramFieldCompra);

				int bRef = 0;
				ParameterField paramFieldRef = new ParameterField ();
				ParameterDiscreteValue ValIdRef = new ParameterDiscreteValue ();
				paramFieldRef.ParameterFieldName = "@Refinanciacion";
				ValIdRef.Value = int.Parse(bRef.ToString());;
				paramFieldRef.CurrentValues.Add (ValIdRef);
				paramFields.Add (paramFieldRef);

				Reporte miReporte = new Reporte(sReporte, "");
				miReporte.MdiParent = MdiParent;
				miReporte.Titulo("Gastos de Cobranza");
				miReporte.crVista.ParameterFieldInfo = paramFields;
				miReporte.ShowDialog();
			}	
			#endregion Gastos De cobranza

			if (this.chkCompromisoDePago.Checked)
			{					
				ParameterFields paramFields = new ParameterFields ();
				
				ParameterField paramFieldCompra = new ParameterField ();
				ParameterDiscreteValue ValIdCompra = new ParameterDiscreteValue ();
				paramFieldCompra.ParameterFieldName = "@idCompra";
				ValIdCompra.Value = int.Parse(IdCompra.ToString());
				paramFieldCompra.CurrentValues.Add (ValIdCompra);
				paramFields.Add (paramFieldCompra);		
				
				
				Reporte miRep = new Reporte("CompromisoLugarDePagoDatosV2.rpt", "");
				//				Reporte miRep = new Reporte("Cre_CompromisoLugarDePago.rpt", "");
				miRep.MdiParent = this.MdiParent;
				miRep.Titulo("Compromiso Lugar de Pago");
				miRep.crVista.ParameterFieldInfo = paramFields;
				miRep.ShowDialog();
			}

			if ( iDyD > 0)
			{
//				if (this.chkSeguroB.Checked )
//				{
//					ParameterFields paramFields = new ParameterFields ();
//					ParameterField paramFieldIdDetCompra = new ParameterField ();
//					ParameterDiscreteValue discreteValIdDetCompra = new ParameterDiscreteValue ();
//					paramFieldIdDetCompra.ParameterFieldName = "@idCompra";
//					discreteValIdDetCompra.Value = int.Parse(IdCompra.ToString()); 
//					paramFieldIdDetCompra.CurrentValues.Add (discreteValIdDetCompra);
//					paramFields.Add (paramFieldIdDetCompra);
//				
//					Reporte Reporte = new Reporte("DatosSeguroDesempleoDatos.rpt", "");//ReporteProteccionDYD.rpt", "");
//					Reporte.MdiParent = this.MdiParent;
//					Reporte.Titulo("Proteccion D&D");
//					Reporte.crVista.ParameterFieldInfo = paramFields;
////					Reporte.Show();		
//		
//					Reporte.MdiParent = this.MdiParent;
//					Reporte.Titulo("Seguro D&D");
//					Reporte.crVista.ParameterFieldInfo = paramFields;
//					Reporte.ShowDialog();
//	
//				}

				if (this.chkSeguroD.Checked )
				{
					ParameterFields paramFields = new ParameterFields ();
					ParameterField paramFieldIdDetCompra = new ParameterField ();
					ParameterDiscreteValue discreteValIdDetCompra = new ParameterDiscreteValue ();
					paramFieldIdDetCompra.ParameterFieldName = "@idCompra";
					discreteValIdDetCompra.Value = int.Parse(IdCompra.ToString());
					paramFieldIdDetCompra.CurrentValues.Add (discreteValIdDetCompra);
					paramFields.Add (paramFieldIdDetCompra);
				
					Reporte Reporte = new Reporte("DatosSeguroDesempleoBlanco.rpt", "");//ReporteProteccionDYD.rpt", "");
					Reporte.MdiParent = this.MdiParent;
					Reporte.Titulo("Proteccion D&D");
					Reporte.crVista.ParameterFieldInfo = paramFields;
//					Reporte.Show();		
		
					Reporte.MdiParent = this.MdiParent;
					Reporte.Titulo("Seguro D&D");
					Reporte.crVista.ParameterFieldInfo = paramFields;
					Reporte.ShowDialog();
	
				}
			}
			
			Funcion.EjecutaSQL(cdsSeteoF, string.Format("Update Compra Set Impreso = Case idTipoFactura When 43 Then 1 When 1 Then Impreso End Where idCompra = {0}", IdCompra), true);

			this.btCancelar.Text = "&Cerrar";			
			//this.btGrabar.Enabled = false;
			this.Close();
			this.Cursor = Cursors.Default;
		}

		private void frmDocsPreFacturaEnBlanco_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Escape) this.Close();
		}
	}
}
