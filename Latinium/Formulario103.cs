using System;
using System.Data.SqlClient;
using System.Globalization;
using C1.Data;
using System.IO;
using System.Xml;
using System.Xml.Schema;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using LibContabilidad.DataObjects;
using LibFacturacion.DataObjects;
using CrystalDecisions.Shared;

namespace Latinium
{
	/// <summary>
	/// Descripción breve de AnexoTransaccional.
	/// </summary>
	public class Formulario103 : DevExpress.XtraEditors.XtraForm
	{
		private Infragistics.Win.Misc.UltraButton btGenerar;
		private C1.Data.C1DataSet cdsSeteo;
		private C1.Data.C1DataSet cdsCompra;
		private Infragistics.Win.Misc.UltraButton btValidar;
		private Infragistics.Win.Misc.UltraButton btForm104;
		private Infragistics.Win.UltraWinSchedule.UltraCalendarCombo cmbFecha;
		private C1.Data.C1DataSet cdsSeteoC;
		private Infragistics.Win.Misc.UltraButton btVerificacion;
		private Infragistics.Win.Misc.UltraButton btRep104;
		private Infragistics.Win.Misc.UltraButton btRep103;
		private Infragistics.Win.UltraWinSchedule.UltraCalendarCombo cmbDesde;
		private Infragistics.Win.UltraWinSchedule.UltraCalendarCombo cmbHasta;
		private Infragistics.Win.UltraWinEditors.UltraCheckEditor chkResumen;
		private Infragistics.Win.Misc.UltraButton btForm101;
		private Infragistics.Win.Misc.UltraButton btRep101;
		private C1.Data.C1DataSet cdsSeteoF;
		/// <summary>
		/// Variable del diseñador requerida.
		/// </summary>
		private System.ComponentModel.Container components = null;

		public Formulario103()
		{
			InitializeComponent();
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
			Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton dateButton1 = new Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton();
			Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton dateButton2 = new Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton();
			Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton dateButton3 = new Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton();
			this.btGenerar = new Infragistics.Win.Misc.UltraButton();
			this.cdsSeteo = new C1.Data.C1DataSet();
			this.cdsCompra = new C1.Data.C1DataSet();
			this.btValidar = new Infragistics.Win.Misc.UltraButton();
			this.btForm104 = new Infragistics.Win.Misc.UltraButton();
			this.cmbFecha = new Infragistics.Win.UltraWinSchedule.UltraCalendarCombo();
			this.cdsSeteoC = new C1.Data.C1DataSet();
			this.btVerificacion = new Infragistics.Win.Misc.UltraButton();
			this.btRep104 = new Infragistics.Win.Misc.UltraButton();
			this.btRep103 = new Infragistics.Win.Misc.UltraButton();
			this.cmbDesde = new Infragistics.Win.UltraWinSchedule.UltraCalendarCombo();
			this.cmbHasta = new Infragistics.Win.UltraWinSchedule.UltraCalendarCombo();
			this.chkResumen = new Infragistics.Win.UltraWinEditors.UltraCheckEditor();
			this.btRep101 = new Infragistics.Win.Misc.UltraButton();
			this.btForm101 = new Infragistics.Win.Misc.UltraButton();
			this.cdsSeteoF = new C1.Data.C1DataSet();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteo)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsCompra)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbFecha)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoC)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbDesde)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbHasta)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).BeginInit();
			this.SuspendLayout();
			// 
			// btGenerar
			// 
			this.btGenerar.Location = new System.Drawing.Point(200, 284);
			this.btGenerar.Name = "btGenerar";
			this.btGenerar.Size = new System.Drawing.Size(64, 25);
			this.btGenerar.TabIndex = 0;
			this.btGenerar.Text = "Generar";
			this.btGenerar.Visible = false;
			this.btGenerar.Click += new System.EventHandler(this.btGenerar_Click);
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
			// cdsCompra
			// 
			this.cdsCompra.BindingContextCtrl = this;
			this.cdsCompra.DataLibrary = "LibFacturacion";
			this.cdsCompra.DataLibraryUrl = "";
			this.cdsCompra.DataSetDef = "dsBodega";
			this.cdsCompra.EnforceConstraints = false;
			this.cdsCompra.Locale = new System.Globalization.CultureInfo("es-EC");
			this.cdsCompra.Name = "cdsCompra";
			this.cdsCompra.SchemaClassName = "LibFacturacion.DataClass";
			this.cdsCompra.SchemaDef = null;
			this.cdsCompra.BeforeFill += new C1.Data.FillEventHandler(this.cdsCompra_BeforeFill);
			// 
			// btValidar
			// 
			this.btValidar.Location = new System.Drawing.Point(80, 103);
			this.btValidar.Name = "btValidar";
			this.btValidar.Size = new System.Drawing.Size(120, 25);
			this.btValidar.TabIndex = 2;
			this.btValidar.Text = "Formulario 10&3";
			this.btValidar.Click += new System.EventHandler(this.btValidar_Click);
			// 
			// btForm104
			// 
			this.btForm104.Location = new System.Drawing.Point(80, 138);
			this.btForm104.Name = "btForm104";
			this.btForm104.Size = new System.Drawing.Size(120, 25);
			this.btForm104.TabIndex = 3;
			this.btForm104.Text = "Formulario 10&4";
			this.btForm104.Click += new System.EventHandler(this.btForm104_Click);
			// 
			// cmbFecha
			// 
			dateButton1.Caption = "Today";
			this.cmbFecha.DateButtons.Add(dateButton1);
			this.cmbFecha.Format = "MMMM/yyyy";
			this.cmbFecha.Location = new System.Drawing.Point(80, 26);
			this.cmbFecha.Name = "cmbFecha";
			this.cmbFecha.NonAutoSizeHeight = 8;
			this.cmbFecha.Size = new System.Drawing.Size(128, 21);
			this.cmbFecha.SpinButtonIncrement = Infragistics.Win.UltraWinSchedule.SpinIncrementUnit.Months;
			this.cmbFecha.SpinButtonsVisible = true;
			this.cmbFecha.TabIndex = 4;
			this.cmbFecha.Value = new System.DateTime(2007, 2, 27, 0, 0, 0, 0);
			// 
			// cdsSeteoC
			// 
			this.cdsSeteoC.BindingContextCtrl = this;
			this.cdsSeteoC.DataLibrary = "LibContabilidad";
			this.cdsSeteoC.DataLibraryUrl = "";
			this.cdsSeteoC.DataSetDef = "dsSeteoC";
			this.cdsSeteoC.Locale = new System.Globalization.CultureInfo("es-EC");
			this.cdsSeteoC.Name = "cdsSeteoC";
			this.cdsSeteoC.SchemaClassName = "LibContabilidad.DataClass";
			this.cdsSeteoC.SchemaDef = null;
			// 
			// btVerificacion
			// 
			this.btVerificacion.Location = new System.Drawing.Point(72, 284);
			this.btVerificacion.Name = "btVerificacion";
			this.btVerificacion.Size = new System.Drawing.Size(120, 25);
			this.btVerificacion.TabIndex = 5;
			this.btVerificacion.Text = "Verificación";
			this.btVerificacion.Click += new System.EventHandler(this.btVerificacion_Click);
			// 
			// btRep104
			// 
			this.btRep104.Location = new System.Drawing.Point(168, 208);
			this.btRep104.Name = "btRep104";
			this.btRep104.Size = new System.Drawing.Size(96, 25);
			this.btRep104.TabIndex = 6;
			this.btRep104.Text = "Reporte 104";
			this.btRep104.Click += new System.EventHandler(this.btRep104_Click);
			// 
			// btRep103
			// 
			this.btRep103.Location = new System.Drawing.Point(208, 103);
			this.btRep103.Name = "btRep103";
			this.btRep103.Size = new System.Drawing.Size(24, 25);
			this.btRep103.TabIndex = 7;
			this.btRep103.Text = "...";
			this.btRep103.Visible = false;
			this.btRep103.Click += new System.EventHandler(this.btRep103_Click);
			// 
			// cmbDesde
			// 
			dateButton2.Caption = "Today";
			this.cmbDesde.DateButtons.Add(dateButton2);
			this.cmbDesde.Format = "dd/MMM/yyyy";
			this.cmbDesde.Location = new System.Drawing.Point(24, 190);
			this.cmbDesde.Name = "cmbDesde";
			this.cmbDesde.NonAutoSizeHeight = 8;
			this.cmbDesde.Size = new System.Drawing.Size(128, 21);
			this.cmbDesde.SpinButtonsVisible = true;
			this.cmbDesde.TabIndex = 8;
			this.cmbDesde.Value = new System.DateTime(2007, 2, 27, 0, 0, 0, 0);
			// 
			// cmbHasta
			// 
			dateButton3.Caption = "Today";
			this.cmbHasta.DateButtons.Add(dateButton3);
			this.cmbHasta.Format = "dd/MMM/yyyy";
			this.cmbHasta.Location = new System.Drawing.Point(24, 224);
			this.cmbHasta.Name = "cmbHasta";
			this.cmbHasta.NonAutoSizeHeight = 8;
			this.cmbHasta.Size = new System.Drawing.Size(128, 21);
			this.cmbHasta.SpinButtonsVisible = true;
			this.cmbHasta.TabIndex = 9;
			this.cmbHasta.Value = new System.DateTime(2007, 2, 27, 0, 0, 0, 0);
			// 
			// chkResumen
			// 
			this.chkResumen.Location = new System.Drawing.Point(32, 250);
			this.chkResumen.Name = "chkResumen";
			this.chkResumen.Size = new System.Drawing.Size(120, 21);
			this.chkResumen.TabIndex = 10;
			this.chkResumen.Text = "Resumen";
			// 
			// btRep101
			// 
			this.btRep101.Location = new System.Drawing.Point(168, 176);
			this.btRep101.Name = "btRep101";
			this.btRep101.Size = new System.Drawing.Size(96, 24);
			this.btRep101.TabIndex = 11;
			this.btRep101.Text = "Reporte 101";
			this.btRep101.Visible = false;
			this.btRep101.Click += new System.EventHandler(this.btRep101_Click);
			// 
			// btForm101
			// 
			this.btForm101.Location = new System.Drawing.Point(80, 69);
			this.btForm101.Name = "btForm101";
			this.btForm101.Size = new System.Drawing.Size(120, 25);
			this.btForm101.TabIndex = 12;
			this.btForm101.Text = "Formulario 10&1";
			this.btForm101.Visible = false;
			this.btForm101.Click += new System.EventHandler(this.btForm101_Click);
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
			// 
			// Formulario103
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 14);
			this.ClientSize = new System.Drawing.Size(296, 324);
			this.Controls.Add(this.btForm101);
			this.Controls.Add(this.btRep101);
			this.Controls.Add(this.chkResumen);
			this.Controls.Add(this.cmbHasta);
			this.Controls.Add(this.cmbDesde);
			this.Controls.Add(this.btRep103);
			this.Controls.Add(this.btRep104);
			this.Controls.Add(this.btVerificacion);
			this.Controls.Add(this.cmbFecha);
			this.Controls.Add(this.btForm104);
			this.Controls.Add(this.btValidar);
			this.Controls.Add(this.btGenerar);
			this.MaximizeBox = false;
			this.Name = "Formulario103";
			this.Text = "Formularios";
			this.Closing += new System.ComponentModel.CancelEventHandler(this.Formulario103_Closing);
			this.Load += new System.EventHandler(this.AnexoTransaccional_Load);
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteo)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsCompra)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbFecha)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoC)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbDesde)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbHasta)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).EndInit();
			this.ResumeLayout(false);

		}
		#endregion

		private string NombreGenerado()
		{
			string stArchivo = "AT";
			DateTime dtFecha = (DateTime) cmbFecha.Value;
			stArchivo += dtFecha.ToString("MMyyyy") + ".XML";
			if (Directory.Exists(@"C:\XML_ANEXO_TRANSACCIONAL\"))
				stArchivo = @"C:\XML_ANEXO_TRANSACCIONAL\" + stArchivo;
			return stArchivo;
		}

		private void btGenerar_Click(object sender, System.EventArgs e)
		{
			Cursor = Cursors.WaitCursor;
			DateTime dtFecha = (DateTime) cmbFecha.Value;
			#region Encabezado
			XmlTextWriter miXML = new XmlTextWriter(NombreGenerado(), null);
			miXML.Formatting = Formatting.Indented;
			miXML.WriteStartDocument(false);
			miXML.WriteStartElement("iva");
			string stMensaje = "";
			string stVerificaRuc = Funcion.VerificaCedula(drSeteo.Ruc, false);
			if (stVerificaRuc.Length == 0) miXML.WriteElementString("numeroRuc", drSeteo.Ruc);
			else stMensaje += "Error en RUC de su Empresa\n";
			if (drSeteo.Empresa.Trim().Length > 2) miXML.WriteElementString("razonSocial", drSeteo.Empresa.ToUpper());
			else stMensaje += "Longitud de Nombre de su Empresa Erronea\n";
			if (drSeteo.Direccion.Trim().Length > 5) miXML.WriteElementString("direccionMatriz", drSeteo.Direccion.ToUpper());
			else stMensaje += "Direccion de su Empresa Erronea\n";
			if (drSeteo.Telefono.Length == 9) miXML.WriteElementString("telefono", drSeteo.Telefono);
			else stMensaje += "Telefono debe tener 9 digitos\n";
			miXML.WriteElementString("fax", drSeteo.Fax);
			miXML.WriteElementString("email", drSeteo.Email);
			miXML.WriteElementString("tpIdRepre", drSeteo.CedPasRep);
			if (drSeteo.CedPasRep != "P")	stVerificaRuc = Funcion.VerificaCedula(drSeteo.CedulaRep, false);
			else stVerificaRuc = "";
			if (stVerificaRuc.Length == 0) miXML.WriteElementString("idRepre", drSeteo.CedulaRep);
			else stMensaje += "Error en Cedula de Representante Legal\n";
			stVerificaRuc = Funcion.VerificaCedula(drSeteo.CedulaContador, false);
			if (stVerificaRuc.Length == 0) miXML.WriteElementString("rucContador", drSeteo.CedulaContador + "001");
			else stMensaje += "Error en Cedula de Contador";

			miXML.WriteElementString("anio", dtFecha.ToString("yyyy"));
			miXML.WriteElementString("mes", dtFecha.ToString("MM"));

			if (stMensaje.Trim().Length > 0)
			{
				MessageBox.Show(stMensaje, "Información");
				using (Seteo miSeteo = new Seteo())
				{
					miSeteo.ShowDialog();
				}
				miXML.Close();
				cdsSeteo.Clear();
				cdsSeteo.Fill();
				C1DataRow dr = this.cdsSeteo.TableViews[0].Rows[0];
				drSeteo = SeteoRow_tableView.Obj(dr);
				Cursor = Cursors.Default;
				return;
			}
			#endregion encabezado

			#region Compras
			miXML.WriteStartElement("compras");
			string stSelect = "Exec CompraIvaServicio '" + dtFecha.ToString("yyyyMMdd") + "'";
			Funcion.EjecutaSQL(cdsCompra, stSelect, true);

			// Creacion de Devolucion de compras
			stSelect = "SELECT IsNull(Compra.idCliente, 0), "
				+ "IsNull(Compra.Comprobante, ''), IsNull(Compra.idCompra, 0) "
				+ "FROM Compra LEFT OUTER JOIN CompraComprobante "
				+ "ON Compra.idComprobante = CompraComprobante.idComprobante "
				+ "Where Compra.idTipoFactura = 6 And Compra.Borrar = 0 And Compra.Fecha >= '" 
				+ dtFecha.ToString("yyyyMM") + "01' And Compra.Fecha < '" 
				+ dtFecha.AddMonths(1).ToString("yyyyMM") + "01'"
				+ " And Not CompraComprobante.Codigo In ('17', '99') ";
			SqlDataReader miReader = Funcion.rEscalarSQL(cdsCompra, stSelect, true);
			int k = 0;
			int iMax = 1000;
			int [] IdDevCliente = new int [iMax];
			int [] IdDevCompra = new int [iMax];
			string [] stDevComprob = new string [iMax];
			string [] stDevSerie = new string[iMax];
			DateTime [] dtDevFecha = new DateTime [iMax];
			string [] stDevAutoriza = new string[iMax];
			string [] stDevCompTipo = new string[iMax];
			while (miReader.Read())
			{
				IdDevCliente[k] = miReader.GetInt32(0);
				stDevComprob[k] = miReader.GetString(1);
				IdDevCompra[k] = miReader.GetInt32(2);
				k++;
			}
			miReader.Close();
			for (int j=0; j< k; j++)
			{
				stSelect = "SELECT IsNull(Compra.SerieFactura, ''), "
					+ "IsNull(Compra.Fecha, DateAdd(year, -100, GetDate())), "
					+ "IsNull(Compra.AutFactura, ''), "
					+ "IsNull(CompraComprobante.Codigo, '') FROM Compra "
					+ "Inner Join CompraComprobante On Compra.idComprobante "
					+ "= CompraComprobante.idComprobante "
					+ "WHERE Compra.Numero = '" + stDevComprob[j].Trim()
					+ "' AND Compra.idCliente = " + IdDevCliente[j].ToString();
				miReader = Funcion.rEscalarSQL(cdsCompra, stSelect, true);
				miReader.Read();
				stDevSerie[j] = miReader.GetString(0);
				stDevSerie[j] = Digito(stDevSerie[j]);
				dtDevFecha[j] = miReader.GetDateTime(1);
				stDevAutoriza[j] = miReader.GetString(2);
				stDevAutoriza[j] = Digito(stDevAutoriza[j]);
				stDevCompTipo[j] = miReader.GetString(3);
				stDevCompTipo[j] = Digito(stDevCompTipo[j]);
				miReader.Close();

				stDevComprob[j] = Digito(stDevComprob[j]);
				#region Verificaciones
				if (stDevCompTipo[j].Trim().Length < 1)
				{
					MessageBox.Show("Tipo de Comprobante, de Comprobante Asociado no existe",	"Información");
//					new Compra(6, IdDevCompra[j]).Show();
					miXML.Close();
					Cursor = Cursors.Default;
					return;
				}
				if (stDevSerie[j].Trim().Length < 6)
				{
					MessageBox.Show("Serie de Comprobante Asociado no existe",	"Información");
//					new Compra(6, IdDevCompra[j]).Show();
					miXML.Close();
					Cursor = Cursors.Default;
					return;
				}
				if (stDevAutoriza[j].Trim().Length < 6)
				{
					MessageBox.Show("Autorizacion de Comprobante Asociado no existe",	"Información");
//					new Compra(6, IdDevCompra[j]).Show();
					miXML.Close();
					Cursor = Cursors.Default;
					return;
				}
				#endregion Verificaciones
			}
			k = 0; // Incializado para que lo recorra nuevamente en devoluciones

			#region Select de Compras
			stSelect = "SELECT IsNull(CompraCredTribut.Codigo, ''), "
				+ "Compra.DerechoDevolucion, IsNull(ClienteTipoRuc.Codigo, '') ,"
				+ "IsNull(Cliente.Ruc, ''), IsNull(CompraComprobante.Codigo, ''), "
				+ "IsNull(Compra.Numero, ''), IsNull(Cliente.idCliente,0), "
				+ "IsNull(Compra.SerieFactura, ''), IsNull(Compra.AutFactura, ''), "
				+ "IsNull(Compra.FechaCaducidad, DateAdd(year, -100,GetDate())), "
				+ "IsNull(Compra.SubtotalExcento, 0), IsNull(Compra.SubtotalIva, 0), "
				+ "IsNull(Compra.Iva, 0), Compra.idCompra, IsNull(RetencionIva.Porcentaje, 0), "
				+ "IsNull(RetencionIva.Codigo, '0'), IsNull(RetIva_1.Porcentaje, 0), "
				+ "IsNull(RetIva_1.Codigo, '0'), "
				+ "IsNull(Compra.Fecha, DateAdd(year, -100,GetDate())), "
				+ "IsNull(Retencion.IvaRetenido, 0), IsNull(Retencion.IvaRetenido1, 0), "
				+ "IsNull(Retencion.idRetencion, 0), IsNull(Compra.IvaServicio, 0), "
				+ "IsNull(RetDetalle.Retenido, 0), IsNull(RetDetalle.Base, 0), "
				+ "IsNull(ArticuloSri.Porcentaje, 0), IsNull(ArticuloSri.Codigo, 0), "
				+ "IsNull(Retencion.SerieRet, ''), IsNull(Retencion.AutorizaRet, ''), "
				+ "IsNull(Retencion.Codigo, ''), "
				+ "IsNull(Retencion.Fecha, DateAdd(year, -100,GetDate())), "
				+ "IsNull(Compra.CuentaRetDet, 0), IsNull(Compra.idTipoFactura, 0) "
				+ "FROM Compra LEFT OUTER JOIN Retencion ON Retencion.idCompra = Compra.idCompra "
				+ "LEFT OUTER JOIN Cliente ON Compra.idCliente = Cliente.idCliente "
				+ "LEFT OUTER JOIN CompraCredTribut ON Compra.idCredTributario = CompraCredTribut.idCredTributario "
				+ "LEFT OUTER JOIN CompraComprobante ON Compra.idComprobante = CompraComprobante.idComprobante "
				+ "LEFT OUTER JOIN ClienteTipoRuc ON Cliente.idTipoRuc = ClienteTipoRuc.idTipoRuc "
				+ "LEFT OUTER JOIN RetencionIva ON Retencion.idRetIva = RetencionIva.idRetIva "
				+ "LEFT OUTER JOIN RetencionIva As RetIva_1  ON Retencion.idRetIva1 = RetIva_1.idRetIva "
				+ "LEFT OUTER JOIN RetDetalle ON RetDetalle.idRetencion = Retencion.idRetencion "
				+ "LEFT OUTER JOIN ArticuloSri ON RetDetalle.idRetRenta = ArticuloSri.idArticuloSri "
				+ "Where Compra.idTipoFactura In (4, 6) And Compra.Borrar = 0 And Compra.Fecha >= '" 
				+ dtFecha.ToString("yyyyMM") + "01' And Compra.Fecha < '" 
				+ dtFecha.AddMonths(1).ToString("yyyyMM") + "01'"
				+ " And Not CompraComprobante.Codigo In ('17', '99') ";
			#endregion Select de Compras
			#region Definicion de Variables de compras
			miReader = Funcion.rEscalarSQL(cdsCompra, stSelect, true);
			string stCodCredTrib = "";
			bool bDerechoDev = false;
			string stCodTipoRuc = "";
			string stRuc = "";
			string stAutFactura = "";
			string stCodComprob = "";
			string stNumero = "";
			string stCompraSerieFact = "";
			DateTime dtCaducidad = DateTime.Today;
			double dSubtotalExcento = 0;
			double dSubtotalIva = 0;
			double dCompraIva = 0;
			int IdCompra = 0;
			double dRetIvaPorc = 0;
			string stCodRetIva = "";
			double dRetIvaPorc1 = 0;
			string stCodRetIva1 = "";
			DateTime dtFechaCompra = DateTime.Today;
			DateTime dtClientCaduca = DateTime.Today;
			double dRetValorIva = 0;
			double dRetValorIva1 = 0;
			int IdRetencion = 0;
			double dIvaServicio = 0;
			int iCuentaRet = 0;
			int IdCliente = 0;
			int IdTipoFactura = 0;
			#endregion fin de definicion de variables
			#region Lectura de Filas de Compras
			while (miReader.Read())
			{
				miXML.WriteStartElement("detalleCompras");
				stCodCredTrib = miReader.GetString(0);
				bDerechoDev = miReader.GetBoolean(1);
				stCodTipoRuc = miReader.GetString(2);
				stRuc = miReader.GetString(3).Trim();
				IdCliente = miReader.GetInt32(6);
				string stVerifica = Funcion.VerificaCedula(stRuc, false);
				if (stRuc.Length != 10 && stCodTipoRuc == "02") stVerifica = "Err";
				if (stRuc.Length != 13 && stCodTipoRuc == "01") stVerifica = "Err";
				if (stCodTipoRuc == "03") stVerifica = "";
				if (stVerifica.Length > 0)
				{
					MessageBox.Show("Error en Cedula o RUC", "Información");
					new Cliente(false, IdCliente).Show();
					miXML.Close();
					Cursor = Cursors.Default;
					return;
				}
				stCodComprob =  miReader.GetString(4);
				stNumero =  miReader.GetString(5);
				stCompraSerieFact = miReader.GetString(7);
				stAutFactura = miReader.GetString(8);
				dtCaducidad =  miReader.GetDateTime(9);
				dSubtotalExcento = miReader.GetDouble(10);
				dSubtotalIva = miReader.GetDouble(11);
				dCompraIva = miReader.GetDouble(12);
				IdCompra = miReader.GetInt32(13);
				dRetIvaPorc = miReader.GetDouble(14);
				stCodRetIva = miReader.GetString(15);
				dRetIvaPorc1 = miReader.GetDouble(16);
				stCodRetIva1 = miReader.GetString(17);
				dtFechaCompra = miReader.GetDateTime(18);
				dRetValorIva = miReader.GetDouble(19);
				dRetValorIva1 = miReader.GetDouble(20);
				IdRetencion = miReader.GetInt32(21);
				dIvaServicio = miReader.GetDouble(22);
				iCuentaRet = miReader.GetInt32(31);
				IdTipoFactura = miReader.GetInt32(32);
				//				dRetValorIva1 = miReader.GetDouble(23);
//				dRetValorRenta1 = miReader.GetDouble(24);
//				stCodRetRenta = miReader.GetString(25);
//				dRetRentaaPorc = miReader.GetDouble(26);
//				stCodRetRenta1 = miReader.GetString(27);
//				dRetRentaaPorc1 = miReader.GetDouble(28);
//				stCodSRI = miReader.GetString(29);
				#endregion Lectura de Filas de Compras
				#region Sustento, Devolucion de Iva
				miXML.WriteElementString("codSustento", stCodCredTrib); // Identificacion del Sustento Tributario
				string stCodigo = "S";
				if (!bDerechoDev) stCodigo = "N";
				miXML.WriteElementString("devIva", stCodigo);
				#endregion fin de sustento
				#region Proveedor
				miXML.WriteElementString("tpIdProv", stCodTipoRuc);
				miXML.WriteElementString("idProv", stRuc);
				if (stCodTipoRuc == "02" && stCodComprob.Trim() != "03")
				{
					MessageBox.Show("Con cedula solo se permite liquidacion de compras",
						"Información");
//					new Compra(IdTipoFactura, IdCompra).Show();
					miXML.Close();
					Cursor = Cursors.Default;
					return;
				}
				
				#endregion Fin de Proveedor
				#region Comprobante, Fecha Asiento, Serie, Fecha Emision
				// Codigo Tipo de Comprobante
				miXML.WriteElementString("tipoComprobante", int.Parse(stCodComprob).ToString());

				miXML.WriteElementString("fechaRegistro", dtFechaCompra.ToString("dd/MM/yyyy"));
				stNumero = Digito(stNumero);
				string stEstablecmiento = "";
				string stPuntoEmision = "";
				string stSecuencial = "";
				if (stNumero.Trim().Length == 13)
				{
					stEstablecmiento = stNumero.Substring(0, 3);
					stPuntoEmision = stNumero.Substring(3,3);
					stSecuencial = stNumero.Substring(6,7);
				}
				else 
				{
					stCompraSerieFact = Digito(stCompraSerieFact);
					if (stCompraSerieFact.Trim().Length == 6)
					{
						stEstablecmiento = stCompraSerieFact.Substring(0, 3);
						stPuntoEmision = stCompraSerieFact.Substring(3,3);
					}
					stSecuencial = stNumero;
				}
				if (stEstablecmiento.Length != 3  || stPuntoEmision.Length != 3 || stNumero.Length == 0)
				{
					MessageBox.Show("Error en Numero o Serie de Factura", "Información");
//					new Compra(IdTipoFactura, IdCompra).Show();
					miXML.Close();
					Cursor = Cursors.Default;
					return;
				}
				miXML.WriteElementString("establecimiento", stEstablecmiento);
				miXML.WriteElementString("puntoEmision", stPuntoEmision);
				miXML.WriteElementString("secuencial", stNumero);

				// Fecha de Emision
				miXML.WriteElementString("fechaEmision", dtFechaCompra.ToString("dd/MM/yyyy"));
				#endregion fin de serie
				#region Autorizacion y caducidad
				if (stAutFactura.Length < 4)
				{
					MessageBox.Show("Error en Numero de Autorizacion de Factura", "Información");
//					new Compra(IdTipoFactura, IdCompra).Show();
					miXML.Close();
					Cursor = Cursors.Default;
					return;
				}
				miXML.WriteElementString("autorizacion", stAutFactura.ToString());
				//fecha de caducidad de factura
				DateTime dtFechaCaducidad;
				if (dtFechaCompra.Year > 2000) dtFechaCaducidad = dtFechaCompra;
				else dtFechaCaducidad = dtClientCaduca;
				miXML.WriteElementString("fechaCaducidad", dtFechaCaducidad.ToString("MM/yyyy"));
				#endregion Autorizacion y caducidad
				#region IVA
				miXML.WriteElementString("baseImponible", dSubtotalExcento.ToString("0.00", us));
				miXML.WriteElementString("baseImpGrav", dSubtotalIva.ToString("0.00", us));
				miXML.WriteElementString("porcentajeIva", "2");
				miXML.WriteElementString("montoIva", dCompraIva.ToString("0.00", us));
				#endregion Fin de IVA
				#region ICE
				miXML.WriteElementString("baseImpIce", "0.00");
				miXML.WriteElementString("porcentajeIce", "0");
				miXML.WriteElementString("montoIce", "0.00");
				#endregion
				#region Iva en Bienes
				//montoIvaBienes
				// Calculo de Iva de Servicios para calcular por diferencia el Iva de Bienes
				double dIvaBienes = dCompraIva - dIvaServicio;
				miXML.WriteElementString("montoIvaBienes", dIvaBienes.ToString("0.00", us));
				// Se asume que el bien siempre se lo pone en la primera retencion
				// Valores validos 0, 30, 100
				miXML.WriteElementString("porRetBienes", stCodRetIva);
				miXML.WriteElementString("valorRetBienes", dRetValorIva.ToString("0.00", us));
				#endregion Iva en Bienes
				#region Iva en Servicios
				miXML.WriteElementString("montoIvaServicios", dIvaServicio.ToString("0.00", us));
				//porRetServicios si existen 2 se asume que es el segundo sino el primero
				double dPorcServ = dRetIvaPorc1;
				if (dPorcServ <= 0.001 && dIvaServicio > 0) dPorcServ = dRetIvaPorc;
				miXML.WriteElementString("porRetServicios", stCodRetIva1);
				//valorRetServicios
				double dvalorRetServicios = dRetValorIva1;
				if (dvalorRetServicios <= 0.001 && dIvaServicio > 0) dvalorRetServicios = dRetValorIva;
				miXML.WriteElementString("valorRetServicios", dvalorRetServicios.ToString("0.00", us));
				#endregion Iva en servicios
				#region Impuesto a la Renta
				miXML.WriteStartElement("air");
				for (int i=0; i<iCuentaRet; i++)
				{
					miXML.WriteStartElement("detalleAir");
					double dRentaRet = miReader.GetDouble(23);
					double dRentaBase = miReader.GetDouble(24);
					double dRentaPorc = miReader.GetDouble(25);
					string stCodRenta = miReader.GetString(26);
					miXML.WriteElementString("codRetAir", stCodRenta);
					miXML.WriteElementString("baseImpAir", dRentaBase.ToString("0.00", us));
					miXML.WriteElementString("porcentajeAir", dRentaPorc.ToString("0.00", us));
					miXML.WriteElementString("valRetAir", dRentaRet.ToString("0.00", us));
					miXML.WriteEndElement(); // Fin de air
					if (i < iCuentaRet - 1)	miReader.Read();
				}
				miXML.WriteEndElement(); // Fin de air
				#endregion Impuesto a la renta
				#region Codigos de Retencion
				if (iCuentaRet > 0)
				{
					string stSerieRet = miReader.GetString(27);
					stSerieRet = Digito(stSerieRet);
					string stAutRet = miReader.GetString(28);
					stAutRet = Digito(stAutRet);
					string stCodRet = miReader.GetString(29);
					stCodRet = Digito(stCodRet);
					DateTime dtFechaRet = miReader.GetDateTime(30);
					if (stSerieRet.Length != 6 || stAutRet.Length < 4 || stCodRet.Length < 1)
					{
						MessageBox.Show("Error en Serie o Numero de Retencion", "Información");
						miReader.Close();
//						using (EgresoAT miEgreso = new EgresoAT())
//						{
//							miEgreso.IdFactura = IdCompra;
//							miEgreso.ShowDialog();
//						}
						miXML.Close();
						Cursor = Cursors.Default;
						return;
					}
					miXML.WriteElementString("estabRetencion1", stSerieRet.Substring(0,3));
					miXML.WriteElementString("ptoEmiRetencion1", stSerieRet.Substring(3,3));
					miXML.WriteElementString("secRetencion1", stCodRet);
					miXML.WriteElementString("autRetencion1", stAutRet);
					miXML.WriteElementString("fechaEmiRet1", dtFechaRet.ToString("dd/MM/yyyy"));
				}
				#endregion Codigos de Retencion

				#region Devoluciones de compras
				if (IdTipoFactura == 6)
				{
					try
					{
						miXML.WriteElementString("docModificado", int.Parse(stDevCompTipo[k]).ToString());
						//					miXML.WriteElementString("docModificado", stDevCompTipo[k]);
						miXML.WriteElementString("fechaEmiModificado", dtDevFecha[k].ToString("dd/MM/yyyy"));
						miXML.WriteElementString("estabModificado", stDevSerie[k].Substring(0,3));
						miXML.WriteElementString("ptoEmiModificado", stDevSerie[k].Substring(3,3));
						miXML.WriteElementString("secModificado", stDevComprob[k]);
						miXML.WriteElementString("autModificado", stDevAutoriza[k]);
						//				miXML.WriteElementString("contratoPartidoPolitico", "");
						miXML.WriteElementString("montoTituloOneroso", "0.00");
						miXML.WriteElementString("montoTituloGratuito", "0.00");
					}
					catch (System.Exception ex)
					{
						MessageBox.Show(ex.Message, "No existe comprobante");
//						new Compra(6, IdCompra).Show();
						return;
					}
					k++;
				}
				else
				{
					miXML.WriteElementString("docModificado", "0");
					miXML.WriteElementString("fechaEmiModificado", "00/00/0000");
					miXML.WriteElementString("estabModificado", "000");
					miXML.WriteElementString("ptoEmiModificado", "000");
					miXML.WriteElementString("secModificado","0000000");
					miXML.WriteElementString("autModificado", "0000000000");
					//				miXML.WriteElementString("contratoPartidoPolitico", "");
					miXML.WriteElementString("montoTituloOneroso", "0.00");
					miXML.WriteElementString("montoTituloGratuito", "0.00");
				}
				#endregion fin de devoluciones de compras
				miXML.WriteEndElement(); // Fin de DetalleCompras
			}
			miReader.Close();
			miXML.WriteEndElement(); // Fin de Compras

			#endregion Fin de Compras

			#region importaciones
			miXML.WriteStartElement("importaciones");
			stSelect = "SELECT CompraCredTribut.Codigo, Compra.DerechoDevolucion, "
				+ "Cliente.idTipoRuc, Cliente.Ruc, Cliente.Nombre, "
				+ "CompraComprobante.Codigo, Compra.Numero, "
				+ "Compra.SerieFactura, Compra.AutFactura, "
				+ "Compra.FechaCaducidad, Compra.Fecha, "
				+ "Compra.SubTotalIva, Compra.SubTotalExcento, "
				+ "Compra.Iva, Compra.idCompra, "
				+ "Compra.Ice, Retencion.IvaRetenido, Retencion.AutorizaRet, "
				+ "Retencion.Fecha as FechaRet, Retencion.SerieRet, "
				+ "Retencion.IvaRetenido1, Retencion.idRetencion, "
				+ "Compra.Refrendo, Compra.Total "
				+ "FROM Compra Inner Join Cliente On Compra.idCliente = Cliente.idCliente "
				+ "LEFT JOIN Retencion ON Compra.idCompra = Retencion.idCompra "
				+ "LEFT OUTER JOIN CompraComprobante ON Compra.idComprobante = "
				+ "CompraComprobante.idComprobante "
				+ "LEFT OUTER JOIN CompraCredTribut ON Compra.idCredTributario = "
				+ "CompraCredTribut.idCredTributario Where Compra.idTipoFactura = 4 "
				+ "And Compra.Borrar = 0 AND Compra.idComprobante = 14";
				//SELECT RetDetalle
//SET ORDER TO IDBN_DET	
//
//SELECT miCompra
//SCAN
//	stTexto = stTexto + "<detalleImportaciones>" + stEnter
//	IF LEN(ALLTRIM(miCompra.Credito_Trib)) <= 1 then
//		MESSAGEBOX("Codigo de Sustento de Credito Tributario")
//		lcDir = IIF(glDepura, "C:\Bodega\", "") + "Compras"
//		DO FORM &lcDir WITH tnTipoFact, -miCompra.id_Compras
//		RETURN 
//	ENDIF 
//	WriteElementString(@stTexto, "codSustento", miCompra.Credito_Trib)
//	WriteElementString(@stTexto, "importacionDe", miCompra.Trans_Prest) && bienes o servicios
//	WriteElementString(@stTexto, "fechaLiquidacion", DTOC(miCompra.Fecha))
//	WriteElementString(@stTexto, "tipoComprobante", val(miCompra.Comp_Tipo))
//	
//	*Refrendo
//	IF (LEN(ALLTRIM(miCompra.Refrendo)) != 16) THEN 
//		MESSAGEBOX("Error en refrendo longitud debe ser de 16")
//		lcDir = IIF(glDepura, "C:\Bodega\", "") + "Compras"
//		DO FORM &lcDir WITH tnTipoFact, -miCompra.id_Compras		
//		RETURN
//	ENDIF 
//	WriteElementString(@stTexto, "distAduanero", SUBSTR(miCompra.Refrendo, 1, 3))
//	WriteElementString(@stTexto, "anio", SUBSTR(miCompra.Refrendo, 4, 4))
//	WriteElementString(@stTexto, "regimen", SUBSTR(miCompra.Refrendo, 8, 2))
//	WriteElementString(@stTexto, "correlativo", SUBSTR(miCompra.Refrendo, 10, 6))
//	WriteElementString(@stTexto, "verificador", SUBSTR(miCompra.Refrendo, 16, 1))
//
//	WriteElementString(@stTexto, "idFiscalProv", miCompra.Ruc)
//	WriteElementString(@stTexto, "valorCIF", miCompra.Total)
//	WriteElementString(@stTexto, "razonSocialProv", miCompra.Desc_Dist)
//	WriteElementString(@stTexto, "tipoSujeto", 2) && Se asume que es juridica en caso de error cambiarlo manualmente
//*	WriteElementString(@stTexto, "valorCIF", miCompra.Total)
//
//	WriteElementString(@stTexto, "baseImponible", STR(miCompra.SubTotal1, 10, 2))
//	WriteElementString(@stTexto, "baseImpGrav", STR(miCompra.SubTotal2, 10, 2))
//	WriteElementString(@stTexto, "porcentajeIva", "2")
//	WriteElementString(@stTexto, "montoIva", STR(miCompra.Iva, 10, 2))
//
//** Calculo de ICE
//	lnBaseIce = 0
//	IF miCompra.Ice > 0 THEN 
//		lnBaseIce = miCompra.Ice / 100.0
//	ENDIF 
//	WriteElementString(@stTexto, "baseImpIce", STR(lnBaseIce, 10, 2))
//	WriteElementString(@stTexto, "porcentajeIce", "0") &&miCompra.Cod_Ice)
//	WriteElementString(@stTexto, "montoIce", STR(miCompra.Ice, 10, 2))
//	
//	stTexto = stTexto + "<air>" + stEnter
//	lnIdBnDet = 0
//	IF NOT ISNULL(miCompra.id_BN_DETALLE) THEN 
//		lnIdBnDet = miCompra.id_BN_DETALLE
//	ENDIF 
//	SELECT RetDetalle
//	SEEK lnIdBnDet
//	IF FOUND() then
//		SCAN WHILE RetDetalle.idBn_Det = miCompra.id_Bn_Detalle
//			IF retDetalle.Retenido = 0 then
//				loop 
//			ENDIF
//			SELECT ArticuloSRI
//			LOCATE FOR ArticuloSRI.idArtSri = RetDetalle.idRetRenta
//			IF NOT FOUND() then
//				loop
//			ENDIF 
//			stTexto = stTexto + "<detalleAir>" + stEnter
//			WriteElementString(@stTexto, "codRetAir", ArticuloSRI.Codigo)
//			WriteElementString(@stTexto, "baseImpAir", RetDetalle.Base)
//			WriteElementString(@stTexto, "porcentajeAir", ArticuloSRI.Porcentaje)
//			WriteElementString(@stTexto, "valRetAir", RetDetalle.Retenido)
//			stTexto = stTexto + "</detalleAir>" + stEnter
//		ENDSCAN 
//	ENDIF 
//	stTexto = stTexto + "</air>" + stEnter
//
//	stTexto = stTexto + "</detalleImportaciones>" + stEnter
//	SELECT miCompra
//ENDSCAN
			miXML.WriteEndElement(); // Fin de importaciones
			#endregion importaciones

			//Write the XML to file and close the myXmlTextWriter
			miXML.WriteEndElement(); // Fin de IVA
			miXML.Flush();
			miXML.Close();
			Cursor = Cursors.Default;
			MessageBox.Show("Archivo Generado", "Información");
		}

		private string Digito(string stNumero)
		{
			string stDigito = "";
			string EsNumero = "0123456789";
			for (int i=0; i< stNumero.Length; i++)
			{
				string cCar = stNumero.Substring(i, 1);
				if (EsNumero.IndexOf(cCar) >= 0)
					stDigito += cCar;
			}
			return stDigito;
		}

		CultureInfo us = new CultureInfo("en-US");

		SeteoRow_tableView drSeteo;
		C1.Data.C1DataRow drSeteoC;
		bool bIncluyeIva = false;

		private void AnexoTransaccional_Load(object sender, System.EventArgs e)
		{
			cmbDesde.Value = DateTime.Today;
			cmbHasta.Value = DateTime.Today;
			bIncluyeIva = Funcion.bEjecutaSQL(cdsCompra, "Select IsNull(IncluyeIva, 0) From SeteoF");

			cmbFecha.Value = DateTime.Today.AddMonths(-1);
			string stAudita = "Exec AuditaCrear 53, 6, 'Formularios'";
			Funcion.EjecutaSQL(cdsCompra, stAudita, true);
			C1DataRow dr = this.cdsSeteo.TableViews[0].Rows[0];
			drSeteo = SeteoRow_tableView.Obj(dr);
//			try
//			{
//				C1DataRow drC = this.cdsCompra.TableViews[0].Rows[0];
//				drCompra = CompraRow_tableView.Obj(drC);
//			}
//			catch{};
			drSeteoC = this.cdsSeteoC.TableViews[0].Rows[0];
		}

		private void cdsSeteo_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			cdsSeteo.Schema.Connections[0].ConnectionString = 
				miFuncion.AccesoBase(MenuLatinium.stDirConta);
		}

		private void cdsCompra_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			this.cdsCompra.Schema.Connections[0].ConnectionString = 
				miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);

//			C1.Data.SchemaObjects.DataSetDef dataSetDef = e.DataSet.Schema.DataSetDefs["dsCompra"];
//
//			e.Filter.Add(new C1.Data.FilterCondition(dataSetDef.TableViews["Compra"],
//				"idCompra < 10"));
		}

		private XmlValidatingReader myXmlValidatingReader = null;
//		private XmlTextReader myXmlTextReader = null;
		private Boolean Success = true;

		private string FormaCampo(int iCampo, string stValor)
		{
			string stCampo = "<campo numero=\"" + iCampo.ToString() +"\">"
				+ stValor + "</campo>";
			return stCampo;
		}
		private void btValidar_Click(object sender, System.EventArgs e)
		{
			#region Encabezado
			DateTime dtFecha = (DateTime) cmbFecha.Value;
			StreamWriter sr = null;
			string stArchivo = "";
			try
			{
				string stDir = @"C:\XML_DECLARACIONES\";
				if (!Directory.Exists(stDir)) stDir = "";
				stArchivo = stDir + "103ORI_" + dtFecha.ToString("MMMyyyy") + ".XML";
				sr= File.CreateText(stArchivo);
			}
			catch
			{
				MessageBox.Show("Esta abierto el archivo Formulario.xml");
				Cursor = Cursors.Default;
				return;
			}
			sr.WriteLine("<?xml version=\"1.0\" encoding=\"UTF-8\"?>");
			sr.WriteLine("<formulario version=\"0.2\">");
			sr.WriteLine("<cabecera>");
			string stVersionFor = "03201202";
			sr.WriteLine("<codigo_version_formulario>" + stVersionFor +  "</codigo_version_formulario>");
			sr.WriteLine("<ruc>"+drSeteo.Ruc.Trim()+"</ruc>");
			sr.WriteLine("<codigo_moneda>1</codigo_moneda>");
      sr.WriteLine("</cabecera>");
			sr.WriteLine("<detalle>");

			sr.WriteLine(FormaCampo(31, "O"));
			sr.WriteLine(FormaCampo(101, dtFecha.Month.ToString()));
			sr.WriteLine(FormaCampo(102, dtFecha.Year.ToString()));
			sr.WriteLine(FormaCampo(104, ""));
			sr.WriteLine(FormaCampo(198, drSeteo.CedulaRep.Trim()));
			sr.WriteLine(FormaCampo(199, drSeteo.CedulaContador.Trim() + "001"));
			sr.WriteLine(FormaCampo(201, drSeteo.Ruc.Trim()));
			sr.WriteLine(FormaCampo(202, "<![CDATA[" + drSeteo.Empresa.Trim() + "]]>"));
//			sr.WriteLine(FormaCampo(301, "0"));
			sr.WriteLine(FormaCampo(302, "0.00"));
			#endregion Encabezado

			#region Codigos del 303 al 349
			string stSelect = "Select Sum(IsNull(RetDetalle.Base, 0)), Substring(LTrim(IsNull(ArticuloSRI.Codigo, '0')), 1, 3), "
				+ "Sum(IsNull(RetDetalle.Retenido, 0)) FROM   (((((Retencion Retencion "
				+ "LEFT OUTER JOIN RetDetalle RetDetalle ON Retencion.idRetencion=RetDetalle.idRetencion) "
				+ "LEFT OUTER JOIN Cliente Cliente ON Retencion.idCliente=Cliente.idCliente) "
				+ "LEFT OUTER JOIN Asiento Asiento ON Retencion.idAsiento=Asiento.idAsiento) "
				+ "LEFT OUTER JOIN Compra Compra ON Retencion.idCompra=Compra.idCompra) "
				+ "LEFT OUTER JOIN Asiento Asiento_1 ON Compra.idAsiento=Asiento_1.idAsiento) "
				+ "LEFT OUTER JOIN ArticuloSRI ArticuloSRI ON RetDetalle.idRetRenta=ArticuloSRI.idArticuloSRI "
				+ "Where Retencion.Fecha >= '" + dtFecha.ToString("yyyyMM")
				+ "01' And Retencion.Fecha < '" + dtFecha.AddMonths(1).ToString("yyyyMM") + "01' "
				+ "And Compra.idTipoFactura = 4 And Compra.Borrar = 0 And Retencion.NoAplica = 0 "
				+ "Group by Substring(LTrim(IsNull(ArticuloSRI.Codigo, '0')), 1, 3)";

			SqlDataReader miReader = Funcion.rEscalarSQL(cdsCompra, stSelect, true);
			double[] dValor = new double[400];
			double[] dPorcentaje = new double[400];
			for (int i=0; i<400; i++)
			{
				dValor[i] = 0;
				dPorcentaje[i] = 0;
			}
			double dAcumula = 0;
			double dAcumula3 = 0;
			double dAcumula4 = 0;
			double dSubtActumula = 0;
			while (miReader.Read())
			{
				int iCodigo = int.Parse(miReader.GetString(1));
				if (iCodigo < 300) continue;
				int j = iCodigo-300;
				dValor[j] = miReader.GetDouble(0);
				dPorcentaje[j] = miReader.GetDouble(2);// * dValor[j] / 100.0;
				if (j < 100)
				{
					dAcumula3 += Math.Round(dPorcentaje[j], 2);
					dSubtActumula += Math.Round(dValor[j], 2);
				}
				else
					dAcumula4 += Math.Round(dPorcentaje[j], 2);
			}
			miReader.Close();
			dAcumula = dAcumula3 + dAcumula4;

			dValor[32] += dValor[33] + dValor[34] + dValor[35] + dValor[36] + + dValor[37];
			for (int i=3; i<=44; i++)
			{
				if (i == 3 || i == 4 || i == 7 || i == 8 || i == 9 || i == 10 || i == 12 || i == 19
					|| i == 20 || i == 22 || i == 23 || i == 24 || i == 25 || i== 27 || i == 28 || i== 29 || i == 30 || i == 32 
					|| i == 40 || i == 41 || i == 42 || i == 43 || i == 44)
					sr.WriteLine(FormaCampo(i + 300, dValor[i].ToString("0.00", us)));
			}
			if (dtFecha.Year>=2009) 
				sr.WriteLine(FormaCampo(349, dSubtActumula.ToString("0.00", us)));

			sr.WriteLine(FormaCampo(352, "0.00"));

			for (int i=3; i<=44; i++)
			{
				if (i == 3 || i == 4 || i == 7 || i == 8 || i == 9 || i == 10 || i == 12 || i == 19
					|| i == 20 || i == 22 || i == 23 || i == 24 || i == 25 || i== 27 || i == 28 || i == 29 || i == 30
					|| i == 40 || i == 41 || i == 42 || i == 43  || i == 44)
						sr.WriteLine(FormaCampo(i + 350, dPorcentaje[i].ToString("0.00", us)));
			}
			sr.WriteLine(FormaCampo(399, dAcumula3.ToString("0.00", us)));

			double d429 = 0;
			for (int i=101; i<128; i++)
			{
				if ( i == 101 || i == 103 || i == 105 || i == 107 || i == 121 || i == 127)
				{
					d429 += dValor[i];
					sr.WriteLine(FormaCampo(i + 300, dValor[i].ToString("0.00", us)));
				}
			}
			sr.WriteLine(FormaCampo(429, d429.ToString("0.00", us)));

			for (int i=101; i<126; i++)
			{
				if ( i == 101 || i == 103 || i == 105 || i == 107 || i == 121)
					sr.WriteLine(FormaCampo(i + 350, dPorcentaje[i].ToString("0.00", us)));
			}
			sr.WriteLine(FormaCampo(498, dAcumula4.ToString("0.00", us)));
			sr.WriteLine(FormaCampo(499, dAcumula.ToString("0.00", us)));
			#endregion Codigos del 303 al 333

			#region Comprobantes de Retencion Emitidos

			stSelect = "Select Count(Distinct(codigo)) From Retencion "
				+ "Where Fecha >= '" + dtFecha.ToString("yyyyMM")
				+ "01' And Fecha < '" + dtFecha.AddMonths(1).ToString("yyyyMM") 
				+ "01' And IngresoEgreso = 2 And NoAplica = 0 And "
				+ "(IsNull(RentaRetenida, 0) + IsNull(IvaRetenido, 0) + IsNull(IvaRetenido1, 0)) >= 0.01";
			int iCuenta = Funcion.iEscalarSQL(cdsCompra, stSelect, true);
			sr.WriteLine(FormaCampo(510, ""));
			sr.WriteLine(FormaCampo(520, ""));
			
			sr.WriteLine(FormaCampo(880, "0.00"));
			sr.WriteLine(FormaCampo(890, "0.00"));
			sr.WriteLine(FormaCampo(897, "0.00"));
			sr.WriteLine(FormaCampo(898, "0.00"));
			sr.WriteLine(FormaCampo(899, "0.00"));

			//sr.WriteLine(FormaCampo(108, iCuenta.ToString()));
//			sr.WriteLine(FormaCampo(499, dAcumula.ToString("0.00", us)));
			sr.WriteLine(FormaCampo(902, dAcumula.ToString("0.00", us)));
			sr.WriteLine(FormaCampo(903, "0.00"));
			sr.WriteLine(FormaCampo(904, "0.00"));
			sr.WriteLine(FormaCampo(905, dAcumula.ToString("0.00", us)));
			sr.WriteLine(FormaCampo(907, "0.00"));
			sr.WriteLine(FormaCampo(908, ""));
			sr.WriteLine(FormaCampo(909, "0.00"));
			sr.WriteLine(FormaCampo(910, ""));
			sr.WriteLine(FormaCampo(911, "0.00"));
			sr.WriteLine(FormaCampo(912, ""));
			sr.WriteLine(FormaCampo(913, "0.00"));
//			sr.WriteLine(FormaCampo(914, ""));
			sr.WriteLine(FormaCampo(915, "0.00"));
			sr.WriteLine(FormaCampo(922, "89"));
			sr.WriteLine(FormaCampo(999, dAcumula.ToString("0.00", us)));
			#endregion Comprobantes de Retencion Emitidos
			sr.WriteLine("</detalle>");
			sr.WriteLine("</formulario>");
			sr.Close();
			MessageBox.Show("Archivo " + stArchivo + " Generado");
		}
		private void Validar()
		{
			try
			{
				// Set the validation event handler
				myXmlValidatingReader.ValidationEventHandler += new ValidationEventHandler (this.ValidationEventHandle);

				// Read XML data
				while (myXmlValidatingReader.Read()){}
				Console.WriteLine ("Validacin finalizada. Validacin {0}", (Success==true ? "correcto" : "error"));
			}
			catch (XmlException e)
			{
				Console.WriteLine ("Excepción XmlException: " + e.ToString());
			}

			catch (Exception e)
			{
				Console.WriteLine ("Excepción: " + e.ToString());
			}
		}

		public void ValidationEventHandle (object sender, ValidationEventArgs args)
		{
			Success = false;

			Console.WriteLine("\tError de validacin: " + args.Message);

			if (args.Severity == XmlSeverityType.Warning)
			{
				Console.WriteLine("No se han encontrado esquemas para forzar la validacin.");
			} 
			else
				if (args.Severity == XmlSeverityType.Error)
			{
				Console.WriteLine("Error de validacin al validar el documento de instancia.");
			} 

			if (args.Exception != null) // XSD schema validation error
			{
				Console.WriteLine(args.Exception.SourceUri + "," +  args.Exception.LinePosition + "," +  args.Exception.LineNumber);
			}

			//if (myXmlValidatingReader.Reader.LineNumber > 0)
			//{
			//    Console.WriteLine("Line: "+ myXmlValidatingReader.Reader.LineNumber + " Position: " + myXmlValidatingReader.Reader.LinePosition);
			//}
		}

		private double ValorSelect(int IdTipo, string stNot, string stIn, string stOperador, int iGrupo, 
			string stCredTribut, int InstPublica)
		{
			string stCliente = "";
			if (InstPublica == 1)
				stCliente = " And IsNull(Cliente.InstPublica, 0) = 1";
			if (InstPublica == 0)
				stCliente = " And IsNull(Cliente.InstPublica, 0) = 0";
			// Ambos if (InstPublica == 2) 

			string stCred = " In (2, 7) ";
			if (stCredTribut == "SC") stCredTribut = " And Compra.idCredTributario " + stCred;
			if (stCredTribut == "C") stCredTribut = " And Compra.idCredTributario Not " + stCred;
			if (stCredTribut == "RG") stCredTribut = " And Compra.idCredTributario In (8) ";
			if (stCredTribut == "NI") stCredTribut = " And Compra.idCredTributario In (11) ";
			if (stCredTribut == "") stCredTribut = " And Compra.idCredTributario Not In (11) ";

			DateTime dtFecha = (DateTime) cmbFecha.Value;
			string stCompraDesc = "(1 - IsNull(Compra.DescPorcIva, 0)/100.0)";
			if (stOperador.Trim().StartsWith("="))
				stCompraDesc = "(1 - IsNull(Compra.DescPorcIva0, 0)/100.0)";

			string stFactor = "1.0 ";
			decimal dIVA = Funcion.decEscalarSQL(cdsSeteoF, "Select dbo.RetornaImpuestoIVAVigenteDecimal(GETDATE(), 1)");
			if (bIncluyeIva == true && (IdTipo == 1||IdTipo == 5)) stFactor = "1.0/1.15 ";
			if (Funcion.bEjecutaSQL(cdsCompra, "Exec SeteoGExiste 'FPIIVA'")) 				
			{
			       stFactor = "1.0 ";
			}

			string stSelect = "Select IsNull(sum((detcompra.cantidad* detcompra.precio "
				+ "- detcompra.descuentoart) * Case When detcompra.Impuesto > 0 then " + stFactor
				+ " else 1 end * "
				+ stCompraDesc + "), 0) from detcompra "
				+ "inner join compra on compra.idcompra = detcompra.idcompra "
				+ "inner join Cliente on Cliente.idCliente = Compra.idCliente "
				+ "inner join Articulo on Articulo.idArticulo = detcompra.idArticulo "
				+ "Inner Join CompraComprobante On Compra.idComprobante = "
				+ "CompraComprobante.idComprobante Where Compra.idTipoFactura = " + IdTipo.ToString()
				+ " And Compra.Borrar = 0 And Compra.Fecha >= '" + dtFecha.ToString("yyyyMM") 
				+ "01' And " + stNot + " CompraComprobante.Codigo In (" + stIn + ") "
				+ "And Compra.Fecha < '" + dtFecha.AddMonths(1).ToString("yyyyMM") + "01' "
				+ "And DetCompra.Impuesto " + stOperador + stCredTribut;
			if (iGrupo > 0)
				stSelect += " And Articulo.idTipoGrupo = " + iGrupo.ToString();
//			if (InstPublica > 0)
				stSelect += stCliente;
			//			else // Necesario si no se resta el valor de servicion y activos al de bienes
//				stSelect += " And Articulo.idTipoGrupo Not In (3, 5) ";
			return Funcion.dEscalarSQL(cdsCompra, stSelect, true);
		}

		private double dServicioCompra(DateTime dtFecha)
		{
			string stSelect = "Select IsNull(Sum(IsNull(c.Servicio, 0)),0) From Compra c "
				+ "Where c.idTipoFactura = 4 And c.Borrar = 0 And c.Fecha >= '"
				+ dtFecha.ToString("yyyyMM") + "01' And c.Fecha < '" 
				+ dtFecha.AddMonths(1).ToString("yyyyMM") + "01'";
			double dServicio = Funcion.dEscalarSQL(cdsCompra, stSelect);
			return dServicio;
		}

		private int NumComprobante(int IdTipo, int iGrupo)
		{
			DateTime dtFecha = (DateTime) cmbFecha.Value;
			string stSelect = "Select count(*) from compra where idtipofactura = " 
				+ IdTipo.ToString()
				+ " And Compra.Fecha >= '" + dtFecha.ToString("yyyyMM") 
				+ "01' And Compra.Fecha < '" + dtFecha.AddMonths(1).ToString("yyyyMM") + "01' ";
			if (iGrupo > 0)
				stSelect += " And idComprobante = " + iGrupo.ToString();
			return Funcion.iEscalarSQL(cdsCompra, stSelect, true);
		}

		// formulario para el 2009
		private void btForm104109()
		{
			DateTime dtFecha = (DateTime) cmbFecha.Value;
			#region Actualizacion de descuento
			string stUpdate = string.Format("Exec DescuentoPorcAct '{0}'", dtFecha.ToString("yyyyMMdd"));
			Funcion.EjecutaSQL(cdsCompra, stUpdate, true);
			#endregion Fin de descuento
			#region Encabezado
			StreamWriter sr = null;
			string stArchivo = "";
			try
			{
				string stDir = @"C:\XML_DECLARACIONES\";
				if (!Directory.Exists(stDir)) stDir = "";
				stArchivo = @stDir + "104ORI_" + dtFecha.ToString("MMMyyyy") + ".XML";
				sr= File.CreateText(stArchivo);
			}
			catch
			{
				MessageBox.Show("Esta abierto el archivo Formulario.xml");
				Cursor = Cursors.Default;
			}
			sr.WriteLine("<?xml version=\"1.0\" encoding=\"UTF-8\"?>");
			sr.WriteLine("<formulario version=\"0.2\">");
			sr.WriteLine("<cabecera>");
			if (drSeteo.PNatural == false)
				sr.WriteLine("<codigo_version_formulario>04201201</codigo_version_formulario>");
			else
				sr.WriteLine("<codigo_version_formulario>04201201</codigo_version_formulario>");

			sr.WriteLine("<ruc>" + drSeteo.Ruc + "</ruc>");
			sr.WriteLine("<codigo_moneda>1</codigo_moneda>");
			sr.WriteLine("</cabecera>");
			sr.WriteLine("<detalle>");

			sr.WriteLine(FormaCampo(31, "O"));
			sr.WriteLine(FormaCampo(101, dtFecha.Month.ToString()));
			sr.WriteLine(FormaCampo(102, dtFecha.Year.ToString()));
			sr.WriteLine(FormaCampo(104, ""));
			sr.WriteLine(FormaCampo(198, drSeteo.CedulaRep));
			if (drSeteo.PNatural == false)
				sr.WriteLine(FormaCampo(199, drSeteo.CedulaContador + "001"));
			sr.WriteLine(FormaCampo(201, drSeteo.Ruc));
			sr.WriteLine(FormaCampo(202, "<![CDATA[" + drSeteo.Empresa + "]]>"));
			#endregion Encabezado

			#region Resumen de Ventas
			#region Ventas Locales 401 al 409
			string stFiltroComp = "'09', '15', '16', '17', '41', '98', '99'";
			double dVentasIva12 = ValorSelect(1, "Not", stFiltroComp, "> 0", 0, "", 0);
			double dVentasAFIva = ValorSelect(1, "Not", stFiltroComp, "> 0", 5, "", 0);
			dVentasIva12 -= dVentasAFIva;
			double dVentasIva0SC = ValorSelect(1, "Not", stFiltroComp, "= 0", 0, "SC", 2);
			double dVentasAF0SC = ValorSelect(1, "Not", stFiltroComp, "= 0", 5, "SC", 2);
      dVentasIva0SC -= dVentasAF0SC;
			double dVentasIva0Cr = ValorSelect(1, "Not", stFiltroComp, "= 0", 0, "C", 2);
			double dVentasAF0SCr = ValorSelect(1, "Not", stFiltroComp, "= 0", 5, "C", 2);
			dVentasIva0Cr -= dVentasAF0SCr;
			double dVentaExpIva = ValorSelect(1, "", "'16', '98'", "> 0", 0, "", 0);
			double dVentaExpIva0 = ValorSelect(1, "", "'16', '98'", "= 0", 0, "", 2);
			double dVentasExp = dVentaExpIva + dVentaExpIva0;
			//Venta de Servicios a Exportadores
			double dVentasExpS = ValorSelect(1, "", "'16', '98'", ">= 0", 3, "", 0);
			double dVentasExpB = dVentasExp - dVentasExpS;
			double dVentaExpImp = dVentaExpIva * .12;

			sr.WriteLine(FormaCampo(401, dVentasIva12.ToString("0.00", us)));
			sr.WriteLine(FormaCampo(402, dVentasAFIva.ToString("0.00", us)));

			if (Funcion.bEjecutaSQL(cdsCompra, "Exec SeteoGExiste 'FPAD'"))
			{
				string stExecD = String.Format("Exec AnexoDonacion '{0}', 1", dtFecha.ToString("yyyyMMdd"));
				double dExcento = Funcion.dEscalarSQL(cdsCompra, stExecD, true);
				sr.WriteLine(FormaCampo(403, dExcento.ToString("0.00", us)));
			}
			else
				sr.WriteLine(FormaCampo(403, dVentasIva0SC.ToString("0.00", us)));

			sr.WriteLine(FormaCampo(404, dVentasAF0SC.ToString("0.00", us)));
			if (Funcion.bEjecutaSQL(cdsCompra, "Exec SeteoGExiste 'FPAD'"))
				dVentasIva0Cr = 0;
			sr.WriteLine(FormaCampo(405, dVentasIva0Cr.ToString("0.00", us)));

			sr.WriteLine(FormaCampo(406, dVentasAF0SCr.ToString("0.00", us)));

			if (drSeteo.PNatural == false)
			{
				sr.WriteLine(FormaCampo(407, dVentasExpB.ToString("0.00", us)));
				sr.WriteLine(FormaCampo(408, dVentasExpS.ToString("0.00", us)));
			}
			else
			{
				dVentasExpB = 0;
				dVentasExpS = 0;
			}
			double dTotal409 = dVentasIva12 + dVentasAFIva + dVentasIva0SC + dVentasAF0SC
				+ dVentasIva0Cr + dVentasAF0SCr + dVentasExpB + dVentasExpS;
			sr.WriteLine(FormaCampo(409, dTotal409.ToString("0.00", us)));
			#endregion Ventas Locales 401 al 409
			#region Devolucion Ventas Locales 411 al 419
			double dN_C_Iva12 = ValorSelect(5, "Not", stFiltroComp, "> 0", 0, "", 0);
			double dN_C_AFIva = ValorSelect(5, "Not", stFiltroComp, "> 0", 5, "", 0);
			dN_C_Iva12 = dVentasIva12 - dN_C_Iva12 + dN_C_AFIva;
			dN_C_AFIva = dVentasAFIva - dN_C_AFIva;

			double dN_C_AFIva0SC = ValorSelect(5, "Not", stFiltroComp, "= 0", 0, "SC", 0);
			double dN_C_AF0SC = ValorSelect(5, "Not", stFiltroComp, "= 0", 5, "SC", 0);
			dN_C_AFIva0SC = dVentasIva0SC - dN_C_AFIva0SC + dN_C_AF0SC;
			dN_C_AF0SC = dVentasAF0SC - dN_C_AF0SC;

			double dN_C_Iva0Cr = ValorSelect(5, "Not", stFiltroComp, "= 0", 0, "C", 0);
			double dN_CAF0SCr = ValorSelect(5, "Not", stFiltroComp, "= 0", 5, "C", 0);
			dN_C_Iva0Cr = dVentasIva0Cr - dN_C_Iva0Cr + dN_CAF0SCr;
//			dVentasAF0SC = dVentasIva0Cr - dVentasAF0SC;
			dN_CAF0SCr = dVentasAF0SCr - dN_CAF0SCr;

			double dN_C_ExpIva = ValorSelect(5, "", "'16', '98'", "> 0", 0, "", 0);
			double dN_C_ExpIva0 = ValorSelect(5, "", "'16', '98'", "= 0", 0, "", 0);
			double dN_C_Exp = dN_C_ExpIva + dN_C_ExpIva0;
			//Venta de Servicios a Exportadores
			double dN_C_ExpS = ValorSelect(5, "", "'16', '98'", ">= 0", 3, "", 0);
			double dN_C_ExpB = dVentasExpB - dN_C_Exp + dN_C_ExpS;
			dN_C_ExpS = dVentasExpS - dN_C_ExpS;

			double dN_C_ExpImp = dN_C_ExpIva * .12;

			sr.WriteLine(FormaCampo(411, dN_C_Iva12.ToString("0.00", us)));
			sr.WriteLine(FormaCampo(412, dN_C_AFIva.ToString("0.00", us)));
			if (Funcion.bEjecutaSQL(cdsCompra, "Exec SeteoGExiste 'FPAD'"))
			{
				string stExecD = String.Format("Exec AnexoDonacion '{0}', 1", dtFecha.ToString("yyyyMMdd"));
				double dExcento = Funcion.dEscalarSQL(cdsCompra, stExecD, true);
//				dExcento = dExcento * .12;
				sr.WriteLine(FormaCampo(413, dExcento.ToString("0.00", us)));
			}
			else
				sr.WriteLine(FormaCampo(413, dN_C_AFIva0SC.ToString("0.00", us)));
			sr.WriteLine(FormaCampo(414, dN_C_AF0SC.ToString("0.00", us)));
			sr.WriteLine(FormaCampo(415, dN_C_Iva0Cr.ToString("0.00", us)));
			sr.WriteLine(FormaCampo(416, dN_CAF0SCr.ToString("0.00", us)));

			if (drSeteo.PNatural == false)
			{
				sr.WriteLine(FormaCampo(417, dN_C_ExpB.ToString("0.00", us)));
				sr.WriteLine(FormaCampo(418, dN_C_ExpS.ToString("0.00", us)));
			}
			else
			{
				dN_C_ExpB = 0;
				dN_C_ExpS = 0;
			}
			double dTotal419 = dN_C_Iva12 + dN_C_AFIva + dN_C_AFIva0SC + dN_C_AF0SC
				+ dN_C_Iva0Cr + dN_CAF0SCr + dN_C_ExpB + dN_C_ExpS;
			sr.WriteLine(FormaCampo(419, dTotal419.ToString("0.00", us)));
			#endregion Ventas Locales 411 al 419
			#region Resumen de Impuestos 421 422 429 y 431 - 434 y 443 444
			double dVentasIvaImp = dN_C_Iva12 * .12;
			double dVentasA_FImp = dN_C_AFIva * .12;
			double dTotal429 = dVentasIvaImp + dVentasA_FImp;
			sr.WriteLine(FormaCampo(421, dVentasIvaImp.ToString("0.00", us)));
			sr.WriteLine(FormaCampo(422, dVentasA_FImp.ToString("0.00", us)));
			sr.WriteLine(FormaCampo(429, dTotal429.ToString("0.00", us)));

			double dDonacion = 0;
			if (Funcion.bEjecutaSQL(cdsCompra, "Exec SeteoGExiste 'FPAD'"))
			{
				string stExecD = String.Format("Exec AnexoDonacion '{0}', 2", dtFecha.ToString("yyyyMMdd"));
				dDonacion = Funcion.dEscalarSQL(cdsCompra, stExecD);
			}

			sr.WriteLine(FormaCampo(431, dDonacion.ToString("0.00", us)));
			sr.WriteLine(FormaCampo(432, "0.00"));
			sr.WriteLine(FormaCampo(433, "0.00"));
			//Reembolso de gastos
			double dVentasReembolso0 = ValorSelect(1, "", "'41'", "= 0", 0, "", 0);
			double dVentasReembolsoIva = ValorSelect(1, "", "'41'", "> 0", 0, "", 0);
//			double dVentasReembolsoImp = dVentasReembolsoIva * .12;
			dVentasReembolsoIva += dVentasReembolso0;
			double dIVA = Funcion.dEscalarSQL(cdsSeteoF, "Select dbo.RetornaImpuestoIVAVigenteDecimal(GETDATE(), 1)");
			double dVentasReembolsoImp = dVentasReembolsoIva - dVentasReembolsoIva / dIVA;
			dVentasReembolsoIva = dVentasReembolsoIva / dIVA;

			sr.WriteLine(FormaCampo(434, dVentasReembolsoIva.ToString("0.00", us)));
			sr.WriteLine(FormaCampo(443, "0.00"));
			sr.WriteLine(FormaCampo(444, dVentasReembolsoImp.ToString("0.00", us)));
			#endregion Resumen de Impuestos 421 422 429
			#region Contado y Credito 480 481 
			string stContado = "Exec Formulario480 1, '" + dtFecha.ToString("yyyyMM")
				+ "01', '" + dtFecha.AddMonths(1).ToString("yyyyMM") + "01'";
			double dContadoVenta480 = Funcion.dEscalarSQL(cdsCompra, stContado, false);
			string stCredito = "Exec Formulario480 0, '" + dtFecha.ToString("yyyyMM")
				+ "01', '" + dtFecha.AddMonths(1).ToString("yyyyMM") + "01'";
			double dCreditoVenta481 = Funcion.dEscalarSQL(cdsCompra, stCredito, false);
			sr.WriteLine(FormaCampo(480, dContadoVenta480.ToString("0.00", us)));
			sr.WriteLine(FormaCampo(481, dCreditoVenta481.ToString("0.00", us)));
			sr.WriteLine(FormaCampo(482, dTotal429.ToString("0.00", us))); 
			sr.WriteLine(FormaCampo(483, "0.00"));
			double dImpEsteMes = dContadoVenta480 * 0.12;
			double dImpProximoMes = 0;//dTotal429 - dImpEsteMes; Kasandra Credimetica
			sr.WriteLine(FormaCampo(484, dImpEsteMes.ToString("0.00", us))); // Contado del mes
			sr.WriteLine(FormaCampo(485, dImpProximoMes.ToString("0.00", us))); // 
			double dValor499 = dImpEsteMes ;
			sr.WriteLine(FormaCampo(499, dValor499.ToString("0.00", us))); // Impuesto a Liquidar este mes
			#endregion ContadoCredito
			#endregion Resumen de Ventas
			#region Resumen de Compras
			#region Compras Locales 501 al 509
			stFiltroComp = "'09', '15', '16', '17', '41', '98', '99'";
			double dComprasIva12 = ValorSelect(4, "Not", stFiltroComp, "> 0", 0, "C", 0);
			double dComprasAFIva = ValorSelect(4, "Not", stFiltroComp, "> 0", 5, "C", 0);
			dComprasIva12 -= dComprasAFIva;
			double dComprasIva12SC = ValorSelect(4, "Not", stFiltroComp, "> 0", 0, "SC", 0);

			double dCompraImpBIva = ValorSelect(4, "", "'15', '17'", "> 0", 0, "", 0);
			double dCompraImpAFIva = ValorSelect(4, "", "'15', '17'", "> 0", 5, "", 0);
			dCompraImpBIva -= dCompraImpAFIva;
			double dCompraImpIva0 = ValorSelect(4, "", "'15', '17'", "= 0", 0, "", 0);

			double dComprasIva0 = ValorSelect(4, "Not", stFiltroComp, "= 0", 0, "", 0);
      
			sr.WriteLine(FormaCampo(501, dComprasIva12.ToString("0.00", us)));
			sr.WriteLine(FormaCampo(502, dComprasAFIva.ToString("0.00", us)));
			sr.WriteLine(FormaCampo(503, dComprasIva12SC.ToString("0.00", us)));
			
			if (drSeteo.PNatural == false)
			{
				sr.WriteLine(FormaCampo(504, dCompraImpBIva.ToString("0.00", us)));
				sr.WriteLine(FormaCampo(505, dCompraImpAFIva.ToString("0.00", us)));
				sr.WriteLine(FormaCampo(506, dCompraImpIva0.ToString("0.00", us)));
			}
			else
			{
				dCompraImpBIva = 0;
				dCompraImpAFIva = 0;
				dCompraImpIva0 = 0;
			}
			
			dComprasIva0 += dServicioCompra(dtFecha);
			sr.WriteLine(FormaCampo(507, dComprasIva0.ToString("0.00", us)));

			double dTotal509 = dComprasIva12 + dComprasAFIva + dComprasIva12SC + dCompraImpBIva
				+ dCompraImpAFIva + dCompraImpIva0 + dComprasIva0;
			sr.WriteLine(FormaCampo(509, dTotal509.ToString("0.00", us)));
			#endregion Compras Locales 501 al 509

			#region Devolucion de compras Locales 511 al 519
			stFiltroComp = "'09', '15', '16', '17', '41', '98', '99'";
			double dN_CC_Iva12 = ValorSelect(6, "Not", stFiltroComp, "> 0", 0, "C", 0);
			double dN_CC_AFIva = ValorSelect(6, "Not", stFiltroComp, "> 0", 5, "C", 0);
			dN_CC_Iva12 = dComprasIva12 - dN_CC_Iva12 + dN_CC_AFIva;
			dN_CC_AFIva = dComprasAFIva - dN_CC_AFIva;
			double dN_CC_Iva12SC = ValorSelect(6, "Not", stFiltroComp, "> 0", 0, "SC", 0);
			dN_CC_Iva12SC = dComprasIva12SC - dN_CC_Iva12SC;

			double dN_CC_ImpBIva = ValorSelect(6, "", "'15', '17'", "> 0", 0, "", 0);
			double dN_CC_ImpAFIva = ValorSelect(6, "", "'15', '17'", "> 0", 5, "", 0);
			dN_CC_ImpBIva = dCompraImpBIva - dN_CC_ImpBIva - dN_CC_ImpAFIva;
			dN_CC_ImpAFIva = dCompraImpAFIva - dN_CC_ImpAFIva;
			double dN_CCImpIva0 = ValorSelect(6, "", "'15', '17'", "= 0", 0, "", 0);
			dN_CCImpIva0 = dCompraImpIva0 - dN_CCImpIva0;

			double dN_CC_Iva0 = ValorSelect(6, "Not", stFiltroComp, "= 0", 0, "", 0);
			dN_CC_Iva0 = dComprasIva0 - dN_CC_Iva0;
			double dN_CC_Rise = ValorSelect(4, "Not", stFiltroComp, "= 0", 0, "", 1);
			dN_CC_Rise += ValorSelect(4, "Not", stFiltroComp, "> 0", 0, "", 1);
			dN_CC_Rise = dN_CC_Rise - ValorSelect(6, "Not", stFiltroComp, "= 0", 0, "", 1);
			dN_CC_Rise = dN_CC_Rise - ValorSelect(6, "Not", stFiltroComp, "> 0", 0, "", 1);

			sr.WriteLine(FormaCampo(511, dN_CC_Iva12.ToString("0.00", us)));
			sr.WriteLine(FormaCampo(512, dN_CC_AFIva.ToString("0.00", us)));
			sr.WriteLine(FormaCampo(513, dN_CC_Iva12SC.ToString("0.00", us)));


			if (drSeteo.PNatural == false)
			{
				sr.WriteLine(FormaCampo(514, dN_CC_ImpBIva.ToString("0.00", us)));
				sr.WriteLine(FormaCampo(515, dN_CC_ImpAFIva.ToString("0.00", us)));
				sr.WriteLine(FormaCampo(516, dN_CCImpIva0.ToString("0.00", us)));
			}
			else
			{
				dN_CC_ImpBIva = 0;
				dN_CC_ImpAFIva = 0;
				dN_CCImpIva0 = 0;
			}

			sr.WriteLine(FormaCampo(517, dN_CC_Iva0.ToString("0.00", us)));
			sr.WriteLine(FormaCampo(518, dN_CC_Rise.ToString("0.00", us)));
			double dTotal519 = dN_CC_Iva12 + dN_CC_AFIva + dN_CC_Iva12SC + dN_CC_ImpBIva
				+ dN_CC_ImpAFIva + dN_CCImpIva0 + dN_CC_Iva0 + dN_CC_Rise;
			sr.WriteLine(FormaCampo(519, dTotal519.ToString("0.00", us)));
			#endregion Compras 511 al 519
			#region Resumen de Impuestos 521 - 529 y 531 - 544
			double dComprasIvaImp = dN_CC_Iva12 * .12;
			double dComprasA_FImp = dN_CC_AFIva * .12;
			double dComprasOtroImp = dN_CC_Iva12SC * 0.12;
			double dComprasImpBImp = dN_CC_ImpBIva * 0.12;
			double dComprasImpAFImp = dN_CC_ImpAFIva * 0.12;
			double dTotal529 = dComprasIvaImp + dComprasA_FImp + dComprasOtroImp
				+ dComprasImpBImp+ dComprasImpAFImp;
			sr.WriteLine(FormaCampo(521, dComprasIvaImp.ToString("0.00", us)));
			sr.WriteLine(FormaCampo(522, dComprasA_FImp.ToString("0.00", us)));
			sr.WriteLine(FormaCampo(523, dComprasOtroImp.ToString("0.00", us)));
			sr.WriteLine(FormaCampo(524, dComprasImpBImp.ToString("0.00", us)));
			sr.WriteLine(FormaCampo(525, dComprasImpAFImp.ToString("0.00", us)));
			sr.WriteLine(FormaCampo(529, dTotal529.ToString("0.00", us)));

			double dCompraNoIva = ValorSelect(4, "", "'01', '02'", "= 0", 0, "NI", 0);
			sr.WriteLine(FormaCampo(531, dCompraNoIva.ToString("0.00", us)));
			// adquisiciones excentas del pago de IVA pendiente febrero 2012
			sr.WriteLine(FormaCampo(532, "0.00"));
			sr.WriteLine(FormaCampo(533, "0.00"));
			sr.WriteLine(FormaCampo(534, "0.00"));
			sr.WriteLine(FormaCampo(535, "0.00"));
			//Reembolso de gastos
			double dComprasReembolso0 = ValorSelect(4, "", "'41'", "= 0", 0, "", 0);
			double dComprasReembolsoIva = ValorSelect(4, "", "'41'", "> 0", 0, "", 0);
			dComprasReembolso0 -= ValorSelect(6, "", "'41'", "= 0", 0, "", 0);
			dComprasReembolsoIva -= ValorSelect(6, "", "'41'", "> 0", 0, "", 0);
			double dComprasReembolsoImp = dComprasReembolsoIva * .12;
			dComprasReembolsoIva += dComprasReembolso0;

			sr.WriteLine(FormaCampo(534, dComprasReembolsoIva.ToString("0.00", us)));
			sr.WriteLine(FormaCampo(544, "0.00"));
			sr.WriteLine(FormaCampo(545, dComprasReembolsoImp.ToString("0.00", us)));
			double dValor553 = 0;
			if (dTotal419 > 0) dValor553 = (dTotal419 - dN_C_AFIva0SC - dN_C_AF0SC) / dTotal419;
			double dValor554 = (dTotal529 - dComprasOtroImp) * Math.Round(dValor553, 4);

			sr.WriteLine(FormaCampo(553, dValor553.ToString("0.0000", us)));
			sr.WriteLine(FormaCampo(554, dValor554.ToString("0.00", us)));
			#endregion Resumen de Impuestos 521 522 529

			#endregion Resumen de Compras			
			#region Resumen Impositivo 600
			string stSelect = "Select IsNull(Sum(IsNull(IvaRetenido, 0)+ "
				+ "IsNull(IvaRetenido1, 0)), 0) from retencion "
				+ "Where Fecha >= '" + dtFecha.ToString("yyyyMM") + "01' "
				+ "And Fecha < '" + dtFecha.AddMonths(1).ToString("yyyyMM") + "01'"
				+ " and IngresoEgreso = 1 And Borrar = 0 And NoAplica = 0";
			double dRetenido = Funcion.dEscalarSQL(cdsCompra, stSelect, true);

			double dValor601 = Math.Round(dValor499, 2) - Math.Round(dValor554, 2);
			double dValor602 = 0;
			if (dValor601 < 0)
			{
				dValor602 = -dValor601;
				dValor601 = 0;
			}

			sr.WriteLine(FormaCampo(601, dValor601.ToString("0.00", us)));
			sr.WriteLine(FormaCampo(602, dValor602.ToString("0.00", us)));
			sr.WriteLine(FormaCampo(605, "0"));
			sr.WriteLine(FormaCampo(607, "0"));
			sr.WriteLine(FormaCampo(609, dRetenido.ToString("0.00", us)));
			sr.WriteLine(FormaCampo(611, "0"));
			sr.WriteLine(FormaCampo(613, "0"));
			double dValor615 = 0; //dTotal529 - dTotal429;
			sr.WriteLine(FormaCampo(615, dValor615.ToString("0.00", us)));
			sr.WriteLine(FormaCampo(617, "0"));
			// 619 = 601 - 602 - 605 - 607 - 609 + 611 IF 611 > 0
			// 619 = 601 - 602 - 609 
			double dSubTotalP = dValor601 - dValor602 - dRetenido;
			if (dSubTotalP < 0)
			{
				dSubTotalP = 0;
			}
			sr.WriteLine(FormaCampo(619, dSubTotalP.ToString("0.00", us)));
			sr.WriteLine(FormaCampo(621, "0"));
			sr.WriteLine(FormaCampo(699, dSubTotalP.ToString("0.00", us)));

			#endregion Resumen Impositivo 600
			#region Retenciones
			double dRetencion30 = 0;
			double dRetencion70 = 0;
			double dRetencion100 = 0;

			stSelect = "exec RetenidoIvaTotal '" + dtFecha.ToString("yyyyMM") + "01', '"
				+ dtFecha.AddMonths(1).ToString("yyyyMM") + "01', 0";

			SqlDataReader miReader = Funcion.rEscalarSQL(cdsCompra, stSelect, true);
			while (miReader.Read())
			{
				double dRetenidoIva = miReader.GetDouble(0);
				double dPorcentaje = miReader.GetDouble(1);
				if (dPorcentaje > 29.9 && dPorcentaje < 30.1)
          dRetencion30 += dRetenidoIva;
				if (dPorcentaje > 69.9 && dPorcentaje < 70.1)
					dRetencion70 += dRetenidoIva;
				if (dPorcentaje > 99.9 && dPorcentaje < 100.1)
					dRetencion100 += dRetenidoIva;
			}
			miReader.Close();

			if (drSeteo.PNatural == false)
			{
				sr.WriteLine(FormaCampo(721, dRetencion30.ToString("0.00", us)));
				sr.WriteLine(FormaCampo(723, dRetencion70.ToString("0.00", us)));
				sr.WriteLine(FormaCampo(725, dRetencion100.ToString("0.00", us)));
			}
			else
			{
				dRetencion30 = 0;
				dRetencion70 = 0;
				dRetencion100 = 0;
			}

			double dTotal799 = dRetencion30 + dRetencion70 + dRetencion100;
			sr.WriteLine(FormaCampo(799, dTotal799.ToString("0.00", us)));
			double dValor859 = dSubTotalP + dTotal799;

			if (drSeteo.PNatural == false)
			{
				sr.WriteLine(FormaCampo(859, dValor859.ToString("0.00", us)));
			}
			else
			{
				dValor859 = 0;
			}
			sr.WriteLine(FormaCampo(880, "0.00"));
			sr.WriteLine(FormaCampo(890, "0.00"));
			sr.WriteLine(FormaCampo(897, "0.00"));
			sr.WriteLine(FormaCampo(898, "0.00"));
			sr.WriteLine(FormaCampo(899, "0.00"));

			sr.WriteLine(FormaCampo(902, dValor859.ToString("0.00", us)));
			sr.WriteLine(FormaCampo(903, "0"));
			sr.WriteLine(FormaCampo(904, "0"));
			sr.WriteLine(FormaCampo(905, dValor859.ToString("0.00", us)));
			sr.WriteLine(FormaCampo(906, "0"));
			sr.WriteLine(FormaCampo(907, "0"));
			sr.WriteLine(FormaCampo(909, "0"));
			sr.WriteLine(FormaCampo(911, "0"));
			sr.WriteLine(FormaCampo(913, "0"));
			sr.WriteLine(FormaCampo(915, "0"));
			sr.WriteLine(FormaCampo(917, "0"));
			sr.WriteLine(FormaCampo(919, "0"));
			sr.WriteLine(FormaCampo(999, dValor859.ToString("0.00", us)));
			#endregion Retenciones

			sr.WriteLine("</detalle>");
			sr.WriteLine("</formulario>");
			sr.Close();
			MessageBox.Show("Archivo: " + stArchivo + " Generado", "Información");		
		}


		private void btForm104108()
		{
//			DateTime dtFecha = (DateTime) cmbFecha.Value;
//			#region Actualizacion de descuento
//			string stUpdate = "Update Compra Set DescPorcIva0 = Desc2 / SubtotalExcento * 100.0 "
//				+ "Where Desc2 > 0 And SubtotalExcento > 0 And Month(Fecha) = " + dtFecha.Month.ToString();
//			Funcion.EjecutaSQL(cdsCompra, stUpdate, true);
//
//			stUpdate = "Update Compra Set DescPorcIva = Desc1 / SubtotalIva * 100.0 "
//				+ "Where Desc1 > 0 And SubtotalIva > 0 And Month(Fecha) = " + dtFecha.Month.ToString();
//			Funcion.EjecutaSQL(cdsCompra, stUpdate, true);
//			#endregion Fin de descuento
//			#region Encabezado
//			StreamWriter sr = null;
//			try
//			{
//				string stDir = @"C:\XML_DECLARACIONES\";
//				if (!Directory.Exists(stDir)) stDir = "";
//				sr= File.CreateText(@stDir + "104ORI_" + dtFecha.ToString("MMMyyyy") + ".XML");
//			}
//			catch
//			{
//				MessageBox.Show("Esta abierto el archivo Formulario.xml");
//				Cursor = Cursors.Default;
//			}
//			sr.WriteLine("<?xml version=\"1.0\" encoding=\"UTF-8\"?>");
//			sr.WriteLine("<formulario version=\"0.2\">");
//			sr.WriteLine("<cabecera>");
//			sr.WriteLine("<codigo_version_formulario>04200801</codigo_version_formulario>");
//			sr.WriteLine("<ruc>" + drSeteo.Ruc + "</ruc>");
//			sr.WriteLine("<codigo_moneda>1</codigo_moneda>");
//			sr.WriteLine("</cabecera>");
//			sr.WriteLine("<detalle>");
//
//			sr.WriteLine(FormaCampo(101, dtFecha.Month.ToString()));
//			sr.WriteLine(FormaCampo(102, dtFecha.Year.ToString()));
//			sr.WriteLine(FormaCampo(31, "O"));
//			sr.WriteLine(FormaCampo(104, ""));
//			sr.WriteLine(FormaCampo(201, drSeteo.Ruc));
//			sr.WriteLine(FormaCampo(202, "<![CDATA[" + drSeteo.Empresa + "]]>"));
//			#endregion Encabezado
//
//			#region Proporcion de credito tributario aplicable este mes
//			sr.WriteLine(FormaCampo(301, "1.0000")); //ventas 12 + exportaciones / total
//			sr.WriteLine(FormaCampo(302, "0")); //Saldo credito tributario mes anterior
//			sr.WriteLine(FormaCampo(303, "0")); //dev de iva de este mes
//			sr.WriteLine(FormaCampo(304, "0")); //
//			sr.WriteLine(FormaCampo(399, "0")); //
//			#endregion Proporcion de credito tributario aplicable este mes
//
//			#region Resumen de Ventas
//			#region Ventas Locales 401 421 441
//			string stFiltroComp = "'09', '15', '16', '17', '41', '99', '98'";
//			double dVentasIva12 = ValorSelect(1, "Not", stFiltroComp, "> 0", 0, "", 0);
//			double dNotaCreditoA12 = ValorSelect(5, "Not", stFiltroComp, "> 0", 0, "", 0);
//			double dVentasIva0 = ValorSelect(1, "Not", stFiltroComp, "= 0", 0, "", 0);
//			double dNotaCreditoA0 = ValorSelect(5, "Not", stFiltroComp, "= 0", 0, "", 0);
//			double dVentasIvaImp = dVentasIva12 * .12;
//			#endregion Ventas Locales 401 431 441
//			#region Ventas Locales Inst Publica 410
//			string stFiltroCompIP = "'09', '15', '16', '17', '41', '99', '98'";
////			double dVentasIva12IP = ValorSelect(1, "Not", stFiltroCompIP, "> 0", 0, "", 1);
////			dVentasIva12IP += -ValorSelect(5, "Not", stFiltroCompIP, "> 0", 0, "", 1);
//			double dVentasIva0IP = ValorSelect(1, "Not", stFiltroCompIP, "= 0", 0, "", 1);
//			double dNotaCredVentasIva0IP = ValorSelect(5, "Not", stFiltroCompIP, "= 0", 0, "", 1);
//			double dVentasIvaImpIP = dVentasIva12 * .12;
//			#endregion Ventas Locales IP 410
//
//			#region Ventas directas a exportadores 503 533 553 511 513
//			double dVentaExpIva0 = ValorSelect(1, "", "'16', '98'", "= 0", 0, "", 0);
//			dVentaExpIva0 += -ValorSelect(5, "", "'16', '98'", "= 0", 0, "", 0);
//			double dVentaExpIva = ValorSelect(1, "", "'16', '98'", "> 0", 0, "", 0);
//			dVentaExpIva += -ValorSelect(5, "", "'16', '98'", "> 0", 0, "", 0);
//
//			double dVentasExp = dVentaExpIva + dVentaExpIva0;
//			//Venta de Servicios a Exportadores
//			double dVentasExpS = ValorSelect(1, "", "'16', '98'", ">= 0", 3, "", 0);
//			double dNotaCreditoB0 = ValorSelect(5, "", "'16', '98'", "= 0", 0, "", 0) + dNotaCredVentasIva0IP;
//			double dNotaCreditoBIva = ValorSelect(5, "", "'16', '98'", "> 0", 0, "", 0);
//			double dVentasExpB = dVentasExp - dVentasExpS;
//			double dVentaExpImp = dVentaExpIva * .12;
//			#endregion Ventas directas a exportadores
//			#region Ventas de Activos Fijos
//			double dVentasAFIva = ValorSelect(1, "Not", stFiltroComp, "> 0", 5, "", 0);
//			dVentasAFIva += -ValorSelect(5, "Not", stFiltroComp, "> 0", 5, "", 0);
//			double dVentasAF0 = ValorSelect(1, "Not", stFiltroComp, "= 0", 5, "", 0);
//			dVentasAF0 += -ValorSelect(5, "Not", stFiltroComp, "= 0", 5, "", 0);
//			// AF Ins Publica
//			double dVentasAF0IP = ValorSelect(1, "Not", stFiltroComp, "= 0", 5, "", 1);
//			dVentasAF0IP += -ValorSelect(5, "Not", stFiltroComp, "= 0", 5, "", 1);
//			// AF Ins Publica
//			double dVentasAFImp = dVentasAFIva * .12;
//			//Reembolso de gastos
//			double dVentasReembolso0 = ValorSelect(1, "", "'41'", "= 0", 0, "", 0);
//			//			dVentasReembolso0 += -ValorSelect(5, "", "'41'", "= 0", 0, "");
//			double dVentasReembolsoIva = ValorSelect(1, "", "'41'", "> 0", 0, "", 0);
//			//			dVentasReembolsoIva += -ValorSelect(5, "", "'41'", "> 0", 0, "");
//			double dVentasReembolsoImp = dVentasReembolsoIva * .12;
//			#endregion Ventas de Activos Fijos
//			#region Graba en Archivo codigos del 400
//			dVentasIva0 += - dVentasAF0;
//			dVentaExpIva0 = 0;
//			dVentaExpIva = 0;
//			dVentaExpImp = 0;
//
//			#region Empresas Ventas y Exportaciones Iva 0
//			sr.WriteLine(FormaCampo(401, dVentasIva0.ToString("0.00", us)));
//			sr.WriteLine(FormaCampo(402, dVentasAF0.ToString("0.00", us)));
//			sr.WriteLine(FormaCampo(403, "0")); //Donaciones
//			sr.WriteLine(FormaCampo(404, dNotaCreditoA0.ToString("0.00", us))); //Nota Credito Grupo A mes Actual sin iva
//			sr.WriteLine(FormaCampo(405, "0")); //Nota Credito Grupo A no compensadas sin iva
//			double dSubTotal401_405 = dVentasIva0 + dVentasAF0 - dNotaCreditoA0;
//			sr.WriteLine(FormaCampo(406, dSubTotal401_405.ToString("0.00", us))); //Subtotal 
//
//			sr.WriteLine(FormaCampo(407, dVentaExpIva0.ToString("0.00", us)));
//			sr.WriteLine(FormaCampo(408, dVentasExpB.ToString("0.00", us)));
//			sr.WriteLine(FormaCampo(409, dVentasExpS.ToString("0.00", us)));
//			#endregion Empresas Ventas y Exportaciones Iva 0
//
//			#region Instituciones Publicas
//			dVentasIva0IP += -dVentasAF0IP;
//			sr.WriteLine(FormaCampo(410, dVentasIva0IP.ToString("0.00", us)));
//			sr.WriteLine(FormaCampo(411, dVentasAF0IP.ToString("0.00", us)));
//			sr.WriteLine(FormaCampo(412, dNotaCreditoB0.ToString("0.00", us))); //Nota Credito Grupo B mes Actual sin iva
//			sr.WriteLine(FormaCampo(413, "0")); //Nota Credito Grupo A no compensadas sin iva
//			double dSubTotal407_413 = dVentaExpIva0 + dVentasExpB + dVentasExpS + dVentasIva0IP + dVentasAF0IP - dNotaCreditoB0;
//			sr.WriteLine(FormaCampo(414, dSubTotal407_413.ToString("0.00", us))); //subtotal
//			#endregion Instituciones Publicas
//
//			#region Nota de Credito y Reembolso Proximo Mes
//			sr.WriteLine(FormaCampo(415, "0.00")); //Nota de Credito Proximo Mes Iva 0
//			double dReembolsoG = ValorSelect(1, "", "41", "= 0", 0, "RG", 0);
//			sr.WriteLine(FormaCampo(416, dReembolsoG.ToString("0.00", us))); //Reembolso Proximo Mes Iva 0
//			sr.WriteLine(FormaCampo(435, "0.00")); //Nota de Credito Proximo Mes Iva 12
//			double dReembolsoGIva = ValorSelect(1, "", "41", "> 0", 0, "RG", 0);
//			sr.WriteLine(FormaCampo(436, dReembolsoGIva.ToString("0.00", us))); //Reembolso Proximo Mes Iva 12
//			double dReembolsoGIva1 = dReembolsoGIva * 0.12;
//			sr.WriteLine(FormaCampo(455, "0.00")); //Impuesto Nota Credito Proximo Mes
//			sr.WriteLine(FormaCampo(456, dReembolsoGIva1.ToString("0.00", us))); //Impuesto Reembolso Proximo Mes
//			sr.WriteLine(FormaCampo(457, "0.00")); //Juegos de Azar
//			#endregion Nota de Credito y Reembolso Proximo Mes
//
//			#region Empresas Ventas y Exportaciones Iva 12
//			dVentasIva12 += - dVentasAFIva;
//			sr.WriteLine(FormaCampo(421, dVentasIva12.ToString("0.00", us)));
//			sr.WriteLine(FormaCampo(422, dVentasAFIva.ToString("0.00", us)));
//			sr.WriteLine(FormaCampo(423, "0")); //Donaciones
//			sr.WriteLine(FormaCampo(424, dNotaCreditoA12.ToString("0.00", us))); //Nota Credito Grupo B mes Actual Base IVa
//			sr.WriteLine(FormaCampo(425, "0")); //Nota Credito Grupo A no compensadas Base Iva
//			sr.WriteLine(FormaCampo(427, dVentaExpIva.ToString("0.00", us)));
//			sr.WriteLine(FormaCampo(432,dNotaCreditoBIva.ToString("0.00", us))); //Nota Credito Grupo B Iva este mes
//			sr.WriteLine(FormaCampo(433, "0")); //Nota Credito Grupo B Iva compensar proximo mes
//			double dSubTotal421_433 = dVentasIva12 + dVentasAFIva - dNotaCreditoA12 + dVentaExpIva - dNotaCreditoBIva;
//			sr.WriteLine(FormaCampo(434, dSubTotal421_433.ToString("0.00", us))); //Nota Credito compensar proximo mes
//
//			dVentasIvaImp = dVentasIva12 * .12;
//			double dNotaCreditoA12Imp = dNotaCreditoA12 * 0.12;
//			double dNotaCreditoB12Imp = dNotaCreditoBIva * 0.12;
//			sr.WriteLine(FormaCampo(441, dVentasIvaImp.ToString("0.00", us)));
//			sr.WriteLine(FormaCampo(442, dVentasAFImp.ToString("0.00", us)));
//			sr.WriteLine(FormaCampo(443, "0")); //Donaciones
//			sr.WriteLine(FormaCampo(444,  dNotaCreditoA12Imp.ToString("0.00", us)));
//			sr.WriteLine(FormaCampo(445,  "0.00"));
//			sr.WriteLine(FormaCampo(447, dVentaExpImp.ToString("0.00", us)));
//			sr.WriteLine(FormaCampo(452, dNotaCreditoB12Imp.ToString("0.00", us)));
//			sr.WriteLine(FormaCampo(453, "0.00"));
//			#endregion Empresas Ventas y Exportaciones Iva 12
//
//			double dValor454 = dVentasIvaImp + dVentaExpImp + dVentasAFImp 
//				- dNotaCreditoA12Imp - dNotaCreditoB12Imp;
//			sr.WriteLine(FormaCampo(454, dValor454.ToString("0.00", us))); 
//
//			#endregion Graba en Archivo codigos del 400
//
//			#region Contado y Credito 480 481
//			string stContado = "Select SUM(IsNull(SubTotalIva, 0)) - Sum(IsNull(Descuento, 0)) "
//				+ "From Compra Inner Join CompraComprobante On Compra.idComprobante = CompraComprobante.idComprobante "
//				+ "Where Borrar = 0 And ContadoCredito = 1 And Fecha >= '" + dtFecha.ToString("yyyyMM")
//				+ "01' and Fecha < '" + dtFecha.AddMonths(1).ToString("yyyyMM") + "01'"
//				+ "And idTipoFactura = 1 And CompraComprobante.Codigo not in ('16', '98', '41')";
//			double dContadoVenta480 = Funcion.dEscalarSQL(cdsCompra, stContado, false);
//			stContado = "Select SUM(IsNull(SubTotalIva, 0)) - Sum(IsNull(Descuento, 0)) "
//				+ "From Compra Inner Join CompraComprobante On Compra.idComprobante = CompraComprobante.idComprobante "
//				+ "Where Borrar = 0 And ContadoCredito = 1 And Fecha >= '" + dtFecha.ToString("yyyyMM")
//				+ "01' and Fecha < '" + dtFecha.AddMonths(1).ToString("yyyyMM") + "01'"
//				+ "And idTipoFactura = 5 "; //And CompraComprobante.Codigo not in ('16', '98')";
//			dContadoVenta480 += -Funcion.dEscalarSQL(cdsCompra, stContado, false);
//
//			string stCredito = "Select SUM(IsNull(SubTotalIva, 0)) - Sum(IsNull(Descuento, 0)) "
//				+ "From Compra Inner Join CompraComprobante On Compra.idComprobante = CompraComprobante.idComprobante "
//				+ "Where Borrar = 0 And ContadoCredito = 2 And Fecha >= '" + dtFecha.ToString("yyyyMM")
//				+ "01' and Fecha < '" + dtFecha.AddMonths(1).ToString("yyyyMM") + "01'"
//				+ "And idTipoFactura = 1 And CompraComprobante.Codigo not in ('16', '98', '41')";
//			double dCreditoVenta481 = Funcion.dEscalarSQL(cdsCompra, stCredito, false);
//			stCredito = "Select SUM(IsNull(SubTotalIva, 0)) - Sum(IsNull(Descuento, 0)) "
//				+ "From Compra Inner Join CompraComprobante On Compra.idComprobante = CompraComprobante.idComprobante "
//				+ "Where Borrar = 0 And ContadoCredito = 2 And Fecha >= '" + dtFecha.ToString("yyyyMM")
//				+ "01' and Fecha < '" + dtFecha.AddMonths(1).ToString("yyyyMM") + "01'"
//				+ "And idTipoFactura = 5 "; //And CompraComprobante.Codigo not in ('16', '98')";
//			dCreditoVenta481 += -Funcion.dEscalarSQL(cdsCompra, stCredito, false);
//			#endregion ContadoCredito
//			#region Numero de comprobantes
//			int iNumFactVenta = NumComprobante(1, 0);
//			int iNumNotaVenta = NumComprobante(1, 2);
//			int iNumDocExp = NumComprobante(1, 13); // Formulario unico de exportacion
//			int iNumNotaCredito = NumComprobante(5, 0);
//			int iNumNotaDebito = NumComprobante(27, 0);
//			#endregion Numero de comprobantes
//			#region Grabacion de Comprobantes 105 - 110
////			iNumFactVenta += -iNumNotaVenta -iNumDocExp;
//			sr.WriteLine(FormaCampo(107, iNumFactVenta.ToString("0.00", us)));
//			sr.WriteLine(FormaCampo(480, dContadoVenta480.ToString("0.00", us)));
//			sr.WriteLine(FormaCampo(481, dCreditoVenta481.ToString("0.00", us)));
//			sr.WriteLine(FormaCampo(482, dValor454.ToString("0.00", us))); 
//			sr.WriteLine(FormaCampo(483, "0.00"));
//			double dImpEsteMes = dContadoVenta480 * 0.12;
//			double dImpProximoMes = dValor454 - dImpEsteMes;
//			sr.WriteLine(FormaCampo(484, dImpEsteMes.ToString("0.00", us))); // Contado del mes
//			sr.WriteLine(FormaCampo(485, dImpProximoMes.ToString("0.00", us))); // 
//			double dValor499 = dImpEsteMes ;
//			sr.WriteLine(FormaCampo(460, "0.00")); // Consultar en Formulario
//			sr.WriteLine(FormaCampo(499, dValor499.ToString("0.00", us))); // Impuesto a Liquidar este mes
//			sr.WriteLine(FormaCampo(107, "0"));
////			sr.WriteLine(FormaCampo(108, iNumDocExp.ToString()));
////			sr.WriteLine(FormaCampo(109, iNumNotaCredito.ToString()));
////			sr.WriteLine(FormaCampo(110, iNumNotaDebito.ToString()));
//			#endregion Grabacion de Comprobantes 105 - 110
//			#endregion Resumen de Ventas
//			
//			#region Resumen de Compras 500
//			#region Compras Locales de Servicios 501 al 526
//			double dCompraIva = ValorSelect(4, "Not", stFiltroComp, "> 0", 0, "C", 0);
//			double dNotaCredCompraIva = ValorSelect(6, "Not", stFiltroComp, "> 0", 0, "C", 0);
//			double dCompraIva0 = ValorSelect(4, "Not", stFiltroComp, "= 0", 0, "C", 0);
//			double dNotaCredCompraIva0 = ValorSelect(6, "Not", stFiltroComp, "= 0", 0, "C", 0);
//
//			double dCompraIvaS12 = ValorSelect(4, "Not", stFiltroComp, "> 0", 3, "C", 0);
////			dCompraIvaS12 += -ValorSelect(6, "Not", stFiltroComp, "> 0", 3, "C", 0);
//			double dCompraIvaSImp = dCompraIvaS12 * .12;
//			double dCompraIvaS0 = ValorSelect(4, "Not", stFiltroComp, "= 0", 3, "C", 0);
////Santiago			dCompraIvaS0 += -ValorSelect(6, "Not", stFiltroComp, "= 0", 3, "C", 0);
//			#endregion Ventas Locales 603 633 653
//
//			#region Importacion de Bienes 609 639 659
//			double dCompraExpBIva0 = ValorSelect(4, "", "'15', '17'", "= 0", 0, "SC", 0);
////Santiago			dCompraExpBIva0 += -ValorSelect(6, "", "'15', '17'", "= 0", 0, "SC", 0);
//			double dCompraExpBIva = ValorSelect(4, "", "'15', '17'", "> 0", 0, "C", 0);
////			dCompraExpBIva += -ValorSelect(6, "", "'15', '17'", "> 0", 0, "C", 0);
//			double dCompraExpSIva0 = ValorSelect(4, "", "'15', '17'", "= 0", 3, "SC", 0);
////			dCompraExpSIva0 += -ValorSelect(6, "", "'15', '17'", "= 0", 3, "SC", 0);
//			double dCompraExpSIva = ValorSelect(4, "", "'15', '17'", "> 0", 3, "C", 0);
//			dCompraExpSIva += -ValorSelect(6, "", "'15', '17'", "> 0", 3, "C", 0);
//			double dCompraExpSImp = dCompraExpSIva * .12;
//			double dCompraExpAFIva0 = ValorSelect(4, "", "'15', '17'", "= 0", 5, "SC", 0);
////			dCompraExpAFIva0 += -ValorSelect(6, "", "'15', '17'", "= 0", 5, "C", 0);
//			double dCompraExpAFIva = ValorSelect(4, "", "'15', '17'", "> 0", 5, "C", 0);
////			dCompraExpAFIva += -ValorSelect(6, "", "'15', '17'", "> 0", 5, "C", 0);
//			double dCompraExpAFImp = dCompraExpAFIva * .12;
//			#endregion Ventas directas a exportadores
//
//			#region Compras de Activos Fijos 605 635 655
//			double dCompraAFIva = ValorSelect(4, "Not", stFiltroComp, "> 0", 5, "C", 0);
//			//dCompraAFIva += -ValorSelect(6, "Not", stFiltroComp, "> 0", 5, "C", 0);
//			double dCompraAF0 = ValorSelect(4, "Not", stFiltroComp, "= 0", 5, "C", 0);
////Santiago			dCompraAF0 += -ValorSelect(6, "Not", stFiltroComp, "= 0", 5, "C", 0);
//			double dCompraG0 = 0;
//			if (drSeteoC["GastoAT_F"] != DBNull.Value && (bool) drSeteoC["GastoAT_F"] == true)
//				dCompraG0 = ValorSelect(4, "Not", stFiltroComp, "= 0", 2, "C", 0);
//			double dCompraAFImp = dCompraAFIva * .12;
//
//			//Reembolso de gastos
//			double dComprasReembolso0 = ValorSelect(4, "", "'41'", "= 0", 0, "C", 0);
//			dComprasReembolso0 += -ValorSelect(6, "", "'41'", "= 0", 0, "C", 0);
//			double dComprasReembolsoIva = ValorSelect(4, "", "'41'", "> 0", 0, "C", 0);
//			dComprasReembolsoIva += -ValorSelect(6, "", "'41'", "> 0", 0, "C", 0);
//			double dComprasReembolsoImp = dComprasReembolsoIva * .12;
//
//			#endregion Compra de Activos Fijos
//
//			#region Graba en Archivo codigos del 500
//			dCompraIva += -dCompraIvaS12 - dCompraAFIva;
//			double dCompraIvaImp = dCompraIva * .12;
//			dCompraIva0 += -dCompraIvaS0 - dCompraAF0 - dCompraG0;
//
//			dCompraExpBIva += -dCompraExpSIva - dCompraExpAFIva;
//			dCompraExpBIva0 += -dCompraExpSIva0 - dCompraExpAFIva0;
//			double dCompraExpBImp = dCompraExpBIva * .12;
//
//			sr.WriteLine(FormaCampo(501, dCompraIva0.ToString("0.00", us)));
//			sr.WriteLine(FormaCampo(502, dCompraAF0.ToString("0.00", us)));
//			sr.WriteLine(FormaCampo(503, dCompraIvaS0.ToString("0.00", us)));
//			sr.WriteLine(FormaCampo(504, dCompraExpBIva0.ToString("0.00", us)));
//			sr.WriteLine(FormaCampo(505, dCompraExpAFIva0.ToString("0.00", us)));
//			sr.WriteLine(FormaCampo(508, dNotaCredCompraIva0.ToString("0.00", us)));
//			sr.WriteLine(FormaCampo(509, "0.00"));
//			double dSubTotal501_509 = dCompraIva0 + dCompraAF0 + dCompraIvaS0 + dCompraExpBIva0
//				+ dCompraExpAFIva0 - dNotaCredCompraIva0;
//			sr.WriteLine(FormaCampo(510, dSubTotal501_509.ToString("0.00", us)));
//			sr.WriteLine(FormaCampo(511, "0.00"));
//
//			double dCompraSCIva = ValorSelect(4, "", "'02', '01', '03', '09', '17'", "> 0", 0, "SC", 0);
//			double dCompraSCIva0 = ValorSelect(4, "", "'02', '01', '03', '09', '17'", "= 0", 0, "SC", 0);
//			sr.WriteLine(FormaCampo(512, dComprasReembolso0.ToString("0.00", us))); //Reembolso de gastos
//			sr.WriteLine(FormaCampo(513, dCompraSCIva0.ToString("0.00", us)));
//
//			sr.WriteLine(FormaCampo(521, dCompraIva.ToString("0.00", us)));
//			sr.WriteLine(FormaCampo(522, dCompraAFIva.ToString("0.00", us)));
//			sr.WriteLine(FormaCampo(523, dCompraIvaS12.ToString("0.00", us)));
//			sr.WriteLine(FormaCampo(524, dCompraExpBIva.ToString("0.00", us)));
//			sr.WriteLine(FormaCampo(525, dCompraExpAFIva.ToString("0.00", us)));
//			sr.WriteLine(FormaCampo(526, "0"));
//			sr.WriteLine(FormaCampo(527, "0"));
//			sr.WriteLine(FormaCampo(528, dNotaCredCompraIva.ToString("0.00", us)));
//			sr.WriteLine(FormaCampo(529, "0.00"));
//			double dSubTotal521_529 = dCompraIva + dCompraAFIva + dCompraIvaS12 + dCompraExpBIva
//				+ dCompraExpAFIva - dNotaCredCompraIva;
//			sr.WriteLine(FormaCampo(530, dSubTotal521_529.ToString("0.00", us)));
//
//			sr.WriteLine(FormaCampo(531, "0.00"));
//			sr.WriteLine(FormaCampo(532, dComprasReembolsoIva.ToString("0.00", us))); //Reembolso de gastos
//			sr.WriteLine(FormaCampo(533, dCompraSCIva.ToString("0.00", us)));
//
//			double dNotaCredCompraIvaImp = dNotaCredCompraIva * 0.12;
//			sr.WriteLine(FormaCampo(541, dCompraIvaImp.ToString("0.00", us)));
//			sr.WriteLine(FormaCampo(542, dCompraAFImp.ToString("0.00", us)));
//			sr.WriteLine(FormaCampo(543, dCompraIvaSImp.ToString("0.00", us)));
//			sr.WriteLine(FormaCampo(544, dCompraExpBImp.ToString("0.00", us)));
//			sr.WriteLine(FormaCampo(545, dCompraExpAFImp.ToString("0.00", us)));
//			sr.WriteLine(FormaCampo(546, "0"));
//			sr.WriteLine(FormaCampo(547, "0"));
//			sr.WriteLine(FormaCampo(548, dNotaCredCompraIvaImp.ToString("0.00", us)));
//			sr.WriteLine(FormaCampo(549, "0.00"));
//			double dSubTotal541_549 = dCompraIvaImp + dCompraAFImp + dCompraIvaSImp + dCompraExpBImp
//				+ dCompraExpAFImp - dNotaCredCompraIvaImp;
//			sr.WriteLine(FormaCampo(550, dSubTotal541_549.ToString("0.00", us)));
//			sr.WriteLine(FormaCampo(551, "0.00"));
//			sr.WriteLine(FormaCampo(552, dComprasReembolsoImp.ToString("0.00", us))); //Reembolso de gastos
//			double dCompraSCImp = dCompraSCIva * 0.12;
//			sr.WriteLine(FormaCampo(553, dCompraSCImp.ToString("0.00", us)));
//			sr.WriteLine(FormaCampo(554, "0.00"));
//			sr.WriteLine(FormaCampo(555, "0.00"));
//			sr.WriteLine(FormaCampo(560, "0.00"));
//			sr.WriteLine(FormaCampo(665, "0"));
//
//			//Calculo de 650
//			double dValor650 = dCompraIva0 + dCompraIvaS0 + dCompraAF0
//				+ dCompraExpBIva0 + dCompraExpSIva0 + dCompraExpAFIva0
//				+ dCompraIva + dCompraIvaS12 + dCompraAFIva
//				+ dCompraExpBIva + dCompraExpSIva + dCompraExpAFIva
//				+ dCompraSCIva + dCompraSCIva0;
//			sr.WriteLine(FormaCampo(650, dValor650.ToString("0.00", us)));
//			double dValor698 = dCompraIvaImp + dCompraIvaSImp + dCompraAFImp
//				+ dCompraExpSImp + dCompraExpBImp + dCompraExpAFImp;
//
//			#endregion Graba en Archivo codigos del 600
//			#region Numero de comprobantes
//			int iNumFactCompra = NumComprobante(4, 1);
//			iNumFactCompra += NumComprobante(4, 18);
//			int iNumNotaCompra = NumComprobante(4, 2);
//			int iNumLiqCompra = NumComprobante(4, 3);
//			int iNumDocImp = NumComprobante(4, 12);
//			iNumDocImp += NumComprobante(4, 14);
//			int iNumOtrosCompra = NumComprobante(4, 0);
//			string stNoAplica = "Select idComprobante From CompraComprobante Where Codigo = '99'";
//			int idNoAplica = Funcion.iEscalarSQL(cdsCompra, stNoAplica, true);
//			iNumOtrosCompra += -NumComprobante(4, idNoAplica);
//			iNumOtrosCompra += -iNumFactCompra - iNumNotaCompra - iNumDocImp - iNumLiqCompra;
//			int iNumNotaCreditoC = NumComprobante(6, 0); // Devolucion de Compra
//			int iNumNotaDebitoC = NumComprobante(28, 0); // Nota de Debito en compra
//			#endregion Numero de comprobantes
//			#region Grabacion de Comprobantes 111 - 117
//			sr.WriteLine(FormaCampo(111, iNumFactCompra.ToString()));
//			sr.WriteLine(FormaCampo(112, iNumNotaCompra.ToString()));
//			sr.WriteLine(FormaCampo(113, iNumLiqCompra.ToString()));
//			sr.WriteLine(FormaCampo(114, iNumOtrosCompra.ToString()));
//			sr.WriteLine(FormaCampo(115, iNumDocImp.ToString()));
//			sr.WriteLine(FormaCampo(116, iNumNotaCreditoC.ToString()));
//			sr.WriteLine(FormaCampo(117, iNumNotaDebitoC.ToString()));
//			#endregion Grabacion de Comprobantes 111 - 117
//			#endregion Resumen de compras
//			
//			#region Resumen Impositivo 600
//			string stSelect = "Select IsNull(Sum(IsNull(IvaRetenido, 0)+ "
//				+ "IsNull(IvaRetenido1, 0)), 0) from retencion "
//				+ "Where Fecha >= '" + dtFecha.ToString("yyyyMM") + "01' "
//				+ "And Fecha < '" + dtFecha.AddMonths(1).ToString("yyyyMM") + "01'"
//				+ " and IngresoEgreso = 1 And Borrar = 0 And NoAplica = 0";
//			double dRetenido = Funcion.dEscalarSQL(cdsCompra, stSelect, true);
//
//			double dValor601 = dValor499;
//			double dValor602 = 0;
//			if (dValor601 < 0)
//			{
//				dValor602 = -dValor601;
//				dValor601 = 0;
//			}
//
//			sr.WriteLine(FormaCampo(601, dValor601.ToString("0.00", us)));
//			sr.WriteLine(FormaCampo(602, dValor602.ToString("0.00", us)));
//			sr.WriteLine(FormaCampo(603, "0")); // Campo 399
//			sr.WriteLine(FormaCampo(604, dRetenido.ToString("0.00", us)));
//			double dSubTotalP = dValor601 - dValor602 - dRetenido;
//			if (dSubTotalP >= 0)
//			{
//				sr.WriteLine(FormaCampo(698, "0"));
//				sr.WriteLine(FormaCampo(699, dSubTotalP.ToString("0.00", us)));
//			}
//			else
//			{
//				dSubTotalP *= -1;
//				sr.WriteLine(FormaCampo(698, dSubTotalP.ToString("0.00", us)));
//				sr.WriteLine(FormaCampo(699, "0"));
//			}
//
//			#endregion Resumen Impositivo 600
//			#region Retenciones
//			double[] dRetencion = new double[100];
//			double[] dPorcentaje = new double[100];
//			for (int i=0; i<100; i++) 
//			{
//				dRetencion[i] = 0;
//				dPorcentaje[i] = 0;
//			}
//
//			#region Retencion Primera
//			stSelect = "Select IsNull(Sum(IsNull(IvaRetenido, 0)), 0), "
//				+ "Max(IsNull(RetencionIva.Porcentaje, 0)), "
//				+ "isNull(RetencionIva.Codigo, '0') from retencion "
//				+ "Left Join RetencionIva On RetencionIva.idRetIva = Retencion.idRetIva "
//				+ "Where Fecha >= '" + dtFecha.ToString("yyyyMM") + "01' And "
//				+ "Fecha < '" + dtFecha.AddMonths(1).ToString("yyyyMM") + "01' "
//				+ " And Retencion.IngresoEgreso = 2 And Retencion.NoAplica = 0 And Retencion.Borrar = 0 "
//				+ "group by RetencionIva.Codigo ";
//
//			SqlDataReader miReader = Funcion.rEscalarSQL(cdsCompra, stSelect, true);
//			double dAcumula = 0;
//			while (miReader.Read())
//			{
//				int iCodigo = int.Parse(miReader.GetString(2));
//				int j = 0;
//				if (iCodigo > 700 && iCodigo < 720) j = iCodigo-700;
//				double dRet2 = miReader.GetDouble(0);
//				dRetencion[j] += dRet2;
//				double dBase = miReader.GetDouble(1);
//				if (dBase > 0)
//					dBase = dRet2 / dBase*100.0;
//				else
//					dBase = 0;
//				dPorcentaje[j] += dBase;
//				dAcumula += dRet2;
//			}
//			miReader.Close();
//			#endregion Retencion Primera
//
//			#region Retencion Segunda
//			stSelect = "Select IsNull(Sum(IsNull(Retencion.IvaRetenido1, 0)), 0), "
//				+ "Max(IsNull(RetencionIva.Porcentaje, 0)), "
//				+ "isNull(RetencionIva.Codigo, '0') from retencion "
//				+ "Left Join RetencionIva On RetencionIva.idRetIva = Retencion.idRetIva1 "
//				+ "Where Retencion.Fecha >= '" + dtFecha.ToString("yyyyMM") + "01' And "
//				+ "Retencion.Fecha < '" + dtFecha.AddMonths(1).ToString("yyyyMM") + "01' "
//				+ " And Retencion.IngresoEgreso = 2 And Retencion.NoAplica = 0 And Retencion.Borrar = 0 "
//				+ "group by RetencionIva.Codigo ";
//
//			miReader = Funcion.rEscalarSQL(cdsCompra, stSelect, true);
//			while (miReader.Read())
//			{
//				int iCodigo = int.Parse(miReader.GetString(2));
//				int j = 0;
//				if (iCodigo > 700 && iCodigo < 720) j = iCodigo-700;
//				double dRet2 = miReader.GetDouble(0);
//				dRetencion[j] += dRet2;
//				double dBase = miReader.GetDouble(1);
//				if (dBase > 0)
//					dBase = dRet2 / dBase*100.0;
//				else
//					dBase = 0;
//				dPorcentaje[j] += dBase;
//				dAcumula += dRet2;
//			}
//			miReader.Close();
//			#endregion Retencion Segunda
//
//			for (int i=1; i<=21; i++)
//			{
//				sr.WriteLine(FormaCampo(i + 700, dPorcentaje[i].ToString("0.00", us)));
//				sr.WriteLine(FormaCampo(i + 720, dRetencion[i].ToString("0.00", us)));
//			}
//			sr.WriteLine(FormaCampo(798, dAcumula.ToString("0.00", us)));
//			// Agregar codigo serie y autorizacion
//			stSelect = "Select Count(distinct(IsNull(SerieRet, '') + IsNull(codigo, ''))) From Retencion "
//				+ "Where Fecha >= '" + dtFecha.ToString("yyyMM")
//				+ "01' And Fecha < '" + dtFecha.AddMonths(1).ToString("yyyMM") 
//				+ "01' And IngresoEgreso = 2 And NoAplica = 0 "
//				+ "And (IvaRetenido + IvaRetenido1) > 0 ";
//			int iCuenta = Funcion.iEscalarSQL(cdsCompra, stSelect, true);
//			sr.WriteLine(FormaCampo(118, iCuenta.ToString("0.00", us)));
//			double dValor899 = dAcumula;
//			if (dSubTotalP > 0) dValor899 += dSubTotalP;
//			sr.WriteLine(FormaCampo(799, dValor899.ToString("0.00", us)));
//
//			sr.WriteLine(FormaCampo(901, "0"));
//			sr.WriteLine(FormaCampo(902, dValor899.ToString("0.00", us)));
//			sr.WriteLine(FormaCampo(903, "0"));
//			sr.WriteLine(FormaCampo(904, "0"));
//			sr.WriteLine(FormaCampo(999, dValor899.ToString("0.00", us)));
//			#endregion Retenciones
//
//			#region Grabacion de Valores
//			sr.WriteLine(FormaCampo(198, drSeteo.CedulaRep));
//			sr.WriteLine(FormaCampo(199, drSeteo.CedulaContador + "001"));
//
//			sr.WriteLine(FormaCampo(905, dValor899.ToString("0.00", us)));
//			sr.WriteLine(FormaCampo(906, "0"));
//			sr.WriteLine(FormaCampo(907, "0"));
//			sr.WriteLine(FormaCampo(908, ""));
//			sr.WriteLine(FormaCampo(909, "0"));
//			sr.WriteLine(FormaCampo(910, ""));
//			sr.WriteLine(FormaCampo(911, "0"));
//			sr.WriteLine(FormaCampo(912, ""));
//			sr.WriteLine(FormaCampo(913, "0"));
//			sr.WriteLine(FormaCampo(914, ""));
//			sr.WriteLine(FormaCampo(915, "0"));
//			sr.WriteLine(FormaCampo(916, ""));
//			sr.WriteLine(FormaCampo(917, "0"));
//			sr.WriteLine(FormaCampo(918, ""));
//			sr.WriteLine(FormaCampo(919, "0"));
//			sr.WriteLine(FormaCampo(922, "16"));
//			#endregion Grabacion de Valores
//
//			sr.WriteLine("</detalle>");
//			sr.WriteLine("</formulario>");
//			sr.Close();
//			MessageBox.Show("Formulario 104 del 2008 Generado");		
		}

		private void btForm104_Click(object sender, System.EventArgs e)
		{
			DateTime dtFecha = (DateTime) cmbFecha.Value;
			btForm104109();
		}

		private void Formulario103_Closing(object sender, System.ComponentModel.CancelEventArgs e)
		{
			string stAudita = "Exec AuditaCrear 53, 7, 'Formularios'";
			Funcion.EjecutaSQL(cdsCompra, stAudita, true);
		}

		private void btVerificacion_Click(object sender, System.EventArgs e)
		{
			DateTime dtFecha = (DateTime) cmbFecha.Value;
			string stFechaIni = dtFecha.Year.ToString();
			if (dtFecha.Month > 9) stFechaIni += dtFecha.Month.ToString();
			else stFechaIni += "0" + dtFecha.Month.ToString();
			stFechaIni += "01'";

			dtFecha = dtFecha.AddMonths(1);
			string stFechaFin = dtFecha.Year.ToString();
			if (dtFecha.Month > 9) stFechaFin += dtFecha.Month.ToString();
			else stFechaFin += "0" + dtFecha.Month.ToString();
			stFechaFin += "01'";

			string stSelect = "Select Count(*) From Compra inner Join  "
				+ " detcompra on compra.idcompra = detcompra.idcompra "
				+ " inner Join CompraComprobante "
				+"On Compra.idComprobante = CompraComprobante.idComprobante Where Compra.Fecha >= '"
				+ stFechaIni + " And Compra.Fecha < '" + stFechaFin 
				+ " And CompraComprobante.Codigo In ('15', '17') "
				+ " And DetCompra.Impuesto = 0 And Compra.idCredTributario Not In (2, 7) "
				+ " And Compra.idTipoFactura = 4";
			int iCuenta = Funcion.iEscalarSQL(cdsCompra, stSelect, true);
			if (iCuenta > 0)
				MessageBox.Show("Existen " + iCuenta.ToString() + " registros en Importaciones "
					+ " con credito tributario y sin iva");
			else
				MessageBox.Show("Revision terminada");
		}

		private void btRep103_Click(object sender, System.EventArgs e)
		{
		
		}

		private void btRep104_Click(object sender, System.EventArgs e)
		{
			Cursor = Cursors.WaitCursor;
			DateTime dtFechaDesde = (DateTime) this.cmbDesde.Value;
			DateTime dtFechaHasta = (DateTime) this.cmbHasta.Value;

			ParameterFields paramFields = new ParameterFields ();
			ParameterField paramField = new ParameterField ();
			ParameterDiscreteValue discreteVal = new ParameterDiscreteValue ();

			paramField.ParameterFieldName = "@FechaDesde";
			discreteVal.Value = dtFechaDesde.ToString("yyyy-MM-dd");
			paramField.CurrentValues.Add (discreteVal);
			paramFields.Add (paramField);

			ParameterField paramField1 = new ParameterField ();
			ParameterDiscreteValue discreteVal1 = new ParameterDiscreteValue ();
			paramField1.ParameterFieldName = "@FechaHasta";
			discreteVal1.Value = dtFechaHasta.ToString("yyyy-MM-dd");
			paramField1.CurrentValues.Add (discreteVal1);
			paramFields.Add (paramField1);

			Reporte miRep = new Reporte("Formulario104.rpt", "");
			miRep.crVista.ParameterFieldInfo = paramFields;
			miRep.MdiParent = this.MdiParent;
			if (chkResumen.Checked) miRep.Resumen();
			miRep.Show();
			Cursor = Cursors.Default;
		}

		private void btForm101_Click(object sender, System.EventArgs e)
		{
			Cursor = Cursors.WaitCursor;
			#region Encabezado
			DateTime dtFecha = (DateTime) cmbFecha.Value;
			StreamWriter sr = null;
			try
			{
				string stDir = @"C:\XML_DECLARACIONES\";
				if (!Directory.Exists(stDir)) stDir = "";
				sr= File.CreateText(@stDir + "101ORI_" + dtFecha.ToString("yyyy") + ".XML");
			}
			catch
			{
				MessageBox.Show("Esta abierto el archivo Formulario.xml");
				Cursor = Cursors.Default;
				return;
			}
			sr.WriteLine("<?xml version=\"1.0\" encoding=\"UTF-8\"?>");
			sr.WriteLine("<formulario version=\"0.2\">");
			sr.WriteLine("<cabecera>");
			string stVersionFor = "01200901";
			if (dtFecha.Year>=2009) stVersionFor = "01200901";
			sr.WriteLine("<codigo_version_formulario>" + stVersionFor +  "</codigo_version_formulario>");
			sr.WriteLine("<ruc>"+drSeteo.Ruc.Trim()+"</ruc>");
			sr.WriteLine("<codigo_moneda>1</codigo_moneda>");
			sr.WriteLine("</cabecera>");
			#endregion Encabezado

			#region Detalle
			sr.WriteLine("<detalle>");
			string stExec = "";
			string stvalor = "0.00";
			for (int j=11; j<31; j++)
			{
				stExec = "Exec Formulario101 " + j + ","+ dtFecha.Year.ToString();
				stvalor = Funcion.sEscalarSQL(cdsCompra, stExec, true);
				sr.WriteLine(FormaCampo(j, stvalor));
				if (j>14 && j<16) {j=30;}
			}
			sr.WriteLine(FormaCampo(31,"O"));
			sr.WriteLine(FormaCampo(102, dtFecha.Year.ToString()));
			sr.WriteLine(FormaCampo(104, ""));

			sr.WriteLine(FormaCampo(198, drSeteo.CedulaRep.Trim()));
			sr.WriteLine(FormaCampo(199, drSeteo.CedulaContador.Trim() + "001"));
			sr.WriteLine(FormaCampo(201, drSeteo.Ruc.Trim()));
			sr.WriteLine(FormaCampo(202, "<![CDATA[" + drSeteo.Empresa.Trim() + "]]>"));
			sr.WriteLine(FormaCampo(203, ""));

			for (int j=311; j<1000; j++)
			{
				stExec = "Exec Formulario101 " + j + ","+ dtFecha.Year.ToString();
				stvalor = Funcion.sEscalarSQL(cdsCompra, stExec, true);
				sr.WriteLine(FormaCampo(j, stvalor));

				if (j == 321 || j == 339 || j == 369 || j == 371 || j == 373 || j == 375 ||  
						j == 379 || j == 397 || j == 439 || j == 499 || j == 501 || j == 503 || 
						j == 505 || j == 507 || j == 513 || j == 515 || j == 517 || j == 599 ||
						j == 699 || j == 713 || j == 740 || j == 792 || j == 799 || j == 839 || 
						j == 849)
								{j++;}
				if ( j == 856 || j == 899 || j == 919 )
						{j=j + 2;}
				if (j == 509)
						{j=j + 3;}
				if (j == 392)
						{j=j + 4;}
				if (j == 332 || j == 692 || j == 822 || j == 842 || j == 890 )
						{j=j + 6;}
				if (j == 802)
						{j=j + 8;}
				if (j == 469 || j == 479 || j == 489 || j == 829 || j == 859 || j == 869)
						{j=j + 9;}
				if (j == 428 || j == 458 || j == 879)
						{j=j + 10;}						
				if (j == 399)
						{j=j + 11;}
				if (j == 350)
						{j=j + 18;}						
				if (j == 922)
						{j=j + 76;}						
				if (j == 519 || j == 612)
						{j=j + 78;}						
		}


			sr.WriteLine("</detalle>");
			#endregion Detalle
	
			sr.WriteLine("</formulario>");
			sr.Close();
			Cursor = Cursors.Default;
			MessageBox.Show("101 Generado");

		}

		private void btRep101_Click(object sender, System.EventArgs e)
		{
			Cursor = Cursors.WaitCursor;
			DateTime dtDesde = (DateTime) cmbDesde.Value;
			DateTime dtHasta = (DateTime) cmbHasta.Value;
			string stExec = "Exec Formulario101Rep '" + dtDesde.ToString("yyyyMMdd") 
				+ "', '" + dtHasta.ToString("yyyyMMdd") + "'";
			Funcion.sEscalarSQL(cdsCompra, stExec, true);
			
			Reporte miRep = new Reporte("Form101Rep.rpt","");
			miRep.MdiParent = this.MdiParent;
			if (chkResumen.Checked) miRep.Resumen();
			miRep.Show();
			Cursor = Cursors.Default;

		}

	}
}
