using System;
using System.Data;
using System.Data.SqlClient;
using System.Globalization;
using C1.Data;
using System.IO;
using System.Text;
using System.Xml;
using System.Xml.Schema;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using LibContabilidad.DataObjects;
using LibFacturacion.DataObjects;

namespace Latinium
{
	/// <summary>
	/// Descripción breve de AnexoTransaccional.
	/// </summary>
	public class AnexoTransaccional : DevExpress.XtraEditors.XtraForm
	{
		private DevExpress.XtraEditors.SimpleButton btGenerar;
		private Infragistics.Win.UltraWinEditors.UltraDateTimeEditor cmbFecha;
		private C1.Data.C1DataSet cdsSeteo;
		private C1.Data.C1DataSet cdsCompra;
		private DevExpress.XtraEditors.SimpleButton btValidar;
		private DevExpress.XtraEditors.SimpleButton btRepetido;
		private DevExpress.XtraEditors.SimpleButton ultraButton1;
		private DevExpress.XtraEditors.SimpleButton btLee;
		private DevExpress.XtraEditors.SimpleButton btRepCliente;
		private Infragistics.Win.UltraWinEditors.UltraDateTimeEditor cmbHasta;
		private Infragistics.Win.Misc.UltraLabel ultraLabel1;
		private Infragistics.Win.Misc.UltraLabel ultraLabel2;
		private Infragistics.Win.UltraWinEditors.UltraCheckEditor chkExcluirVenta;
		private Infragistics.Win.UltraWinEditors.UltraCheckEditor chkAjusteIva;
		/// <summary>
		/// Variable del diseñador requerida.
		/// </summary>
		private System.ComponentModel.Container components = null;

		public AnexoTransaccional()
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
			this.btGenerar = new DevExpress.XtraEditors.SimpleButton();
			this.cmbFecha = new Infragistics.Win.UltraWinEditors.UltraDateTimeEditor();
			this.cdsSeteo = new C1.Data.C1DataSet();
			this.cdsCompra = new C1.Data.C1DataSet();
			this.btValidar = new DevExpress.XtraEditors.SimpleButton();
			this.btRepetido = new DevExpress.XtraEditors.SimpleButton();
			this.ultraButton1 = new DevExpress.XtraEditors.SimpleButton();
			this.btLee = new DevExpress.XtraEditors.SimpleButton();
			this.btRepCliente = new DevExpress.XtraEditors.SimpleButton();
			this.chkAjusteIva = new Infragistics.Win.UltraWinEditors.UltraCheckEditor();
			this.cmbHasta = new Infragistics.Win.UltraWinEditors.UltraDateTimeEditor();
			this.ultraLabel1 = new Infragistics.Win.Misc.UltraLabel();
			this.ultraLabel2 = new Infragistics.Win.Misc.UltraLabel();
			this.chkExcluirVenta = new Infragistics.Win.UltraWinEditors.UltraCheckEditor();
			((System.ComponentModel.ISupportInitialize)(this.cmbFecha)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteo)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsCompra)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbHasta)).BeginInit();
			this.SuspendLayout();
			// 
			// btGenerar
			// 
			this.btGenerar.Location = new System.Drawing.Point(40, 120);
			this.btGenerar.Name = "btGenerar";
			this.btGenerar.Size = new System.Drawing.Size(96, 25);
			this.btGenerar.TabIndex = 0;
			this.btGenerar.Text = "Generar";
			this.btGenerar.Click += new System.EventHandler(this.btGenerar_Click);
			// 
			// cmbFecha
			// 
			this.cmbFecha.DateTime = new System.DateTime(2013, 11, 29, 0, 0, 0, 0);
			this.cmbFecha.FormatString = "dd/MM/yyyy";
			this.cmbFecha.Location = new System.Drawing.Point(120, 16);
			this.cmbFecha.Name = "cmbFecha";
			this.cmbFecha.PromptChar = ' ';
			this.cmbFecha.Size = new System.Drawing.Size(160, 22);
			this.cmbFecha.SpinButtonDisplayStyle = Infragistics.Win.ButtonDisplayStyle.Always;
			this.cmbFecha.TabIndex = 1;
			this.cmbFecha.Value = new System.DateTime(2013, 11, 29, 0, 0, 0, 0);
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
			this.btValidar.Location = new System.Drawing.Point(200, 120);
			this.btValidar.Name = "btValidar";
			this.btValidar.Size = new System.Drawing.Size(96, 25);
			this.btValidar.TabIndex = 2;
			this.btValidar.Text = "&Salir";
			this.btValidar.Click += new System.EventHandler(this.btValidar_Click);
			// 
			// btRepetido
			// 
			this.btRepetido.Location = new System.Drawing.Point(40, 152);
			this.btRepetido.Name = "btRepetido";
			this.btRepetido.Size = new System.Drawing.Size(96, 25);
			this.btRepetido.TabIndex = 3;
			this.btRepetido.Text = "Repetidos";
			this.btRepetido.Click += new System.EventHandler(this.btRepetido_Click);
			// 
			// ultraButton1
			// 
			this.ultraButton1.Location = new System.Drawing.Point(200, 152);
			this.ultraButton1.Name = "ultraButton1";
			this.ultraButton1.Size = new System.Drawing.Size(96, 25);
			this.ultraButton1.TabIndex = 4;
			this.ultraButton1.Text = "&Avanzado";
			this.ultraButton1.Click += new System.EventHandler(this.ultraButton1_Click);
			// 
			// btLee
			// 
			this.btLee.Location = new System.Drawing.Point(200, 184);
			this.btLee.Name = "btLee";
			this.btLee.Size = new System.Drawing.Size(96, 25);
			this.btLee.TabIndex = 5;
			this.btLee.Text = "&Lee Xml";
			this.btLee.Visible = false;
			this.btLee.Click += new System.EventHandler(this.btLee_Click);
			// 
			// btRepCliente
			// 
			this.btRepCliente.Location = new System.Drawing.Point(40, 184);
			this.btRepCliente.Name = "btRepCliente";
			this.btRepCliente.Size = new System.Drawing.Size(96, 25);
			this.btRepCliente.TabIndex = 6;
			this.btRepCliente.Text = "Rep. Clientes";
			this.btRepCliente.Click += new System.EventHandler(this.btRepCliente_Click);
			this.btRepCliente.MouseUp += new System.Windows.Forms.MouseEventHandler(this.btRepCliente_MouseDown);
			// 
			// chkAjusteIva
			// 
			this.chkAjusteIva.Checked = true;
			this.chkAjusteIva.CheckState = System.Windows.Forms.CheckState.Checked;
			this.chkAjusteIva.Location = new System.Drawing.Point(32, 88);
			this.chkAjusteIva.Name = "chkAjusteIva";
			this.chkAjusteIva.Size = new System.Drawing.Size(168, 21);
			this.chkAjusteIva.TabIndex = 7;
			this.chkAjusteIva.Text = "Recalculo de Iva en Servicios";
			this.chkAjusteIva.Visible = false;
			// 
			// cmbHasta
			// 
			this.cmbHasta.DateTime = new System.DateTime(2013, 11, 29, 0, 0, 0, 0);
			this.cmbHasta.FormatString = "dd/MM/yyyy";
			this.cmbHasta.Location = new System.Drawing.Point(120, 56);
			this.cmbHasta.Name = "cmbHasta";
			this.cmbHasta.PromptChar = ' ';
			this.cmbHasta.Size = new System.Drawing.Size(160, 22);
			this.cmbHasta.SpinButtonDisplayStyle = Infragistics.Win.ButtonDisplayStyle.Always;
			this.cmbHasta.TabIndex = 8;
			this.cmbHasta.Value = new System.DateTime(2013, 11, 29, 0, 0, 0, 0);
			// 
			// ultraLabel1
			// 
			this.ultraLabel1.AutoSize = true;
			this.ultraLabel1.Location = new System.Drawing.Point(56, 16);
			this.ultraLabel1.Name = "ultraLabel1";
			this.ultraLabel1.Size = new System.Drawing.Size(39, 15);
			this.ultraLabel1.TabIndex = 9;
			this.ultraLabel1.Text = "Desde:";
			// 
			// ultraLabel2
			// 
			this.ultraLabel2.AutoSize = true;
			this.ultraLabel2.Location = new System.Drawing.Point(56, 56);
			this.ultraLabel2.Name = "ultraLabel2";
			this.ultraLabel2.Size = new System.Drawing.Size(37, 15);
			this.ultraLabel2.TabIndex = 10;
			this.ultraLabel2.Text = "Hasta:";
			// 
			// chkExcluirVenta
			// 
			this.chkExcluirVenta.Location = new System.Drawing.Point(208, 88);
			this.chkExcluirVenta.Name = "chkExcluirVenta";
			this.chkExcluirVenta.Size = new System.Drawing.Size(96, 21);
			this.chkExcluirVenta.TabIndex = 11;
			this.chkExcluirVenta.Text = "Excluir Ventas";
			this.chkExcluirVenta.Visible = false;
			// 
			// AnexoTransaccional
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 14);
			this.ClientSize = new System.Drawing.Size(336, 224);
			this.Controls.Add(this.chkExcluirVenta);
			this.Controls.Add(this.ultraLabel2);
			this.Controls.Add(this.ultraLabel1);
			this.Controls.Add(this.cmbHasta);
			this.Controls.Add(this.chkAjusteIva);
			this.Controls.Add(this.btRepCliente);
			this.Controls.Add(this.btLee);
			this.Controls.Add(this.ultraButton1);
			this.Controls.Add(this.btRepetido);
			this.Controls.Add(this.btValidar);
			this.Controls.Add(this.cmbFecha);
			this.Controls.Add(this.btGenerar);
			this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle;
			this.KeyPreview = true;
			this.MaximizeBox = false;
			this.MinimizeBox = false;
			this.Name = "AnexoTransaccional";
			this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
			this.Text = "Anexo Transaccional Simplificado";
			this.KeyDown += new System.Windows.Forms.KeyEventHandler(this.AnexoTransaccional_KeyDown);
			this.Closing += new System.ComponentModel.CancelEventHandler(this.AnexoTransaccional_Closing);
			this.Load += new System.EventHandler(this.AnexoTransaccional_Load);
			((System.ComponentModel.ISupportInitialize)(this.cmbFecha)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteo)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsCompra)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbHasta)).EndInit();
			this.ResumeLayout(false);

		}
		#endregion

		private string NombreGenerado(string stArchivo)
		{
			DateTime dtFecha = (DateTime) cmbFecha.Value;
			stArchivo += dtFecha.ToString("MMyyyy") + ".XML";
			if (Directory.Exists(@"C:\XML_ANEXO_TRANSACCIONAL\")) stArchivo = @"C:\XML_ANEXO_TRANSACCIONAL\" + stArchivo;
			return stArchivo;
		}

		private int[] miLoteidPagoLot = new Int32[1000];
		private string[] miLoteStCodRenta = new string [1000];
		private double[] miLotedRentaPorc = new double [1000];
		private int miLoteiCuenta;
		private string[] miLoteStSerieRet = new string [1000];
		private string[] miLoteStAutRet = new string [1000];
		private string[] miLoteStCodRet = new string [1000];

		private void AnexoReoc2008()
		{
			Cursor = Cursors.WaitCursor;
			DateTime dtFecha = (DateTime) cmbFecha.Value;
			DateTime dtHasta = (DateTime) cmbHasta.Value;
			miLoteiCuenta = 0;
			Funcion.EjecutaSQL(cdsCompra, "Exec VerificaLote", true);

			#region Encabezado
			string stArchivo = NombreGenerado("REOC");
			XmlTextWriter miXML = new XmlTextWriter(stArchivo, 
				Encoding.GetEncoding("ISO-8859-1"));
			miXML.Formatting = Formatting.Indented;
			miXML.WriteStartDocument(false);
			miXML.WriteStartElement("reoc");
			string stMensaje = "";
			string stVerificaRuc = Funcion.VerificaCedula(drSeteo.Ruc, false);
			if (stVerificaRuc.Length == 0) miXML.WriteElementString("numeroRuc", drSeteo.Ruc);
			else stMensaje += "Error en RUC de su Empresa\n";

			miXML.WriteElementString("anio", dtFecha.ToString("yyyy"));
			miXML.WriteElementString("mes", dtFecha.ToString("MM"));

			if (stMensaje.Trim().Length > 0)
			{
				MessageBox.Show(stMensaje, "Información");
				using(Seteo miSeteo = new Seteo())
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
			string stExec = "Exec DevolucionImportaciones ";
			Funcion.EjecutaSQL(cdsCompra, stExec, true);
			string stFechaRet = "Retencion.Fecha";
			if (drSeteo.At_Fecha_Cont == true)
				stFechaRet = "Retencion.FechaContable";

			#region Compras
			#region Devolucion de compras
			miXML.WriteStartElement("compras");
			// Creacion de Devolucion de compras
			string stSelect = "SELECT IsNull(Compra.idCliente, 0), "
				+ "IsNull(Compra.Comprobante, ''), IsNull(Compra.idCompra, 0), "
				+ "IsNull(Compra.Comprob1, '') "
				+ "FROM Compra LEFT OUTER JOIN CompraComprobante "
				+ "ON Compra.idComprobante = CompraComprobante.idComprobante "
				+ "Where Compra.idTipoFactura = 6 And Compra.Borrar = 0 And Compra.Fecha >= '" 
				+ dtFecha.ToString("yyyyMMdd") + "' And Compra.Fecha < '" 
				+ dtHasta.AddDays(1).ToString("yyyyMMdd") + "'"
				+ " And Not CompraComprobante.Codigo In ('17', '99') "
				+ " Order By Compra.IdCompra";
			SqlDataReader miReader = Funcion.rEscalarSQL(cdsCompra, stSelect, true);
			int k = 0;
			int iMax = 4000;
			int [] IdDevCliente = new int [iMax];
			int [] IdDevCompra = new int [iMax];
			int [] IdCompra1 = new int [iMax];
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
				string stComp1 = miReader.GetString(3);
				if (stComp1.Trim().Length > 0)
					stDevComprob[k] = stComp1.Trim();
				k++;
			}
			miReader.Close();
			for (int j=0; j< k; j++)
			{
				stSelect = "Select Count(*) From Compra "
					+ "WHERE Compra.Numero = '" + stDevComprob[j].Trim()
					+ "' AND Compra.idCliente = " + IdDevCliente[j].ToString();
				int iCuenta = Funcion.iEscalarSQL(cdsCompra, stSelect, true);
				if (iCuenta == 0)
				{
					MessageBox.Show("Numero de Comprobante no existe",	"Información");
//					new Compra(6, IdDevCompra[j]).Show();
					miXML.Close();
					Cursor = Cursors.Default;
					return;
				}
				stSelect = "SELECT Count(*) FROM Compra "
					+ "Inner Join CompraComprobante On Compra.idComprobante "
					+ "= CompraComprobante.idComprobante "
					+ "WHERE Compra.Numero = '" + stDevComprob[j].Trim()
					+ "' And Compra.idTipoFactura = 4 AND Compra.idCliente = " + IdDevCliente[j].ToString();

				int iCuentaComp = Funcion.iEscalarSQL(cdsCompra, stSelect);
				if (iCuentaComp == 0)
				{
					MessageBox.Show("Comprobante Asociado no existe",	"Información");
//					new Compra(6, IdDevCompra[j]).Show();
					miXML.Close();
					Cursor = Cursors.Default;
					return;
				}

				stSelect = "SELECT IsNull(Compra.SerieFactura, ''), "
					+ "IsNull(Compra.Fecha, DateAdd(year, -100, GetDate())), "
					+ "IsNull(Compra.AutFactura, ''), "
					+ "IsNull(CompraComprobante.Codigo, ''), IsNull(Compra.idCompra, 0) FROM Compra "
					+ "Inner Join CompraComprobante On Compra.idComprobante "
					+ "= CompraComprobante.idComprobante "
					+ "WHERE Compra.Numero = '" + stDevComprob[j].Trim()
					+ "' And Compra.idTipoFactura = 4 AND Compra.idCliente = " + IdDevCliente[j].ToString();
				miReader = Funcion.rEscalarSQL(cdsCompra, stSelect, true);
				miReader.Read();
				stDevSerie[j] = miReader.GetString(0);
				stDevSerie[j] = Digito(stDevSerie[j]);
				dtDevFecha[j] = miReader.GetDateTime(1);
				stDevAutoriza[j] = miReader.GetString(2);
				stDevAutoriza[j] = Digito(stDevAutoriza[j]);
				stDevCompTipo[j] = miReader.GetString(3);
				stDevCompTipo[j] = Digito(stDevCompTipo[j]);
				IdCompra1[j] = miReader.GetInt32(4);
				miReader.Close();

				stDevComprob[j] = Digito(stDevComprob[j]);
				#endregion Devolucion de compras

				#region Verificaciones de devoluciones
				if (stDevCompTipo[j].Trim().Length < 2)
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
			stSelect = "Update Compra Set CuentaRetDet = "
				+ "(SELECT COUNT(*) FROM RETDETALLE INNER JOIN RETENCION ON RETDETALLE.IDRETENCION = RETENCION.IDRETENCION "
				+ "WHERE RETENCION.IDCOMPRA = COMPRA.IDCOMPRA  AND RETENCION.INGRESOEGRESO = 2) "
				+ "WHERE Compra.Fecha >= '" + dtFecha.ToString("yyyyMMdd")
				+ "' AND COMPRA.FECHA < '" + dtHasta.AddDays(1).ToString("yyyyMMdd")
				+ "' AND COMPRA.IDTIPOFACTURA = 4 ";
			Funcion.EjecutaSQL(cdsCompra, stSelect, true);

			stSelect = "SELECT IsNull(CompraCredTribut.Codigo, ''), "
				+ "Compra.DerechoDevolucion, IsNull(ClienteTipoRuc.Codigo, '') ,"
				+ "IsNull(substring(Cliente.Ruc, 1, 13), ''), IsNull(CompraComprobante.Codigo, ''), "
				+ "IsNull(Compra.Numero, ''), IsNull(Cliente.idCliente,0), "
				+ "IsNull(Compra.SerieFactura, ''), IsNull(Compra.AutFactura, ''), "
				+ "IsNull(Compra.FechaCaducidad, DateAdd(year, -100,GetDate())), "
				+ "IsNull(Compra.SubtotalExcento - Compra.Descuento0, 0), "
				+ "IsNull(Compra.SubtotalIva - Compra.Descuento, 0), "
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
				+ "IsNull( " + stFechaRet + ", DateAdd(year, -100,GetDate())), "
				+ "IsNull(Compra.CuentaRetDet, 0), IsNull(Compra.idTipoFactura, 0), "
				+ "IsNull(Retencion.IvaTotal, 0), IsNull(Retencion.IvaTotal1, 0), "
				+ "IsNull(Retencion.Tarjeta, 0), IsNull(Retencion.NoAplica, 0), "
				+ "IsNull(Compra.Retenido, 0), IsNull(Compra.IdPagoLot, 0), "
				+ "IsNull(Compra.Ice, 0), IsNull(PorcIce.porcentaje, 0) "
				+ "FROM Compra LEFT OUTER JOIN Retencion ON Retencion.idCompra = Compra.idCompra "
				+ "LEFT OUTER JOIN Cliente ON Compra.idCliente = Cliente.idCliente "
				+ "LEFT OUTER JOIN CompraCredTribut ON Compra.idCredTributario = CompraCredTribut.idCredTributario "
				+ "LEFT OUTER JOIN CompraComprobante ON Compra.idComprobante = CompraComprobante.idComprobante "
				+ "LEFT OUTER JOIN ClienteTipoRuc ON Cliente.idTipoRuc = ClienteTipoRuc.idTipoRuc "
				+ "LEFT OUTER JOIN RetencionIva ON Retencion.idRetIva = RetencionIva.idRetIva "
				+ "LEFT OUTER JOIN RetencionIva As RetIva_1  ON Retencion.idRetIva1 = RetIva_1.idRetIva "
				+ "LEFT OUTER JOIN RetDetalle ON RetDetalle.idRetencion = Retencion.idRetencion "
				+ "LEFT OUTER JOIN ArticuloSri ON RetDetalle.idRetRenta = ArticuloSri.idArticuloSri "
				+ "LEFT JOIN porcice ON compra.idporcice = porcice.idPorcIce "
				+ "Where Compra.idTipoFactura In (4, 6) And Compra.Borrar = 0 And Compra.Fecha >= '" 
				+ dtFecha.ToString("yyyyMMdd") + "' And Compra.Fecha < '" + dtHasta.AddDays(1).ToString("yyyyMMdd") + "'"
				+ " And Not CompraComprobante.Codigo In ('17', '99') "
				+ " Order By Compra.IdCompra";
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
			double dFacIva = 0;
			double dFacServ = 0;
			bool bTarjeta = false;
			bool bNoAplica = false;
			double dRetenidoLote = 0;
			int idPagoLot = 0;
			double dCompraIce = 0;
			double dPorcIce = 0;
			#endregion fin de definicion de variables
			#region Lectura de Filas de Compras
			while (miReader.Read())
			{
				IdTipoFactura = miReader.GetInt32(32);
				IdCompra = miReader.GetInt32(13);
				miXML.WriteStartElement("detalleCompras");
				stCodCredTrib = miReader.GetString(0);
				bDerechoDev = miReader.GetBoolean(1);
				stCodTipoRuc = miReader.GetString(2);
				stRuc = miReader.GetString(3).Trim();
				string stPasaporte = stRuc;
				IdCliente = miReader.GetInt32(6);
				stRuc = Digito(stRuc);
				if (IdCliente < 1)
				{
					MessageBox.Show("Proveedor no definido", "Información");
//					new Compra(IdTipoFactura, IdCompra).Show();
					miXML.Close();
					Cursor = Cursors.Default;
					return;
				}

				string stVerifica = Funcion.VerificaCedula(stRuc, false);
				if (stRuc.Length != 10 && (stCodTipoRuc == "02" || stCodTipoRuc == "05")) stVerifica = "Err";
//				if (stCodTipoRuc == "01" || stCodTipoRuc == "04") stRuc += "000";
				if (stRuc.Length != 13 && (stCodTipoRuc == "01" || stCodTipoRuc == "04"))	stVerifica = "Err";
				if (stCodTipoRuc == "03" || stCodTipoRuc == "06" || stCodTipoRuc == "07" || stCodTipoRuc == "08") stVerifica = "";
				if (stVerifica.Length > 0)
				{
					MessageBox.Show("Error en Cedula o RUC", "Información");
					new Cliente(false, IdCliente).Show();
					miXML.Close();
					Cursor = Cursors.Default;
					return;
				}
				if (stCodTipoRuc.Length != 2)
				{
					MessageBox.Show("Error en Tipo Cedula o RUC", "Información");
					new Cliente(false, IdCliente).Show();
					miXML.Close();
					Cursor = Cursors.Default;
					return;
				}
				if (stRuc.Length == 13 && (stCodTipoRuc == "01" || stCodTipoRuc == "04"))
					stRuc = stRuc.Substring(0,10) + "001";
				stCodComprob =  miReader.GetString(4);
				stNumero =  miReader.GetString(5);
				stCompraSerieFact = miReader.GetString(7);
				stAutFactura = miReader.GetString(8);
				dtCaducidad =  miReader.GetDateTime(9);
				dSubtotalExcento = miReader.GetDouble(10);
				dSubtotalIva = miReader.GetDouble(11);
				dCompraIva = miReader.GetDouble(12);
				dRetIvaPorc = miReader.GetDouble(14);
				stCodRetIva = miReader.GetString(15);
				if (Math.Round(dRetIvaPorc, 0) == 0) stCodRetIva = "0";
				if (Math.Round(dRetIvaPorc, 0) == 30) stCodRetIva = "1";
				if (Math.Round(dRetIvaPorc, 0) == 70) stCodRetIva = "2";
				if (Math.Round(dRetIvaPorc, 0) == 100) stCodRetIva = "3";

				dRetIvaPorc1 = miReader.GetDouble(16);
				stCodRetIva1 = miReader.GetString(17);
				if (Math.Round(dRetIvaPorc1, 0) == 0) stCodRetIva1 = "0";
				if (Math.Round(dRetIvaPorc1, 0) == 30) stCodRetIva1 = "1";
				if (Math.Round(dRetIvaPorc1, 0) == 70) stCodRetIva1 = "2";
				if (Math.Round(dRetIvaPorc1, 0) == 100) stCodRetIva1 = "3";

				dtFechaCompra = miReader.GetDateTime(18);
				dRetValorIva = miReader.GetDouble(19);
				dRetValorIva1 = miReader.GetDouble(20);
				IdRetencion = miReader.GetInt32(21);
				dIvaServicio = miReader.GetDouble(22);
				iCuentaRet = miReader.GetInt32(31);
				dFacIva = miReader.GetDouble(33);
				dFacServ = miReader.GetDouble(34);
				bTarjeta = miReader.GetBoolean(35);
				bNoAplica = miReader.GetBoolean(36);
				dRetenidoLote = miReader.GetDouble(37);
				idPagoLot = miReader.GetInt32(38);
				dCompraIce = miReader.GetDouble(39);
				dPorcIce = miReader.GetDouble(40);
				#endregion Lectura de Filas de Compras
			#region Proveedor
				if (stCodTipoRuc == "07" || stCodTipoRuc == "08") 
					stCodTipoRuc = "03";

				if (stCodTipoRuc == "07" || stCodTipoRuc == "08") 
					stCodTipoRuc = "03";

				if (int.Parse(stCodTipoRuc) > 3)
				{
					int iCodTipoRuc = int.Parse(stCodTipoRuc) - 3;
					stCodTipoRuc = "0" + iCodTipoRuc.ToString().Trim();
				}
				miXML.WriteElementString("tpIdProv", stCodTipoRuc);
				if (stRuc.Trim().Length < 2)
				{
					MessageBox.Show("Codigo de RUC no existe",
						"Información");
//					new Compra(IdTipoFactura, IdCompra).Show();
					miXML.Close();
					Cursor = Cursors.Default;
					return;
				}
				miXML.WriteElementString("idProv", stRuc);
				if (stCodComprob.Trim().Length < 2)
				{
					MessageBox.Show("Codigo de Comprobante no existe",
						"Información");
//					new Compra(IdTipoFactura, IdCompra).Show();
					miXML.Close();
					Cursor = Cursors.Default;
					return;
				}

				bool bExterior = false;
				if (stCodComprob.Trim() == "15") bExterior = true;

				if ((stCodTipoRuc == "02") 
					&& !(stCodComprob.Trim() == "03" || stCodComprob.Trim() == "04"))
				{
					MessageBox.Show("Con cedula solo se permite liquidacion de compras",
						"Información");
//					new Compra(IdTipoFactura, IdCompra).Show();
					miXML.Close();
					Cursor = Cursors.Default;
					return;
				}

				if (stCodComprob.Trim() == "15" && stCodTipoRuc != "03")
				{
					MessageBox.Show("Con Comprobantes emitidos en el exterior solo puede usar pasaporte",
						"Información");
//					new Compra(IdTipoFactura, IdCompra).Show();
					miXML.Close();
					Cursor = Cursors.Default;
					return;
				}

				if ((stCodTipoRuc == "03") 
					&& !(stCodComprob.Trim() == "03" || stCodComprob.Trim() == "04" || stCodComprob.Trim() == "15"))
				{
					MessageBox.Show("Con Pasaporte solo se permite liquidacion de compras y Comprobantes emitidos en el exterior",
						"Información");
//					new Compra(IdTipoFactura, IdCompra).Show();
					miXML.Close();
					Cursor = Cursors.Default;
					return;
				}

				if ((stCodTipoRuc == "01") && stCodComprob.Trim() == "03")
				{
					MessageBox.Show("Con RUC No se permite liquidacion de compras",
						"Información");
//					new Compra(IdTipoFactura, IdCompra).Show();
					miXML.Close();
					Cursor = Cursors.Default;
					return;
				}
				
				#endregion Fin de Proveedor
			#region Comprobante, Autorizacion, Serie, Fecha Emision
				// Codigo Tipo de Comprobante
				if (stCodComprob.Trim() == "02")
					if (stCodCredTrib.Trim() != "02")
					{
						MessageBox.Show("Nota de Venta no sustenta Credito Tributario", "Información");
//						new Compra(IdTipoFactura, IdCompra).Show();
						miXML.Close();
						Cursor = Cursors.Default;
						return;
					}
				if (stCodComprob.Trim() == "15")
					if (stCodCredTrib.Trim() != "02")
					{
						MessageBox.Show("Comprobantes en el exterior no sustentan Credito Tributario", "Información");
//						new Compra(IdTipoFactura, IdCompra).Show();
						miXML.Close();
						Cursor = Cursors.Default;
						return;
					}

//					if (stCodComprob.Trim() == "03" || stCodComprob.Trim() == "12" || stCodComprob.Trim() == "11" || stCodComprob.Trim() == "11" || stCodComprob.Trim() == "11")
				if (stCodCredTrib.Trim() == "00"	&& stCodTipoRuc != "01")
					{
						MessageBox.Show("Crédito Tributario 'No Aplica' solo se usa en ruc", "Información");
//						new Compra(IdTipoFactura, IdCompra).Show();
						miXML.Close();
						Cursor = Cursors.Default;
						return;
					}

				if (stCodCredTrib.Trim() == "00"	&& stCodTipoRuc == "01")
					if (!(stCodComprob.Trim() == "01" || stCodComprob.Trim() == "02" || stCodComprob.Trim() == "11"))
					{
						MessageBox.Show("Crédito Tributario 'No Aplica' solo se usa en: Facturas, Notas de Venta, y Documentos Agente de retencion presuntiva", "Información");
//						new Compra(IdTipoFactura, IdCompra).Show();
						miXML.Close();
						Cursor = Cursors.Default;
						return;
					}

				miXML.WriteElementString("tipoComp", int.Parse(stCodComprob).ToString());

				if (!bExterior)
				{
					stAutFactura = stAutFactura.Trim();
					stAutFactura = Digito(stAutFactura);
					if (stAutFactura.Length != 10)
					{
						MessageBox.Show("Error en Numero de Autorizacion de Factura", "Información");
//						new Compra(IdTipoFactura, IdCompra).Show();
						miXML.Close();
						Cursor = Cursors.Default;
						return;
					}
					miXML.WriteElementString("aut", stAutFactura.ToString());
				}
				else
				{
					miXML.WriteElementString("aut", "9999999999");
				}

				stNumero = Digito(stNumero);
				string stEstablecmiento = "";
				string stPuntoEmision = "";
				string stSecuencial = stNumero;
				if (!bExterior)
				{
					if (stNumero.Trim().Length > 9 && stNumero.Trim().Length < 15)
					{
						MessageBox.Show("Error en Numero o Serie de Factura", "Información");
//						new Compra(IdTipoFactura, IdCompra).Show();
						miXML.Close();
						Cursor = Cursors.Default;
						return;
					}

					if (stNumero.Trim().Length > 9)
					{
						stEstablecmiento = stNumero.Substring(0, 3);
						stPuntoEmision = stNumero.Substring(3,3);
						stSecuencial = stNumero.Substring(6);
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
					if (stEstablecmiento.Length != 3  || stPuntoEmision.Length != 3 
						|| stSecuencial.Length == 0 || stSecuencial.Length > 9)
					{
						MessageBox.Show("Error en Numero o Serie de Factura", "Información");
//						new Compra(IdTipoFactura, IdCompra).Show();
						miXML.Close();
						Cursor = Cursors.Default;
						return;
					}
					miXML.WriteElementString("estab", stEstablecmiento);
					miXML.WriteElementString("ptoEmi", stPuntoEmision);
					miXML.WriteElementString("sec", stSecuencial);
				}
				else
				{
					miXML.WriteElementString("estab", "001");
					miXML.WriteElementString("ptoEmi", "001");
					if (stNumero.Trim().Length > 7) stSecuencial = stNumero.Substring(6);
					miXML.WriteElementString("sec", stSecuencial);
				}

				miXML.WriteElementString("fechaEmiCom", dtFechaCompra.ToString("dd/MM/yyyy"));
				#endregion fin de serie
				DateTime dtFechaRet = miReader.GetDateTime(30);

			#region Impuesto a la Renta Retenido
				miXML.WriteStartElement("air");
				double dRentaRetTotal = 0;
				if (idPagoLot > 0) iCuentaRet = 1;
				bool bRetPeriodo = true;

				if (!(bTarjeta || bNoAplica || !bRetPeriodo))
				{
					double[] dAcRentaRet = new double[iCuentaRet];
					double[] dAcRentaBase = new double[iCuentaRet];
					double[] dAcRentaPorc = new double[iCuentaRet];
					string[] stAcCodRenta = new string[iCuentaRet];
					for (int i=0; i<iCuentaRet; i++)
					{
						double dRentaRet = miReader.GetDouble(23);
						double dRentaBase = miReader.GetDouble(24);
						double dRentaPorc = miReader.GetDouble(25);
						string stCodRenta = miReader.GetString(26);

						#region  Retencion de varias Facturas
						if (idPagoLot > 0)
						{
							dRentaRet = dRetenidoLote;
							if (dRentaPorc >= 0.01 && dRentaRet >= 0.01) 
							{
								dRentaBase = dSubtotalExcento + dSubtotalIva;
								dRentaRet = dRentaPorc * dRentaBase / 100.0;
								miLoteidPagoLot[miLoteiCuenta] = idPagoLot;
								miLoteStCodRenta[miLoteiCuenta] = stCodRenta;
								miLotedRentaPorc[miLoteiCuenta] = dRentaPorc;
								miLoteiCuenta ++;
							}
							else // Busca Primer idPagoLot
							{
								for (int iLot =0; iLot< miLoteiCuenta; iLot++)
								{
									if (miLoteidPagoLot[iLot] == idPagoLot)
									{
										stCodRenta = miLoteStCodRenta[iLot];
										dRentaPorc = miLotedRentaPorc[iLot];
										break;
									} //Fin de if
								} // Fin de For
								dRentaBase = dSubtotalExcento + dSubtotalIva;
								dRentaRet = dRentaPorc * dRentaBase / 100.0;
							} // Busca Primer idPagoLot
						}
						#endregion

						if (dRentaBase < .0001) 
						{
							if (i < iCuentaRet - 1) miReader.Read();
							continue;
						}
						dRentaRetTotal += dRentaRet;

						stAcCodRenta[i] = stCodRenta;
						dAcRentaBase[i] = dRentaBase;
						dAcRentaPorc[i] = dRentaPorc;
						dAcRentaRet[i] = dRentaRet;
						if (i < iCuentaRet - 1) miReader.Read();
					}

					for (int i=0; i<iCuentaRet; i++)
					{
						for (int j=i + 1; j<iCuentaRet; j++)
						{
							if (stAcCodRenta[i] == stAcCodRenta[j])
							{
								dAcRentaBase[i] += dAcRentaBase[j];
								dAcRentaRet[i] += dAcRentaRet[j];
								stAcCodRenta[j] = "NoE";
							}
						}
					}
					for (int i=0; i<iCuentaRet; i++)
					{
						if (stAcCodRenta[i] == null) continue;
						if (stAcCodRenta[i] == "NoE") continue;
						if (stAcCodRenta[i].Trim() == "0")
						{
							MessageBox.Show("Error en Codigo de Retencion", "Información");
//							new Compra(IdTipoFactura, IdCompra).Show();
							miXML.Close();
							Cursor = Cursors.Default;
							return;
						}
						miXML.WriteStartElement("detalleAir");
						miXML.WriteElementString("codRetAir", stAcCodRenta[i].Trim());
						miXML.WriteElementString("porcentaje", Math.Round(dAcRentaPorc[i], 0).ToString("0.0", us));
//						if (dSubtotalExcento > 0)
						if (dSubtotalExcento > 0 && Math.Abs(dAcRentaBase[i] - dSubtotalExcento) < 0.01)
						{
							miXML.WriteElementString("base0", dAcRentaBase[i].ToString("0.00", us));
							miXML.WriteElementString("baseGrav", "0.00");
						}
						else
						{
							miXML.WriteElementString("base0", "0.00");
							miXML.WriteElementString("baseGrav", dAcRentaBase[i].ToString("0.00", us));
						}
						miXML.WriteElementString("baseNoGrav", "0.00");
						miXML.WriteElementString("valRetAir", dAcRentaRet[i].ToString("0.00", us));
						miXML.WriteEndElement();
						// Fin de detalleAir
					}
				}
				miXML.WriteEndElement(); // Fin de air
				#endregion Impuesto a la renta
			#region Codigos de Retencion
				if ((iCuentaRet > 0 || bRetPeriodo) && !(bTarjeta || bNoAplica) )
				{
					string stSerieRet = miReader.GetString(27);
					string stAutRet = miReader.GetString(28);
					string stCodRet = miReader.GetString(29);
					if (idPagoLot > 0)
					{
						if (stSerieRet.Trim().Length > 0)
						{
							miLoteStSerieRet[miLoteiCuenta-1] = stSerieRet;
							miLoteStAutRet[miLoteiCuenta-1] = stAutRet;
							miLoteStCodRet[miLoteiCuenta-1] = stCodRet;
						}
						else
						{
							for (int iLot =0; iLot< miLoteiCuenta; iLot++)
							{
								if (miLoteidPagoLot[iLot] == idPagoLot)
								{
									stSerieRet = miLoteStSerieRet[iLot];
									stAutRet = miLoteStAutRet[iLot];
									stCodRet = miLoteStCodRet[iLot];
									break;
								} //Fin de if
							} // Fin de For
						} // fin de else
					}
					stSerieRet = Digito(stSerieRet);
					stAutRet = Digito(stAutRet);
					stCodRet = Digito(stCodRet);

					if (dRentaRetTotal >= 0.01 && (stSerieRet.Length != 6 || stAutRet.Length != 10 || stCodRet.Length < 1))
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
					// Fecha de Emision de Retencion mayor o igual qu efactura
					DateTime dtFechaNula = DateTime.Today.AddYears(-100);
					dtFechaRet = dtFechaRet.Date;
					dtFechaCompra = dtFechaCompra.Date;
					if (dtFechaRet < dtFechaCompra && 
						!(dtFechaRet > dtFechaNula.AddDays(-1) && dtFechaRet < dtFechaNula.AddDays(1)))
					{
						MessageBox.Show("Fecha de Retencion debe ser mayor que " + dtFechaCompra.ToString("dd/MMM/yyyy"), "Información");
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
					if (stSerieRet.Trim().Length == 6 && stCodRet.Trim().Length > 0
						&& (dRentaRetTotal >= 0.01)) // || bRetPeriodo))
					{
						miXML.WriteElementString("autRet", stAutRet);
						miXML.WriteElementString("estabRet", stSerieRet.Substring(0,3));
						miXML.WriteElementString("ptoEmiRet", stSerieRet.Substring(3,3));
						miXML.WriteElementString("secRet", stCodRet);
						miXML.WriteElementString("fechaEmiRet", dtFechaRet.ToString("dd/MM/yyyy"));
					}
				}
				#endregion Codigos de Retencion

			#region Devoluciones de compras
				if (IdTipoFactura == 6)
				{
					try
					{
						if (stDevComprob[k].Trim().Length > 15 || stDevComprob[k].Trim().Length < 1)
						{
							MessageBox.Show("Error en Comprobante Asociado", "Información");
//							new Compra(6, IdCompra1[k]).Show();
							miXML.Close();
							Cursor = Cursors.Default;
							return;
						}

						if (stDevComprob[k].Trim().Length == 15)
						{
							stDevComprob[k] = stDevComprob[k].Substring(6, 9);
						}
						int iSerie = int.Parse(stDevSerie[k].Substring(0,3));
						int iPuntoE = int.Parse(stDevSerie[k].Substring(3,3));
						if (iSerie < 1 || iSerie > 999 || iPuntoE < 1 || iPuntoE > 999)
						{
							MessageBox.Show("Error en Serie de Comprobante Asociado", "Información");
//							new Compra(6, IdCompra).Show();
							miXML.Close();
							Cursor = Cursors.Default;
							return;
						}
						if (Int64.Parse(stDevAutoriza[k]) == 0 || stDevAutoriza[k].Length != 10)
						{
							MessageBox.Show("Error en Autorizacion de Comprobante Asociado", "Información");
//							new Compra(6, IdCompra).Show();
							miXML.Close();
							Cursor = Cursors.Default;
							return;
						}
						if (Int64.Parse(stDevComprob[k]) == 0 || stDevComprob[k].Length > 9)
						{
							MessageBox.Show("Error en Secuencial de Comprobante Asociado: " + stDevComprob[k], "Información");
//							new Compra(6, IdCompra).Show();
							miXML.Close();
							Cursor = Cursors.Default;
							return;
						}

					}
					catch (System.Exception ex)
					{
						MessageBox.Show(ex.Message, "No existe comprobante");
//						new Compra(6, IdCompra1[k] IdCompra).Show();
//						new Compra(6, IdCompra1[k]).Show();
						miXML.Close();
						Cursor = Cursors.Default;
						return;
					}
					k++;
				}
				#endregion fin de devoluciones de compras
				miXML.WriteEndElement(); // Fin de DetalleCompras
			}
			miReader.Close();
			#endregion Fin de Compras

			#region Retencion de Tarjetas en Compras
			stSelect = "Select Max(Cl.Ruc), Sum(IsNull(R.RentaRetenida, 0) "
				+ "+ IsNull(R.RentaRetenida1, 0)), "
				+ "Sum(IsNull(R.IvaTotal, 0)), Sum(IsNull(R.IvaRetenido, 0)), "
				+ "Max(IsNull(R.idRetIva, 0)), Sum(IsNull(R.IvaTotal1, 0)), "
				+ "Sum(IsNull(R.IvaRetenido1, 0)), Max(IsNull(R.idRetIva1, 0)) "
				+ "From Retencion r Inner Join Cliente cl On R.idCliente = Cl.idCliente "
				+ "where R.Tarjeta = 1 And R.NoAplica = 0 "
				+ "And (IsNull(R.IvaRetenido, 0) +IsNull(R.IvaRetenido1, 0) >= 0) "
				+ "And R.IngresoEgreso = 2 "
				+ "group by Cl.Ruc, R.idRetIva, R.idRetIva1";
			miReader = Funcion.rEscalarSQL(cdsCompra, stSelect, true);
			while(miReader.Read())
			{
				string stRucTarjeta = miReader.GetString(0);
				double dRentaTarjeta = miReader.GetDouble(1);
				double dIvaTarjeta = miReader.GetDouble(2);
				double dIvaRetTarjeta = miReader.GetDouble(3);
				int iIdRetTarjeta = miReader.GetInt32(4);
				double dIva1Tarjeta = miReader.GetDouble(5);
				double dIvaRet1Tarjeta = miReader.GetDouble(6);
				int iIdRet1Tarjeta = miReader.GetInt32(7);
			}
			miReader.Close();
			#endregion Retencion de Tarjetas en Compras

			string[] stCodTipoRuc1 = new string[iMax];
			string[] stRuc1 = new string[iMax];
			string [] stComprobante = new string[iMax];

			double[] dIvaBienes1 = new double[iMax];

			miXML.WriteEndElement(); // Fin de Compras
			miXML.WriteEndElement(); // Fin de IVA
			miXML.Flush();
			miXML.Close();
			Cursor = Cursors.Default;
			MessageBox.Show("Archivo Generado: " + stArchivo, "Información");
		}

		private void AnexoReocSql2008()
		{
			Cursor = Cursors.WaitCursor;
			DateTime dtFecha = (DateTime) cmbFecha.Value;
			DateTime dtHasta = (DateTime) cmbHasta.Value;

			#region Encabezado
			string stArchivo = NombreGenerado("REOC");
			XmlTextWriter miXML = new XmlTextWriter(NombreGenerado("REOC"), 
				Encoding.GetEncoding("ISO-8859-1"));
			miXML.Formatting = Formatting.Indented;
			miXML.WriteStartDocument(false);
			miXML.WriteStartElement("reoc");
			string stMensaje = "";
			string stVerificaRuc = Funcion.VerificaCedula(drSeteo.Ruc, false);
			if (stVerificaRuc.Length == 0) miXML.WriteElementString("numeroRuc", drSeteo.Ruc);
			else stMensaje += "Error en RUC de su Empresa\n";

			miXML.WriteElementString("anio", dtFecha.ToString("yyyy"));
			miXML.WriteElementString("mes", dtFecha.ToString("MM"));

			if (stMensaje.Trim().Length > 0)
			{
				MessageBox.Show(stMensaje, "Información");
				using(Seteo miSeteo = new Seteo())
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
			#region Devolucion de compras
			miXML.WriteStartElement("compras");
			// Creacion de Devolucion de compras
			string stSelect = "SELECT IsNull(Compra.idCliente, 0), "
				+ "IsNull(Compra.Comprobante, ''), IsNull(Compra.idCompra, 0), "
				+ "IsNull(Compra.Comprob1, '') "
				+ "FROM Compra LEFT OUTER JOIN CompraComprobante "
				+ "ON Compra.idComprobante = CompraComprobante.idComprobante "
				+ "Where Compra.idTipoFactura = 6 And Compra.Borrar = 0 And Compra.Fecha >= '" 
				+ dtFecha.ToString("yyyyMMdd") + "' And Compra.Fecha < '" 
				+ dtHasta.AddDays(1).ToString("yyyyMMdd") + "'"
				+ " And Not CompraComprobante.Codigo In ('17', '99', '15') "
				+ " Order By Compra.IdCompra";
			SqlDataReader miReader = Funcion.rEscalarSQL(cdsCompra, stSelect, true);
			int k = 0;
			int iMax = 4000;
			int [] IdDevCliente = new int [iMax];
			int [] IdDevCompra = new int [iMax];
			int [] IdCompra1 = new int [iMax];
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
				string stComp1 = miReader.GetString(3);
				if (stComp1.Trim().Length > 0)
					stDevComprob[k] = stComp1.Trim();
				k++;
			}
			miReader.Close();
			for (int j=0; j< k; j++)
			{
				stSelect = "Select Count(*) From Compra "
					+ "WHERE Compra.Numero = '" + stDevComprob[j].Trim()
					+ "' AND Compra.idCliente = " + IdDevCliente[j].ToString();
				int iCuenta = Funcion.iEscalarSQL(cdsCompra, stSelect, true);
				if (iCuenta == 0)
				{
					MessageBox.Show("Numero de Comprobante no existe",	"Información");
//					new Compra(6, IdDevCompra[j]).Show();
					miXML.Close();
					Cursor = Cursors.Default;
					return;
				}
				stSelect = "SELECT IsNull(Compra.SerieFactura, ''), "
					+ "IsNull(Compra.Fecha, DateAdd(year, -100, GetDate())), "
					+ "IsNull(Compra.AutFactura, ''), "
					+ "IsNull(CompraComprobante.Codigo, ''), IsNull(Compra.idCompra, 0) FROM Compra "
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
				IdCompra1[j] = miReader.GetInt32(4);
				miReader.Close();

				stDevComprob[j] = Digito(stDevComprob[j]);

				#region Verificaciones de devoluciones
				if (stDevCompTipo[j].Trim().Length < 2)
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
			#endregion Devolucion de compras
			#region Select de Compras
			stSelect = "Exec Reoc '" + dtFecha.ToString("yyyyMMdd") + "', '" + dtHasta.AddDays(1).ToString("yyyyMMdd") + "'";
			#endregion Select de Compras
			#region Definicion de Variables de compras
			miReader = Funcion.rEscalarSQL(cdsCompra, stSelect, true);
			string stCodTipoRuc = "";
			string stRuc = "";
			string stCodComprob = "";
			string stAutFactura = "";
			string stCompraSerieFact = "";
			string stNumero = "";
			double dSubtotalExcento = 0;
			double dSubtotalIva = 0;
			int IdCompra = 0;
			DateTime dtFechaCompra = DateTime.Today;
			int IdRetencion = 0;
			int iCuentaRet = 0;
			int IdCliente = 0;
			int IdTipoFactura = 0;
			bool bRetNoAplica = false;
			#endregion fin de definicion de variables
			#region Lectura de Filas de Compras
			while (miReader.Read())
			{
				miXML.WriteStartElement("detalleCompras");
				stCodTipoRuc = miReader.GetString(0);
				stRuc = miReader.GetString(1).Trim();
				IdCliente = miReader.GetInt32(8);
				stRuc = Digito(stRuc);
				if (IdCliente < 1)
				{
					MessageBox.Show("Proveedor no definido", "Información");
//					new Compra(IdTipoFactura, IdCompra).Show();
					miXML.Close();
					Cursor = Cursors.Default;
					return;
				}

				string stVerifica = Funcion.VerificaCedula(stRuc, false);
				if (stRuc.Length != 10 && stCodTipoRuc == "02") stVerifica = "Err";
				//if (stCodTipoRuc == "01") stRuc += "000";
				if (stRuc.Length != 13 && (stCodTipoRuc == "01" || stCodTipoRuc == "04"))	stVerifica = "Err";
				if (stCodTipoRuc == "03" || stCodTipoRuc == "06" || stCodTipoRuc == "07" || stCodTipoRuc == "08") stVerifica = "";
				if (stVerifica.Length > 0)
				{
					MessageBox.Show("Error en Cedula o RUC", "Información");
					new Cliente(false, IdCliente).Show();
					miXML.Close();
					Cursor = Cursors.Default;
					return;
				}
				if (stCodTipoRuc.Length != 2)
				{
					MessageBox.Show("Error en Tipo Cedula o RUC", "Información");
					new Cliente(false, IdCliente).Show();
					miXML.Close();
					Cursor = Cursors.Default;
					return;
				}
				if (stRuc.Length == 13 && stCodTipoRuc == "01") 
					stRuc = stRuc.Substring(0,10) + "001";
				stCodComprob =  miReader.GetString(2);
				stAutFactura =  miReader.GetString(3);
				stCompraSerieFact =  miReader.GetString(4);
				stNumero =  miReader.GetString(5);
				dSubtotalExcento = miReader.GetDouble(6);
				dSubtotalIva = miReader.GetDouble(7);

				IdCompra = miReader.GetInt32(9);
				dtFechaCompra = miReader.GetDateTime(10);
				IdRetencion = miReader.GetInt32(11);
				IdTipoFactura = miReader.GetInt32(19);
				iCuentaRet = miReader.GetInt32(20);
				bRetNoAplica = miReader.GetBoolean(21);
				#endregion Lectura de Filas de Compras
				#region Proveedor
				if (stCodTipoRuc == "07" || stCodTipoRuc == "08") 
					stCodTipoRuc = "03";
				if (int.Parse(stCodTipoRuc) > 3)
				{
					int iCodTipoRuc = int.Parse(stCodTipoRuc) - 3;
					stCodTipoRuc = "0" + iCodTipoRuc.ToString().Trim();
				}
				miXML.WriteElementString("tpIdProv", stCodTipoRuc);
				if (stRuc.Trim().Length < 2)
				{
					MessageBox.Show("Codigo de RUC no existe",
						"Información");
//					new Compra(IdTipoFactura, IdCompra).Show();
					miXML.Close();
					Cursor = Cursors.Default;
					return;
				}
				miXML.WriteElementString("idProv", stRuc);
				if (stCodComprob.Trim().Length < 2)
				{
					MessageBox.Show("Codigo de Comprobante no existe",
						"Información");
//					new Compra(IdTipoFactura, IdCompra).Show();
					miXML.Close();
					Cursor = Cursors.Default;
					return;
				}

				bool bExterior =  false;
				if (stCodComprob.Trim() == "15") bExterior = true;

				if (stCodComprob.Trim() == "15" && stCodTipoRuc != "03")
				{
					MessageBox.Show("Con Comprobantes emitidos en el exterior solo puede usar pasaporte",
						"Información");
//					new Compra(IdTipoFactura, IdCompra).Show();
					miXML.Close();
					Cursor = Cursors.Default;
					return;
				}

				if ((stCodTipoRuc == "02") && 
					!(stCodComprob.Trim() == "03" || stCodComprob.Trim() == "04"))
				{
					MessageBox.Show("Con cédula sólo se permite liquidación de compras",
						"Informacion");
//					new Compra(IdTipoFactura, IdCompra).Show();
					miXML.Close();
					Cursor = Cursors.Default;
					return;
				}

				if ((stCodTipoRuc == "03") 
					&& !(stCodComprob.Trim() == "03" || stCodComprob.Trim() == "04" || stCodComprob.Trim() == "15"))
				{
					MessageBox.Show("Con Pasaporte solo se permite liquidacion de compras y Comprobantes emitidos en el exterior",
						"Información");
//					new Compra(IdTipoFactura, IdCompra).Show();
					miXML.Close();
					Cursor = Cursors.Default;
					return;
				}

				if ((stCodTipoRuc == "01") && stCodComprob.Trim() == "03")
				{
					MessageBox.Show("Con RUC No se permite liquidacion de compras",
						"Información");
//					new Compra(IdTipoFactura, IdCompra).Show();
					miXML.Close();
					Cursor = Cursors.Default;
					return;
				}
				
				#endregion Fin de Proveedor
				#region Comprobante, Autorizacion, Serie, Fecha Emision
				miXML.WriteElementString("tipoComp", int.Parse(stCodComprob).ToString());

				if (!bExterior)
				{
					stAutFactura = stAutFactura.Trim();
					stAutFactura = Digito(stAutFactura);
					if (stAutFactura.Length != 10)
					{
						MessageBox.Show("Error en Numero de Autorizacion de Factura", "Información");
//						new Compra(IdTipoFactura, IdCompra).Show();
						miXML.Close();
						Cursor = Cursors.Default;
						return;
					}
					miXML.WriteElementString("aut", stAutFactura.ToString());
				}
				else
				{
					miXML.WriteElementString("aut", "9999999999");
				}

				stNumero = Digito(stNumero);
				string stEstablecmiento = "";
				string stPuntoEmision = "";
				string stSecuencial = stNumero;
				if (!bExterior)
				{
					if (stNumero.Trim().Length > 9 && stNumero.Trim().Length < 15)
					{
						MessageBox.Show("Error en Numero o Serie de Factura", "Información");
//						new Compra(IdTipoFactura, IdCompra).Show();
						miXML.Close();
						Cursor = Cursors.Default;
						return;
					}

					if (stNumero.Trim().Length > 9)
					{
						stEstablecmiento = stNumero.Substring(0, 3);
						stPuntoEmision = stNumero.Substring(3,3);
						stSecuencial = stNumero.Substring(6);
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
					if (stEstablecmiento.Length != 3  || stPuntoEmision.Length != 3 
						|| stSecuencial.Length == 0 || stSecuencial.Length > 9)
					{
						MessageBox.Show("Error en Numero o Serie de Factura", "Información");
//						new Compra(IdTipoFactura, IdCompra).Show();
						miXML.Close();
						Cursor = Cursors.Default;
						return;
					}
					miXML.WriteElementString("estab", stEstablecmiento);
					miXML.WriteElementString("ptoEmi", stPuntoEmision);
				}
				else
				{
					miXML.WriteElementString("estab", "001");
					miXML.WriteElementString("ptoEmi", "001");
				}
				miXML.WriteElementString("sec", stSecuencial);

				miXML.WriteElementString("fechaEmiCom", dtFechaCompra.ToString("dd/MM/yyyy"));
				#endregion fin de serie
				DateTime dtFechaRet = miReader.GetDateTime(18);

				#region Impuesto a la Renta Retenido
				miXML.WriteStartElement("air");
				double dRentaRetTotal = 0;

				if (bRetNoAplica == false)
				{
					#region Variables de Detalle de Retencion
					double[] dAcRentaRet = new double[iCuentaRet];
					double[] dAcRentaBase = new double[iCuentaRet];
					double[] dAcRentaPorc = new double[iCuentaRet];
					string[] stAcCodRenta = new string[iCuentaRet];
					#endregion Variables de Detalle de Retencion
					for (int i=0; i<iCuentaRet; i++)
					{
						double dRentaRet = miReader.GetDouble(12);
						double dRentaBase = miReader.GetDouble(13);
						double dRentaPorc = 0;
						if (dRentaBase > 0) dRentaPorc = dRentaRet / dRentaBase * 100.0 + 0.001;
						string stCodRenta = miReader.GetString(14);

						if (dRentaBase < .0001) 
						{
							if (i < iCuentaRet - 1) miReader.Read();
							continue;
						}
						dRentaRetTotal += dRentaRet;

						stAcCodRenta[i] = stCodRenta;
						dAcRentaBase[i] = dRentaBase;
						dAcRentaPorc[i] = dRentaPorc;
						dAcRentaRet[i] = dRentaRet;
						if (i < iCuentaRet - 1) miReader.Read();
					}

					for (int i=0; i<iCuentaRet; i++)
					{
						for (int j=i + 1; j<iCuentaRet; j++)
						{
							if (stAcCodRenta[i] == stAcCodRenta[j])
							{
								dAcRentaBase[i] += dAcRentaBase[j];
								dAcRentaRet[i] += dAcRentaRet[j];
								stAcCodRenta[j] = "NoE";
							}
						}
					}
					for (int i=0; i<iCuentaRet; i++)
					{
						if (stAcCodRenta[i] == null) continue;
						if (stAcCodRenta[i] == "NoE") continue;
						if (stAcCodRenta[i].Trim() == "0")
						{
							MessageBox.Show("Error en Codigo de Retencion", "Información");
//							new Compra(IdTipoFactura, IdCompra).Show();
							miXML.Close();
							Cursor = Cursors.Default;
							return;
						}
						miXML.WriteStartElement("detalleAir");
						miXML.WriteElementString("codRetAir", stAcCodRenta[i].Trim());
						miXML.WriteElementString("porcentaje", Math.Round(dAcRentaPorc[i], 0).ToString("0.0", us));
						if (dSubtotalExcento > 0)
						{
							miXML.WriteElementString("base0", dAcRentaBase[i].ToString("0.00", us));
							miXML.WriteElementString("baseGrav", "0.00");
						}
						else
						{
							miXML.WriteElementString("base0", "0.00");
							miXML.WriteElementString("baseGrav", dAcRentaBase[i].ToString("0.00", us));
						}
						miXML.WriteElementString("baseNoGrav", "0.00");
						miXML.WriteElementString("valRetAir", dAcRentaRet[i].ToString("0.00", us));
						miXML.WriteEndElement();
						// Fin de detalleAir
					}
				}
				miXML.WriteEndElement(); // Fin de air
				#endregion Impuesto a la renta
				#region Codigos de Retencion
				if (iCuentaRet > 0 )
				{
					#region Lectura de Autorizacion Serie y Codigo
					string stAutRet = miReader.GetString(15);
					string stSerieRet = miReader.GetString(16);
					string stCodRet = miReader.GetString(17);
					stSerieRet = Digito(stSerieRet);
					stAutRet = Digito(stAutRet);
					stCodRet = Digito(stCodRet);
					#endregion Lectura de Autorizacion Serie y Codigo

					#region Verificacion de Autorizacion Serie Codigo Retencion
					if (dRentaRetTotal >= 0.01 && (stSerieRet.Length != 6 || stAutRet.Length != 10 || stCodRet.Length < 1))
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
					#endregion Verificacion de Autorizacion Serie Codigo Retencion
					#region Verificacion de Fecha de Emision de Retencion mayor o igual que factura
					DateTime dtFechaNula = DateTime.Today.AddYears(-100);
					if (dtFechaRet < dtFechaCompra && 
						!(dtFechaRet > dtFechaNula.AddDays(-1) && dtFechaRet < dtFechaNula.AddDays(1)))
					{
						MessageBox.Show("Fecha de Retencion debe ser mayor que " + dtFechaCompra.ToString("dd/MMM/yyyy"), "Información");
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
					#endregion Fecha de Emision de Retencion mayor o igual que factura

					#region Grabacion de Autorizacion, etc de Retencion
					if (stSerieRet.Trim().Length == 6 && stCodRet.Trim().Length > 0
						&& (dRentaRetTotal >= 0.01))
					{
						miXML.WriteElementString("autRet", stAutRet);
						miXML.WriteElementString("estabRet", stSerieRet.Substring(0,3));
						miXML.WriteElementString("ptoEmiRet", stSerieRet.Substring(3,3));
						miXML.WriteElementString("secRet", stCodRet);
						miXML.WriteElementString("fechaEmiRet", dtFechaRet.ToString("dd/MM/yyyy"));
					}
					#endregion Grabacion de Autorizacion, etc de Retencion
				}
				#endregion Codigos de Retencion

				#region Devoluciones de compras
				if (IdTipoFactura == 6)
				{
					try
					{
						if (stDevComprob[k].Trim().Length > 15 || stDevComprob[k].Trim().Length < 1)
						{
							MessageBox.Show("Error en Comprobante Asociado", "Información");
//							new Compra(6, IdCompra1[k]).Show();
							miXML.Close();
							Cursor = Cursors.Default;
							return;
						}

						if (stDevComprob[k].Trim().Length == 15)
						{
							stDevComprob[k] = stDevComprob[k].Substring(6, 9);
						}
						int iSerie = int.Parse(stDevSerie[k].Substring(0,3));
						int iPuntoE = int.Parse(stDevSerie[k].Substring(3,3));
						if (iSerie < 1 || iSerie > 999 || iPuntoE < 1 || iPuntoE > 999)
						{
							MessageBox.Show("Error en Serie de Comprobante Asociado", "Información");
//							new Compra(6, IdCompra).Show();
							miXML.Close();
							Cursor = Cursors.Default;
							return;
						}
						if (Int64.Parse(stDevAutoriza[k]) == 0 || stDevAutoriza[k].Length != 10)
						{
							MessageBox.Show("Error en Autorizacion de Comprobante Asociado", "Información");
//							new Compra(6, IdCompra).Show();
							miXML.Close();
							Cursor = Cursors.Default;
							return;
						}
						if (Int64.Parse(stDevComprob[k]) == 0 || stDevComprob[k].Length > 7)
						{
							MessageBox.Show("Error en Secuencial de Comprobante Asociado: " + stDevComprob[k], "Información");
//							new Compra(6, IdCompra).Show();
							miXML.Close();
							Cursor = Cursors.Default;
							return;
						}

					}
					catch (System.Exception ex)
					{
						MessageBox.Show(ex.Message, "No existe comprobante");
//						new Compra(6, IdCompra).Show();
						miXML.Close();
						Cursor = Cursors.Default;
						return;
					}
					k++;
				}
				#endregion fin de devoluciones de compras
				miXML.WriteEndElement(); // Fin de DetalleCompras
			}
			miReader.Close();

			#endregion Fin de Compras
			#region Retencion de Tarjetas en Compras
			stSelect = "Select Max(Cl.Ruc), Sum(IsNull(R.RentaRetenida, 0) "
				+ "+ IsNull(R.RentaRetenida1, 0)), "
				+ "Sum(IsNull(R.IvaTotal, 0)), Sum(IsNull(R.IvaRetenido, 0)), "
				+ "Max(IsNull(R.idRetIva, 0)), Sum(IsNull(R.IvaTotal1, 0)), "
				+ "Sum(IsNull(R.IvaRetenido1, 0)), Max(IsNull(R.idRetIva1, 0)) "
				+ "From Retencion r Inner Join Cliente cl On R.idCliente = Cl.idCliente "
				+ "where R.Tarjeta = 1 And R.NoAplica = 0 "
				+ "And (IsNull(R.IvaRetenido, 0) +IsNull(R.IvaRetenido1, 0) >= 0) "
				+ "And R.IngresoEgreso = 2 "
				+ "group by Cl.Ruc, R.idRetIva, R.idRetIva1";
			miReader = Funcion.rEscalarSQL(cdsCompra, stSelect, true);
			while(miReader.Read())
			{
				string stRucTarjeta = miReader.GetString(0);
				double dRentaTarjeta = miReader.GetDouble(1);
				double dIvaTarjeta = miReader.GetDouble(2);
				double dIvaRetTarjeta = miReader.GetDouble(3);
				int iIdRetTarjeta = miReader.GetInt32(4);
				double dIva1Tarjeta = miReader.GetDouble(5);
				double dIvaRet1Tarjeta = miReader.GetDouble(6);
				int iIdRet1Tarjeta = miReader.GetInt32(7);
			}
			miReader.Close();
			#endregion Retencion de Tarjetas en Compras

			string[] stCodTipoRuc1 = new string[iMax];
			string[] stRuc1 = new string[iMax];
			string [] stComprobante = new string[iMax];

			double[] dIvaBienes1 = new double[iMax];

			miXML.WriteEndElement(); // Fin de Compras
			miXML.WriteEndElement(); // Fin de IVA
			miXML.Flush();
			miXML.Close();
			Cursor = Cursors.Default;
			MessageBox.Show("Archivo por procedimiento Generado: " + stArchivo, "Información");
		}

		private void ATS()
		{
			Cursor = Cursors.WaitCursor;

			if (DialogResult.No == MessageBox.Show("¿Desea Generar el Anexo Transaccional Simplificado?", "Point Technology", MessageBoxButtons.YesNo, MessageBoxIcon.Question)) return;

			Funcion.TiempoInicia();

			#region Variables
			DateTime dtFecha = (DateTime) cmbFecha.Value;
			DateTime dtHasta = (DateTime) cmbHasta.Value;
			char Pad = '0';
			int idCompra = 0;
			int idOrigen = 0;
			int idTipoFactura = 0;
			DateTime dtFechaRegistro = DateTime.Today;
			DateTime dtFechaEmision = DateTime.Today;
			double dBaseNGrvIva = 0;
			double dBase0 = 0;
			double dBase = 0;
			double dICE = 0;
			double dIVA = 0;
			double dIVABienes = 0;
			double dIVAServicios = 0;
			double dIVABienesServicios = 0;

			double dVentaBaseNGIVA = 0;
			double dVentaBase0 = 0;
			double dVentaBase = 0;
			double dVentaIVA = 0;
			double dVentaRetIVA = 0;
			double dVentaRetRenta = 0;

			DateTime dtFechaRet = DateTime.Today;

			DataSet oDSCompras = new DataSet();
			cdsSeteo.Schema.Connections[0].Open();
			C1.Data.SchemaObjects.Connection miConeccion = cdsSeteo.Schema.Connections[0];
			miConeccion = cdsSeteo.Schema.Connections[0];
			#endregion Inicia Variables

			#region IVA
			string stArchivo = NombreGenerado("AT");
			XmlTextWriter miXML = new XmlTextWriter(stArchivo, Encoding.GetEncoding("ISO-8859-1"));
			miXML.Formatting = Formatting.Indented;
			miXML.WriteStartDocument(true);
			miXML.WriteStartElement("iva");
			
			#region Encabezado
			
			miXML.WriteElementString("TipoIDInformante", drSeteo.TipoId);
			miXML.WriteElementString("IdInformante", drSeteo.Ruc);
			miXML.WriteElementString("razonSocial", drSeteo.Empresa.ToUpper());
			miXML.WriteElementString("Anio", dtFecha.ToString("yyyy"));
			miXML.WriteElementString("Mes", dtFecha.ToString("MM"));
			miXML.WriteElementString("numEstabRuc", Funcion.iEscalarSQL(cdsSeteo, "Select COUNT(*) From Bodega Where Factura = 1 And Activo = 1 And Bodega Not In (50)").ToString().PadLeft(3, Pad)); //+ Iva - DescPorcIva  - DescPorcIva0 
			miXML.WriteElementString("totalVentas", Funcion.dEscalarSQL(cdsSeteo, string.Format("Select SUM((SubtotalIva - Descuento) + (SubtotalExcento - Descuento0)) From Compra Where idTipoFactura = 1 And Borrar = 0 And Bodega Not In (50) And Estado In (10, 11) And Fecha >= '{0}' And Fecha <= '{1}'", dtFecha.ToString("yyyyMMdd"), dtHasta.ToString("yyyyMMdd"))).ToString("0.00", us));
			miXML.WriteElementString("codigoOperativo", "IVA");
			#endregion encabezado

			#region Compras
			string sSQLCompras = string.Format("Exec ATSCompras '{0}', '{1}'", dtFecha.ToString("yyyyMMdd"), dtHasta.ToString("yyyyMMdd"));
			SqlDataAdapter oDACompras = new SqlDataAdapter(sSQLCompras, (SqlConnection) miConeccion.DbConnection);
			oDACompras.Fill(oDSCompras);

			if (oDSCompras.Tables[0].Rows.Count > 1)
			{
				miXML.WriteStartElement("compras");

				foreach(DataRow oDRCompras in oDSCompras.Tables[0].Rows)
				{
					#region Inicia Variables
					idCompra = (int)oDRCompras["idCompra"];
					idOrigen = (int)oDRCompras["idOrigen"];
					idTipoFactura = (int)oDRCompras["idTipoFactura"];
					dBaseNGrvIva = 0;
					dBase0 = 0;
					dBase = 0;
					dICE = 0;
					dIVA = 0;
					dIVABienes = 0;
					dIVAServicios = 0;
					dIVABienesServicios = 0;
					#endregion Inicia Variables

					miXML.WriteStartElement("detalleCompras");
					
					#region Generales
					miXML.WriteElementString("codSustento", oDRCompras["Sustento"].ToString().Trim());
					miXML.WriteElementString("tpIdProv", oDRCompras["TipoRuc"].ToString().Trim());
					miXML.WriteElementString("idProv", oDRCompras["Ruc"].ToString().Trim());
					miXML.WriteElementString("tipoComprobante", oDRCompras["Comprobante"].ToString().Trim());
					
					if (oDRCompras["TipoRuc"].ToString().Trim() == "03")
					{
						miXML.WriteElementString("tipoProv", oDRCompras["TipoProveedor"].ToString().Trim());
						miXML.WriteElementString("parteRel", "NO");
					}

					dtFechaRegistro = (DateTime)oDRCompras["FechaRegistro"];
					
					miXML.WriteElementString("fechaRegistro", dtFechaRegistro.ToString("dd/MM/yyyy"));

					if (oDRCompras["TipoRuc"].ToString().Trim() != "03")
					{
						miXML.WriteElementString("establecimiento", oDRCompras["SerieFactura"].ToString().Trim().Substring(0, 3));
						miXML.WriteElementString("puntoEmision", oDRCompras["SerieFactura"].ToString().Trim().Substring(3, 3));
					}
					else
					{
						miXML.WriteElementString("establecimiento", "999");
						miXML.WriteElementString("puntoEmision", "999");
					}

					miXML.WriteElementString("secuencial", oDRCompras["Numero"].ToString().Trim());

					dtFechaEmision = (DateTime)oDRCompras["Fecha"];
					
					miXML.WriteElementString("fechaEmision", dtFechaEmision.ToString("dd/MM/yyyy"));
					
					if (oDRCompras["TipoRuc"].ToString().Trim() != "03") miXML.WriteElementString("autorizacion", oDRCompras["AutFactura"].ToString().Trim());
					else miXML.WriteElementString("autorizacion", "9999999999");
					#endregion Generales

					#region Base IVA
					dBaseNGrvIva = (double)oDRCompras["BaseNoImp"];
					dBase0 = (double)oDRCompras["Base0"];
					dBase = (double)oDRCompras["Base"];
					dICE = (double)oDRCompras["Ice"];
					dIVA = (double)oDRCompras["Iva"];

					miXML.WriteElementString("baseNoGraIva", dBaseNGrvIva.ToString("0.00", us));
					miXML.WriteElementString("baseImponible", dBase0.ToString("0.00", us));
					miXML.WriteElementString("baseImpGrav", dBase.ToString("0.00", us));
					
					miXML.WriteElementString("montoIce", dICE.ToString("0.00", us));
					miXML.WriteElementString("montoIva", dIVA.ToString("0.00", us));
					#endregion Base IVA

					#region Retenciones de IVA
					dIVABienes = Funcion.dEscalarSQL(cdsSeteo, string.Format("Exec ATSRetencionIVA {0}, 30", idCompra));
					dIVAServicios = Funcion.dEscalarSQL(cdsSeteo, string.Format("Exec ATSRetencionIVA {0}, 70", idCompra));;
					dIVABienesServicios = Funcion.dEscalarSQL(cdsSeteo, string.Format("Exec ATSRetencionIVA {0}, 100", idCompra));;

					miXML.WriteElementString("valorRetBienes", dIVABienes.ToString("0.00", us));
					miXML.WriteElementString("valorRetServicios", dIVAServicios.ToString("0.00", us));
					miXML.WriteElementString("valRetServ100", dIVABienesServicios.ToString("0.00", us));
					#endregion Retenciones de IVA

					#region Pago Exterior o Local
					miXML.WriteStartElement("pagoExterior");
					miXML.WriteElementString("pagoLocExt", oDRCompras["TipoDePago"].ToString());

					if (oDRCompras["TipoDePago"].ToString() == "01")
					{
						miXML.WriteElementString("paisEfecPago", "NA");
						miXML.WriteElementString("aplicConvDobTrib", "NA");
						miXML.WriteElementString("pagExtSujRetNorLeg", "NA");
					}
					else
					{
						miXML.WriteElementString("paisEfecPago", oDRCompras["CodigoPais"].ToString());
						string ACDT = "NO";
						miXML.WriteElementString("aplicConvDobTrib", ACDT);
						if (ACDT == "NO") miXML.WriteElementString("pagExtSujRetNorLeg", "SI");
					}				
					miXML.WriteEndElement();
					#endregion Pago Exterior o Local

					#region Forma de pago
					if ((double)oDRCompras["Total"] > 1000)
					{
						miXML.WriteStartElement("formasDePago");
						miXML.WriteElementString("formaPago", oDRCompras["TipoFormaPago"].ToString());
						miXML.WriteEndElement();
					}
					#endregion Forma de pago					

					#region Retencion de Renta
					miXML.WriteStartElement("air");

					string sSQLAIR = string.Format("Exec ATSRetencionRenta {0}", idCompra);
					SqlDataReader drAIR = Funcion.rEscalarSQL(cdsSeteo, sSQLAIR, true);

					while(drAIR.Read())
					{
						miXML.WriteStartElement("detalleAir");
					
						miXML.WriteElementString("codRetAir", drAIR.GetString(0).Trim());
						miXML.WriteElementString("baseImpAir", drAIR.GetDouble(1).ToString("0.00", us));
						miXML.WriteElementString("porcentajeAir", drAIR.GetDouble(2).ToString("0.00", us));
						miXML.WriteElementString("valRetAir", drAIR.GetDouble(3).ToString("0.00", us));
					
						miXML.WriteEndElement();
					}
					drAIR.Close();

					miXML.WriteEndElement();
					#endregion Retencion de Renta

					#region Comprobante de Retencion
					if ((double)oDRCompras["Retenido"] > 0)
					{
						miXML.WriteElementString("estabRetencion1", oDRCompras["Serie"].ToString().Trim().Substring(0,3));
						miXML.WriteElementString("ptoEmiRetencion1", oDRCompras["Serie"].ToString().Trim().Substring(3,3));
						miXML.WriteElementString("secRetencion1", oDRCompras["NumRet"].ToString().Trim());
						miXML.WriteElementString("autRetencion1", oDRCompras["Autorizacion"].ToString().Trim());
						dtFechaRet = (DateTime)oDRCompras["FecRet"];
						miXML.WriteElementString("fechaEmiRet1", dtFechaRet.ToString("dd/MM/yyyy"));
					}
					#endregion Comprobante de Retencion

					#region Documento modificado
					if (idTipoFactura != 4)
					{
						string sSQLMod = string.Format("Select cc.Codigo, c.SerieFactura, Numero, AutFactura From Compra c Inner Join CompraComprobante cc On cc.idComprobante = c.idComprobante Where idTipoFactura = 4 And idCompra = {0}", idOrigen);
						SqlDataReader drModificado = Funcion.rEscalarSQL(cdsSeteo, sSQLMod, true);
						drModificado.Read();
						if (drModificado.HasRows)
						{
							miXML.WriteElementString("docModificado", drModificado.GetString(0).ToString().Trim());// int.Parse(stDevCompTipo[k]).ToString());
							miXML.WriteElementString("estabModificado", drModificado.GetString(1).ToString().Trim().Substring(0,3));
							miXML.WriteElementString("ptoEmiModificado", drModificado.GetString(1).ToString().Trim().Substring(3,3));
							miXML.WriteElementString("secModificado", drModificado.GetString(2).ToString().Trim());
							miXML.WriteElementString("autModificado", drModificado.GetString(3).ToString().Trim());
						}
						drModificado.Close();
					}
					#endregion Documento modificado

					miXML.WriteEndElement();
				}			
				miXML.WriteEndElement();
			}
			#endregion Compras

			#region Ventas
			miXML.WriteStartElement("ventas");

			string sSQLVentas = string.Format("Exec ATSVentas '{0}', '{1}'", dtFecha.ToString("yyyyMMdd"), dtHasta.ToString("yyyyMMdd"));
			SqlDataReader drVentas = Funcion.rEscalarSQL(cdsSeteo, sSQLVentas, true);

			while (drVentas.Read())
			{
				dVentaBaseNGIVA = 0;
				dVentaBase0 = 0;
				dVentaBase = 0;
				dIVA = 0;
				dVentaRetIVA = 0;
				dVentaRetRenta = 0;

				miXML.WriteStartElement("detalleVentas");

				miXML.WriteElementString("tpIdCliente", drVentas.GetString(0).ToString().Trim());
				miXML.WriteElementString("idCliente", drVentas.GetString(1).ToString().Trim());
				miXML.WriteElementString("tipoComprobante", drVentas.GetString(2).ToString().Trim());
				miXML.WriteElementString("numeroComprobantes", drVentas.GetValue(3).ToString().Trim());

				dVentaBaseNGIVA = drVentas.GetDouble(4);
				dVentaBase0 = drVentas.GetDouble(5);
				dVentaBase = drVentas.GetDouble(6);
				dVentaIVA = drVentas.GetDouble(7);
				dVentaRetIVA = drVentas.GetDouble(8);
				dVentaRetRenta = drVentas.GetDouble(9);

				miXML.WriteElementString("baseNoGraIva", dVentaBaseNGIVA.ToString("0.00", us));
				miXML.WriteElementString("baseImponible", dVentaBase0.ToString("0.00", us));
				miXML.WriteElementString("baseImpGrav", dVentaBase.ToString("0.00", us));
				miXML.WriteElementString("montoIva", dVentaIVA.ToString("0.00", us));
				miXML.WriteElementString("valorRetIva", dVentaRetIVA.ToString("0.00", us));
				miXML.WriteElementString("valorRetRenta", dVentaRetRenta.ToString("0.00", us));

				miXML.WriteEndElement();
			}
			drVentas.Close();

			miXML.WriteEndElement();
			#endregion Ventas

			#region Ventas Establecimiento
			miXML.WriteStartElement("ventasEstablecimiento");
			
			int iContBodegas = (int)Funcion.iEscalarSQL(cdsSeteo, "Select Count(*) From Bodega");
			
			bool bFactura = false;

			for (int l = 1; l <= iContBodegas; l++)
			{	
				if (l != 50) bFactura = Funcion.bEscalarSQL(cdsSeteo, string.Format("Select IsNull(Factura, 0) From Bodega Where Bodega = {0}", l), true);
				else bFactura = false;
				if (bFactura)
				{
					miXML.WriteStartElement("ventaEst");

					miXML.WriteElementString("codEstab", Funcion.sEscalarSQL(cdsSeteo, string.Format("Select SUBSTRING(Serie, 1, 3) From Bodega Where Bodega = {0} And Activo = 1 And Factura = 1", l)));
					miXML.WriteElementString("ventasEstab", Funcion.dEscalarSQL(cdsSeteo, string.Format("Select ISNULL(SUM((SubtotalIva - Descuento) + (SubtotalExcento - Descuento0)), 0) From Compra Where idTipoFactura = 1 And Borrar = 0 And Estado In (10, 11) And Fecha >= '{1}' And Fecha <= '{2}' And Bodega = {0}", l, dtFecha.ToString("yyyyMMdd"), dtHasta.ToString("yyyyMMdd"))).ToString("0.00", us));
					
					miXML.WriteEndElement();
				}
			}
	
			miXML.WriteEndElement();
			#endregion Ventas Establecimiento

			#region Comprobantes Anulados
			#region Facturas - Notas de crédito
			string sSQLFacAnuladas = string.Format("Exec ATSComprobantesAnulados '{0}', '{1}', 0", dtFecha.ToString("yyyyMMdd"), dtHasta.ToString("yyyyMMdd"));
			SqlDataReader drFacAnuladas = Funcion.rEscalarSQL(cdsSeteo, sSQLFacAnuladas, true);

			miXML.WriteStartElement("anulados");

			while (drFacAnuladas.Read())
			{
				miXML.WriteStartElement("detalleAnulados");
				
				miXML.WriteElementString("tipoComprobante", drFacAnuladas.GetString(0).ToString().Trim());
				miXML.WriteElementString("establecimiento", drFacAnuladas.GetString(1).ToString().Trim().Substring(0, 3));
				miXML.WriteElementString("puntoEmision", drFacAnuladas.GetString(1).ToString().Trim().Substring(3, 3));
				miXML.WriteElementString("secuencialInicio", drFacAnuladas.GetString(2).ToString().Trim());
				miXML.WriteElementString("secuencialFin", drFacAnuladas.GetString(2).ToString().Trim());
				miXML.WriteElementString("autorizacion", drFacAnuladas.GetString(3).ToString().Trim());

				miXML.WriteEndElement();
			}
			drFacAnuladas.Close();
			#endregion Facturas - Notas de crédito
			
			#region Retenciones
			string sSQLRetAnuladas = string.Format("Exec ATSComprobantesAnulados '{0}', '{1}', 1", dtFecha.ToString("yyyyMMdd"), dtHasta.ToString("yyyyMMdd"));
			SqlDataReader drRetAnuladas = Funcion.rEscalarSQL(cdsSeteo, sSQLRetAnuladas, true);

			while (drRetAnuladas.Read())
			{
				miXML.WriteStartElement("detalleAnulados");
				
				miXML.WriteElementString("tipoComprobante", drRetAnuladas.GetString(0).ToString().Trim());
				miXML.WriteElementString("establecimiento", drRetAnuladas.GetString(1).ToString().Trim().Substring(0, 3));
				miXML.WriteElementString("puntoEmision", drRetAnuladas.GetString(1).ToString().Trim().Substring(3, 3));
				miXML.WriteElementString("secuencialInicio", drRetAnuladas.GetString(2).ToString().Trim());
				miXML.WriteElementString("secuencialFin", drRetAnuladas.GetString(2).ToString().Trim());
				miXML.WriteElementString("autorizacion", drRetAnuladas.GetString(3).ToString().Trim());

				miXML.WriteEndElement();
			}
			drRetAnuladas.Close();
			#endregion Retenciones
			#endregion Comprobantes Anulados

			miXML.WriteEndElement();

			miXML.WriteEndElement(); 
			miXML.WriteEndDocument();

			miXML.Flush();
			miXML.Close();
			#endregion IVA
			
			MessageBox.Show("Anexo Transaccional Simplificado " + stArchivo + " Generado " + Funcion.Tiempo("En"), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Information);

			Cursor = Cursors.Default;
		}

		private void Anexo2008()
		{
			Cursor = Cursors.WaitCursor;
			DateTime dtInicio = DateTime.Now;
			string stTiempos = "\nTiempos de Generación:";
			DateTime dtFecha = (DateTime) cmbFecha.Value;
			DateTime dtHasta = (DateTime) cmbHasta.Value;
			miLoteiCuenta = 0;
			Funcion.EjecutaSQL(cdsCompra, "Exec VerificaLote", true);

			#region Encabezado
			string stArchivo = NombreGenerado("AT");
			XmlTextWriter miXML = new XmlTextWriter(stArchivo, Encoding.GetEncoding("ISO-8859-1"));
			miXML.Formatting = Formatting.Indented;
			miXML.WriteStartDocument(true);
			miXML.WriteStartElement("iva");
			string stMensaje = "";
			string stVerificaRuc = Funcion.VerificaCedula(drSeteo.Ruc, false);

			miXML.WriteElementString("TipoIDInformante", drSeteo.TipoId);
			if (stVerificaRuc.Length == 0) miXML.WriteElementString("IdInformante", drSeteo.Ruc);
			else stMensaje += "Error en RUC de su Empresa\n";
			if (drSeteo.Empresa.Trim().Length > 2) miXML.WriteElementString("razonSocial", drSeteo.Empresa.ToUpper());
			else stMensaje += "Longitud de Nombre de su Empresa Erronea\n";

			miXML.WriteElementString("Anio", dtFecha.ToString("yyyy"));
			miXML.WriteElementString("Mes", dtFecha.ToString("MM"));
			miXML.WriteElementString("numEstabRuc", Funcion.iEscalarSQL(cdsSeteo, "Select COUNT(*) From Bodega Where Factura = 1 And Activo = 1 And Bodega Not In (50)").ToString()); //+ Iva - DescPorcIva  - DescPorcIva0 
			miXML.WriteElementString("totalVentas", Funcion.dEscalarSQL(cdsSeteo, string.Format("Select SUM((SubtotalIva - Descuento) + (SubtotalExcento - Descuento0)) From Compra Where idTipoFactura = 1 And Borrar = 0 And Estado In (10, 11) And Fecha >= '{0}' And Fecha <= '{1}'", dtFecha.ToString("yyyyMMdd"), dtHasta.ToString("yyyyMMdd"))).ToString("0.00", us));
			miXML.WriteElementString("codigoOperativo", "IVA");

			if (stMensaje.Trim().Length > 0)
			{
				MessageBox.Show(stMensaje, "Información");
				using(Seteo miSeteo = new Seteo())
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

			string stExec = "Exec DevolucionImportaciones ";
			Funcion.EjecutaSQL(cdsCompra, stExec, true);
			string stFechaRet = "Retencion.Fecha";
			if (drSeteo.At_Fecha_Cont == true) stFechaRet = "Retencion.FechaContable";

			#region Compras
			#region Devolucion de compras
			miXML.WriteStartElement("compras");
			string stSelect = string.Format("Exec CompraIvaServicio '{0}', '{1}'", dtFecha.ToString("yyyyMMdd"), dtHasta.ToString("yyyyMMdd"));

			if (chkAjusteIva.Checked) Funcion.EjecutaSQL(cdsCompra, stSelect, true);
//			MessageBox.Show("Filas Compras", "Informacion");

			// Creacion de Devolucion de compras
			stSelect = "SELECT IsNull(Compra.idCliente, 0), "
				+ "IsNull(Compra.Comprobante, ''), IsNull(Compra.idCompra, 0), "
				+ "IsNull(Compra.Comprob1, '') "
				+ "FROM Compra LEFT OUTER JOIN CompraComprobante "
				+ "ON Compra.idComprobante = CompraComprobante.idComprobante "
				+ "Where Compra.idTipoFactura = 6 And Compra.Borrar = 0 And Compra.Fecha >= '" 
				+ dtFecha.ToString("yyyyMMdd") + "' And Compra.Fecha < '" 
				+ dtHasta.AddDays(1).ToString("yyyyMMdd") + "'"
				+ " And Not CompraComprobante.Codigo In ('17', '99') "
				+ " Order By Compra.IdCompra ";
			SqlDataReader miReader = Funcion.rEscalarSQL(cdsCompra, stSelect, true);
			int k = 0;
			int iMax = 20000;
			int [] IdDevCliente = new int [iMax];
			int [] IdDevCompra = new int [iMax];
			int [] IdCompra1 = new int [iMax];
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
				string stComp1 = miReader.GetString(3);
				if (stComp1.Trim().Length > 0) stDevComprob[k] = stComp1.Trim();
				k++;
			}
			miReader.Close();
			for (int j=0; j< k; j++)
			{
				stSelect = "Select Count(*) From Compra "
					+ "WHERE Compra.Numero = '" + stDevComprob[j].Trim()
					+ "' AND Compra.idCliente = " + IdDevCliente[j].ToString();
				int iCuenta = Funcion.iEscalarSQL(cdsCompra, stSelect, true);
				if (iCuenta == 0)
				{
					MessageBox.Show("Numero de Comprobante no existe",	"Información");
//					new Compra(6, IdDevCompra[j]).Show();
					miXML.Close();
					Cursor = Cursors.Default;
					return;
				}
				stSelect = "SELECT IsNull(Compra.SerieFactura, ''), "
					+ "IsNull(Compra.Fecha, DateAdd(year, -100, GetDate())), "
					+ "IsNull(Compra.AutFactura, ''), "
					+ "IsNull(CompraComprobante.Codigo, ''), IsNull(Compra.idCompra, 0) FROM Compra "
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
				IdCompra1[j] = miReader.GetInt32(4);
				miReader.Close();

				stDevComprob[j] = Digito(stDevComprob[j]);
				#endregion Devolucion de compras

			#region Verificaciones de devoluciones
				if (stDevCompTipo[j].Trim().Length < 2)
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
//				if (stDevAutoriza[j].Trim().Length < 6)
//				{
//					MessageBox.Show("Autorizacion de Comprobante Asociado no existe",	"Información");
//					new Compra(6, IdDevCompra[j]).Show();
//					miXML.Close();
//					Cursor = Cursors.Default;
//					return;
//				}
				#endregion Verificaciones
			}
			k = 0; // Incializado para que lo recorra nuevamente en devoluciones

			#region Select de Compras
			stSelect = "SELECT IsNull(CompraCredTribut.Codigo, ''), "
				+ "Compra.DerechoDevolucion, IsNull(ClienteTipoRuc.Codigo, '') ,"
				+ "IsNull(substring(Cliente.Ruc, 1, 13), ''), IsNull(CompraComprobante.Codigo, ''), "
				+ "IsNull(Compra.Numero, ''), IsNull(Cliente.idCliente,0), "
				+ "IsNull(Compra.SerieFactura, ''), IsNull(Compra.AutFactura, ''), "
				+ "IsNull(Compra.FechaCaducidad, DateAdd(year, -100,GetDate())), "
				+ "IsNull(Compra.SubtotalExcento - Compra.Descuento0, 0), "
				+ "IsNull(Compra.SubtotalIva - Compra.Descuento, 0), "
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
				+ "IsNull( " + stFechaRet + ", DateAdd(year, -100,GetDate())), "
				+ "(Select Count(*) From RetDetalle rd Inner Join Retencion r On rd.idRetencion = r.idRetencion Where r.idCompra = Compra.idCompra), IsNull(Compra.idTipoFactura, 0), " //IsNull(Compra.CuentaRetDet, 0)
				+ "IsNull(Retencion.IvaTotal, 0), IsNull(Retencion.IvaTotal1, 0), "
				+ "IsNull(Retencion.Tarjeta, 0), IsNull(Retencion.NoAplica, 0), "
				+ "IsNull(Compra.Retenido, 0), IsNull(Compra.IdPagoLot, 0), "
				+ "IsNull(Compra.Ice, 0), IsNull(PorcIce.porcentaje, 0), "
				+ "IsNull(Compra.Servicio, 0), IsNull(TiposDePagos.Codigo, ''), IsNull(ClientePais.Codigo, ''), IsNull(TipoFormaDePago.Codigo, ''), IsNull(TipoDeProveedor.Codigo, '') "
				+ "FROM Compra LEFT OUTER JOIN Retencion ON Retencion.idCompra = Compra.idCompra "
				+ "LEFT OUTER JOIN Cliente ON Compra.idCliente = Cliente.idCliente "
				+ "LEFT OUTER JOIN CompraCredTribut ON Compra.idCredTributario = CompraCredTribut.idCredTributario "
				+ "LEFT OUTER JOIN CompraComprobante ON Compra.idComprobante = CompraComprobante.idComprobante "
				+ "LEFT OUTER JOIN ClienteTipoRuc ON Cliente.idTipoRuc = ClienteTipoRuc.idTipoRuc "
				+ "LEFT OUTER JOIN RetencionIva ON Retencion.idRetIva = RetencionIva.idRetIva "
				+ "LEFT OUTER JOIN RetencionIva As RetIva_1  ON Retencion.idRetIva1 = RetIva_1.idRetIva "
				+ "LEFT OUTER JOIN RetDetalle ON RetDetalle.idRetencion = Retencion.idRetencion "
				+ "LEFT OUTER JOIN ArticuloSri ON RetDetalle.idRetRenta = ArticuloSri.idArticuloSri "
				+ "LEFT JOIN porcice ON compra.idporcice = porcice.idPorcIce "
				+ "LEFT JOIN TiposDePagos On Compra.idTipoDePago = TiposDePagos.idTipoDePago "
				+ "LEFT JOIN ClientePais On Cliente.idPais = ClientePais.idClientePais "
				+ "LEFT JOIN TipoFormaDePago On TipoFormaDePago.idTipoFormaPago = Compra.idTipoFormaPago "
				+ "LEFT JOIN TipoDeProveedor On TipoDeProveedor.idTipoProveedor = Cliente.idTipoProveedor "
				+ "Where Compra.idTipoFactura In (4, 6) And Compra.Borrar = 0 And Compra.Fecha >= '" 
				+ dtFecha.ToString("yyyyMMdd") + "' And Compra.Fecha < '" + dtHasta.AddDays(1).ToString("yyyyMMdd") + "'"
				//				Quitar estos comentarios para retenciones del mes con compra del mes pasado
				//				+ "Where Compra.idTipoFactura In (4, 6) And Compra.Borrar = 0 And ((Compra.Fecha >= '" 
				//				+ dtFecha.ToString("yyyyMMdd") + "' And Compra.Fecha < '" 
				//				+ dtHasta.AddDays(1).ToString("yyyyMMdd") + "') || ("
				//				+ " And Retencion.Fecha < '" + dtHasta.AddDays(1).ToString("yyyyMMdd") + "'"
				//				+ " And Retencion.Fecha >= '" + dtFecha.ToString("yyyyMMdd") + "'"
				//				+ " And Compra.Fecha < '" + dtFecha.ToString("yyyyMMdd") + "'))"
				+ " And Not CompraComprobante.Codigo In ('17', '99') "
				+ " Order By Compra.IdCompra";
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
			double dFacIva = 0;
			double dFacServ = 0;
			bool bTarjeta = false;
			bool bNoAplica = false;
			double dRetenidoLote = 0;
			int idPagoLot = 0;
			double dCompraIce = 0;
			double dPorcIce = 0;
			double dServicioC = 0;
			string CodPago = "";
			string CodPais = "";
			string CodTipoFormaPago = "";
			string stTipoProveedor = "";
			#endregion fin de definicion de variables
			#region Lectura de Filas de Compras
			while (miReader.Read())
			{
				IdTipoFactura = miReader.GetInt32(32);
				IdCompra = miReader.GetInt32(13);
				miXML.WriteStartElement("detalleCompras");
				stCodCredTrib = miReader.GetString(0);
				bDerechoDev = miReader.GetBoolean(1);
				stCodTipoRuc = miReader.GetString(2);
				stRuc = miReader.GetString(3).Trim();
				string stPasaporte = stRuc;
				IdCliente = miReader.GetInt32(6);
				stRuc = Digito(stRuc);
				if (IdCliente < 1)
				{
					MessageBox.Show("Proveedor no definido", "Información");
//					new Compra(IdTipoFactura, IdCompra).Show();
					miXML.Close();
					Cursor = Cursors.Default;
					return;
				}

				string stVerifica = Funcion.VerificaCedula(stRuc, false);
				if (stRuc.Length != 10 && stCodTipoRuc == "02") stVerifica = "Err";
				//if (stCodTipoRuc == "01") stRuc += "000";
				if (stRuc.Length != 13 && (stCodTipoRuc == "01" || stCodTipoRuc == "04"))	stVerifica = "Err";
				if (stCodTipoRuc == "03" || stCodTipoRuc == "07" || stCodTipoRuc == "08") stVerifica = "";
				if (stVerifica.Length > 0)
				{
					MessageBox.Show("Error en Cedula o RUC", "Información");
					new Cliente(false, IdCliente).Show();
					miXML.Close();
					Cursor = Cursors.Default;
					return;
				}
				if (stCodTipoRuc.Length != 2)
				{
					MessageBox.Show("Error en Tipo Cedula o RUC", "Información");
					new Cliente(false, IdCliente).Show();
					miXML.Close();
					Cursor = Cursors.Default;
					return;
				}
				if (stRuc.Length == 13 && stCodTipoRuc == "01") stRuc = stRuc.Substring(0,10) + "001";
				stCodComprob =  miReader.GetString(4);
				stNumero =  miReader.GetString(5);
				stCompraSerieFact = miReader.GetString(7);
				if (stCodComprob.Trim() == "15") stAutFactura = "9999999999";
				else stAutFactura = miReader.GetString(8);
				dtCaducidad =  miReader.GetDateTime(9);
				dSubtotalExcento = miReader.GetDouble(10);
				dSubtotalIva = miReader.GetDouble(11);
				dCompraIva = miReader.GetDouble(12);
				dRetIvaPorc = miReader.GetDouble(14);
				stCodRetIva = miReader.GetString(15);
				if (Math.Round(dRetIvaPorc, 0) == 0) stCodRetIva = "0";
				if (Math.Round(dRetIvaPorc, 0) == 30) stCodRetIva = "1";
				if (Math.Round(dRetIvaPorc, 0) == 70) stCodRetIva = "2";
				if (Math.Round(dRetIvaPorc, 0) == 100) stCodRetIva = "3";

				dRetIvaPorc1 = miReader.GetDouble(16);
				stCodRetIva1 = miReader.GetString(17);
				if (Math.Round(dRetIvaPorc1, 0) == 0) stCodRetIva1 = "0";
				if (Math.Round(dRetIvaPorc1, 0) == 30) stCodRetIva1 = "1";
				if (Math.Round(dRetIvaPorc1, 0) == 70) stCodRetIva1 = "2";
				if (Math.Round(dRetIvaPorc1, 0) == 100) stCodRetIva1 = "3";

				dtFechaCompra = miReader.GetDateTime(18);
				dRetValorIva = miReader.GetDouble(19);
				dRetValorIva1 = miReader.GetDouble(20);
				IdRetencion = miReader.GetInt32(21);
				dIvaServicio = miReader.GetDouble(22);
				iCuentaRet = miReader.GetInt32(31);
				dFacIva = miReader.GetDouble(33);
				dFacServ = miReader.GetDouble(34);
				bTarjeta = miReader.GetBoolean(35);
				bNoAplica = miReader.GetBoolean(36);
				dRetenidoLote = miReader.GetDouble(37);
				idPagoLot = miReader.GetInt32(38);
				dCompraIce = miReader.GetDouble(39);
				dPorcIce = miReader.GetDouble(40);
				dServicioC = miReader.GetDouble(41);
				CodPago = miReader.GetString(42);
				CodPais = miReader.GetString(43);
				CodTipoFormaPago = miReader.GetString(44);
				stTipoProveedor = miReader.GetString(45);
			#endregion Lectura de Filas de Compras

			#region Sustento, Devolucion de Iva
			miXML.WriteElementString("codSustento", stCodCredTrib); // Identificacion del Sustento Tributario
			#endregion fin de sustento
			#region Proveedor
				if (stCodTipoRuc == "07" || stCodTipoRuc == "08") stCodTipoRuc = "03";
				if (int.Parse(stCodTipoRuc) > 3)
				{
					int iCodTipoRuc = int.Parse(stCodTipoRuc) - 3;
					stCodTipoRuc = "0" + iCodTipoRuc.ToString().Trim();
				}
				miXML.WriteElementString("tpIdProv", stCodTipoRuc);
				if (stCodTipoRuc == "03") // Pasaporte
					miXML.WriteElementString("idProv", stPasaporte);
				else
					miXML.WriteElementString("idProv", stRuc);

				if (stCodComprob.Trim().Length < 2)
				{
					MessageBox.Show("Codigo de Comprobante no existe", "Información");
//					new Compra(IdTipoFactura, IdCompra).Show();
					miXML.Close();
					Cursor = Cursors.Default;
					return;
				}
				if ((stCodTipoRuc == "02") && !(stCodComprob.Trim() == "03" || stCodComprob.Trim() == "04"))
				{
					MessageBox.Show("Con cédula sólo se permite liquidación de compras", "Informacion");
//					new Compra(IdTipoFactura, IdCompra).Show();
					miXML.Close();
					Cursor = Cursors.Default;
					return;
				}
//				if ((stCodTipoRuc == "03") && !(stCodComprob.Trim() == "03" || stCodComprob.Trim() == "04" || stCodComprob.Trim() == "15"))
//				{
//					MessageBox.Show("Con Pasaporte sólo se permite liquidación de compras o comprobantes emitidos en el exterior", "Informacion");
//					new Compra(IdTipoFactura, IdCompra).Show();
//					miXML.Close();
//					Cursor = Cursors.Default;
//					return;
//				}

				if ((stCodTipoRuc == "01") && stCodComprob.Trim() == "03")
				{
					MessageBox.Show("Con RUC No se permite liquidacion de compras", "Información");
//					new Compra(IdTipoFactura, IdCompra).Show();
					miXML.Close();
					Cursor = Cursors.Default;
					return;
				}				
				#endregion Fin de Proveedor
			#region Comprobante, Fecha Asiento, Serie, Fecha Emision
				// Codigo Tipo de Comprobante
				if (stCodComprob.Trim() == "02")
					if (stCodCredTrib.Trim() != "02")
					{
						MessageBox.Show("Nota de Venta no sustenta Credito Tributario", "Información");
//						new Compra(IdTipoFactura, IdCompra).Show();
						miXML.Close();
						Cursor = Cursors.Default;
						return;
					}

				if (stCodComprob.Trim() == "15")
					if (!(stCodCredTrib.Trim() == "02" || stCodCredTrib.Trim() == "04" || stCodCredTrib.Trim() == "07"))
					{
						MessageBox.Show("Comprobantes en el exterior no sustentan Credito Tributario", "Información");
//							new Compra(IdTipoFactura, IdCompra).Show();
						miXML.Close();
						Cursor = Cursors.Default;
						return;
					}

				miXML.WriteElementString("tipoComprobante", stCodComprob.Trim()); // int.Parse(stCodComprob).ToString());

				if (stCodTipoRuc == "03")
				{
					miXML.WriteElementString("tipoProv", stTipoProveedor);
					miXML.WriteElementString("parteRel", "NO");
				}

				
				miXML.WriteElementString("fechaRegistro", dtFechaCompra.ToString("dd/MM/yyyy"));
			
				stNumero = Digito(stNumero);
				string stEstablecmiento = "";
				string stPuntoEmision = "";
				string stSecuencial = "";
				
				if (stNumero.Trim().Length > 9 && stNumero.Trim().Length < 15)
				{
					MessageBox.Show("Error en Numero o Serie de Factura", "Información");
//					new Compra(IdTipoFactura, IdCompra).Show();
					miXML.Close();
					Cursor = Cursors.Default;
					return;
				}

				if (stCodComprob.Trim() == "15")
				{
					stEstablecmiento = "999";
					stPuntoEmision = "999";
				}
				else	if (stNumero.Trim().Length > 9)
				{
					stEstablecmiento = stNumero.Substring(0, 3);
					stPuntoEmision = stNumero.Substring(3,3);
					stSecuencial = stNumero.Substring(6);
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

				if (stEstablecmiento.Length != 3  || stPuntoEmision.Length != 3 || stSecuencial.Length == 0 || stSecuencial.Length > 9)
				{
					MessageBox.Show("Error en Numero o Serie de Factura", "Información");
//					new Compra(IdTipoFactura, IdCompra).Show();
					miXML.Close();
					Cursor = Cursors.Default;
					return;
				}
				
					miXML.WriteElementString("establecimiento", stEstablecmiento);
					miXML.WriteElementString("puntoEmision", stPuntoEmision);
				
					miXML.WriteElementString("secuencial", stSecuencial);

					// Fecha de Emision
					miXML.WriteElementString("fechaEmision", dtFechaCompra.ToString("dd/MM/yyyy"));
					#endregion fin de serie
			#region Autorizacion y caducidad
					stAutFactura = stAutFactura.Trim();
					stAutFactura = Digito(stAutFactura);
					if (stAutFactura.Length != 10)
					{
						MessageBox.Show("Error en Numero de Autorizacion de Factura", "Información");
//						new Compra(IdTipoFactura, IdCompra).Show();
						miXML.Close();
						Cursor = Cursors.Default;
						return;
					}
					miXML.WriteElementString("autorizacion", stAutFactura.ToString());
					#endregion Autorizacion y caducidad
				
			#region IVA
				miXML.WriteElementString("baseNoGraIva", dServicioC.ToString("0.00", us));
				miXML.WriteElementString("baseImponible", dSubtotalExcento.ToString("0.00", us));
				miXML.WriteElementString("baseImpGrav", dSubtotalIva.ToString("0.00", us));
				#endregion Fin de IVA
			#region ICE
				double dBaseIce = 0;
				if (dCompraIce > 0)
				{
					if (dPorcIce == 0)
					{
						MessageBox.Show("Codigo de Retencion del ICE No Existe. Ingreselo en la Ficha AT",
							"Información");
//						new Compra(IdTipoFactura, IdCompra).Show();
						miXML.Close();
						Cursor = Cursors.Default;
						return;
					}

					dBaseIce = dCompraIce / dPorcIce * 100.0;
				}
				if (dPorcIce > 0)
				{
					if (Math.Round(dPorcIce, 2) == 0) dPorcIce = 1;
					else if (Math.Round(dPorcIce, 2) == 18.54) dPorcIce = 2;
					else if (Math.Round(dPorcIce, 2) == 30.9) dPorcIce = 3;
					else if (Math.Round(dPorcIce, 2) == 10.3) dPorcIce = 4;
					else if (Math.Round(dPorcIce, 2) == 5.15) dPorcIce = 6;
					else if (Math.Round(dPorcIce, 2) == 10.3) dPorcIce = 7;
					else if (Math.Round(dPorcIce, 2) == 15) dPorcIce = 8;
					else if (Math.Round(dPorcIce, 2) == 98) dPorcIce = 9;
					else if (Math.Round(dPorcIce, 2) == 32) dPorcIce = 10;
				}
				miXML.WriteElementString("montoIce", dCompraIce.ToString("0.00", us));
				miXML.WriteElementString("montoIva", dCompraIva.ToString("0.00", us));
				#endregion
			#region Retencion de Iva 
				double dIvaBienes = 0;
				double dvalorRetServicios = 0;
				double dvalorRetServ100 = 0;
				if (stCodRetIva == "1") dIvaBienes += dRetValorIva;
				if (stCodRetIva == "2") dvalorRetServicios += dRetValorIva;
				if (stCodRetIva == "3") dvalorRetServ100 += dRetValorIva;
				if (stCodRetIva1 == "1") dIvaBienes += dRetValorIva1;
				if (stCodRetIva1 == "2") dvalorRetServicios += dRetValorIva1;
				if (stCodRetIva1 == "3") dvalorRetServ100 += dRetValorIva1;
				miXML.WriteElementString("valorRetBienes", dIvaBienes.ToString("0.00", us));
				miXML.WriteElementString("valorRetServicios", dvalorRetServicios.ToString("0.00", us));
				miXML.WriteElementString("valRetServ100", dvalorRetServ100.ToString("0.00", us));
				#endregion Iva 
			#region Pagos Exterior
				miXML.WriteStartElement("pagoExterior");
				miXML.WriteElementString("pagoLocExt", CodPago);

				if (CodPago == "01")
				{
					miXML.WriteElementString("paisEfecPago", "NA");
					miXML.WriteElementString("aplicConvDobTrib", "NA");
					miXML.WriteElementString("pagExtSujRetNorLeg", "NA");
				}
				else
				{
					miXML.WriteElementString("paisEfecPago", CodPais);
					string ACDT = "NO";
//					string PESRNL = "NO";
					miXML.WriteElementString("aplicConvDobTrib", ACDT);
					if (ACDT == "NO") 
					{						
						miXML.WriteElementString("pagExtSujRetNorLeg", "SI");
					}
				}				
				miXML.WriteEndElement();
				#endregion Pagos Exterior
			#region Formas de pago
				if (CodTipoFormaPago.Length > 0)
				{
					miXML.WriteStartElement("formasDePago");
					miXML.WriteElementString("formaPago", CodTipoFormaPago);
					miXML.WriteEndElement();
				}
				#endregion Formas de pago
			DateTime dtFechaRet = miReader.GetDateTime(30);
			#region Impuesto a la Renta Retenido
				miXML.WriteStartElement("air");
				double dRentaRetTotal = 0;
				if (idPagoLot > 0) iCuentaRet = 1;
				bool bRetPeriodo = true;

				if (!(bTarjeta || bNoAplica || !bRetPeriodo))
				{
					double[] dAcRentaRet = new double[iCuentaRet];
					double[] dAcRentaBase = new double[iCuentaRet];
					double[] dAcRentaPorc = new double[iCuentaRet];
					string[] stAcCodRenta = new string[iCuentaRet];
					for (int i = 0; i < iCuentaRet; i++)
					{
						double dRentaRet = miReader.GetDouble(23);
						double dRentaBase = miReader.GetDouble(24);
						double dRentaPorc = miReader.GetDouble(25);
						string stCodRenta = miReader.GetString(26);

						#region  Retencion de varias Facturas
						if (idPagoLot > 0)
						{
							dRentaRet = dRetenidoLote;
							if (dRentaPorc >= 0.01 && dRentaRet >= 0.01) 
							{
								dRentaBase = dSubtotalExcento + dSubtotalIva;
								dRentaRet = dRentaPorc * dRentaBase / 100.0;
								// Graba valores de lotes stcodrenta, dRentaPorc e idLote para uso futuro
								miLoteidPagoLot[miLoteiCuenta] = idPagoLot;
								miLoteStCodRenta[miLoteiCuenta] = stCodRenta;
								miLotedRentaPorc[miLoteiCuenta] = dRentaPorc;
								miLoteiCuenta ++;
							}
							else // Busca Primer idPagoLot
							{
								for (int iLot =0; iLot< miLoteiCuenta; iLot++)
								{
									if (miLoteidPagoLot[iLot] == idPagoLot)
									{
										stCodRenta = miLoteStCodRenta[iLot];
										dRentaPorc = miLotedRentaPorc[iLot];
										//dRentaBase = dRentaRet / dRentaPorc * 100.0;
										break;
									} //Fin de if
								} // Fin de For
								dRentaBase = dSubtotalExcento + dSubtotalIva;
								dRentaRet = dRentaPorc * dRentaBase / 100.0;

							} // Busca Primer idPagoLot
						}
						#endregion

						if (dRentaBase < .0001) 
						{
							if (i < iCuentaRet - 1) miReader.Read();
							continue;
						}
						dRentaRetTotal += dRentaRet;

						stAcCodRenta[i] = stCodRenta;
						dAcRentaBase[i] = dRentaBase;
						dAcRentaPorc[i] = dRentaPorc;
						dAcRentaRet[i] = dRentaRet;
						if (i < iCuentaRet - 1) miReader.Read();
					}

					for (int i=0; i<iCuentaRet; i++)
					{
						for (int j=i + 1; j<iCuentaRet; j++)
						{
							if (stAcCodRenta[i] == stAcCodRenta[j])
							{
								dAcRentaBase[i] += dAcRentaBase[j];
								dAcRentaRet[i] += dAcRentaRet[j];
								stAcCodRenta[j] = "NoE";
							}
						}
					}

					for (int i=0; i<iCuentaRet; i++)
					{
						if (stAcCodRenta[i] == null) continue;
						if (stAcCodRenta[i].Trim() == "0")
						{
							MessageBox.Show("Error en Codigo de Retencion", "Información");
//							new Compra(IdTipoFactura, IdCompra).Show();
							miXML.Close();
							Cursor = Cursors.Default;
							return;
						}
						if (stAcCodRenta[i] == "NoE") continue;
						miXML.WriteStartElement("detalleAir");
						miXML.WriteElementString("codRetAir", stAcCodRenta[i].Trim());
						miXML.WriteElementString("baseImpAir", dAcRentaBase[i].ToString("0.00", us));
						miXML.WriteElementString("porcentajeAir", dAcRentaPorc[i].ToString("0.00", us));
						miXML.WriteElementString("valRetAir", dAcRentaRet[i].ToString("0.00", us));
						miXML.WriteEndElement(); // Fin de air
					}
				}
				miXML.WriteEndElement(); // Fin de air
				#endregion Impuesto a la renta
			#region Codigos de Retencion
				if ((iCuentaRet > 0 || bRetPeriodo) && IdTipoFactura == 4 && !(bTarjeta || bNoAplica) )
				{
					string stSerieRet = miReader.GetString(27);
					string stAutRet = miReader.GetString(28);
					string stCodRet = miReader.GetString(29);
					if (idPagoLot > 0)
					{
						if (stSerieRet.Trim().Length > 0)
						{
							miLoteStSerieRet[miLoteiCuenta-1] = stSerieRet;
							miLoteStAutRet[miLoteiCuenta-1] = stAutRet;
							miLoteStCodRet[miLoteiCuenta-1] = stCodRet;
						}
						else
						{
							for (int iLot =0; iLot< miLoteiCuenta; iLot++)
							{
								if (miLoteidPagoLot[iLot] == idPagoLot)
								{
									stSerieRet = miLoteStSerieRet[iLot];
									stAutRet = miLoteStAutRet[iLot];
									stCodRet = miLoteStCodRet[iLot];
									break;
								} //Fin de if
							} // Fin de For
						} // fin de else
					}
					stSerieRet = Digito(stSerieRet);
					stAutRet = Digito(stAutRet);
					stCodRet = Digito(stCodRet);

					if (dRentaRetTotal >= 0.01 && (stSerieRet.Trim().Length != 6 || stAutRet.Trim().Length != 10 || stCodRet.Trim().Length < 1))
					{
						MessageBox.Show("Error en Serie o Numero de Retencion.", "Información");
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
	
					// Fecha de Emision de Retencion mayor o igual que factura
					DateTime dtFechaNula = DateTime.Today.AddYears(-100);
					dtFechaRet = dtFechaRet.Date;
					dtFechaCompra = dtFechaCompra.Date;
					if (dtFechaRet < dtFechaCompra && 
						!(dtFechaRet > dtFechaNula.AddDays(-1) && dtFechaRet < dtFechaNula.AddDays(1)))
					{
						MessageBox.Show("Fecha de Retencion debe ser mayor que " + dtFechaCompra.ToString("dd/MMM/yyyy"), "Información");
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
					if (stSerieRet.Trim().Length == 6 && stCodRet.Trim().Length > 0
						&& (dRentaRetTotal + dIvaBienes + dvalorRetServicios + dvalorRetServ100 >= 0.01)) // || bRetPeriodo))
					{
						miXML.WriteElementString("estabRetencion1", stSerieRet.Substring(0,3));
						miXML.WriteElementString("ptoEmiRetencion1", stSerieRet.Substring(3,3));
						miXML.WriteElementString("secRetencion1", stCodRet);
						miXML.WriteElementString("autRetencion1", stAutRet);
						miXML.WriteElementString("fechaEmiRet1", dtFechaRet.ToString("dd/MM/yyyy"));
					}
				}
				#endregion Codigos de Retencion

			#region Devoluciones de compras
				if (IdTipoFactura == 6 && stDevComprob[k] != null)
				{
					try
					{
						if (stDevComprob[k].Trim().Length > 15 || stDevComprob[k].Trim().Length < 1)
						{
							MessageBox.Show("Error en Comprobante Asociado", "Información");
//							new Compra(4, IdCompra1[k]).Show();
							//new Compra(6, IdCompra).Show();
							miXML.Close();
							Cursor = Cursors.Default;
							return;
						}

						if (stDevComprob[k].Trim().Length == 15)
						{
							stDevComprob[k] = stDevComprob[k].Substring(6, 9);
						}
						int iSerie = int.Parse(stDevSerie[k].Substring(0,3));
						int iPuntoE = int.Parse(stDevSerie[k].Substring(3,3));
						if (iSerie < 1 || iSerie > 999 || iPuntoE < 1 || iPuntoE > 999)
						{
							MessageBox.Show("Error en Serie de Comprobante Asociado", "Información");
//							new Compra(6, IdCompra).Show();
							miXML.Close();
							Cursor = Cursors.Default;
							return;
						}
						if (Int64.Parse(stDevAutoriza[k]) == 0 || stDevAutoriza[k].Length != 10)
						{
							MessageBox.Show("Error en Autorizacion de Comprobante Asociado", "Información");
//							new Compra(6, IdCompra).Show();
							miXML.Close();
							Cursor = Cursors.Default;
							return;
						}
						if (Int64.Parse(stDevComprob[k]) == 0 || stDevComprob[k].Length > 15)
						{
							MessageBox.Show("Error en Secuencial de Comprobante Asociado: " + stDevComprob[k], "Información");
//							new Compra(6, IdCompra).Show();
							miXML.Close();
							Cursor = Cursors.Default;
							return;
						}

						miXML.WriteElementString("docModificado", stDevCompTipo[k].ToString());// int.Parse(stDevCompTipo[k]).ToString());
						miXML.WriteElementString("estabModificado", stDevSerie[k].Substring(0,3));
						miXML.WriteElementString("ptoEmiModificado", stDevSerie[k].Substring(3,3));
						miXML.WriteElementString("secModificado", stDevComprob[k].Trim());
						miXML.WriteElementString("autModificado", stDevAutoriza[k]);
					}
					catch (System.Exception ex)
					{
						MessageBox.Show(ex.Message, "No existe comprobante");
//						new Compra(6, IdCompra).Show();
						miXML.Close();
						Cursor = Cursors.Default;
						return;
					}
					k++;
				}
				else
				{
					miXML.WriteElementString("docModificado", "00");
					miXML.WriteElementString("estabModificado", "000");
					miXML.WriteElementString("ptoEmiModificado", "000");
//					miXML.WriteElementString("secModificado","0000000"); Problema en Sol y Luna
					miXML.WriteElementString("secModificado","0");
					miXML.WriteElementString("autModificado", "0000000000");
				}
				#endregion fin de devoluciones de compras
				miXML.WriteEndElement(); // Fin de DetalleCompras
			}
			miReader.Close();
			#endregion Fin de Compras

			#region Retencion de Tarjetas en Compras
			stSelect = "Select Max(Cl.Ruc), Sum(IsNull(R.RentaRetenida, 0) "
				+ "+ IsNull(R.RentaRetenida1, 0)), "
				+ "Sum(IsNull(R.IvaTotal, 0)), Sum(IsNull(R.IvaRetenido, 0)), "
				+ "Max(IsNull(R.idRetIva, 0)), Sum(IsNull(R.IvaTotal1, 0)), "
				+ "Sum(IsNull(R.IvaRetenido1, 0)), Max(IsNull(R.idRetIva1, 0)) "
				+ "From Retencion r Inner Join Cliente cl On R.idCliente = Cl.idCliente "
				+ "where R.Tarjeta = 1 And R.NoAplica = 0 "
				+ "And (IsNull(R.IvaRetenido, 0) +IsNull(R.IvaRetenido1, 0) >= 0) "
				+ "And R.IngresoEgreso = 2 "
				+ "group by Cl.Ruc, R.idRetIva, R.idRetIva1";
			miReader = Funcion.rEscalarSQL(cdsCompra, stSelect, true);
			while(miReader.Read())
			{
				string stRucTarjeta = miReader.GetString(0);
				double dRentaTarjeta = miReader.GetDouble(1);
				double dIvaTarjeta = miReader.GetDouble(2);
				double dIvaRetTarjeta = miReader.GetDouble(3);
				int iIdRetTarjeta = miReader.GetInt32(4);
				double dIva1Tarjeta = miReader.GetDouble(5);
				double dIvaRet1Tarjeta = miReader.GetDouble(6);
				int iIdRet1Tarjeta = miReader.GetInt32(7);
			}
			miReader.Close();
			#endregion Retencion de Tarjetas en Compras

			#region Tiempo Compras
			TimeSpan ts = DateTime.Now.Subtract(dtInicio);
			stTiempos += string.Format("\nCompras: {0}:{1} mm:ss", ts.Minutes.ToString("n0"), ts.Seconds.ToString("n0"));
			dtInicio = DateTime.Now;
			#endregion Tiempo Compras

			string[] stCodTipoRuc1 = new string[iMax];
			string[] stRuc1 = new string[iMax];
			string[] stComprobante = new string[iMax];

			int j1 = -1;
			double[] dIvaBienes1 = new double[iMax];

			miXML.WriteEndElement(); // Fin de Compras
			DateTime dtFechaFinMes = dtFecha;
			miXML.WriteStartElement("ventas");
			if (!chkExcluirVenta.Checked)
			{
				#region Venta
				#region Select de Venta
				stSelect = "SELECT IsNull(ClienteTipoRuc.Codigo, '0'), "
				+ "IsNull(Cliente.Ruc, '0'), "
				+ "IsNull(Cliente.IdCliente, 0), "
				+ "Count(Distinct Compra.idCompra), "
				+ "IsNull(Round(sum(Compra.SubTotalExcento - Compra.Descuento0), 2), 0), "
				+ "IsNull(Round(sum(Compra.SubTotalIva - Compra.Descuento), 2), 0), "
				+ "IsNUll(Round(sum(Compra.Iva), 2), 0), "
				+ "IsNUll(Sum(Compra.Ice), 0), MAX(Compra.Ice), "
				+ "IsNull(Sum(Compra.IvaServicio), 0), "
				+ "IsNull(Sum(Retencion.IvaRetenido), 0), "
				+ "IsNull(MAX(RetencionIva.Porcentaje), 0), "
				+ "IsNull(Sum(Retencion.IvaRetenido1), 0), "
				+ "IsNull(MAX(RetIva_1.Porcentaje), 0), "
				+ "IsNull(MAX(CompraComprobante.Codigo), '18'), "
				+ "IsNull(Compra.idTipoFactura, 0), "
				+ "IsNull(Sum(Compra.Servicio), 0) "
				+ "FROM Compra LEFT JOIN Retencion ON Compra.idCompra = Retencion.idCompra "
				+ "Left Join CompraComprobante On Compra.idComprobante = CompraComprobante.idComprobante "
				+ "Inner Join Cliente On Compra.idCliente = Cliente.idCliente "
				+ "LEFT JOIN RetencionIva ON Retencion.idRetIva = RetencionIva.idRetIva "
				+ "LEFT JOIN RetencionIva As RetIva_1  ON Retencion.idRetIva1 = RetIva_1.idRetIva "
				+ "Left Join ClienteTipoRuc On Cliente.idTipoRuc = ClienteTipoRuc.idTipoRuc "
				+ "Where Compra.idTipoFactura In (1, 5, 27) And Compra.Borrar = 0 "
				+ "And Compra.Fecha >= '" + dtFecha.ToString("yyyyMMdd") 
				+ "' And Not CompraComprobante.Codigo In ('17', '99', '15', '16', '98') "
				+ " And Compra.Fecha < '" + dtHasta.AddDays(1).ToString("yyyyMMdd") + "' And Estado In (10, 11) "
				+ "GROUP BY IsNull(Cliente.IdCliente, 0), ClienteTipoRuc.Codigo, Cliente.Ruc, Compra.idComprobante, Compra.IdTipoFactura ";

				#region Select comentada de ventas
				//				stSelect = "SELECT IsNull(Max(ClienteTipoRuc.Codigo), '0'), "
//					+ "IsNull(Substring(Max(Cliente.Ruc),1, 10), '0'), "
//					+ "IsNull(Max(Cliente.IdCliente), 0), "
//					+ "Count(Distinct Compra.idCompra), "
//					+ "IsNull(Round(sum(Compra.SubTotalExcento - Compra.Descuento0), 2), 0), "
//					+ "IsNull(Round(sum(Compra.SubTotalIva - Compra.Descuento), 2), 0), "
//					+ "IsNUll(Round(sum(Compra.Iva), 2), 0), "
//					+ "IsNUll(Sum(Compra.Ice), 0), MAX(Compra.Ice), "
//					+ "IsNull(Sum(Compra.IvaServicio), 0), "
//					+ "IsNull(Sum(Retencion.IvaRetenido), 0), "
//					+ "IsNull(MAX(RetencionIva.Porcentaje), 0), "
//					+ "IsNull(Sum(Retencion.IvaRetenido1), 0), "
//					+ "IsNull(MAX(RetIva_1.Porcentaje), 0), "
//					+ "IsNull(MAX(CompraComprobante.Codigo), '18'), "
//					+ "IsNull(Compra.idTipoFactura, 0), "
//					+ "IsNull(Sum(Compra.Servicio), 0) "
//					+ "FROM Compra LEFT JOIN Retencion ON Compra.idCompra = Retencion.idCompra "
//					+ "Left Join CompraComprobante On Compra.idComprobante = CompraComprobante.idComprobante "
//					+ "Inner Join Cliente On Compra.idCliente = Cliente.idCliente "
//					+ "LEFT JOIN RetencionIva ON Retencion.idRetIva = RetencionIva.idRetIva "
//					+ "LEFT JOIN RetencionIva As RetIva_1  ON Retencion.idRetIva1 = RetIva_1.idRetIva "
//					+ "Left Join ClienteTipoRuc On Cliente.idTipoRuc = ClienteTipoRuc.idTipoRuc "
//					+ "Where Compra.idTipoFactura In (1, 5, 27) And Compra.Borrar = 0 "
//					+ "And Compra.Fecha >= '" + dtFecha.ToString("yyyyMMdd") 
//					+ "' And Not CompraComprobante.Codigo In ('17', '99', '15', '16', '98') "
//					+ " And Compra.Fecha < '" + dtHasta.AddDays(1).ToString("yyyyMMdd") + "' And Estado In (10, 11) "
//					+ "GROUP BY Substring(Cliente.Ruc,1, 10), Compra.idComprobante, Compra.IdTipoFactura ";
				//				+ "And Retencion.IngresoEgreso = 1 "
				#endregion Select comentada de ventas
				#endregion asignacion variables
				#region Creacion de variables
				miReader = Funcion.rEscalarSQL(cdsCompra, stSelect, true);
				dtFechaFinMes = dtFechaFinMes.AddMonths(1);
				dtFechaFinMes = dtFechaFinMes.AddDays(-dtFechaFinMes.Day);

				int [] iCuentaComprob1 = new int [iMax];
				double [] dSubtotalExcento1 = new double [iMax];
				double [] dSubtotalIva1 = new double [iMax];
				double [] dIva1 = new double [iMax];
				double [] dIvaServicio1 = new double [iMax];
				double [] dRetValorIva2 = new double [iMax];
				int [] dRetIvaPorc2 = new int [iMax];
				double [] dRetValorIva11 = new double [iMax];
				int [] dRetIvaPorc11 = new int [iMax];
				int [] iCuentaRet1 = new int [iMax];
				int [] IdCliente1 = new int [iMax];
				int [] IdTipoFac = new int [iMax];
				string [] stPasaporte = new string [iMax];
				double [] dServicio = new double [iMax];
				//			string [] stComprobante = new string [iMax];
				k = 0;
				#endregion creacion de variables
				while (miReader.Read())
				{
					#region asignacion variables
					stCodTipoRuc1[k] = miReader.GetString(0);
					stRuc1[k] = miReader.GetString(1).Trim();
					stPasaporte[k] = stRuc1[k];
					iCuentaComprob1[k] = miReader.GetInt32(3);
					dSubtotalExcento1[k] = Math.Round(miReader.GetDouble(4), 2);
					dSubtotalIva1[k] = Math.Round(miReader.GetDouble(5), 2);
					dIva1[k] = Math.Round(miReader.GetDouble(6), 2);
					dIvaServicio1[k] = Math.Round(miReader.GetDouble(9), 2);
					dRetValorIva2[k] = Math.Round(miReader.GetDouble(10), 2);
					dRetIvaPorc2[k] = int.Parse(miReader.GetDouble(11).ToString());
					dRetValorIva11[k] = Math.Round(miReader.GetDouble(12), 2);
					dRetIvaPorc11[k] = int.Parse(miReader.GetDouble(13).ToString());
					stComprobante[k] = miReader.GetString(14).ToString().Trim();
					IdTipoFac[k] = miReader.GetInt32(15);
					dServicio[k] = miReader.GetDouble(16);
					if (dRetIvaPorc2[k] == 0) dRetIvaPorc2[k] = 0;
					if (dRetIvaPorc2[k] == 30) dRetIvaPorc2[k] = 1;
					if (dRetIvaPorc2[k] == 70) dRetIvaPorc2[k] = 2;
					if (dRetIvaPorc2[k] == 100) dRetIvaPorc2[k] = 3;

					if (dRetIvaPorc11[k] == 0) dRetIvaPorc11[k] = 0;
					if (dRetIvaPorc11[k] == 30) dRetIvaPorc11[k] = 1;
					if (dRetIvaPorc11[k] == 70) dRetIvaPorc11[k] = 2;
					if (dRetIvaPorc11[k] == 100) dRetIvaPorc11[k] = 3;

					#endregion asignacion variables
					#region Verificaciones
					IdCliente1[k] = miReader.GetInt32(2);

					if (stCodTipoRuc1[k].Trim() == "0" ) 
					{
						MessageBox.Show("Codigo de Tipo de Ruc Erroneo", stVerificaRuc);
						new Cliente(true, IdCliente1[k]).Show();
						miXML.Close();
						Cursor = Cursors.Default;					
						return;
					}
					if (IdTipoFac[k] == 5 && !(stComprobante[k] == "04" || stComprobante[k] == "05") ) // Devolucion de venta
					{
						MessageBox.Show("Devolucion en Venta debe tener comprobante nota de credito", stVerificaRuc);
						new Cliente(true, IdCliente1[k]).Show();
						miXML.Close();
						Cursor = Cursors.Default;					
						return;
					}
					if (IdTipoFac[k] == 1 && !(stComprobante[k] == "18" || stComprobante[k] == "41" || 
						stComprobante[k] == "04" || stComprobante[k] == "05" || stComprobante[k] == "02") ) // Devolucion de venta
					{
						MessageBox.Show("Venta debe tener Documentos Autorizados o Reembolso o Nota de Debito/Credito", stVerificaRuc);
						new Cliente(true, IdCliente1[k]).Show();
						miXML.Close();
						Cursor = Cursors.Default;					
						return;
					}
					if (int.Parse(stCodTipoRuc1[k]) < 4) 
						stCodTipoRuc1[k] = "0" + Convert.ToString(int.Parse(stCodTipoRuc1[k]) + 3);
					stVerificaRuc = "";
					if (stCodTipoRuc1[k] == "07") stRuc1[k] = "9999999999999";
					else stVerificaRuc = Funcion.VerificaCedula(stRuc1[k], false);

					if (stCodTipoRuc1[k] == "06") stRuc1[k] = Digito(stRuc1[k]);
					if (stCodTipoRuc1[k] == "04" && stRuc1[k].Length ==10) stRuc1[k] += "001";
					if (stCodTipoRuc1[k] == "04" && stRuc1[k].Length !=13) stVerificaRuc = "Longitud de ruc debe ser de 13 digitos";
					if (stCodTipoRuc1[k] == "05" && stRuc1[k].Length !=10) stVerificaRuc = "Longitud de cedula debe ser de 10 digitos";
					if  (stVerificaRuc.Length > 0 && !(stCodTipoRuc1[k] == "07" || stCodTipoRuc1[k] == "08" || stCodTipoRuc1[k] == "09" || stCodTipoRuc1[k] == "03" || (stCodTipoRuc1[k] == "06"))) // && stRuc1[k].Length > 4)) )
					{
						MessageBox.Show("Ruc o Cedula Erroneo", stVerificaRuc);
						new Cliente(true, IdCliente1[k]).Show();
						miXML.Close();
						Cursor = Cursors.Default;
						return;
					}
					#endregion verificacion
					k++;
				}
				miReader.Close();
				double dIvaPorc = Funcion.dEscalarSQL(cdsCompra, "Select isnull(iva, 12) from seteof") / 100.0;
				for (int j=0; j< k; j++)
				{
					miXML.WriteStartElement("detalleVentas");
					#region Generales
					miXML.WriteElementString("tpIdCliente", stCodTipoRuc1[j]);
					if (stCodTipoRuc1[j] == "04" && stRuc1[j].Length ==13) miXML.WriteElementString("idCliente", stRuc1[j].Substring(0,10) + "001");
					else
					{
						if (stCodTipoRuc1[j] == "06") miXML.WriteElementString("idCliente", stPasaporte[j]);
						else miXML.WriteElementString("idCliente", stRuc1[j]);
					}
//					if (stComprobante[j]=="02") stComprobante[j]="18";
					miXML.WriteElementString("tipoComprobante", int.Parse(stComprobante[j]).ToString());
					miXML.WriteElementString("numeroComprobantes", iCuentaComprob1[j].ToString());

					miXML.WriteElementString("baseNoGraIva", dServicio[j].ToString("0.00", us));
					miXML.WriteElementString("baseImponible", dSubtotalExcento1[j].ToString("0.00", us));
					miXML.WriteElementString("baseImpGrav", dSubtotalIva1[j].ToString("0.00", us));
					dIva1[j] = dIvaPorc * dSubtotalIva1[j];
					miXML.WriteElementString("montoIva", dIva1[j].ToString("0.00", us));
					#endregion Generales

					double dIvaBienes = dIva1[j] - dIvaServicio1[j];
					dRetValorIva2[j] = dIvaBienes * dRetIvaPorc2[j];

					RetencionIvaV2008(dIvaServicio1[j], stRuc1[j], miXML, 1, stComprobante[j]);
					RetencionRentaV2008(miXML, stRuc1[j], 1, stComprobante[j]);

					miXML.WriteEndElement(); // Fin de Detalle de Ventas
				}
				miXML.WriteEndElement(); 
				#endregion Ventas

				#region Venta Establecimiento
				miXML.WriteStartElement("ventasEstablecimiento");
				int iContBodegas = (int)Funcion.iEscalarSQL(cdsSeteo, "Select Count(*) From Bodega");
				bool bFactura = false;
				for (int l = 1; l <= iContBodegas; l++)
				{	
					if (l != 50) bFactura = Funcion.bEscalarSQL(cdsSeteo, string.Format("Select IsNull(Factura, 0) From Bodega Where Bodega = {0}", l), true);
					else bFactura = false;

					if (bFactura)
					{
						miXML.WriteStartElement("ventaEst");//+ Iva - DescPorcIva  - DescPorcIva0
						miXML.WriteElementString("codEstab", Funcion.sEscalarSQL(cdsSeteo, string.Format("Select Establecimiento From Bodega Where Bodega = {0} And Activo = 1 And Factura = 1", l)));
						miXML.WriteElementString("ventasEstab", Funcion.dEscalarSQL(cdsSeteo, string.Format("Select ISNULL(SUM((SubtotalIva - Descuento) + (SubtotalExcento - Descuento0)), 0) From Compra Where idTipoFactura = 1 And Borrar = 0 And Estado In (10, 11) And Fecha >= '{1}' And Fecha <= '{2}' And Bodega = {0}", l, dtFecha.ToString("yyyyMMdd"), dtHasta.ToString("yyyyMMdd"))).ToString("0.00", us));
						miXML.WriteEndElement();
					}
				}
				#endregion Venta Establecimiento

				#region Tarjeta en Ventas
				stSelect = "Exec TarjetaAts '" + dtFecha.ToString("yyyyMMdd") + "', '" + dtHasta.ToString("yyyyMMdd") + "'";
				miReader = Funcion.rEscalarSQL(cdsCompra, stSelect, true);
				int k1 = 0;
				while (miReader.Read())
				{
					#region asignacion variables
					stRuc1[k1] = miReader.GetString(0).Trim();
					iCuentaComprob1[k1] = miReader.GetInt32(1);
					dRetValorIva11[k1] = Math.Round(miReader.GetDouble(2), 2);
					#endregion asignacion variables
					k1++;
				}
				miReader.Close();

				for (int j=0; j< k1; j++)
				{
					miXML.WriteStartElement("detalleVentas");
					#region Generales
					miXML.WriteElementString("tpIdCliente", "04");
					miXML.WriteElementString("idCliente", stRuc1[j]);
					miXML.WriteElementString("tipoComprobante", "18");
					miXML.WriteElementString("numeroComprobantes", iCuentaComprob1[j].ToString());

					miXML.WriteElementString("baseNoGraIva", "0.00");
					miXML.WriteElementString("baseImponible", "0.00");
					miXML.WriteElementString("baseImpGrav", "0.00");
					miXML.WriteElementString("montoIva", "0.00");
					#endregion Generales

					miXML.WriteElementString("valorRetIva", "0.00");
					miXML.WriteElementString("valorRetRenta", dRetValorIva11[j].ToString("0.00", us));

					miXML.WriteEndElement(); // Fin de Detalle de Tarjeta en Ventas
				}
				#endregion Tarjeta en Ventas

				#region Retencion que no coincide con fecha
				#region Select de Retenciones de otro periodo
				stSelect = "Select Max(IsNull(ClienteTipoRuc.Codigo, '')), Max(IsNull(Cliente.Ruc, '')), "
					+ "IsNull(Max(CompraComprobante.Codigo), '18')  "
					+ "From Retencion inner Join RetencionIva On Retencion.idRetIva = RetencionIva.idRetIva "
					+ "Inner Join Compra On Compra.idCompra = Retencion.idCompra "
					+ "Inner Join CompraComprobante On CompraComprobante.idComprobante = Compra.idComprobante "
					+ "Inner Join Cliente On Cliente.idCliente = Compra.idCliente "
					+ "Inner Join ClienteTipoRuc On ClienteTipoRuc.idTipoRuc = Cliente.idTipoRuc "
					+ "Where Retencion.IngresoEgreso = 1 And Retencion.Borrar = 0 "
					+ "And Compra.Borrar = 0 And " + stFechaRet + " >= '" + dtFecha.ToString("yyyyMMdd") + "' "
					+ "And " + stFechaRet + " < '" + dtHasta.AddDays(1).ToString("yyyyMMdd") + "' "
					+ "And Compra.Fecha < '" + dtFecha.ToString("yyyyMMdd") 
					+ "' And (Retencion.IvaRetenido + Retencion.IvaRetenido1 + Retencion.RentaRetenida) > 0 "
					+ "Group by Cliente.Ruc ";
				miReader = Funcion.rEscalarSQL(cdsCompra, stSelect, true);
				#endregion Select de Retenciones de otro periodo
				#region Lectura de Variables
				j1 = -1;
				//			double[] dIvaBienes1 = new double[1000];
				while(miReader.Read())
				{
					j1++;
					stCodTipoRuc1[j1] = miReader.GetString(0).Trim();
					if (int.Parse(miReader.GetString(0)) < 4)
					{
						int iCod = int.Parse(miReader.GetString(0))+3;
						stCodTipoRuc1[j1] = "0" + iCod.ToString();
					}
					stRuc1[j1] = miReader.GetString(1).Trim();
					stComprobante[j1] = miReader.GetString(2).Trim();
				}
				miReader.Close();
				#endregion Lectura de Variables
				for (int i = 0; i <= j1; i++)
				{
					if (ExisteAntes(stRuc1[i])) continue;
					miXML.WriteStartElement("detalleVentas");
					#region Generales
					miXML.WriteElementString("tpIdCliente", stCodTipoRuc1[i]);
					if (stCodTipoRuc1[i] == "04" && stRuc1[i].Length ==13) 
						miXML.WriteElementString("idCliente", stRuc1[i].Substring(0,10) + "001");
					else
						miXML.WriteElementString("idCliente", stRuc1[i]);
					miXML.WriteElementString("tipoComprobante", int.Parse(stComprobante[i]).ToString());
					miXML.WriteElementString("numeroComprobantes", "0");

					miXML.WriteElementString("baseNoGraIva", "0.00");
					miXML.WriteElementString("baseImponible", "0.00");
					miXML.WriteElementString("baseImpGrav", "0.00"); // dSubtotalIva.ToString("0.00", us)); Grijalva
					miXML.WriteElementString("montoIva", "0.00");
					#endregion Generales   
					RetencionIvaV2008(0, stRuc1[i], miXML, 2, stComprobante[i]);
					RetencionRentaV2008(miXML, stRuc1[i], 2, stComprobante[i]);
					miXML.WriteEndElement(); // Fin de Detalle de Ventas
				}
				miReader.Close();
				#endregion Retencion que no coincide con fecha
			}
			miXML.WriteEndElement(); // Fin de Ventas

			#region importaciones
			int iLinea = -1;
						string[] stRefrendo = new string[1000];
			#endregion importaciones

			#region exportaciones
			stSelect = "SELECT Compra.Fecha, Compra.BienServicio, "
				+ "CompraComprobante.Codigo, IsNull(Compra.Refrendo, ''), "
				+ "IsNull(Cliente.Nombre, ''), IsNull(Cliente.Ruc, ''), IsNull(Compra.Total, 0) - IsNull(Compra.Servicio, 0) - IsNull(Compra.Propina, 0), "
				+ "Compra.FechaCaducidad, Compra.DerechoDevolucion, "
				+ "IsNull(Compra.SerieFactura, ''), IsNull(Compra.AutFactura, ''), "
				+ "IsNull(Compra.Numero, ''), Compra.idCompra, Compra.idTipoFactura, IsNull(Compra.Departamento, ''), "
				+ "IsNull(Compra.Comprob1, '') "
				+ "FROM Compra Inner Join Cliente On Compra.idCliente = Cliente.idCliente "
				+ "LEFT JOIN Retencion ON Compra.idCompra = Retencion.idCompra "
				+ "LEFT OUTER JOIN CompraComprobante ON Compra.idComprobante = "
				+ "CompraComprobante.idComprobante "
				+ "LEFT OUTER JOIN CompraCredTribut ON Compra.idCredTributario = "
				+ "CompraCredTribut.idCredTributario Where Compra.idTipoFactura In (1, 5) "
				+ "And Compra.Fecha >= '" + dtFecha.ToString("yyyyMMdd") 
				+ "' And Compra.Fecha < '" + dtHasta.AddDays(1).ToString("yyyyMMdd")
				+ "' And Compra.Borrar = 0 AND CompraComprobante.Codigo In ('16', '98')";
			
			miXML.WriteStartElement("exportaciones");
			miReader = Funcion.rEscalarSQL(cdsCompra, stSelect, true);
			while(miReader.Read())
			{
				miXML.WriteStartElement("detalleExportaciones");
				DateTime dtFechaExp = miReader.GetDateTime(0);
				int IdCompraBienS = miReader.GetInt32(1);
//				int IdCompraBienS = 3 - miReader.GetInt32(1);
				string stCompTipo = miReader.GetString(2);
				if (stCompTipo == "98") stCompTipo = "1";
				string stRefrendoExp = miReader.GetString(3).Trim();
				string stNombreExp = miReader.GetString(4).Trim();
				string stRucExp = miReader.GetString(5).Trim();
				double dValoFob = miReader.GetDouble(6);
				int idCompraExp = miReader.GetInt32(12);
				int idTipoFacturaExp = miReader.GetInt32(13);
				string stFue = miReader.GetString(14);
				string stDocTransp = miReader.GetString(15);

				string stSerieExp = miReader.GetString(9);
				string stAutorizaExp = miReader.GetString(10);
				string stSecuenciaExp = miReader.GetString(11);
				stSerieExp = Digito(stSerieExp);
				stAutorizaExp = Digito(stAutorizaExp);
				stSecuenciaExp = Digito(stSecuenciaExp);
				if (stSerieExp.Length != 6 || stAutorizaExp.Length != 10 || stSecuenciaExp.Length > 7)
				{
					MessageBox.Show("Error en serie, autorizacion o secuencial", "Información");
//					Compra miCompra = new Compra(idTipoFacturaExp, idCompraExp);
//					miCompra.Show();
					miXML.Close();
					Cursor = Cursors.Default;
					return;
				}
//				if (stFue.Length != 13)
//				{
//					MessageBox.Show("Error en FUE debe tener 13 digitos. Pestaña Entrega frente a botón Importación", "Información");
//					Compra miCompra = new Compra(idTipoFacturaExp, idCompraExp);
//					miCompra.Show();
//					miXML.Close();
//					Cursor = Cursors.Default;
//					return;
//				}

				if (stRefrendoExp.Length != 16 && IdCompraBienS == 1)
				{
					MessageBox.Show("Error en refrendo (Pestaña A.T.) longitud debe ser de 16 en Exportacion de Bienes", "Información");
//					new Compra(idTipoFacturaExp, idCompraExp).Show();
					miXML.Close();
					Cursor = Cursors.Default;
					return;
				}
				if (IdCompraBienS == 2)
				{
					if (iLinea < 0) iLinea = 0;
					stRefrendoExp = "                ";
				}
				int idCambio = IdCompraBienS;
				miXML.WriteElementString("exportacionDe", idCambio.ToString());
				if (idTipoFacturaExp == 5) miXML.WriteElementString("tipoComprobante", "4");
				else miXML.WriteElementString("tipoComprobante", stCompTipo);

				if (IdCompraBienS == 1)
				{
					if (stDocTransp.Trim().Length == 0)
					{
						MessageBox.Show("Error en Doc. Transportista (Pestaña Entrega - Abajo de Boton Comprob1) longitud debe ser mayor a 1", "Información");
//						new Compra(idTipoFacturaExp, idCompraExp).Show();
						miXML.Close();
						Cursor = Cursors.Default;
						return;
					}

					miXML.WriteElementString("distAduanero", stRefrendoExp.Substring(0,3));
					miXML.WriteElementString("anio", stRefrendoExp.Substring(3,4));
					miXML.WriteElementString("regimen", stRefrendoExp.Substring(7,2));
					miXML.WriteElementString("correlativo", stRefrendoExp.Substring(9,7));
//					miXML.WriteElementString("verificador", stRefrendoExp.Substring(15,1));
					stDocTransp = stDocTransp.PadLeft(13, '0');
					miXML.WriteElementString("docTransp", stDocTransp);
				}

				miXML.WriteElementString("fechaEmbarque", dtFechaExp.ToString("dd/MM/yyyy"));
//				miXML.WriteElementString("fue", stFue);
				miXML.WriteElementString("valorFOB", dValoFob.ToString("0.00", us));
				miXML.WriteElementString("valorFOBComprobante", dValoFob.ToString("0.00", us));
				miXML.WriteElementString("establecimiento", stSerieExp.Substring(0,3));
				miXML.WriteElementString("puntoEmision", stSerieExp.Substring(3,3));
				miXML.WriteElementString("secuencial", stSecuenciaExp);
				miXML.WriteElementString("autorizacion", stAutorizaExp);
				miXML.WriteElementString("fechaEmision", dtFechaExp.ToString("dd/MM/yyyy"));
				miXML.WriteEndElement(); // Fin de de exportaciones
			}
			miReader.Close();
			miXML.WriteEndElement(); // Fin de exportaciones
			#endregion exportaciones

			#region anulados
			#region Ventas y Devoluciones
			miXML.WriteStartElement("anulados");
			stSelect = "SELECT IsNull(CompraComprobante.Codigo, ''), "
				+ "IsNull(Compra.Numero, ''), "
				+ "IsNull(Compra.SerieFactura, ''), "
				+ "IsNull(Compra.AutFactura, ''), "
				+ "IsNull(Compra.FechaCaducidad, dateadd(year, -100, getdate())), "
				+ "IsNull(Compra.Fecha, dateadd(year, -100, getdate())), "
				+ "IsNull(Compra.idTipoFactura, 0), IsNull(Compra.idCompra, 0) "
				+ " FROM Compra Inner Join Cliente On Compra.idCliente = Cliente.idCliente "
				+ " LEFT JOIN Retencion ON Compra.idCompra = Retencion.idCompra "
				+ " Inner Join CompraComprobante On CompraComprobante.idComprobante = Compra.idComprobante "
				+ " Where Compra.idTipoFactura  in (1, 5) And Compra.Borrar = 1 "
				+ " And Not CompraComprobante.Codigo In ('17', '99', '15') "
				+ " And Compra.Fecha >= '" + dtFecha.ToString("yyyyMMdd")
				+ "' And Compra.Fecha < '" + dtHasta.AddDays(1).ToString("yyyyMMdd")
				+ "' ";

			miReader = Funcion.rEscalarSQL(cdsCompra, stSelect, true);
			while(miReader.Read())
			{
				miXML.WriteStartElement("detalleAnulados");
				string stCompTipo = miReader.GetString(0);
				string stNumAnu = miReader.GetString(1);
				string stSerie = miReader.GetString(2);
				string stAutoriza = miReader.GetString(3);
				DateTime dtFechaCad = miReader.GetDateTime(4);
				DateTime dtFechaAnu = miReader.GetDateTime(5);
				int IdTipoF = miReader.GetInt32(6);
				int IdCompraA = miReader.GetInt32(7);

				miXML.WriteElementString("tipoComprobante", int.Parse(stCompTipo).ToString());

				#region Serie
				string stEstablecmiento = "";
				string stPuntoEmision = "";
				string stSecuencial = "";
				stNumAnu = Digito(stNumAnu);
				if (stNumAnu.Length == 15)
				{
					stEstablecmiento = stNumAnu.Substring(0, 3);
					stPuntoEmision = stNumAnu.Substring(3,3);
					stSecuencial = stNumAnu.Substring(6,9);
				}
				else if (stNumAnu.Length >= 10)
				{
					MessageBox.Show("Numero Secuencial Factura " + stNumAnu + " Erronea");
//					new Compra(IdTipoF, IdCompraA).Show();
					miXML.Close();
					Cursor = Cursors.Default;
					return;
				}
				else
				{
					stSerie = Digito(stSerie);
					if (stSerie.Length != 6)
					{
						MessageBox.Show("Numero de Serie de Factura " + stNumAnu + " Erronea", "Información");
//						new Compra(IdTipoF, IdCompraA).Show();
						miXML.Close();
						Cursor = Cursors.Default;
						return;
					}
					stEstablecmiento = stSerie.Substring(0, 3);
					stPuntoEmision = stSerie.Substring(3, 3);
					stNumAnu= Digito(stNumAnu);
					stSecuencial = stNumAnu;
				}
				if (stSecuencial.Length == 0 || Int64.Parse(stSecuencial) == 0)
				{
					MessageBox.Show("Numero de Secuencial de Factura " + stSecuencial + " Erronea", "Información");
//					new Compra(IdTipoF, IdCompraA).Show();
					miXML.Close();
					Cursor = Cursors.Default;
					return;
				}
				miXML.WriteElementString("establecimiento", stEstablecmiento);
				miXML.WriteElementString("puntoEmision", stPuntoEmision);
				miXML.WriteElementString("secuencialInicio", stSecuencial);
				miXML.WriteElementString("secuencialFin", stSecuencial);
				#endregion Serie
		
				stAutoriza = Digito(stAutoriza);
				if (stAutoriza.Length < 4)
				{
					MessageBox.Show("Numero de Autorizacion de Factura " + stNumAnu + " Erronea");
//					new Compra(IdTipoF, IdCompraA).Show();
					miXML.Close();
					Cursor = Cursors.Default;
					return;
				}
				miXML.WriteElementString("autorizacion", stAutoriza);
//				miXML.WriteElementString("fechaAnulacion", dtFechaAnu.ToString("dd/MM/yyyy"));
				miXML.WriteEndElement(); // Fin de Detalle de anulados
			}	
			miReader.Close();

			#endregion Ventas y Devoluciones

			#region Retenciones de Compra
			stSelect = "SELECT IsNull(Retencion.Codigo, ''), "
				+ "IsNull(Retencion.SerieRet, ''), "
				+ "IsNull(Retencion.AutorizaRet, ''), "
				+ "IsNull(Retencio,n.Fecha, dateadd(year, -100, getdate())), "
				+ "IsNull(Retencion.Fecha, dateadd(year, -100, getdate())), "
				+ "IsNull(Retencion.IdRetencion, 0) "
				+ " FROM Retencion where Retencion.Borrar = 1 "
				+ " And Retencion.NoAplica = 0 And Retencion.IngresoEgreso = 2 "
				+ " And Retencion.Fecha >= '" + dtFecha.ToString("yyyyMMdd")
				+ "' And Retencion.Fecha < '" + dtHasta.AddDays(1).ToString("yyyyMMdd")
				+ "' ";

			miReader = Funcion.rEscalarSQL(cdsCompra, stSelect, true);
			while(miReader.Read())
			{
				miXML.WriteStartElement("detalleAnulados");
				string stCompTipo = "7";
				string stNumAnu = miReader.GetString(0);
				string stSerie = miReader.GetString(1);
				string stAutoriza = miReader.GetString(2);
				DateTime dtFechaCad = miReader.GetDateTime(3);
				DateTime dtFechaAnu = miReader.GetDateTime(4);
				int IRetencion = miReader.GetInt32(5);

				miXML.WriteElementString("tipoComprobante", int.Parse(stCompTipo).ToString());

				#region Serie
				string stEstablecmiento = "";
				string stPuntoEmision = "";
				string stSecuencial = "";
				stNumAnu = Digito(stNumAnu);
				stSerie = Digito(stSerie);
				if (stNumAnu.Length < 1)
				{
					MessageBox.Show("Numero de Retencion Anulado " + stNumAnu + " Erroneo", "Información");
					miReader.Close();
					miXML.Close();
//					new EgresoAT(IRetencion.ToString()).Show();
					Cursor = Cursors.Default;
					return;
				}

				if (stSerie.Length != 6)
				{
					MessageBox.Show("Numero de Serie de Retencion Anulado " + stNumAnu + " Erroneo", "Información");
					miReader.Close();
					miXML.Close();
//					new EgresoAT(IRetencion.ToString()).Show();
					Cursor = Cursors.Default;
					return;
				}
				stEstablecmiento = stSerie.Substring(0, 3);
				stPuntoEmision = stSerie.Substring(3, 3);
				stNumAnu= Digito(stNumAnu);
				stSecuencial = stNumAnu;
				if (int.Parse(stSecuencial) == 0)
				{
					MessageBox.Show("Numero de Secuencial de Retencion Anulado " + stSecuencial + " Erroneo", "Información");
					miReader.Close();
					miXML.Close();
//					new EgresoAT(IRetencion.ToString()).Show();
					Cursor = Cursors.Default;
					return;
				}

				miXML.WriteElementString("establecimiento", stEstablecmiento);
				miXML.WriteElementString("puntoEmision", stPuntoEmision);
				miXML.WriteElementString("secuencialInicio", stSecuencial);
				miXML.WriteElementString("secuencialFin", stSecuencial);
				#endregion Serie
		
				stAutoriza = Digito(stAutoriza);
				if (stAutoriza.Length != 10)
				{
					MessageBox.Show("Numero de Autorizacion de Retencion Anulado " + stNumAnu + " Erroneo");
					miXML.Close();
					miReader.Close();
//					new EgresoAT(IRetencion.ToString()).Show();
					Cursor = Cursors.Default;
					return;
				}
				miXML.WriteElementString("autorizacion", stAutoriza);
//				miXML.WriteElementString("fechaAnulacion", dtFechaAnu.ToString("dd/MM/yyyy"));
				miXML.WriteEndElement(); // Fin de Detalle de anulados
			}	
			miReader.Close();

			#endregion Retenciones de Compras
			
			miXML.WriteEndElement(); // Fin de anulados
			#endregion anulados

			//Write the XML to file and close the myXmlTextWriter
			miXML.WriteEndElement(); // Fin de IVA
			miXML.Flush();
			miXML.Close();

			#region Tiempo Venta
				ts = DateTime.Now.Subtract(dtInicio);
			stTiempos += string.Format("\nVentas: {0}:{1} mm:ss", ts.Minutes.ToString("n0"), ts.Seconds.ToString("n0"));
			dtInicio = DateTime.Now;
				#endregion Tiempo Compras

			if (!stArchivo.StartsWith(@"C:\XML_ANEXO_TRANSACCIONAL\")) stArchivo = Application.StartupPath + "\\" + stArchivo;
			Cursor = Cursors.Default;
			MessageBox.Show("Archivo Generado: " + stArchivo + stTiempos, "Información", MessageBoxButtons.OK, MessageBoxIcon.Information);
		}


		private void btGenerar_Click(object sender, System.EventArgs e)
		{			
			DateTime dtFecha = (DateTime) cmbFecha.Value;
			DateTime dtHasta = (DateTime) cmbHasta.Value;
			if (drSeteo.Anexo2008)
			{
				//Anexo2008();
				ATS();
				return;
			}
			else
			{
				if (DialogResult.No == MessageBox.Show("¿Desea Generar REOC?", "Confirmación",
					MessageBoxButtons.YesNo, MessageBoxIcon.Question)) return;
				if ((bool) drSeteo.StoreProcedure == false)
					AnexoReoc2008();
				else
          AnexoReocSql2008();
				return;
			}
//			MessageBox.Show("Archivo Generado", "Información");
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
//		CompraRow_tableView drCompra;
		private void RetencionRentaV2008(XmlTextWriter miXML, string stRuc, int iPeriodo, string stComprobante)
		{
			string stFechaRet = "Retencion.Fecha";
			if (drSeteo.At_Fecha_Cont == true) stFechaRet = "Retencion.FechaContable";

			DateTime dtFecha = (DateTime) cmbFecha.Value;
			DateTime dtHasta = (DateTime) cmbHasta.Value;
			string stFecha = "And Compra.Fecha >= '" + dtFecha.ToString("yyyyMMdd") + "' And Compra.Fecha < '" + dtHasta.AddDays(1).ToString("yyyyMMdd");
			string stRet = "";
			if (iPeriodo == 1)
			{
				stRet = "Select Count(*) "
					+ "FROM Compra LEFT JOIN Retencion ON Compra.idCompra = Retencion.idCompra "
					+ "Left Join CompraComprobante cc On Compra.idComprobante = cc.idComprobante "
					+ "Left JOIN RetDetalle ON Retencion.idRetencion = RetDetalle.idRetencion "
					+ "Left Join ArticuloSRI On ArticuloSRI.idArticuloSri = RetDetalle.idRetRenta "
					+ "Inner Join Cliente On Cliente.idCliente = Compra.idCliente "
					+ "Where Compra.idTipoFactura = 1 And Compra.Borrar = 0 "
					+ "And " + stFechaRet + " >= '" + dtFecha.ToString("yyyyMMdd") 
					+ "' And " + stFechaRet + " < '" + dtHasta.AddDays(1).ToString("yyyyMMdd") + "' "
					+ " And Compra.Fecha < '" + dtFecha.ToString("yyyyMMdd") + "' And Cliente.Ruc = '" + stRuc
					+ "' And cc.Codigo = '" + stComprobante + "'";
				int iCuenta = Funcion.iEscalarSQL(cdsCompra, stRet, false);
				if (iCuenta > 0) stFecha = "And Compra.Fecha < '" + dtHasta.AddDays(1).ToString("yyyyMMdd");
			}
			if (iPeriodo == 2) // Compras Anteriores
				stFecha = "And Compra.Fecha < '" + dtFecha.ToString("yyyyMMdd") ;

			//miXML.WriteStartElement("air");
			#region Creacion de Sql de retencion renta
			string stSelect = "SELECT IsNull(Sum(RetDetalle.Base), 0), "
				+ "IsNull(Sum(RetDetalle.Retenido), 0), "
				+ "Max(IsNull(ArticuloSRI.Codigo, '')), "
				+ "MAX(IsNull(ArticuloSRI.Porcentaje, 0)), "
				+ "Max(IsNull(Compra.CuentaRetDet, 0)) "
				+ "FROM Compra LEFT JOIN Retencion ON Compra.idCompra = Retencion.idCompra "
				+ "Left JOIN RetDetalle ON Retencion.idRetencion = RetDetalle.idRetencion "
				+ "Left Join ArticuloSRI On ArticuloSRI.idArticuloSri = RetDetalle.idRetRenta "
				+ "Inner Join Cliente On Cliente.idCliente = Compra.idCliente "
				+ "Inner Join CompraComprobante On CompraComprobante.idComprobante = Compra.idComprobante "
				+ "Where Compra.idTipoFactura = 1 And Compra.Borrar = 0 "
				+ "And " + stFechaRet + " >= '" + dtFecha.ToString("yyyyMMdd") 
				+ "' And " + stFechaRet + " < '" + dtHasta.AddDays(1).ToString("yyyyMMdd") + "' "
				+ stFecha + "' And Cliente.Ruc = '" + stRuc
				+ "' And Retencion.IngresoEgreso In (1, 2) "
				+ " And CompraComprobante.Codigo = " + stComprobante
				+ " Group By ArticuloSri.Codigo ";
			if (stComprobante == "04" || stComprobante == "05") stSelect = "Select * From Compra Where 0 = 1 ";
			SqlDataReader miReader = Funcion.rEscalarSQL(cdsCompra, stSelect, true);
			#endregion Creacion de Sql de retencion renta
			double dRetRenta2008 = 0;
			miReader.Read();
			if (miReader.HasRows)
			{
				while (miReader.Read())
				{
					double dRentaBase = miReader.GetDouble(0);
					double dRentaRet = miReader.GetDouble(1);
					string stCodRet = miReader.GetString(2);
					double dRentaPorc = miReader.GetDouble(3);

					if (dRentaBase > 0)
					{
						//					miXML.WriteStartElement("detalleAir");
						//					miXML.WriteElementString("codRetAir", stCodRet);
						//					miXML.WriteElementString("baseImpAir", dRentaBase.ToString("0.00", us));
						//					miXML.WriteElementString("porcentajeAir", dRentaPorc.ToString("0.00", us));
						//					miXML.WriteElementString("valRetAir", dRentaRet.ToString("0.00", us));
						//					miXML.WriteEndElement(); // Fin de Detalle de Retenciones
						dRetRenta2008 += dRentaRet;
					}
				}
			}
			miXML.WriteElementString("valorRetRenta", dRetRenta2008.ToString("0.00", us));
			miReader.Close();
			miReader = null;
//			miXML.WriteEndElement(); // Fin de Retenciones
		}


		private void RetencionRentaV(XmlTextWriter miXML, string stRuc, int iPeriodo, string stComprobante)
		{
			string stFechaRet = "Retencion.Fecha";
			if (drSeteo.At_Fecha_Cont == true)
				stFechaRet = "Retencion.FechaContable";

			DateTime dtFecha = (DateTime) cmbFecha.Value;
			DateTime dtHasta = (DateTime) cmbHasta.Value;
			string stFecha = "And Compra.Fecha >= '" + dtFecha.ToString("yyyyMMdd") 
				+ "' And Compra.Fecha < '" + dtHasta.AddDays(1).ToString("yyyyMMdd");
			string stRet = "";
			if (iPeriodo == 1)
			{
				stRet = "Select Count(*) "
				+ "FROM Compra LEFT JOIN Retencion ON Compra.idCompra = Retencion.idCompra "
				+ "Left JOIN RetDetalle ON Retencion.idRetencion = RetDetalle.idRetencion "
				+ "Left Join ArticuloSRI On ArticuloSRI.idArticuloSri = RetDetalle.idRetRenta "
				+ "Inner Join Cliente On Cliente.idCliente = Compra.idCliente "
				+ "Where Compra.idTipoFactura = 1 And Compra.Borrar = 0 "
				+ "And " + stFechaRet + " >= '" + dtFecha.ToString("yyyyMMdd") 
				+ "' And " + stFechaRet + " < '" + dtHasta.AddDays(1).ToString("yyyyMMdd") + "' "
        + " And Compra.Fecha < '" + dtFecha.ToString("yyyyMMdd") + "' And Cliente.Ruc = '" + stRuc
				+ "'";
				int iCuenta = Funcion.iEscalarSQL(cdsCompra, stRet, false);
				if (iCuenta > 0) 
					stFecha = "And Compra.Fecha < '" + dtHasta.AddDays(1).ToString("yyyyMMdd");
			}
			if (iPeriodo == 2) // Compras Anteriores
				stFecha = "And Compra.Fecha < '" + dtFecha.ToString("yyyyMMdd") ;

			miXML.WriteStartElement("air");
			#region Creacion de Sql de retencion renta
			string stSelect = "SELECT IsNull(Sum(RetDetalle.Base), 0), "
				+ "IsNull(Sum(RetDetalle.Retenido), 0), "
				+ "Max(IsNull(ArticuloSRI.Codigo, '')), "
				+ "MAX(IsNull(ArticuloSRI.Porcentaje, 0)), "
				+ "Max(IsNull(Compra.CuentaRetDet, 0)) "
				+ "FROM Compra LEFT JOIN Retencion ON Compra.idCompra = Retencion.idCompra "
				+ "Left JOIN RetDetalle ON Retencion.idRetencion = RetDetalle.idRetencion "
				+ "Left Join ArticuloSRI On ArticuloSRI.idArticuloSri = RetDetalle.idRetRenta "
				+ "Inner Join Cliente On Cliente.idCliente = Compra.idCliente "
				+ "Where Compra.idTipoFactura = 1 And Compra.Borrar = 0 "
				+ "And " + stFechaRet + " >= '" + dtFecha.ToString("yyyyMMdd") 
				+ "' And " + stFechaRet + " < '" + dtHasta.AddDays(1).ToString("yyyyMMdd") + "' "
        + stFecha + "' And Cliente.Ruc = '" + stRuc
				+ "' And Retencion.IngresoEgreso In (1,2) "
				+ " Group By ArticuloSri.Codigo ";
			if (stComprobante == "04" || stComprobante == "05")
				stSelect = "Select * From Compra Where 0 = 1 ";
			SqlDataReader miReader = Funcion.rEscalarSQL(cdsCompra, stSelect, true);
			#endregion Creacion de Sql de retencion renta
			while (miReader.Read())
			{
				double dRentaBase = miReader.GetDouble(0);
				double dRentaRet = miReader.GetDouble(1);
				string stCodRet = miReader.GetString(2);
				double dRentaPorc = miReader.GetDouble(3);

				if (dRentaBase > 0)
				{
					miXML.WriteStartElement("detalleAir");
					miXML.WriteElementString("codRetAir", stCodRet);
					miXML.WriteElementString("baseImpAir", dRentaBase.ToString("0.00", us));
					miXML.WriteElementString("porcentajeAir", dRentaPorc.ToString("0.00", us));
					miXML.WriteElementString("valRetAir", dRentaRet.ToString("0.00", us));
					miXML.WriteEndElement(); // Fin de Detalle de Retenciones
				}
			}
			miReader.Close();
			miReader = null;
			miXML.WriteEndElement(); // Fin de Retenciones
		}

		private bool ExisteAntes(string stRuc)
		{
			bool bExiste = false;
			DateTime dtFecha = (DateTime) cmbFecha.Value;
			DateTime dtHasta = (DateTime) cmbHasta.Value;
			string stFecha = "And Compra.Fecha >= '" + dtFecha.ToString("yyyyMMdd") 
				+ "' And Compra.Fecha < '" + dtHasta.AddDays(1).ToString("yyyyMMdd");
			string stRet = "Select Count(*) "
				+ "FROM Compra LEFT JOIN Retencion ON Compra.idCompra = Retencion.idCompra "
				+ "Left JOIN RetDetalle ON Retencion.idRetencion = RetDetalle.idRetencion "
				+ "Left Join ArticuloSRI On ArticuloSRI.idArticuloSri = RetDetalle.idRetRenta "
				+ "Inner Join Cliente On Cliente.idCliente = Compra.idCliente "
				+ "Where Compra.idTipoFactura = 1 And Compra.Borrar = 0 "
	//			+ "And Retencion.Fecha >= '" + dtFecha.ToString("yyyyMMdd") 
	//			+ "' And Retencion.Fecha < '" + dtHasta.AddDays(1).ToString("yyyyMMdd") + "' "
				+ " And Compra.Fecha >= '" + dtFecha.ToString("yyyyMMdd") 
				+ "' And Compra.Fecha < '" + dtHasta.AddDays(1).ToString("yyyyMMdd") 
				+ "' And Cliente.Ruc = '" + stRuc	+ "'";
			int iCuenta = Funcion.iEscalarSQL(cdsCompra, stRet, false);
			if (iCuenta > 0) bExiste = true;
			return bExiste;
		}

		private void AnexoTransaccional_Load(object sender, System.EventArgs e)
		{
			C1DataRow dr = this.cdsSeteo.TableViews[0].Rows[0];
			drSeteo = SeteoRow_tableView.Obj(dr);
//			try
//			{
//				C1DataRow drC = this.cdsCompra.TableViews[0].Rows[0];
//				drCompra = CompraRow_tableView.Obj(drC);
//			}
//			catch(Exception ex)
//			{
//				MessageBox.Show(ex.Message, "Error en compra");
//			}
			string stAudita = "Exec AuditaCrear 53, 6, 'Anexo Trans'";
			Funcion.EjecutaSQL(cdsSeteo, stAudita, true);

			this.cmbFecha.Value = "01/" + DateTime.Now.Month.ToString() + "/" + DateTime.Now.Year.ToString();
			this.cmbHasta.Value = DateTime.DaysInMonth(DateTime.Now.Year, DateTime.Now.Month).ToString() + "/" + DateTime.Now.Month.ToString() + "/" + DateTime.Now.Year.ToString();
		}

		private void RetencionIvaV(double dIvaServicio, string RucCliente, XmlTextWriter miXML, int iPeriodo, string stComprobante)
		{
			string stFechaRet = "Retencion.Fecha";
			if (drSeteo.At_Fecha_Cont == true)
				stFechaRet = "Retencion.FechaContable";

			DateTime dtFecha = (DateTime) cmbFecha.Value;
			DateTime dtHasta = (DateTime) cmbHasta.Value;
			string stFiltroCompra = "' And Compra.Fecha >= '" + dtFecha.ToString("yyyyMMdd") 
				+ "' And Compra.Fecha < '" + dtHasta.AddDays(1).ToString("yyyyMMdd");
			string stRet = "";
			if (iPeriodo == 1)
			{
				stRet = "Select Count(*) "
					+ "From Retencion inner Join RetencionIva On Retencion.idRetIva = RetencionIva.idRetIva "
					+ "Inner Join Compra On Compra.idCompra = Retencion.idCompra "
					+ "Inner Join Cliente On Compra.idCliente = Cliente.idCliente "
					+ "Where Retencion.IngresoEgreso = 1 And Retencion.Borrar = 0 And Compra.Borrar = 0 "
					+ "And " + stFechaRet + " >= '" + dtFecha.ToString("yyyyMMdd") 
					+ "' And " + stFechaRet + " < '" + dtHasta.AddDays(1).ToString("yyyyMMdd")
					+ "' And Compra.Fecha < '" + dtFecha.ToString("yyyyMMdd") + "' And Cliente.Ruc = '" + RucCliente.Trim()
					+ "' And RetencionIva.Porcentaje > 0 ";
				int iCuenta = Funcion.iEscalarSQL(cdsCompra, stRet, false);
				if (iCuenta > 0) 
					stFiltroCompra = "' And Compra.Fecha < '" + dtHasta.AddDays(1).ToString("yyyyMMdd");
			}
			if (iPeriodo == 2)
			{
				stFiltroCompra = "' And Compra.Fecha < '" + dtFecha.ToString("yyyyMMdd");
			}

			if (stComprobante == "04" || stComprobante == "05")
				stFiltroCompra = "' And '0'='1";
			double[] dValor = new double[] {0, 0, 0, 0, 0, 0};
			#region Select de retencion 1
			stRet = "Select Max(RetencionIva.Porcentaje), Sum(Retencion.IvaRetenido) "
				+ "From Retencion inner Join RetencionIva On Retencion.idRetIva = RetencionIva.idRetIva "
				+ "Inner Join Compra On Compra.idCompra = Retencion.idCompra "
				+ "Inner Join Cliente On Compra.idCliente = Cliente.idCliente "
				+ "Where Retencion.IngresoEgreso = 1 And Retencion.Borrar = 0 And Compra.Borrar = 0 "
				+ "And " + stFechaRet + " >= '" + dtFecha.ToString("yyyyMMdd") 
				+ "' And " + stFechaRet + " < '" + dtHasta.AddDays(1).ToString("yyyyMMdd")
				+ stFiltroCompra + "' And Cliente.Ruc = '" + RucCliente.Trim()
				+ "' And RetencionIva.Porcentaje > 0 Group by RetencionIva.Porcentaje Order by RetencionIva.Porcentaje";
			SqlDataReader miReader = Funcion.rEscalarSQL(cdsCompra, stRet, true);
			#endregion Select de retencion 1
			#region Valores Iniciales
			bool bExisteServ = false;
			if (dIvaServicio >= 0.01) bExisteServ = true;
			#endregion Valores Iniciales
			#region Calculo de Valores
			while (miReader.Read())
			{
				double dPorcentaje = Math.Round(miReader.GetDouble(0), 0);
				double dRetenido = Math.Round(miReader.GetDouble(1), 2);
				if (dPorcentaje == 30.0) 
				{
					dValor[0] = dRetenido / .3;
					dValor[1] = 1;
					dValor[2] = dRetenido;
				}
				if (dPorcentaje == 70.0) 
				{
					dValor[3] = dRetenido / .7;
					dValor[4] = 2;
					dValor[5] = dRetenido;
				}
				if (dPorcentaje == 100.0) 
				{
					//Existe servicio o ya hay retencion del 30%
					if (bExisteServ || dValor[0] > 0)
					{
						if (dValor[4] == 0) dValor[4] = 3;
						else dValor[4] = 4;
						dValor[3] += dRetenido;
						dValor[5] += dRetenido;
					}
					else
					{
						dValor[3] = dRetenido;
						dValor[4] = 3;
						dValor[5] = dRetenido;
					}
				}
			}
			miReader.Close();
			#endregion Calculo de Valores
			#region Select de Retencion 2
			stRet = "Select Max(RetencionIva.Porcentaje), Sum(Retencion.IvaRetenido1) "
				+ "From Retencion inner Join RetencionIva On Retencion.idRetIva1 = RetencionIva.idRetIva "
				+ "Inner Join Compra On Compra.idCompra = Retencion.idCompra "
				+ "Inner Join Cliente On Compra.idCliente = Cliente.idCliente "
				+ "Where Retencion.IngresoEgreso = 1 And Retencion.Borrar = 0 And Compra.Borrar = 0 "
				+ "And " + stFechaRet + " >= '" + dtFecha.ToString("yyyyMMdd") 
				+ "' And " + stFechaRet + " < '" + dtHasta.AddDays(1).ToString("yyyyMMdd")
				+ stFiltroCompra + "' And Cliente.Ruc = '" + RucCliente.Trim()
				+ "' And RetencionIva.Porcentaje > 0 Group by RetencionIva.Porcentaje Order by RetencionIva.Porcentaje";
			miReader = Funcion.rEscalarSQL(cdsCompra, stRet, true);
			#endregion Select de Retencion 2
			#region Calculo de Valores
			while (miReader.Read())
			{
				double dPorcentaje = Math.Round(miReader.GetDouble(0), 0);
				double dRetenido = Math.Round(miReader.GetDouble(1), 2);
				if (dPorcentaje == 30.0) 
				{
					dValor[0] += dRetenido / .3;
					dValor[1] = 1;
					dValor[2] += dRetenido;
				}
				if (dPorcentaje == 70.0) 
				{
					dValor[3] += dRetenido / .7;
					dValor[4] = 2;
					dValor[5] += dRetenido;
				}
				if (dPorcentaje == 100.0) 
				{
					//Existe servicio o ya hay retencion del 30%
					if (bExisteServ || dValor[0] > 0)
					{
						if (dValor[4] == 0) dValor[4] = 3;
						else dValor[4] = 4;
						dValor[3] += dRetenido;
						dValor[5] += dRetenido;
					}
					else
					{
						dValor[3] += dRetenido;
						dValor[4] = 3;
						dValor[5] += dRetenido;
					}
				}
			}
			miReader.Close();
			#endregion Calculo de Valores
			miXML.WriteElementString("montoIvaBienes", dValor[0].ToString("0.00", us));
			miXML.WriteElementString("porRetBienes", dValor[1].ToString("0", us));
			miXML.WriteElementString("valorRetBienes", dValor[2].ToString("0.00", us));
			miXML.WriteElementString("montoIvaServicios", dValor[3].ToString("0.00", us));
			miXML.WriteElementString("porRetServicios", dValor[4].ToString("0", us));
			miXML.WriteElementString("valorRetServicios", dValor[5].ToString("0.00", us));
			miXML.WriteElementString("retPresuntiva", "N");
		}


		private void RetencionIvaV2008(double dIvaServicio, string RucCliente, XmlTextWriter miXML, int iPeriodo, string stComprobante)
		{
			string stFechaRet = "Retencion.Fecha";
			if (drSeteo.At_Fecha_Cont == true) stFechaRet = "Retencion.FechaContable";

			DateTime dtFecha = (DateTime) cmbFecha.Value;
			DateTime dtHasta = (DateTime) cmbHasta.Value;
			string stFiltroCompra = "' And Compra.Fecha >= '" + dtFecha.ToString("yyyyMMdd") + "' And Compra.Fecha < '" + dtHasta.AddDays(1).ToString("yyyyMMdd");
			string stRet = "";
			if (iPeriodo == 1)
			{
				stRet = "Select Count(*) "
					+ "From Retencion inner Join RetencionIva On Retencion.idRetIva = RetencionIva.idRetIva "
					+ "Inner Join Compra On Compra.idCompra = Retencion.idCompra "
					+ "Left Join CompraComprobante cc On Compra.idComprobante = cc.idComprobante "
					+ "Inner Join Cliente On Compra.idCliente = Cliente.idCliente "
					+ "Where Retencion.IngresoEgreso = 1 And Retencion.Borrar = 0 And Compra.Borrar = 0 "
					+ "And " + stFechaRet + " >= '" + dtFecha.ToString("yyyyMMdd") 
					+ "' And " + stFechaRet + " < '" + dtHasta.AddDays(1).ToString("yyyyMMdd")
					+ "' And Compra.Fecha < '" + dtFecha.ToString("yyyyMMdd") + "' And Cliente.Ruc = '" + RucCliente.Trim()
					+ "' And RetencionIva.Porcentaje > 0 And cc.Codigo = '" + stComprobante + "' ";
				int iCuenta = Funcion.iEscalarSQL(cdsCompra, stRet, false);
				if (iCuenta > 0) 
					stFiltroCompra = "' And Compra.Fecha < '" + dtHasta.AddDays(1).ToString("yyyyMMdd");
			}
			if (iPeriodo == 2)
			{
				stFiltroCompra = "' And Compra.Fecha < '" + dtFecha.ToString("yyyyMMdd");
			}

			if (stComprobante == "04" || stComprobante == "05")
				stFiltroCompra = "' And '0'='1";
			double[] dValor = new double[] {0, 0, 0, 0, 0, 0};
			#region Select de retencion 1
			stRet = "Select Max(RetencionIva.Porcentaje), Sum(Retencion.IvaRetenido) "
				+ "From Retencion inner Join RetencionIva On Retencion.idRetIva = RetencionIva.idRetIva "
				+ "Inner Join Compra On Compra.idCompra = Retencion.idCompra "
				+ "Left Join CompraComprobante cc On Compra.idComprobante = cc.idComprobante "
				+ "Inner Join Cliente On Compra.idCliente = Cliente.idCliente "
				+ "Where Retencion.IngresoEgreso = 1 And Retencion.Borrar = 0 And Compra.Borrar = 0 "
				+ "And " + stFechaRet + " >= '" + dtFecha.ToString("yyyyMMdd") 
				+ "' And " + stFechaRet + " < '" + dtHasta.AddDays(1).ToString("yyyyMMdd")
				+ stFiltroCompra + "' And Cliente.Ruc = '" + RucCliente.Trim()
				+ "' And cc.Codigo = '" + stComprobante 
				+ "' And RetencionIva.Porcentaje > 0 Group by RetencionIva.Porcentaje Order by RetencionIva.Porcentaje";
			SqlDataReader miReader = Funcion.rEscalarSQL(cdsCompra, stRet, true);
			#endregion Select de retencion 1
			#region Valores Iniciales
			bool bExisteServ = false;
			if (dIvaServicio >= 0.01) bExisteServ = true;
			#endregion Valores Iniciales
			#region Calculo de Valores
			while (miReader.Read())
			{
				double dPorcentaje = Math.Round(miReader.GetDouble(0), 0);
				double dRetenido = Math.Round(miReader.GetDouble(1), 2);
				if (dPorcentaje == 30.0) 
				{
					dValor[0] = dRetenido / .3;
					dValor[1] = 1;
					dValor[2] = dRetenido;
				}
				if (dPorcentaje == 70.0) 
				{
					dValor[3] = dRetenido / .7;
					dValor[4] = 2;
					dValor[5] = dRetenido;
				}
				if (dPorcentaje == 100.0) 
				{
					//Existe servicio o ya hay retencion del 30%
					if (bExisteServ || dValor[0] > 0)
					{
						if (dValor[4] == 0) dValor[4] = 3;
						else dValor[4] = 4;
						dValor[3] += dRetenido;
						dValor[5] += dRetenido;
					}
					else
					{
						dValor[3] = dRetenido;
						dValor[4] = 3;
						dValor[5] = dRetenido;
					}
				}
			}
			miReader.Close();
			#endregion Calculo de Valores
			#region Select de Retencion 2
			stRet = "Select Max(RetencionIva.Porcentaje), Sum(Retencion.IvaRetenido1) "
				+ "From Retencion inner Join RetencionIva On Retencion.idRetIva1 = RetencionIva.idRetIva "
				+ "Inner Join Compra On Compra.idCompra = Retencion.idCompra "
				+ "Inner Join Cliente On Compra.idCliente = Cliente.idCliente "
				+ "Where Retencion.IngresoEgreso = 1 And Retencion.Borrar = 0 And Compra.Borrar = 0 "
				+ "And " + stFechaRet + " >= '" + dtFecha.ToString("yyyyMMdd") 
				+ "' And " + stFechaRet + " < '" + dtHasta.AddDays(1).ToString("yyyyMMdd")
				+ stFiltroCompra + "' And Cliente.Ruc = '" + RucCliente.Trim()
				+ "' And RetencionIva.Porcentaje > 0 Group by RetencionIva.Porcentaje Order by RetencionIva.Porcentaje";
			miReader = Funcion.rEscalarSQL(cdsCompra, stRet, true);
			#endregion Select de Retencion 2
			#region Calculo de Valores
			while (miReader.Read())
			{
				double dPorcentaje = Math.Round(miReader.GetDouble(0), 0);
				double dRetenido = Math.Round(miReader.GetDouble(1), 2);
				if (dPorcentaje == 30.0) 
				{
					dValor[0] += dRetenido / .3;
					dValor[1] = 1;
					dValor[2] += dRetenido;
				}
				if (dPorcentaje == 70.0) 
				{
					dValor[3] += dRetenido / .7;
					dValor[4] = 2;
					dValor[5] += dRetenido;
				}
				if (dPorcentaje == 100.0) 
				{
					//Existe servicio o ya hay retencion del 30%
					if (bExisteServ || dValor[0] > 0)
					{
						if (dValor[4] == 0) dValor[4] = 3;
						else dValor[4] = 4;
						dValor[3] += dRetenido;
						dValor[5] += dRetenido;
					}
					else
					{
						dValor[3] += dRetenido;
						dValor[4] = 3;
						dValor[5] += dRetenido;
					}
				}
			}
			miReader.Close();
			#endregion Calculo de Valores
//			miXML.WriteElementString("montoIvaBienes", dValor[0].ToString("0.00", us));
//			miXML.WriteElementString("porRetBienes", dValor[1].ToString("0", us));
//			miXML.WriteElementString("valorRetBienes", dValor[2].ToString("0.00", us));
//			miXML.WriteElementString("montoIvaServicios", dValor[3].ToString("0.00", us));
//			miXML.WriteElementString("porRetServicios", dValor[4].ToString("0", us));
//			miXML.WriteElementString("valorRetServicios", dValor[5].ToString("0.00", us));
			double dValorRetIva = dValor[2] + dValor[5];
			miXML.WriteElementString("valorRetIva", dValorRetIva.ToString("0.00", us));
			//			miXML.WriteElementString("retPresuntiva", "N");
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
			cdsCompra.Schema.Connections[0].ConnectionString = 
				miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);

//			C1.Data.SchemaObjects.DataSetDef dataSetDef = e.DataSet.Schema.DataSetDefs["dsCompra"];
//
//			e.Filter.Add(new C1.Data.FilterCondition(dataSetDef.TableViews["Compra"],
//				"idCompra < 10"));
		}

		private XmlValidatingReader myXmlValidatingReader = null;
		private Boolean Success = true;

		private void btValidar_Click(object sender, System.EventArgs e)
		{
			Close();
			Dispose();
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
				Console.WriteLine ("Excepcin XmlException: " + e.ToString());
			}

			catch (Exception e)
			{
				Console.WriteLine ("Excepcin: " + e.ToString());
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

		private void btRepetido_Click(object sender, System.EventArgs e)
		{
			#region Select
			DateTime dtFecha = (DateTime) cmbFecha.Value;
			DateTime dtHasta = (DateTime) cmbHasta.Value;
			string stSelect = "SELECT IsNull(Compra.Numero, ''), Compra.idCompra, "
				+ "IsNull(Compra.idCliente, 0), IsNull(Compra.idTipoFactura, 0) "
				+ "FROM Compra Inner Join Cliente On Compra.idCliente = Cliente.idCliente "
				+ "Inner Join CompraComprobante On CompraComprobante.idComprobante = "
				+ "Compra.idComprobante Where Compra.Borrar = 0 And Compra.Fecha >= '"
				+ dtFecha.ToString("yyyyMMdd") 
				+ "' And Compra.Fecha < '" + dtHasta.AddDays(1).ToString("yyyyMMdd")
				+ "' And Compra.idTipoFactura In (1, 4, 5, 6) "
				+ " AND NOT CompraComprobante.Codigo IN ('99') ";
			SqlDataReader miReader = Funcion.rEscalarSQL(cdsCompra, stSelect, true);
			#endregion Select
			#region Lectura de Variables
			long [] iNumero = new long[10000];
			int [] IdCompra = new int[10000];
			int [] IdCliente = new int[10000];
			int [] IdTipoFac = new int[10000];
			int iCuenta = 0;
			while (miReader.Read())
			{
				string stNumero = miReader.GetString(0);
				iNumero[iCuenta] = long.Parse(Digito(stNumero).ToString());
				IdCompra[iCuenta] = miReader.GetInt32(1);
				IdCliente[iCuenta] = miReader.GetInt32(2);
				IdTipoFac[iCuenta] = miReader.GetInt32(3);
				iCuenta ++;
			}
			miReader.Close();
			#endregion Lectura de Variables
			#region Recorrido de Facturas
			for (int i=0; i< iCuenta; i++)
			{
				string stBusca = "Select Numero From Compra Where Compra.idCliente = "
					+ IdCliente[i].ToString() + " And Compra.idTipoFactura = "
					+ IdTipoFac[i].ToString() + " And Compra.Borrar = 0 And Compra.Numero Like '%"
					+ iNumero[i].ToString().Trim() + "%' And IdCompra <> " + IdCompra[i].ToString();

				miReader = Funcion.rEscalarSQL(cdsCompra, stBusca, true);
				while(miReader.Read())
				{
					string stNumero = miReader.GetString(0);
					long iNumero1 = long.Parse(Digito(stNumero).ToString());
					if (iNumero1 == iNumero[i])
					{
						MessageBox.Show("Codigo Repetido en Factura " + stNumero, "Información");
						miReader.Close();
//						new Compra(IdTipoFac[i], IdCompra[i]).Show();
						return;
					}
				}
				miReader.Close();
			}
			#endregion Recorrido de Facturas
			#region Repetidos Retenciones
			stSelect = "select Retencion.IdRetencion, Max(Retencion.idCompra), "
				+ "Max(Compra.idTipoFactura) "
				+ "From Retencion Inner join RetDetalle "
				+ "On Retencion.idRetencion = RetDetalle.idRetencion "
				+ "Inner Join Compra On Compra.idCompra = Retencion.idCompra "
				+ "Where Retencion.Fecha >= '" + dtFecha.ToString("yyyyMMdd") 
				+ "' And Retencion.Fecha < '" + dtHasta.AddDays(1).ToString("yyyyMMdd")
				+ "' Group by Retencion.idRetencion "
				+ "Having Count(RetDetalle.idRetRenta) <> Count(distinct RetDetalle.idRetRenta) ";
			miReader = Funcion.rEscalarSQL(cdsCompra, stSelect, true);

			while(miReader.Read())
			{
				MessageBox.Show("Error en Retencion", "Información");
				int IdRetencion = miReader.GetInt32(0);
				int IdCompra1 = miReader.GetInt32(1);
				int IdTipoFac1 = miReader.GetInt32(2);
//				using(Compra miCompra = new Compra(IdTipoFac1, IdCompra1))
//				{
//					miCompra.ShowDialog();
//				}
				miReader.Close();
				return;
			}
			miReader.Close();
			#endregion Repetidos Retenciones

			MessageBox.Show("Repetidos Generados", "Información");
		}

		private void ultraButton1_Click(object sender, System.EventArgs e)
		{
			AnexoTransAvanzado miAvanzado = new AnexoTransAvanzado();
			miAvanzado.MdiParent = this.MdiParent;
			miAvanzado.Show();
		}

		private void btLee_Click(object sender, System.EventArgs e)
		{
			AnexoTransImporta miAnexo = new AnexoTransImporta();
			miAnexo.MdiParent = this.MdiParent;
			miAnexo.Show();
		}

		private void btRepCliente_Click(object sender, System.EventArgs e)
		{
			#region Cliente Repetidos
			string stSelect = "Select Max(IsNull(Ruc, '')), Max(idCliente), max(cast(proveedor as int)) from Cliente "
				+ "where Ambos = 0 Group by LTRIM(RTRIM(ruc)), Proveedor having count(*) > 1 ";
			SqlDataReader miReader = Funcion.rEscalarSQL(cdsCompra, stSelect, true);
			while(miReader.Read())
			{
				string stRucC = miReader.GetString(0);
				MessageBox.Show("Cedula o Ruc Repetido: " +stRucC, "Información");
				int IdClienteR = miReader.GetInt32(1);
				int iProveedor = miReader.GetInt32(2);
				bool bProveedor = false;
				if (iProveedor == 0) bProveedor = true;
				using (Cliente miCliente = new Cliente(bProveedor, IdClienteR))
				{
					miCliente.ShowDialog();
				}
				miReader.Close();
				return;
			}
			miReader.Close();
			MessageBox.Show("Registros Repetidos Revisados",
				"Información", MessageBoxButtons.OK, MessageBoxIcon.Information);
			#endregion
		}

		private void btRepCliente_MouseDown(object sender, System.Windows.Forms.MouseEventArgs e)
		{
			if (e.Button != MouseButtons.Right) return;
			if (DialogResult.No == MessageBox.Show("Desea Reemplazar Ruc de Consumidor Final con 9999999999999",
				"Confirmación", MessageBoxButtons.YesNo, MessageBoxIcon.Question)) return;
			string stSelect = "Update Cliente Set Ruc = '9999999999999' Where idTipoRuc = 7 And Proveedor = 0";
			Funcion.EjecutaSQL(cdsCompra, stSelect, true);
			MessageBox.Show("Registros Modificados",
				"Información", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
		}

		private void AnexoTransaccional_Closing(object sender, System.ComponentModel.CancelEventArgs e)
		{
			string stAudita = "Exec AuditaCrear 53, 6, 'Anexo Trans'";
			Funcion.EjecutaSQL(cdsCompra, stAudita, true);
		}

		private void AnexoTransaccional_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Escape) this.Close();
		}

	}
}
