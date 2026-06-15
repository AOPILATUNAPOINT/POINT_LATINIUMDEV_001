using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using System.Xml;
using System.IO;
using System.Data.SqlClient;
using System.Xml.Schema;
using System.Globalization;
using LibContabilidad.DataObjects;
using C1.Data;

namespace Latinium
{
	/// <summary>
	/// Descripción breve de ReporteSRI.
	/// </summary>
	public class ReporteSRI : DevExpress.XtraEditors.XtraForm
	{
		private Infragistics.Win.Misc.UltraButton btGenerar;
		private Infragistics.Win.Misc.UltraLabel ultraLabel2;
		private Infragistics.Win.Misc.UltraLabel ultraLabel1;
		private Infragistics.Win.UltraWinEditors.UltraDateTimeEditor cmbHasta;
		private Infragistics.Win.UltraWinEditors.UltraDateTimeEditor cmbFecha;
		private C1.Data.C1DataSet cdsSeteo;
		private DevExpress.XtraEditors.RadioGroup optTipo;
		/// <summary>
		/// Variable del diseñador requerida.
		/// </summary>
		private System.ComponentModel.Container components = null;

		public ReporteSRI()
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
			this.btGenerar = new Infragistics.Win.Misc.UltraButton();
			this.ultraLabel2 = new Infragistics.Win.Misc.UltraLabel();
			this.ultraLabel1 = new Infragistics.Win.Misc.UltraLabel();
			this.cmbHasta = new Infragistics.Win.UltraWinEditors.UltraDateTimeEditor();
			this.cmbFecha = new Infragistics.Win.UltraWinEditors.UltraDateTimeEditor();
			this.cdsSeteo = new C1.Data.C1DataSet();
			this.optTipo = new DevExpress.XtraEditors.RadioGroup();
			((System.ComponentModel.ISupportInitialize)(this.cmbHasta)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbFecha)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteo)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.optTipo.Properties)).BeginInit();
			this.SuspendLayout();
			// 
			// btGenerar
			// 
			this.btGenerar.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			this.btGenerar.Location = new System.Drawing.Point(104, 248);
			this.btGenerar.Name = "btGenerar";
			this.btGenerar.Size = new System.Drawing.Size(96, 23);
			this.btGenerar.TabIndex = 1;
			this.btGenerar.Text = "Generar";
			this.btGenerar.Click += new System.EventHandler(this.btGenerar_Click);
			// 
			// ultraLabel2
			// 
			this.ultraLabel2.AutoSize = true;
			this.ultraLabel2.Location = new System.Drawing.Point(32, 48);
			this.ultraLabel2.Name = "ultraLabel2";
			this.ultraLabel2.Size = new System.Drawing.Size(37, 14);
			this.ultraLabel2.TabIndex = 14;
			this.ultraLabel2.Text = "Hasta:";
			// 
			// ultraLabel1
			// 
			this.ultraLabel1.AutoSize = true;
			this.ultraLabel1.Location = new System.Drawing.Point(32, 16);
			this.ultraLabel1.Name = "ultraLabel1";
			this.ultraLabel1.Size = new System.Drawing.Size(40, 14);
			this.ultraLabel1.TabIndex = 13;
			this.ultraLabel1.Text = "Desde:";
			// 
			// cmbHasta
			// 
			this.cmbHasta.DateTime = new System.DateTime(2009, 7, 30, 0, 0, 0, 0);
			this.cmbHasta.FormatString = "dd/MMM/yyyy";
			this.cmbHasta.Location = new System.Drawing.Point(96, 48);
			this.cmbHasta.Name = "cmbHasta";
			this.cmbHasta.PromptChar = ' ';
			this.cmbHasta.Size = new System.Drawing.Size(160, 21);
			this.cmbHasta.SpinButtonDisplayStyle = Infragistics.Win.ButtonDisplayStyle.Always;
			this.cmbHasta.TabIndex = 12;
			this.cmbHasta.Value = new System.DateTime(2009, 7, 30, 0, 0, 0, 0);
			// 
			// cmbFecha
			// 
			this.cmbFecha.DateTime = new System.DateTime(2009, 7, 30, 0, 0, 0, 0);
			this.cmbFecha.FormatString = "dd/MMM/yyyy";
			this.cmbFecha.Location = new System.Drawing.Point(96, 16);
			this.cmbFecha.Name = "cmbFecha";
			this.cmbFecha.PromptChar = ' ';
			this.cmbFecha.Size = new System.Drawing.Size(160, 21);
			this.cmbFecha.SpinButtonDisplayStyle = Infragistics.Win.ButtonDisplayStyle.Always;
			this.cmbFecha.TabIndex = 11;
			this.cmbFecha.Value = new System.DateTime(2009, 7, 30, 0, 0, 0, 0);
			// 
			// cdsSeteo
			// 
			this.cdsSeteo.BindingContextCtrl = this;
			this.cdsSeteo.DataLibrary = "LibContabilidad";
			this.cdsSeteo.DataLibraryUrl = "";
			this.cdsSeteo.DataSetDef = "dsSeteo";
			this.cdsSeteo.Locale = new System.Globalization.CultureInfo("es-MX");
			this.cdsSeteo.Name = "cdsSeteo";
			this.cdsSeteo.SchemaClassName = "LibContabilidad.DataClass";
			this.cdsSeteo.SchemaDef = null;
			// 
			// optTipo
			// 
			this.optTipo.EditValue = 0;
			this.optTipo.Location = new System.Drawing.Point(96, 88);
			this.optTipo.Name = "optTipo";
			// 
			// optTipo.Properties
			// 
			this.optTipo.Properties.Items.AddRange(new DevExpress.XtraEditors.Controls.RadioGroupItem[] {
																																																		new DevExpress.XtraEditors.Controls.RadioGroupItem(0, "Factura"),
																																																		new DevExpress.XtraEditors.Controls.RadioGroupItem(1, "Nota de Venta"),
																																																		new DevExpress.XtraEditors.Controls.RadioGroupItem(2, "Nota de Debito"),
																																																		new DevExpress.XtraEditors.Controls.RadioGroupItem(3, "Nota de Crédito"),
																																																		new DevExpress.XtraEditors.Controls.RadioGroupItem(4, "Retención"),
																																																		new DevExpress.XtraEditors.Controls.RadioGroupItem(5, "Guía de Remisión")});
			this.optTipo.Properties.Style = new DevExpress.Utils.ViewStyle("ControlStyle", null, new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0))), "", DevExpress.Utils.StyleOptions.StyleEnabled, true, false, false, DevExpress.Utils.HorzAlignment.Default, DevExpress.Utils.VertAlignment.Center, null, System.Drawing.Color.Transparent, System.Drawing.SystemColors.WindowText);
			this.optTipo.Size = new System.Drawing.Size(142, 144);
			this.optTipo.TabIndex = 64;
			// 
			// ReporteSRI
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 13);
			this.ClientSize = new System.Drawing.Size(292, 286);
			this.Controls.Add(this.optTipo);
			this.Controls.Add(this.ultraLabel2);
			this.Controls.Add(this.ultraLabel1);
			this.Controls.Add(this.cmbHasta);
			this.Controls.Add(this.cmbFecha);
			this.Controls.Add(this.btGenerar);
			this.Name = "ReporteSRI";
			this.Text = "Reportes del SRI";
			this.Load += new System.EventHandler(this.ReporteSRI_Load);
			((System.ComponentModel.ISupportInitialize)(this.cmbHasta)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbFecha)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteo)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.optTipo.Properties)).EndInit();
			this.ResumeLayout(false);

		}
		#endregion

		SeteoRow_tableView drSeteo;
		private void ReporteSRI_Load(object sender, System.EventArgs e)
		{
			C1DataRow dr = this.cdsSeteo.TableViews[0].Rows[0];
			drSeteo = SeteoRow_tableView.Obj(dr);
			cmbFecha.Value=DateTime.Today;
			cmbHasta.Value=DateTime.Today;
		}

		private void Factura(int idTipo, int idCompra, string stCodigo)
		{
			#region Crea Documento XML
			string stTipo = "factura";
			string stCodDoc = "1";
			if (idTipo == 1) 
			{
				stTipo = "notaVenta";
				stCodDoc = "2";
			}
			bool bAnulado = Funcion.bEjecutaSQL(cdsSeteo, "Select Borrar From Compra Where idCompra = "
				+ idCompra.ToString());
			if (bAnulado) stTipo = "Anulado" + stTipo;
			XmlTextWriter miXML = new XmlTextWriter("FacturasSri\\"+stTipo + stCodigo + ".Xml", null);
			miXML.Formatting = Formatting.Indented;
			miXML.WriteProcessingInstruction("xml", "version=\"1.0\" encoding=\"UTF-8\"");
			#endregion Crea Documento XML
			#region Datos de Empresa
			string stMensaje = "";
			string stVerificaRuc = Funcion.VerificaCedula(drSeteo.Ruc, false);
			if (drSeteo.Empresa.Trim().Length < 3) stMensaje += "Longitud de Nombre de su Empresa Erronea\n";
			if (stVerificaRuc.Length != 0) stMensaje += "Error en RUC de su Empresa\n";
			if (stMensaje.Length > 0)
			{
				miXML.Close();
				MessageBox.Show(stMensaje, "Información");
				Seteo miSeteo = new Seteo();
				miSeteo.MdiParent = this.MdiParent;
				miSeteo.Show();
				return;
			}
			#endregion Datos de Empresa
			#region Valores Iniciales
			string stExec = "Exec AutoImpresorFactura1 " + idTipo.ToString() + ", " + 
				idCompra.ToString();
			SqlDataReader dr = Funcion.rEscalarSQL(cdsSeteo, stExec, true);
			string stEstabPtoEmi = "";
			if(dr.Read()) stEstabPtoEmi = dr.GetString(1);
			#endregion ValoresIniciales
			while(stEstabPtoEmi.Length > 0)
			{
				#region Encabezado
				miXML.WriteStartElement(stTipo);
				miXML.WriteStartElement("infoTributaria");
				miXML.WriteElementString("razonSocial", drSeteo.Empresa.ToUpper().Trim());
				miXML.WriteElementString("ruc", drSeteo.Ruc);
				miXML.WriteElementString("numAut", dr.GetString(0));
				miXML.WriteElementString("codDoc", stCodDoc);
				stEstabPtoEmi = dr.GetString(1);
				if (stEstabPtoEmi.Length != 13)
				{
					dr.Close();
					miXML.Flush();
					miXML.Close();
					MessageBox.Show("Numero: " + stEstabPtoEmi + " debe ser de 13 digitos", "Información");
					return;
				}
				miXML.WriteElementString("estab", stEstabPtoEmi.Substring(0, 3));
				miXML.WriteElementString("ptoEmi", stEstabPtoEmi.Substring(3, 3));
				miXML.WriteElementString("secuencial", stEstabPtoEmi.Substring(6, 7));
				miXML.WriteElementString("caducidad",dr.GetDateTime(5).ToString("MM/yyyy"));
				miXML.WriteElementString("fechaEmision",dr.GetDateTime(2).ToString("dd/MM/yyyy"));
				miXML.WriteElementString("dirMatriz",drSeteo.Direccion);

				miXML.WriteElementString("razonSocialComprador", dr.GetString(3));
				miXML.WriteElementString("rucCedulaComprador",dr.GetString(4));
				#endregion Encabezado
				#region Variables de Pie
				double dTotalSinImp = dr.GetDouble(9);
				double dDescuentos = dr.GetDouble(10);
				double dIce = dr.GetDouble(11);
				double dIva0 = dr.GetDouble(12);
				double dIva12 = dr.GetDouble(13);
				double dPropina = dr.GetDouble(14);
				double dTotalConImp = dr.GetDouble(15);
				#endregion Variables de Pie
				#region Grabacion de Pie
				miXML.WriteElementString("totalSinImpuestos", dTotalSinImp.ToString("0.00", us));
				if (dIce >= 0.01)
					miXML.WriteElementString("ICE", dIce.ToString("0.00", us));
				if (dIva0 >= 0.01) // SubTotalExcento
					miXML.WriteElementString("baseIVA0", dIva0.ToString("0.00", us));
				if (dIva12 >= 0.01) // Iva
				{
					double dSubTotalIva = dTotalSinImp-dIva0;
					miXML.WriteElementString("baseIVA12", dSubTotalIva.ToString("0.00", us));
					miXML.WriteElementString("IVA12", dIva12.ToString("0.00", us));
				}
				if (dPropina >= 0.01)
					miXML.WriteElementString("propina", dPropina.ToString("0.00", us));
				miXML.WriteElementString("totalConImpuestos", dTotalConImp.ToString("0.00", us));
				#endregion Grabacion de Pie
				miXML.WriteEndElement(); // Fin de infoTributaria
				#region Detalles
				miXML.WriteStartElement("detalles");
//				miXML.WriteStartElement("detalle");
				string stCodAntes = stEstabPtoEmi;
				while(stCodAntes == stEstabPtoEmi)
				{
					miXML.WriteStartElement("detalle");
					miXML.WriteElementString("concepto", dr.GetString(6));
					double dCantidad = dr.GetDouble(7);
					double dPrecio = dr.GetDouble(8);
					double dPrecioT = dCantidad*dPrecio;
					miXML.WriteElementString("cantidad", dCantidad.ToString("0.00", us));
					miXML.WriteElementString("precioUnitario", dPrecio.ToString("0.00", us));
					miXML.WriteElementString("precioTotal", dPrecioT.ToString("0.00", us));
					miXML.WriteEndElement(); // Detalle
					if (dr.Read()) stEstabPtoEmi = dr.GetString(1);
					else stEstabPtoEmi = "";
				}
//				miXML.WriteEndElement(); // Detalle
				miXML.WriteEndElement(); // Detalles
				#endregion Detalles
				miXML.WriteEndElement(); // Fin de Factura
			}
			#region Cierre
			dr.Close();
			miXML.Flush();
			miXML.Close();
			#endregion Cierre
		}


		private void NotaVenta(int idTipo, int idCompra, string stCodigo)
		{
			#region Crea Documento XML
			string stTipo = "factura";
			string stCodDoc = "1";
			if (idTipo == 1) 
			{
				stTipo = "notaVenta";
				stCodDoc = "2";
			}
			bool bAnulado = Funcion.bEjecutaSQL(cdsSeteo, "Select Borrar From Compra Where idCompra = "
				+ idCompra.ToString());
			if (bAnulado) stTipo = "Anulado" + stTipo;
			XmlTextWriter miXML = new XmlTextWriter("FacturasSri\\"+stTipo + stCodigo + ".Xml", null);
			miXML.Formatting = Formatting.Indented;
			miXML.WriteProcessingInstruction("xml", "version=\"1.0\" encoding=\"UTF-8\"");
			#endregion Crea Documento XML
			#region Datos de Empresa
			string stMensaje = "";
			string stVerificaRuc = Funcion.VerificaCedula(drSeteo.Ruc, false);
			if (drSeteo.Empresa.Trim().Length < 3) stMensaje += "Longitud de Nombre de su Empresa Erronea\n";
			if (stVerificaRuc.Length != 0) stMensaje += "Error en RUC de su Empresa\n";
			if (stMensaje.Length > 0)
			{
				miXML.Close();
				MessageBox.Show(stMensaje, "Información");
				Seteo miSeteo = new Seteo();
				miSeteo.MdiParent = this.MdiParent;
				miSeteo.Show();
				return;
			}
			#endregion Datos de Empresa
			#region Valores Iniciales
			string stExec = "Exec AutoImpresorFactura1 " + idTipo.ToString() + ", " + 
				idCompra.ToString();
			SqlDataReader dr = Funcion.rEscalarSQL(cdsSeteo, stExec, true);
			string stEstabPtoEmi = "";
			if(dr.Read()) stEstabPtoEmi = dr.GetString(1);
			#endregion ValoresIniciales
			while(stEstabPtoEmi.Length > 0)
			{
				#region Encabezado
				miXML.WriteStartElement(stTipo);
				miXML.WriteElementString("razonSocial", "![DATA[" + drSeteo.Empresa.ToUpper().Trim() + "]]");
				miXML.WriteElementString("ruc", drSeteo.Ruc);
				miXML.WriteElementString("numAut", dr.GetString(0));
				miXML.WriteElementString("codDoc", stCodDoc);
				stEstabPtoEmi = dr.GetString(1);
				if (stEstabPtoEmi.Length != 13)
				{
					dr.Close();
					miXML.Flush();
					miXML.Close();
					MessageBox.Show("Numero: " + stEstabPtoEmi + " debe ser de 13 digitos", "Información");
					return;
				}
				miXML.WriteElementString("estab", stEstabPtoEmi.Substring(0, 3));
				miXML.WriteElementString("ptoEmi", stEstabPtoEmi.Substring(3, 3));
				miXML.WriteElementString("secuencial", stEstabPtoEmi.Substring(6, 7));
				miXML.WriteElementString("fechaEmision",dr.GetDateTime(2).ToString("dd/MM/yyyy"));

				miXML.WriteElementString("razonSocialComprador", "![DATA["+ dr.GetString(3) + "]]");
				miXML.WriteElementString("rucCedulaComprador",dr.GetString(4));
				miXML.WriteElementString("caducidad",dr.GetDateTime(5).ToString("MM/yyyy"));
				#endregion Encabezado
				#region Variables de Pie
				double dTotalSinImp = dr.GetDouble(9);
				double dDescuentos = dr.GetDouble(10);
				double dIce = dr.GetDouble(11);
				double dIva0 = dr.GetDouble(12);
				double dIva12 = dr.GetDouble(13);
				double dPropina = dr.GetDouble(14);
				double dTotalConImp = dr.GetDouble(15);
				#endregion Variables de Pie
				#region Detalles
				miXML.WriteStartElement("detalles");
				string stCodAntes = stEstabPtoEmi;
				while(stCodAntes == stEstabPtoEmi)
				{
					miXML.WriteStartElement("detalle");
					miXML.WriteElementString("concepto", "![DATA[" +dr.GetString(6)+ "]]");
					double dCantidad = dr.GetDouble(7);
					double dPrecio = dr.GetDouble(8);
					double dPrecioT = dCantidad*dPrecio;
					miXML.WriteElementString("cantidad", dCantidad.ToString("0.00", us));
					miXML.WriteElementString("precioUnitario", dPrecio.ToString("0.00", us));
					miXML.WriteEndElement(); // Detalle
					if (dr.Read()) stEstabPtoEmi = dr.GetString(1);
					else stEstabPtoEmi = "";
				}
				miXML.WriteEndElement(); // Detalles
				#endregion Detalles
				#region Grabacion de Pie
				miXML.WriteElementString("totalSinImpuestos", dTotalSinImp.ToString("0.00", us));
//				if (dIce >= 0.01)
					miXML.WriteElementString("ICE", dIce.ToString("0.00", us));
//				if (dIva0 >= 0.01) // SubTotalExcento
					miXML.WriteElementString("IVA0", dIva0.ToString("0.00", us));
//				if (dIva12 >= 0.01) // Iva
					miXML.WriteElementString("IVA12", dIva12.ToString("0.00", us));
//				if (dPropina >= 0.01)
					miXML.WriteElementString("propina10", dPropina.ToString("0.00", us));
				miXML.WriteElementString("totalConImpuestos", dTotalConImp.ToString("0.00", us));
				#endregion Grabacion de Pie

				miXML.WriteEndElement(); // Fin de Factura
			}
			#region Cierre
			dr.Close();
			miXML.Flush();
			miXML.Close();
			#endregion Cierre
		}


		private void NotaCredito(int idTipo, int idCompra, string stCodigo)
		{
			#region Crea Documento XML
			string stTipo = "notaCredito";
			bool bAnulado = Funcion.bEjecutaSQL(cdsSeteo, "Select Borrar From Compra Where idCompra = "
				+ idCompra.ToString());
			if (bAnulado) stTipo = "Anulado" + stTipo;
			XmlTextWriter miXML = new XmlTextWriter("NotaCredito\\"+stTipo + stCodigo + ".Xml", null);
			miXML.Formatting = Formatting.Indented;
			miXML.WriteProcessingInstruction("xml", "version=\"1.0\" encoding=\"UTF-8\"");
			#endregion Crea Documento XML
			#region Datos de Empresa
			string stMensaje = "";
			string stVerificaRuc = Funcion.VerificaCedula(drSeteo.Ruc, false);
			if (drSeteo.Empresa.Trim().Length < 3) stMensaje += "Longitud de Nombre de su Empresa Erronea\n";
			if (stVerificaRuc.Length != 0) stMensaje += "Error en RUC de su Empresa\n";
			if (stMensaje.Length > 0)
			{
				miXML.Close();
				MessageBox.Show(stMensaje, "Información");
				Seteo miSeteo = new Seteo();
				miSeteo.MdiParent = this.MdiParent;
				miSeteo.Show();
				return;
			}
			#endregion Datos de Empresa
			#region Valores Iniciales
			string stExec = "Exec AutoImpresorNotaDebito1 " + idTipo.ToString() + ", " + 
				idCompra.ToString();
			SqlDataReader dr = Funcion.rEscalarSQL(cdsSeteo, stExec, true);
			string stEstabPtoEmi = "";
			if(dr.Read()) stEstabPtoEmi = dr.GetString(1);
			#endregion ValoresIniciales
			while(stEstabPtoEmi.Length > 0)
			{
				#region Encabezado
				miXML.WriteStartElement(stTipo);
				miXML.WriteStartElement("infoTributaria");
				miXML.WriteElementString("razonSocial", drSeteo.Empresa.ToUpper().Trim());
				miXML.WriteElementString("ruc", drSeteo.Ruc);
				miXML.WriteElementString("numAut", dr.GetString(0));
				miXML.WriteElementString("codDoc", "4");
				stEstabPtoEmi = dr.GetString(1);
				if (stEstabPtoEmi.Length != 13)
				{
					dr.Close();
					miXML.Flush();
					miXML.Close();
					MessageBox.Show("Numero: " + stEstabPtoEmi + " debe ser de 13 digitos", "Información");
					return;
				}
				miXML.WriteElementString("estab", stEstabPtoEmi.Substring(0, 3));
				miXML.WriteElementString("ptoEmi", stEstabPtoEmi.Substring(3, 3));
				miXML.WriteElementString("secuencial", stEstabPtoEmi.Substring(6, 7));
				miXML.WriteElementString("caducidad",dr.GetDateTime(5).ToString("MM/yyyy"));
				miXML.WriteElementString("fechaEmision",dr.GetDateTime(2).ToString("dd/MM/yyyy"));
				miXML.WriteElementString("dirMatriz",drSeteo.Direccion);

				miXML.WriteElementString("razonSocialComprador", dr.GetString(3));
				miXML.WriteElementString("rucCedulaComprador",dr.GetString(4));

				miXML.WriteElementString("codDocModificado",dr.GetString(16));
				miXML.WriteElementString("numDocModificado",dr.GetString(17));
				miXML.WriteElementString("numAutDocSustento",dr.GetString(18));
				DateTime dtEmisionSustento = dr.GetDateTime(19);
				miXML.WriteElementString("fechaEmisionDocSustento", dtEmisionSustento.ToString("dd/MM/yyyy"));
				#endregion Encabezado

				#region Variables de Pie
				double dTotalSinImp = dr.GetDouble(9);
				double dDescuentos = dr.GetDouble(10);
				double dIce = dr.GetDouble(11);
				double dIva0 = dr.GetDouble(12);
				double dIva12 = dr.GetDouble(13);
				double dPropina = dr.GetDouble(14);
				double dTotalConImp = dr.GetDouble(15);
				#endregion Variables de Pie
				#region Grabacion de Pie
				miXML.WriteElementString("valorModificacion", dTotalSinImp.ToString("0.00", us));
				miXML.WriteElementString("ICE", dIce.ToString("0.00", us));
				miXML.WriteElementString("baseIVA0", dIva0.ToString("0.00", us));
				miXML.WriteElementString("baseIVA12", dIva12.ToString("0.00", us));
				miXML.WriteElementString("IVA12", dIva12.ToString("0.00", us));
				miXML.WriteElementString("valorTotal", dTotalConImp.ToString("0.00", us));

				#endregion Grabacion de Pie
				miXML.WriteEndElement(); // Fin infoTributaria 

				#region Detalles
				miXML.WriteStartElement("motivos");
				string stCodAntes = stEstabPtoEmi;
				while(stCodAntes == stEstabPtoEmi)
				{
					string stMotivo = dr.GetString(6);
					if (stMotivo.Length == 0) stMotivo="Devolucion";
					miXML.WriteElementString("motivo", stMotivo);
					if (dr.Read()) stEstabPtoEmi = dr.GetString(1);
					else stEstabPtoEmi = "";
				}
				miXML.WriteEndElement(); // motivos
				#endregion Detalles

				#region infoAdicional
//				miXML.WriteStartElement("infoAdicional");
//				string stDireccion = "Av. Colón y América";
//				string stAdicional = "<campoAdicional nombre=\"Direccion Establecimiento\">"
//					+ stDireccion + "</campoAdicional>";
//				miXML.WriteString(stAdicional);
//				stAdicional = "<campoAdicional nombre=\"Nombre Comercial\">"
//					+ stDireccion + "</campoAdicional>";
//				miXML.WriteString(stAdicional);
//				stAdicional = "<campoAdicional nombre=\"Cedula Firmante\">"
//					+ stDireccion + "</campoAdicional>";
//				stAdicional = "<campoAdicional nombre=\"Ruc Firmante\">"
//					+ stDireccion + "</campoAdicional>";
//				miXML.WriteString(stAdicional);
//				miXML.WriteEndElement(); // Fin de infoAdicional
				#endregion infoAdicional

				miXML.WriteEndElement(); // Fin de Factura

			}
			#region Cierre
			dr.Close();
			miXML.Flush();
			miXML.Close();
			#endregion Cierre
		}


		private void NotaDebito(int idTipo)
		{
			DateTime dtDesde = (DateTime) cmbFecha.Value;
			DateTime dtHasta = (DateTime) cmbHasta.Value;
			#region Datos de Empresa
			string stMensaje = "";
			string stVerificaRuc = Funcion.VerificaCedula(drSeteo.Ruc, false);
			if (drSeteo.Empresa.Trim().Length < 3) stMensaje += "Longitud de Nombre de su Empresa Erronea\n";
			if (stVerificaRuc.Length != 0) stMensaje += "Error en RUC de su Empresa\n";
			if (stMensaje.Length > 0)
			{
				MessageBox.Show(stMensaje, "Información");
				Seteo miSeteo = new Seteo();
				miSeteo.MdiParent = this.MdiParent;
				miSeteo.Show();
				return;
			}
			#endregion Datos de Empresa
			#region Valores Iniciales
			string stExec = "Exec AutoImpresorNotaDebito '" + dtDesde.ToString("yyyyMMdd") + "', '"
				+ dtHasta.AddDays(1).ToString("yyyyMMdd") + "', " + idTipo.ToString();
			SqlDataReader dr = Funcion.rEscalarSQL(cdsSeteo, stExec, true);
			#endregion ValoresIniciales
			while(dr.Read())
			{
				#region Crea Documento XML
				string stTipo = "notaDebito";
				string stCodDoc = "5";
				if (idTipo == 2) 
				{
					stTipo = "notaCredito";
					stCodDoc = "4";
				}
				if (!Directory.Exists(stTipo)) Directory.CreateDirectory(stTipo);
				string stEstabPtoEmi = dr.GetString(1);
				string stArchivo = stTipo + "\\"+ stTipo + "_" + stEstabPtoEmi + ".Xml";
				XmlTextWriter miXML = new XmlTextWriter(stArchivo, null);
				miXML.Formatting = Formatting.Indented;
				miXML.WriteProcessingInstruction("xml", "version=\"1.0\" encoding=\"UTF-8\"");
				#endregion Crea Documento XML
				#region Encabezado
				miXML.WriteStartElement(stTipo);
				miXML.WriteElementString("razonSocial", "![DATA[" + drSeteo.Empresa.ToUpper().Trim() + "]]");
				miXML.WriteElementString("ruc", drSeteo.Ruc);
				miXML.WriteElementString("numAut", dr.GetString(0));
				miXML.WriteElementString("codDoc", stCodDoc);
				stEstabPtoEmi = dr.GetString(1);
				if (stEstabPtoEmi.Length != 13)
				{
					miXML.Close();
					dr.Close();
					MessageBox.Show("Numero debe ser de 13 digitos", "Información");
					return;
				}
				miXML.WriteElementString("establ", stEstabPtoEmi.Substring(0, 3));
				miXML.WriteElementString("ptoEmi", stEstabPtoEmi.Substring(3, 3));
				miXML.WriteElementString("secuencial", stEstabPtoEmi.Substring(6, 7));
				miXML.WriteElementString("fechaEmision",dr.GetDateTime(2).ToString("dd-MM-yyyy"));
				miXML.WriteElementString("razonSocialComprador", "![DATA["+ dr.GetString(3) + "]]");
				miXML.WriteElementString("rucCedulaComprador",dr.GetString(4));
				miXML.WriteElementString("caducidad",dr.GetDateTime(5).ToString("MM/yyyy"));
				#endregion Encabezado
				#region Doc Modificacdo
				int iCodModificado = dr.GetInt32(6);
				if (iCodModificado==15) iCodModificado = 1;
				else if (iCodModificado == 2) iCodModificado=2;
				else if (iCodModificado == 3) iCodModificado=3;
				else if (iCodModificado == 4) iCodModificado=4;
				else if (iCodModificado == 5) iCodModificado=5;
				miXML.WriteElementString("codDocModificado",iCodModificado.ToString());
				miXML.WriteElementString("codDocModificado","1");
				string stNumDocMod = dr.GetString(7);
				if (stNumDocMod.Length != 13)
				{
					miXML.Close();
					dr.Close();
					MessageBox.Show("Numero de comprobante asociado de "
						+ stEstabPtoEmi +" debe ser de 13 digitos", "Información");
					return;
				}
				stNumDocMod = stNumDocMod.Substring(0, 3) + "-" + stNumDocMod.Substring(3, 3)
					+ "-" + stNumDocMod.Substring(6);
				miXML.WriteElementString("numDocModificado", stNumDocMod);
				miXML.WriteElementString("numAutDocSustento",dr.GetString(8));
				miXML.WriteElementString("fechaEmisionDocSustento",dr.GetDateTime(9).ToString("dd-MM-yyyy"));
				#endregion Doc Modificacdo

				#region Variables de Pie
				double dValorModificacion = dr.GetDouble(10);
				double dIce = dr.GetDouble(11);
				double dIva0 = dr.GetDouble(12);
				double dIva12 = dr.GetDouble(13);
				double dValorTotal = dr.GetDouble(14);
				#endregion Variables de Pie
				#region Detalles
				miXML.WriteStartElement("detalles");
//				string stCodAntes = stEstabPtoEmi;
//				while(stCodAntes == stEstabPtoEmi)
				{
					miXML.WriteStartElement("detalle");
					miXML.WriteElementString("motivoModificacion", "![CDATA[Devolucion]]");
					miXML.WriteEndElement(); // Detalle
//					if (dr.Read()) stEstabPtoEmi = dr.GetString(1);
//					else stEstabPtoEmi = "";
				}
				miXML.WriteEndElement(); // Detalles
				#endregion Detalles
				#region Grabacion de Pie
				miXML.WriteElementString("ValorModificacion", dValorModificacion.ToString("0.00", us));
				miXML.WriteElementString("ICE", dIce.ToString("0.00", us));
				miXML.WriteElementString("IVA0", dIva0.ToString("0.00", us));
				miXML.WriteElementString("IVA12", dIva12.ToString("0.00", us));
				miXML.WriteElementString("valorTotal", dValorTotal.ToString("0.00", us));
				miXML.WriteEndElement(); // Fin de Factura
				#endregion Grabacion de Pie
				miXML.Flush();
				miXML.Close();
			}
			#region Cierre
			dr.Close();
			MessageBox.Show("Archivos Xml Generados", "Información");
			#endregion Cierre
		}

		CultureInfo us = new CultureInfo("en-US");

		private void FacturasFecha(int iTipo)
		{
			DateTime dtDesde = (DateTime) cmbFecha.Value;
			DateTime dtHasta = (DateTime) cmbHasta.Value;
			string stExec = "Exec AutoImpresorFactura '" + dtDesde.ToString("yyyyMMdd") + "', '" 
				+ dtHasta.AddDays(1).ToString("yyyyMMdd") + "', " + iTipo.ToString();
			SqlDataReader dr = Funcion.rEscalarSQL(cdsSeteo, stExec, true);
			#region Crea Documento XML
			string stTipo = "factura";
			string stCodDoc = "1";
//			if (idTipo == 1) 
//			{
//				stTipo = "notaVenta";
//				stCodDoc = "2";
//			}
			string stArchivo = stTipo + "_" + dtDesde.ToString("ddMMMyyyy") + "_"
				+ dtHasta.ToString("ddMMMyyyy")+ ".Xml";
			XmlTextWriter miXML = new XmlTextWriter(stArchivo, null);
			miXML.Formatting = Formatting.Indented;
			miXML.WriteProcessingInstruction("xml", "version=\"1.0\" encoding=\"UTF-8\"");
			#endregion Crea Documento XML
			#region Datos de Empresa
			miXML.WriteStartElement("facturas");
			string stMensaje = "";
			string stVerificaRuc = Funcion.VerificaCedula(drSeteo.Ruc, false);
			if (drSeteo.Empresa.Trim().Length < 3) stMensaje += "Longitud de Nombre de su Empresa Erronea\n";
			if (stVerificaRuc.Length != 0) stMensaje += "Error en RUC de su Empresa\n";
			if (stMensaje.Length > 0)
			{
				miXML.Close();
				MessageBox.Show(stMensaje, "Información");
				Seteo miSeteo = new Seteo();
				miSeteo.MdiParent = this.MdiParent;
				miSeteo.Show();
				return;
			}
			#endregion Datos de Empresa

			int [] idCompra = new int [100000];
			string [] stNumero = new string [100000];
			int i = 0;
			while(dr.Read())
			{
				idCompra[i] = dr.GetInt32(0);
				stNumero[i++] = dr.GetString(1);
			}
			dr.Close();
			for (int j = 0; j < i; j++)
			{
				//Factura(iTipo, idCompra[j], stNumero[j]);
				#region Valores Iniciales
				stExec = "Exec AutoImpresorFactura1 " + iTipo.ToString() + ", " + 
					idCompra[j].ToString();
				dr = Funcion.rEscalarSQL(cdsSeteo, stExec, true);
				string stEstabPtoEmi = "";
				if(dr.Read()) stEstabPtoEmi = dr.GetString(1);
				#endregion ValoresIniciales
				while(stEstabPtoEmi.Length > 0)
				{
					#region Encabezado
					miXML.WriteStartElement(stTipo);
					miXML.WriteElementString("razonSocial", "![DATA[" + drSeteo.Empresa.ToUpper().Trim() + "]]");
					miXML.WriteElementString("ruc", drSeteo.Ruc);
					miXML.WriteElementString("numAut", dr.GetString(0));
					miXML.WriteElementString("codDoc", stCodDoc);
					stEstabPtoEmi = dr.GetString(1);
					if (stEstabPtoEmi.Length != 13)
					{
						dr.Close();
						miXML.Flush();
						miXML.Close();
						MessageBox.Show("Numero: " + stEstabPtoEmi + " debe ser de 13 digitos", "Información");
						return;
					}
					miXML.WriteElementString("establ", stEstabPtoEmi.Substring(0, 3));
					miXML.WriteElementString("ptoEmi", stEstabPtoEmi.Substring(3, 3));
					miXML.WriteElementString("secuencial", stEstabPtoEmi.Substring(6, 7));
					miXML.WriteElementString("fechaEmision",dr.GetDateTime(2).ToString("dd-MM-yyyy"));
					miXML.WriteElementString("razonSocialComprador", "![DATA["+ dr.GetString(3) + "]]");
					miXML.WriteElementString("rucCedulaComprador",dr.GetString(4));
					miXML.WriteElementString("caducidad",dr.GetDateTime(5).ToString("MM/yyyy"));
					#endregion Encabezado
					#region Variables de Pie
					double dTotalSinImp = dr.GetDouble(9);
					double dDescuentos = dr.GetDouble(10);
					double dIce = dr.GetDouble(11);
					double dIva0 = dr.GetDouble(12);
					double dIva12 = dr.GetDouble(13);
					double dPropina = dr.GetDouble(14);
					double dTotalConImp = dr.GetDouble(15);
					#endregion Variables de Pie
					#region Detalles
					miXML.WriteStartElement("detalles");
					string stCodAntes = stEstabPtoEmi;
					while(stCodAntes == stEstabPtoEmi)
					{
						miXML.WriteStartElement("detalle");
						miXML.WriteElementString("concepto", dr.GetString(6));
						miXML.WriteElementString("cantidad", dr.GetDouble(7).ToString("0.00", us));
						miXML.WriteElementString("precioUnitario", dr.GetDouble(8).ToString("0.00", us));
						miXML.WriteEndElement(); // Detalle
						if (dr.Read()) stEstabPtoEmi = dr.GetString(1);
						else stEstabPtoEmi = "";
					}
					miXML.WriteEndElement(); // Detalles
					#endregion Detalles
					#region Grabacion de Pie
					miXML.WriteElementString("totalSinImpuestos", dTotalSinImp.ToString("0.00", us));
					miXML.WriteElementString("descuentos", dDescuentos.ToString("0.00", us));
					miXML.WriteElementString("ICE", dIce.ToString("0.00", us));
					miXML.WriteElementString("IVA0", dIva0.ToString("0.00", us));
					miXML.WriteElementString("IVA12", dIva12.ToString("0.00", us));
					miXML.WriteElementString("propina10", dPropina.ToString("0.00", us));
					miXML.WriteElementString("totalConImpuestos", dTotalConImp.ToString("0.00", us));
					miXML.WriteEndElement(); // Fin de Factura
					#endregion Grabacion de Pie
				}
				dr.Close();
			}
			#region Cierre
			miXML.WriteEndElement(); // Fin de Factura
			dr.Close();
			miXML.Flush();
			miXML.Close();
			#endregion Cierre

			MessageBox.Show("Facturas Generadas", "Información");
		}

		private void FacturaFecha(int iTipo)
		{
			DateTime dtDesde = (DateTime) cmbFecha.Value;
			DateTime dtHasta = (DateTime) cmbHasta.Value;
			#region Datos de Empresa
			string stMensaje = "";
			string stVerificaRuc = Funcion.VerificaCedula(drSeteo.Ruc, false);
			if (drSeteo.Empresa.Trim().Length < 3) stMensaje += "Longitud de Nombre de su Empresa Erronea\n";
			if (stVerificaRuc.Length != 0) stMensaje += "Error en RUC de su Empresa\n";
			if (stMensaje.Length > 0)
			{
				MessageBox.Show(stMensaje, "Información");
				Seteo miSeteo = new Seteo();
				miSeteo.MdiParent = this.MdiParent;
				miSeteo.Show();
				return;
			}
			#endregion Datos de Empresa
			#region Crea Carpeta si no existe
      if (!Directory.Exists("FacturasSri"))
				Directory.CreateDirectory("FacturasSri");
			#endregion Crea Carpeta si no existe
			#region Lee id y numero
			string stExec = "Exec AutoImpresorFactura '" + dtDesde.ToString("yyyyMMdd") + "', '" 
				+ dtHasta.AddDays(1).ToString("yyyyMMdd") + "', " + iTipo.ToString();
			
			SqlDataReader dr = Funcion.rEscalarSQL(cdsSeteo, stExec, true);
			int [] idCompra = new int [100000];
			string [] stNumero = new string [100000];
			int i = 0;
			while(dr.Read())
			{
				idCompra[i] = dr.GetInt32(0);
				stNumero[i++] = dr.GetString(1);
			}
			dr.Close();
			#endregion Lee id y numero

			for (int j = 0; j < i; j++)
			{
				if (iTipo == 0)
					Factura(iTipo, idCompra[j], stNumero[j]);
				else
					NotaVenta(iTipo, idCompra[j], stNumero[j]);
			}
			#region Cierre
			dr.Close();
			#endregion Cierre

			MessageBox.Show("Facturas Generadas", "Información");
		}


		private void NotaCreditoFecha(int iTipo)
		{
			DateTime dtDesde = (DateTime) cmbFecha.Value;
			DateTime dtHasta = (DateTime) cmbHasta.Value;
			#region Datos de Empresa
			string stMensaje = "";
			string stVerificaRuc = Funcion.VerificaCedula(drSeteo.Ruc, false);
			if (drSeteo.Empresa.Trim().Length < 3) stMensaje += "Longitud de Nombre de su Empresa Erronea\n";
			if (stVerificaRuc.Length != 0) stMensaje += "Error en RUC de su Empresa\n";
			if (stMensaje.Length > 0)
			{
				MessageBox.Show(stMensaje, "Información");
				Seteo miSeteo = new Seteo();
				miSeteo.MdiParent = this.MdiParent;
				miSeteo.Show();
				return;
			}
			#endregion Datos de Empresa
			#region Crea Carpeta si no existe
			if (!Directory.Exists("NotaCredito"))
				Directory.CreateDirectory("NotaCredito");
			#endregion Crea Carpeta si no existe
			#region Lee id y numero
			string stExec = "Exec AutoImpresorNotaDebito '" + dtDesde.ToString("yyyyMMdd") + "', '" 
				+ dtHasta.AddDays(1).ToString("yyyyMMdd") + "', " + iTipo.ToString();
			
			SqlDataReader dr = Funcion.rEscalarSQL(cdsSeteo, stExec, true);
			int [] idCompra = new int [100000];
			string [] stNumero = new string [100000];
			int i = 0;
			while(dr.Read())
			{
				idCompra[i] = dr.GetInt32(0);
				stNumero[i++] = dr.GetString(1);
			}
			dr.Close();
			#endregion Lee id y numero

			for (int j = 0; j < i; j++)
			{
				NotaCredito(iTipo, idCompra[j], stNumero[j]);
			}
			#region Cierre
			dr.Close();
			#endregion Cierre

			MessageBox.Show("Notas de Credito Generadas", "Información");
		}


		private void btGenerar_Click(object sender, System.EventArgs e)
		{
			int iTipo = (int) optTipo.EditValue;
			switch(iTipo)
			{
				case 0:
				case 1:
					FacturaFecha(iTipo);
//					FacturasFecha(iTipo);
					break;
				case 2:
				case 3:
					NotaCreditoFecha(iTipo);
					break;
			}
		}

	}
}
