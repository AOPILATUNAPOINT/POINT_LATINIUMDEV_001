using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using LibContabilidad.DataObjects;
using LibFacturacion.DataObjects;
using C1.Data;
using System.IO;

namespace Latinium
{
	/// <summary>
	/// Descripción breve de CompraAutoImpresor.
	/// </summary>
	public class CompraAutoImpresor : System.Windows.Forms.Form
	{
		private Infragistics.Win.Misc.UltraButton btAceptar;
		private Infragistics.Win.UltraWinEditors.UltraOptionSet optTipo;
		private Infragistics.Win.Misc.UltraButton btCancelar;
		private System.Windows.Forms.Label label1;
		private System.Windows.Forms.Label label2;
		private Infragistics.Win.UltraWinSchedule.UltraCalendarCombo cmbDesde;
		private Infragistics.Win.UltraWinSchedule.UltraCalendarCombo cmbHasta;
		private C1.Data.C1DataSet cdsSeteo;
		/// <summary>
		/// Variable del diseñador requerida.
		/// </summary>
		private System.ComponentModel.Container components = null;

		public CompraAutoImpresor()
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
			Infragistics.Win.Appearance appearance1 = new Infragistics.Win.Appearance();
			Infragistics.Win.ValueListItem valueListItem1 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.ValueListItem valueListItem2 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.ValueListItem valueListItem3 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.ValueListItem valueListItem4 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.ValueListItem valueListItem5 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.ValueListItem valueListItem6 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton dateButton1 = new Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton();
			Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton dateButton2 = new Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton();
			this.btAceptar = new Infragistics.Win.Misc.UltraButton();
			this.optTipo = new Infragistics.Win.UltraWinEditors.UltraOptionSet();
			this.btCancelar = new Infragistics.Win.Misc.UltraButton();
			this.cmbDesde = new Infragistics.Win.UltraWinSchedule.UltraCalendarCombo();
			this.cmbHasta = new Infragistics.Win.UltraWinSchedule.UltraCalendarCombo();
			this.label1 = new System.Windows.Forms.Label();
			this.label2 = new System.Windows.Forms.Label();
			this.cdsSeteo = new C1.Data.C1DataSet();
			((System.ComponentModel.ISupportInitialize)(this.optTipo)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbDesde)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbHasta)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteo)).BeginInit();
			this.SuspendLayout();
			// 
			// btAceptar
			// 
			this.btAceptar.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			this.btAceptar.Location = new System.Drawing.Point(40, 248);
			this.btAceptar.Name = "btAceptar";
			this.btAceptar.TabIndex = 0;
			this.btAceptar.Text = "&Aceptar";
			this.btAceptar.Click += new System.EventHandler(this.btAceptar_Click);
			// 
			// optTipo
			// 
			this.optTipo.CheckedIndex = 0;
			this.optTipo.ItemAppearance = appearance1;
			this.optTipo.ItemOrigin = new System.Drawing.Point(4, 4);
			valueListItem1.DataValue = 0;
			valueListItem1.DisplayText = "Factura";
			valueListItem2.DataValue = 1;
			valueListItem2.DisplayText = "Nota de Venta";
			valueListItem3.DataValue = 2;
			valueListItem3.DisplayText = "Nota de Credito";
			valueListItem4.DataValue = 3;
			valueListItem4.DisplayText = "Nota de Debito";
			valueListItem5.DataValue = 4;
			valueListItem5.DisplayText = "Guia de Remision";
			valueListItem6.DataValue = 5;
			valueListItem6.DisplayText = "Retencion";
			this.optTipo.Items.Add(valueListItem1);
			this.optTipo.Items.Add(valueListItem2);
			this.optTipo.Items.Add(valueListItem3);
			this.optTipo.Items.Add(valueListItem4);
			this.optTipo.Items.Add(valueListItem5);
			this.optTipo.Items.Add(valueListItem6);
			this.optTipo.ItemSpacingVertical = 6;
			this.optTipo.Location = new System.Drawing.Point(80, 88);
			this.optTipo.Name = "optTipo";
			this.optTipo.Size = new System.Drawing.Size(128, 136);
			this.optTipo.TabIndex = 1;
			this.optTipo.Text = "Factura";
			// 
			// btCancelar
			// 
			this.btCancelar.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			this.btCancelar.DialogResult = System.Windows.Forms.DialogResult.Cancel;
			this.btCancelar.Location = new System.Drawing.Point(168, 248);
			this.btCancelar.Name = "btCancelar";
			this.btCancelar.TabIndex = 2;
			this.btCancelar.Text = "&Cancelar";
			// 
			// cmbDesde
			// 
			this.cmbDesde.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			dateButton1.Caption = "Today";
			this.cmbDesde.DateButtons.Add(dateButton1);
			this.cmbDesde.Format = "dd/MMM/yyyy";
			this.cmbDesde.Location = new System.Drawing.Point(104, 24);
			this.cmbDesde.Name = "cmbDesde";
			this.cmbDesde.NonAutoSizeHeight = 23;
			this.cmbDesde.Size = new System.Drawing.Size(128, 21);
			this.cmbDesde.TabIndex = 3;
			// 
			// cmbHasta
			// 
			this.cmbHasta.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			dateButton2.Caption = "Today";
			this.cmbHasta.DateButtons.Add(dateButton2);
			this.cmbHasta.Format = "dd/MMM/yyyy";
			this.cmbHasta.Location = new System.Drawing.Point(104, 56);
			this.cmbHasta.Name = "cmbHasta";
			this.cmbHasta.NonAutoSizeHeight = 23;
			this.cmbHasta.Size = new System.Drawing.Size(128, 21);
			this.cmbHasta.TabIndex = 4;
			// 
			// label1
			// 
			this.label1.AutoSize = true;
			this.label1.Location = new System.Drawing.Point(40, 24);
			this.label1.Name = "label1";
			this.label1.Size = new System.Drawing.Size(37, 16);
			this.label1.TabIndex = 5;
			this.label1.Text = "Desde";
			// 
			// label2
			// 
			this.label2.AutoSize = true;
			this.label2.Location = new System.Drawing.Point(40, 56);
			this.label2.Name = "label2";
			this.label2.Size = new System.Drawing.Size(37, 16);
			this.label2.TabIndex = 6;
			this.label2.Text = "Hasta:";
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
			// CompraAutoImpresor
			// 
			this.AcceptButton = this.btAceptar;
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 13);
			this.CancelButton = this.btCancelar;
			this.ClientSize = new System.Drawing.Size(280, 293);
			this.Controls.Add(this.label2);
			this.Controls.Add(this.label1);
			this.Controls.Add(this.cmbHasta);
			this.Controls.Add(this.cmbDesde);
			this.Controls.Add(this.btCancelar);
			this.Controls.Add(this.optTipo);
			this.Controls.Add(this.btAceptar);
			this.MaximizeBox = false;
			this.Name = "CompraAutoImpresor";
			this.Text = "AutoImpresor";
			this.Load += new System.EventHandler(this.CompraAutoImpresor_Load);
			((System.ComponentModel.ISupportInitialize)(this.optTipo)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbDesde)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbHasta)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteo)).EndInit();
			this.ResumeLayout(false);

		}
		#endregion

		SeteoRow_tableView drSeteo;

		private void CompraAutoImpresor_Load(object sender, System.EventArgs e)
		{
			cmbDesde.Value = DateTime.Today;
			cmbHasta.Value = DateTime.Today;
			C1DataRow dr = this.cdsSeteo.TableViews[0].Rows[0];
			drSeteo = SeteoRow_tableView.Obj(dr);
		}

		private void btAceptar_Click(object sender, System.EventArgs e)
		{
//			string stSelect = "SELECT Compra.Numero, Compra.Fecha, Cliente.Nombre, Cliente.Ruc, "
//				+ "DetCompra.Cantidad, Articulo.Articulo, DetCompra.Precio "
//				+ "FROM Compra INNER JOIN Cliente ON Compra.idCliente = Cliente.idCliente "
//				+ "INNER JOIN DetCompra ON DetCompra.idCompra = Compra.idCompra "
//				+ "INNER JOIN Articulo ON Articulo.idArticulo = DetCompra.idArticulo";
//
			#region Factura
			DateTime dtFecha = (DateTime) cmbDesde.Value;
			StreamWriter sr = null;
			try
			{
				string stDir = @"C:\XML_DECLARACIONES\";
				if (!Directory.Exists(stDir)) stDir = "";
				sr= File.CreateText(@stDir + "Factura_" + dtFecha.ToString("MMMyyyy") + ".XML");
			}
			catch
			{
				MessageBox.Show("Esta abierto el archivo Formulario.xml");
				Cursor = Cursors.Default;
			}
			
			sr.WriteLine("<?xml version=\"1.0\" encoding=\"UTF-8\"?>");
			sr.WriteLine("<factura>");
			sr.WriteLine("<razonSocial><![CDATA[" + drSeteo.Empresa.Trim() + "]]></razonSocial >");
			sr.WriteLine("<ruc>"+drSeteo.Ruc.Trim()+"</ruc>");
			sr.WriteLine("<numAut>" + drSeteo.SriAutFact.Trim() + "</numAut>");
			sr.WriteLine("<codDoc>1</codDoc>");
			sr.WriteLine("<estab>" + drSeteo.SriSerie.Substring(0, 3) + "</estab>");
			sr.WriteLine("<ptoEmi>" + drSeteo.SriSerie.Substring(4, 3) + "</ptoEmi>");
			sr.WriteLine("<secuencial>0000001</secuencial>");
			sr.WriteLine("<fechaEmision>28-02-2006</fechaEmision>");
			sr.WriteLine("<razonSocialComprador><![CDATA[Empresa XYZ S.A.]]></razonSocialComprador>");
			sr.WriteLine("<rucCedulaComprador>1795698038001</rucCedulaComprador>");
//			sr.WriteLine("<moneda>><![CDATA[euro]]></moneda>");
			sr.WriteLine("<caducidad>02/2007</caducidad>");
//			sr.WriteLine("<guiaRemision>001-001-0000001</ guiaRemision>");
			sr.WriteLine("<detalles>");
			sr.WriteLine("<detalle>");
			sr.WriteLine("<concepto><![CDATA[taladro]]></concepto>");
			sr.WriteLine("<cantidad>5</cantidad>");
			sr.WriteLine("<precioUnitario>100.00</precioUnitario>");
			sr.WriteLine("</detalle>");
			sr.WriteLine("</detalles>");
			sr.WriteLine("<totalSinImpuestos>500.00</totalSinImpuestos>");
			sr.WriteLine("<descuentos>10</descuentos>");
			sr.WriteLine("<ICE>0</ICE>");
			sr.WriteLine("<IVA0>0</IVA0>");
			sr.WriteLine("<IVA12>58.80</IVA12>");
			sr.WriteLine("<propina10>0</propina10>");
			sr.WriteLine("<totalConImpuestos>558.80</totalConImpuestos>");
			sr.WriteLine("</factura>");

			#endregion Encabezado

		}

		private void cdsSeteo_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			cdsSeteo.Schema.Connections[0].ConnectionString = 
				miFuncion.AccesoBase(MenuLatinium.stDirConta);
		}
	}
}
