using System;
using System.IO;
using System.Data.OleDb;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;

namespace Latinium
{
	/// <summary>
	/// Descripción breve de COA.
	/// </summary>
	public class COA : System.Windows.Forms.Form
	{
		private Infragistics.Win.UltraWinSchedule.UltraCalendarCombo cmbFecha;
		private Infragistics.Win.Misc.UltraLabel ultraLabel1;
		private Infragistics.Win.UltraWinTabControl.UltraTabSharedControlsPage ultraTabSharedControlsPage1;
		private Infragistics.Win.UltraWinTabControl.UltraTabPageControl ultraTabPageControl1;
		private Infragistics.Win.Misc.UltraButton btTransLocal;
		private Infragistics.Win.Misc.UltraButton btComprobantes;
		private Infragistics.Win.Misc.UltraButton btIdentificacion;
		private Infragistics.Win.Misc.UltraButton btTransExterior;
		private Infragistics.Win.Misc.UltraButton btCopiar;
		private Infragistics.Win.UltraWinTabControl.UltraTabControl ultraTabControl1;
		private Infragistics.Win.UltraWinTabControl.UltraTabPageControl ultraTabPageControl2;
		private C1.Data.C1DataSet cdsSeteo;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtIdentificacion;
		private System.Windows.Forms.ErrorProvider errorProvider;
		/// <summary>
		/// Variable del diseñador requerida.
		/// </summary>
		private System.ComponentModel.Container components = null;

		public COA()
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
			Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton dateButton1 = new Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton();
			Infragistics.Win.UltraWinTabControl.UltraTab ultraTab1 = new Infragistics.Win.UltraWinTabControl.UltraTab();
			Infragistics.Win.UltraWinTabControl.UltraTab ultraTab2 = new Infragistics.Win.UltraWinTabControl.UltraTab();
			this.cmbFecha = new Infragistics.Win.UltraWinSchedule.UltraCalendarCombo();
			this.ultraLabel1 = new Infragistics.Win.Misc.UltraLabel();
			this.ultraTabSharedControlsPage1 = new Infragistics.Win.UltraWinTabControl.UltraTabSharedControlsPage();
			this.ultraTabPageControl1 = new Infragistics.Win.UltraWinTabControl.UltraTabPageControl();
			this.btTransLocal = new Infragistics.Win.Misc.UltraButton();
			this.btComprobantes = new Infragistics.Win.Misc.UltraButton();
			this.btIdentificacion = new Infragistics.Win.Misc.UltraButton();
			this.btTransExterior = new Infragistics.Win.Misc.UltraButton();
			this.btCopiar = new Infragistics.Win.Misc.UltraButton();
			this.ultraTabControl1 = new Infragistics.Win.UltraWinTabControl.UltraTabControl();
			this.ultraTabPageControl2 = new Infragistics.Win.UltraWinTabControl.UltraTabPageControl();
			this.cdsSeteo = new C1.Data.C1DataSet();
			this.txtIdentificacion = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.errorProvider = new System.Windows.Forms.ErrorProvider();
			((System.ComponentModel.ISupportInitialize)(this.cmbFecha)).BeginInit();
			this.ultraTabPageControl1.SuspendLayout();
			((System.ComponentModel.ISupportInitialize)(this.ultraTabControl1)).BeginInit();
			this.ultraTabControl1.SuspendLayout();
			this.ultraTabPageControl2.SuspendLayout();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteo)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIdentificacion)).BeginInit();
			this.SuspendLayout();
			// 
			// cmbFecha
			// 
			dateButton1.Caption = "Today";
			this.cmbFecha.DateButtons.Add(dateButton1);
			this.cmbFecha.Format = "dd/MMM/yyyy";
			this.cmbFecha.Location = new System.Drawing.Point(144, 32);
			this.cmbFecha.Name = "cmbFecha";
			this.cmbFecha.NonAutoSizeHeight = 24;
			this.cmbFecha.Size = new System.Drawing.Size(112, 21);
			this.cmbFecha.SpinButtonIncrement = Infragistics.Win.UltraWinSchedule.SpinIncrementUnit.Months;
			this.cmbFecha.SpinButtonsVisible = true;
			this.cmbFecha.TabIndex = 5;
			// 
			// ultraLabel1
			// 
			this.ultraLabel1.AutoSize = true;
			this.ultraLabel1.BackColor = System.Drawing.Color.Transparent;
			this.ultraLabel1.Location = new System.Drawing.Point(72, 32);
			this.ultraLabel1.Name = "ultraLabel1";
			this.ultraLabel1.Size = new System.Drawing.Size(36, 14);
			this.ultraLabel1.TabIndex = 6;
			this.ultraLabel1.Text = "Fecha";
			// 
			// ultraTabSharedControlsPage1
			// 
			this.ultraTabSharedControlsPage1.Location = new System.Drawing.Point(-10000, -10000);
			this.ultraTabSharedControlsPage1.Name = "ultraTabSharedControlsPage1";
			this.ultraTabSharedControlsPage1.Size = new System.Drawing.Size(532, 318);
			// 
			// ultraTabPageControl1
			// 
			this.ultraTabPageControl1.Controls.Add(this.btIdentificacion);
			this.ultraTabPageControl1.Controls.Add(this.btTransLocal);
			this.ultraTabPageControl1.Controls.Add(this.btTransExterior);
			this.ultraTabPageControl1.Controls.Add(this.btComprobantes);
			this.ultraTabPageControl1.Controls.Add(this.btCopiar);
			this.ultraTabPageControl1.Controls.Add(this.cmbFecha);
			this.ultraTabPageControl1.Controls.Add(this.ultraLabel1);
			this.ultraTabPageControl1.Location = new System.Drawing.Point(-10000, -10000);
			this.ultraTabPageControl1.Name = "ultraTabPageControl1";
			this.ultraTabPageControl1.Size = new System.Drawing.Size(532, 318);
			// 
			// btTransLocal
			// 
			this.btTransLocal.Location = new System.Drawing.Point(192, 64);
			this.btTransLocal.Name = "btTransLocal";
			this.btTransLocal.Size = new System.Drawing.Size(104, 40);
			this.btTransLocal.TabIndex = 1;
			this.btTransLocal.Text = "Transacciones Locales";
			this.btTransLocal.Click += new System.EventHandler(this.btTransLocal_Click);
			// 
			// btComprobantes
			// 
			this.btComprobantes.Location = new System.Drawing.Point(32, 120);
			this.btComprobantes.Name = "btComprobantes";
			this.btComprobantes.Size = new System.Drawing.Size(104, 40);
			this.btComprobantes.TabIndex = 2;
			this.btComprobantes.Text = "Comprobantes Afectados ND/NC";
			// 
			// btIdentificacion
			// 
			this.btIdentificacion.Location = new System.Drawing.Point(32, 64);
			this.btIdentificacion.Name = "btIdentificacion";
			this.btIdentificacion.Size = new System.Drawing.Size(104, 40);
			this.btIdentificacion.TabIndex = 0;
			this.btIdentificacion.Text = "Identificacion";
			this.btIdentificacion.Click += new System.EventHandler(this.btIdentificacion_Click);
			// 
			// btTransExterior
			// 
			this.btTransExterior.Location = new System.Drawing.Point(192, 120);
			this.btTransExterior.Name = "btTransExterior";
			this.btTransExterior.Size = new System.Drawing.Size(104, 40);
			this.btTransExterior.TabIndex = 3;
			this.btTransExterior.Text = "Transacciones en el Exterior";
			// 
			// btCopiar
			// 
			this.btCopiar.Location = new System.Drawing.Point(32, 208);
			this.btCopiar.Name = "btCopiar";
			this.btCopiar.Size = new System.Drawing.Size(104, 40);
			this.btCopiar.TabIndex = 4;
			this.btCopiar.Text = "Copiar Archivos a Disquete";
			// 
			// ultraTabControl1
			// 
			this.ultraTabControl1.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.ultraTabControl1.Controls.Add(this.ultraTabSharedControlsPage1);
			this.ultraTabControl1.Controls.Add(this.ultraTabPageControl1);
			this.ultraTabControl1.Controls.Add(this.ultraTabPageControl2);
			this.ultraTabControl1.Location = new System.Drawing.Point(32, 24);
			this.ultraTabControl1.Name = "ultraTabControl1";
			this.ultraTabControl1.SharedControlsPage = this.ultraTabSharedControlsPage1;
			this.ultraTabControl1.Size = new System.Drawing.Size(536, 344);
			this.ultraTabControl1.TabIndex = 7;
			ultraTab1.TabPage = this.ultraTabPageControl1;
			ultraTab1.Text = "Opciones";
			ultraTab2.TabPage = this.ultraTabPageControl2;
			ultraTab2.Text = "Archivos Generados";
			this.ultraTabControl1.Tabs.AddRange(new Infragistics.Win.UltraWinTabControl.UltraTab[] {
																																															 ultraTab1,
																																															 ultraTab2});
			// 
			// ultraTabPageControl2
			// 
			this.ultraTabPageControl2.Controls.Add(this.txtIdentificacion);
			this.ultraTabPageControl2.Location = new System.Drawing.Point(1, 23);
			this.ultraTabPageControl2.Name = "ultraTabPageControl2";
			this.ultraTabPageControl2.Size = new System.Drawing.Size(532, 318);
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
			// txtIdentificacion
			// 
			this.txtIdentificacion.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.txtIdentificacion.Location = new System.Drawing.Point(24, 16);
			this.txtIdentificacion.Multiline = true;
			this.txtIdentificacion.Name = "txtIdentificacion";
			this.txtIdentificacion.Size = new System.Drawing.Size(496, 64);
			this.txtIdentificacion.TabIndex = 0;
			// 
			// errorProvider
			// 
			this.errorProvider.ContainerControl = this;
			// 
			// COA
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 13);
			this.ClientSize = new System.Drawing.Size(600, 406);
			this.Controls.Add(this.ultraTabControl1);
			this.Name = "COA";
			this.Text = "COA";
			this.Load += new System.EventHandler(this.COA_Load);
			((System.ComponentModel.ISupportInitialize)(this.cmbFecha)).EndInit();
			this.ultraTabPageControl1.ResumeLayout(false);
			((System.ComponentModel.ISupportInitialize)(this.ultraTabControl1)).EndInit();
			this.ultraTabControl1.ResumeLayout(false);
			this.ultraTabPageControl2.ResumeLayout(false);
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteo)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIdentificacion)).EndInit();
			this.ResumeLayout(false);

		}
		#endregion

		private void btIdentificacion_Click(object sender, System.EventArgs e)
		{
			string stArchivo;
			DateTime dtFecha = (DateTime) cmbFecha.Value;

			stArchivo = "ID" + NombreArchivo();
		
			string stTexto = drSeteo["Ruc"].ToString().Trim();
			string stTemp = Funcion.VerificaCedula(stTexto, false);
			string stMensaje = "";
			bool bOkEmpresa = true;
			if (stTemp.Length > 0)
			{
				stMensaje += stTemp + "\n";
				bOkEmpresa = false;
			}
			int iCod = (dtFecha.Year - 1995) * 12 + dtFecha.Month;
			if (iCod < 10) stTexto += "0";
			stTexto += iCod.ToString();
		
			//*Razon Social de 60 caracteres
			stTemp = drSeteo["Empresa"].ToString().Trim();
			if (stTemp.Length == 0)
			{
				stMensaje += "Ingrese Nombre de la Empresa\n";
				bOkEmpresa = false;
			}
			if (stTemp.Length >= 60) stTemp = stTemp.Substring(1,60);
			stTemp = stTemp.PadRight(60);
			stTexto += stTemp;
		
			//*Telefono responsable
			stTemp = drSeteo["Telefono"].ToString().Trim();
			if (stTemp.Length != 9)
			{
				stMensaje += "Telefono debe Tener 9 Caracteres de Longitud\n";
				bOkEmpresa = false;
			}
			stTexto += stTemp;

			//*Fax responsable
			stTemp = drSeteo["Fax"].ToString().Trim();
			if (stTemp.Length != 9)
			{
				stMensaje += "Fax debe Tener 9 Caracteres de Longitud\n";
				bOkEmpresa = false;
			}
			stTexto += stTemp;
		
			//
			//*Email de 60 caracteres
			stTemp = drSeteo["EMail"].ToString().Trim();
			if (stTemp.Length == 0)
			{
				stMensaje += "Ingrese EMail de la Empresa\n";
				bOkEmpresa = false;
			}
			if (stTemp.Length >= 60) stTemp = stTemp.Substring(1,60);
			stTemp = stTemp.PadRight(60);
			stTexto += stTemp;

			//STRTOFILE(lcTexto, lcArchivo)
			if (!bOkEmpresa)
			{
				MessageBox.Show(stMensaje, "Informacion");
				using(Seteo miSeteo = new Seteo())
				{
					miSeteo.ShowDialog();
				}
				cdsSeteo.Clear();
				cdsSeteo.Fill();
				return;
			}
			txtIdentificacion.Text = stTexto;
			MessageBox.Show(stTexto);
		}

		private void cdsSeteo_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
		
		}

		C1.Data.C1DataRow drSeteo;
		private void COA_Load(object sender, System.EventArgs e)
		{
			drSeteo = cdsSeteo.TableViews["Seteo"].Rows[0];
		}

		private string CodigoPeriodo()
		{
			DateTime dtFecha = (DateTime) cmbFecha.Value;
			string stCod = "";
			int iCod = (dtFecha.Year - 1995)*12 + dtFecha.Month;
			if (iCod < 10) stCod = "0";
			stCod += iCod.ToString();
			return stCod;
		}

		private string NombreArchivo()
		{
			DateTime dtFecha = (DateTime) cmbFecha.Value;
			string stArchivo = "";
			if (dtFecha.Month < 10) stArchivo = "0";
			stArchivo += dtFecha.Month.ToString().Trim()
				+dtFecha.Year.ToString().Trim() + ".ANE";
			return stArchivo;
		}

		private bool VerificaFecha()
		{
			bool bOk = true;
			DateTime dtFecha = (DateTime) cmbFecha.Value;
			errorProvider.SetError(cmbFecha, "");
			if (dtFecha.Year < 1995)
			{
				errorProvider.SetError(cmbFecha, "Fecha no puede ser menor a 1995");
				bOk = false;
			}
			return bOk;
		}

		private void btTransLocal_Click(object sender, System.EventArgs e)
		{
			if (!VerificaFecha()) return;

			string stArchivo, stTexto, sTemp;
			DateTime dtFecha = (DateTime) cmbFecha.Value;
			stArchivo = "TL" + NombreArchivo();

			stTexto = drSeteo["Ruc"].ToString().Trim();
			sTemp = Funcion.VerificaCedula(stTexto, false);
			if (sTemp.Length > 0 || stTexto.Length != 13)
			{
				if (stTexto.Length != 13) sTemp = "Ruc Debe Tener 13 Digitos";
				MessageBox.Show(sTemp, "Error en RUC");
				using(Seteo miSeteo = new Seteo())
				{
					miSeteo.ShowDialog();
				}
				return;
			}
			stTexto += CodigoPeriodo();

			string stEncabezado = stTexto;
			stTexto = "";
			string stSelect = "SELECT Compra.Fecha, Compra.Numero, "
				+ "ISNULL(Cliente.Nombre, ''), "
				+ "ISNULL(ClienteTipoRuc.Codigo, ''), Compra.idTipoFactura, "
				+ "ISNULL(Compra.SerieFactura, ''), "
				+ "ISNULL(CompraComprobante.Codigo, ''), "
				+ "ISNULL(Compra.AutFactura, ''), "
				+ "ISNULL(CompraCredTribut.Codigo, ''), "
				+ "Isnull(Compra.Descuento, 0), "
				+ "IsNull(Compra.Descuento0, 0), "
				+ "IsNull(Compra.SubtotalIva, 0), "
				+ "IsNull(Compra.SubtotalExcento, 0), "
				+ "IsNull(Compra.Iva, 0), "
				+ "IsNull(Compra.DerechoDevolucion, 0), "
				+ "IsNull(Compra.Ice, 0), IsNull(Compra.idCompra, 0), "
				+ "IsNull(Cliente.Ambos, 0), "
				+ "IsNull(Cliente.Ruc, ''), "
				+ "IsNull(Cliente.AutorizaFactura, ''), "
				+ "IsNull(Cliente.SerieFactura, ''), "
				+ "IsNull(Compra.BienServicio, 0), "
				+ "IsNull(RetencionIva.Porcentaje, 0), "
				+ "IsNull(Retencion.IvaRetenido, 0), "
				+ "IsNull(Cliente.idCliente, 0) "
				+ "FROM Compra INNER JOIN "
				+ "Cliente ON Compra.idCliente = Cliente.idCliente INNER JOIN "
				+ "ClienteTipoRuc ON Cliente.idTipoRuc = ClienteTipoRuc.idTipoRuc INNER JOIN "
				+ "CompraComprobante ON Compra.idComprobante = CompraComprobante.idComprobante "
				+ "Inner JOIN CompraCredTribut ON Compra.idCredTributario = CompraCredTribut.idCredTributario "
				+ "Left Outer JOIN Retencion ON Compra.idCompra = Retencion.idCompra "
				+ "INNER JOIN RetencionIva ON Retencion.idRetIva = RetencionIva.idRetIva "
				+ "WHERE Compra.idTipoFactura = 4 And DatePart(Month, Compra.Fecha) = "
				+ dtFecha.Month.ToString() + " And DatePart(Year, Compra.Fecha) = "
				+ dtFecha.Year.ToString();

			cdsSeteo.Schema.Connections[0].Open();
			OleDbDataReader miReader;
			C1.Data.SchemaObjects.Connection miConeccion = cdsSeteo.Schema.Connections[0];
			OleDbCommand cmdCant = new OleDbCommand (stSelect
				, (OleDbConnection) miConeccion.DbConnection);
			//			udsAnalisis.Rows.Clear();

			StreamWriter sr= File.CreateText(stArchivo);
			string stLinea = "";
			miReader = cmdCant.ExecuteReader();
			try
			{
				while (miReader.Read())
				{
					DateTime dtFechaC = miReader.GetDateTime(0);
					string stNumeroC = miReader.GetString(1);
					string stClienteN = miReader.GetString(2);
					string stCodTipoRuc = miReader.GetString(3);
					int IdTipoF = miReader.GetInt32(4);
					string stSerieF = miReader.GetString(5);
					string stComprobCod = miReader.GetString(6);
					string stAutFactC = miReader.GetString(7);
					string stCredTribCod = miReader.GetString(8);
					double dDescIva = miReader.GetDouble(9);
					double dDescIva0 = miReader.GetDouble(10);
					double dSubTotalIva = miReader.GetDouble(11);
					double dSubTotalIva0 = miReader.GetDouble(12);
					double dIva = miReader.GetDouble(13);
					bool bDerechoDev = miReader.GetBoolean(14);
					double dIce = miReader.GetDouble(15);
					int IdCompra = miReader.GetInt32(16);
					bool bAmbos = miReader.GetBoolean(17);
					string stClienteRuc = miReader.GetString(18);
					string stClienteAutF = miReader.GetString(19);
					string stClienteSerF = miReader.GetString(20);
					int iBienServicio = miReader.GetInt32(21);
					double dPorcIva = miReader.GetDouble(22);
					double dIvaRetenido1 = miReader.GetDouble(23);
					int IdCliente = miReader.GetInt32(24);

					stTexto = stEncabezado;
					string stMensaje = "";
					bool bVerificaCompra = true;

					#region Verificacion de Datos Generales
					sTemp = stNumeroC.Trim();
					if (sTemp.Length == 0)
					{
						stMensaje += "Ingrese numero de Factura\n";
						bVerificaCompra = false;
					}
					//					stTexto += sTemp;

					sTemp = stClienteN.Trim();
					if (sTemp.Length == 0)
					{
						stMensaje += "Ingrese Cliente\n";
						bVerificaCompra = false;
					}
					//					stTexto += sTemp;
					#endregion

					#region Codigo de tipo de proveedor [Secuencial de Transaccion]
					sTemp = stCodTipoRuc.Trim();
					if (sTemp.Length == 0)
					{
						//						sTemp = "0";
						//						stCodTipoRuc = "1";
						//						stMensaje = "No existe Tipo de RUC\n";
						//						bVerificaCompra = false;
						MessageBox.Show("No existe tipo de RUC", "Informacion");
						miReader.Close();
						cdsSeteo.Schema.Connections[0].Close();
						sr.Close();
						using (Cliente miCliente = new Cliente(false, IdCliente))
						{
							miCliente.ShowDialog();
						}
						return;
					}

					if (int.Parse(sTemp) > 3 && bAmbos) // Marcado Como Proveedor y cliente
					{
						int iTemp = int.Parse(sTemp) - 3;
						sTemp = "0" + iTemp.ToString();
					}

					if (sTemp.Length != 2)
					{
						stMensaje = "Error en Codigo de Cliente\n";
						bVerificaCompra = false;
					}

					if (int.Parse(sTemp) < 1 || int.Parse(sTemp) > 3)
					{
						stMensaje = "Error en Codigo de Cliente\n";
						bVerificaCompra = false;
					}
					stTexto += sTemp;
					#endregion	Fin de codigo de tipo de proveedor

					#region	RUC, cedula o pasaporte
					sTemp = stClienteRuc.Trim();
					switch(int.Parse(stCodTipoRuc))
					{
						case 1:
						case 4:
							if (sTemp.Length != 13)
							{
								stMensaje += "Ruc de Cliente debe tener 13 caractares\n";
								bVerificaCompra = false;
							}
							break;
						case 2:
						case 5:
							if (sTemp.Length != 10)
							{
								stMensaje += "Cedula debe tener 10 caracteres\n";
								bVerificaCompra = false;
							}
							sTemp = "000" + sTemp;
							break;
						case 3:
						case 6:
							if (sTemp.Length < 5)
							{
								stMensaje += "Pasaporte no puede tener menos de 5 caracteres\n";
								bVerificaCompra = false;
							}
							sTemp = sTemp.PadRight(13, '0');
							break;
						case 7:
							sTemp = "9999999999999";
							break;
					}
					//						Funcion.VerificaCedula();
					//	*Comprobante de venta emitido en el exterior
					if (stComprobCod == "15") sTemp = "0000000000001";
					stTexto += sTemp;
					#endregion // Fin de Ruc Cedula Pasaporte

					#region Tipo de Comprobante
					sTemp = stComprobCod.Trim();
					if (sTemp.Length != 2)
					{
						stMensaje += "Codigo de Comprobante debe tener 5 caracteres\n";
						bVerificaCompra = false;
					}
					stTexto += sTemp;
					#endregion Tipo de Comprobante

					#region Fechas
					// Fecha de Emision
					stTexto += dtFechaC.ToString("ddMMyyyy");
					// Fecha de Contabilizacion
					stTexto += dtFechaC.ToString("ddMMyyyy");
					#endregion
				
					#region //	*No. de serie de Factura
					sTemp = stSerieF.Trim();
					if (sTemp.Length == 0)
					{
						sTemp = stClienteSerF.Trim();
						if (sTemp.Length != 6)
						{
							stMensaje += "Ingrese Numero de Serie de Factura\n";
							bVerificaCompra = false;
						}
					}
					stTexto += sTemp;
					#endregion

					# region // *No Secuencial del comprobante
					sTemp = stNumeroC.Trim();
					if (sTemp.Length == 0)
					{
						stMensaje += "Ingrese Numero de Factura\n";
						bVerificaCompra = false;
					}
					
					if (sTemp.Length < 7) sTemp = sTemp.PadLeft(7, '0');
					stTexto += sTemp;
					#endregion 

					#region //	*No. de autorización del comprobante 
					sTemp = stAutFactC.Trim();
					if (sTemp.Length == 0) sTemp = stClienteAutF.Trim();
					if (sTemp.Length == 9) sTemp = "0" + sTemp;
					if (sTemp.Length != 10)
					{
						stMensaje += "Ingrese Numero de Autorizacion de Factura\n";
						bVerificaCompra = false;
					}
					stTexto += sTemp;
					#endregion

					#region //	*Identificacion del credito tributario o Gasto
					sTemp = stCredTribCod.Trim();
					if (sTemp.Length != 2)
					{
						stMensaje += "Error en Identificacion de Credito Tributario\n";
						bVerificaCompra = false;
					}
					stTexto += sTemp;
					#endregion

					#region	//	*Base imponible grabable
					double dSubt1 = dSubTotalIva - dDescIva;
					double dSubt2 = dSubTotalIva0 - dDescIva0;

					if (dIvaRetenido1 - dPorcIva*dIva/100.0 > .01)
					{
						stMensaje += "Error en Identificacion de Credito Tributario\n";
						bVerificaCompra = false;
					}
					stTexto += FormatoNum(dSubt1);
					#endregion //Base Imponible

					#region //	*Codigo Porcentaje de IVA no esta completo ***********
					if (dtFechaC.Year < 2000) sTemp = "1";
					else if (dtFechaC.Year < 2001) sTemp = "2";
					else if (dtFechaC.Year < 2001) sTemp = "3";
					else sTemp = "2";
					stTexto += sTemp;
					#endregion //Codigo % Iva

					#region	//	*Base imponible tarifa 0
					stTexto += FormatoNum(dSubt2);
					//	*Monto IVA
					stTexto += FormatoNum(dIva);
					//	*MONTO ICE
					stTexto += FormatoNum(dIce);

					//	*Monto IVA (1) Transferencia de bienes Existen Compras separadas para bien y servicio
					// Tratar de Modificar con los dos valores de retenciones
					if (iBienServicio == 1) // Bienes
						stTexto += FormatoNum(dIva);
					else 
						stTexto += FormatoNum(0);
					
					#region //	*Codigo porcentaje IVA en Transferencia de bienes
					if (dPorcIva == 0)	sTemp = "0";
					else if (dPorcIva == 30)	sTemp = "1";
					else if (dPorcIva == 70) sTemp = "2";
					else sTemp = "3";
					
					// Liquidacion de compras
					if (stComprobCod =="03") sTemp = "3";
					if (iBienServicio == 1)
					{
						stTexto += sTemp;
						//	*Monto de retencion de IVA (1) Transferencia de bienes
						stTexto += FormatoNum(dIvaRetenido1);
					}
					else 
					{
						stTexto += "0";
						stTexto += FormatoNum(0);
					}
					#endregion

					//	*Monto IVA (2) Prestacion de servicios
					if (iBienServicio == 2)
						stTexto += FormatoNum(dIva);
					else 
						stTexto += FormatoNum(0);
				
					//	*Codigo porcentaje IVA en Prestacion de servicios
					if (dPorcIva == 0)	sTemp = "0";
					else if (dPorcIva == 30)	sTemp = "1";
					else if (dPorcIva == 70) sTemp = "2";
					else sTemp = "3";
					
					// Liquidacion de compras
					if (stComprobCod =="03") sTemp = "3";
					if (iBienServicio == 2)
					{
						stTexto += sTemp;
						//	*Monto de retencion de IVA (1) Prestacion de Servicios
						stTexto += FormatoNum(dIvaRetenido1);
					}
					else 
					{
						stTexto += "0";
						stTexto += FormatoNum(0);
					}

					//	Transaccion con derecho a devolucion
					sTemp = "N";
					if (bDerechoDev) sTemp = "S";
					stTexto += sTemp;

					//	*Cantidad de comprobantes de venta emitidos al mes
					sTemp = "            ";
					stTexto += sTemp;

					//	*Cantidad de Notas de Credito emitidas al mes
					stTexto += sTemp;

					//	*Cantidad de Notas de Debito emitidas al mes
					stTexto += sTemp;
					#endregion

					if (!bVerificaCompra)
					{
						sr.Close();
						miReader.Close();
						cdsSeteo.Schema.Connections[0].Close();
						MessageBox.Show(stMensaje, "Informacion");
						using (Compra miCompra = new Compra())
						{
							miCompra.ShowDialog();
						}
						return;
					}
					sr.WriteLine(stTexto);
					stLinea += stTexto + "\n";
				}
			}
			catch(System.Exception ex)
			{
				MessageBox.Show(ex.Message, "Select Compra");
			}
			miReader.Close();
			sr.Close();
			cdsSeteo.Schema.Connections[0].Close();
			txtIdentificacion.Value = stTexto;
			MessageBox.Show("Archivo de Transacciones Locales Generado");
		}
	

		private string FormatoNum(double dValor)
		{
			string sTemp = dValor.ToString("0.00");
			sTemp = sTemp.PadLeft(12);
			return sTemp;
		}


			
			
			
			
			
			//
//*Filtro de Ventas No Exportaciones
//SELECT Compras
//ldFecha = ThisForm.Mes_anio1.FechaRet.Value
//lcFiltro = "SET FILTER TO MONTH(Compras.Fecha) =" + ;
//	ALLTRIM(STR(MONTH(ldFecha))) + " AND INLIST(Id_Tipo, 1)" ;
//	+ " AND YEAR(Compras.Fecha) =" + ALLTRIM(STR(YEAR(ldFecha))) ;
//	+ " AND NOT INLIST(VAL(Compras.Comp_Tipo), 16, 17)"
//&lcFiltro
//SET ORDER TO id_dist
//SCAN
//	IF INLIST(Compras.Comp_Tipo, "04", "05")
//		lnTemp = ""
//	ENDIF
//	lcTemp = ALLTRIM(Compras.Numero)
//	ldFecha = Compras.Fecha
//	lcTexto = lcTexto + lcEncabezado 
//	IF LEN(lcTemp) = 0
//		MESSAGEBOX(ThisForm.Subr.lcMensaje(36) + ;
//			DTOC(ldFecha) + ThisForm.Subr.lcMensaje(37) + ALLTRIM(Distribuidor.Desc_Dist))
//		DO FORM &lcDir WITH 1, -Compras.Id_Compras
//		RETURN
//	ENDIF
//	SELECT Distribuidor
//	IF LEN(ALLTRIM(Distribuidor.Desc_Dist)) = 0
//		MESSAGEBOX(ThisForm.Subr.lcMensaje(38) + Compras.Numero)
//		DO FORM &lcDir WITH 1, -Compras.Id_Compras
//		RETURN
//	ENDIF
//	
//	*Codigo de tipo de proveedor
//	lcTemp = ALLTRIM(Distribuidor.Cod_Tipo)
//	IF VAL(lcTemp) < 4 AND Distribuidor.Ambos
//		lcTemp = "0"+ALLTRIM(STR(VAL(lcTemp) + 3))
//	ENDIF
//	IF LEN(lcTemp) != 2
//		MESSAGEBOX(ThisForm.Subr.lcMensaje(39) + Distribuidor.Desc_Dist)
//		DO FORM &lcDir WITH 1, -Compras.Id_Compras
//		RETURN
//	ENDIF
//	IF NOT (VAL(lcTemp) >= 4 AND VAL(lcTemp) <= 7)
//		MESSAGEBOX(ThisForm.Subr.lcMensaje(40) + Distribuidor.Desc_Dist)
//		DO FORM &lcDir WITH 1, -Compras.Id_Compras
//		RETURN
//	ENDIF	
//	lcTexto = lcTexto + lcTemp
//	lcCodDist = lcTemp
//	*Fin de codigo de tipo de proveedor
//	
//	*No de documento RUC, cedula pasaporte
//	lcTemp = ALLTRIM(Distribuidor.Ruc)
//	DO CASE 
//		*Venta con RUC
//		CASE INLIST(lcCodDist, "04")
//			*Longitud de 13 caracteres
//			IF LEN(lcTemp) != 13
//				MESSAGEBOX(ThisForm.Subr.lcMensaje(41) + ;
//					ALLTRIM(Distribuidor.Desc_Dist) ;
//					+ ThisForm.Subr.lcMensaje(42))
//				DO FORM &lcDir WITH 1, -Compras.Id_Compras
//				RETURN
//			ENDIF
//			*Debe terminar en 001
//			IF SUBSTR(lcTexto,11,13) != "001"
//				MESSAGEBOX(ThisForm.Subr.lcMensaje(41) + ;
//					ALLTRIM(Distribuidor.Desc_Dist) ;
//					+ ThisForm.Subr.lcMensaje(43))
//				DO FORM &lcDir WITH 1, -Compras.Id_Compras
//				RETURN
//			ENDIF
//		*Venta con Cedula
//		CASE INLIST(lcCodDist, "05")
//			*Longitud de 10 caracteres
//			IF LEN(lcTemp) != 10
//				MESSAGEBOX(ThisForm.Subr.lcMensaje(44) + ;
//					ALLTRIM(Distribuidor.Desc_Dist) ;
//					+ ThisForm.Subr.lcMensaje(45))
//				DO FORM &lcDir WITH 1, -Compras.Id_Compras
//				RETURN
//			ENDIF
//			lcTemp = "000" + lcTemp
//		*Venta con Pasaporte
//		CASE INLIST(lcCodDist, "06")
//			*Longitud minima 5 caracteres
//			IF LEN(lcTemp) < 5
//				MESSAGEBOX(ThisForm.Subr.lcMensaje(46) + ;
//					ALLTRIM(Distribuidor.Desc_Dist) ;
//					+ ThisForm.Subr.lcMensaje(47))
//				DO FORM &lcDir WITH 1, -Compras.Id_Compras
//				RETURN
//			ENDIF
//			IF LEN(lcTemp) < 13
//				lcTemp = Substr("000000000",1,13-LEN(lcTemp)) + lcTemp
//			ENDIF
//		*Consumidor Final
//		CASE INLIST(lcCodDist, "07")
//			lcTemp = "9999999999999"
//	ENDCASE
//	lcTexto = lcTexto + lcTemp
//	*Fin de RUC, cedula, etc
//	
//	*Tipo de Comprobante
//	SELECT Compras
//	lcTemp = compras.comp_tipo && "18"
//	lcTexto = lcTexto + lcTemp 
//	
//	*Fecha de emision de comprobante
//	ldFecha = Compras.Fecha
//	IF ldFecha < DATE(1995,1,1) && OR ldFecha > DATE(2004,12,31)
//		MESSAGEBOX(ThisForm.Subr.lcMensaje(48) + ;
//			ALLTRIM(Compras.Numero) )
//		DO FORM &lcDir WITH 1, -Compras.Id_Compras
//		RETURN
//	ENDIF
//	lcTemp = ALLTRIM(STR(DAY(ldFecha)))
//	*Caso de Compras
//	IF LEN(lcTemp) < 2 THEN
//		lcTemp = "0" + lcTemp
//	ENDIF
//	IF Compras.Id_Tipo = 1 && Venta
//		IF MONTH(Compras.Fecha) != 12
//			lcTemp = Alltrim(STR(DAY(DATE(YEAR(Compras.Fecha),;
//				MONTH(Compras.Fecha)+1,1)-1) ))
//		ELSE
//			lcTemp = Alltrim(STR(DAY(DATE(YEAR(Compras.Fecha),12,31) )))
//		ENDIF
//	ENDIF
//
//	lcFecha1 = lcTemp
//	lcTemp = ALLTRIM(STR(MONTH(ldFecha)))
//	IF LEN(lcTemp) < 2 THEN
//		lcTemp = "0" + lcTemp
//	ENDIF
//	lcFecha1 = lcFecha1 + lcTemp
//	
//	lcFecha1 = lcFecha1 + ALLTRIM(STR(YEAR(ldFecha)))
//	lcTexto = lcTexto + lcFecha1
//	
//	*Fecha de contabilizacion
//	lcTexto = lcTexto + lcFecha1
//	
//	*No. de serie de comprobante
//	IF Compras.Id_Tipo = 1 && Venta
//		lcTemp = SPACE(6)
//	ELSE
//		lcTemp = ALLTRIM(Compras.Serie_Fact)
//		IF LEN(lcTemp) = 0
//			lcTemp = ALLTRIM(Distribuidor.Serie_Fact)
//		ENDIF
//	ENDIF
//	lcTexto = lcTexto + lcTemp
//	
//	*No Secuencial del comprobante
//	lcTemp = SPACE(7)
//	lcTexto = lcTexto + lcTemp
//	
//	*No. de autorización del comprobante 
//	lcTemp = SPACE(10)
//	lcTexto = lcTexto + lcTemp
//	
//	*Identificacion del credito tributario o Gasto
//	lcTemp = "00"
//	lcTexto = lcTexto + lcTemp
//	
//	*Base imponible grabada
//	lnIdDist = Compras.Id_Dist
//	lnSubt1 = 0
//	lnSubt2 = 0
//	lnNotasDeb = 0
//	lnNotasCre = 0
//	lnIva = 0
//	lnIce = 0
//	lnIvaRet1 = 0
//	lnIvaRet2 = 0
//	lnCuenta = 0
//	SELECT Compras
//	SCAN WHILE Compras.Id_Dist = lnIdDist
//*		IF NOT INLIST(Compras.Comp_Tipo, "04", "05")
//			lnSubt1 = lnSubt1 + Compras.SubTotal1 - Compras.Desc1
//			lnSubt2 = lnSubt2 + Compras.SubTotal2 - Compras.Desc2
//			lnIva = lnIva + Compras.Iva
//			lnIce = lnIce + Compras.Ice
//*		ENDIF
//		IF Compras.Comp_Tipo = "04"
//			lnNotasCre = lnNotasCre + 1
//		ENDIF
//		IF Compras.Comp_Tipo = "05"
//			lnNotasDeb = lnNotasDeb + 1
//		ENDIF
//		lnCuenta = lnCuenta + 1
//		SELECT Compras
//	ENDSCAN
//	SKIP -1
//	lcTemp = ThisForm.CmdFormatoNum.Click(lnSubt2)
//	lcTexto = lcTexto + lcTemp
//	
//	*Codigo Porcentaje de IVA
//	DO CASE
//		CASE ldFecha < DATE(2000,1,1)
//			lcTemp = "1"
//		CASE ldFecha < DATE(2001,6,1)
//			lcTemp = "2"
//		CASE ldFecha < DATE(2001,9,1)
//			lcTemp = "3"
//		OTHERWISE
//			lcTemp = "2"
//	ENDCASE		
//	lcTexto = lcTexto + lcTemp
//
//	*Base imponible tarifa 0
//	lcTemp = ThisForm.CmdFormatoNum.Click(lnSubt1)
//	lcTexto = lcTexto + lcTemp
//	
//	*Monto IVA
//	lcTemp = ThisForm.CmdFormatoNum.Click(lnIva)
//	lcTexto = lcTexto + lcTemp
//	
//	*MONTO ICE
//	lcTemp = ThisForm.CmdFormatoNum.Click(lnIce)
//	lcTexto = lcTexto + lcTemp
//	
//	*Monto IVA (1) Transferencia de bienes
//	lcTemp = SPACE(12)
//	lcTexto = lcTexto + lcTemp
//	
//	*Codigo porcentaje IVA en Transferencia de bienes
//	lcTemp = " "
//	lcTexto = lcTexto + lcTemp
//	
//	*Monto de retencion de IVA (1) Transferencia de bienes
//	lcTemp = SPACE(12)
//	lcTexto = lcTexto + lcTemp
//	
//	*Monto IVA (2) Prestacion de servicios
//	lcTemp = SPACE(12)
//	lcTexto = lcTexto + lcTemp
//	
//	*Codigo porcentaje IVA en Prestacion de servicios
//	lcTemp = " "
//	lcTexto = lcTexto + lcTemp
//	
//	*Monto de retencion de IVA (2) Transferencia de bienes
//	lcTemp = SPACE(12)
//	lcTexto = lcTexto + lcTemp
//
//	*Transaccion con derecho a devolucion
//	lcTemp = "N"
//	lcTexto = lcTexto + lcTemp
//
//	*Cantidad de comprobantes de venta emitidos al mes
//	lcTemp = ThisForm.CmdFormatoNum.Click(lnCuenta, .T.)
//	lcTexto = lcTexto + lcTemp
//	
//	*Cantidad de Notas de Credito emitidas al mes
//	lcTemp = ThisForm.CmdFormatoNum.Click(lnNotasCre, .T.)
//	lcTexto = lcTexto + lcTemp
//
//	*Cantidad de Notas de Debito emitidas al mes
//	lcTemp = ThisForm.CmdFormatoNum.Click(lnNotasDeb, .T.)
//	lcTexto = lcTexto + lcTemp + CHR(13) + CHR(10)
//	SELECT Compras
//ENDSCAN
//
//MESSAGEBOX(ThisForm.Subr.lcMensaje(49) + lcArchivo)
//STRTOFILE(lcTexto, lcArchivo)
//SELECT Compras
//SET FILTER TO 

		
	}
}
