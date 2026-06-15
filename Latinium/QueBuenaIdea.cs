using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using System.Data;
using System.Data.SqlClient;
using System.Data.OleDb;
using System.Globalization;

namespace Latinium
{
	/// <summary>
	/// Descripción breve de QueBuenaIdea.
	/// </summary>
	public class QueBuenaIdea : DevExpress.XtraEditors.XtraForm
	{
		private Infragistics.Win.Misc.UltraButton btCliente;
		private Infragistics.Win.UltraWinGrid.UltraGrid ultraGrid1;
		private C1.Data.C1DataSet cdsCliente;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtPago;
		private Infragistics.Win.Misc.UltraButton btSubir;
		private Infragistics.Win.Misc.UltraButton btSubirTarjeta;
		private Infragistics.Win.UltraWinEditors.UltraCheckEditor chkVerifica;
		private Infragistics.Win.Misc.UltraButton btTarjeta;
		private Infragistics.Win.Misc.UltraButton btSubirRecarga;
		private Infragistics.Win.Misc.UltraButton btRecarga;
		private Infragistics.Win.Misc.UltraButton btAmortizacion;
		private Infragistics.Win.Misc.UltraButton btSubirAmort;
		private Infragistics.Win.Misc.UltraButton btEspejo;
		private Infragistics.Win.Misc.UltraButton btArchivo;
		private System.Windows.Forms.OpenFileDialog openFileDialog;
		private Infragistics.Win.Misc.UltraButton btSubirEspejo;
		private Infragistics.Win.Misc.UltraButton btEspejoCobro;
		private Infragistics.Win.Misc.UltraButton btSubirP;
		/// <summary>
		/// Variable del diseñador requerida.
		/// </summary>
		private System.ComponentModel.Container components = null;

		public QueBuenaIdea()
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
			this.btCliente = new Infragistics.Win.Misc.UltraButton();
			this.ultraGrid1 = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.cdsCliente = new C1.Data.C1DataSet();
			this.txtPago = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.btSubir = new Infragistics.Win.Misc.UltraButton();
			this.btTarjeta = new Infragistics.Win.Misc.UltraButton();
			this.btSubirTarjeta = new Infragistics.Win.Misc.UltraButton();
			this.chkVerifica = new Infragistics.Win.UltraWinEditors.UltraCheckEditor();
			this.btSubirRecarga = new Infragistics.Win.Misc.UltraButton();
			this.btRecarga = new Infragistics.Win.Misc.UltraButton();
			this.btAmortizacion = new Infragistics.Win.Misc.UltraButton();
			this.btSubirAmort = new Infragistics.Win.Misc.UltraButton();
			this.btEspejo = new Infragistics.Win.Misc.UltraButton();
			this.btArchivo = new Infragistics.Win.Misc.UltraButton();
			this.openFileDialog = new System.Windows.Forms.OpenFileDialog();
			this.btSubirEspejo = new Infragistics.Win.Misc.UltraButton();
			this.btEspejoCobro = new Infragistics.Win.Misc.UltraButton();
			this.btSubirP = new Infragistics.Win.Misc.UltraButton();
			((System.ComponentModel.ISupportInitialize)(this.ultraGrid1)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsCliente)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtPago)).BeginInit();
			this.SuspendLayout();
			// 
			// btCliente
			// 
			this.btCliente.Location = new System.Drawing.Point(32, 16);
			this.btCliente.Name = "btCliente";
			this.btCliente.Size = new System.Drawing.Size(72, 32);
			this.btCliente.TabIndex = 0;
			this.btCliente.Text = "Cliente";
			this.btCliente.Visible = false;
			this.btCliente.Click += new System.EventHandler(this.btCliente_Click);
			// 
			// ultraGrid1
			// 
			this.ultraGrid1.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.ultraGrid1.Cursor = System.Windows.Forms.Cursors.Default;
			this.ultraGrid1.Location = new System.Drawing.Point(24, 104);
			this.ultraGrid1.Name = "ultraGrid1";
			this.ultraGrid1.Size = new System.Drawing.Size(648, 200);
			this.ultraGrid1.TabIndex = 1;
			// 
			// cdsCliente
			// 
			this.cdsCliente.BindingContextCtrl = this;
			this.cdsCliente.DataLibrary = "LibFacturacion";
			this.cdsCliente.DataLibraryUrl = "";
			this.cdsCliente.DataSetDef = "dsSeteoF";
			this.cdsCliente.Locale = new System.Globalization.CultureInfo("es-EC");
			this.cdsCliente.Name = "cdsCliente";
			this.cdsCliente.SchemaClassName = "LibFacturacion.DataClass";
			this.cdsCliente.SchemaDef = null;
			this.cdsCliente.BeforeFill += new C1.Data.FillEventHandler(this.cdsCliente_BeforeFill);
			// 
			// txtPago
			// 
			this.txtPago.Location = new System.Drawing.Point(552, 56);
			this.txtPago.Name = "txtPago";
			this.txtPago.Size = new System.Drawing.Size(80, 21);
			this.txtPago.TabIndex = 2;
			this.txtPago.Text = "c:\\Idea.xls";
			// 
			// btSubir
			// 
			this.btSubir.Location = new System.Drawing.Point(32, 56);
			this.btSubir.Name = "btSubir";
			this.btSubir.Size = new System.Drawing.Size(72, 32);
			this.btSubir.TabIndex = 3;
			this.btSubir.Text = "Subir  Cl.";
			this.btSubir.Visible = false;
			this.btSubir.Click += new System.EventHandler(this.ultraButton1_Click);
			// 
			// btTarjeta
			// 
			this.btTarjeta.Location = new System.Drawing.Point(120, 16);
			this.btTarjeta.Name = "btTarjeta";
			this.btTarjeta.Size = new System.Drawing.Size(72, 32);
			this.btTarjeta.TabIndex = 4;
			this.btTarjeta.Text = "Tarjeta";
			this.btTarjeta.Click += new System.EventHandler(this.btFactura_Click);
			// 
			// btSubirTarjeta
			// 
			this.btSubirTarjeta.Location = new System.Drawing.Point(120, 56);
			this.btSubirTarjeta.Name = "btSubirTarjeta";
			this.btSubirTarjeta.Size = new System.Drawing.Size(72, 32);
			this.btSubirTarjeta.TabIndex = 5;
			this.btSubirTarjeta.Text = "Subir T.";
			this.btSubirTarjeta.Click += new System.EventHandler(this.btSubirTarjeta_Click);
			// 
			// chkVerifica
			// 
			this.chkVerifica.Checked = true;
			this.chkVerifica.CheckState = System.Windows.Forms.CheckState.Checked;
			this.chkVerifica.Location = new System.Drawing.Point(584, 24);
			this.chkVerifica.Name = "chkVerifica";
			this.chkVerifica.Size = new System.Drawing.Size(64, 20);
			this.chkVerifica.TabIndex = 6;
			this.chkVerifica.Text = "Verifica";
			// 
			// btSubirRecarga
			// 
			this.btSubirRecarga.Location = new System.Drawing.Point(208, 56);
			this.btSubirRecarga.Name = "btSubirRecarga";
			this.btSubirRecarga.Size = new System.Drawing.Size(72, 32);
			this.btSubirRecarga.TabIndex = 8;
			this.btSubirRecarga.Text = "Subir R.";
			this.btSubirRecarga.Visible = false;
			this.btSubirRecarga.Click += new System.EventHandler(this.btSubirRecarga_Click);
			// 
			// btRecarga
			// 
			this.btRecarga.Location = new System.Drawing.Point(208, 16);
			this.btRecarga.Name = "btRecarga";
			this.btRecarga.Size = new System.Drawing.Size(72, 32);
			this.btRecarga.TabIndex = 7;
			this.btRecarga.Text = "Recarga";
			this.btRecarga.Visible = false;
			this.btRecarga.Click += new System.EventHandler(this.btRecarga_Click);
			// 
			// btAmortizacion
			// 
			this.btAmortizacion.Location = new System.Drawing.Point(296, 16);
			this.btAmortizacion.Name = "btAmortizacion";
			this.btAmortizacion.Size = new System.Drawing.Size(72, 32);
			this.btAmortizacion.TabIndex = 9;
			this.btAmortizacion.Text = "Amortiza";
			this.btAmortizacion.Click += new System.EventHandler(this.btAmortizacion_Click);
			// 
			// btSubirAmort
			// 
			this.btSubirAmort.Location = new System.Drawing.Point(296, 56);
			this.btSubirAmort.Name = "btSubirAmort";
			this.btSubirAmort.Size = new System.Drawing.Size(72, 32);
			this.btSubirAmort.TabIndex = 10;
			this.btSubirAmort.Text = "Subir A.";
			this.btSubirAmort.Click += new System.EventHandler(this.btSubirAmort_Click);
			// 
			// btEspejo
			// 
			this.btEspejo.Location = new System.Drawing.Point(384, 16);
			this.btEspejo.Name = "btEspejo";
			this.btEspejo.Size = new System.Drawing.Size(72, 32);
			this.btEspejo.TabIndex = 11;
			this.btEspejo.Text = "Espejo";
			this.btEspejo.Click += new System.EventHandler(this.btEspejo_Click);
			// 
			// btArchivo
			// 
			this.btArchivo.Location = new System.Drawing.Point(648, 56);
			this.btArchivo.Name = "btArchivo";
			this.btArchivo.Size = new System.Drawing.Size(24, 24);
			this.btArchivo.TabIndex = 12;
			this.btArchivo.Text = "...";
			this.btArchivo.Click += new System.EventHandler(this.btArchivo_Click);
			// 
			// btSubirEspejo
			// 
			this.btSubirEspejo.Location = new System.Drawing.Point(384, 56);
			this.btSubirEspejo.Name = "btSubirEspejo";
			this.btSubirEspejo.Size = new System.Drawing.Size(72, 32);
			this.btSubirEspejo.TabIndex = 13;
			this.btSubirEspejo.Text = "Subir E.";
			this.btSubirEspejo.Click += new System.EventHandler(this.btSubirEspejo_Click);
			// 
			// btEspejoCobro
			// 
			this.btEspejoCobro.Location = new System.Drawing.Point(472, 16);
			this.btEspejoCobro.Name = "btEspejoCobro";
			this.btEspejoCobro.Size = new System.Drawing.Size(72, 32);
			this.btEspejoCobro.TabIndex = 14;
			this.btEspejoCobro.Text = "Espejo";
			this.btEspejoCobro.Click += new System.EventHandler(this.btEspejoCobro_Click);
			// 
			// btSubirP
			// 
			this.btSubirP.Location = new System.Drawing.Point(472, 56);
			this.btSubirP.Name = "btSubirP";
			this.btSubirP.Size = new System.Drawing.Size(72, 32);
			this.btSubirP.TabIndex = 15;
			this.btSubirP.Text = "Subir P.";
			this.btSubirP.Click += new System.EventHandler(this.btSubirP_Click);
			// 
			// QueBuenaIdea
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 13);
			this.ClientSize = new System.Drawing.Size(688, 326);
			this.Controls.Add(this.btSubirP);
			this.Controls.Add(this.btEspejoCobro);
			this.Controls.Add(this.btSubirEspejo);
			this.Controls.Add(this.btArchivo);
			this.Controls.Add(this.btEspejo);
			this.Controls.Add(this.btSubirAmort);
			this.Controls.Add(this.btAmortizacion);
			this.Controls.Add(this.btSubirRecarga);
			this.Controls.Add(this.btRecarga);
			this.Controls.Add(this.chkVerifica);
			this.Controls.Add(this.btSubirTarjeta);
			this.Controls.Add(this.btTarjeta);
			this.Controls.Add(this.btSubir);
			this.Controls.Add(this.txtPago);
			this.Controls.Add(this.ultraGrid1);
			this.Controls.Add(this.btCliente);
			this.Name = "QueBuenaIdea";
			this.Text = "Que Buena Idea";
			this.Load += new System.EventHandler(this.QueBuenaIdea_Load);
			((System.ComponentModel.ISupportInitialize)(this.ultraGrid1)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsCliente)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtPago)).EndInit();
			this.ResumeLayout(false);

		}
		#endregion

		private void btCliente_Click(object sender, System.EventArgs e)
		{
			string strConn ="Provider=Microsoft.Jet.OLEDB.4.0;" +
				"Data Source=" + txtPago.Text + ";" +
				"Extended Properties='Excel 8.0;'";
			OleDbDataAdapter myData = new OleDbDataAdapter("SELECT * FROM [A$]", strConn);

			myData.TableMappings.Add("Tabla", "ExcelSube");
			miDataSet.Reset();
			miDataSet.Clear();
			try
			{
				myData.Fill(miDataSet);
			}
			catch(Exception ex)
			{
				MessageBox.Show(ex.Message, "Error al abrir la Hoja A del Archivo: " + txtPago.Text);
				return;
			}
			ultraGrid1.DataSource = miDataSet.Tables[0].DefaultView;
			ultraGrid1.DataBind();
		}

		DataSet miDataSet = new DataSet();
		private void QueBuenaIdea_Load(object sender, System.EventArgs e)
		{
		
		}

		private void ultraButton1_Click(object sender, System.EventArgs e)
		{
			foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.ultraGrid1.Rows.All)
			{
				try
				{
					if (dr.Cells[0].Value == DBNull.Value) continue;
					string stNombre = dr.Cells[2].Value.ToString().Trim();
					string stCodigo = dr.Cells[7].Value.ToString().Trim();
					string stExec = "Exec ClienteCodigoCrear '" + stNombre.Trim()
						+ "', '" + stCodigo + "'";
					Funcion.EjecutaSQL(cdsCliente, stExec, true);
				}
				catch(System.Exception ex)
				{
					MessageBox.Show(ex.Message, "Información");
				}
			}
			MessageBox.Show("Clientes subidos", "Información");
		}

		private void cdsCliente_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			cdsCliente.Schema.Connections[0].ConnectionString = 
				miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);
		}

		private void btFactura_Click(object sender, System.EventArgs e)
		{
			string strConn ="Provider=Microsoft.Jet.OLEDB.4.0;" +
				"Data Source=" + txtPago.Text + ";" +
				"Extended Properties='Excel 8.0;'";
			OleDbDataAdapter myData = new OleDbDataAdapter("SELECT * FROM [T$]", strConn);
			myData.TableMappings.Add("Tabla", "ExcelSube");
			miDataSet.Reset();
			miDataSet.Clear();
			try
			{
				myData.Fill(miDataSet);
			}
			catch(Exception ex)
			{
				MessageBox.Show(ex.Message, "Error al abrir la Hoja T del Archivo: " + txtPago.Text);
				return;
			}
			ultraGrid1.DataSource = miDataSet.Tables[0].DefaultView;
			ultraGrid1.DataBind();
		}

		private void btSubirTarjeta_Click(object sender, System.EventArgs e)
		{
			string stCodAntes = "";
			int IdCompra = 0;
			int idCompraAntes = Funcion.iEscalarSQL(cdsCliente, 
				"Select Top 1 idCompra From Compra Order by idCompra Desc", true);
		
			Funcion.EjecutaSQL(cdsCliente, "Delete ErrorSalida", true);
			string stVerifica = ", 0";
			if (chkVerifica.Checked) stVerifica = ", 1";
			#region Recorrido
			foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.ultraGrid1.Rows.All)
			{
				try
				{
					if (dr.Cells[0].Value == DBNull.Value) continue;
					string stCodigoCliente = dr.Cells[4].Value.ToString().Trim();
					string stCodigoVendedor = dr.Cells[5].Value.ToString().Trim();
					DateTime dtFecha = DateTime.Parse(dr.Cells[7].Value.ToString().Trim());
					if (stCodigoCliente != stCodAntes)
					{
						string stSelect = "Select Top 1 idCompra From Compra "
							+ "Where idTipoFactura = 1 "
							+ "Order by idCompra Desc";
						idCompraAntes = Funcion.iEscalarSQL(cdsCliente, stSelect, true);
						// Crea Factura Encabezado
						string stExec = "Exec QBIFactura '" + stCodigoCliente.Trim()
							+ "', '" + stCodigoVendedor + "', '" 
							+ dtFecha.ToString("yyyyMMdd") + "'" + stVerifica;
						Funcion.EjecutaSQL(cdsCliente, stExec, true);
						IdCompra = Funcion.iEscalarSQL(cdsCliente, stSelect, true);
						if (idCompraAntes == IdCompra) continue;
					}
					stCodAntes = stCodigoCliente;
					int iCantidad = int.Parse(dr.Cells[8].Value.ToString());
					string stDescripcion = dr.Cells[9].Value.ToString().Trim();
					string stLinea = "Exec QBIFacturaLinea " + IdCompra.ToString() + ", " +
						iCantidad.ToString() + ", '" + stDescripcion + "'" + stVerifica;
					Funcion.EjecutaSQL(cdsCliente, stLinea, true);
				}
				catch(System.Exception ex)
				{
					MessageBox.Show(ex.Message, "Información");
				}
			}		
			#endregion Recorrido
			if (chkVerifica.Checked) MuestraErrores();
			else
			{
				MessageBox.Show("Tarjetas subidas", "Información");
				string stFact = "Exec FacturaTotales " + idCompraAntes.ToString();
				Funcion.EjecutaSQL(cdsCliente, stFact, true);
			}		
		}

		private void MuestraErrores()
		{
			TesterError miError = new TesterError();
			miError.MdiParent = this.MdiParent;
			miError.Show();
		}

		private void btSubirTarjeta_MouseDown(object sender, System.Windows.Forms.MouseEventArgs e)
		{
			MuestraErrores();
		}

		private void btRecarga_Click(object sender, System.EventArgs e)
		{
			string strConn ="Provider=Microsoft.Jet.OLEDB.4.0;" +
				"Data Source=" + txtPago.Text + ";" +
				"Extended Properties='Excel 8.0;'";
			OleDbDataAdapter myData = new OleDbDataAdapter("SELECT * FROM [R$] Where [Nodo_Origen] = 31031", strConn);
			myData.TableMappings.Add("Tabla", "ExcelSube");
			miDataSet.Reset();
			miDataSet.Clear();
			try
			{
				myData.Fill(miDataSet);
			}
			catch(Exception ex)
			{
				MessageBox.Show(ex.Message, "Error al abrir la Hoja R del Archivo: " + txtPago.Text);
				return;
			}
			ultraGrid1.DataSource = miDataSet.Tables[0].DefaultView;
			ultraGrid1.DataBind();		
		}

		private void btSubirRecarga_Click(object sender, System.EventArgs e)
		{
			string stCodAntes = "";
			int IdCompra = 0;
			int idCompraAntes = 0;
			Funcion.EjecutaSQL(cdsCliente, "Delete ErrorSalida", true);
			string stVerifica = ", 0";
			if (chkVerifica.Checked) stVerifica = ", 1";
			#region Recorrido
			foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.ultraGrid1.Rows.All)
			{
				try
				{
					if (dr.Cells[0].Value == DBNull.Value) continue;
					string stCodigoCliente = dr.Cells[4].Value.ToString().Trim();
					string stCodigoVendedor = dr.Cells[4].Value.ToString().Trim();
					DateTime dtFecha = DateTime.Parse(dr.Cells[10].Value.ToString().Trim());
					if (stCodigoCliente != stCodAntes)
					{
						string stSelect = "Select Top 1 idCompra From Compra "
							+ "Where idTipoFactura = 1 "
							+ "Order by idCompra Desc";
						idCompraAntes = Funcion.iEscalarSQL(cdsCliente, stSelect, true);
						// Crea Factura Encabezado
						string stExec = "Exec QBIFactura '" + stCodigoCliente.Trim()
							+ "', '" + stCodigoVendedor + "', '" 
							+ dtFecha.ToString("yyyyMMdd") + "'" + stVerifica;
						Funcion.EjecutaSQL(cdsCliente, stExec, true);
						IdCompra = Funcion.iEscalarSQL(cdsCliente, stSelect, true);
						if (idCompraAntes == IdCompra) continue;
					}
					stCodAntes = stCodigoCliente;
					int iCantidad = int.Parse(dr.Cells[8].Value.ToString());
					string stDescripcion = dr.Cells[9].Value.ToString().Trim();
					string stLinea = "Exec QBIFacturaLinea " + IdCompra.ToString() + ", " +
						iCantidad.ToString() + ", '" + stDescripcion + "'" + stVerifica;
					Funcion.EjecutaSQL(cdsCliente, stLinea, true);
				}
				catch(System.Exception ex)
				{
					MessageBox.Show(ex.Message, "Información");
				}
			}		
			#endregion Recorrido
			if (chkVerifica.Checked) MuestraErrores();
			else MessageBox.Show("Recarga subida", "Información");
		}

		CultureInfo us = new CultureInfo("en-US");
		private void btSubirAmort_Click(object sender, System.EventArgs e)
		{
			Funcion.EjecutaSQL(cdsCliente, "Delete ErrorSalida", true);
			string stVerifica = ", 0";
			if (chkVerifica.Checked) stVerifica = ", 1";
			#region Recorrido
			foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.ultraGrid1.Rows.All)
			{
				try
				{
					if (dr.Cells[0].Value == DBNull.Value) continue;
					string stNumero = dr.Cells[0].Value.ToString().Trim();
					DateTime dtFecha = DateTime.Parse(dr.Cells[1].Value.ToString().Trim());
					double dCapital = double.Parse(dr.Cells[2].Value.ToString().Trim());
					double dInteres = double.Parse(dr.Cells[3].Value.ToString().Trim());
					DateTime dtFechaP = DateTime.Parse(dr.Cells[4].Value.ToString().Trim());
					int iCuota = int.Parse(dr.Cells[5].Value.ToString().Trim());
					int iPendiente = int.Parse(dr.Cells[6].Value.ToString().Trim());

					string stLinea = "Exec AmortizacionImportar '" + stNumero + "', '"
						+ dtFecha.ToString("yyyyMMdd") + "', " + dCapital.ToString("0.00", us) 
						+ ", " + dInteres.ToString("0.00", us) + ", '" 
						+ dtFechaP.ToString("yyyyMMdd") + "', " + iCuota.ToString() 
						+ ", " + iPendiente.ToString()
						+ stVerifica;
					Funcion.EjecutaSQL(cdsCliente, stLinea, true);
				}
				catch(System.Exception ex)
				{
					MessageBox.Show(ex.Message, "Información");
				}
			}			
			#endregion Recorrido
			if (chkVerifica.Checked) MuestraErrores();
			else MessageBox.Show("Amortizacion subido", "Información");
		}

		private void btAmortizacion_Click(object sender, System.EventArgs e)
		{
			string strConn ="Provider=Microsoft.Jet.OLEDB.4.0;" +
				"Data Source=" + txtPago.Text + ";" +
				"Extended Properties='Excel 8.0;'";
			OleDbDataAdapter myData = new OleDbDataAdapter("SELECT * FROM [Hoja1$]", strConn);
			myData.TableMappings.Add("Tabla", "ExcelSube");
			miDataSet.Reset();
			miDataSet.Clear();
			try
			{
				myData.Fill(miDataSet);
			}
			catch(Exception ex)
			{
				MessageBox.Show(ex.Message, "Error al abrir la Hoja1 del Archivo: " + txtPago.Text);
				return;
			}
			ultraGrid1.DataSource = miDataSet.Tables[0].DefaultView;
			ultraGrid1.DataBind();		
		}

		private void btEspejo_Click(object sender, System.EventArgs e)
		{
			string strConn ="Provider=Microsoft.Jet.OLEDB.4.0;" +
				"Data Source=" + txtPago.Text + ";" +
				"Extended Properties='Excel 8.0;'";
			OleDbDataAdapter myData = new OleDbDataAdapter("SELECT * FROM [Detalle$]", strConn);
			myData.TableMappings.Add("Tabla", "ExcelSube");
			miDataSet.Reset();
			miDataSet.Clear();
			try
			{
				myData.Fill(miDataSet);
			}
			catch(Exception ex)
			{
				MessageBox.Show(ex.Message, "Error al abrir la Detalle del Archivo: " + txtPago.Text);
				return;
			}
			ultraGrid1.DataSource = miDataSet.Tables[0].DefaultView;
			ultraGrid1.DataBind();
		}

		private void btArchivo_Click(object sender, System.EventArgs e)
		{
			this.openFileDialog.Filter = "Excel (*.xls)|*.xls";
			this.openFileDialog.InitialDirectory = MenuLatinium.stDirInicio;
			if (DialogResult.OK == this.openFileDialog.ShowDialog())
			{
				string stArchivo = this.openFileDialog.FileName.ToString();
				txtPago.Text = stArchivo;
			}
		}

		private void btSubirEspejo_Click(object sender, System.EventArgs e)
		{
			Funcion.EjecutaSQL(cdsCliente, "Delete ErrorSalida", true);
			string stVerifica = ", 0";
			if (chkVerifica.Checked) stVerifica = ", 1";
			#region Recorrido
			foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.ultraGrid1.Rows.All)
			{
				try
				{
					if (dr.Cells[0].Value == DBNull.Value) continue;
					string stNumero = dr.Cells[0].Value.ToString().Trim();
					string stArticulo = dr.Cells[1].Value.ToString().Trim();
					double dTotal = double.Parse(dr.Cells[2].Value.ToString().Trim());
					double dMulta = double.Parse(dr.Cells[3].Value.ToString().Trim());
					double dCantidad = double.Parse(dr.Cells[4].Value.ToString().Trim());

					string stLinea = "Exec EspejoDetalle '" + stNumero + "', '"
						+ stArticulo + "', " + dTotal.ToString("0.00", us) 
						+ ", " + dMulta.ToString("0.00", us) + ", " 
						+ dCantidad.ToString("0.00", us) 
						+ stVerifica;
					Funcion.EjecutaSQL(cdsCliente, stLinea, true);
				}
				catch(System.Exception ex)
				{
					MessageBox.Show(ex.Message, "Información");
				}
			}			
			#endregion Recorrido
			if (chkVerifica.Checked) MuestraErrores();
			else MessageBox.Show("Detalle de compra subido", "Información");
		}

		private void btEspejoCobro_Click(object sender, System.EventArgs e)
		{
			string strConn ="Provider=Microsoft.Jet.OLEDB.4.0;" +
				"Data Source=" + txtPago.Text + ";" +
				"Extended Properties='Excel 8.0;'";
			OleDbDataAdapter myData = new OleDbDataAdapter("SELECT * FROM [Mov Cuotas 2008-2009$]", strConn);
			myData.TableMappings.Add("Tabla", "ExcelSube");
			miDataSet.Reset();
			miDataSet.Clear();
			try
			{
				myData.Fill(miDataSet);
			}
			catch(Exception ex)
			{
				MessageBox.Show(ex.Message, "Error al abrir la Detalle del Archivo: " + txtPago.Text);
				return;
			}
			ultraGrid1.DataSource = miDataSet.Tables[0].DefaultView;
			ultraGrid1.DataBind();
		}

		private void btSubirP_Click(object sender, System.EventArgs e)
		{
			Funcion.EjecutaSQL(cdsCliente, "Delete ErrorSalida", true);
			string stVerifica = "', 0";
			if (chkVerifica.Checked) stVerifica = "', 1";
			#region Recorrido
			foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.ultraGrid1.Rows.All)
			{
				try
				{
					if (dr.Cells[0].Value == DBNull.Value) continue;
					string stVerdad = dr.Cells[4].Value.ToString().Trim();
					if (stVerdad == "True")
					{
						string stNumero = dr.Cells[8].Value.ToString().Trim();
						DateTime dtFecha = (DateTime) dr.Cells[3].Value;
						double dTotal = double.Parse(dr.Cells[7].Value.ToString().Trim());
						double dMulta = double.Parse(dr.Cells[13].Value.ToString().Trim());
						DateTime dtFechaFin = (DateTime) dr.Cells[17].Value;

						string stLinea = "Exec EspejoPago '" + stNumero + "', '"
							+ dtFecha.ToString("yyyyMMdd") + "', " + dTotal.ToString("0.00", us) 
							+ ", " + dMulta.ToString("0.00", us) + ", '" 
							+ dtFechaFin.ToString("yyyyMMdd") 
							+ stVerifica;
						Funcion.EjecutaSQL(cdsCliente, stLinea, true);
					}
				}
				catch(System.Exception ex)
				{
					MessageBox.Show(ex.Message, "Información");
				}
			}			
			#endregion Recorrido
			if (chkVerifica.Checked) MuestraErrores();
			else MessageBox.Show("Detalle de compra subido", "Información");
		}
	}
}
