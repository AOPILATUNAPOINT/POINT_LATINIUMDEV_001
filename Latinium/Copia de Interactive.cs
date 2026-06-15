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
	/// Descripción breve de TelcoSan.
	/// </summary>
	public class Interactive : DevExpress.XtraEditors.XtraForm
	{
		private Infragistics.Win.Misc.UltraButton ultraButton1;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtFile;
		private Infragistics.Win.UltraWinGrid.UltraGrid ultraGrid1;
		private Infragistics.Win.Misc.UltraButton btLee;
		private C1.Data.C1DataSet cdsCliente;
		/// <summary>
		/// Variable del diseñador requerida.
		/// </summary>
		private System.ComponentModel.Container components = null;
		private Infragistics.Win.Misc.UltraButton ultraButton2;
		private Infragistics.Win.Misc.UltraButton ultraButton3;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtPago;
		CultureInfo us = new CultureInfo("en-US");
		private Infragistics.Win.UltraWinEditors.UltraCheckEditor chkCreaCliente;
		private Infragistics.Win.Misc.UltraButton btErrores;
		private Infragistics.Win.UltraWinEditors.UltraCheckEditor chkBorrarError;
		private Infragistics.Win.Misc.UltraButton btAsiPago;
		private Infragistics.Win.Misc.UltraButton btAsiFact;
		private Infragistics.Win.Misc.UltraButton btExcelRet;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtRetencion;
		private Infragistics.Win.Misc.UltraButton btRetencion;
		private Infragistics.Win.Misc.UltraButton btVerificaRet;
		private Infragistics.Win.Misc.UltraButton btAsiRet;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtAsiento;
		private Infragistics.Win.Misc.UltraButton btFactRet;
		private Infragistics.Win.Misc.UltraButton btReservaPago;
		private Infragistics.Win.UltraWinEditors.UltraCheckEditor chkClienteFactPago;
		private Infragistics.Win.Misc.UltraButton btExcelFactRapida;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtFactRapida;
		CultureInfo es = new CultureInfo("es");

		public Interactive()
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
			Infragistics.Win.Appearance appearance1 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance2 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance3 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance4 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance5 = new Infragistics.Win.Appearance();
			this.ultraButton1 = new Infragistics.Win.Misc.UltraButton();
			this.txtFile = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.ultraGrid1 = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.btLee = new Infragistics.Win.Misc.UltraButton();
			this.cdsCliente = new C1.Data.C1DataSet();
			this.ultraButton2 = new Infragistics.Win.Misc.UltraButton();
			this.txtPago = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.ultraButton3 = new Infragistics.Win.Misc.UltraButton();
			this.chkCreaCliente = new Infragistics.Win.UltraWinEditors.UltraCheckEditor();
			this.btErrores = new Infragistics.Win.Misc.UltraButton();
			this.chkBorrarError = new Infragistics.Win.UltraWinEditors.UltraCheckEditor();
			this.btAsiPago = new Infragistics.Win.Misc.UltraButton();
			this.btAsiFact = new Infragistics.Win.Misc.UltraButton();
			this.btExcelRet = new Infragistics.Win.Misc.UltraButton();
			this.txtRetencion = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.btRetencion = new Infragistics.Win.Misc.UltraButton();
			this.btVerificaRet = new Infragistics.Win.Misc.UltraButton();
			this.btAsiRet = new Infragistics.Win.Misc.UltraButton();
			this.txtAsiento = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.btFactRet = new Infragistics.Win.Misc.UltraButton();
			this.btReservaPago = new Infragistics.Win.Misc.UltraButton();
			this.chkClienteFactPago = new Infragistics.Win.UltraWinEditors.UltraCheckEditor();
			this.btExcelFactRapida = new Infragistics.Win.Misc.UltraButton();
			this.txtFactRapida = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			((System.ComponentModel.ISupportInitialize)(this.txtFile)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraGrid1)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsCliente)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtPago)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtRetencion)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtAsiento)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtFactRapida)).BeginInit();
			this.SuspendLayout();
			// 
			// ultraButton1
			// 
			this.ultraButton1.Location = new System.Drawing.Point(40, 40);
			this.ultraButton1.Name = "ultraButton1";
			this.ultraButton1.TabIndex = 0;
			this.ultraButton1.Text = "Excel Fact.";
			this.ultraButton1.Click += new System.EventHandler(this.ultraButton1_Click);
			// 
			// txtFile
			// 
			this.txtFile.Location = new System.Drawing.Point(136, 40);
			this.txtFile.Name = "txtFile";
			this.txtFile.Size = new System.Drawing.Size(128, 21);
			this.txtFile.TabIndex = 1;
			this.txtFile.Text = "C:\\Factura.Xls";
			// 
			// ultraGrid1
			// 
			this.ultraGrid1.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.ultraGrid1.Cursor = System.Windows.Forms.Cursors.Default;
			appearance1.BackColor = System.Drawing.Color.White;
			this.ultraGrid1.DisplayLayout.Appearance = appearance1;
			appearance2.BackColor = System.Drawing.Color.Transparent;
			this.ultraGrid1.DisplayLayout.Override.CardAreaAppearance = appearance2;
			appearance3.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance3.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance3.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance3.FontData.BoldAsString = "True";
			appearance3.FontData.Name = "Arial";
			appearance3.FontData.SizeInPoints = 10F;
			appearance3.ForeColor = System.Drawing.Color.White;
			appearance3.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.ultraGrid1.DisplayLayout.Override.HeaderAppearance = appearance3;
			appearance4.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance4.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance4.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ultraGrid1.DisplayLayout.Override.RowSelectorAppearance = appearance4;
			appearance5.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance5.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance5.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ultraGrid1.DisplayLayout.Override.SelectedRowAppearance = appearance5;
			this.ultraGrid1.DisplayLayout.ViewStyleBand = Infragistics.Win.UltraWinGrid.ViewStyleBand.OutlookGroupBy;
			this.ultraGrid1.Location = new System.Drawing.Point(8, 152);
			this.ultraGrid1.Name = "ultraGrid1";
			this.ultraGrid1.Size = new System.Drawing.Size(472, 128);
			this.ultraGrid1.TabIndex = 2;
			// 
			// btLee
			// 
			this.btLee.Location = new System.Drawing.Point(288, 40);
			this.btLee.Name = "btLee";
			this.btLee.TabIndex = 3;
			this.btLee.Text = "Facturas";
			this.btLee.Click += new System.EventHandler(this.btLee_Click);
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
			// ultraButton2
			// 
			this.ultraButton2.Location = new System.Drawing.Point(288, 8);
			this.ultraButton2.Name = "ultraButton2";
			this.ultraButton2.TabIndex = 9;
			this.ultraButton2.Text = "Pago";
			this.ultraButton2.Click += new System.EventHandler(this.ultraButton2_Click);
			// 
			// txtPago
			// 
			this.txtPago.Location = new System.Drawing.Point(136, 8);
			this.txtPago.Name = "txtPago";
			this.txtPago.Size = new System.Drawing.Size(128, 21);
			this.txtPago.TabIndex = 8;
			this.txtPago.Text = "C:\\Cobro.Xls";
			// 
			// ultraButton3
			// 
			this.ultraButton3.Location = new System.Drawing.Point(40, 8);
			this.ultraButton3.Name = "ultraButton3";
			this.ultraButton3.TabIndex = 7;
			this.ultraButton3.Text = "Excel Pago";
			this.ultraButton3.Click += new System.EventHandler(this.ultraButton3_Click);
			// 
			// chkCreaCliente
			// 
			this.chkCreaCliente.Location = new System.Drawing.Point(280, 104);
			this.chkCreaCliente.Name = "chkCreaCliente";
			this.chkCreaCliente.Size = new System.Drawing.Size(96, 20);
			this.chkCreaCliente.TabIndex = 10;
			this.chkCreaCliente.Text = "Crea Cliente";
			// 
			// btErrores
			// 
			this.btErrores.Location = new System.Drawing.Point(136, 104);
			this.btErrores.Name = "btErrores";
			this.btErrores.Size = new System.Drawing.Size(64, 23);
			this.btErrores.TabIndex = 11;
			this.btErrores.Text = "Errores";
			this.btErrores.Click += new System.EventHandler(this.btErrores_Click);
			// 
			// chkBorrarError
			// 
			this.chkBorrarError.Checked = true;
			this.chkBorrarError.CheckState = System.Windows.Forms.CheckState.Checked;
			this.chkBorrarError.Location = new System.Drawing.Point(40, 104);
			this.chkBorrarError.Name = "chkBorrarError";
			this.chkBorrarError.Size = new System.Drawing.Size(88, 20);
			this.chkBorrarError.TabIndex = 12;
			this.chkBorrarError.Text = "Borrar Error";
			// 
			// btAsiPago
			// 
			this.btAsiPago.Location = new System.Drawing.Point(392, 8);
			this.btAsiPago.Name = "btAsiPago";
			this.btAsiPago.TabIndex = 13;
			this.btAsiPago.Text = "Asi. Pago";
			this.btAsiPago.Click += new System.EventHandler(this.btAsiPago_Click);
			// 
			// btAsiFact
			// 
			this.btAsiFact.Location = new System.Drawing.Point(392, 40);
			this.btAsiFact.Name = "btAsiFact";
			this.btAsiFact.TabIndex = 14;
			this.btAsiFact.Text = "Asi. Fact.";
			this.btAsiFact.Click += new System.EventHandler(this.btAsiFact_Click);
			// 
			// btExcelRet
			// 
			this.btExcelRet.Location = new System.Drawing.Point(40, 72);
			this.btExcelRet.Name = "btExcelRet";
			this.btExcelRet.TabIndex = 15;
			this.btExcelRet.Text = "Excel Ret.";
			this.btExcelRet.Click += new System.EventHandler(this.btExcelRet_Click);
			// 
			// txtRetencion
			// 
			this.txtRetencion.Location = new System.Drawing.Point(136, 72);
			this.txtRetencion.Name = "txtRetencion";
			this.txtRetencion.Size = new System.Drawing.Size(128, 21);
			this.txtRetencion.TabIndex = 16;
			this.txtRetencion.Text = "C:\\Retencion.Xls";
			// 
			// btRetencion
			// 
			this.btRetencion.Location = new System.Drawing.Point(392, 72);
			this.btRetencion.Name = "btRetencion";
			this.btRetencion.Size = new System.Drawing.Size(72, 23);
			this.btRetencion.TabIndex = 17;
			this.btRetencion.Text = "Retencion";
			this.btRetencion.Click += new System.EventHandler(this.btRetencion_Click);
			// 
			// btVerificaRet
			// 
			this.btVerificaRet.Location = new System.Drawing.Point(264, 48);
			this.btVerificaRet.Name = "btVerificaRet";
			this.btVerificaRet.Size = new System.Drawing.Size(32, 23);
			this.btVerificaRet.TabIndex = 18;
			this.btVerificaRet.Text = "Verifica";
			this.btVerificaRet.Click += new System.EventHandler(this.btVerificaRet_Click);
			// 
			// btAsiRet
			// 
			this.btAsiRet.Location = new System.Drawing.Point(208, 104);
			this.btAsiRet.Name = "btAsiRet";
			this.btAsiRet.Size = new System.Drawing.Size(56, 23);
			this.btAsiRet.TabIndex = 19;
			this.btAsiRet.Text = "Asi. Ret.";
			this.btAsiRet.Visible = false;
			this.btAsiRet.Click += new System.EventHandler(this.btAsiRet_Click);
			// 
			// txtAsiento
			// 
			this.txtAsiento.Location = new System.Drawing.Point(408, 104);
			this.txtAsiento.Name = "txtAsiento";
			this.txtAsiento.Size = new System.Drawing.Size(64, 21);
			this.txtAsiento.TabIndex = 20;
			// 
			// btFactRet
			// 
			this.btFactRet.Location = new System.Drawing.Point(288, 72);
			this.btFactRet.Name = "btFactRet";
			this.btFactRet.Size = new System.Drawing.Size(72, 23);
			this.btFactRet.TabIndex = 21;
			this.btFactRet.Text = "Fact. Retn.";
			this.btFactRet.Click += new System.EventHandler(this.btFactRet_Click);
			// 
			// btReservaPago
			// 
			this.btReservaPago.Location = new System.Drawing.Point(224, 104);
			this.btReservaPago.Name = "btReservaPago";
			this.btReservaPago.Size = new System.Drawing.Size(48, 23);
			this.btReservaPago.TabIndex = 22;
			this.btReservaPago.Text = "R -> P";
			this.btReservaPago.Click += new System.EventHandler(this.btReservaPago_Click);
			// 
			// chkClienteFactPago
			// 
			this.chkClienteFactPago.Location = new System.Drawing.Point(280, 128);
			this.chkClienteFactPago.Name = "chkClienteFactPago";
			this.chkClienteFactPago.Size = new System.Drawing.Size(184, 20);
			this.chkClienteFactPago.TabIndex = 23;
			this.chkClienteFactPago.Text = "Cliente de la Factura en Pago";
			// 
			// btExcelFactRapida
			// 
			this.btExcelFactRapida.Location = new System.Drawing.Point(40, 128);
			this.btExcelFactRapida.Name = "btExcelFactRapida";
			this.btExcelFactRapida.TabIndex = 24;
			this.btExcelFactRapida.Text = "Fact. Rapida";
			this.btExcelFactRapida.Click += new System.EventHandler(this.ultraButton4_Click);
			this.btExcelFactRapida.MouseDown += new System.Windows.Forms.MouseEventHandler(this.btExcelFactRapida_MouseDown);
			// 
			// txtFactRapida
			// 
			this.txtFactRapida.Location = new System.Drawing.Point(136, 128);
			this.txtFactRapida.Name = "txtFactRapida";
			this.txtFactRapida.Size = new System.Drawing.Size(128, 21);
			this.txtFactRapida.TabIndex = 25;
			this.txtFactRapida.Text = "C:\\FacturaRapida.Xls";
			// 
			// Interactive
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 13);
			this.ClientSize = new System.Drawing.Size(488, 294);
			this.Controls.Add(this.txtFactRapida);
			this.Controls.Add(this.btExcelFactRapida);
			this.Controls.Add(this.chkClienteFactPago);
			this.Controls.Add(this.btReservaPago);
			this.Controls.Add(this.btFactRet);
			this.Controls.Add(this.txtAsiento);
			this.Controls.Add(this.btAsiRet);
			this.Controls.Add(this.btVerificaRet);
			this.Controls.Add(this.btRetencion);
			this.Controls.Add(this.txtRetencion);
			this.Controls.Add(this.btExcelRet);
			this.Controls.Add(this.btAsiFact);
			this.Controls.Add(this.btAsiPago);
			this.Controls.Add(this.chkBorrarError);
			this.Controls.Add(this.btErrores);
			this.Controls.Add(this.chkCreaCliente);
			this.Controls.Add(this.ultraButton2);
			this.Controls.Add(this.txtPago);
			this.Controls.Add(this.ultraButton3);
			this.Controls.Add(this.btLee);
			this.Controls.Add(this.ultraGrid1);
			this.Controls.Add(this.txtFile);
			this.Controls.Add(this.ultraButton1);
			this.Name = "Interactive";
			this.Text = "Interactive";
			this.Closing += new System.ComponentModel.CancelEventHandler(this.Interactive_Closing);
			this.Load += new System.EventHandler(this.TelcoSan_Load);
			((System.ComponentModel.ISupportInitialize)(this.txtFile)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraGrid1)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsCliente)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtPago)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtRetencion)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtAsiento)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtFactRapida)).EndInit();
			this.ResumeLayout(false);

		}
		#endregion

		DataSet miDataSet = new DataSet();
		private void ultraButton1_Click(object sender, System.EventArgs e)
		{
       string strConn ="Provider=Microsoft.Jet.OLEDB.4.0;" +
         "Data Source=" + txtFile.Text + ";" +
         "Extended Properties='Excel 8.0;'";
			OleDbDataAdapter myData = new OleDbDataAdapter("SELECT * FROM [Hoja1$]", strConn);
			myData.TableMappings.Add("Tabla", "ExcelSube");
			miDataSet.Clear();
			try
			{
				myData.Fill(miDataSet);
			}
			catch(Exception ex)
			{
				MessageBox.Show(ex.Message, "Error al abrir la Hoja1 del archivo " + txtFile.Text);
				return;
			}
			ultraGrid1.DataSource = miDataSet.Tables[0].DefaultView;
			ultraGrid1.DataBind();
		}

		private bool EsRepetido(string stNumero, string stSerie)
		{
			string stSelect = "Select Count(*) From Compra Where Numero = '"
				+ stNumero.Trim() + "' And SerieFactura = '" + stSerie.Trim() + "'";
			int iCuenta = Funcion.iEscalarSQL(cdsCliente, stSelect, true);
			if (iCuenta > 0) return true;
			else return false;
		}

		private int CalculoId (string stNumero, string stSerie)
		{
			string stSelect = "Select IdCompra From Compra Where Numero = '"
				+ stNumero.Trim() + "' And SerieFactura = '" + stSerie.Trim() + "'";
			int IdCompra = Funcion.iEscalarSQL(cdsCliente, stSelect, true);
			return IdCompra;
		}

		private void btLee_Click(object sender, System.EventArgs e)
		{
			if (chkBorrarError.Checked)
			{
				string stBorra = "Delete ErrorSalida";
				Funcion.EjecutaSQL(cdsCliente, stBorra, true);
			}
			int iLineas = 0;
			string stExec, stSerie = "";
			int IdCompraInicio = 0;

			foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.ultraGrid1.Rows.All)
			{
				if (dr.Cells[0].Value == DBNull.Value || dr.Cells[0].Value.ToString().Length == 0) continue;
				string stRuc = dr.Cells[0].Value.ToString().Trim();
				
				// crea la variable para asignarle un grupo al nuevo cliente
				int iGrupoCliente = int.Parse(dr.Cells[9].Value.ToString());
				if (chkCreaCliente.Checked)
				{
					string stNombre = dr.Cells[8].Value.ToString().Trim();
					string stSelect1 = "Exec ClienteCrear '" + stNombre + "', '" + stRuc + "', 0" + " , " + iGrupoCliente;
					Funcion.EjecutaSQL(cdsCliente, stSelect1, true);
				}

				double dSubTotalIva = double.Parse(dr.Cells[1].Value.ToString());
				double dIva = double.Parse(dr.Cells[2].Value.ToString());
				double dIce = double.Parse(dr.Cells[3].Value.ToString());
				string stNumero = dr.Cells[4].Value.ToString();
				

				DateTime dtFecha = (DateTime ) dr.Cells[5].Value;
				string stBorrar = ", 0";
				string stContadoCredito = ", 2";
				double dTotal = dSubTotalIva + dIva + dIce;
	
				int iCuenta = Funcion.iEscalarSQL(cdsCliente, "Select Count(*) From Compra", true);
				stExec = "Exec CreaFactura '" + stRuc + "', '"
					+ "', '" + stNumero + "', '" + dtFecha.ToString("yyyyMMdd") + "', "
					+ dSubTotalIva.ToString("0.00", us) + ", 0," + dIva.ToString("0.00", us) 
					+ ", " + dIce.ToString("0.00", us) + ", "
					+ dTotal.ToString("0.00", us) + ", 1" + stBorrar + stContadoCredito + ", -1";
				Funcion.EjecutaSQL(cdsCliente, stExec, true);

				string stArticulo = dr.Cells[7].Value.ToString();

				int iCuenta1 = Funcion.iEscalarSQL(cdsCliente, "Select Count(*) From Compra", true);
				if (iCuenta1 > iCuenta)
				{
					if (IdCompraInicio == 0) IdCompraInicio = CalculoId(stNumero, stSerie);
					iLineas++;
					stExec = "Exec CreaFacturaDetalle 1, '" + stArticulo + "', 12, "
						+ dSubTotalIva.ToString("0.00", us);
					Funcion.EjecutaSQL(cdsCliente, stExec, true);
				}
			}
			//Ingresa totales en facturas posteriores a la inicial
			string stSelect = "EXEC FacturaTotales " + IdCompraInicio.ToString();
			Funcion.EjecutaSQL(cdsCliente, stSelect, true);
			MessageBox.Show("Lineas Subidas: " + iLineas.ToString());
		}

		private void TelcoSan_Load(object sender, System.EventArgs e)
		{
			string stAudita = "Exec AuditaCrear 53, 6, 'InterActive'";
			Funcion.EjecutaSQL(cdsCliente, stAudita, true);
			btVerificaRet.Width = 0;
		}

		private void cdsCliente_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			cdsCliente.Schema.Connections[0].ConnectionString = 
				miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);
		}

		private void ultraButton3_Click(object sender, System.EventArgs e)
		{
			string strConn ="Provider=Microsoft.Jet.OLEDB.4.0;" +
				"Data Source=" + txtPago.Text + ";" +
				"Extended Properties='Excel 8.0;'";
			OleDbDataAdapter myData = new OleDbDataAdapter("SELECT * FROM [Hoja1$]", strConn);
			myData.TableMappings.Add("Tabla", "ExcelSube");
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

//		private bool 
		private void ultraButton2_Click(object sender, System.EventArgs e)
		{
			int iLineas = 0;
			if (chkBorrarError.Checked)
			{
				string stBorra = "Delete ErrorSalida";
				Funcion.EjecutaSQL(cdsCliente, stBorra, true);
			}

			string stNumeroPago = "SELECT LTRIM(RTRIM(CODIGO) + RTRIM(CAST(NUMERO AS VARCHAR(20)))) from COMPRANUMERO WHERE IDTIPOFACTURA = 21";
			stNumeroPago = Funcion.sEscalarSQL(cdsCliente, stNumeroPago, true);
			Funcion.EjecutaSQL(cdsCliente, "Declare @Salida char(20);Exec CompraNumera 21, @Salida OutPut", true);

			int iCol = ultraGrid1.DisplayLayout.Bands[0].Columns.Count;
			foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.ultraGrid1.Rows.All)
			{
				try
				{
					if (dr.Cells[0].Value == DBNull.Value) continue;
					//campo mandatorio ??????
					string stRuc = dr.Cells[0].Value.ToString().Trim();
					if (chkCreaCliente.Checked)
					{
						// variable para asignar el grupo del cliente
						int iGrupoCliente = int.Parse(dr.Cells[7].Value.ToString());

						string stNombre = dr.Cells[6].Value.ToString().Trim();
						string stSelect1 = "Exec ClienteCrear '" + stNombre + "', '" + stRuc + "', 0 , " + iGrupoCliente;
						Funcion.EjecutaSQL(cdsCliente, stSelect1, true);
					}

					double dPago = double.Parse(dr.Cells[1].Value.ToString());
					string stNumero = dr.Cells[2].Value.ToString();
					DateTime dtFecha = new DateTime(1900, 01, 01);
					if (dr.Cells[3].Value != DBNull.Value)
						dtFecha = (DateTime) dr.Cells[3].Value;
					string stNumCuenta = dr.Cells[4].Value.ToString();
					string stDeposito = dr.Cells[5].Value.ToString();
					string stAsiento = txtAsiento.Text;
					string stCheque = "";
					string stBanco = "";
					string stFormaPago = "";
					string stCobrador = "";
					string stReciboCaja = "";
					if (iCol > 8) stCheque = dr.Cells[8].Value.ToString();
					if (iCol > 9) stBanco = dr.Cells[9].Value.ToString();
					if (iCol > 10) stFormaPago = dr.Cells[10].Value.ToString();
					if (iCol > 11) stCobrador = dr.Cells[11].Value.ToString();
					if (iCol > 12) stReciboCaja = dr.Cells[12].Value.ToString();
					string stCreaCliente = "-1";
					if (chkClienteFactPago.Checked) stCreaCliente = "-2";

					string stSelect = "Exec CreaPago '" + stNumero + "', 1, '" + stRuc + "', "
						+ dPago.ToString("0.00", us) + ", '" + stDeposito + "', '"
						+ dtFecha.ToString("yyyyMMdd") + "', " + stCreaCliente 
						+ ", '" + stNumCuenta + "', '"
						+ stDeposito + "', '" + stNumeroPago + "', '" + stAsiento+ "', '"
						+ stBanco + "', '" + stFormaPago + "', '"
						+ stCobrador + "', '" + stReciboCaja + "'";
					Funcion.EjecutaSQL(cdsCliente, stSelect, true);
					iLineas ++;
				}
				catch (Exception ex)
				{
					MessageBox.Show(ex.Message, "Error en linea " + iLineas.ToString());
				}
			}
			MessageBox.Show("Lineas Procesadas: " + iLineas.ToString(), "Información");
		}

		private void btErrores_Click(object sender, System.EventArgs e)
		{
			TesterError miError = new TesterError();
			miError.MdiParent = this.MdiParent;
			miError.Show();
		}

		private void btAsiPago_Click(object sender, System.EventArgs e)
		{
			if (ultraGrid1.Rows.Count < 1)
			{
				MessageBox.Show("Ingrese Filas");
				return;
			}
			if (chkBorrarError.Checked)
			{
				string stBorra = "Delete ErrorSalida";
				Funcion.EjecutaSQL(cdsCliente, stBorra, true);
			}
			string stSelect = "Exec AsientoGenera 2, 'Resumen de Cobros'";
			Funcion.EjecutaSQL(cdsCliente, stSelect, true);
			int IdAsiento = Funcion.iEscalarSQL(cdsCliente, 
				"Select Top 1 idAsiento From Asiento Order by idAsiento DESC", true);

			string stCodAsiento = "Select CodAsiento From Asiento Where idAsiento = " 
				+ IdAsiento.ToString();
			stCodAsiento = Funcion.sEscalarSQL(cdsCliente, stCodAsiento, true);
			txtAsiento.Text = stCodAsiento.Trim();

			int iLineas = 0;
			double dTotal = 0;
			string stNumCuenta="";
			string stDepAntes = "Ninguno";
			double dDepTotal = 0;

			string stRuc = "";
			double dHaber = 0;
			string stNumero = "";
			string stDeposito = "";

			Infragistics.Win.UltraWinGrid.UltraGridRow dr1 = this.ultraGrid1.Rows[0];
			{
				stRuc = dr1.Cells[0].Value.ToString();
				dHaber = double.Parse(dr1.Cells[1].Value.ToString());
				dTotal += dHaber;
				stNumero = dr1.Cells[2].Value.ToString();
				stNumCuenta = dr1.Cells[4].Value.ToString();
				stDepAntes = dr1.Cells[5].Value.ToString();
			}

			foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.ultraGrid1.Rows.All)
			{
				try
				{
					if (dr.Cells[0].Value == DBNull.Value) continue;

					stRuc = dr.Cells[0].Value.ToString().Trim();
					string stNombre = dr.Cells[6].Value.ToString().Trim();
					if (chkCreaCliente.Checked)
					{
						stSelect = "Exec ClienteCrear '" + stNombre + "', '" + stRuc + "', 0";
						Funcion.EjecutaSQL(cdsCliente, stSelect, true);
					}

					stDeposito = dr.Cells[5].Value.ToString();
					if (stDepAntes != stDeposito)
					{
						stSelect = "Exec AsientoDetalleGenera "
							+ dDepTotal.ToString("0.00", us) + ", 0, -3, " + IdAsiento.ToString() + ", ''"
							+ ", 0, 0, '" + stDepAntes + "', '" + stNumCuenta + "'";
						Funcion.EjecutaSQL(cdsCliente, stSelect, true);
						dDepTotal = 0;
						stDepAntes = stDeposito;
					}
					
					dHaber = double.Parse(dr.Cells[1].Value.ToString());
					dTotal += dHaber;
					stNumero = dr.Cells[2].Value.ToString();
					stNumCuenta = dr.Cells[4].Value.ToString();
					//if (stDepAntes == "Ninguno") stDepAntes = stDeposito;

					dDepTotal += dHaber;
				
					stSelect = "Exec AsientoDetalleGenera 0, "
						+ dHaber.ToString("0.00", us) + ", -1, " + IdAsiento.ToString() 
						+ ", 'Fac: " + stNumero.Trim() + " " + stNombre
						+ "', 0, 0, '', '" + stRuc + "'";
					Funcion.EjecutaSQL(cdsCliente, stSelect, true);

					iLineas ++;
				}
				catch (Exception ex)
				{
					MessageBox.Show(ex.Message, "Error en linea " + iLineas.ToString());
				}
			}
			#region Genera Suma de Total Ultimo 
			stSelect = "Exec AsientoDetalleGenera "
				+ dDepTotal.ToString("0.00", us) + ", 0, -3, " + IdAsiento.ToString() + ", ''"
				+ ", 0, 0, '" + stDepAntes + "', '" + stNumCuenta + "'";
			Funcion.EjecutaSQL(cdsCliente, stSelect, true);
			#endregion Fin de Generacion de Suma de Total Ultimo 
		
			if (dTotal > 0)
			{
				stSelect = "Exec AsientoTotal " + IdAsiento.ToString();
				Funcion.EjecutaSQL(cdsCliente, stSelect, true);
				iLineas ++;
			}
			MessageBox.Show("Asiento Generado con " + iLineas.ToString() + " lineas", "Información");
		}

		private void btAsiFact_Click(object sender, System.EventArgs e)
		{
			if (chkBorrarError.Checked)
			{
				string stBorra = "Delete ErrorSalida";
				Funcion.EjecutaSQL(cdsCliente, stBorra, true);
			}
			string stSelect = "Exec AsientoGenera 1, 'Resumen de Facturas'";
			Funcion.EjecutaSQL(cdsCliente, stSelect, true);
			int IdAsiento = Funcion.iEscalarSQL(cdsCliente, 
				"Select Top 1 idAsiento From Asiento Order by idAsiento DESC", true);

			int iLineas = 0;
			double dSumSubT = 0;
			double dIvaAc = 0;
			double dIceAc = 0;
			foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.ultraGrid1.Rows.All)
			{
				try
				{
					if (dr.Cells[0].Value == DBNull.Value || dr.Cells[0].Value.ToString().Length == 0) continue;
					iLineas++;
					string stNombreCliente = dr.Cells[0].Value.ToString();
					double dSubTotal = double.Parse(dr.Cells[1].Value.ToString());
					double dIva = double.Parse(dr.Cells[2].Value.ToString());
					double dIce = double.Parse(dr.Cells[3].Value.ToString());
					string stNumero = dr.Cells[4].Value.ToString();
					DateTime dtFecha = (DateTime) dr.Cells[5].Value;
					string stDetalle = dr.Cells[6].Value.ToString() + " " 
						+ stNumero;// + " " + dtFecha.ToString("dd/MMM/yyyy");
					double dTotal = dSubTotal + dIva + dIce;
					dSumSubT += dSubTotal;
					dIceAc += dIce;
					dIvaAc += dIva;

					stSelect = "Exec AsientoDetalleGenera "
						+ dTotal.ToString("0.00", us) + ", 0, -1, " + IdAsiento.ToString() + ", @Detalle = '"
						+ stDetalle + "', @Parametro = '" + stNombreCliente + "'";
					Funcion.EjecutaSQL(cdsCliente, stSelect, true);
				}
				catch (Exception ex)
				{
					MessageBox.Show(ex.Message, "Error en linea " + iLineas.ToString());
				}
			}

			if (dSumSubT + dIvaAc + dIceAc > 0)
			{
				stSelect = "Exec AsientoDetalleGenera 0, "
					+ dSumSubT.ToString("0.00", us) + ", -5, " + IdAsiento.ToString() + 
					", @Parametro = '5'";
				Funcion.EjecutaSQL(cdsCliente, stSelect, true);

				stSelect = "Exec AsientoDetalleGenera 0, "
					+ dIvaAc.ToString("0.00", us) + ", -5, " + IdAsiento.ToString()
					+ ", @Parametro = '2'";
				Funcion.EjecutaSQL(cdsCliente, stSelect, true);

				stSelect = "Exec AsientoDetalleGenera 0, "
					+ dIceAc.ToString("0.00", us) + ", -5, " + IdAsiento.ToString()
					+ ", @Parametro = '9'";
				Funcion.EjecutaSQL(cdsCliente, stSelect, true);

				stSelect = "Exec AsientoTotal " + IdAsiento.ToString();
				Funcion.EjecutaSQL(cdsCliente, stSelect, true);
				iLineas ++;
			}
			MessageBox.Show("Asiento Generado con " + iLineas.ToString() + " lineas", "Información");		
		}

		private void btExcelRet_Click(object sender, System.EventArgs e)
		{
			string strConn ="Provider=Microsoft.Jet.OLEDB.4.0;" +
				"Data Source=" + txtRetencion.Text + ";" +
				"Extended Properties='Excel 8.0;'";
			OleDbDataAdapter myData = new OleDbDataAdapter("SELECT * FROM [Hoja1$]", strConn);
			myData.TableMappings.Add("Tabla", "ExcelSube");
			miDataSet.Clear();
			try
			{
				myData.Fill(miDataSet);
			}
			catch(Exception ex)
			{
				MessageBox.Show(ex.Message, "Error al abrir la Hoja1 del archivo " + txtRetencion.Text);
				return;
			}
			ultraGrid1.DataSource = null;
			ultraGrid1.DataBind();
			ultraGrid1.DataSource = miDataSet.Tables[0].DefaultView;
			ultraGrid1.DataBind();
		}

		private int idPagoLote = 0;
		private void btRetencion_Click(object sender, System.EventArgs e)
		{
			int iCuentaTodo = 0;
			foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.ultraGrid1.Rows.All)
			{
				if (dr.Cells[4].Value == DBNull.Value || dr.Cells[4].Value.ToString().Length == 0) continue;
				iCuentaTodo ++;
			}
			if (iCuentaTodo == 0)
			{
				MessageBox.Show("No hay datos", "Información");
				return;
			}
			string stFactura = "";
			string stRuc = "";
			string stSelect = "";
			int idCliente = 0;
			idPagoLote = 0;
			string stRetencion = "";
			string stRetencionAntes = "";

			Infragistics.Win.UltraWinGrid.UltraGridRow dr1 = this.ultraGrid1.Rows[0];
			{
				stRetencionAntes = dr1.Cells[10].Value.ToString();
				stRuc = dr1.Cells[0].Value.ToString();
				stSelect = "Insert Into PagoLote (idCliente, idTipoFactura, Codigo, CodAsiento, Nota) "
					+ "Values (" + idCliente.ToString() + ", 4, '', '', '')";
				Funcion.EjecutaSQL(cdsCliente, stSelect, true);
				stSelect = "Select Top 1 idPagoLote From PagoLote Order by idPagoLote DESC";
				idPagoLote = Funcion.iEscalarSQL(cdsCliente, stSelect, true);
				stSelect = "Select idCliente From Cliente Where LTRIM(RTRIM(Ruc)) = '"
					+ stRuc.Trim() + "'";
				idCliente = Funcion.iEscalarSQL(cdsCliente, stSelect, true);
			}

			int iLineas = 0;
			foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.ultraGrid1.Rows.All)
			{
				if (dr.Cells[4].Value == DBNull.Value || dr.Cells[4].Value.ToString().Length == 0) continue;
				stFactura = dr.Cells[4].Value.ToString();
				if (stFactura == "324")
					MessageBox.Show("ultimo");

				if (dr.Cells[0].Value != DBNull.Value && dr.Cells[0].Value.ToString().Length > 0)
					stRuc = dr.Cells[0].Value.ToString();

				stSelect = "Select Count(*) From Compra c Inner Join Cliente cl "
					+ " On c.idCliente = cl.idCliente Where LTrim(Rtrim(c.Numero)) = '"
					+ stFactura.Trim() + "' And c.idTipoFactura In (4)"
					+ " And LTrim(RTrim(cl.Ruc)) = '" + stRuc.Trim() + "'";
				int iCuenta = Funcion.iEscalarSQL(cdsCliente, stSelect, true);
				if (iCuenta != 1) 
				{
					if (iCuenta ==0)
						MessageBox.Show("No existe Factura " + stFactura, "Información");
					else
						MessageBox.Show("Existen " + iCuenta.ToString() + " Facturas con el numero " 
							+ stFactura, "Información");
					continue;
				}
				iLineas ++;

				stRetencion = dr.Cells[10].Value.ToString();
				if (stRetencion != stRetencionAntes || iCuentaTodo == iLineas)
				{
					int idPagoLote1 = idPagoLote;
					if (iCuentaTodo != iLineas)
					{
						stSelect = "Insert Into PagoLote (idCliente, idTipoFactura, Codigo, CodAsiento, Nota) "
							+ "Values (" + idCliente.ToString() + ", 4, '', '', '')";
						Funcion.EjecutaSQL(cdsCliente, stSelect, true);
						stSelect = "Select Top 1 idPagoLote From PagoLote Order by idPagoLote DESC";
						idPagoLote = Funcion.iEscalarSQL(cdsCliente, stSelect, true);
					}
					else
					{
						if (stRetencion != stRetencionAntes)
						{
							stSelect = "Insert Into PagoLote (idCliente, idTipoFactura, Codigo, CodAsiento, Nota) "
								+ "Values (" + idCliente.ToString() + ", 4, '', '', '')";
							Funcion.EjecutaSQL(cdsCliente, stSelect, true);
							stSelect = "Select Top 1 idPagoLote From PagoLote Order by idPagoLote DESC";
							idPagoLote = Funcion.iEscalarSQL(cdsCliente, stSelect, true);

							stSelect = "Select c.IdCompra From Compra c Inner Join Cliente cl "
								+ " On c.idCliente = cl.idCliente Where LTrim(Rtrim(c.Numero)) = '"
								+ stFactura.Trim() + "' And c.idTipoFactura In (4)"
								+ " And LTrim(RTrim(cl.Ruc)) = '" + stRuc.Trim() + "'";
							int idCompraFin = Funcion.iEscalarSQL(cdsCliente, stSelect, true);
							stSelect = "Update Compra Set idPagoLot = " + idPagoLote.ToString()
								+ " Where idCompra = " + idCompraFin.ToString();
							Funcion.EjecutaSQL(cdsCliente, stSelect, true);
							
							stSelect = "EXEC RetencionCreaLote " + idPagoLote.ToString();
							Funcion.EjecutaSQL(cdsCliente, stSelect, true);
							string stDistribuye1 = "EXEC DistribuyeRet " + idPagoLote.ToString();
							Funcion.EjecutaSQL(cdsCliente, stDistribuye1, true);

							DateTime dtFecha1 = (DateTime) dr.Cells[11].Value;
							string stAutRet1 = dr.Cells[13].Value.ToString();
							string stSerieRet1 = dr.Cells[12].Value.ToString();
							stSelect = "Select Top 1 idRetencion From Retencion Order by idRetencion Desc"; 
							int IdRetencion1 = Funcion.iEscalarSQL(cdsCliente, stSelect, true);
							stSelect = "Update Retencion Set AutorizaRet = '" + stAutRet1 + "', "
								+ "SerieRet = '" + stSerieRet1 + "', Codigo = '" + stRetencion
								+ "', IngresoEgreso = 2, Fecha = '" + dtFecha1.ToString("yyyyMMdd")
								+ "', RentaRetenida = (Select Sum(Retenido) From RetDetalle Where idRetencion = "
								+ IdRetencion1.ToString() + ") From Retencion "
								+ "Where Retencion.idRetencion = " + IdRetencion1.ToString();
							Funcion.EjecutaSQL(cdsCliente, stSelect, true);
						}
						else
						{
							stSelect = "Select c.IdCompra From Compra c Inner Join Cliente cl "
								+ " On c.idCliente = cl.idCliente Where LTrim(Rtrim(c.Numero)) = '"
								+ stFactura.Trim() + "' And c.idTipoFactura In (4)"
								+ " And LTrim(RTrim(cl.Ruc)) = '" + stRuc.Trim() + "'";
							int idCompra1 = Funcion.iEscalarSQL(cdsCliente, stSelect, true);
							stSelect = "Update Compra Set idPagoLot = " + idPagoLote.ToString()
								+ " Where idCompra = " + idCompra1.ToString();
							Funcion.EjecutaSQL(cdsCliente, stSelect, true);
						}
					}

					stSelect = "EXEC RetencionCreaLote " + idPagoLote1.ToString();
					Funcion.EjecutaSQL(cdsCliente, stSelect, true);
					stSelect = "Select Top 1 idRetencion From Retencion Order by idRetencion Desc"; 
					int IdRetencion = Funcion.iEscalarSQL(cdsCliente, stSelect, true);
					DateTime dtFecha = (DateTime) dr.Cells[11].Value;
					string stAutRet = dr.Cells[13].Value.ToString();
					string stSerieRet = dr.Cells[12].Value.ToString();
					stSelect = "Update Retencion Set AutorizaRet = '" + stAutRet + "', "
						+ "SerieRet = '" + stSerieRet + "', Codigo = '" + stRetencionAntes
						+ "', IngresoEgreso = 2, Fecha = '" + dtFecha.ToString("yyyyMMdd")
						+ "', RentaRetenida = (Select Sum(Retenido) From RetDetalle Where idRetencion = "
						+ IdRetencion.ToString() + ") From Retencion "
						+ "Where Retencion.idRetencion = " + IdRetencion.ToString();
					Funcion.EjecutaSQL(cdsCliente, stSelect, true);

					string stDistribuye = "EXEC DistribuyeRet " + idPagoLote1.ToString();
					Funcion.EjecutaSQL(cdsCliente, stDistribuye, true);
					stRetencionAntes = stRetencion;
				}

				if (iCuentaTodo != iLineas)
				{
					stSelect = "Select c.IdCompra From Compra c Inner Join Cliente cl "
						+ " On c.idCliente = cl.idCliente Where LTrim(Rtrim(c.Numero)) = '"
						+ stFactura.Trim() + "' And c.idTipoFactura In (4)"
						+ " And LTrim(RTrim(cl.Ruc)) = '" + stRuc.Trim() + "'";
					int idCompra = Funcion.iEscalarSQL(cdsCliente, stSelect, true);
					stSelect = "Update Compra Set idPagoLot = " + idPagoLote.ToString()
						+ " Where idCompra = " + idCompra.ToString();
					Funcion.EjecutaSQL(cdsCliente, stSelect, true);
				}
			}
			MessageBox.Show("Facturas Afectadas por retencion: " + iLineas.ToString());
		}

		private void btVerificaRet_Click(object sender, System.EventArgs e)
		{
			int iLineas = 0;
			string stFactura = "";
			string stRuc = "";
			foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.ultraGrid1.Rows.All)
			{
				if (dr.Cells[0].Value == DBNull.Value || dr.Cells[0].Value.ToString().Length == 0)
					stFactura = "Vacio";
				else
					stFactura = dr.Cells[0].Value.ToString();

				if (dr.Cells[1].Value != DBNull.Value && dr.Cells[1].Value.ToString().Length > 0)
					stRuc = dr.Cells[1].Value.ToString();
				string stSelect = "Select Count(*) From Compra c Inner Join Cliente cl "
					+ " On c.idCliente = cl.idCliente Where LTrim(Rtrim(c.Numero)) = '"
					+ stFactura.Trim() + "' And c.idTipoFactura In (4)"
					+ " And LTrim(RTrim(cl.Ruc)) = '" + stRuc.Trim() + "'";
				int iCuenta = Funcion.iEscalarSQL(cdsCliente, stSelect, true);
				if (iCuenta != 1) 
				{
					if (iCuenta ==0)
						MessageBox.Show("No existe Factura " + stFactura, "Información");
					else
						MessageBox.Show("Existen " + iCuenta.ToString() + " facturas con el numero " 
							+ stFactura, "Información");
					continue;
				}
				iLineas ++;
			}
			MessageBox.Show("Facturas Validas: " + iLineas.ToString());
		}

		private void btAsiRet_Click(object sender, System.EventArgs e)
		{
			if (idPagoLote == 0)
			{
				MessageBox.Show("Antes de generar el asiento debe presionar el boton Ret.", "Información");
				return;
			}
			string stExec = "EXEC AsientoRetencion 0, "
				+ idPagoLote.ToString() + ", 4";
			Funcion.EjecutaSQL(cdsCliente, stExec, true);
			MessageBox.Show("Asiento Generado", "Información");
		}

		private void btFactRet_Click(object sender, System.EventArgs e)
		{
			int iCuentaTodo = 0;
			foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.ultraGrid1.Rows.All)
			{
				if (dr.Cells[4].Value == DBNull.Value || dr.Cells[4].Value.ToString().Length == 0) continue;
				iCuentaTodo ++;
			}
			if (iCuentaTodo == 0)
			{
				MessageBox.Show("No hay datos", "Información");
				return;
			}
			if (chkBorrarError.Checked)
			{
				string stBorra = "Delete ErrorSalida";
				Funcion.EjecutaSQL(cdsCliente, stBorra, true);
			}

			int iLineas = 0;
			string stExec;
			int IdCompraInicio = 0;
			
			foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.ultraGrid1.Rows.All)
			{
				if (dr.Cells[0].Value == DBNull.Value || dr.Cells[0].Value.ToString().Length == 0) continue;
				string stRuc = dr.Cells[0].Value.ToString().Trim();

				double dSubTotalIva = double.Parse(dr.Cells[1].Value.ToString());
				double dIva = double.Parse(dr.Cells[2].Value.ToString());
				double dIce = double.Parse(dr.Cells[3].Value.ToString());
				string stNumero = dr.Cells[4].Value.ToString();
				DateTime dtFecha = (DateTime ) dr.Cells[5].Value;
				string stBorrar = ", 0";
				string stContadoCredito = ", 2";
				double dTotal = dSubTotalIva + dIva + dIce;
				string stSerie1 = dr.Cells[6].Value.ToString();
				string stAutoriza = dr.Cells[8].Value.ToString();
				DateTime dtFechaCad = (DateTime ) dr.Cells[9].Value;
				string stCredTribut = dr.Cells[14].Value.ToString();
	
				int iCuenta = Funcion.iEscalarSQL(cdsCliente, "Select Count(*) From Compra", true);
				stExec = "Exec CreaFactura '" + stRuc + "', '" + stSerie1
					+ "', '" + stNumero + "', '" + dtFecha.ToString("yyyyMMdd") + "', "
					+ dSubTotalIva.ToString("0.00", us) + ", 0," + dIva.ToString("0.00", us) 
					+ ", " + dIce.ToString("0.00", us) + ", "
					+ dTotal.ToString("0.00", us) + ", 4" + stBorrar + stContadoCredito + ", -1, '"
					+ stAutoriza + "', '" + dtFechaCad.ToString("yyyyMMdd") + "', "
					+ stCredTribut;
				Funcion.EjecutaSQL(cdsCliente, stExec, true);
				if (IdCompraInicio == 0) IdCompraInicio = CalculoId(stNumero, stSerie1);

				string stArticulo = dr.Cells[7].Value.ToString();

				int iCuenta1 = Funcion.iEscalarSQL(cdsCliente, "Select Count(*) From Compra", true);
				if (iCuenta1 > iCuenta)
				{
					iLineas++;
					stExec = "Exec CreaFacturaDetalle 1, '" + stArticulo + "', 12, "
						+ dSubTotalIva.ToString("0.00", us);
					Funcion.EjecutaSQL(cdsCliente, stExec, true);
				}
			}
			//Ingresa totales en facturas posteriores a la inicial
			string stSelect = "EXEC FacturaTotales " + IdCompraInicio.ToString();
			Funcion.EjecutaSQL(cdsCliente, stSelect, true);
			MessageBox.Show("Lineas Subidas: " + iLineas.ToString());
		}

		private void btReservaPago_Click(object sender, System.EventArgs e)
		{
			if (DialogResult.No == MessageBox.Show("Desea Transformar las reservaciones a pagos",
				"Información", MessageBoxButtons.YesNo, MessageBoxIcon.Question)) return;

			string stExec = "Exec ReservaPago";
			Funcion.EjecutaSQL(cdsCliente, stExec, true);
			MessageBox.Show("Reserva trasformadas a Pagos", "Información");
		}

		private void Interactive_Closing(object sender, System.ComponentModel.CancelEventArgs e)
		{
			string stAudita = "Exec AuditaCrear 53, 7, 'Interactive'";
			Funcion.EjecutaSQL(cdsCliente, stAudita, true);

		}

		private void ultraButton4_Click(object sender, System.EventArgs e)
		{
			string strConn ="Provider=Microsoft.Jet.OLEDB.4.0;" +
				"Data Source=" + txtFactRapida.Text + ";" +
				"Extended Properties='Excel 8.0;'";
			OleDbDataAdapter myData = new OleDbDataAdapter("SELECT * FROM [Hoja1$]", strConn);
			myData.TableMappings.Add("Tabla", "ExcelSube");
			miDataSet.Clear();
			try
			{
				myData.Fill(miDataSet);
			}
			catch(Exception ex)
			{
				MessageBox.Show(ex.Message, "Error al abrir la Hoja1 del archivo " + txtFile.Text);
				return;
			}
			ultraGrid1.DataSource = miDataSet.Tables[0].DefaultView;
			ultraGrid1.DataBind();
		}

		private void btExcelFactRapida_MouseDown(object sender, System.Windows.Forms.MouseEventArgs e)
		{
			if (e.Button != MouseButtons.Right) return;

			if (chkBorrarError.Checked)
			{
				string stBorra = "Delete ErrorSalida";
				Funcion.EjecutaSQL(cdsCliente, stBorra, true);
			}
			int iLineas = 0;
			int IdCompraInicio = 0;

//			foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.ultraGrid1.Rows.All)
//			{
//				if (dr.Cells[0].Value != DBNull.Value && dr.Cells[0].Value.ToString().Length > 0)
//				{
//					string stNumero = dr.Cells[0].Value.ToString().Trim();
//					DateTime dtFecha = (DateTime) dr.Cells[1].Value;
//					string stNombre = dr.Cells[2].Value.ToString().Trim();
//
//					double dSubTotalIva = 0;
//					double dIva = 0;
//					double dIce = 0;
//					string stBorrar = ", 0";
//					string stContadoCredito = ", 2";
//					double dTotal = dSubTotalIva + dIva + dIce;
//	
//					int iCuenta = Funcion.iEscalarSQL(cdsCliente, "Select Count(*) From Compra", true);
//					stExec = "Exec CreaFactura '" + stRuc + "', '"
//						+ "', '" + stNumero + "', '" + dtFecha.ToString("yyyyMMdd") + "', "
//						+ dSubTotalIva.ToString("0.00", us) + ", 0," + dIva.ToString("0.00", us) 
//						+ ", " + dIce.ToString("0.00", us) + ", "
//						+ dTotal.ToString("0.00", us) + ", 1" + stBorrar + stContadoCredito + ", -1";
//					Funcion.EjecutaSQL(cdsCliente, stExec, true);
//				}
//				else
//				{
//					if (dr.Cells[3].Value != DBNull.Value && dr.Cells[3].Value.ToString().Length > 0) 
//						continue;
//					string stArticulo = dr.Cells[4].Value.ToString();
//
//					int iCuenta1 = Funcion.iEscalarSQL(cdsCliente, "Select Count(*) From Compra", true);
//					if (iCuenta1 > iCuenta)
//					{
//						if (IdCompraInicio == 0) IdCompraInicio = CalculoId(stNumero, stSerie);
//						iLineas++;
//						stExec = "Exec CreaFacturaDetalle 1, '" + stArticulo + "', 12, "
//							+ dSubTotalIva.ToString("0.00", us);
//						Funcion.EjecutaSQL(cdsCliente, stExec, true);
//					}
//				}
//			}
			//Ingresa totales en facturas posteriores a la inicial
			string stSelect = "EXEC FacturaTotales " + IdCompraInicio.ToString();
			Funcion.EjecutaSQL(cdsCliente, stSelect, true);
			MessageBox.Show("Lineas Subidas: " + iLineas.ToString());
		}

		/*private void btnExcelPrueba_Click(object sender, System.EventArgs e)
		{
			string strConn ="Provider=Microsoft.Jet.OLEDB.4.0;" +
				"Data Source=" + txtPago.Text + ";" +
				"Extended Properties='Excel 8.0;'";
			OleDbDataAdapter myData = new OleDbDataAdapter("SELECT * FROM [Hoja1$]", strConn);
			myData.TableMappings.Add("Tabla", "ExcelSube");
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
			grabaExcel();
		}
		public void grabaExcel ()
		{
			foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.ultraGrid1.Rows.All)
			{
				try
				{
					//if (dr.Cells[0].Value == DBNull.Value) continue;
					double stDias = Double.Parse(dr.Cells[0].Value.ToString());
					string stNombre = dr.Cells[1].Value.ToString().Trim();
					DateTime stFecha = DateTime.Parse(dr.Cells[2].Value.ToString());
			
					string graba = "Exec grabaPrueba " + stDias.ToString("0.00",us) + " , ' "+ stFecha+" ', '"+stNombre +" '";
					string recibe = Funcion.sEscalarSQL(cdsCliente, graba, true);
					if (recibe=="OK")
					{
						MessageBox.Show("Cliente Grabado....");
					}
					else
					{
						MessageBox.Show("Cliente Grabado....","GRABAR",MessageBoxButtons.AbortRetryIgnore,MessageBoxIcon.Error);
					}
//					MessageBox.Show(" ..: "+stDias.ToString("0.00",us) + "  : "+stNombre +"  :" +stFecha.ToString("yyyy/MMM/dd"));
				}
				catch (Exception ex)
				{
					MessageBox.Show(ex.Message, "Error en linea ");
				}
			}
		}*/
	}
}
