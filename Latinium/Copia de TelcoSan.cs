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
	public class TelcoSan : DevExpress.XtraEditors.XtraForm
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
		private Infragistics.Win.Misc.UltraButton btAsiento;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtArchAsiento;
		private Infragistics.Win.UltraWinEditors.UltraComboEditor cmbTipo;
		private Infragistics.Win.Misc.UltraButton ultraButton2;
		private Infragistics.Win.Misc.UltraButton ultraButton3;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtPago;
		CultureInfo us = new CultureInfo("en-US");
		private Infragistics.Win.UltraWinEditors.UltraCheckEditor chkCreaCliente;
		private Infragistics.Win.Misc.UltraButton btErrores;
		private Infragistics.Win.UltraWinEditors.UltraCheckEditor chkBorrarError;
		CultureInfo es = new CultureInfo("es");

		public TelcoSan()
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
			Infragistics.Win.Appearance appearance2 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance3 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance4 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance5 = new Infragistics.Win.Appearance();
			Infragistics.Win.ValueListItem valueListItem1 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.ValueListItem valueListItem2 = new Infragistics.Win.ValueListItem();
			this.ultraButton1 = new Infragistics.Win.Misc.UltraButton();
			this.txtFile = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.ultraGrid1 = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.btLee = new Infragistics.Win.Misc.UltraButton();
			this.cdsCliente = new C1.Data.C1DataSet();
			this.btAsiento = new Infragistics.Win.Misc.UltraButton();
			this.txtArchAsiento = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.cmbTipo = new Infragistics.Win.UltraWinEditors.UltraComboEditor();
			this.ultraButton2 = new Infragistics.Win.Misc.UltraButton();
			this.txtPago = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.ultraButton3 = new Infragistics.Win.Misc.UltraButton();
			this.chkCreaCliente = new Infragistics.Win.UltraWinEditors.UltraCheckEditor();
			this.btErrores = new Infragistics.Win.Misc.UltraButton();
			this.chkBorrarError = new Infragistics.Win.UltraWinEditors.UltraCheckEditor();
			((System.ComponentModel.ISupportInitialize)(this.txtFile)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraGrid1)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsCliente)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtArchAsiento)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbTipo)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtPago)).BeginInit();
			this.SuspendLayout();
			// 
			// ultraButton1
			// 
			this.ultraButton1.Location = new System.Drawing.Point(40, 40);
			this.ultraButton1.Name = "ultraButton1";
			this.ultraButton1.TabIndex = 0;
			this.ultraButton1.Text = "Excel";
			this.ultraButton1.Click += new System.EventHandler(this.ultraButton1_Click);
			// 
			// txtFile
			// 
			this.txtFile.Location = new System.Drawing.Point(136, 40);
			this.txtFile.Name = "txtFile";
			this.txtFile.Size = new System.Drawing.Size(128, 21);
			this.txtFile.TabIndex = 1;
			this.txtFile.Text = "C:\\FacturaPlano.Xls";
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
			this.ultraGrid1.Location = new System.Drawing.Point(8, 104);
			this.ultraGrid1.Name = "ultraGrid1";
			this.ultraGrid1.Size = new System.Drawing.Size(448, 152);
			this.ultraGrid1.TabIndex = 2;
			// 
			// btLee
			// 
			this.btLee.Location = new System.Drawing.Point(376, 40);
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
			// btAsiento
			// 
			this.btAsiento.Location = new System.Drawing.Point(376, 8);
			this.btAsiento.Name = "btAsiento";
			this.btAsiento.TabIndex = 4;
			this.btAsiento.Text = "Asientos";
			this.btAsiento.Click += new System.EventHandler(this.btAsiento_Click);
			// 
			// txtArchAsiento
			// 
			this.txtArchAsiento.Location = new System.Drawing.Point(136, 8);
			this.txtArchAsiento.Name = "txtArchAsiento";
			this.txtArchAsiento.Size = new System.Drawing.Size(128, 21);
			this.txtArchAsiento.TabIndex = 5;
			this.txtArchAsiento.Text = "C:\\Prueba.dbf";
			// 
			// cmbTipo
			// 
			this.cmbTipo.DropDownStyle = Infragistics.Win.DropDownStyle.DropDownList;
			valueListItem1.DataValue = 0;
			valueListItem1.DisplayText = "Ventas";
			valueListItem2.DataValue = 1;
			valueListItem2.DisplayText = "Recaudador";
			this.cmbTipo.Items.Add(valueListItem1);
			this.cmbTipo.Items.Add(valueListItem2);
			this.cmbTipo.Location = new System.Drawing.Point(16, 8);
			this.cmbTipo.Name = "cmbTipo";
			this.cmbTipo.Size = new System.Drawing.Size(112, 21);
			this.cmbTipo.TabIndex = 6;
			// 
			// ultraButton2
			// 
			this.ultraButton2.Location = new System.Drawing.Point(376, 72);
			this.ultraButton2.Name = "ultraButton2";
			this.ultraButton2.TabIndex = 9;
			this.ultraButton2.Text = "Pago";
			this.ultraButton2.Click += new System.EventHandler(this.ultraButton2_Click);
			// 
			// txtPago
			// 
			this.txtPago.Location = new System.Drawing.Point(136, 72);
			this.txtPago.Name = "txtPago";
			this.txtPago.Size = new System.Drawing.Size(128, 21);
			this.txtPago.TabIndex = 8;
			this.txtPago.Text = "C:\\PagoPlano.Xls";
			// 
			// ultraButton3
			// 
			this.ultraButton3.Location = new System.Drawing.Point(40, 72);
			this.ultraButton3.Name = "ultraButton3";
			this.ultraButton3.TabIndex = 7;
			this.ultraButton3.Text = "Excel Pago";
			this.ultraButton3.Click += new System.EventHandler(this.ultraButton3_Click);
			// 
			// chkCreaCliente
			// 
			this.chkCreaCliente.Location = new System.Drawing.Point(280, 72);
			this.chkCreaCliente.Name = "chkCreaCliente";
			this.chkCreaCliente.Size = new System.Drawing.Size(88, 20);
			this.chkCreaCliente.TabIndex = 10;
			this.chkCreaCliente.Text = "Crea Cliente";
			// 
			// btErrores
			// 
			this.btErrores.Location = new System.Drawing.Point(280, 40);
			this.btErrores.Name = "btErrores";
			this.btErrores.TabIndex = 11;
			this.btErrores.Text = "Errores";
			this.btErrores.Click += new System.EventHandler(this.btErrores_Click);
			// 
			// chkBorrarError
			// 
			this.chkBorrarError.Location = new System.Drawing.Point(280, 8);
			this.chkBorrarError.Name = "chkBorrarError";
			this.chkBorrarError.Size = new System.Drawing.Size(88, 20);
			this.chkBorrarError.TabIndex = 12;
			this.chkBorrarError.Text = "Borrar Error";
			// 
			// TelcoSan
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 13);
			this.ClientSize = new System.Drawing.Size(464, 270);
			this.Controls.Add(this.chkBorrarError);
			this.Controls.Add(this.btErrores);
			this.Controls.Add(this.chkCreaCliente);
			this.Controls.Add(this.ultraButton2);
			this.Controls.Add(this.txtPago);
			this.Controls.Add(this.ultraButton3);
			this.Controls.Add(this.cmbTipo);
			this.Controls.Add(this.txtArchAsiento);
			this.Controls.Add(this.btAsiento);
			this.Controls.Add(this.btLee);
			this.Controls.Add(this.ultraGrid1);
			this.Controls.Add(this.txtFile);
			this.Controls.Add(this.ultraButton1);
			this.Name = "TelcoSan";
			this.Text = "TelcolSam";
			this.Closing += new System.ComponentModel.CancelEventHandler(this.TelcoSan_Closing);
			this.Load += new System.EventHandler(this.TelcoSan_Load);
			((System.ComponentModel.ISupportInitialize)(this.txtFile)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraGrid1)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsCliente)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtArchAsiento)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbTipo)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtPago)).EndInit();
			this.ResumeLayout(false);

		}
		#endregion

		DataSet miDataSet = new DataSet();
		private void ultraButton1_Click(object sender, System.EventArgs e)
		{
       string strConn ="Provider=Microsoft.Jet.OLEDB.4.0;" +
         "Data Source=" + txtFile.Text + ";" +
         "Extended Properties='Excel 8.0;'";
			OleDbDataAdapter myData = new OleDbDataAdapter("SELECT * FROM [Trafico$]", strConn);
			myData.TableMappings.Add("Tabla", "ExcelSube");
			miDataSet.Clear();
			try
			{
				myData.Fill(miDataSet);
			}
			catch(Exception ex)
			{
				MessageBox.Show("Debe tener abierto el archivo " + txtFile.Text + " en Exel con la hoja trafico",
					ex.Message);
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
			string stNumAntes = "";
			int iLineas = 0;
			string stExec, stSerie = "", stRepetido = "";
			int IdCompraInicio = 0;
		
			foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.ultraGrid1.Rows.All)
			{
				double dIva2=0;
				string stNumero = dr.Cells[2].Value.ToString();
				if (stNumero.Trim() == "" || stNumero == stRepetido) 
				{
					if (stNumero == stRepetido)	dr.Cells[9].Value = "Rep";
					continue;
				}
				if (stNumAntes != stNumero)
				{
					stSerie = dr.Cells[1].Value.ToString();
					stSerie = Funcion.Digito(stSerie);
					if (EsRepetido(stNumero, stSerie)) 
					{
						stRepetido = stNumero;
						dr.Cells[9].Value = "Rep";
						continue;
					}
					iLineas++;
					string stNombreCliente = dr.Cells[0].Value.ToString();
					DateTime dtFecha = (DateTime ) dr.Cells[3].Value;
					DateTime dtHora = (DateTime ) dr.Cells[4].Value;
					double dIce = double.Parse(dr.Cells[6].Value.ToString());
					double dIva = double.Parse(dr.Cells[7].Value.ToString());
					dIva2=dIva;  
					string stAnulado = dr.Cells[9].Value.ToString();
					string stBorrar = "1";
					if (stAnulado.ToUpper() == "NO") stBorrar = "0"; 
					
					stExec = "Exec CreaFactura '" + stNombreCliente + "', '"
						+ stSerie + "', '" + stNumero + "', '" + dtFecha.ToString("yyyyMMdd") + "', "
						+ "0, 0," + dIva.ToString("0.00", us) + ", " + dIce.ToString("0.00", us) + ", "
						+ "0, 1, " + stBorrar;
					Funcion.EjecutaSQL(cdsCliente, stExec, true);
					stNumAntes = stNumero;
				}
				
				if (IdCompraInicio == 0) IdCompraInicio = CalculoId(stNumero, stSerie);
				
				string stArticulo = dr.Cells[5].Value.ToString();
				double dSubTotalIva = double.Parse(dr.Cells[8].Value.ToString());
				double dTotal = double.Parse(dr.Cells[10].Value.ToString());
				dIva2 = double.Parse(dr.Cells[7].Value.ToString());
				stExec = "Exec CreaFacturaDetalleTelco " + dIva2.ToString("0.00", us)+ ",1, '" + stArticulo + "',12, "
					+ dSubTotalIva.ToString("0.00", us) ; 
				
				Funcion.EjecutaSQL(cdsCliente, stExec, true);
			}
		//Ingresa totales en facturas posteriores a la inicial
		
			string stSelect = "EXEC FacturaTotalesTELCO " + IdCompraInicio.ToString() + ", 0";
			Funcion.EjecutaSQL(cdsCliente, stSelect, true);
	
//		stSelect = "EXEC FacturasIce " + IdCompraInicio.ToString();
//		Funcion.EjecutaSQL(cdsCliente, stSelect, true);
			MessageBox.Show("Lineas Subidas: " + iLineas.ToString());
		}

		private void TelcoSan_Load(object sender, System.EventArgs e)
		{
			string stAudita = "Exec AuditaCrear 53, 6, 'Telcosan'";
			Funcion.EjecutaSQL(cdsCliente, stAudita, true);
			cmbTipo.Value = 0;		
		}

		private void cdsCliente_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			cdsCliente.Schema.Connections[0].ConnectionString = 
				miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);
		}

		private void btAsiento_Click(object sender, System.EventArgs e)
		{
			string stNombre = cmbTipo.Text;
			if (DialogResult.No == MessageBox.Show("¿Desea Genera Asientos de " + stNombre
				+ " generadas en el sistema SARII?", "Confirmación", MessageBoxButtons.YesNo
				, MessageBoxIcon.Question)) return;
			string stArchivo = txtArchAsiento.Text;
			string stSelect =	"Select Count(*) From Asiento";
			int iCuenta = Funcion.iEscalarSQL(cdsCliente, stSelect, true);
			string stTipo = "TelcolSamAsiento";
			if ((int) cmbTipo.Value == 1) stTipo = "TelcolSamRecaudador";
			stSelect = "EXEC " + stTipo;
			Funcion.EjecutaSQL(cdsCliente, stSelect, true);
			stSelect =	"Select Count(*) From Asiento";
			iCuenta = Funcion.iEscalarSQL(cdsCliente, stSelect, true) - iCuenta;
			MessageBox.Show("Creados " + iCuenta.ToString() + " Asientos", "Información");
		}

		private void ultraButton3_Click(object sender, System.EventArgs e)
		{
			string strConn ="Provider=Microsoft.Jet.OLEDB.4.0;" +
				"Data Source=" + txtPago.Text + ";" +
				"Extended Properties='Excel 8.0;'";
			OleDbDataAdapter myData = new OleDbDataAdapter("SELECT * FROM [Trafico$]", strConn);
			myData.TableMappings.Add("Tabla", "ExcelSube");
			miDataSet.Clear();
			try
			{
				myData.Fill(miDataSet);
			}
			catch(Exception ex)
			{
				MessageBox.Show("Debe tener abierto el archivo " + txtPago.Text + " en Exel",
					ex.Message);
				return;
			}
			ultraGrid1.DataSource = miDataSet.Tables[0].DefaultView;
			ultraGrid1.DataBind();
		}

		private void ultraButton2_Click(object sender, System.EventArgs e)
		{
			int iLineas = 0;
			if (chkBorrarError.Checked)
			{
				string stBorra = "Delete ErrorSalida";
				Funcion.EjecutaSQL(cdsCliente, stBorra, true);
			}

			foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.ultraGrid1.Rows.All)
			{
				if (dr.Cells[1].Value.ToString() != "Total Recaudador") continue;
				string stCliente = dr.Cells[0].Value.ToString();
				string stNumero = dr.Cells[2].Value.ToString();
				string stFecha = "01/01/1900";
				if (dr.Cells[4].Value != DBNull.Value)
					dr.Cells[4].Value.ToString().Substring(0, 10);
				DateTime dtFecha = Funcion.CreaFecha(stFecha);
				double dPago = double.Parse(dr.Cells[7].Value.ToString());
				string stCreaCliente = "";
				if (chkCreaCliente.Checked) stCreaCliente = ", 1";
				string stSelect = "Exec CreaPago '', 1, '" + stCliente + "', "
          + dPago.ToString("0.00", us) + ", 'P" + stNumero + "', '"
					+ dtFecha.ToString("yyyyMMdd") + "'" + stCreaCliente;
				Funcion.EjecutaSQL(cdsCliente, stSelect, true);
				iLineas ++;
			}
			MessageBox.Show("Lineas Procesadas: " + iLineas.ToString(), "Información");
		}

		private void btErrores_Click(object sender, System.EventArgs e)
		{
			TesterError miError = new TesterError();
			miError.MdiParent = this.MdiParent;
			miError.Show();
		}

		private void TelcoSan_Closing(object sender, System.ComponentModel.CancelEventArgs e)
		{
			string stAudita = "Exec AuditaCrear 53, 7, 'Telcosan'";
			Funcion.EjecutaSQL(cdsCliente, stAudita, true);
		}
	}
}
