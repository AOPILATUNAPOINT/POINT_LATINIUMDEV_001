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
	/// Descripción breve de Seprofin.
	/// </summary>
	public class Seprofin : DevExpress.XtraEditors.XtraForm
	{
		private Infragistics.Win.UltraWinGrid.UltraGrid ultraGrid1;
		private C1.Data.C1DataSet cdsCliente;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtPago;
		private Infragistics.Win.Misc.UltraButton btExcel;
		private Infragistics.Win.Misc.UltraButton btPago;
		/// <summary>
		/// Variable del diseñador requerida.
		/// </summary>
		private System.ComponentModel.Container components = null;
		DataSet miDataSet = new DataSet();

		public Seprofin()
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
			this.ultraGrid1 = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.cdsCliente = new C1.Data.C1DataSet();
			this.btExcel = new Infragistics.Win.Misc.UltraButton();
			this.txtPago = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.btPago = new Infragistics.Win.Misc.UltraButton();
			((System.ComponentModel.ISupportInitialize)(this.ultraGrid1)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsCliente)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtPago)).BeginInit();
			this.SuspendLayout();
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
			this.ultraGrid1.Location = new System.Drawing.Point(24, 56);
			this.ultraGrid1.Name = "ultraGrid1";
			this.ultraGrid1.Size = new System.Drawing.Size(472, 312);
			this.ultraGrid1.TabIndex = 3;
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
			// btExcel
			// 
			this.btExcel.Location = new System.Drawing.Point(24, 16);
			this.btExcel.Name = "btExcel";
			this.btExcel.TabIndex = 8;
			this.btExcel.Text = "Excel";
			this.btExcel.Click += new System.EventHandler(this.ultraButton3_Click);
			// 
			// txtPago
			// 
			this.txtPago.Location = new System.Drawing.Point(112, 16);
			this.txtPago.Name = "txtPago";
			this.txtPago.Size = new System.Drawing.Size(128, 21);
			this.txtPago.TabIndex = 9;
			this.txtPago.Text = "C:\\Latinium\\seprofin.xls";
			// 
			// btPago
			// 
			this.btPago.Location = new System.Drawing.Point(256, 16);
			this.btPago.Name = "btPago";
			this.btPago.TabIndex = 10;
			this.btPago.Text = "Pago";
			this.btPago.Click += new System.EventHandler(this.btPago_Click);
			// 
			// Seprofin
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 13);
			this.ClientSize = new System.Drawing.Size(528, 413);
			this.Controls.Add(this.btPago);
			this.Controls.Add(this.txtPago);
			this.Controls.Add(this.btExcel);
			this.Controls.Add(this.ultraGrid1);
			this.Name = "Seprofin";
			this.Text = "Seprofin";
			this.Load += new System.EventHandler(this.Seprofin_Load);
			((System.ComponentModel.ISupportInitialize)(this.ultraGrid1)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsCliente)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtPago)).EndInit();
			this.ResumeLayout(false);

		}
		#endregion

		private void ultraButton3_Click(object sender, System.EventArgs e)
		{
			string strConn ="Provider=Microsoft.Jet.OLEDB.4.0;" +
				"Data Source=" + txtPago.Text + ";" +
				"Extended Properties='Excel 8.0;'";
			OleDbDataAdapter myData = new OleDbDataAdapter("Select * FROM [seprofin$]", strConn);
			myData.TableMappings.Add("Tabla", "ExcelSube");
			miDataSet.Clear();
			try
			{
				myData.Fill(miDataSet);
			}
			catch(Exception ex)
			{
				MessageBox.Show(ex.Message, "Error al abrir la Hoja seprofin del Archivo: " + txtPago.Text);
				return;
			}
			ultraGrid1.DataSource = miDataSet.Tables[0].DefaultView;
			ultraGrid1.DataBind();
		}

		private void btPago_Click(object sender, System.EventArgs e)
		{
			int iLineas = 0;

			foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.ultraGrid1.Rows.All)
			{
				try
				{
					if (dr.Cells[0].Value == DBNull.Value) continue;
					//campo mandatorio ??????
					string stBeneficiario = dr.Cells[0].Value.ToString();
					string stFuneraria = dr.Cells[1].Value.ToString();
					string stRuc = dr.Cells[2].Value.ToString();
					string stDireccion = dr.Cells[3].Value.ToString();
					string stCiudad = dr.Cells[4].Value.ToString();
					string stTelefono = dr.Cells[5].Value.ToString();
					string stCaso = dr.Cells[6].Value.ToString();
					string stDifunto = dr.Cells[7].Value.ToString();
					string stFacturaNota = dr.Cells[8].Value.ToString();
					string stSerie = dr.Cells[9].Value.ToString();
					string stNum = dr.Cells[10].Value.ToString();
					string stSri = dr.Cells[11].Value.ToString();
					DateTime dtEmision = DateTime.Today;
					if (dr.Cells[12].Value != DBNull.Value)
						dtEmision = (DateTime) dr.Cells[12].Value;
					double dSin_Iva = double.Parse(dr.Cells[13].Value.ToString());
					double dCon_Iva = double.Parse(dr.Cells[14].Value.ToString());
					double dValor = double.Parse(dr.Cells[15].Value.ToString());
					double dIvaRet = double.Parse(dr.Cells[16].Value.ToString());
					double dFuente = double.Parse(dr.Cells[17].Value.ToString());
					double dIvaFuente = double.Parse(dr.Cells[18].Value.ToString());
					double dRetFuente = double.Parse(dr.Cells[19].Value.ToString());
					double dCancelar = double.Parse(dr.Cells[20].Value.ToString());
					string stCodArticulo = dr.Cells[21].Value.ToString();
					int iCreditoContado = int.Parse(dr.Cells[22].Value.ToString());
					string stDebito = dr.Cells[23].Value.ToString();
					string stCredito = dr.Cells[24].Value.ToString();
					string stConcepto = dr.Cells[25].Value.ToString();
					DateTime dtPago = DateTime.Today;
					if (dr.Cells[26].Value != DBNull.Value)
						dtPago = (DateTime) dr.Cells[26].Value;
					string stCodAsi = dr.Cells[27].Value.ToString();
					string stAutRet = dr.Cells[28].Value.ToString();
					string stSerieRet = dr.Cells[29].Value.ToString();
					string stCodRet = dr.Cells[30].Value.ToString();
					string stCodConIva = dr.Cells[31].Value.ToString();
					DateTime dtCaduca = DateTime.Today;
					if (dr.Cells[32].Value != DBNull.Value)
						dtCaduca = (DateTime) dr.Cells[32].Value;
					string stCodAsiento = dr.Cells[33].Value.ToString();

					string stSelect = String.Format(
"Exec Seprofin '{0}', '{1}', '{2}', '{3}', '{4}', '{5}', '{6}', '{7}', '{8}', '{9}', '{10}', '{11}', '{12}', {13}, {14}, {15}, {16}, {17}, {18}, {19}, {20}, '{21}', {22}, '{23}', '{24}', '{25}', '{26}', '{27}', '{28}', '{29}', '{30}', '{31}', '{32}', '{33}'",
            stBeneficiario, stFuneraria, stRuc, stDireccion, stCiudad,
						stTelefono, stCaso, stDifunto, stFacturaNota, stSerie,
						stNum, stSri, dtEmision.ToString("yyyyMMdd"), dSin_Iva.ToString("0.00", us), dCon_Iva.ToString("0.00", us),
            dValor.ToString("0.00", us), dIvaRet.ToString("0.00", us), dIvaFuente.ToString("0.00", us), dFuente.ToString("0.00", us),	dRetFuente.ToString("0.00", us),
            dCancelar.ToString("0.00", us), stCodArticulo, iCreditoContado,	stDebito, stCredito,
						stConcepto, dtPago.ToString("yyyyMMdd"), stCodAsi, stAutRet, stSerieRet, stCodRet, stCodConIva, dtCaduca.ToString("yyyyMMdd"), stCodAsiento);

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

		private void cdsCliente_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			cdsCliente.Schema.Connections[0].ConnectionString = 
				miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);
		}
		CultureInfo us = new CultureInfo("en-US");

		private void Seprofin_Load(object sender, System.EventArgs e)
		{
		
		}
	}
}
