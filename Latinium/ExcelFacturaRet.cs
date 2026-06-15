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
	public class ExcelFacturaRet : DevExpress.XtraEditors.XtraForm
	{
		private Infragistics.Win.UltraWinGrid.UltraGrid ultraGrid1;
		private C1.Data.C1DataSet cdsCliente;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtPago;
		private Infragistics.Win.Misc.UltraButton btExcel;
		/// <summary>
		/// Variable del diseñador requerida.
		/// </summary>
		private System.ComponentModel.Container components = null;
		private Infragistics.Win.Misc.UltraButton btCompras;
		private Infragistics.Win.Misc.UltraButton btVentas;
		private Infragistics.Win.Misc.UltraButton ultraButton1;
		DataSet miDataSet = new DataSet();

		public ExcelFacturaRet()
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
			this.btCompras = new Infragistics.Win.Misc.UltraButton();
			this.btVentas = new Infragistics.Win.Misc.UltraButton();
			this.ultraButton1 = new Infragistics.Win.Misc.UltraButton();
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
			this.ultraGrid1.Location = new System.Drawing.Point(16, 56);
			this.ultraGrid1.Name = "ultraGrid1";
			this.ultraGrid1.Size = new System.Drawing.Size(656, 368);
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
			this.btExcel.Size = new System.Drawing.Size(88, 25);
			this.btExcel.TabIndex = 8;
			this.btExcel.Text = "Leer Compras";
			this.btExcel.Click += new System.EventHandler(this.ultraButton3_Click);
			// 
			// txtPago
			// 
			this.txtPago.Location = new System.Drawing.Point(216, 16);
			this.txtPago.Name = "txtPago";
			this.txtPago.Size = new System.Drawing.Size(192, 22);
			this.txtPago.TabIndex = 9;
			this.txtPago.Text = "C:\\Latinium\\Fact_Ret.xls";
			// 
			// btCompras
			// 
			this.btCompras.Location = new System.Drawing.Point(504, 16);
			this.btCompras.Name = "btCompras";
			this.btCompras.Size = new System.Drawing.Size(75, 25);
			this.btCompras.TabIndex = 10;
			this.btCompras.Text = "Compras";
			this.btCompras.Click += new System.EventHandler(this.btPago_Click);
			// 
			// btVentas
			// 
			this.btVentas.Location = new System.Drawing.Point(584, 16);
			this.btVentas.Name = "btVentas";
			this.btVentas.Size = new System.Drawing.Size(75, 25);
			this.btVentas.TabIndex = 11;
			this.btVentas.Text = "Ventas";
			this.btVentas.Click += new System.EventHandler(this.btVentas_Click);
			// 
			// ultraButton1
			// 
			this.ultraButton1.Location = new System.Drawing.Point(120, 16);
			this.ultraButton1.Name = "ultraButton1";
			this.ultraButton1.Size = new System.Drawing.Size(88, 25);
			this.ultraButton1.TabIndex = 12;
			this.ultraButton1.Text = "Leer Ventas";
			this.ultraButton1.Click += new System.EventHandler(this.ultraButton1_Click);
			// 
			// ExcelFacturaRet
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 14);
			this.ClientSize = new System.Drawing.Size(688, 444);
			this.Controls.Add(this.ultraButton1);
			this.Controls.Add(this.btVentas);
			this.Controls.Add(this.btCompras);
			this.Controls.Add(this.txtPago);
			this.Controls.Add(this.btExcel);
			this.Controls.Add(this.ultraGrid1);
			this.Name = "ExcelFacturaRet";
			this.Text = "Factura - Retención";
			((System.ComponentModel.ISupportInitialize)(this.ultraGrid1)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsCliente)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtPago)).EndInit();
			this.ResumeLayout(false);

		}
		#endregion

		private void ultraButton3_Click(object sender, System.EventArgs e)
		{
			LeerExcel("Compras");
		}

		private void btPago_Click(object sender, System.EventArgs e)
		{
			GeneraFactura(4);
		}

		private void cdsCliente_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			cdsCliente.Schema.Connections[0].ConnectionString = 
				miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);
		}

		CultureInfo us = new CultureInfo("en-US");
		private void GeneraFactura(int idTipoFactura)
		{
			int iLineas = 0;

			foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.ultraGrid1.Rows.All)
			{
				try
				{
					if (dr.Cells[0].Value == DBNull.Value) continue;
					//campo mandatorio ??????
					DateTime dtFecha = DateTime.Parse(dr.Cells[0].Value.ToString());
					string stSerieFact = dr.Cells[1].Value.ToString();
					string stFactura = dr.Cells[2].Value.ToString();
					string stAutorizaFact = dr.Cells[3].Value.ToString();
					string stRuc = dr.Cells[4].Value.ToString();
					string stCliente = dr.Cells[5].Value.ToString();
					string stConcepto = dr.Cells[6].Value.ToString();
					double dSin_Iva = dr.Cells[7].Value == DBNull.Value ? 0: double.Parse(dr.Cells[7].Value.ToString());
					double dCon_Iva = dr.Cells[8].Value == DBNull.Value ? 0: double.Parse(dr.Cells[8].Value.ToString());
					double dIva = dr.Cells[9].Value == DBNull.Value ? 0: double.Parse(dr.Cells[9].Value.ToString());
					double dTotal = dr.Cells[10].Value == DBNull.Value ? 0: double.Parse(dr.Cells[10].Value.ToString());
					string stSerieRet = dr.Cells[11].Value.ToString();
					string stNumRet = dr.Cells[12].Value.ToString();
					string stAutorizaRet = dr.Cells[13].Value.ToString();
					string stCodRetSRI = dr.Cells[14].Value.ToString();
					double dPorcRenta = dr.Cells[15].Value == DBNull.Value ? 0: double.Parse(dr.Cells[15].Value.ToString());
					double dRentaRet = dr.Cells[16].Value == DBNull.Value ? 0: double.Parse(dr.Cells[16].Value.ToString());
					double dPorcIva = dr.Cells[17].Value == DBNull.Value ? 0: double.Parse(dr.Cells[17].Value.ToString());
					double dIvaRet = dr.Cells[18].Value == DBNull.Value ? 0: double.Parse(dr.Cells[18].Value.ToString());
					double dContadoCredito = dr.Cells[19].Value == DBNull.Value ? 1.0: double.Parse(dr.Cells[19].Value.ToString());
					double dCreditoTrib = dr.Cells[20].Value == DBNull.Value ? 1.0: double.Parse(dr.Cells[20].Value.ToString());
					DateTime dtFechaCad = dr.Cells[21].Value == DBNull.Value ? DateTime.Today: DateTime.Parse(dr.Cells[21].Value.ToString());

					string stSelect = String.Format(
						"Exec Fact_Ret '{0}', '{1}', '{2}', '{3}', '{4}', '{5}', '{6}', {7}, {8}, {9}, {10}, '{11}', '{12}', '{13}', '{14}', {15}, {16}, {17}, {18}, {19}, {20}, {21}, '{22}'",
						dtFecha.ToString("yyyyMMdd"), stSerieFact, stFactura, stAutorizaFact, stRuc, 
						stCliente, stConcepto, dSin_Iva.ToString("0.00", us), dCon_Iva.ToString("0.00", us), 
						dIva.ToString("0.00", us), dTotal.ToString("0.00", us), stSerieRet, stNumRet, stAutorizaRet, stCodRetSRI,
						dPorcRenta.ToString("0.00", us), dRentaRet.ToString("0.00", us), dPorcIva.ToString("0.00", us), 
						dIvaRet.ToString("0.00", us), idTipoFactura, dContadoCredito.ToString("0"), dCreditoTrib.ToString("0"), 
						dtFechaCad.ToString("yyyyMMdd"));

					int idCompra = Funcion.iEscalarSQL(cdsCliente, stSelect, true);
					if (idCompra > 0)
					{
						stSelect = string.Format("Exec Fact_Ret_Detalle {0}, {1}, {2}", idCompra,
							dSin_Iva.ToString("0.00", us), dCon_Iva.ToString("0.00", us));
						Funcion.EjecutaSQL(cdsCliente, stSelect, true);
						iLineas ++;
					}
				}
				catch (Exception ex)
				{
					MessageBox.Show(ex.Message, "Error en linea " + iLineas.ToString());
				}
			}
			MessageBox.Show("Lineas Procesadas: " + iLineas.ToString(), "Información");
		}

		private void btVentas_Click(object sender, System.EventArgs e)
		{
			GeneraFactura(1);
		}

		private void LeerExcel(string stTipo)
		{
			string strConn ="Provider=Microsoft.Jet.OLEDB.4.0;" +
				"Data Source=" + txtPago.Text + ";" +
				"Extended Properties='Excel 8.0;'";
			string stHoja = string.Format("Select * FROM [{0}$]", stTipo);
			OleDbDataAdapter myData = new OleDbDataAdapter(stHoja, strConn);
			myData.TableMappings.Add("Tabla", "ExcelSube");
			miDataSet.Clear();
			try
			{
				myData.Fill(miDataSet);
			}
			catch(Exception ex)
			{
				MessageBox.Show(ex.Message, "Error al abrir la Hoja Ventas del Archivo: " + txtPago.Text);
				return;
			}
			ultraGrid1.DataSource = miDataSet.Tables[0].DefaultView;
			ultraGrid1.DataBind();
		}

		private void ultraButton1_Click(object sender, System.EventArgs e)
		{
			LeerExcel("Ventas");
		}
	}
}
