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
	public class Fidelius : DevExpress.XtraEditors.XtraForm
	{
		private Infragistics.Win.UltraWinGrid.UltraGrid ultraGrid1;
		private C1.Data.C1DataSet cdsCliente;
		/// <summary>
		/// Variable del diseñador requerida.
		/// </summary>
		private System.ComponentModel.Container components = null;
		private Infragistics.Win.Misc.UltraButton btAsiento;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtArchAsiento;
		private Infragistics.Win.Misc.UltraButton ultraButton2;
		private Infragistics.Win.Misc.UltraButton ultraButton3;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtPago;
		CultureInfo us = new CultureInfo("en-US");
		private Infragistics.Win.UltraWinEditors.UltraCheckEditor chkCreaCliente;
		private Infragistics.Win.UltraWinEditors.UltraCheckEditor chkBorrarError;
		private System.Windows.Forms.OpenFileDialog openFileDialog;
		private Infragistics.Win.Misc.UltraButton txtDevolucion;
		private Infragistics.Win.Misc.UltraButton btnExcelAsi;


		public Fidelius()
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
			this.btAsiento = new Infragistics.Win.Misc.UltraButton();
			this.txtArchAsiento = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.ultraButton2 = new Infragistics.Win.Misc.UltraButton();
			this.txtPago = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.ultraButton3 = new Infragistics.Win.Misc.UltraButton();
			this.chkCreaCliente = new Infragistics.Win.UltraWinEditors.UltraCheckEditor();
			this.chkBorrarError = new Infragistics.Win.UltraWinEditors.UltraCheckEditor();
			this.openFileDialog = new System.Windows.Forms.OpenFileDialog();
			this.btnExcelAsi = new Infragistics.Win.Misc.UltraButton();
			this.txtDevolucion = new Infragistics.Win.Misc.UltraButton();
			((System.ComponentModel.ISupportInitialize)(this.ultraGrid1)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsCliente)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtArchAsiento)).BeginInit();
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
			this.ultraGrid1.Location = new System.Drawing.Point(8, 96);
			this.ultraGrid1.Name = "ultraGrid1";
			this.ultraGrid1.Size = new System.Drawing.Size(448, 144);
			this.ultraGrid1.TabIndex = 2;
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
			this.btAsiento.Location = new System.Drawing.Point(272, 8);
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
			this.txtArchAsiento.Text = "C:\\Fidelius2.xls";
			this.txtArchAsiento.ValueChanged += new System.EventHandler(this.txtArchAsiento_ValueChanged);
			// 
			// ultraButton2
			// 
			this.ultraButton2.Location = new System.Drawing.Point(272, 32);
			this.ultraButton2.Name = "ultraButton2";
			this.ultraButton2.TabIndex = 9;
			this.ultraButton2.Text = "Facturas";
			this.ultraButton2.Click += new System.EventHandler(this.ultraButton2_Click);
			// 
			// txtPago
			// 
			this.txtPago.Location = new System.Drawing.Point(136, 32);
			this.txtPago.Name = "txtPago";
			this.txtPago.Size = new System.Drawing.Size(128, 21);
			this.txtPago.TabIndex = 8;
			this.txtPago.Text = "C:\\Fidelius1.Xls";
			this.txtPago.ValueChanged += new System.EventHandler(this.txtPago_ValueChanged);
			// 
			// ultraButton3
			// 
			this.ultraButton3.Location = new System.Drawing.Point(8, 32);
			this.ultraButton3.Name = "ultraButton3";
			this.ultraButton3.Size = new System.Drawing.Size(88, 23);
			this.ultraButton3.TabIndex = 7;
			this.ultraButton3.Text = "Excel Factura";
			this.ultraButton3.Click += new System.EventHandler(this.ultraButton3_Click);
			// 
			// chkCreaCliente
			// 
			this.chkCreaCliente.Location = new System.Drawing.Point(168, 56);
			this.chkCreaCliente.Name = "chkCreaCliente";
			this.chkCreaCliente.Size = new System.Drawing.Size(88, 20);
			this.chkCreaCliente.TabIndex = 10;
			this.chkCreaCliente.Text = "Crea Cliente";
			// 
			// chkBorrarError
			// 
			this.chkBorrarError.Location = new System.Drawing.Point(80, 56);
			this.chkBorrarError.Name = "chkBorrarError";
			this.chkBorrarError.Size = new System.Drawing.Size(88, 20);
			this.chkBorrarError.TabIndex = 12;
			this.chkBorrarError.Text = "Borrar Error";
			// 
			// btnExcelAsi
			// 
			this.btnExcelAsi.Location = new System.Drawing.Point(8, 8);
			this.btnExcelAsi.Name = "btnExcelAsi";
			this.btnExcelAsi.Size = new System.Drawing.Size(88, 23);
			this.btnExcelAsi.TabIndex = 13;
			this.btnExcelAsi.Text = "Excel Asiento";
			this.btnExcelAsi.Click += new System.EventHandler(this.btnExcelAsi_Click);
			// 
			// txtDevolucion
			// 
			this.txtDevolucion.Location = new System.Drawing.Point(272, 56);
			this.txtDevolucion.Name = "txtDevolucion";
			this.txtDevolucion.TabIndex = 14;
			this.txtDevolucion.Text = "Devolucion";
			this.txtDevolucion.Click += new System.EventHandler(this.txtDevolucion_Click);
			// 
			// Fidelius
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 13);
			this.ClientSize = new System.Drawing.Size(464, 254);
			this.Controls.Add(this.txtDevolucion);
			this.Controls.Add(this.btnExcelAsi);
			this.Controls.Add(this.chkBorrarError);
			this.Controls.Add(this.chkCreaCliente);
			this.Controls.Add(this.ultraButton2);
			this.Controls.Add(this.txtPago);
			this.Controls.Add(this.ultraButton3);
			this.Controls.Add(this.txtArchAsiento);
			this.Controls.Add(this.btAsiento);
			this.Controls.Add(this.ultraGrid1);
			this.Name = "Fidelius";
			this.Text = "Micro Fidelius";
			this.Closing += new System.ComponentModel.CancelEventHandler(this.TelcoSan_Closing);
			this.Load += new System.EventHandler(this.TelcoSan_Load);
			((System.ComponentModel.ISupportInitialize)(this.ultraGrid1)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsCliente)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtArchAsiento)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtPago)).EndInit();
			this.ResumeLayout(false);

		}
		#endregion

		DataSet miDataSet = new DataSet();

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

		private int Cuentita (string stNumero)
		{
			string stSelect = "Select Idcuenta From cuenta Where codigo = '"
				+ stNumero.Trim() + "'";
			int IdCuenta = Funcion.iEscalarSQL(cdsCliente, stSelect, true);
			return IdCuenta;
		}
		private void btLee_Click(object sender, System.EventArgs e)
		{
	   	}

		private void TelcoSan_Load(object sender, System.EventArgs e)
		{
			string stAudita = "Exec AuditaCrear 53, 6, 'Telcosan'";
			Funcion.EjecutaSQL(cdsCliente, stAudita, true);
		
		}

		private void cdsCliente_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			cdsCliente.Schema.Connections[0].ConnectionString = 
				miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);
		}

		private void btAsiento_Click(object sender, System.EventArgs e)
		{
			Cursor =Cursors.WaitCursor;
			int banini=1;
			int lineas =0;
			foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.ultraGrid1.Rows.All)
			{
				if (banini==0)
				{
	//				CodVenta= ultraGrid1.DisplayLayout.Bands[0].Columns[4].ToString();      
	//				CodVenta=CodVenta.Replace("#","."); 
					banini=1;
				}
				else
				{
					string Cantidad = dr.Cells[0].Value.ToString();																	
					string CheckDesde = dr.Cells[1].Value.ToString();																	
					string CheckHasta = dr.Cells[2].Value.ToString();	
					
					string stFecha = "01/01/1900";
					if (dr.Cells[3].Value != DBNull.Value)
						stFecha=dr.Cells[3].Value.ToString().Substring(0, 10);
					stFecha =stFecha.Trim(); 
					DateTime dtFecha = Funcion.CreaFecha(stFecha);
					
					double Venta = double.Parse(dr.Cells[4].Value.ToString());																	
					double Iva = double.Parse(dr.Cells[5].Value.ToString());																	
					double Servicio = double.Parse(dr.Cells[6].Value.ToString());																	
					double Propina = double.Parse(dr.Cells[18].Value.ToString());																	
				
					double TotalDes = double.Parse(dr.Cells[7].Value.ToString());																	
					double Efectivo = double.Parse(dr.Cells[8].Value.ToString());																	
					double Cheque = double.Parse(dr.Cells[9].Value.ToString());																	
					double Diners = double.Parse(dr.Cells[10].Value.ToString());																	
					double American = double.Parse(dr.Cells[11].Value.ToString());																	
					double Mastercard = double.Parse(dr.Cells[13].Value.ToString());																	
					double Visa = double.Parse(dr.Cells[14].Value.ToString());																	
					double Canje = double.Parse(dr.Cells[15].Value.ToString());																	
					double Credito = double.Parse(dr.Cells[16].Value.ToString());																	
					double Retencion = double.Parse(dr.Cells[17].Value.ToString());																	
				
					string detallito="Ventas del " + dtFecha.ToString("dd-MMM-yyyy");
					
					string stSelect="select count(*) from asiento where descripcion='" + detallito + "' and fecha=' "+dtFecha.ToString("yyyyMMdd")+"'";  
					int cuantos=Funcion.iEscalarSQL(cdsCliente,stSelect);
					if(cuantos==0)
					{
						lineas++;
						stSelect = "Exec AsientoGenera 1,'" + detallito.ToString() + "',@fecha='" +dtFecha.ToString("yyyyMMdd")+"'";
						int IdAsiento= Funcion.iEscalarSQL(cdsCliente , stSelect, true);

						stSelect = "Exec AsientoDetalleGenera "
							+ Efectivo.ToString("0.00", us) + ",0,-7," + IdAsiento.ToString() + ", @Detalle = 'Efectivo'"+ ", @Parametro = 1";
						Funcion.EjecutaSQL(cdsCliente, stSelect, true);

						stSelect = "Exec AsientoDetalleGenera "
							+ Cheque.ToString("0.00", us) + ",0,-7,"  + IdAsiento.ToString() + ", @Detalle = 'Cheque'"+ ", @Parametro = 2";
						Funcion.EjecutaSQL(cdsCliente, stSelect, true);

						stSelect = "Exec AsientoDetalleGenera "
							+ Diners.ToString("0.00", us) + ",0,-7, " + IdAsiento.ToString() + ", @Detalle = 'Diners'"+ ", @Parametro = 3";
						Funcion.EjecutaSQL(cdsCliente, stSelect, true);

						stSelect = "Exec AsientoDetalleGenera "
							+ American.ToString("0.00", us) + ",0,-7, "  + IdAsiento.ToString() + ", @Detalle = 'American'"+ ", @Parametro = 5";
						Funcion.EjecutaSQL(cdsCliente, stSelect, true);
						stSelect = "Exec AsientoDetalleGenera "
							+ Mastercard.ToString("0.00", us) + ",0,-7, "  + IdAsiento.ToString() + ", @Detalle = 'Mastercard'"+ ", @Parametro = 4";
						Funcion.EjecutaSQL(cdsCliente, stSelect, true);
						stSelect = "Exec AsientoDetalleGenera "
							+ Visa.ToString("0.00", us) + ",0,-7, "  + IdAsiento.ToString() + ", @Detalle = 'Visa'"+ ", @Parametro = 6";
						Funcion.EjecutaSQL(cdsCliente, stSelect, true);
						stSelect = "Exec AsientoDetalleGenera "
							+ Canje.ToString("0.00", us) + ",0,-7, " + IdAsiento.ToString() + ", @Detalle = 'Canje'"+ ", @Parametro = 10";
						Funcion.EjecutaSQL(cdsCliente, stSelect, true);
						stSelect = "Exec AsientoDetalleGenera "
							+ Credito.ToString("0.00", us) + ",0,-7, "  + IdAsiento.ToString() + ", @Detalle = 'Credito'"+ ", @Parametro = 7";
						Funcion.EjecutaSQL(cdsCliente, stSelect, true);
						stSelect = "Exec AsientoDetalleGenera "
							+ Retencion.ToString("0.00", us) + ",0,-8, "  + IdAsiento.ToString() + ", @Detalle = 'Retencion'"+ ", @Parametro = 5";
						Funcion.EjecutaSQL(cdsCliente, stSelect, true);

						stSelect = "Exec AsientoDetalleGenera "
							+ TotalDes.ToString("0.00", us) + ",0,-5, "  + IdAsiento.ToString() + ", @Detalle = 'Descuento'"+ ", @Parametro = 3";
						Funcion.EjecutaSQL(cdsCliente, stSelect, true);


//TotalDes 

						////////////////// son los del haber: venta iva servicio propina
						stSelect = "Exec AsientoDetalleGenera 0,"
							+ Venta.ToString("0.00", us) + ",-5, " + IdAsiento.ToString() + ", @Detalle = 'Venta'"+ ", @Parametro = 5";
						Funcion.EjecutaSQL(cdsCliente, stSelect, true);

						stSelect = "Exec AsientoDetalleGenera 0,"
							+ Iva.ToString("0.00", us) + ",-5 ,"  + IdAsiento.ToString() + ", @Detalle = 'Iva'"+ ", @Parametro = 2";
						Funcion.EjecutaSQL(cdsCliente, stSelect, true);

						stSelect = "Exec AsientoDetalleGenera 0,"
							+ Servicio.ToString("0.00", us) + ",-5, "  + IdAsiento.ToString() + ", @Detalle = 'Servicio'"+ ", @Parametro = 41";
						Funcion.EjecutaSQL(cdsCliente, stSelect, true);

						stSelect = "Exec AsientoDetalleGenera 0,"
							+ Propina.ToString("0.00", us) + ",-5 ,"  + IdAsiento.ToString() + ", @Detalle = 'Propina'"+ ", @Parametro = 42";
						Funcion.EjecutaSQL(cdsCliente, stSelect, true);
				
						//aqui debo actualizar el asiento con los totales
						stSelect = "Exec AsientoTotal " + IdAsiento.ToString();
						Funcion.EjecutaSQL(cdsCliente, stSelect, true);
					}
					else
					{
						MessageBox.Show("Este asiento ya existe");  
					}
				}
			}
			Cursor =Cursors.Default ;
			MessageBox.Show("Se han subido " + lineas + " asientos");  
		
		}

		private void ultraButton3_Click(object sender, System.EventArgs e)
		{
			//boton de facturas para subir del excel a la grilla

			string strConn ="Provider=Microsoft.Jet.OLEDB.4.0;" +
				"Data Source=" + txtPago.Text + ";" +
				"Extended Properties='Excel 8.0;'";
			OleDbDataAdapter myData = new OleDbDataAdapter
			("select [Check Micros]AS FACTURA, MAX(CLIENTE)AS CLIENTE ,MAX([PRODUCTO FECHA]) AS FECHA  ,SUM(CANTIDAD)AS CANTIDAD ,SUM(PRECIO) AS PRECIO,CODIGO , max(ruc) as RUC ,max(direccion)as DIRECCION,max(telefono)as TELEFONO from [FACTURA$] gROUP BY [Check Micros],CODIGO ORDER BY [Check Micros],CODIGO", strConn);
			myData.TableMappings.Add("Tabla", "ExcelSube");
			miDataSet.Clear();
			try
			{
				myData.Fill(miDataSet);
			}
			catch(Exception ex)
			{
				MessageBox.Show("Debe tener abierto el archivo " + txtPago.Text + " en ExCel",
					ex.Message);
				return;
			}
			ultraGrid1.DataSource = miDataSet.Tables[0].DefaultView;
			ultraGrid1.DataBind();
		}

		private void ultraButton2_Click(object sender, System.EventArgs e)
		{
			//crea el asiento de facturas

			Cursor = Cursors.WaitCursor;  
			string stSelect =	"Select Count(*) From compra where idtipofactura=1";
			int iCuenta = Funcion.iEscalarSQL(cdsCliente, stSelect, true);
			string stExec="";
			if (chkBorrarError.Checked)
			{
				string stBorra = "Delete ErrorSalida";
				Funcion.EjecutaSQL(cdsCliente, stBorra, true);
			}

			int FactAnt=0;
			int IdCompraInicio=0;
			int IdCompra=0;
			int IdCompraInicio1=0;
			double descuento=0;
			foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.ultraGrid1.Rows.All)
			{
				string sFactura=dr.Cells[0].Value.ToString();
				int iFactura=int.Parse (dr.Cells[0].Value.ToString());

				string sCliente = dr.Cells[1].Value.ToString();

				int iCodigo=int.Parse (dr.Cells[5].Value.ToString());
				string sCodigo = dr.Cells[5].Value.ToString();

				string stFecha = "01/01/1900";
				if (dr.Cells[2].Value != DBNull.Value)
					 stFecha=dr.Cells[2].Value.ToString().Substring(0, 10);
				DateTime dtFecha = Funcion.CreaFecha(stFecha);

				string sCantidad = dr.Cells[3].Value.ToString();
				int iCantidad    = int.Parse (dr.Cells[3].Value.ToString());
				double fTotal    = float.Parse (dr.Cells[4].Value.ToString());
				double ddes      = float.Parse (dr.Cells[4].Value.ToString());
				string sPrecio   = dr.Cells[4].Value.ToString();
				if (iCantidad==0 )
				fTotal=0;
				else
				fTotal=fTotal/iCantidad;
				

		 //  	if (dr.Cells[6].Value != DBNull.Value)
					string sRuc=dr.Cells[6].Value.ToString();
		 //		if (dr.Cells[7].Value != DBNull.Value)
					 string sDireccion=dr.Cells[7].Value.ToString();
		 // 	if(dr.Cells[8].Value != DBNull.Value)
					 string sTelefono=dr.Cells[8].Value.ToString();

				if(iFactura!=FactAnt) //es nueva factura			
				{
					descuento=0;
					stExec = "Exec CreaFacturaFidelius '" + sCliente + "', '"
						+ "', '" + sFactura + "', '" + dtFecha.ToString("yyyyMMdd") 
						+"','" +sRuc +"','" + sDireccion +"','" + sTelefono
						+ "',0,0,0,0,0"
						+ ", 1,0,1,1 ";
					IdCompra = Funcion.iEscalarSQL(cdsCliente, stExec, true);
				}
				FactAnt=iFactura;
				if(IdCompra>0)
				{
					if (IdCompraInicio==0)
						IdCompraInicio=IdCompra;
						
					/// para sumar el descuento
				if(iCodigo>=101 && iCodigo<=104)	
				{
					descuento = descuento + Math.Abs(ddes) ;
					stExec = "update compra set desc1= "
						+ descuento.ToString("0.00", us) + " where idcompra=" +IdCompra.ToString() ;
					Funcion.EjecutaSQL(cdsCliente, stExec, true);
	
				}
				else					
				{
					stExec = "Exec CreaFacturaDetalle " + sCantidad + ",'" + sCodigo + "', 12, "
					+ fTotal.ToString("0.00", us) + ",@IdCompra=" +IdCompra.ToString() ;
					Funcion.EjecutaSQL(cdsCliente, stExec, true);
				}

				}
			}//foreach
		
			//ingreso del consumo real
			foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr1 in this.ultraGrid1.Rows.All)
			{
				int iFactura=int.Parse(dr1.Cells[0].Value.ToString());
				stSelect =" select count(*) from compra where idtipofactura=10 and numero='" + iFactura.ToString()+ "'"; 
				int cuantos = Funcion.iEscalarSQL(cdsCliente,stSelect);

				if(cuantos==0)
				{
					stExec="Select IdCompra From Compra where numero='" + iFactura.ToString() + "'";
					IdCompraInicio1=Funcion.iEscalarSQL(cdsCliente,stExec);
					Funcion.EjecutaSQL(cdsCliente , "EXEC ConsumoFacturaFidelius " 
						+ IdCompraInicio1.ToString()  + ",10 ");
				}
			}		

			if(IdCompraInicio>0)
			{
				stExec = "EXEC FacturaTotalesFidelius " + IdCompraInicio.ToString();
				Funcion.EjecutaSQL(cdsCliente, stExec, true);
			}
			stExec = "Select Count(*) From compra where idtipofactura=1";
			iCuenta = Funcion.iEscalarSQL(cdsCliente, stExec, true) - iCuenta;
			Cursor =Cursors.Default;
			MessageBox.Show("Proceso realizado exitosamente, se han creado " + iCuenta.ToString() + " Facturas con sus consumos ", " Factura");

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

		private void txtPago_ValueChanged(object sender, System.EventArgs e)
		{
		
		}

		private void txtArchAsiento_ValueChanged(object sender, System.EventArgs e)
		{
		
		}

		private void btnExcelAsi_Click(object sender, System.EventArgs e)
		{
			//boton de facturas para subir del excel a la grilla

			string strConn ="Provider=Microsoft.Jet.OLEDB.4.0;" +
				"Data Source=" + txtArchAsiento.Text + ";" +
				"Extended Properties='Excel 8.0;'";
			OleDbDataAdapter myData = new OleDbDataAdapter("select * from [ASIENTO$]", strConn);
			myData.TableMappings.Add("Tabla", "ExcelSube");
			miDataSet.Clear();
		
			try
			{
				myData.Fill(miDataSet);
			}
			catch(Exception ex)
			{
				MessageBox.Show("Debe tener abierto el archivo " + txtArchAsiento.Text + " en Excel",
					ex.Message);
				return;
			}
			ultraGrid1.DataSource = miDataSet.Tables[0].DefaultView;
			ultraGrid1.DataBind();

		}

		private void txtDevolucion_Click(object sender, System.EventArgs e)
		{
			//crea el asiento de facturas es un ingreso porque es de los negativos o sea facturas que se revierten
      /// va ingreso bodega 9
			Cursor = Cursors.WaitCursor;  
			string stSelect =	"Select Count(*) From compra where idtipofactura=5";  ///5 DEV VENTA;
			int iCuenta = Funcion.iEscalarSQL(cdsCliente, stSelect, true);
			string stExec="";
			if (chkBorrarError.Checked)
			{
				string stBorra = "Delete ErrorSalida";
				Funcion.EjecutaSQL(cdsCliente, stBorra, true);
			}

			int FactAnt=0;
			int IdCompraInicio=0;
			int IdCompra=0;
			int IdCompraInicio1=0;
			double descuento=0;
			foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in this.ultraGrid1.Rows.All)
			{
				string sFactura=dr.Cells[0].Value.ToString();
				int iFactura=int.Parse (dr.Cells[0].Value.ToString());

				string sCliente = dr.Cells[1].Value.ToString();

				int iCodigo=int.Parse (dr.Cells[5].Value.ToString());
				string sCodigo = dr.Cells[5].Value.ToString();

				string stFecha = "01/01/1900";
				if (dr.Cells[2].Value != DBNull.Value)
					stFecha=dr.Cells[2].Value.ToString().Substring(0, 10);
				DateTime dtFecha = Funcion.CreaFecha(stFecha);

				string sCantidad = dr.Cells[3].Value.ToString();
				int iCantidad    = int.Parse (dr.Cells[3].Value.ToString());
				double fTotal    = float.Parse (dr.Cells[4].Value.ToString());
				double ddes      = float.Parse (dr.Cells[4].Value.ToString());
				string sPrecio   = dr.Cells[4].Value.ToString();

				
				iCantidad=Math.Abs(iCantidad);
        sCantidad=iCantidad.ToString(); 
		
				if (iCantidad==0 )
					fTotal=0;
				else
					fTotal=Math.Abs(fTotal)/Math.Abs(iCantidad);

				//  	if (dr.Cells[6].Value != DBNull.Value)
				string sRuc=dr.Cells[6].Value.ToString();
				//		if (dr.Cells[7].Value != DBNull.Value)
				string sDireccion=dr.Cells[7].Value.ToString();
				// 	if(dr.Cells[8].Value != DBNull.Value)
				string sTelefono=dr.Cells[8].Value.ToString();

				if(iFactura!=FactAnt) //es nueva factura			
				{
					sFactura="D"+sFactura;
					descuento=0;
					stExec = "Exec CreaFacturaFidelius '" + sCliente + "', '"
						+ "', '" + sFactura + "', '" + dtFecha.ToString("yyyyMMdd") 
						+"','" +sRuc +"','" + sDireccion +"','" + sTelefono
						+ "',0,0,0,0,0"
						+ ", 5,0,1,1 ";
					IdCompra = Funcion.iEscalarSQL(cdsCliente, stExec, true);
				}
				FactAnt=iFactura;
				if(IdCompra>0)
				{
					if (IdCompraInicio==0)
						IdCompraInicio=IdCompra;
						
					/// para sumar el descuento
					if(iCodigo>=101 && iCodigo<=104)	
					{
						descuento = descuento + Math.Abs(ddes) ;
						stExec = "update compra set desc1= "
							+ descuento.ToString("0.00", us) + " where idcompra=" +IdCompra.ToString() ;
						Funcion.EjecutaSQL(cdsCliente, stExec, true);
	
					}
					else					
					{
						stExec = "Exec CreaFacturaDetalle " + sCantidad + ",'" + sCodigo + "', 12, "
							+ fTotal.ToString("0.00", us) + ",@IdCompra=" +IdCompra.ToString() ;
						Funcion.EjecutaSQL(cdsCliente, stExec, true);
					}

				}
			}//foreach
		
			//ingreso del consumo real
			foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr1 in this.ultraGrid1.Rows.All)
			{
				int iFactura=int.Parse(dr1.Cells[0].Value.ToString());
				stSelect =" select count(*) from compra where idtipofactura=9 and numero='D" + iFactura.ToString()+ "'"; 
				int cuantos = Funcion.iEscalarSQL(cdsCliente,stSelect);

				if(cuantos==0)
				{
					stExec="Select IdCompra From Compra where idtipofactura=5 and numero='D" + iFactura.ToString() + "'";
					IdCompraInicio1=Funcion.iEscalarSQL(cdsCliente,stExec);
					Funcion.EjecutaSQL(cdsCliente , "EXEC ConsumoFacturaFidelius " 
						+ IdCompraInicio1.ToString()  + ",9 ");//9 porque va a ingreso
				}
			}		

			if(IdCompraInicio>0)
			{
				stExec = "EXEC FacturaTotalesFidelius " + IdCompraInicio.ToString();
				Funcion.EjecutaSQL(cdsCliente, stExec, true);
			}
			stExec = "Select Count(*) From compra where idtipofactura=5";
			iCuenta = Funcion.iEscalarSQL(cdsCliente, stExec, true) - iCuenta;
			Cursor =Cursors.Default;
			MessageBox.Show("Proceso realizado exitosamente, se han creado " + iCuenta.ToString() + " Facturas con sus consumos ", " Factura");

		}

	}
}
