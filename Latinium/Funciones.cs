using System;
using System.IO;
using System.Security.Cryptography;
using System.Text;
using System.Globalization;
using System.Data.SqlClient;
using System.Windows.Forms;
using System.Data;
using System.Data.OleDb;
using CrystalDecisions.Shared;
using System.Xml;
using Infragistics.Win.UltraWinGrid;
using System.Text.RegularExpressions;
using System.Drawing;

using System.Collections;
using System.ComponentModel;
using DevExpress.XtraEditors;
using System.Net;
using System.Web;
using System.Xml.Schema;
using System.Diagnostics;
using System.Threading;



namespace Latinium
{
	class Funcion
	{
		#region Variables
		public enum TipoFactura 
		{
			Venta = 1, OrdenCompra = 2, Cotizacion = 3, Compra = 4 , DevolucionVenta = 5, DevolucionCompra = 6, 
			OrdenProd = 7, EgresoBodega = 8, IngresoBodega = 9, ConsumoReal = 10, ReservaCliente = 11, ProdTerminado = 12,
			PedidoCliente = 13, PedidoProveedor = 14, IngresoBodega1 = 15, EgresoBodega1 = 16, IngresoProceso = 17,
			EgresoProceso = 18, NotaVenta = 19, EgresoSemiterminado = 20, CuentasxCobrar = 21, CuentasxPagar = 22,
			Importacion = 23, AnticipoCliente = 24, AnticipoProveedor = 25, ReservaProveedor = 26, DebitoVenta = 27,
			DebitoCompra = 28, PedidoEmpleado = 29, GuiaRemision = 30, ImpTransito = 31, Liquida = 32, CodigoCliente = 33,
			CodigoProveedor = 34, CodigoArticulo = 35, Requisición = 36, NIIF = 37,Aprobacion=38,OrdenTrabajo=39,Conclusion=40,
			OrdenTrabajoServicio=41
		}

		public static string NombreSevidor = "";
		public static bool bAyudaRemota = false;
		public static bool Seguridad = false;
		public static bool[,] Menu = new bool[10,20];
		public static bool[] MenuP = new bool[10];
		public static string ClaveMaestra =  "G7C0iPCTBp6uX9NAxQLyHeo/KTA=";
		static private StreamWriter writer;
		#endregion Variables
		#region EjecutaSQL
		public static void EjecutaSQL(C1.Data.C1DataSet ds, string stComando, bool bError)
		{
			DateTime dtIni = DateTime.Now;
			ds.Schema.Connections[0].Open();
			C1.Data.SchemaObjects.Connection miConeccion = ds.Schema.Connections[0];
			miConeccion = ds.Schema.Connections[0];

			SqlCommand cmdBusca = new SqlCommand (stComando,	
				(SqlConnection) miConeccion.DbConnection);

			try
			{
				cmdBusca.CommandTimeout = 240;
				cmdBusca.ExecuteNonQuery();
			}
			catch(System.Exception ex)
			{
				DateTime dtFin = DateTime.Now;
				TimeSpan ts = dtFin.Subtract(dtIni);
				string stRetorno = "";
				if (ts.Minutes > 0)
					stRetorno += ts.Minutes + " m ";
				if (ts.Seconds > 0)
					stRetorno += ts.Seconds + " s";

				if (bError)
					MessageBox.Show(MensajeError(ex) + stComando, "Ejecuta SQL " + stRetorno);
			}
			ds.Schema.Connections[0].Close();
			return;
		}

		public static SqlDataReader rInstanciaSQL(string stConeccion, string stComando, bool bError)
		{
			SqlDataReader rSalida = null;
			SqlConnection connection = new SqlConnection(stConeccion);
			SqlCommand cmdBusca = new SqlCommand (stComando, connection);
			connection.Open();

			try
			{
				cmdBusca.CommandTimeout = 120;
				rSalida = cmdBusca.ExecuteReader();
			}
			catch (System.Exception ex)
			{
				if (bError)
				{
					MessageBox.Show(MensajeError(ex) + stComando, "rInstanciaSQL");
					connection.Close();
					return rSalida;
				}
				else return rSalida;
			}
			return rSalida;
		}

		public static SqlDataReader rEscalarSQL(C1.Data.C1DataSet ds, string stComando, bool bError)
		{
			ds.Schema.Connections[0].Open();
			C1.Data.SchemaObjects.Connection miConeccion = ds.Schema.Connections[0];
			miConeccion = ds.Schema.Connections[0];

			SqlCommand cmdBusca = new SqlCommand (stComando,	
				(SqlConnection) miConeccion.DbConnection);

			SqlDataReader rSalida = null;
			try
			{
				cmdBusca.CommandTimeout = 0;
				rSalida = cmdBusca.ExecuteReader();
			}
			catch (System.Exception ex)
			{
				if (bError)
				{
					MessageBox.Show(MensajeError(ex) + stComando, "rEscalarSQL");
					ds.Schema.Connections[0].Close();
					return rSalida;
				}
				else return rSalida;
			}
			//			ds.Schema.Connections[0].Close();
			return rSalida;
		}

		public static string sEscalarSQL(C1.Data.C1DataSet ds, string stComando)
		{
			ds.Schema.Connections[0].Open();
			C1.Data.SchemaObjects.Connection miConeccion = ds.Schema.Connections[0];
			miConeccion = ds.Schema.Connections[0];

			SqlCommand cmdBusca = new SqlCommand (stComando,	
				(SqlConnection) miConeccion.DbConnection);

			string stSalida = "";
			try
			{
				cmdBusca.CommandTimeout = 120;
				stSalida = (string) cmdBusca.ExecuteScalar();
			}
			catch (System.Exception ex)
			{
				MessageBox.Show(MensajeError(ex) + stComando, "sEscalarSQL");
				ds.Schema.Connections[0].Close();
				return "";
			}
			ds.Schema.Connections[0].Close();
			return stSalida;
		}

		public static string sEscalarSQL(C1.Data.C1DataSet ds, string stComando, bool bError)
		{
			ds.Schema.Connections[0].Open();
			C1.Data.SchemaObjects.Connection miConeccion = ds.Schema.Connections[0];
			miConeccion = ds.Schema.Connections[0];

			SqlCommand cmdBusca = new SqlCommand (stComando,	
				(SqlConnection) miConeccion.DbConnection);

			string stSalida = "";
			try
			{
				cmdBusca.CommandTimeout = 120;
				stSalida = (string) cmdBusca.ExecuteScalar();
			}
			catch (System.Exception ex)
			{
				if (bError)
				{
					MessageBox.Show(MensajeError(ex) + stComando, "sEscalarSQL");
					ds.Schema.Connections[0].Close();
				}
				return "";
			}
			ds.Schema.Connections[0].Close();
			return stSalida;
		}

		public static decimal decEscalarSQL(C1.Data.C1DataSet ds, string stComando)
		{
			ds.Schema.Connections[0].Open();
			C1.Data.SchemaObjects.Connection miConeccion = ds.Schema.Connections[0];
			miConeccion = ds.Schema.Connections[0];

			SqlCommand cmdBusca = new SqlCommand (stComando,	
				(SqlConnection) miConeccion.DbConnection);

			decimal dSalida = 0;
			try
			{
				cmdBusca.CommandTimeout = 120;
				dSalida = decimal.Parse(cmdBusca.ExecuteScalar().ToString());
			}
			catch (System.Exception ex)
			{
				MessageBox.Show(MensajeError(ex) + stComando, "dEscalarSQL");
				ds.Schema.Connections[0].Close();
				return -1;
			}
			ds.Schema.Connections[0].Close();
			return dSalida;
		}

		public static double dEscalarSQL(C1.Data.C1DataSet ds, string stComando)
		{
			ds.Schema.Connections[0].Open();
			C1.Data.SchemaObjects.Connection miConeccion = ds.Schema.Connections[0];
			miConeccion = ds.Schema.Connections[0];

			SqlCommand cmdBusca = new SqlCommand (stComando,	
				(SqlConnection) miConeccion.DbConnection);

			double dSalida = 0;
			try
			{
				cmdBusca.CommandTimeout = 120;
				dSalida = double.Parse(cmdBusca.ExecuteScalar().ToString());
			}
			catch (System.Exception ex)
			{
				MessageBox.Show(MensajeError(ex) + stComando, "dEscalarSQL");
				ds.Schema.Connections[0].Close();
				return -1;
			}
			ds.Schema.Connections[0].Close();
			return dSalida;
		}

		public static DateTime dtEscalarSQL(C1.Data.C1DataSet ds, string stComando)
		{
			ds.Schema.Connections[0].Open();
			C1.Data.SchemaObjects.Connection miConeccion = ds.Schema.Connections[0];
			miConeccion = ds.Schema.Connections[0];

			SqlCommand cmdBusca = new SqlCommand (stComando,	
				(SqlConnection) miConeccion.DbConnection);

			DateTime dtSalida = DateTime.Today;
			try
			{
				cmdBusca.CommandTimeout = 120;
				dtSalida = DateTime.Parse(cmdBusca.ExecuteScalar().ToString());
			}
			catch (System.Exception ex)
			{
				MessageBox.Show(MensajeError(ex) + stComando, "dEscalarSQL");
				ds.Schema.Connections[0].Close();
				return DateTime.Today.AddYears(-100);
			}
			ds.Schema.Connections[0].Close();
			return dtSalida;
		}

		public static double diEscalarSQL(C1.Data.C1DataSet ds, string stComando)
		{
			ds.Schema.Connections[0].Open();
			C1.Data.SchemaObjects.Connection miConeccion = ds.Schema.Connections[0];
			miConeccion = ds.Schema.Connections[0];

			SqlCommand cmdBusca = new SqlCommand (stComando,	
				(SqlConnection) miConeccion.DbConnection);

			double dSalida = 0;
			try
			{
				cmdBusca.CommandTimeout = 120;
				dSalida = double.Parse(cmdBusca.ExecuteScalar().ToString());
			}
			catch (System.Exception ex)
			{
				MessageBox.Show(MensajeError(ex) + stComando, "dEscalarSQL");
				ds.Schema.Connections[0].Close();
				return -1;
			}
			ds.Schema.Connections[0].Close();
			return dSalida;
		}

		public static double dEscalarSQL(C1.Data.C1DataSet ds, string stComando, bool bError)
		{
			ds.Schema.Connections[0].Open();
			C1.Data.SchemaObjects.Connection miConeccion = ds.Schema.Connections[0];
			miConeccion = ds.Schema.Connections[0];

			SqlCommand cmdBusca = new SqlCommand (stComando,	
				(SqlConnection) miConeccion.DbConnection);

			double dSalida = 0;
			try
			{
				cmdBusca.CommandTimeout = 120;
				dSalida = (double) cmdBusca.ExecuteScalar();
			}
			catch (System.Exception ex)
			{
				if (bError)
				{
					MessageBox.Show(MensajeError(ex) + stComando, "dEscalarSQL");
					ds.Schema.Connections[0].Close();
					return -1;
				}
				else return 0;
			}
			ds.Schema.Connections[0].Close();
			return dSalida;
		}

		public static string VerificaSQL(C1.Data.C1DataSet ds, string stComando)
		{
			string bOk = "";
			ds.Schema.Connections[0].Open();
			C1.Data.SchemaObjects.Connection miConeccion = ds.Schema.Connections[0];
			miConeccion = ds.Schema.Connections[0];

			SqlCommand cmdBusca = new SqlCommand (stComando,	
				(SqlConnection) miConeccion.DbConnection);

			try
			{
				cmdBusca.CommandTimeout = 120;
				cmdBusca.ExecuteNonQuery();
			}
			catch(System.Exception ex)
			{
				bOk = ex.Message;
			}
			ds.Schema.Connections[0].Close();
			return bOk;
		}

		static public DataView dvProcedimiento(C1.Data.C1DataSet ds, string stProcedimiento)
		{
			ds.Schema.Connections[0].Open();
			C1.Data.SchemaObjects.Connection miConeccion = ds.Schema.Connections[0];
			miConeccion = ds.Schema.Connections[0];
			DataSet dsArticulo = new DataSet();
			SqlDataAdapter myData = new SqlDataAdapter(stProcedimiento, miConeccion.ConnectionString);
			myData.SelectCommand.CommandTimeout = 120;
			if(MenuLatinium.stIdDB == 29) myData.SelectCommand.CommandTimeout = 240;
			if(MenuLatinium.stIdDB == 34) myData.SelectCommand.CommandTimeout = 240;
			myData.TableMappings.Add("Tabla", "ExcelSube");

			dsArticulo.Reset();
			dsArticulo.Clear();
			try
			{
				myData.Fill(dsArticulo);
			}
			catch (Exception ex)
			{
				MessageBox.Show(MensajeError(ex), "Error al ejecutar procedimiento: ");
				return null;
			}
			return (DataView)dsArticulo.Tables[0].DefaultView;
		}

		#region Google documentos

		public static void ConfigurarConexionRed()
		{
			// Configurar la conexión de red
			Process.Start("cmd.exe", "/C net use \\\\192.168.1.10 /user:latinium 10dgpv15.p");
		}

		public static string ObtenerCarpeta(string ruc, string numero, decimal numIdCompra, string carpetaBase)
		{
			string carpeta = carpetaBase + "\\" + MenuLatinium.stDirFacturacion + "\\" +
				ruc.Replace(" ", "_") + "_" + numero + "_" + numIdCompra;
			return carpeta;
		}


		public static string GenerarNombreArchivoNuevo(string numero, decimal numIdCompra, string extension)
		{
			Random random = new Random();
			int randomNumber = random.Next(1, 10001);
			return numero + "_" + numIdCompra + "_" + randomNumber + extension;
		}


		public static void MoverArchivo(string origen, string destino)
		{
			FileInfo fi = new FileInfo(origen);
			string nuevoArchivo = Path.GetDirectoryName(destino) + "\\" + Path.GetFileName(destino);

			if (File.Exists(nuevoArchivo))
			{
				File.Delete(nuevoArchivo); // Eliminar si ya existe
			}

			File.Move(origen, nuevoArchivo);
		}


		public static void ValidarYSubirArchivo(string carpetaDestino, string archivoOrigen, string nombreArchivo)
		{
			if (File.Exists(carpetaDestino))
			{
				MessageBox.Show(string.Format("El archivo '{0}' ya existe. ", nombreArchivo), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				return;
			}

			File.Copy(archivoOrigen, carpetaDestino);
		}

		public static void SubirArchivo(string ruc, string numero, decimal numIdCompra, bool esNuevo, string carpetaBase, ref string sUrlNube)
		{
			string carpeta = ObtenerCarpeta(ruc, numero, numIdCompra, carpetaBase);
			string txtURL = string.Empty;

			try
			{
				ConfigurarConexionRed();

				if (!Directory.Exists(carpeta))
				{
					Directory.CreateDirectory(carpeta);
				}

				OpenFileDialog buscar = new OpenFileDialog();
				buscar.Filter = "Documentos e Imágenes|*.txt;*.pdf;*.doc;*.docx;*.xls;*.xlsx;*.csv;*.jpg;*.jpeg;*.png;*.gif;*.bmp"; // Filtros para documentos e imágenes

				if (buscar.ShowDialog() == DialogResult.OK)
				{
					string txtDirectorioarchivo = buscar.FileName;
					string extension = Path.GetExtension(txtDirectorioarchivo); 
					string nuevoNombreArchivo = esNuevo ? GenerarNombreArchivoNuevo(numero, numIdCompra, extension) : Path.GetFileName(txtDirectorioarchivo);
					string directorioDestino = carpeta + "\\" + nuevoNombreArchivo; // Concatenación manual

					if (esNuevo)
					{
						// Mover el archivo y subir
						MoverArchivo(txtDirectorioarchivo, directorioDestino);
						txtURL = nuevoNombreArchivo; // usar el nuevo nombre
						sUrlNube = directorioDestino.Replace(" ", "_"); // Asignar la URL a la variable global
						MessageBox.Show("Archivo subido exitosamente.", "Confirmación", MessageBoxButtons.OK, MessageBoxIcon.Information);
					}
					else
					{
						// Subir el archivo existente
						ValidarYSubirArchivo(directorioDestino, txtDirectorioarchivo, nuevoNombreArchivo);
						sUrlNube = directorioDestino.Replace(" ", "_"); // Asignar la URL a la variable global
						MessageBox.Show("Archivo actualizado exitosamente.", "Confirmación", MessageBoxButtons.OK, MessageBoxIcon.Information);
					}
				}
			}
			catch (Exception ex)
			{
				MessageBox.Show("Error: " + ex.Message);
			}
		}

		#endregion Googledocumentos

		public static string MensajeError(System.Exception ex)
		{
			string stMensaje = ex.Message;
			if (ex.InnerException != null) stMensaje = ex.InnerException.Message;
			return stMensaje;
		}
		public static bool EjecutaSQL(C1.Data.C1DataSet ds, string stComando)
		{
			C1.Data.SchemaObjects.Connection miConeccion = ds.Schema.Connections[0];
			miConeccion.Open();
			SqlCommand cmdBusca = new SqlCommand (stComando,	
				(SqlConnection) miConeccion.DbConnection);

			try
			{
				cmdBusca.CommandTimeout = 120;
				cmdBusca.ExecuteNonQuery();
			}
			catch (System.Exception ex)
			{
				MessageBox.Show(Funcion.MensajeError(ex), "Error en EjecutaSql");
				miConeccion.Close();
				//				ds.Schema.Connections[0].Close();
				return false;
			}
			//			ds.Schema.Connections[0].Close();
			miConeccion.Close();
			return true;
		}

		public static bool bEjecutaSQL(C1.Data.C1DataSet ds, string stComando)
		{
			ds.Schema.Connections[0].Open();
			C1.Data.SchemaObjects.Connection miConeccion = ds.Schema.Connections[0];
			miConeccion = ds.Schema.Connections[0];
			//	miConeccion.Open();
			SqlCommand cmdBusca = new SqlCommand (stComando,	
				(SqlConnection) miConeccion.DbConnection);

			bool bSalida = false;

			try
			{
				cmdBusca.CommandTimeout = 120;
				bSalida = (bool) cmdBusca.ExecuteScalar();
			}
			catch 
			{
				ds.Schema.Connections[0].Close();
				return false;
			}
			ds.Schema.Connections[0].Close();
			return bSalida;
		}

		public static int iEscalarSQL(C1.Data.C1DataSet ds, string stComando)
		{
			ds.Schema.Connections[0].Open();
			C1.Data.SchemaObjects.Connection miConeccion = ds.Schema.Connections[0];
			miConeccion = ds.Schema.Connections[0];

			SqlCommand cmdBusca = new SqlCommand (stComando, 
				(SqlConnection) miConeccion.DbConnection);

			int iSalida = 0;
			try
			{
				cmdBusca.CommandTimeout = 120;
				iSalida = (int) cmdBusca.ExecuteScalar();
			}
			catch (System.Exception ex)
			{
				MessageBox.Show(MensajeError(ex) + stComando, "iEscalarSQL");
				ds.Schema.Connections[0].Close();
				return -1;
			}
			ds.Schema.Connections[0].Close();
			return iSalida;
		}

		public static int iEscalarSQL(C1.Data.C1DataSet ds, string stComando, bool bError)
		{
			ds.Schema.Connections[0].Open();
			C1.Data.SchemaObjects.Connection miConeccion = ds.Schema.Connections[0];
			miConeccion = ds.Schema.Connections[0];

			SqlCommand cmdBusca = new SqlCommand (stComando,	
				(SqlConnection) miConeccion.DbConnection);

			int iSalida = 0;
			try
			{
				iSalida = (int) cmdBusca.ExecuteScalar();
			}
			catch (System.Exception ex)
			{
				if (bError)
				{
					MessageBox.Show(MensajeError(ex) + stComando, "iEscalarSQL");
					ds.Schema.Connections[0].Close();
					return -1;
				}
				else return 0;
			}
			ds.Schema.Connections[0].Close();
			return iSalida;
		}
		
		public static bool bEscalarSQL(C1.Data.C1DataSet ds, string stComando, bool bError)
		{
			ds.Schema.Connections[0].Open();
			C1.Data.SchemaObjects.Connection miConeccion = ds.Schema.Connections[0];
			miConeccion = ds.Schema.Connections[0];

			SqlCommand cmdBusca = new SqlCommand (stComando,	
				(SqlConnection) miConeccion.DbConnection);

			bool bSalida = false;
			try
			{
				cmdBusca.CommandTimeout = 120;
				bSalida = (bool) cmdBusca.ExecuteScalar();
			}
			catch (System.Exception ex)
			{
				if (bError)
				{
					MessageBox.Show(MensajeError(ex) + stComando, "bEscalarSQL");
					ds.Schema.Connections[0].Close();
					return false;
				}
				else return false;
			}
			ds.Schema.Connections[0].Close();
			return bSalida;
		}

		#endregion
		#region Formularios Remotos
		static private void CreaFGrilla(UltraGrid Grilla, string stPadre)
		{
			writer.WriteLine("{0}~{1}~{2}~{3}~{4}~{5}~{6}{7}",
				Grilla.Top, Grilla.Left, Grilla.Width, Grilla.Height, Grilla.Name, Grilla.Text, Grilla.GetType().ToString(), stPadre); 
			string stLista = "";

			string stArchivo = String.Format("C:\\Latinium\\Control{0}.xml", Grilla.Name);
			XmlTextWriter myXmlTextWriter = new XmlTextWriter(stArchivo, null);
			myXmlTextWriter.Formatting = Formatting.Indented;
			myXmlTextWriter.WriteStartDocument(false);

			myXmlTextWriter.WriteStartElement("Grilla");
			if (Grilla.Rows.Count == 0) // Llena una fila vacia
			{
				myXmlTextWriter.WriteStartElement("Fila");
				foreach (UltraGridColumn col in Grilla.DisplayLayout.Bands[0].Columns)
				{
					if (!col.Hidden)
					{
						stLista += "~";
						string stCol = col.Header.Caption;
						string stColValido = "";
						for (int i=0; i< stCol.Length; i++)
							if (EsAlfaNumerico(stCol.Substring(i, 1) ))
								stColValido += stCol.Substring(i, 1);
						myXmlTextWriter.WriteElementString(stColValido, "");
					}
				}
			}
			foreach(UltraGridRow dr in Grilla.Rows.All)
			{
				myXmlTextWriter.WriteStartElement("Fila");
				foreach (UltraGridColumn col in Grilla.DisplayLayout.Bands[0].Columns)
				{
					if (!col.Hidden)
					{
						string stFila = dr.Cells[col].Text.ToString();
						stLista += "~" + stFila;
						string stCol = col.Header.Caption;
						string stColValido = "";
						for (int i=0; i< stCol.Length; i++)
							if (EsAlfaNumerico(stCol.Substring(i, 1) ))
								stColValido += stCol.Substring(i, 1);

						myXmlTextWriter.WriteElementString(stColValido, stFila);
					}
				}
				myXmlTextWriter.WriteEndElement();
			}
			myXmlTextWriter.WriteEndElement();
			myXmlTextWriter.Flush();
			myXmlTextWriter.Close();
		}

		static private void CreaFTabControl(Infragistics.Win.UltraWinTabControl.UltraTabControl pn)
		{
			writer.WriteLine("{0}~{1}~{2}~{3}~{4}~{5}~{6}~{7}",
				pn.Top, pn.Left, pn.Width, pn.Height, pn.Name, pn.Text, 
				pn.GetType().ToString(), pn.Tabs.Count); 
			foreach (Control cn3 in pn.Controls)
			{
				if (cn3.GetType().ToString().EndsWith("UltraTabPageControl"))
					CreaFTabPage((Infragistics.Win.UltraWinTabControl.UltraTabPageControl)cn3, cn3.Name);
				else
					writer.WriteLine("{0}~{1}~{2}~{3}~{4}~{5}~{6}",
						cn3.Top + pn.Top, cn3.Left + pn.Left, cn3.Width, cn3.Height, 
						cn3.Name, cn3.Text, cn3.GetType().ToString());
			}
		}		
		
		static private void CreaFTabPage(Infragistics.Win.UltraWinTabControl.UltraTabPageControl pn, string stPagina)
		{
			writer.WriteLine("{0}~{1}~{2}~{3}~{4}~{5}~{6}",
				pn.Top, pn.Left, pn.Width, pn.Height, pn.Name, pn.Tab.Text, pn.GetType().ToString()); 
			foreach (Control cn3 in pn.Controls)
			{
				if (cn3.GetType().ToString().EndsWith("UltraTabPageControl"))
					CreaFTabPage((Infragistics.Win.UltraWinTabControl.UltraTabPageControl)cn3, cn3.Name);
				else if (cn3.GetType().ToString().EndsWith("UltraGrid"))
					CreaFGrilla((UltraGrid)cn3, "~" + cn3.Name);
				else
					writer.WriteLine("{0}~{1}~{2}~{3}~{4}~{5}~{6}~{7}",
						cn3.Top + pn.Top, cn3.Left + pn.Left, cn3.Width, cn3.Height, 
						cn3.Name, cn3.Text, cn3.GetType().ToString(), stPagina);
			}
		}

		static public void CreaFormulario(Form miForma)
		{
			writer = File.CreateText("C:\\Latinium\\Controles.txt");
			writer.WriteLine("{0}~{1}~{2}", miForma.Width, miForma.Height, miForma.Text);

			foreach (Control cn in miForma.Controls)
			{
				string stLista = "";
				if (cn.GetType().ToString().EndsWith("UltraGrid"))
					CreaFGrilla((UltraGrid)cn, "");
					//				else if (cn.GetType().ToString().EndsWith("LayoutControl"))
					//					CreaFLayout((DevExpress.XtraLayout.LayoutControl)cn);
				else if (cn.GetType().ToString().EndsWith("UltraTabControl"))
					CreaFTabControl((Infragistics.Win.UltraWinTabControl.UltraTabControl)cn);
				else
					if (cn.Visible)
					writer.WriteLine("{0}~{1}~{2}~{3}~{4}~{5}~{6}{7}",
						cn.Top, cn.Left, cn.Width, cn.Height, cn.Name, cn.Text, cn.GetType().ToString(), stLista);
			}
			writer.Close();
		}

		#endregion Formularios Remotos
		#region Generales
		static private bool EsAlfaNumerico(String str)
		{
			Regex regexAlphaNum=new Regex("[^a-zA-Z0-9]");
			return !regexAlphaNum.IsMatch(str);    
		}

		static public string Digito(string stNumero)
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
		static public string CreaHash(string stCadena)
		{
			UnicodeEncoding uEncode = new UnicodeEncoding();
			byte[] ByteCadena = uEncode.GetBytes(stCadena);

			SHA1CryptoServiceProvider sha = new SHA1CryptoServiceProvider(); 
			byte[] result = sha.ComputeHash(ByteCadena);
			return Convert.ToBase64String(result);
		}
		
		static public string CreaHash(string stCadena, int idGrupo)
		{
			stCadena += idGrupo.ToString() + "Bl2$";
			UnicodeEncoding uEncode = new UnicodeEncoding();
			byte[] ByteCadena = uEncode.GetBytes(stCadena);

			SHA1CryptoServiceProvider sha = new SHA1CryptoServiceProvider(); 
			byte[] result = sha.ComputeHash(ByteCadena);
			return Convert.ToBase64String(result);
		}

		static public DateTime CreaFecha(string stFecha)
		{
			int iDia = int.Parse(stFecha.Substring(0, 2));
			int iMes = int.Parse(stFecha.Substring(3, 2));
			int iAnio = int.Parse(stFecha.Substring(6, 4));
			DateTime dtFecha = new DateTime(iAnio, iMes, iDia);
			return dtFecha;
		}

		static public double DifFecha360(DateTime dtFechaFin, DateTime dtFechaIni)
		{
			int iAnio1, iAnio2, iMes1, iMes2, iDia1, iDia2, lMes;
			int lAnios;
			if(dtFechaFin.AddDays(1).Day == 1) dtFechaFin = dtFechaFin.AddDays(1);
			iAnio1 = dtFechaIni.Year;
			iMes1 = dtFechaIni.Month;
			iDia1 = dtFechaIni.Day;

			iAnio2 = dtFechaFin.Year;
			iMes2 = dtFechaFin.Month;
			iDia2 = dtFechaFin.Day;
			// Excepcion del visiesto
			DateTime dtFechaIni2 = DateTime.Today;
			if (iMes2 == 12)
				dtFechaIni2= new System.DateTime(iAnio2 + 1, 1, 1);
			else
				dtFechaIni2= new System.DateTime(iAnio2, iMes2+1, 1);

			dtFechaIni2 = dtFechaIni2.AddDays(-1);
			int iUltimoDia = dtFechaIni2.Day;
			if (iDia1 > iUltimoDia) iDia1 = iUltimoDia;

			dtFechaIni2= new System.DateTime(iAnio2, iMes2, iDia1);
			TimeSpan iNumDias = dtFechaFin.Subtract(dtFechaIni2);
			lMes = iMes2 - iMes1;
			lAnios = iAnio2 - iAnio1;
			lAnios = lAnios * 360 + lMes * 30 + iNumDias.Days;
			return lAnios;
		}

		static public string VerificaCedula(string stRuc, bool bMensaje)
		{
			stRuc = stRuc.Trim();
			string stError = "";
			Regex regex = new Regex(@"^[0-9]+$");
			if (!regex.IsMatch(stRuc))
			{
				stError = "Cédula debe tener solo números";
				return stError;
			}
			if (stRuc.Length < 10)
			{
				stError = "Cédula no puede tener menos de 10 dígitos";
				return stError;
			}

			if (!(stRuc.Length == 10 || stRuc.Length == 13))
			{
				stError = "Cédula o Ruc no puede tener " + stRuc.Length.ToString().Trim() + " caracteres";
				return stError;
			}

			string stTercero = stRuc.Substring(2,1);
			//identificacion de Sociedad, Cedula, o Institucion Publica
			if (stTercero == "7" || stTercero == "8")
			{
				stError = "Error en tercer digito";
				return stError;
			}
			if (stRuc.StartsWith("0000"))
			{
				stError = "No puede empezar con 0000";
				return stError;
			}

			string stSerie;
			int iCar;
			int iMult;
			int iSuma = 0;
			switch(int.Parse(stTercero))
			{
				case 6: // Institucion pública
					stSerie = "32765432";
					for(int i=0; i<8; i++)
					{
						iCar = int.Parse(stRuc.Substring(i,1));
						iMult = int.Parse(stSerie.Substring(i,1));
						iSuma += iCar* iMult;
					}
          
					iSuma += -(int) Math.Floor(iSuma/ 11) * 11;
					if (iSuma == 0 && stRuc.Substring(8 ,1) == "0")
					{
						if (bMensaje)	stError = "RUC de Institución publica Correcto";
						return stError;
					}

					if (iSuma == 11 - int.Parse(stRuc.Substring(8 ,1)))
					{
						if (bMensaje)	stError = "RUC de Institución publica Correcto";
						return stError;
					}
					else
					{
						stError = "RUC de Institución publica InCorrecto";
						return stError;
					}

					//					break;
				case 9: // Caso de RUC de sociedades
					stSerie = "432765432";
					for(int i=0; i < 9; i++)
					{
						iCar = int.Parse(stRuc.Substring(i,1));
						iMult = int.Parse(stSerie.Substring(i,1));
						iSuma += iCar*iMult;
					}

					iSuma += - (int) Math.Floor(iSuma/ 11) * 11;
					if (iSuma == 0 && int.Parse(stRuc.Substring(9, 1)) == 0)
					{
						if (bMensaje)	stError = "RUC Correcto";
						return stError;
					}

					if (iSuma == 11 - int.Parse(stRuc.Substring(9,1)))
					{
						if (bMensaje)	stError = "RUC Correcto";
						return stError;
					}
					else
					{
						stError = "RUC ERRONEO";
						return stError;
					}
					//					break;
				default: // Cédula
	
					for (int i=0; i< 9; i++)
					{
						iCar = int.Parse(stRuc.Substring(i,1));
						if (i%2 == 0)
							iCar = iCar * 2;
						else
							iCar = iCar;

						if (iCar > 9)
							iCar -= 9;

						iSuma += iCar;
					}

					iSuma = iSuma% 10;
					if( iSuma == 0) iSuma = 10;

					if (iSuma == 10- int.Parse(stRuc.Substring(9,1)))
					{
						if (stRuc.Length == 10)
						{
							if (bMensaje) stError = ""; //"Cédula Correcta";
						}
						else
							if (bMensaje) stError = "RUC Correcto";
						return stError;
					}
					else
					{
						if (stRuc.Length == 10)
						{
							stError = "Cédula ERRONEA";
						}
						else
							stError = "RUC ERRONEO";
						return stError;
					}
			} // Fin de switch
		}

		#endregion Generales
		#region facturacion
		static public void AuditoriaImprime(int IdTipo, int IdUnica, bool bEditar)
		{
			if (MenuLatinium.bAuditoria == false)
			{
				MessageBox.Show("No tiene acceso a Auditoria", "Información",
					MessageBoxButtons.OK, MessageBoxIcon.Information);
				return;
			}
			// Definicion de Parametros
			ParameterFields paramFields = new ParameterFields ();
			ParameterField paramField1 = new ParameterField ();
			ParameterField paramField2 = new ParameterField ();

			ParameterDiscreteValue discreteVal = new ParameterDiscreteValue ();
			paramField1.ParameterFieldName = "@idTipo";
			discreteVal.Value = IdTipo;
			paramField1.CurrentValues.Add (discreteVal);
			paramFields.Add (paramField1);

			ParameterDiscreteValue discreteVal2 = new ParameterDiscreteValue ();
			paramField2.ParameterFieldName = "@IdUnica";
			discreteVal2.Value = IdUnica;
			paramField2.CurrentValues.Add (discreteVal2);
			paramFields.Add (paramField2);

			// Fin de definicion de parametros
			string stFiltro = "";
			if (bEditar) stFiltro = "{AuditoriaGenera;1.idAuditaEstado} = 2";

			Reporte miRep = new Reporte("AuditoriaPantalla.Rpt", stFiltro);
			//			miRep.MdiParent = MdiParent;
			miRep.crVista.ParameterFieldInfo = paramFields;
			miRep.ShowDialog();
		}

		static public bool EsProveedor(int IdTipoFactura)
		{
			bool bProveedor = true;
			switch (IdTipoFactura)
			{
				case 1: // Venta
				case 3: // Cotizacion
				case 5: // Dev. Venta
				case 8: // Egreso de Bodega
				case 11: // Reservacion Cliente
				case 12: // Producto Terminado
				case 13: // Pedido Cliente
				case 24: // Anticipo Cliente
				case 27: // Nota de Debito en Venta
				case 43: // Pre Factura
				case 99: // Nota de Debito en Venta
					bProveedor = false;
					break; 
			}
			return bProveedor;
		}

		static public int IngresoEgreso(int IdTipoFactura)
		{
			int iIngEgr = 2;
			if (IdTipoFactura == 1 || IdTipoFactura == 11 || IdTipoFactura == 27 || IdTipoFactura == 5) // Ingreso
				iIngEgr = 1;
			return iIngEgr;
		}

		public static string NumeraFactura(C1.Data.C1DataSet cdsCompra, int IdTipoFactura)
		{
			return Funcion.sEscalarSQL(cdsCompra, "EXEC FacturaNumeroEstacion " + IdTipoFactura);
		}

		#endregion facturacion

		static public Decimal Redondeo(decimal dValor, int Digitos)
		{
			decimal dRound = decimal.Round(dValor, Digitos);

			return dRound;
		}

		static public DateTime FechaCompra(C1.Data.C1DataSet ds, int iTipo)
		{
			ds.Schema.Connections[0].Open();
			C1.Data.SchemaObjects.Connection miConeccion = ds.Schema.Connections[0];

			string stOrden = "";
			if (iTipo == 1) stOrden = "DESC";

			DateTime dtFecha = (DateTime) DateTime.Now;
			try
			{
				string stSelect = "SELECT TOP 1 Fecha FROM Compra " + 
					" ORDER BY Fecha " + stOrden;
				SqlCommand cmdGrupo = new SqlCommand (
					stSelect, (SqlConnection) miConeccion.DbConnection);
				dtFecha = (DateTime) cmdGrupo.ExecuteScalar();
			}
			catch(System.Exception ex)
			{
				Console.WriteLine(ex.Message + ". En Función FechaCompra");
			}
			ds.Schema.Connections[0].Close();
			return dtFecha;
		}

		public static void CreaDetalle(C1.Data.C1DataSet cdsRetencion, int iCuenta, double dDebe, Double dHaber, int iCuentaP)
		{
			cdsRetencion.Schema.Connections[0].Open();
			C1.Data.SchemaObjects.Connection miConeccion = cdsRetencion.Schema.Connections[0];
			CultureInfo us = new CultureInfo("en-US");

			string stSelect = "";
			try
			{
				stSelect = "INSERT INTO AsientoDetalle (idCuenta, idAsiento, Debe, Haber)" +
					" VALUES (" + iCuenta.ToString() + ", " + iCuentaP.ToString() 
					+ ", " + dDebe.ToString("0.00", us) + ", " + dHaber.ToString("0.00", us) + ")";
				SqlCommand cmdGrupo = new SqlCommand (
					stSelect, (SqlConnection) miConeccion.DbConnection);
				cmdGrupo.Transaction = (SqlTransaction)
					miConeccion.DbTransaction;
				cmdGrupo.ExecuteNonQuery();
			}
			catch(System.Exception ex)
			{
				MessageBox.Show(ex.Message + " " + stSelect, "CreaDetalle");
			}
			cdsRetencion.Schema.Connections[0].Close();
		}

		public static void RetencionCuenta(C1.Data.C1DataSet cdsRetencion, int IdFactura, int IdTipo, int IdCuentaP)
		{
			cdsRetencion.Schema.Connections[0].Open();
			C1.Data.SchemaObjects.Connection miConeccion = cdsRetencion.Schema.Connections[0];

			double dRentaRet = 0;
			double dIvaRet = 0;
			double dRentaRet1 = 0;
			double dIvaRet1 = 0;
			int iIdRentaRet = 0;
			int iIdIvaRet = 0;
			int iIdRentaRet1 = 0;
			int iIdIvaRet1 = 0;
			int iIdRetencion = 0;

			string stSelect = "Select FormMarzo From Seteo ";
			bool bFormMarzo = false;
			bFormMarzo = Funcion.bEscalarSQL(cdsRetencion, stSelect, false);
			try
			{
				stSelect = "SELECT RentaRetenida, IvaRetenido, " +
					"RentaRetenida1, IvaRetenido1, idRetRenta, idRetIva, " +
					"idRetRenta1, idRetIva1, idRetencion " +
					"FROM Retencion WHERE idCompra = " +
					IdFactura.ToString().Trim();
				SqlCommand cmdGrupo = new SqlCommand (
					stSelect, (SqlConnection) miConeccion.DbConnection);

				SqlDataReader miReader;
				miReader = cmdGrupo.ExecuteReader();
				miReader.Read();
				dRentaRet =  miReader.GetDouble(0);
				dIvaRet =  miReader.GetDouble(1);
				dRentaRet1 =  miReader.GetDouble(2);
				dIvaRet1 =  miReader.GetDouble(3);
				iIdRentaRet =  miReader.GetInt32(4);
				iIdIvaRet =  miReader.GetInt32(5);
				iIdRentaRet1 =  miReader.GetInt32(6);
				iIdIvaRet1 =  miReader.GetInt32(7);
				iIdRetencion =  miReader.GetInt32(8);
				miReader.Close();
			}
			catch(System.Exception ex)
			{
				Console.WriteLine(ex.Message, "Retencion Cuenta");
			}
			cdsRetencion.Schema.Connections[0].Close();

			int iCuenta = 0;
			#region Retenciones Renta
			if (bFormMarzo)
			{
				stSelect = "Exec AsientoRetencionC " + IdCuentaP.ToString()
					+ ", " + IdFactura.ToString() + ", " + IdTipo.ToString();
				Funcion.EjecutaSQL(cdsRetencion, stSelect, true);
			}
			else
			{
				if (dRentaRet > 0) 
				{
					if (IdTipo == 1)
					{
						iCuenta = CuentaRetencion(cdsRetencion, iIdRentaRet,3);
						CreaDetalle(cdsRetencion, iCuenta, dRentaRet, 0, IdCuentaP);
					}
					else
					{
						iCuenta = CuentaRetencion(cdsRetencion, iIdRentaRet,4);
						CreaDetalle(cdsRetencion, iCuenta, 0, dRentaRet, IdCuentaP);
					}
				}
				if (dRentaRet1 > 0) 
				{
					if (IdTipo == 1)
					{
						iCuenta = CuentaRetencion(cdsRetencion, iIdRentaRet1, 3);
						CreaDetalle(cdsRetencion, iCuenta, dRentaRet1, 0, IdCuentaP);
					}
					else
					{
						iCuenta = CuentaRetencion(cdsRetencion, iIdRentaRet1, 4);
						CreaDetalle(cdsRetencion, iCuenta, 0, dRentaRet1, IdCuentaP);
					}
				}
			}
			#endregion Retenciones

			if (bFormMarzo) return;
			if (dIvaRet > 0) 
			{
				if (IdTipo == 1)
				{
					iCuenta = CuentaRetencion(cdsRetencion, iIdIvaRet, 1);
					CreaDetalle(cdsRetencion, iCuenta, dIvaRet, 0, IdCuentaP);
				}
				else
				{
					iCuenta = CuentaRetencion(cdsRetencion, iIdIvaRet, 2);
					CreaDetalle(cdsRetencion, iCuenta, 0, dIvaRet, IdCuentaP);
				}
			}
			if (dIvaRet1 > 0) 
			{
				if (IdTipo == 1)
				{
					iCuenta = CuentaRetencion(cdsRetencion, iIdIvaRet1, 1);
					CreaDetalle(cdsRetencion, iCuenta, dIvaRet1, 0, IdCuentaP);
				}
				else
				{
					iCuenta = CuentaRetencion(cdsRetencion, iIdIvaRet1, 2);
					CreaDetalle(cdsRetencion, iCuenta, 0, dIvaRet1, IdCuentaP);
				}
			}

		}

		static public string RolCampo(C1.Data.C1DataSet ds, string stConcepto, int iRolGrupo)
		{
			ds.Schema.Connections[0].Open();
			C1.Data.SchemaObjects.Connection miConeccion = ds.Schema.Connections[0];

			int iPos = 0;
			int iTipo = 0;
			stConcepto = stConcepto.Trim();
			string stCampo = "";
			try
			{
				string stSelect = "Select Posicion from RolConcepto WHERE LTRIM(Concepto) = '" 
					+ stConcepto +"' And idRolGrupo = "
					+ iRolGrupo.ToString();
				SqlCommand cmdGrupo = new SqlCommand (
					stSelect, (SqlConnection) miConeccion.DbConnection);
				iPos = (int) cmdGrupo.ExecuteScalar();

				stSelect = "Select idRolTipo from RolConcepto WHERE LTRIM(Concepto) = '" 
					+ stConcepto +"' And idRolGrupo = "
					+ iRolGrupo.ToString();
				cmdGrupo = new SqlCommand (
					stSelect, (SqlConnection) miConeccion.DbConnection);
				iTipo = (int) cmdGrupo.ExecuteScalar();

				switch( iTipo)
				{
					case 1:
						stCampo = "Ingreso";
						break;
					case 2:
						stCampo = "Egreso";
						break;
					case 3:
						stCampo = "Provision";
						break;
					case 4:
						stCampo = "Valor";
						break;
				}
				stCampo += iPos.ToString().Trim();
			}
			catch(System.Exception ex)
			{
				stCampo = stConcepto;
				Console.WriteLine(ex.Message + ". En Función RolCampo");
			}
			ds.Schema.Connections[0].Close();
			return stCampo;
		}

		public static bool Permiso(string stCodigo, C1.Data.C1DataSet ds)
		{
			string nivel = "n_" + MenuLatinium.iNivel.ToString().Trim();

			string stSelect = "SELECT " + nivel + " FROM Seguridad WHERE id_seg = " + stCodigo;
			bool bAcceso = Funcion.bEjecutaSQL(ds, stSelect);
			return bAcceso;
		}

		public static int CuentaRetencion(C1.Data.C1DataSet cdsRetencion, int IdRet, int IdTipo)
		{
			string stTipoCuenta = "idCuentaI";
			string stIdPrinc = "idRetRenta";
			string stTipoRet = "RetencionRenta";
			if (IdTipo < 3) 
			{
				stTipoRet = "RetencionIva";
				stIdPrinc = "idRetIva";
			}
			if (IdTipo == 2 || IdTipo == 4) stTipoCuenta = "idCuentaE";
			cdsRetencion.Schema.Connections[0].Open();
			C1.Data.SchemaObjects.Connection miConeccion = cdsRetencion.Schema.Connections[0];

			int IdCuenta = 0;
			try
			{
				string stSelect = "SELECT " + stTipoCuenta + " FROM " + stTipoRet  + 
					" WHERE " + stIdPrinc + " = " +
					IdRet.ToString().Trim();
				SqlCommand cmdGrupo = new SqlCommand (
					stSelect, (SqlConnection) miConeccion.DbConnection);
				IdCuenta = (int) cmdGrupo.ExecuteScalar();
			}
			catch(System.Exception ex)
			{
				IdCuenta = 0;
				Console.WriteLine(ex.Message + ". En Función CuentaRetencion");
			}
			cdsRetencion.Schema.Connections[0].Close();
			return IdCuenta;
		}
			
		public static int CuentaGrupoCliente(C1.Data.C1DataSet cdsCompra, int IdCliente)
		{
			cdsCompra.Schema.Connections[0].Open();
			C1.Data.SchemaObjects.Connection miConeccion = cdsCompra.Schema.Connections[0];

			string stSelect = "SELECT idGrupoCliente FROM Cliente WHERE idCliente = " +
				IdCliente.ToString().Trim();
			SqlCommand cmdGrupo = new SqlCommand (
				stSelect, (SqlConnection) miConeccion.DbConnection);
			int IdGrupo = 0;
			try
			{
				IdGrupo = (int) cmdGrupo.ExecuteScalar();
			}
			catch{}

			stSelect = "SELECT idCuenta FROM ClienteGrupo WHERE idGrupoCliente = " +
				IdGrupo.ToString().Trim();
			SqlCommand cmdCuenta = new SqlCommand (
				stSelect, (SqlConnection) miConeccion.DbConnection);
			int IdCuenta = 0;
			try
			{
				IdCuenta = (int) cmdCuenta.ExecuteScalar();
			}
			catch{}

			cdsCompra.Schema.Connections[0].Close();
			return IdCuenta;
		}

		static private DateTime dtInicio = DateTime.Now;
		static public void TiempoInicia()
		{
			dtInicio = DateTime.Now;
		}
		static public string Tiempo(string stAccion)
		{
			TimeSpan ts = DateTime.Now.Subtract(dtInicio);
			return String.Format("{0} {1:n0} SEGUNDOS", stAccion, ts.TotalSeconds);
		}

		#region Menus
		public static void SetMenu(int i, int j, bool Valor)
		{
			Menu[i,j] = Valor;
		}
		public static bool GetMenu(int i, int j)
		{
			return Menu[i,j];
		}
		public static bool GetMenu(int i)
		{
			return MenuP[i];
		}
		public static void SetMenu(int i, bool Valor)
		{
			MenuP[i] = Valor;
		}
		public string AccesoEmpresaBase()	
		{
			string strBase = "EmpresasBase";
			string strServidor = MenuLatinium.stDirServidor;
			string strUsuario = MenuLatinium.stUsuario;

			string stRetorno = @"Password=Bl45o6$9;"
				+ "Persist Security Info=True;User ID=" + strUsuario.Trim()
				+ ";Initial Catalog=" + strBase.Trim()
				+ ";Data Source=" + strServidor.Trim();

			return stRetorno;
		}

		public string AccesoBase(string BaseDato, string idTipo)	
		{
			if (idTipo == "3")
			{
				string strBase = BaseDato;
				string strServidor = MenuLatinium.stDirServidor;
				string strUsuario = MenuLatinium.stUsuario;

				string stRetorno = @"Provider=SQLOLEDB.1;Password=Bl45o6$9;"
					+ "Persist Security Info=True;User ID=" + strUsuario.Trim()
					+ ";Initial Catalog=" + strBase.Trim()
					+ ";Data Source=" + strServidor.Trim()
					+ ";Connect Timeout=120";

				return stRetorno;
			}
			return "";
		}

		#endregion Menus

		#region Acceso
		public string AccesoBase(string BaseDato)
		{
			int idTipo = MenuLatinium.iTipoDB;

			if (idTipo == 1)
			{
				string strBase = MenuLatinium.stDirNumero;
				string strServidor = MenuLatinium.stDirServidor;
				return @"Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + 
					strBase + ";Persist Security Info=False";
			}

			if (idTipo == 2)
			{
				string strBase = MenuLatinium.stDirNumero;
				string strServidor = MenuLatinium.stDirServidor;

				return	@"Integrated Security=SSPI;Persist "
					+ "Security Info=False;" 
					+ "Pooling=False; Connection Timeout=1;"
					+	"Initial Catalog=" + strBase.Trim() 
					+ ";Data Source=" + strServidor.Trim();
			}

			if (idTipo == 3)
			{
				string strBase = BaseDato;
				string strServidor = MenuLatinium.stDirServidor;
				string strUsuario = MenuLatinium.stUsuario;

				string stRetorno = @"Password=Bl45o6$9;"
					+ "Persist Security Info=True;User ID=" + strUsuario.Trim()
					+ ";Initial Catalog=" + strBase.Trim()
					+ ";Data Source=" + strServidor.Trim();

				return stRetorno;
			}

			if (idTipo == 4)
			{
				return @"Provider=OraOLEDB.Oracle.1;Password=sa;Persist Security Info=True;"
					+"User ID=sa;Data Source=oracle";
			}
			return "";
		}

		public string AccesoBaseWin(string BaseDato)
		{
			string strBase = BaseDato;
			string strServidor = MenuLatinium.stDirServidor;
			string strUsuario = MenuLatinium.stUsuario;
			string stRetorno = @"Integrated Security=SSPI;"
				+ "Persist Security Info=False;User ID=" + strUsuario.Trim()
				+ ";Initial Catalog=" + strBase.Trim()
				+ ";Data Source=" + strServidor.Trim();

			return stRetorno;
		}

		public string AccesoBase(string BaseDato, int idTipo)
		{
			return @"Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + 
				BaseDato + ";Persist Security Info=False";
		}

		public string AccesoEmpresa()
		{
			string st = Application.ProductName.ToString().Trim();
			string st1 = Application.ExecutablePath.ToString().Trim();
			string strBase = st1.Substring(0,st1.Length - st.Length - 4);
			//			this.cdsAsiento.Schema.Connections[0].ConnectionString = 
			return strBase + "Empresa.mdb";
		}

		public static void Mensajes(string strMensaje, int Tipo)
		{
			switch(Tipo)
			{
				case 1:
					MessageBox.Show(strMensaje);
					break;
				case 2:
//					CronosHMenu.MensajePanel(strMensaje);
					break;
				case 3:
					MessageBox.Show(strMensaje);
					break;
				default:
					MessageBox.Show(strMensaje,"Error en tipo");
					break;
			}
		}
		#endregion Acceso

		#region Validacion
		#region Campos Numericos
		static public void CamposNumericos(object sender, KeyPressEventArgs e)
		{
			if (Char.IsDigit(e.KeyChar))
			{
				e.Handled = false;
			}
			else if (Char.IsControl(e.KeyChar))
			{
				e.Handled = false;
			}
			else if (Char.IsSeparator(e.KeyChar))
			{
				e.Handled = false;
			}
			else
			{
				e.Handled = true;
			}
		}
		
		static public void CamposNumericos_1(object sender, KeyPressEventArgs e)
		{
			if (Char.IsDigit(e.KeyChar))
			{
				e.Handled = false;
			}
			else if (Char.IsControl(e.KeyChar))
			{
				e.Handled = false;
			}
			else if (Char.IsSeparator(e.KeyChar))
			{
				e.Handled = false;
			}
			else if (e.KeyChar == 45)
			{
				e.Handled = false;
			}
			else
			{
				e.Handled = true;
			}
		}

			static public void CamposNumericosPasaporte(object sender, KeyPressEventArgs e)
			{
				if (Char.IsDigit(e.KeyChar))
				{
					e.Handled = false;
				}
				else if (Char.IsControl(e.KeyChar))
				{
					e.Handled = false;
				}
//				else if (Char.IsSeparator(e.KeyChar))
//				{
//					e.Handled = false;
//				}
				else if (Char.IsLetter(e.KeyChar))
				{
					e.Handled = false;
				}
				else
				{
					e.Handled = true;
				}
			}

		static public void CaracteresEspeciales(object sender, KeyPressEventArgs e)
		{
			if (!char.IsLetter(e.KeyChar) && !char.IsWhiteSpace(e.KeyChar) && e.KeyChar !='\b')
			{
				e.Handled = true;
			}
		}
		#endregion Validacion
		#endregion Campos Numericos
		
		static public void Resolucion(System.Windows.Forms.Form Form, int Alto, int Ancho)
		{
			int iAlto = Screen.PrimaryScreen.Bounds.Height;
			int iAncho = Screen.PrimaryScreen.Bounds.Width;

			if (iAlto < Alto && iAncho < Ancho) Form.Dock = DockStyle.Fill;
		}

		public static DateTime RetornaFecha()
		{
			using (SqlConnection oConexion = new SqlConnection("Password=Bl45o6$9;Persist Security Info=True;User ID= " + MenuLatinium.stUsuario + ";Initial Catalog=" + MenuLatinium.stDirFacturacion + ";Data Source=" + MenuLatinium.stDirServidor))
			{
				oConexion.Open();				
				SqlCommand oCmdActualiza = oConexion.CreateCommand();
				oCmdActualiza.Connection = oConexion;
				oCmdActualiza.CommandType = CommandType.Text;
				oCmdActualiza.CommandTimeout = 0;				
				oCmdActualiza.CommandType = CommandType.StoredProcedure;
				oCmdActualiza.CommandText = "RetornaFechaServidor";
				DateTime dtFecha = (DateTime)oCmdActualiza.ExecuteScalar(); 
				oConexion.Close();
				return dtFecha;
			}			
		}

		public static bool ValidaPeriodo(DateTime dtFecha, C1.Data.C1DataSet ds, string sModulo)
		{
			string sSQL = string.Format("Exec ValidaPeriodoActivo '{0}', '{1}'", dtFecha.ToString("yyyyMMdd"), sModulo);	
			bool bActivo = Funcion.bEjecutaSQL(ds, sSQL);
			return bActivo;
		}

		static public SqlConnection Coneccion()
		{
//			string Conexion = new SqlConnection(@"Password=Bl45o6$9;Persist Security Info=True;User ID= " + MenuLatinium.stUsuario + ";Initial Catalog=" + MenuLatinium.stDirFacturacion + ";Data Source=" + MenuLatinium.stDirServidor);
			SqlConnection oConexion = new SqlConnection("Password=Bl45o6$9;Persist Security Info=True;User ID= " + MenuLatinium.stUsuario + ";Initial Catalog=" + MenuLatinium.stDirFacturacion + ";Data Source=" + MenuLatinium.stDirServidor);

//			string Conexion = @"Password=Bl45o6$9;Persist Security Info=True;User ID= " + MenuLatinium.stUsuario + ";Initial Catalog=" + MenuLatinium.stDirFacturacion + ";Data Source=" + MenuLatinium.stDirServidor;
			return oConexion;
		}

		public static bool ValidaBloqueo(int idOrigen, int Origen, C1.Data.C1DataSet ds)
		{
			if (Funcion.iEscalarSQL(ds, string.Format("Select COUNT(*) From BloqueoTransacciones Where idOrigen = {0} And Origen = {1} And Bloqueado = 1", idOrigen, Origen)) > 0)
			{
				string sUsuario = "";
				string sEstacion = "";
				DateTime dtFecha = DateTime.Today;

				string sSQL = string.Format("Select Usuario, Estacion, Fecha From BloqueoTransacciones Where idOrigen = {0} And Origen = {1} And Bloqueado = 1", idOrigen, Origen);
				SqlDataReader drBloqueo = Funcion.rEscalarSQL(ds, sSQL, true);
				drBloqueo.Read();
				if (drBloqueo.HasRows)
				{
					sUsuario = drBloqueo.GetValue(0).ToString();
					sEstacion = drBloqueo.GetValue(1).ToString();
					dtFecha = drBloqueo.GetDateTime(2);
				}
				drBloqueo.Close();

				MessageBox.Show(string.Format("Este Registro esta siendo Editado : \n\n Usuario : '{0}' \n Estacion : '{1}' \n Desde : '{2}'", sUsuario, sEstacion, dtFecha), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
					
				return true;
			}
			else return false;
		}

		public static bool ValidaBloqueoERMEnTransacciones(int idCompra, C1.Data.C1DataSet ds)
		{			
			if (Funcion.iEscalarSQL(ds, string.Format("Exec ValidaBloqueoERMEnTransacciones {0}", idCompra)) > 0)
			{
				MessageBox.Show("Este Registro esta en proceso de carga de seriales.\n\nNo puede editarlo", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
					
				return true;
			}
			else return false;
		}		

		#region Ejecucion configuracion requerimientos
		//Configuraciones requerimiento
		public static DataSet dsEjecucion(string sconexion, string scomando)
		{
			DataSet oDSPoint = new DataSet();
			try
			{
				SqlConnection oConexionPoint = new SqlConnection(sconexion);
				SqlCommand oCmdPoint = new SqlCommand();
				oCmdPoint.Connection = oConexionPoint;
				oCmdPoint.CommandType = CommandType.Text;
				//Cargar Tipo Requerimiento
				oCmdPoint.CommandText = scomando;
				oConexionPoint.Open();
				SqlDataAdapter oDAPoint = new SqlDataAdapter(oCmdPoint.CommandText, oConexionPoint);
				oDAPoint.Fill(oDSPoint);
				oConexionPoint.Close();
			}
			catch(Exception ex)
			{
				Console.WriteLine(ex.Message);
			}
			return oDSPoint;
		}

		public static bool bEjecucion(string sconexion, string scomando, bool bError)
		{
			bool resultado = false;
			try
			{
				SqlConnection oConexionPoint = new SqlConnection(sconexion);
				SqlCommand oCmdPoint = new SqlCommand();
				oCmdPoint.Connection = oConexionPoint;
				oCmdPoint.CommandType = CommandType.Text;
				//Cargar Tipo Requerimiento
				oCmdPoint.CommandText = scomando;
				oConexionPoint.Open();
				SqlCommand oCmdEmp = new SqlCommand(scomando, oConexionPoint);
				resultado = (bool)oCmdEmp.ExecuteScalar();
				oConexionPoint.Close();
			}
			catch(Exception ex)
			{
				Console.WriteLine(ex.Message);
			}
			return resultado;
		}

		public static int iEjecucion(string sconexion, string scomando, bool bError)
		{
			int resultado = -1;
			try
			{
				SqlConnection oConexionPoint = new SqlConnection(sconexion);
				SqlCommand oCmdPoint = new SqlCommand();
				oCmdPoint.Connection = oConexionPoint;
				oCmdPoint.CommandType = CommandType.Text;
				//Cargar Tipo Requerimiento
				oCmdPoint.CommandText = scomando;
				oConexionPoint.Open();
				SqlCommand oCmdEmp = new SqlCommand(scomando, oConexionPoint);
				resultado = (int)oCmdEmp.ExecuteScalar();
				oConexionPoint.Close();
			}
			catch(Exception ex)
			{
				Console.WriteLine(ex.Message);
			}
			return resultado;
		}

		public static string sEjecucion(string sconexion, string scomando)
		{
			string resultado = "";
			try
			{
				SqlConnection oConexionPoint = new SqlConnection(sconexion);
				SqlCommand oCmdPoint = new SqlCommand();
				oCmdPoint.Connection = oConexionPoint;
				oCmdPoint.CommandType = CommandType.Text;
				//Cargar Tipo Requerimiento
				oCmdPoint.CommandText = scomando;
				oConexionPoint.Open();
				SqlCommand oCmdEmp = new SqlCommand(scomando, oConexionPoint);
				resultado = (string)oCmdEmp.ExecuteScalar();
				oConexionPoint.Close();
			}
			catch(Exception ex)
			{
				Console.WriteLine(ex.Message);
			}
			return resultado;
		}

		public static DateTime dtEjecucion(string sconexion, string scomando)
		{
			DateTime resultado = DateTime.Now;
			try
			{
				SqlConnection oConexionPoint = new SqlConnection(sconexion);
				SqlCommand oCmdPoint = new SqlCommand();
				oCmdPoint.Connection = oConexionPoint;
				oCmdPoint.CommandType = CommandType.Text;
				//Cargar Tipo Requerimiento
				oCmdPoint.CommandText = scomando;
				oConexionPoint.Open();
				SqlCommand oCmdEmp = new SqlCommand(scomando, oConexionPoint);
				resultado = DateTime.Parse(oCmdEmp.ExecuteScalar().ToString());
				oConexionPoint.Close();
			}
			catch(Exception ex)
			{
				Console.WriteLine(ex.Message);
			}
			return resultado;
		}

		public static DateTime dtEjecucion(string sconexion, string scomando, bool berror)
		{
			DateTime resultado = DateTime.Now;
			DateTime fechaDefault = new DateTime();
			try
			{
				SqlConnection oConexionPoint = new SqlConnection(sconexion);
				SqlCommand oCmdPoint = new SqlCommand();
				oCmdPoint.Connection = oConexionPoint;
				oCmdPoint.CommandType = CommandType.Text;
				//Cargar Tipo Requerimiento
				oCmdPoint.CommandText = scomando;
				oConexionPoint.Open();
				SqlCommand oCmdEmp = new SqlCommand(scomando, oConexionPoint);
				resultado = DateTime.Parse(oCmdEmp.ExecuteScalar().ToString());
				oConexionPoint.Close();
			}
			catch(Exception ex)
			{
				Console.WriteLine(ex.Message);
				if(berror == true)
				{
					return fechaDefault;
				}
			}
			return resultado;
		}

		//Ejecutar sql por conexión requerimiento
		public static void EjecusionEjecutaSQL(string sconexion, string stComando, bool bError)
		{
			DateTime dtIni = DateTime.Now;
			SqlConnection oConexionPoint = new SqlConnection(sconexion);
			SqlCommand oCmdPoint = new SqlCommand();
			oCmdPoint.CommandType = CommandType.Text;
			oCmdPoint.CommandText = stComando;
			oConexionPoint.Open();
			SqlCommand oCmdEmp = new SqlCommand(stComando, oConexionPoint);
			try
			{
				oCmdEmp.CommandTimeout = 120;
				oCmdEmp.ExecuteNonQuery();
			}
			catch(System.Exception ex)
			{
				DateTime dtFin = DateTime.Now;
				TimeSpan ts = dtFin.Subtract(dtIni);
				string stRetorno = "";
				if (ts.Minutes > 0)
					stRetorno += ts.Minutes + " m ";
				if (ts.Seconds > 0)
					stRetorno += ts.Seconds + " s";

				if (bError)
					MessageBox.Show(MensajeError(ex) + stComando, "Ejecuta SQL " + stRetorno);
			}
			oConexionPoint.Close();
			return;
		}
		#endregion Ejecucion configuracion requerimientos
		//Ejecución de log
		public static string slogtipoerror = "ERROR";
		public static string slogtipoadvertencia = "ADVERTENCIA";
		public static string slogtipoexito = "ÉXITO";
		public static string slogtipoinformacion = "INFORMACIÓN";
		public static string smensajeerror = "Los registros no se ejecutaron correctamente, intentelo de nuevo."; 

		public static void LogSistema(C1.Data.C1DataSet ds, string Pantalla, string Mensaje, string Procedimiento, string Tipo)
		{
			string query = "";
			try
			{
				//Escribirlo en un log del .10
				WriteLog(Mensaje, Procedimiento , Pantalla);
			}
			catch(Exception ex)
			{
				Procedimiento = Procedimiento.Replace("'", "");
				Mensaje = Mensaje.Replace("'", " ");
				string sms = ex.Message;
				//Escribir el log en la tabla LogSistema
				query = String.Format("EXEC LogSistemaCrear '{0}','{1}','{2}','{3}'",
					Pantalla, @Mensaje, @Procedimiento, Tipo);
				EjecutaSQL(ds, query, false);
			}
		}

		//Log pero solo a point
		public static void LogSistema(string sconexion, string Pantalla, string Mensaje, string Procedimiento, string Tipo)
		{
			string query = "";
			try
			{
				//Escribirlo en un log del .10
				WriteLog(Mensaje, Procedimiento, Pantalla);
			}
			catch(Exception ex)
			{
				Procedimiento = Procedimiento.Replace("'", "");
				Mensaje = Mensaje.Replace("'", " ");
				string sms = ex.Message;
				//Escribir el log en la tabla LogSistema
				query = String.Format("EXEC LogSistemaCrear '{0}','{1}','{2}','{3}'",
					Pantalla, @Mensaje, @Procedimiento, Tipo);
				EjecusionEjecutaSQL(sconexion, query, false);
			}
		}

		public static void ValidaLetraDigito(object sender, System.Windows.Forms.KeyPressEventArgs e, char getExcepcion)
		{
			if (Char.IsControl(e.KeyChar) || e.KeyChar == getExcepcion)
			{
				e.Handled = false;
			}
			else if (Char.IsLetterOrDigit(e.KeyChar))
			{
				e.Handled = false;
			}			
			else
			{
				e.Handled = true;
			}
		}

		//Esto debe ir en base la ubicación
		//Producción
		//static public string  LogPath = @"\\192.168.1.10\Proyectos\LogLatinium\Desarrollo\";
		//Desarrollo


		public static void tratarFicheroLog()
		{
			try
			{
				if(MenuLatinium.LogPath != "" && MenuLatinium.LogFileName != "" && MenuLatinium.stUsuario != "")
				{
					string sLogComprueba1MesAnt = MenuLatinium.LogPath + MenuLatinium.LogFileName + MenuLatinium.stUsuario + DateTime.Today.AddMonths(-1).ToString("yyyyMM") + ".txt";
					string sLogComprueba2MesAnt = MenuLatinium.LogPath + MenuLatinium.LogFileName + MenuLatinium.stUsuario + DateTime.Today.AddMonths(-2).ToString("yyyyMM") + ".txt";

					if (File.Exists(sLogComprueba2MesAnt))
						File.Delete(sLogComprueba2MesAnt);

					if (!File.Exists(sLogComprueba1MesAnt))
						File.Move(MenuLatinium.LogPath + MenuLatinium.LogFileName + MenuLatinium.stUsuario + ".txt", sLogComprueba1MesAnt);
				}
			}
			catch (Exception ex)
			{
				string sms = ex.Message;
			}
		}
		public static void WriteLog(string message, string Procedimiento,string Pantalla)
		{
			if(MenuLatinium.LogPath != "" && MenuLatinium.LogFileName != "" && MenuLatinium.stUsuario != "")
			{
				string path = MenuLatinium.LogPath + MenuLatinium.LogFileName + MenuLatinium.stUsuario + ".txt";
				StreamWriter stream = null;
				try
				{
					stream = File.AppendText(path);
					stream.WriteLine(string.Format("{0} - {1} - {2} - {3} - {4} - {5}.", DateTime.Now, Pantalla,MenuLatinium.stUsuario, MenuLatinium.stDirFacturacion, message, Procedimiento));
				}
				catch (Exception ex)
				{
					throw ex;
				}
				finally
				{
					if (stream != null)
						stream.Close();
				}
			}
		}

		static public DataView dvProcedimiento(string sconexion, string stProcedimiento)
		{

			DataSet dsArticulo = new DataSet();
			SqlDataAdapter myData = new SqlDataAdapter(stProcedimiento, sconexion);
			myData.SelectCommand.CommandTimeout = 120;
			myData.TableMappings.Add("Tabla", "ExcelSube");

			dsArticulo.Reset();
			dsArticulo.Clear();
			try
			{
				myData.Fill(dsArticulo);
			}
			catch (Exception ex)
			{
				MessageBox.Show(MensajeError(ex), "Error al ejecutar procedimiento: ");
				return null;
			}
			return (DataView)dsArticulo.Tables[0].DefaultView;
		}
		
		//Contador de filas para un ultradatagrid
		public static Infragistics.Win.UltraWinGrid.UltraGrid ContadorFilas(Infragistics.Win.UltraWinGrid.UltraGrid ugd, string Columna)
		{
			Infragistics.Win.UltraWinGrid.UltraGrid getugd = ugd;

			int Contador = 0;
			foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in ugd.Rows.All)	
			{
				Contador = Contador + 1;
				dr.Cells[Columna].Value = Contador;
			}
			return getugd;
		}

		public static void VersionarReportesCarpetasYArchivos(string rutaBaseLocal, string rutaBaseServidor, string Servidor, string Usuario, string Clave)
		{
			DateTime inicio = DateTime.Now;
			string mensajeInicio = "Inicio del proceso: " + inicio.ToString("yyyy-MM-dd HH:mm:ss");
			Console.WriteLine(mensajeInicio);
			//Funcion.LogSistema(frmRequerimiento.sconexionPoint, "VersionarReportesCarpetasYArchivos", mensajeInicio, "", Funcion.slogtipoinformacion);

			try
			{
				ConectarAlServidor(Servidor, Usuario, Clave);
				VerificarYActualizarArchivosYCarpetas(rutaBaseLocal, rutaBaseServidor);
			}
			catch (Exception ex)
			{
				Funcion.LogSistema(frmRequerimiento.sconexionPoint, "Error general", ex.Message, "", Funcion.slogtipoerror);
			}
			finally
			{
				DesconectarServidor(Servidor);

				DateTime fin = DateTime.Now;
				TimeSpan tiempoTotal = fin - inicio;

//				string mensajeFin = "Fin del proceso: " + fin.ToString("yyyy-MM-dd HH:mm:ss") +
//					" | Tiempo total: " + tiempoTotal.TotalSeconds.ToString("F2") + " segundos.";
//				Console.WriteLine(mensajeFin);
//				Funcion.LogSistema(frmRequerimiento.sconexionPoint, "VersionarReportesCarpetasYArchivos", mensajeFin, "", Funcion.slogtipoinformacion);
			}
		}
		
		private static void VerificarYActualizarArchivosYCarpetas(string rutaLocal, string rutaServidor)
		{
			string[] archivosExcluidos = new string[] { "thumbs.db", "desktop.ini", "~$tempfile.xlsx" };

			if (!Directory.Exists(rutaServidor))
			{
				Funcion.LogSistema(frmRequerimiento.sconexionPoint, "Ruta no encontrada", "Carpeta servidor no existe: " + rutaServidor, "", Funcion.slogtipoerror);
				return;
			}

			if (!Directory.Exists(rutaLocal))
			{
				Directory.CreateDirectory(rutaLocal);
			}

			string[] subdirectorios = Directory.GetDirectories(rutaServidor);
			for (int i = 0; i < subdirectorios.Length; i++)
			{
				string nombreSub = Path.GetFileName(subdirectorios[i]);
				string subLocal = Path.Combine(rutaLocal, nombreSub);
				VerificarYActualizarArchivosYCarpetas(subLocal, subdirectorios[i]);
			}

			string[] archivosServidor = Directory.GetFiles(rutaServidor);
			for (int i = 0; i < archivosServidor.Length; i++)
			{
				string archivoServidor = archivosServidor[i];
				string nombreArchivo = Path.GetFileName(archivoServidor);

				bool ignorar = false;
				for (int j = 0; j < archivosExcluidos.Length; j++)
				{
					if (string.Compare(nombreArchivo, archivosExcluidos[j], true) == 0)
					{
						ignorar = true;
						break;
					}
				}
				if (ignorar)
					continue;

				string archivoLocal = Path.Combine(rutaLocal, nombreArchivo);

				try
				{
					bool copiar = false;

					if (!File.Exists(archivoLocal))
					{
						copiar = true;
					}
					else
					{
						FileInfo infoLocal = new FileInfo(archivoLocal);
						FileInfo infoServidor = new FileInfo(archivoServidor);

						if (infoServidor.LastWriteTime > infoLocal.LastWriteTime || infoServidor.Length != infoLocal.Length)
						{
							copiar = true;
						}
					}

					if (copiar)
					{
						DateTime inicioCopia = DateTime.Now;

						File.Copy(archivoServidor, archivoLocal, true);

						DateTime finCopia = DateTime.Now;
						TimeSpan tiempoArchivo = finCopia - inicioCopia;

						string logCopia = "Copiado: " + nombreArchivo +
							" | Tiempo: " + tiempoArchivo.TotalSeconds.ToString("F2") + " segundos" +
							" | Origen: " + archivoServidor +
							" | Destino: " + archivoLocal;

						Console.WriteLine(logCopia);
						Funcion.LogSistema(frmRequerimiento.sconexionPoint, "Archivo Copiado", logCopia, "", Funcion.slogtipoinformacion);
					}
				}
				catch (Exception ex)
				{
					Funcion.LogSistema(frmRequerimiento.sconexionPoint, "Error al copiar archivo", ex.Message, nombreArchivo, Funcion.slogtipoerror);
				}
			}
		}

		private static void ConectarAlServidor(string servidor, string usuario, string clave)
		{
			try
			{
				ProcessStartInfo conectar = new ProcessStartInfo("cmd.exe", "/C net use " + servidor + " /user:" + usuario + " " + clave);
				conectar.CreateNoWindow = true;
				conectar.UseShellExecute = false;
				conectar.RedirectStandardOutput = true;
				conectar.RedirectStandardError = true;
				Process.Start(conectar);
			}
			catch (Exception ex)
			{
				Funcion.LogSistema(frmRequerimiento.sconexionPoint, "Error al conectar al servidor", ex.Message, "", Funcion.slogtipoerror);
			}
		}

		private static void DesconectarServidor(string servidor)
		{
			try
			{
				ProcessStartInfo desconectar = new ProcessStartInfo("cmd.exe", "/C net use " + servidor + " /delete /y");
				desconectar.CreateNoWindow = true;
				desconectar.UseShellExecute = false;
				desconectar.RedirectStandardOutput = true;
				desconectar.RedirectStandardError = true;
				Process.Start(desconectar);
			}
			catch (Exception ex)
			{
				Funcion.LogSistema(frmRequerimiento.sconexionPoint, "Error al desconectar del servidor", ex.Message, "", Funcion.slogtipoerror);
			}
		}

		public class ParametrosCopiaArchivos
		{
			private string rutaLocal;
			private string rutaServidor;
			private string servidorReportes;
			private string usuarioReportes;
			private string claveReportes;

			public string RutaLocal
			{
				get { return rutaLocal; }
				set { rutaLocal = value; }
			}

			public string RutaServidor
			{
				get { return rutaServidor; }
				set { rutaServidor = value; }
			}

			public string ServidorReportes
			{
				get { return servidorReportes; }
				set { servidorReportes = value; }
			}

			public string UsuarioReportes
			{
				get { return usuarioReportes; }
				set { usuarioReportes = value; }
			}

			public string ClaveReportes
			{
				get { return claveReportes; }
				set { claveReportes = value; }
			}

			// Constructor
			public ParametrosCopiaArchivos(string rutaLocal, string rutaServidor, string servidorReportes, string usuarioReportes, string claveReportes)
			{
				this.rutaLocal = rutaLocal;
				this.rutaServidor = rutaServidor;
				this.servidorReportes = servidorReportes;
				this.usuarioReportes = usuarioReportes;
				this.claveReportes = claveReportes;
			}
		}
		public static ParametrosCopiaArchivos ParametrosCopiaArchivosThread;  // Campo estático
		public static void VersionarReportesCarpetasYArchivosThread()
		{
			// Accedemos a los parámetros a través de la variable estática
			ParametrosCopiaArchivos parametros = ParametrosCopiaArchivosThread;

			// Aquí llamamos a tu función original con los parámetros
			VersionarReportesCarpetasYArchivos(parametros.RutaLocal, parametros.RutaServidor, parametros.ServidorReportes, parametros.UsuarioReportes, parametros.ClaveReportes);
		}

		public static void IniciarCopiaArchivos(string rutaBaseLocalReportes, string rutaBaseServidorReportes, string ServidorReportes, string UsuarioReportes, string ClaveReportes)
		{
			// Asignamos los parámetros al campo estático
			ParametrosCopiaArchivosThread = new ParametrosCopiaArchivos(rutaBaseLocalReportes, rutaBaseServidorReportes, ServidorReportes, UsuarioReportes, ClaveReportes);

			// Creamos y arrancamos el hilo
			Thread hiloCopia = new Thread(new ThreadStart(VersionarReportesCarpetasYArchivosThread));
			hiloCopia.Start();
		}

		public static string ObtenerCadenaConexionEmpresa(int idEmpresa, string conexionInicio, string plantillaConexion)
		{
			string baseDeDatos = "";
			string servidor = "";
			string hostName = System.Environment.MachineName;

			using (SqlConnection conexion = new SqlConnection(conexionInicio))
			{
				string query = string.Format("SELECT Empresa, BaseDeDatos, Servidor FROM Empresas WHERE idEmpresa = {0}", idEmpresa);
				SqlCommand command = new SqlCommand(query, conexion);
				conexion.Open();
				using (SqlDataReader reader = command.ExecuteReader())
				{
					if (reader.Read())
					{
						baseDeDatos = reader["BaseDeDatos"].ToString();
						servidor = @"192.168.1.165,56399";
					}
				}
			}

			string usuario = MenuLatinium.stUsuario;
			return string.Format(plantillaConexion, usuario, baseDeDatos, servidor);
		}
		

		public static void  ConsumirAPISilencioso(string urlBase)
		{
			try
			{


				HttpWebRequest request = (HttpWebRequest)WebRequest.Create(urlBase);
				request.Method = "GET";
				request.Timeout = 20000;

				using (HttpWebResponse response = (HttpWebResponse)request.GetResponse())
				{
					using (StreamReader reader = new StreamReader(response.GetResponseStream()))
					{
						reader.ReadToEnd();
					}
				}
			}
			catch
			{
			}
		}

	

		public static string ConsumirAPI(string urlBase)
		{
			string respuesta = "";

			try
			{
				// Construir la URL con el par?metro
				string urlCompleta = urlBase ;

				// Crear la solicitud
				HttpWebRequest request = (HttpWebRequest)WebRequest.Create(urlCompleta);
				request.Method = "GET";
				request.Timeout = 20000; // 10 segundos (ajustable)

				// Obtener la respuesta
				using (HttpWebResponse response = (HttpWebResponse)request.GetResponse())
				{
					if (response.StatusCode == HttpStatusCode.OK)
					{
						using (StreamReader reader = new StreamReader(response.GetResponseStream()))
						{
							respuesta = reader.ReadToEnd();
						}
					}
					else
					{
						MessageBox.Show("Error en el servicio: " + response.StatusCode.ToString(), "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
						return "Error";
					}
				}
			}
			catch (WebException ex)
			{
				MessageBox.Show("Error de conexi?n al servicio:\n" + ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
				return "Error";
			}
			catch (Exception ex)
			{
				MessageBox.Show("Error inesperado:\n" + ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
				return "Error";
			}

			return "OK";
		}



	}
}
