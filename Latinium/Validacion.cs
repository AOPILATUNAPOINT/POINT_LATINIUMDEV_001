using System;
using System.Net;
using System.Net.Sockets;
using System.Data;
using System.Data.SqlClient;
using System.Windows.Forms;
using Infragistics;
using Infragistics.Win.UltraWinGrid;
using System.Globalization;
using System.Data.OleDb;
using System.Data.Odbc;
using System.Text;
using System.Text.RegularExpressions;


namespace Latinium
{
	/// <summary>
	/// Descripción breve de Validacion.
	/// </summary>
	public class Validacion
	{
		public Validacion()
		{

		}	

		static public bool bNumeros(string sCadena)
		{
			String expresion;
			expresion = @"[0-9]";

			if (!Regex.IsMatch(sCadena, expresion)) return false;			
			//			if (Regex.Replace(sCadena, expresion, String.Empty).Length > 0) 
			return true;
		}

		static public bool bEMail(String email)
		{			
			String expresion;
			expresion = @"^([a-zA-Z0-9_\-\.]+)@((\[[0-9]{1,3}\.[0-9]{1,3}\.[0-9]{1,3}\.)|(([a-zA-Z0-9\-]+\.)+))([a-zA-Z]{2,4}|[0-9]{1,3})(\]?)$";
			if (Regex.IsMatch(email, expresion))
			{
				if (Regex.Replace(email, expresion, String.Empty).Length == 0)
				{
//					string[] host = (email.Split('@'));
//					string hostname = host[1];
//     
//					IPHostEntry IPhst = Dns.Resolve(hostname);
//					IPEndPoint endPt = new IPEndPoint(IPhst.AddressList[0], 25);
//					Socket s = new Socket(endPt.AddressFamily, SocketType.Stream,ProtocolType.Tcp);
//					s.Connect(endPt); 

					return true;
				}
				else
				{
					MessageBox.Show("E-Mail incorrecto", "Validacion", MessageBoxButtons.OK, MessageBoxIcon.Stop);
					return false;
				}
			}
			else
			{
				MessageBox.Show("E-Mail incorrecto", "Validacion", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				return false;
			}
		}

		static public bool vbValidaTelefono(Infragistics.Win.UltraWinEditors.UltraTextEditor uTXT, int iTipo, bool bObligatorio)
		{			
			if (uTXT.Text.ToString().Trim().Length == 0 && bObligatorio)
			{
				MessageBox.Show("Ingrese el Numero Telefonico", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);				
				uTXT.Focus();
				return false;
			}

			if (uTXT.Text.ToString().Trim().Length > 0 && uTXT.Text.ToString().Trim().Length < 9)
			{
				MessageBox.Show("Para convencional :\n\nEl Número de Teléfono debe tener minimo 9 digitos \n\nLos dos primeros digitos corresponden al Codigo Provincial\n\n\nPara Celular :\n\nDebe tener minimo 10 digitos y los dos primeros deben ser '09'", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Error);
				uTXT.Focus();
				return false;
			}

			#region Convencional
			if (iTipo == 1 && uTXT.Text.ToString().Trim().Length > 0)
			{
				if (uTXT.Text.ToString().Trim().Length < 9)
				{
					MessageBox.Show("Numero Convencional debe tener 9 digitos", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);									
					uTXT.Focus();
					return false;
				}

				#region Valida Codigo Provincial
				string sCodProv = uTXT.Text.ToString().Trim().Substring(0, 2);
				string[] sCodigos = { "02", "03", "04", "05", "06", "07" };
				int iIndex = 0;
				bool bCodProv = false;
				for (int i = 0; i < sCodigos.Length; i++)
				{
					iIndex = sCodigos[i].IndexOf(sCodProv);
					if (iIndex >= 0)
					{
						bCodProv = true;
						break;
					}
				}

				if (!bCodProv)
				{
					MessageBox.Show("El Codigo Provincial es Incorrecto", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Error);
					uTXT.Focus();
					return false;
				}
				#endregion Valida Codigo Provincial

				#region Valida Numero Inicial
				string sNumIniTelf = uTXT.Text.ToString().Trim().Substring(2, 1);
				string[] sInicial = { "2", "3", "4", "5", "6", "7" };
				iIndex = 0;
				bool bNumIni = false;
				for (int i = 0; i < sInicial.Length; i++)
				{
					iIndex = sInicial[i].IndexOf(sNumIniTelf);
					if (iIndex >= 0)
					{
						bNumIni = true;
						break;
					}
				}

				if (!bNumIni)
				{
					MessageBox.Show("El Numero Telefónico es Incorrecto", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Error);
					uTXT.Focus();
					return false;
				}
				#endregion Valida Numero Inicial

				#region Valida Cadena Repetida
				string[] sNumConv = { "0000000", "1111111", "2222222", "3333333", "4444444", "5555555", "6666666", "7777777", "8888888", "9999999" };	
				string sNumTelf = uTXT.Text.ToString().Trim().Substring(2, 7);
				iIndex = 0;
				bool bNumConv = false;
				for (int i = 0; i < sNumConv.Length; i++)
				{
					iIndex = sNumConv[i].IndexOf(sNumTelf);
					if (iIndex >= 0)
					{
						bNumConv = true;
						break;
					}
				}

				if (bNumConv)
				{
					MessageBox.Show("El Numero Telefónico es Incorrecto", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Error);
					uTXT.Focus();
					return false;
				}	
				#endregion Valida Cadena Repetida
			}
			#endregion Convencional

			#region Celular
			if (iTipo == 2 && uTXT.Text.ToString().Trim().Length > 0)
			{
				if (uTXT.Text.ToString().Trim().Length < 10)
				{
					MessageBox.Show("Numero Celular debe tener 10 digitos", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);														
					uTXT.Focus();
					return false;					
				}

				if (uTXT.Text.ToString().Trim().Substring(0, 2) != "09") 
				{
					MessageBox.Show("Numero Celular empieza con 09", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);														
					uTXT.Focus();
					return false;					
				}

				#region Valida Cadena Repetida
				string[] sNumCel = { "00000000", "11111111", "22222222", "33333333", "44444444", "55555555", "66666666", "77777777", "88888888", "99999999" };	
				string sNumCelular = uTXT.Text.ToString().Trim().Substring(2, 7);
				int iIndex = 0;
				bool bNumCelular = false;
				for (int i = 0; i < sNumCel.Length; i++)
				{
					iIndex = sNumCel[i].IndexOf(sNumCelular);
					if (iIndex >= 0)
					{
						bNumCelular = true;
						break;
					}
				}

				if (bNumCelular)
				{
					MessageBox.Show("El Número de Celular es Incorrecto", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Error);
					uTXT.Focus();
					return false;
				}	
				#endregion Valida Cadena Repetida
			}
			#endregion Celular

			return true;
		}

		static public void ModificaDescuentos(Infragistics.Win.UltraWinGrid.BeforeCellUpdateEventArgs e, bool bGuarda, bool bEmpleado)
		{		
			if (e.Cell.Row.Cells["Tipo"].Value == System.DBNull.Value) return;

			if ((int)e.Cell.Row.Cells["Tipo"].Value == 5)
			{
				MessageBox.Show("Garantia Extendida no aplica descuentos", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				e.Cancel = true;
				return;
			}

			if ((int)e.Cell.Row.Cells["Tipo"].Value == 8)
			{
				MessageBox.Show("D&FI no aplica descuentos", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				e.Cancel = true;
				return;
			}

			if (!bEmpleado)
			{
				if (!bGuarda)
				{
					if ((int)e.Cell.Row.Cells["Tipo"].Value >= 1 && !(bool)e.Cell.Row.Cells["Principal"].Value)
					{
						MessageBox.Show("Producto promocional no puede modificar descuentos", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
						e.Cancel = true;
						return;
					}
				}
			}
		}

		static public bool vbTexto(Infragistics.Win.UltraWinEditors.UltraTextEditor uTXT, int iMinLongitud, int iMaxLongitud, string sCampo)
		{
			if (uTXT.Text.ToString().Trim().Length == 0)
			{
				MessageBox.Show(string.Format("El campo '{0}' no puede estar vacio", sCampo), "Información", MessageBoxButtons.OK,MessageBoxIcon.Stop);
				uTXT.Focus();
				return false;
			}

			if (uTXT.Text.ToString().Length < iMinLongitud)
			{
				MessageBox.Show(string.Format("El campo '{0}' debe tener minimo {1} caracteres", sCampo, iMinLongitud), "Información", MessageBoxButtons.OK,MessageBoxIcon.Stop);
				uTXT.Focus();
				return false;
			}

			if (uTXT.Text.ToString().Length > iMaxLongitud)
			{
				MessageBox.Show(string.Format("El campo '{0}' debe tener maximo {1} caracteres", sCampo, iMaxLongitud), "Información", MessageBoxButtons.OK,MessageBoxIcon.Stop);
				uTXT.Focus();
				return false;
			}

			return true;
		}

		static public bool vbVariableTexto(string uTXT, int iMinLongitud, int iMaxLongitud, string sCampo)
		{
			if (uTXT.Length == 0)
			{
				MessageBox.Show(string.Format("Seleccione el archivo  '{0}' .", sCampo), "Información", MessageBoxButtons.OK,MessageBoxIcon.Stop);
			//	uTXT.Focus();
				return false;
			}

			if (uTXT.Length  < iMinLongitud)
			{
				MessageBox.Show(string.Format("Seleccione el archivo  '{0}' !!.", sCampo), "Información", MessageBoxButtons.OK,MessageBoxIcon.Stop);
			//	uTXT.Focus();
				return false;
			}

//			if (uTXT.Text.ToString().Length > iMaxLongitud)
//			{
//				MessageBox.Show(string.Format("El campo '{0}' debe tener maximo {1} caracteres", sCampo, iMaxLongitud), "Información", MessageBoxButtons.OK,MessageBoxIcon.Stop);
//				uTXT.Focus();
//				return false;
//			}

			return true;
		}

		static public bool vNotas(Infragistics.Win.UltraWinEditors.UltraTextEditor uTXT)
		{
			if (uTXT.Text.ToString().Length == 0)
			{
				MessageBox.Show("Ingrese una nota","Información", MessageBoxButtons.OK,MessageBoxIcon.Stop);
				uTXT.Focus();
				return false;
			}
			if (uTXT.Text.ToString().Length < 20)
			{
				MessageBox.Show("La Nota debe constar de 20 caracteres","Información", MessageBoxButtons.OK,MessageBoxIcon.Stop);
				uTXT.Focus();
				return false;
			}

			return true;
		}

		static public bool vbCampoEnteroVacio(Infragistics.Win.UltraWinEditors.UltraNumericEditor uNumEditor, string sMensaje, int iMinimo, int iMaximo)
		{
			if (uNumEditor.Value == System.DBNull.Value)
			{
				MessageBox.Show(sMensaje, "Validacion", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				uNumEditor.Focus();
				return false;
			}

			if ((int)uNumEditor.Value < iMinimo)
			{
				MessageBox.Show(string.Format("El valor minimo es {0}", iMinimo), "Validacion", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				uNumEditor.Focus();
				return false;
			}

			if ((int)uNumEditor.Value > iMaximo)
			{
				MessageBox.Show(string.Format("El valor maximo es {0}", iMaximo), "Validacion", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				uNumEditor.Focus();
				return false;
			}

			return true;
		}

		static public bool vbCampoDecimalVacio(Infragistics.Win.UltraWinEditors.UltraNumericEditor uNumEditor, string sMensaje, decimal dMinimo, decimal dMaximo)
		{
			if (uNumEditor.Value == System.DBNull.Value)
			{
				MessageBox.Show(sMensaje, "Validacion", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				uNumEditor.Focus();
				return false;
			}

			if (Convert.ToDecimal(uNumEditor.Value) < dMinimo)
			{
				MessageBox.Show(string.Format("El valor minimo es {0}", dMinimo), "Validacion", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				uNumEditor.Focus();
				return false;
			}

			if (Convert.ToDecimal(uNumEditor.Value) > dMaximo)
			{
				MessageBox.Show(string.Format("El valor maximo es {0}", dMaximo), "Validacion", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				uNumEditor.Focus();
				return false;
			}

			return true;
		}

		static public bool vbComboVacio(Infragistics.Win.UltraWinGrid.UltraCombo uCombo, string sMensaje)
		{
			bool bValida = true;

			if (uCombo.ActiveRow == null)
			{
				MessageBox.Show(sMensaje, "Validacion", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				uCombo.Focus();
				bValida = false;
			}

			return bValida;
		}

		static public bool vbFechaEnDocumentos(Infragistics.Win.UltraWinSchedule.UltraCalendarCombo uCCombo, string sMensaje, bool bValidaPeriodo, string sOperacion, C1.Data.C1DataSet ds)
		{
			bool bValida = true;

			if (uCCombo.Value == System.DBNull.Value)
			{
				MessageBox.Show(sMensaje, "Validacion", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				uCCombo.Focus();
				bValida = false;
			}
			else 
			{
				if (bValidaPeriodo)
				{
					if (Funcion.ValidaPeriodo(Convert.ToDateTime(uCCombo.Value), ds, sOperacion))
					{						
						MessageBox.Show(string.Format("No puede ingresar documentos en esta fecha, El Período '{0}' esta Cerrado,\n\n Consulte al Administrador", Convert.ToDateTime(uCCombo.Value).ToString("MMMM").ToUpper()), "Validacion", MessageBoxButtons.OK, MessageBoxIcon.Stop);
						uCCombo.Focus();
						bValida = false;
					}
				}
			}
			return bValida;
		}

		static public bool vbNumeracion(bool bAutomatico, int idCompra, int idTipoFactura, int iBodega, Infragistics.Win.UltraWinEditors.UltraTextEditor uTXT, C1.Data.C1DataSet ds)
		{
			bool bValida = true;

			if (!bAutomatico)
			{
				if (uTXT.Text.ToString().Trim().Length == 0)
				{
					MessageBox.Show("Ingrese el número del documento", "Validacion", MessageBoxButtons.OK, MessageBoxIcon.Stop);
					uTXT.Focus();
					bValida = false;
				}

				if (bValida)
				{
					char Pad = '0';
					uTXT.Text = uTXT.Text.ToString().PadLeft(9, Pad);

					if(Funcion.bEscalarSQL(ds, string.Format("Exec FacturacionValidaNumeracion {0}, {1}, {2}, '{3}'", idCompra, idTipoFactura, iBodega, uTXT.Text.ToString()), true))
					{
						MessageBox.Show(string.Format("El numero '{0}' ya existe", uTXT.Text), "Validacion", MessageBoxButtons.OK, MessageBoxIcon.Stop);
						uTXT.Focus();
						bValida = false;
					}
				}
			}

			return bValida;
		}

		static public bool vbIdentificacion(Infragistics.Win.UltraWinEditors.UltraTextEditor uTXT, int idTipoRuc, C1.Data.C1DataSet ds)
		{
			#region Valida longitud Campo
			if (uTXT.Text.ToString().Length == 0)
			{
				MessageBox.Show("Ingrese el numero de Identificacion", "Validacion", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				uTXT.Focus();
				return false;
			}
			
			if ((idTipoRuc == 1 || idTipoRuc == 4) && uTXT.Text.ToString().Length < 13)
			{
				MessageBox.Show("Para RUC debe ingresar 13 digitos", "Validacion", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				uTXT.Focus();
				return false;
			}
			
			if ((idTipoRuc == 3 || idTipoRuc == 6) && uTXT.Text.ToString().Length < 13)
			{
				MessageBox.Show("Para Pasaporte debe ingresar 13 digitos", "Validacion", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				uTXT.Focus();
				return false;
			}
			
			if ((idTipoRuc == 2 || idTipoRuc == 5) && uTXT.Text.ToString().Length != 10)
			{
				MessageBox.Show("Para Cedula debe ingresar 10 digitos", "Validacion", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				uTXT.Focus();
				return false;
			}
			
			int iTercerDigito = 0;

			if (idTipoRuc != 6)
				if (uTXT.Text.Length > 0) 
					iTercerDigito = int.Parse(uTXT.Text.ToString().Substring(2, 1));

			if (MenuLatinium.stIdDB != 38  )
			{
				if ((idTipoRuc == 1 /*|| idTipoRuc == 4*/) && iTercerDigito != 6 && uTXT.Text.ToString().Substring(10, 3) != "001")
				{
					MessageBox.Show("Para RUC los ultimos tres digitos deben ser 001", "Validacion", MessageBoxButtons.OK, MessageBoxIcon.Stop);
					uTXT.Focus();
					return false;
				}
			}
			
			if ((idTipoRuc == 1 /*|| idTipoRuc == 4*/) && iTercerDigito == 6 && uTXT.Text.ToString().Substring(9, 4) != "0001")
			{
				MessageBox.Show("Para RUC de empresa publica los ultimos cuatro digitos deben ser 0001", "Validacion", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				uTXT.Focus();
				return false;
			}
			
			if (idTipoRuc == 11 && uTXT.Text.ToString().Length != 13 && uTXT.Text.ToString() != "9999999999999")
			{
				MessageBox.Show("Para Consumidor Final debe registrar 13 digitos y todos deben ser 9", "Validacion", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				uTXT.Focus();
				return false;
			}
			#endregion Valida longitud Campo

			if (idTipoRuc != 11  && idTipoRuc != 4 )
			{
				if (!Funcion.sEscalarSQL(ds, string.Format("Select dbo.ValidaIdentificacion('{0}', {1})", uTXT.Text.ToString(), idTipoRuc)).StartsWith("Correcto"))
				{
					MessageBox.Show("Identificacion no valida", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
					uTXT.Focus();
					return false;
				}
			}

			return true;
		}	

		#region Valida Cliente Existe
		static public bool vbClienteExiste(Infragistics.Win.UltraWinEditors.UltraTextEditor uTXT, bool bProveedor, int idTipoFactura, int idFormaPago, C1.Data.C1DataSet ds)
		{
			bool bValida = Funcion.bEscalarSQL(ds, string.Format("Exec FacturacionValidaClienteExiste '{0}', {1}", uTXT.Text.ToString(), bProveedor), true);

			string sClienteProveedor = "Cliente";

			if (bProveedor) sClienteProveedor = "Proveedor";

			if (!bValida)
			{
				MessageBox.Show(string.Format("{0} no existe", sClienteProveedor), "Validacion", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				uTXT.Focus();
				return bValida;
			}

			if (idTipoFactura == 13 && idFormaPago == 6 && Funcion.iEscalarSQL(ds, string.Format("Select IsNull(ContadoCred, 0) From Cliente Where RUC = '{0}' And Proveedor = 0", uTXT.Text.ToString())) == 1)
			{											
				MessageBox.Show("El Cliente no tiene autorización para crédito", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Information);
				uTXT.Focus();
				return bValida;
			}

			return bValida;
		}
		#endregion Valida Cliente Existe

		static public bool vCosto(Infragistics.Win.UltraWinGrid.UltraGrid uGrid, decimal dBaseImponible, DateTime dtFecha, C1.Data.C1DataSet ds)
		{	
			bool bRes = true;
			decimal dCostoArticulo = 0.00m;
			decimal dCostoTotal = 0.00m;
			decimal dPorc5 = 0.00m;
     
			foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in uGrid.Rows.All)
			{
				dCostoArticulo = 
					Funcion.decEscalarSQL(ds, string.Format("Select dbo.CalculaUltimoCosto({0}, '{1}')", 
					(int)dr.Cells["idArticulo"].Value, dtFecha.ToString("yyyyMMdd")));					
				dCostoTotal = dCostoTotal + dCostoArticulo;
			}
      
			//dPorc5 = (dCostoTotal * 5.00m)/100.00m;
			
			if (dCostoTotal > dBaseImponible)			
			{
				MessageBox.Show("El valor de la factura es inferior al valor permitido", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Information);				
				return false;
			}

			return true;
		}

		static public bool vCostoCrediPoint(Infragistics.Win.UltraWinGrid.UltraGrid uGrid, decimal dBaseImponible, C1.Data.C1DataSet ds)
		{	
			bool bRes = true;
			decimal dCostoArticulo = 0.00m;
			decimal dCostoTotal = 0.00m;
			decimal dPorc5 = 0.00m;
     
			foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in uGrid.Rows.All)
			{
				dCostoArticulo = 
					Funcion.decEscalarSQL(ds, string.Format("Select Precio6 From Articulo Where idArticulo = {0}", 
					(int)dr.Cells["idArticulo"].Value));					
				dCostoTotal = dCostoTotal + dCostoArticulo;
			}
      
			//dPorc5 = (dCostoTotal * 5.00m)/100.00m;
			
			if (dCostoTotal > dBaseImponible)			
			{
				MessageBox.Show("El valor de la factura es inferior al valor permitido", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Information);				
				return false;
			}

			return true;
		}

		static public void vEliminaFilasVacias(Infragistics.Win.UltraWinGrid.UltraGrid uGrid, string sColumna)
		{	
			foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in uGrid.Rows.All)
				if (dr.Cells[sColumna].Value == System.DBNull.Value) dr.Delete();
		}

		static public bool vbGridVacio(Infragistics.Win.UltraWinGrid.UltraGrid uGrid, string sMensaje, string sColumna)
		{			
			if (uGrid.Rows.Count == 0)
			{
				MessageBox.Show(sMensaje, "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);

				if (sColumna.Length > 0)
				{
					uGrid.DisplayLayout.Bands[0].AddNew();
					uGrid.ActiveCell = uGrid.Rows[0].Cells[sColumna];
				}
				return false;
			}

			return true;
		}

		static public bool vbPointMovil(Infragistics.Win.UltraWinGrid.UltraGrid uGrid, C1.Data.C1DataSet ds)
		{
			foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in uGrid.Rows.All)
			{	
				if (dr.Cells["Codigo"].Value.ToString()!="COMP00001")
				{
					if (!Funcion.bEscalarSQL(ds, string.Format("Select Preciomovil from Articulo where Codigo = '{0}'", dr.Cells["Codigo"].Value.ToString()), true)) 
					{
						MessageBox.Show("El prodcuto seleccionado no se puede facturar con Point Movil", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Error);
						uGrid.DisplayLayout.Bands[0].AddNew();
						uGrid.ActiveCell = uGrid.Rows[0].Cells["Codigo"];								
						return false;
					}
				}
			}

			return true;
		}

		static public bool vbGrids(Infragistics.Win.UltraWinGrid.UltraGrid uGrid, string sColumna, int idTipoFactura)
		{
			#region Valida Articulos en el grid
			if (uGrid.Rows.Count == 0)
			{
				MessageBox.Show("Ingrese el detalle del documento", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Error);
				uGrid.DisplayLayout.Bands[0].AddNew();
				uGrid.ActiveCell = uGrid.Rows[0].Cells[sColumna];		
				return false;
			}
						
			if (idTipoFactura == 42)
			{		
				int iCont = 0;
				foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in uGrid.Rows.All)
				{				
					if (dr.Cells["idArticulo"].Value != System.DBNull.Value)				
						if (!(bool)dr.Cells["Borrar"].Value) 
							iCont++;
				}

				if (iCont == 0)
				{
					MessageBox.Show("Ingrese el detalle del documento", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Error);
					uGrid.DisplayLayout.Bands[0].AddNew();
					uGrid.ActiveCell = uGrid.Rows[0].Cells[sColumna];		
					return false;
				}
			}
			#endregion Valida Articulos en el grid

			#region Valida Cantidad - Precio
			foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in uGrid.Rows.All)
			{				
				if (dr.Cells["idArticulo"].Value != System.DBNull.Value)
				{
					if (idTipoFactura != 42)
					{
						if (Convert.ToDecimal(dr.Cells["Precio"].Value) == 0.00m)
						{
							MessageBox.Show(string.Format("Digite el precio del item '{0}'.", dr.Cells["Articulo"].Value.ToString()), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Error);
							uGrid.ActiveRow = dr;
							uGrid.ActiveRow.Selected = true;
							return false;
						}		
			
//						if (Convert.ToDecimal(dr.Cells["Subtotal"].Value) < 0.00m)
//						{	
//							MessageBox.Show(string.Format("El Subtotal no puede ser menor a Cero, '{0}'", dr.Cells["Codigo"].Value.ToString()), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Error);
//							uGrid.ActiveRow = dr;
//							uGrid.ActiveRow.Selected = true;
//							return false;
//						}
					}
					
					if (idTipoFactura != 54)
					{
						if ((int)dr.Cells["Cantidad"].Value == 0)
						{
							MessageBox.Show(string.Format("Digite la cantidad del item '{0}'.", dr.Cells["Articulo"].Value.ToString()), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Error);
							uGrid.ActiveRow = dr;
							uGrid.ActiveRow.Selected = true;
							return false;
						}
					}				
				}
			}
			#endregion Valida Cantidad - Precio
			
			return true;
		}

		static public bool vbValidaCantidadCargaDYD(Infragistics.Win.UltraWinGrid.UltraGrid uGrid)
		{
			int iContP = 0;
			int iContT = 0;

			foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in uGrid.Rows.All)
			{
				if (dr.Cells["idArticulo"].Value != System.DBNull.Value)
				{
					if ((int)dr.Cells["idArticulo"].Value == 8335)
					{
						iContP++;
						
						if ((int)dr.Cells["Tipo"].Value == 8) iContT++;
					}					
				}
			}

			if (iContP > 1)
			{
				MessageBox.Show("No puede cargar mas de un Producto D&D", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Error);				
				return false;
			}

			if (iContP != iContT)
			{
				MessageBox.Show("Producto D&D no puede ser cargado de forma manual", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Error);				
				return false;
			}

			return true;
		}

		static public bool ValidaTotales(decimal dTotal)
		{
			if (dTotal == 0.00m) 
			{
				MessageBox.Show("No puede Grabar el documento en Cero", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Error);				
				return false;
			}

			return true;
		}

		#region Entrega a Doimicilio
		static public bool EntregaDomicilio(Infragistics.Win.UltraWinEditors.UltraCheckEditor uCheck, Infragistics.Win.UltraWinGrid.UltraCombo uCombo, Infragistics.Win.UltraWinSchedule.UltraCalendarCombo uCCombo, Infragistics.Win.UltraWinSchedule.UltraCalendarCombo uCCombo1, int iTipoVenta)
		{
			if (uCheck.Checked)
			{
//				if (iTipoVenta == 1 )
//				{
//					if (uCombo.Rows.Count == 0)
//					{
//						MessageBox.Show("El Cliente No Tiene Direcciones para la Entrega", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
//						uCheck.Checked	= false;
//						uCheck.Focus();
//						return false;
//					}
//				
//					if (uCombo.ActiveRow == null)
//					{
//						MessageBox.Show("Seleccione la Dirección de Entrega", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
//						uCombo.Focus();				
//						return false;
//					}
//				}

				if (uCCombo.Value == System.DBNull.Value)
				{
					MessageBox.Show("Ingrese la Fecha De Entrega", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
					uCCombo.Focus();					
					return false;
				}
				
				if (uCCombo1.Value == System.DBNull.Value)
				{
					MessageBox.Show("Ingrese la Hora de Entrega", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
					uCCombo1.Focus();					
					return false;
				}				
			}

			return true;
		}
		#endregion Entrega a Doimicilio

		#region Curier
		static public bool Curier(Infragistics.Win.UltraWinEditors.UltraCheckEditor uCheck, Infragistics.Win.UltraWinGrid.UltraCombo uCombo, Infragistics.Win.UltraWinSchedule.UltraCalendarCombo uCCombo,Infragistics.Win.UltraWinEditors.UltraTextEditor uTXT, int iTipoVenta)
		{
			bool bValida = true;
			if (uCheck.Checked)
			{
				if (iTipoVenta == 2)
				{
					if (uCombo.Rows.Count == 0)
					{
						MessageBox.Show("Seleccione el cliente para servicio de Curier", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
						uCheck.Checked	= false;
						uCheck.Focus();
						return false;
					}
					if (uTXT.Text.ToString().Trim().Length == 0)
					{
						MessageBox.Show("Ingrese el número del guia", "Validacion", MessageBoxButtons.OK, MessageBoxIcon.Stop);
						uTXT.Focus();
						bValida = false;
					}
					/*if (uCombo.ActiveRow == null)
					{
						MessageBox.Show("Seleccione la Dirección de Entrega", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
						uCombo.Focus();				
						return false;
					}*/
				

					if (uCCombo.Value == System.DBNull.Value)
					{
						MessageBox.Show("Ingrese la Fecha De Entrega al curier", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
						uCCombo.Focus();					
						return false;
					}
				
					/*if (uCCombo1.Value == System.DBNull.Value)
					{
						MessageBox.Show("Ingrese la Hora de Entrega", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
						uCCombo1.Focus();					
						return false;
					}			*/	
				}
			}
			return true;
		}
		#endregion Curier

		#region Valida Existencia Inventario
		static public bool bValidaFacturacionExistencias(Infragistics.Win.UltraWinGrid.UltraGrid uGrid, DateTime dtFecha, int idCompra, int iBodega, C1.Data.C1DataSet ds)
		{
			bool bSaldo0 = false;
			bool bSaldo = false;
			int iSaldo = 0;
			string CodArticulo = "";

			#region Valida Saldos
			foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in uGrid.Rows.All)
			{
				if (bSaldo0 || bSaldo) break;
				
				int iCantidad = 0;
				int idArticulo = (int)dr.Cells["idArticulo"].Value;
				
				iSaldo = Funcion.iEscalarSQL(ds, string.Format("Exec FacturacionRetornaExistencia '{0}', {1}, {2}, {3}", 
					dtFecha.ToString("yyyyMMdd"), idCompra, iBodega, idArticulo));

				foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr1 in uGrid.Rows.All)
				{					
					if (Funcion.bEscalarSQL(ds, string.Format("Select Inventario From Articulo Where idArticulo = {0}", idArticulo), true))
					{
						if (idArticulo == (int)dr1.Cells["idArticulo"].Value) iCantidad = iCantidad + int.Parse(dr1.Cells["Cantidad"].Value.ToString());
																	
						if (iSaldo == 0) 
						{
							CodArticulo = dr.Cells["Codigo"].Value.ToString();
							uGrid.ActiveRow = dr;
							uGrid.ActiveRow.Selected = true;
							bSaldo0 = true;
							break;
						}
																												
						if (iCantidad > iSaldo) 
						{
							CodArticulo = dr.Cells["Codigo"].Value.ToString();
							uGrid.ActiveRow = dr;
							uGrid.ActiveRow.Selected = true;
							bSaldo = true;
							break;
						}
					}
				}				
			}		
			#endregion Valida Saldos								
			
			#region Saldos
			if (bSaldo0)
			{
				MessageBox.Show(string.Format("El artículo '{0}' no tiene existencia", CodArticulo), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Error);
				return false;
			}
			
			if (bSaldo)
			{
				MessageBox.Show(string.Format("El artículo '{0}' tiene en existencia {1}", CodArticulo, iSaldo), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Error);
				return false;
			}
			#endregion Saldos

			return true;
		}

		static public bool bValidaExistencias(Infragistics.Win.UltraWinGrid.UltraGrid uGrid, DateTime dtFecha, int idCompra, int iBodega, int idTipoFactura, int idOrigen, C1.Data.C1DataSet ds)
		{
			bool bSaldo0 = false;
			bool bSaldo = false;
			int iSaldo = 0;
			string CodArticulo = "";

			#region Valida Saldos
			foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in uGrid.Rows.All)
			{
				if (bSaldo0 || bSaldo) break;
				
				int iCantidad = 0;
				int idArticulo = (int)dr.Cells["idArticulo"].Value;
				
				iSaldo = Funcion.iEscalarSQL(ds, string.Format("Select dbo.RetornaExistencia('{0}', {1}, {2}, {3}, {4}, {5})", 
					dtFecha.ToString("yyyyMMdd"), idCompra, iBodega, idArticulo, idTipoFactura, idOrigen)); 

				foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr1 in uGrid.Rows.All)
				{					
					if (Funcion.bEscalarSQL(ds, string.Format("Select Inventario From Articulo Where idArticulo = {0}", idArticulo), true))
					{
						if (idArticulo == (int)dr1.Cells["idArticulo"].Value) iCantidad = iCantidad + int.Parse(dr1.Cells["Cantidad"].Value.ToString());
																	
						if (iSaldo == 0) 
						{
							CodArticulo = dr.Cells["Codigo"].Value.ToString();
							uGrid.ActiveRow = dr;
							uGrid.ActiveRow.Selected = true;
							bSaldo0 = true;
							break;
						}
																												
						if (iCantidad > iSaldo) 
						{
							CodArticulo = dr.Cells["Codigo"].Value.ToString();
							uGrid.ActiveRow = dr;
							uGrid.ActiveRow.Selected = true;
							bSaldo = true;
							break;
						}
					}
				}				
			}		
			#endregion Valida Saldos								
			
			#region Saldos
			if (bSaldo0)
			{
				MessageBox.Show(string.Format("El artículo '{0}' no tiene existencia", CodArticulo), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Error);
				return false;
			}
			
			if (bSaldo)
			{
				MessageBox.Show(string.Format("El artículo '{0}' tiene en existencia {1}", CodArticulo, iSaldo), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Error);
				return false;
			}
			#endregion Saldos

			return true;
		}
		#endregion Valida Existencia Inventario
	
		static public bool bValidaRepetidosTF(Infragistics.Win.UltraWinGrid.UltraGrid uGrid)
		{
			int iCont = 0;
			string sCodArticulo = "";
	
			#region Valida repetidos
			foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in uGrid.Rows.All)
			{					
				int idArticulo = (int)dr.Cells["idArticulo"].Value;
				sCodArticulo = dr.Cells["Codigo"].Value.ToString();

				iCont = 0;

				foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr1 in uGrid.Rows.All)
					if (idArticulo == (int)dr1.Cells["idArticulo"].Value) iCont++;
			
				if (iCont > 1) break;
			}	
			#endregion Valida repetidos
						
			if (iCont > 1)
			{
				MessageBox.Show(string.Format("El artículo '{0}' ha sido registrado {1} veces\n\nno puede repetir los articulos en la transferencia\n\nSi requiere transferir mas de 1 escriba en la columna CANTIDAD el numero de articulos necesarios", sCodArticulo, iCont), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Error);
				return false;
			}		

			return true;
		}

		static public bool bValidaAbreListaArticulos(Infragistics.Win.UltraWinGrid.UltraGrid uGrid)
		{
			if (uGrid.ActiveRow.Cells["idArticulo"].Value != System.DBNull.Value)
			{
				if ((int)uGrid.ActiveRow.Cells["idArticulo"].Value == 0) 
				{
					if ((bool)uGrid.ActiveRow.Cells["Principal"].Value)
					{
						MessageBox.Show("No puede modificar el producto principal de un combo o promoción", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
						return false;
					}

					if ((int)uGrid.ActiveRow.Cells["Tipo"].Value == 5)
					{
						MessageBox.Show("No puede modificar garantia extendida", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
						return false;
					}
			
					if ((int)uGrid.ActiveRow.Cells["Tipo"].Value > 0)
					{
						MessageBox.Show("No puede modificar productos de un combo o promoción", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
						return false;
					}
				}
			}

			return true;
		}

		static public bool bValidaCambioArticuloVenta(Infragistics.Win.UltraWinGrid.BeforeCellUpdateEventArgs e, bool bModificaColor, bool bAcceso)
		{	
			if (e.Cell.Row.Cells["idArticulo"].Value != System.DBNull.Value)
			{			
				if (!bModificaColor && (int)e.Cell.Row.Cells["idArticulo"].Value > 0)
				{
					MessageBox.Show("Para reemplazar un articulo, elimine el articulo actual", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
					return false;
				}			
			}

			if (!bModificaColor && (bool)e.Cell.Row.Cells["Principal"].Value)
			{
				MessageBox.Show("No puede modificar el producto principal de un combo o promoción", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				return false;
			}

			if ((int)e.Cell.Row.Cells["Tipo"].Value == 5)
			{
				MessageBox.Show("No puede modificar garantia extendida", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				return false;
			}


			if ((int)e.Cell.Row.Cells["Tipo"].Value != 0)
			{
				if ((int)e.Cell.Row.Cells["Tipo"].Value != 8)
				{
					if ((int)e.Cell.Row.Cells["Tipo"].Value != 11)
					{
						if (!bModificaColor && !bAcceso)
						{
							MessageBox.Show("No puede modificar productos de un combo o promoción", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
							return false;
						}
					}
				}
			}

			return true;
		}

		static public bool bFechas(object oFecha)
		{            
			try
			{
				DateTime dtFecha = DateTime.Parse(oFecha.ToString());
				return true;
			}
			catch (FormatException e)
			{
				return false;
			}            
		}

		static public bool vbComparaDecimales(decimal dValor1, decimal dValor2, string sMensaje, bool bValidaIguales)
		{
			bool bRes = false;

			if (bValidaIguales) if (dValor1 >= dValor2) bRes = true;
			if (!bValidaIguales) if (dValor1 > dValor2) bRes = true;

			if (bRes)
			{
				MessageBox.Show(sMensaje, "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				return false;
			}

			return true;
		}

		static public bool vbComparaFechas(DateTime dtFecha1, DateTime dtFecha2, Infragistics.Win.UltraWinSchedule.UltraCalendarCombo uCCombo,  string sMensaje,  bool bValidaIguales)
		{
			bool bRes = false;

			if (bValidaIguales) if (dtFecha1 >= dtFecha2) bRes = true;
			if (!bValidaIguales) if (dtFecha1 > dtFecha2) bRes = true;

			if (bRes)
			{
				MessageBox.Show(sMensaje, "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);				
				return false;
			}

			return true;
		}

		static public bool vbCarteraEnTransito(int idEntidadFinanciera, int iNumeroLote, bool bConfirmado)
		{	
			string sMensaje = "";

			if (idEntidadFinanciera == 2 && iNumeroLote > 0 && !bConfirmado) sMensaje = "Cartera solidario en transito";
			if (idEntidadFinanciera == 2 && iNumeroLote > 0 && bConfirmado) sMensaje = "Cartera solidario vendida";
			if (idEntidadFinanciera == 4 && iNumeroLote > 0 && !bConfirmado) sMensaje = "Cartera en transito";
			if (idEntidadFinanciera == 13 && iNumeroLote > 0 && !bConfirmado) sMensaje = "Cartera en transito";
			if (idEntidadFinanciera == 14 && iNumeroLote > 0 && !bConfirmado) sMensaje = "Cartera en transito";
			//if (idEntidadFinanciera == 15 && iNumeroLote > 0 && !bConfirmado) sMensaje = "Cartera en transito";
			if (idEntidadFinanciera == 16 && iNumeroLote > 0 && !bConfirmado) sMensaje = "Cartera en transito";

			if (sMensaje.Length > 0)
			{
				MessageBox.Show(sMensaje, "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);				
				return false;
			}

			return true;
		}

		static public bool vbValidaValorDctoOrigen(int idOrigen, decimal dTotal, C1.Data.C1DataSet ds)
		{
			if (idOrigen > 0)
			{				
				decimal dTotalOrigen = Funcion.decEscalarSQL(ds, string.Format("Select Total From Compra Where idTipoFactura = 1 And idCompra = {0}", idOrigen));

				if (Math.Round(dTotal, 2) > Math.Round(dTotalOrigen, 2))
				{
					MessageBox.Show(string.Format("El valor de la Nota de Crédito {0} no puede ser mayor al valor de la Factura {1}", 
						Math.Round(dTotal, 2), Math.Round(dTotalOrigen, 2)), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);					
					return false;
				}
			}

			return true;
		}

		static public bool vbValidaAnulacionCobrosCuotasPosteriores(int idAnticipo, int iTipo, C1.Data.C1DataSet ds)
		{
			string sSQL = string.Format("Exec ValidaAnulacionCobrosCuotasPosteriores {0}, {1}", idAnticipo, iTipo);
			bool bRes = Funcion.bEscalarSQL(ds, sSQL, true);

			if (!bRes)
			{
				MessageBox.Show("Existen cuotas posteriores ya cobradas", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);				
				return false;
			}

			return true;
		}

		static public bool vbValidaSelGrids(Infragistics.Win.UltraWinGrid.UltraGrid uGrid, string sColumna, string sMensaje)
		{			
			int iCont = 0;
			foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in uGrid.Rows.All)
				if ((bool)dr.Cells[sColumna].Value) iCont++;

			if (iCont == 0)
			{
				MessageBox.Show(sMensaje, "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);				
				return false;
			}

			return true;
		}

		static public bool vbValidaValoresNotasDeCredito(object sender, Infragistics.Win.UltraWinGrid.CellEventArgs e, int iCampo, C1.Data.C1DataSet ds) 
		{
			string sSQL = string.Format("Exec NDCValidaValoresFacturacion {0}, {1}, {2}", 
										(int)e.Cell.Row.Cells["idDetCompra"].Value, (int)e.Cell.Row.Cells["idArticulo"].Value, iCampo);
			decimal dValorOrigen = Funcion.decEscalarSQL(ds, sSQL);	

			if (iCampo == 1)
			{
				if (Convert.ToDecimal(e.Cell.Row.Cells["Precio"].Value) > dValorOrigen)
				{
					MessageBox.Show(string.Format("El precio del articulo no puede ser mayor a {0}", dValorOrigen), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);				
					e.Cell.Row.Cells["Precio"].Value = dValorOrigen;
					return false;
				}
			}

			if (iCampo == 2)
			{
				if (Convert.ToDecimal(e.Cell.Row.Cells["DescuentoPorc"].Value) < dValorOrigen)
				{
					MessageBox.Show(string.Format("El descuento del articulo no puede ser menor a {0}", dValorOrigen), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);				
					e.Cell.Row.Cells["DescuentoPorc"].Value = dValorOrigen;
					return false;
				}
			}

			return true;
		}

		static public bool vbValidaFechasDesdeHasta(Infragistics.Win.UltraWinSchedule.UltraCalendarCombo dtDesde, Infragistics.Win.UltraWinSchedule.UltraCalendarCombo dtHasta)
		{
			if (dtDesde.Value == System.DBNull.Value)
			{
				MessageBox.Show("Seleccione la Fecha Inicial", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				dtDesde.Focus();
				return false;
			}

			if (dtHasta.Value == System.DBNull.Value)
			{
				MessageBox.Show("Seleccione la Fecha Final", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				dtHasta.Focus();
				return false;
			}

			if (Convert.ToDateTime(dtDesde.Value) > Convert.ToDateTime(dtHasta.Value))
			{
				MessageBox.Show("La Fecha Inicial no puede ser mayor a la Fecha Final", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				dtHasta.Focus();
				return false;
			}

			return true;
		}

		static public bool vbValidaFechasDesdeHasta(Infragistics.Win.UltraWinSchedule.UltraCalendarCombo dtDesde, Infragistics.Win.UltraWinSchedule.UltraCalendarCombo dtHasta, string lblDesde, string lblHasta)
		{
			if (dtDesde.Value == System.DBNull.Value)
			{
				MessageBox.Show("Seleccione la "+ lblDesde, "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				dtDesde.Focus();
				return false;
			}

			if (dtHasta.Value == System.DBNull.Value)
			{
				MessageBox.Show("Seleccione la "+ lblHasta, "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				dtHasta.Focus();
				return false;
			}

			if (Convert.ToDateTime(dtDesde.Value) > Convert.ToDateTime(dtHasta.Value))
			{
				MessageBox.Show("La "+lblDesde+" no puede ser mayor a la "+lblHasta+"", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				dtHasta.Focus();
				return false;
			}

			return true;
		}

		static public bool vbValidaPlaca(Infragistics.Win.UltraWinEditors.UltraTextEditor uTXT)
		{
			if (uTXT.Text.ToString().Length == 0)
			{
				MessageBox.Show("Ingrese el número de Placa", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				uTXT.Focus();
				return false;
			}

			if (uTXT.Text.ToString().Length < 7)
			{
				MessageBox.Show("El número de Placa debe tener 7 caracteres", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				uTXT.Focus();
				return false;
			}

			string sLetras = uTXT.Text.ToString().Substring(0, 3);
			string sNumeros = uTXT.Text.ToString().Substring(3, 4);

			if (!Regex.IsMatch(sLetras, @"^[a-zA-Z]+$"))
			{
				MessageBox.Show("La placa de todo vehiculo empieza con 3 caracteres", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				uTXT.Focus();
				return false;
			}

			if (!Regex.IsMatch(sLetras.Substring(0, 1), @"[AWQBGSUIPCLYHRJXMKOVTENZ]"))
			{
				MessageBox.Show("La placa registrada no es valida", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				uTXT.Focus();
				return false;
			}

			if (Regex.IsMatch(sLetras.Substring(1, 1), @"[EX]"))
			{
				MessageBox.Show("La placa registrada no es valida", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				uTXT.Focus();
				return false;
			}

			if (!Regex.IsMatch(sNumeros, @"^[0-9]+$"))
			{
				MessageBox.Show("La placa de todo vehiculo termina con 4 digitos", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				uTXT.Focus();
				return false;
			}
			
			return true;
		}

		static public bool vbValidaFacturasCliente (string sCedula, C1.Data.C1DataSet ds)
		{
			if (Funcion.iEscalarSQL(ds, string.Format("Select COUNT(*) From Compra Where idTipoFactura = 1 And idFormaPago In (3, 9, 11) And Borrar = 0 And Estado In (10, 11) And Convert(Date, Fecha) >= '20150101' And Ruc = '{0}'", 
				sCedula)) > 0) return false;	

			return true;
		}

		static public bool bValidaMargen(Infragistics.Win.UltraWinGrid.UltraGrid uGrid, int idTipoPrecio, string sUsuario, decimal dSubtotal, decimal dPrecio, int idBanco, int idTarjeta, int idPlan, int idPlazo, int idCompra, int iEstado, C1.Data.C1DataSet ds)
		{
			decimal dCosto = 0;

			decimal dComision = Funcion.decEscalarSQL(ds, string.Format("Select dbo.PorcComisionTarjeta({0}, {1}, {2}, {3})", idBanco, idTarjeta, idPlan, idPlazo));

			foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in uGrid.Rows.All)
				if (dr.Cells["idArticulo"].Value != System.DBNull.Value)
					dCosto = dCosto + Funcion.decEscalarSQL(ds, string.Format("Select ISNULL(CostoUltimo, 0) From Articulo Where idArticulo = {0}", (int)dr.Cells["idArticulo"].Value));
			
			dCosto = dCosto + dComision;

			string sSQL = string.Format("Exec DescuentoValidaFacturacion {0}, '{1}', {2}, {3}, {4}, {5}, {6}", idTipoPrecio, sUsuario, dSubtotal, dCosto, idCompra, iEstado, dPrecio);
			bool bResultado = Funcion.bEscalarSQL(ds, sSQL, true);

			if (!bResultado)
			{
				MessageBox.Show("No puede facturar bajo el margen asignado", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				return false;
			}
			else return true;
		}
	}
}