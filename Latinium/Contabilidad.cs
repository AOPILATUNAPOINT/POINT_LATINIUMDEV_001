using System;
using System.Globalization;
using C1.Data;
using System.Data;
using System.Data.Common;
using System.Data.SqlClient;
using System.Windows.Forms;

namespace Latinium
{
	public class Contabilidad
	{
		/// <summary>
		/// Busca la cuenta contable que tiene el cliente basada en el grupo
		/// </summary>
		/// <param name="IdCliente"></param>
		/// <returns></returns>
		public static int CuentaCliente(C1.Data.C1DataSet ds, int IdCliente)
		{
			int IdGrupo = 0;
			int IdCuenta = 0;
			string stBusca = "SELECT idGrupoCliente FROM Cliente WHERE idCliente = "
				+ IdCliente.ToString();
			ds.Schema.Connections[0].Open();
			C1.Data.SchemaObjects.Connection miConeccion = ds.Schema.Connections[0];
			miConeccion = ds.Schema.Connections[0];

			SqlCommand cmdBusca = new 
				SqlCommand (stBusca,	(SqlConnection) miConeccion.DbConnection);

			try
			{
				IdGrupo = (int) cmdBusca.ExecuteScalar();
			}
			catch (System.Exception ex)
			{
				MessageBox.Show(ex.Message);
			}
			// Fin de busqueda de grupo en cliente

			if (IdGrupo > 0) 
			{
				string stBuscaCuenta = "SELECT idCuenta FROM ClienteGrupo WHERE idGrupoCliente = "
					+ IdGrupo.ToString();
				SqlCommand cmdBuscaCuenta = new 
					SqlCommand (stBuscaCuenta,	(SqlConnection) miConeccion.DbConnection);
				try
				{
					IdCuenta = (int) cmdBuscaCuenta.ExecuteScalar();
				}
				catch (System.Exception ex)
				{
					MessageBox.Show(ex.Message);
				}
			} // Fin de busqueda de cuenta de grupo
			ds.Schema.Connections[0].Close();
			return IdCuenta;
		}

		/// <summary>
		/// Borra Asiento
		/// </summary>
		/// <param name="IdAsiento"></param>
		/// <returns></returns>
		public static bool BorrarAsiento(C1.Data.C1DataSet ds, int IdAsiento)
		{
			if (IdAsiento == 0) return true;
			string stBusca = "DELETE Asiento WHERE idAsiento = "
				+ IdAsiento.ToString();
			return Funcion.EjecutaSQL(ds, stBusca);
		}

		/// <summary>
		/// Borra Retencion
		/// </summary>
		/// <param name="IdAsiento"></param>
		/// <returns></returns>
		public static bool BorrarRetencion(C1.Data.C1DataSet ds, int IdFactura)
		{
			if (IdFactura == 0) return true;
			string stBusca = "DELETE Retencion WHERE idCompra = "
				+ IdFactura.ToString();
			return Funcion.EjecutaSQL(ds, stBusca);
		}


		/// <summary>
		/// Asiento de Retencion
		/// </summary>
		/// <param name="ds"></param>
		/// <param name="IdCompra"></param>
		/// <param name="IdAsiento"></param>
		/// <param name="ITipo"></param>
		static public void AsientoRetencion(C1.Data.C1DataSet ds, int IdCompra, int IdTipo)
		{
			CultureInfo us = new CultureInfo("en-US");
			string stSelect = "Select Retencion From Compra Where idCompra = " 
				+ IdCompra.ToString();
			bool bRetencion = Funcion.bEscalarSQL(ds, stSelect, true);

			if (!bRetencion)
			{
				MessageBox.Show("No existe retencion de esta factura");
				return;
			}
			stSelect = "Select Fecha From Retencion Where idCompra = " 
				+ IdCompra.ToString();
			DateTime dtFecha = Funcion.dtEscalarSQL(ds, stSelect);

			int IdProyecto = 0;
			int IdSubProyecto = 0;
			string stNumero = "";
			string stNombre = "";
			stSelect = "Select idProyecto, idSubProyecto, Numero, Cliente.Nombre "
				+ "From Compra Inner Join Cliente "
				+ "On Compra.idCliente = Cliente.idCliente "
				+ "Where idCompra = " + IdCompra.ToString();

			SqlDataReader dr = Funcion.rEscalarSQL(ds, stSelect, true);
			while(dr.Read())
			{
				IdProyecto = dr.GetInt32(0);
				IdSubProyecto = dr.GetInt32(1);
				stNumero = "";
				if (dr.GetValue(2) != System.DBNull.Value)
					stNumero = dr.GetString(2);
				stNombre = "";
				if (dr.GetValue(3) != System.DBNull.Value)
					stNombre = dr.GetString(3);
			}
			dr.Close();
			ds.Schema.Connections[0].Close();
		
			string stCodAsiento = Contabilidad.NumeroAsiento(ds, 1);
			stSelect = "Insert Into Asiento "
				+ "(Descripcion, Fecha, idTipoAsiento, CodAsiento, Numero"
				+ ", idProyecto, idSubProyecto) VALUES ('"
				+ stNombre.Trim() + " Retencion', '" + dtFecha.ToString("yyyyMMdd")
				+ "', 1, '"	+ stCodAsiento + "', '" + stNumero.Trim() + "', " 
				+ IdProyecto.ToString()	+ ", " + IdSubProyecto.ToString() + ")";

			Funcion.EjecutaSQL(ds, stSelect);
		}

		/// <summary>
		/// Precio * Cantidad agrupado por grupo de articulo
		/// </summary>
		/// 
		static public void AsientoFactura(C1.Data.C1DataSet ds, int IdCompra, int IdAsiento, int ITipo, C1.Data.C1DataRow drSeteo)
		{
			CultureInfo us = new CultureInfo("en-US");

			string stGrupo = "3";
			string stCalculo = "DetCompra.Precio*DetCompra.Cantidad + DetCompra.Flete";
			string stDebe = "Haber";
			switch(ITipo)
			{
				case 1: // Factura de Venta
					stGrupo = "3";
					break;
				case 2: // Factura de Compra
					stGrupo = "2";
					stDebe = "Debe";
					break;
				case 3: // Transferencia Recibida
					stGrupo = "2, ArticuloGrupo.idCuentas2";
					stDebe = "Debe";
					break;
				case 4: // Devolucion de Compra
					stGrupo = "2";
					stDebe = "Haber";
					break;
				case 5: // Devolucion de Venta
					stGrupo = "6";
					stDebe = "Debe";
					break;
				case 6: // Transferencia Enviada
					stGrupo = "2, ArticuloGrupo.idCuentas2";
					stDebe = "Debe";
					break;
			}
		
			string stSelect = "SELECT Sum(" + stCalculo + ") AS "
				+ " TotalGrupo, ArticuloGrupo.idCuentas" + stGrupo
				+ ", IdProyecto, IdSubProyecto "
				+ "FROM DetCompra INNER JOIN Articulo "
				+ "ON DetCompra.idArticulo = Articulo.idArticulo "
				+ "LEFT JOIN ArticuloGrupo ON ArticuloGrupo.idGrupoArticulo "
				+ "= Articulo.idGrupoArticulo "
				+ "Where DetCompra.idCompra = " + IdCompra.ToString().Trim()
				+ " GROUP BY ArticuloGrupo.idCuentas2, IdProyecto, IdSubProyecto";

			ds.Schema.Connections[0].Open();
			C1.Data.SchemaObjects.Connection miConeccion = ds.Schema.Connections[0];
			SqlCommand cmdCant = new SqlCommand (
				stSelect, (SqlConnection) miConeccion.DbConnection);
			try
			{
				SqlDataReader miReader = cmdCant.ExecuteReader();
				double [] dSubTotal = new double[500];
				int [] idGrupo = new int[500];
				int [] iProy = new int[500];
				int [] iSubP = new int[500];
				int [] idGrupo2 = new int[500];
				int i = 0;
				
				while (miReader.Read())
				{
					dSubTotal[i] = miReader.GetDouble(0);
					if (ITipo == 3)
					{
						if (miReader.GetValue(3) != System.DBNull.Value)
							iProy[i] = miReader.GetInt32(3);
						if (miReader.GetValue(4) != System.DBNull.Value)
							iSubP[i] = miReader.GetInt32(4);
					}
					else
					{
						if (miReader.GetValue(2) != System.DBNull.Value)
							iProy[i] = miReader.GetInt32(2);
						if (miReader.GetValue(3) != System.DBNull.Value)
							iSubP[i] = miReader.GetInt32(3);
					}
					try
					{
						idGrupo[i] = miReader.GetInt32(1);
						if (ITipo == 3)
							idGrupo2[i] = miReader.GetInt32(2);
					}
					catch
					{
						idGrupo[i] = 0;
					}
					i++;
				}
				miReader.Close();

				bool bIngEgrInterno = false;
				if (ITipo == 3 || ITipo == 6)
				{
					string stDestino = "SELECT IngEgrInterno FROM SeteoF";
					bIngEgrInterno = Funcion.bEscalarSQL(ds, stDestino, false);
				}

				for (int j=0; j< i; j++)
				{
					int iCuentaDebe = idGrupo[j];
					if (ITipo == 3) // Egreso de Bodega
						if (!bIngEgrInterno) iCuentaDebe = 0;

					string stValor = dSubTotal[j].ToString("0.00", us);
					string stProy = ", " + iProy[j].ToString("0.00", us);
					string stSubP = ", " + iSubP[j].ToString("0.00", us);
					stSelect = "INSERT INTO AsientoDetalle (idCuenta, idAsiento, " 
						+ stDebe + ", IdProyecto, IdSubProyecto) VALUES ("
						+ iCuentaDebe.ToString() + ", "
						+ IdAsiento.ToString().Trim() + ", "
						+ stValor + stProy + stSubP + ")";
					SqlCommand cmdActualiza = new SqlCommand (
						stSelect, (SqlConnection) miConeccion.DbConnection);
					cmdActualiza.Transaction = (SqlTransaction)
						miConeccion.DbTransaction;
					cmdActualiza.ExecuteNonQuery();
					if (ITipo == 3 || ITipo == 6)
					{
						// Destino de Egreso o Ingreso
						int iCuentaHaber = idGrupo[j];
						// Ingreso a Bodega sin cuenta
						if (!bIngEgrInterno && ITipo == 6)	iCuentaHaber = 0;

						stSelect = "INSERT INTO AsientoDetalle (idCuenta, idAsiento, " 
							+ "Haber, idProyecto, idSubProyecto) VALUES ("
							+ iCuentaHaber.ToString() + ", "
							+ IdAsiento.ToString().Trim() + ", "
							+ stValor + stProy + stSubP + ")";
						cmdActualiza = new SqlCommand (
							stSelect, (SqlConnection) miConeccion.DbConnection);
						cmdActualiza.Transaction = (SqlTransaction)
							miConeccion.DbTransaction;
						cmdActualiza.ExecuteNonQuery();
					}
				}
			}
			catch (System.Exception ex)
			{
				MessageBox.Show(ex.Message);
			}
			ds.Schema.Connections[0].Close();

		}

		/// <summary>
		/// devuelve cuenta de asigncacion de cuentas
		/// </summary>
		/// <param name="ds"></param>
		/// <param name="IdCuenta"></param>
		/// <returns></returns>
		static public int AsignaCuenta(C1.Data.C1DataSet ds, int IdCuenta)
		{
			int iCuenta = 0;

			try
			{
				C1DataRow dr = ds.TableViews["AsignaCuenta"].Rows[IdCuenta-1];
				if (dr["idCuenta"] != System.DBNull.Value) iCuenta = (int) dr["idCuenta"];
			}
			catch(System.Exception ex)
			{
				MessageBox.Show(ex.Message, "Error en "+ IdCuenta.ToString());
			}
			return iCuenta;
		}

		/// <summary>
		/// Devuelve el valor maximo o minimo de la fecha de los asientos
		/// si iTipo = 1 devuelve la  fecha maxima en caso contrario la fecha minima
		/// </summary>
		/// <param name="ds"></param>
		static public DateTime FechaAsiento(C1.Data.C1DataSet ds, int iTipo)
		{
			ds.Schema.Connections[0].Open();
			C1.Data.SchemaObjects.Connection miConeccion = ds.Schema.Connections[0];

			string stOrden = "";
			if (iTipo == 1) stOrden = "DESC";

			DateTime dtFecha = (DateTime) DateTime.Now;
			try
			{
				string stSelect = "SELECT TOP 1 Fecha FROM Asiento " + 
					" ORDER BY Fecha " + stOrden;
				SqlCommand cmdGrupo = new SqlCommand (
					stSelect, (SqlConnection) miConeccion.DbConnection);
				dtFecha = (DateTime) cmdGrupo.ExecuteScalar();
			}
			catch(System.Exception ex)
			{
				Console.WriteLine(ex.Message + ". En Función FechaAsiento");
			}
			ds.Schema.Connections[0].Close();
			return dtFecha;
		}
		/// <summary>
		/// Calculo de totales de Debe y de Haber de Asiento
		/// </summary>
		/// <param name="ds"></param>
		static public void TotalAsiento(C1.Data.C1DataSet ds, int iTipo)
		{
			CultureInfo us = new CultureInfo("en-US");

			double fTotalDebe = 0;
			double fTotalHaber = 0;
			try
			{
				string stSelect = "SELECT SUM(Debe) FROM AsientoDetalle" + 
					" WHERE idAsiento = " + iTipo.ToString();
				fTotalDebe = Funcion.dEscalarSQL(ds, stSelect, false);

				stSelect = "SELECT SUM(Haber) FROM AsientoDetalle" + 
					" WHERE idAsiento = " + iTipo.ToString();
				fTotalHaber = Funcion.dEscalarSQL(ds, stSelect, false);

				fTotalDebe = Math.Round(fTotalDebe,2);
				fTotalHaber = Math.Round(fTotalHaber,2);
				stSelect = "UPDATE Asiento SET Total = " +
					fTotalDebe.ToString("0.00",us) + 
					", TotalH = " + fTotalHaber.ToString("0.00", us) + 
					" WHERE idAsiento = " + iTipo.ToString();
				Funcion.EjecutaSQL(ds, stSelect);
			}
			catch(System.Exception ex)
			{
				MessageBox.Show(ex.Message, "En Función CalculoTotal");
			}
		}


		/// <summary>
		/// Devuelve el valor del cheque de la retencion
		/// y asigna  el idAsiento a la retencion
		/// </summary>
		/// <param name="ds"></param>
		static public string FechaAsiento(C1.Data.C1DataSet ds, int IdFactura, int IdAsiento)
		{
			ds.Schema.Connections[0].Open();
			C1.Data.SchemaObjects.Connection miConeccion = ds.Schema.Connections[0];


			string stFecha = "";
			try
			{
				string stSelect = "SELECT Cheque FROM Retencion " + 
					" WHERE idCompra = " + IdFactura.ToString();
				SqlCommand cmdGrupo = new SqlCommand (
					stSelect, (SqlConnection) miConeccion.DbConnection);
				stFecha = (string) cmdGrupo.ExecuteScalar();

				stSelect = "UPDATE Retencion SET idAsiento = " 
					+ IdAsiento.ToString() + 
					" WHERE idCompra = " + IdFactura.ToString();
				cmdGrupo = new SqlCommand (
					stSelect, (SqlConnection) miConeccion.DbConnection);
				cmdGrupo.ExecuteNonQuery();

				// Actualiza Asiento en Ingreso y Egreso
				stSelect = "UPDATE Asiento SET Cheque = " + stFecha + 
					" WHERE idAsiento = " + IdAsiento.ToString()
					+ " And Asiento.idTipo IN (2,3) ";
				cmdGrupo = new SqlCommand (
					stSelect, (SqlConnection) miConeccion.DbConnection);
				cmdGrupo.ExecuteNonQuery();

			}
			catch(System.Exception ex)
			{
				Console.WriteLine(ex.Message + ". En Función ChequeRetencion");
			}
			ds.Schema.Connections[0].Close();
			return stFecha;
		}


		/// <summary>
		/// Configura tipo de cuentas del Plan Activo Pasivo etc.
		/// </summary>
		/// <param name="ds"></param>
		static public void TipoCuenta(C1.Data.C1DataSet ds, string stGrupo, int idTipo)
		{
			foreach(string stCodigo in stGrupo.Split(','))
			{
				ds.Schema.Connections[0].Open();
				C1.Data.SchemaObjects.Connection miConeccion = ds.Schema.Connections[0];
				SqlCommand cmdActualiza = new SqlCommand (
					"UPDATE Cuenta SET Tipo = " + idTipo.ToString().Trim() +
					" WHERE (Tipo IS NULL OR Tipo = 0) AND Codigo Like '" 
					+ stCodigo.Trim() + "%'"
					, (SqlConnection) miConeccion.DbConnection);
				cmdActualiza.Transaction = (SqlTransaction)
					miConeccion.DbTransaction;
				cmdActualiza.ExecuteNonQuery();
				ds.Schema.Connections[0].Close();
			}
		}

		/// <summary>
		/// Crea la numeracion del asiento
		/// </summary>
		/// <param name="dt"></param>
		/// <param name="IdTipo"></param>
		static public string NumeroAsiento(C1.Data.C1DataSet ds, int IdTipo)
		{
			
			string stCodigo = "";
			bool bAuto = false;
			int iNumero = 0;

			string stSelect = "";
			SqlDataReader dr = null;
			try
			{
				stSelect =	"SELECT Automatico, Numero, Codigo FROM AsientoTipo "
					+ "WHERE idTipoAsiento = " + IdTipo.ToString();
				dr = Funcion.rEscalarSQL(ds, stSelect, true);
				dr.Read();
				bAuto = dr.GetBoolean(0);
				iNumero = dr.GetInt32(1);
				stCodigo = dr.GetString(2);
				dr.Close();
			}
			catch(System.Exception ex)
			{
				if (dr != null) dr.Close();
				MessageBox.Show(ex.Message, "No se puede leer numeracion");
				return "Err.SN";
			}

			if (!bAuto) return "";

			stCodigo += iNumero.ToString().Trim();
			iNumero ++;

			stSelect = "UPDATE AsientoTipo SET Numero = " + iNumero.ToString().Trim() + 
				" WHERE idTipoAsiento = " + IdTipo.ToString();
			try
			{
				Funcion.EjecutaSQL(ds, stSelect, true);
			}
			catch(System.Exception ex)
			{
				MessageBox.Show(ex.Message, "No se puede aumentar numeracion");
				return "Err";
			}
			return stCodigo;
		}

		/// <summary>
		/// Devuelve IdCuenta del plan de cuentas de la cuenta corriente
		/// </summary>
		/// <param name="ds"></param>
		/// <param name="IdCuentaCorr"></param>
		/// <returns></returns>
		static public int CuentaCuentaCorr(C1.Data.C1DataSet ds, int IdCuentaCorr)
		{
			ds.Schema.Connections[0].Open();
			C1.Data.SchemaObjects.Connection miConeccion = ds.Schema.Connections[0];
			SqlCommand cmdAuto = new SqlCommand (
				"SELECT idCuenta FROM CuentaCorriente WHERE idCuentaCorriente = " +
				IdCuentaCorr.ToString(), (SqlConnection) miConeccion.DbConnection);
			int iNumCuenta = 0;
			try
			{
				iNumCuenta = (int) cmdAuto.ExecuteScalar();
			}
			catch(System.Exception ex)
			{
				Console.WriteLine(ex.Message + ". Funcion Lee CuentaCuentaCorr");
			}
			return iNumCuenta;
		}
		
		/// <summary>
		/// Crea la numeracion de cheque de acuerdo a cuenta corriente
		/// </summary>
		/// <param name="dt"></param>
		/// <param name="IdTipo"></param>
		static public int NumeroCheque(C1.Data.C1DataSet ds, int IdCuentaCorr, bool bGrabar)
		{
			ds.Schema.Connections[0].Open();
			C1.Data.SchemaObjects.Connection miConeccion = ds.Schema.Connections[0];
			string stSelect = "SELECT Automatico FROM CuentaCorriente WHERE idCuentaCorriente = " +
				IdCuentaCorr.ToString();
			bool bAuto = false;
			SqlCommand cmdAuto = new SqlCommand (
				stSelect, (SqlConnection) miConeccion.DbConnection);
			try
			{
				bAuto = (bool) cmdAuto.ExecuteScalar();				
			}
			catch{}
			if (!bAuto) 
			{
				ds.Schema.Connections[0].Close();
				return -1;
			}
			cmdAuto = new SqlCommand (
				"SELECT NumCheque FROM CuentaCorriente WHERE idCuentaCorriente = " +
				IdCuentaCorr.ToString(), (SqlConnection) miConeccion.DbConnection);
			int iNumCheque = 0;
			try
			{
				iNumCheque = (int) cmdAuto.ExecuteScalar();
			}
			catch(System.Exception ex)
			{
				Console.WriteLine(ex.Message + ". Funcion Lee NumeraCheque");
			}

			if (bGrabar)
			{
				string stUpdate = "UPDATE CuentaCorriente SET NumCheque = " 
					+ iNumCheque.ToString().Trim() + " + 1 WHERE idCuentaCorriente = " +
					IdCuentaCorr.ToString();
				SqlCommand cmdActualiza = new SqlCommand (stUpdate, (SqlConnection) miConeccion.DbConnection);
				try
				{
					cmdActualiza.Transaction = (SqlTransaction)
						miConeccion.DbTransaction;
					cmdActualiza.ExecuteNonQuery();
				}
				catch(System.Exception ex)
				{
					Console.WriteLine(ex.Message + ". Funcion Grabar NumeraCheque");
				}
			}

			ds.Schema.Connections[0].Close();
			return iNumCheque;
		}
		/// <summary>
		/// Retorna el numero de cuenta corriente
		/// </summary>
		/// <param name="dt"></param>
		/// <param name="IdTipo"></param>
		static public int NumeroCuentaCorr(C1.Data.C1DataSet ds, int IdCuenta)
		{
			ds.Schema.Connections[0].Open();
			C1.Data.SchemaObjects.Connection miConeccion = ds.Schema.Connections[0];
			string stSelect = "SELECT idCuentaCorriente FROM CuentaCorriente "
				+ " WHERE idCuenta = " +	IdCuenta.ToString();
			SqlCommand cmdAuto = new SqlCommand (stSelect, (SqlConnection) miConeccion.DbConnection);
			int iNumCuenta = 0;
			try
			{
				iNumCuenta = (int) cmdAuto.ExecuteScalar();
			}
			catch(System.Exception ex)
			{
				Console.WriteLine(ex.Message + ". Funcion NumeraCuentaCorr");
			}

			ds.Schema.Connections[0].Close();
			return iNumCuenta;
		}

		/// <summary>
		/// Retorna el número de veces que existe la cuenta
		/// en los asientos
		/// </summary>
		/// <param name="ds"></param>
		/// <param name="IdTipo"></param>
		/// <returns></returns>
		static public int LineasAsiento(C1.Data.C1DataSet ds, int IdCuenta)
		{
			int iNumAsientos = 0;
			ds.Schema.Connections[0].Open();
			C1.Data.SchemaObjects.Connection miConeccion = ds.Schema.Connections[0];
			SqlCommand cmdAuto = new SqlCommand (
				"SELECT COUNT(*) FROM AsientoDetalle WHERE idCuenta = " +
				IdCuenta.ToString().Trim(), 
				(SqlConnection) miConeccion.DbConnection);
			try
			{
				iNumAsientos = (int) cmdAuto.ExecuteScalar();
			}
			catch(System.Exception ex)
			{
				Console.WriteLine(ex.Message + ". Funcion LineaAsiento");
			}
			ds.Schema.Connections[0].Close();
			return iNumAsientos;
		}

		/// <summary>
		/// Verifica que no exista código
		/// </summary>
		/// <param name="ds"></param>
		/// <param name="IdTipo"></param>
		/// <returns></returns>
		static public bool ExisteCodigo(C1.Data.C1DataSet ds, string stCodigo, int IdAsiento)
		{
			ds.Schema.Connections[0].Open();
			C1.Data.SchemaObjects.Connection miConeccion = ds.Schema.Connections[0];
			SqlCommand cmdAuto = new SqlCommand (
				"SELECT CodAsiento FROM Asiento WHERE CodAsiento = '" +
				stCodigo.Trim() + "' AND idAsiento <> " + IdAsiento.ToString().Trim()
				, (SqlConnection) miConeccion.DbConnection);
			string stCod = (string) cmdAuto.ExecuteScalar();
			ds.Schema.Connections[0].Close();
			if (stCod != null) return true;
			return false;
		}

		/// <summary>
		/// Verifica que no exista existan subcuentas de esta cuenta
		/// </summary>
		/// <returns></returns>
		static public bool ExisteSubCuenta(string stCodigo, DataTable dtCuenta)
		{
			if (stCodigo == "") return false;
			DataRow[] dr = dtCuenta.Select("Codigo like '" + stCodigo + "*'");
			if (dr.Length < 2)
				return false;
			else
				return true;
		}
		/// <summary>
		/// Verifica que no exista asientos creados con esta cuenta
		/// </summary>
		/// <returns></returns>
		static public bool MovimientoCuenta(int idCuenta, DataTable dtDetAsiento)
		{
			DataRow[] dr = dtDetAsiento.Select("IdCuenta = " + idCuenta.ToString());
			if (dr.Length == 0)
				return false;
			else
				return true;

		}
		/// <summary>
		/// Genera en la tabla cuentas la columna Padre para mostrar en el Arbol
		/// La columna CodRapido para usarlo como alias
		/// La columna grupo para filtrar las cuentas de grupo
		/// Tambien borra espacios en blanco de Codigo y Descripcion
		/// </summary>
		/// <param name="dsCuenta"></param>
		static public void GeneraArbol(C1DataSet dsCuenta)
		{
			string stCodActual = "";
			string stCodPadre = "";

			DataRow[] drs;
			DataTable t = dsCuenta.StorageDataSet.Tables["Cuenta"].Copy();
			foreach(C1DataRow dr in dsCuenta.Tables["Cuenta"].Rows)
			{
				// Borra espacios blancos de Codigo y Descripcion
				dr["Codigo"] = dr["Codigo"].ToString().Trim();
				dr["Descripcion"] = dr["Descripcion"].ToString().Trim();

				// Busqueda de codigo padre
				stCodActual = dr["codigo"].ToString();
				string stCodActual1=stCodActual;
				int iLong = stCodActual.Length - 1;
				if (stCodActual.ToString().EndsWith(".")) 
					stCodActual1 = stCodActual.Remove(iLong, 1);
				stCodPadre = Contabilidad.CuentaResumen(stCodActual1);
				drs = t.Select("Codigo = '" + stCodPadre + "'");

				// Cuando existe genera Padre del IdCuenta Anterior
				if (drs.Length > 0)	dr["Padre"] = drs[0]["IdCuenta"];

				// Asigna Cuentas que son de Grupo las que terminan con "."
				if (dr["Codigo"].ToString().EndsWith(".")) 
					dr["Grupo"] = true;
				else
					dr["Grupo"] = false;

				// Genera CodRapido
				dr["CodRapido"] = dr["Codigo"].ToString().Replace(".","");
			}
			dsCuenta.Update();
		}

		/// <summary>
		/// Genera el codigo de la cuenta que la contiene.
		/// Recibe como parametro el codigo completo de la cuenta por ejemplo 1.1.02
		/// Devuelve el codigo 1.1.
		/// </summary>C:\Documents and Settings\Roberto\Mis documentos\Proyectos de Visual Studio\Ciad2004\AssemblyInfo.cs
		static public string CuentaResumen(string cadena )
		{
			int PosPunto = 0;
			cadena = cadena.Trim();
			if (cadena.EndsWith(".")){
				return cadena+ "0";
//				cadena = cadena.Substring(0,cadena.Length-1);
			}
			PosPunto = cadena.LastIndexOf(".");
			if (PosPunto == 0) PosPunto = 1;
			return cadena.Substring(0, PosPunto+1);
		}
	}
}
