using System;
using System.Data;
using System.Data.SqlClient;
using System.IO;

namespace Latinium.Services.Transferencia
{
	public class ReporteTransferenciaService
	{
		private string _connectionString;

		public ReporteTransferenciaService(string connectionString)
		{
			_connectionString = connectionString;
		}

		// ============================
		// REPORTE PRINCIPAL
		// ============================
		public DataTable ObtenerReporte(DateTime desde, DateTime hasta, int estado, int bodega, int grupo, int articulo)
		{
			SqlConnection cn = new SqlConnection(_connectionString);
			SqlCommand cmd = new SqlCommand("ReporteDeTransferenciasDetalle", cn);
			DataTable dt = new DataTable();

			try
			{
				cmd.CommandType = CommandType.StoredProcedure;

				cmd.Parameters.Add("@Desde", SqlDbType.DateTime).Value = desde;
				cmd.Parameters.Add("@Hasta", SqlDbType.DateTime).Value = hasta;
				cmd.Parameters.Add("@Estado", SqlDbType.Int).Value = estado;
				cmd.Parameters.Add("@Bodega", SqlDbType.Int).Value = bodega;
				cmd.Parameters.Add("@Grupo", SqlDbType.Int).Value = grupo;
				cmd.Parameters.Add("@Articulo", SqlDbType.Int).Value = articulo;

				SqlDataAdapter da = new SqlDataAdapter(cmd);
				da.Fill(dt);
			}
			catch (Exception ex)
			{
				throw new Exception("Error al obtener reporte: " + ex.Message);
			}
			finally
			{
				if (cn.State == ConnectionState.Open)
					cn.Close();
			}

			return dt;
		}
		// ============================
		// BODEGAS
		// ============================
		public DataTable ObtenerBodegas()
		{
			SqlConnection cn = new SqlConnection(_connectionString);
			DataTable dt = new DataTable();

			try
			{
				string sql = "SELECT Bodega, Nombre FROM Bodega ORDER BY Nombre";

				SqlDataAdapter da = new SqlDataAdapter(sql, cn);
				da.Fill(dt);

				DataRow row = dt.NewRow();
				row["Bodega"] = 0;
				row["Nombre"] = "TODOS";
				dt.Rows.InsertAt(row, 0);
			}
			catch (Exception ex)
			{
				throw new Exception("Error al obtener bodegas: " + ex.Message);
			}
			finally
			{
				if (cn.State == ConnectionState.Open)
					cn.Close();
			}

			return dt;
		}

		// ============================
		// GRUPOS
		// ============================
		public DataTable ObtenerGrupos()
		{
			SqlConnection cn = new SqlConnection(_connectionString);
			DataTable dt = new DataTable();

			try
			{
				string sql = "SELECT idGrupoArticulo, Grupo FROM ArticuloGrupo WHERE Borrar = 0 ORDER BY Grupo";

				SqlDataAdapter da = new SqlDataAdapter(sql, cn);
				da.Fill(dt);

				DataRow row = dt.NewRow();
				row["idGrupoArticulo"] = 0;
				row["Grupo"] = "TODOS";
				dt.Rows.InsertAt(row, 0);
			}
			catch (Exception ex)
			{
				throw new Exception("Error al obtener grupos: " + ex.Message);
			}
			finally
			{
				if (cn.State == ConnectionState.Open)
					cn.Close();
			}

			return dt;
		}

		public DataTable ObtenerBodegasPorUsuario(int idUsuario, int idTipoFactura, DateTime fecha, bool recibeConsignacion, out int bodegaAsignada)
		{
			SqlConnection cn = new SqlConnection(_connectionString);
			DataTable dt = new DataTable();
			bodegaAsignada = 0;

			try
			{
				// 1. Traer bodegas del SP
				string sql = string.Format(
					"Exec FacturacionListaBodegas {0}, {1}, '{2}', {3}",
					idUsuario,
					idTipoFactura,
					fecha.ToString("yyyyMMdd"),
					recibeConsignacion ? 1 : 0
					);

				SqlDataAdapter da = new SqlDataAdapter(sql, cn);
				da.Fill(dt);

				// 2. Obtener bodega asignada
				cn.Open();

				string sqlAsignada = string.Format(
					"Exec BodegaAsignadaPorUsuarioBuscaFacturas {0}",
					idUsuario
					);

				SqlCommand cmd = new SqlCommand(sqlAsignada, cn);

				object result = cmd.ExecuteScalar();

				if (result != null && result != DBNull.Value)
					bodegaAsignada = Convert.ToInt32(result);

				// 3. AQUÍ VA TU LÓGICA
				if (bodegaAsignada == 0)
				{
					DataRow row = dt.NewRow();
					row["Bodega"] = 0;
					row["Nombre"] = "TODOS";
					dt.Rows.InsertAt(row, 0);
				}
			}
			catch (Exception ex)
			{
				throw new Exception("Error al obtener bodegas: " + ex.Message);
			}
			finally
			{
				if (cn.State == ConnectionState.Open)
					cn.Close();
			}

			return dt;
		}

		public void ExportarExcelHtml(DataTable dt, string ruta)
		{
			try
			{
				using (StreamWriter sw = new StreamWriter(ruta, false, System.Text.Encoding.UTF8))
				{
					// Cabecera HTML compatible con Excel
					sw.WriteLine("<html>");
					sw.WriteLine("<head>");
					sw.WriteLine("<meta http-equiv='Content-Type' content='text/html; charset=utf-8'/>");

					// Estilo Excel
					sw.WriteLine("<style>");
					sw.WriteLine(".text { mso-number-format:\"\\@\"; }"); // FORZAR TEXTO
					sw.WriteLine("</style>");

					sw.WriteLine("</head>");
					sw.WriteLine("<body>");
					sw.WriteLine("<table border='1'>");

					// ============================
					// ENCABEZADOS
					// ============================
					sw.WriteLine("<tr>");
					foreach (DataColumn col in dt.Columns)
					{
						sw.WriteLine("<th>" + col.ColumnName + "</th>");
					}
					sw.WriteLine("</tr>");

					// ============================
					// FILAS
					// ============================
					foreach (DataRow row in dt.Rows)
					{
						sw.WriteLine("<tr>");

						foreach (DataColumn col in dt.Columns)
						{
							string valor = row[col].ToString();

							// CAMPOS QUE DEBEN SER TEXTO
							if (col.ColumnName == "Serial" ||
								col.ColumnName == "CodigoTransferencia" ||
								col.ColumnName == "CodigoProducto")
							{
								sw.WriteLine("<td class='text'>" + valor + "</td>");
							}
							else
							{
								sw.WriteLine("<td>" + valor + "</td>");
							}
						}

						sw.WriteLine("</tr>");
					}

					sw.WriteLine("</table>");
					sw.WriteLine("</body>");
					sw.WriteLine("</html>");
				}
			}
			catch (Exception ex)
			{
				throw new Exception("Error al exportar Excel: " + ex.Message);
			}
		}
		// ============================
		// ARTICULOS
		// ============================
		public DataTable ObtenerArticulos()
		{
			SqlConnection cn = new SqlConnection(_connectionString);
			DataTable dt = new DataTable();

			try
			{
				string sql = "SELECT idArticulo, Articulo FROM Articulo WHERE Borrar = 0 ORDER BY Articulo";

				SqlDataAdapter da = new SqlDataAdapter(sql, cn);
				da.Fill(dt);

				DataRow row = dt.NewRow();
				row["idArticulo"] = 0;
				row["Articulo"] = "TODOS";
				dt.Rows.InsertAt(row, 0);
			}
			catch (Exception ex)
			{
				throw new Exception("Error al obtener artículos: " + ex.Message);
			}
			finally
			{
				if (cn.State == ConnectionState.Open)
					cn.Close();
			}

			return dt;
		}
	}
}