using System;
using System.Collections;
using System.Data;
using System.Data.SqlClient;
using Latinium.Models;
using Latinium.Models.FamilyFriends.Estado;

namespace Latinium.Services
{
	/// <summary>
	/// Servicio para consultar el reporte Family & Friends
	/// usando el procedimiento almacenado [dbo].[FamilyFriends_ReporteCon].
	/// Totalmente compatible con .NET Framework 1.1.
	/// </summary>
	public class FamilyFriendsReporteService
	{
		private string connectionString;

		/// <summary>
		/// Constructor: recibe la cadena de conexión a la base de datos POINT.
		/// </summary>
		public FamilyFriendsReporteService(string connectionString)
		{
			this.connectionString = connectionString;
		}

		/// <summary>
		/// Ejecuta el procedimiento almacenado [dbo].[FamilyFriends_ReporteCon]
		/// y devuelve una lista (ArrayList) de objetos FamilyFriendsReporteModel.
		/// </summary>
		public ArrayList ObtenerReporte(FamilyFriendsReporteFiltroModel param)
		{
			ArrayList lista = new ArrayList();

			SqlConnection conn = null;
			SqlCommand cmd = null;
			SqlDataReader reader = null;

			try
			{
				conn = new SqlConnection(connectionString);

				// SP NUEVO V5
				cmd = new SqlCommand(
					"FamilyFriends_ReporteCon_V6",
					conn);

				cmd.CommandType =
					CommandType.StoredProcedure;

				// =====================================
				// CABECERA
				// =====================================
				cmd.Parameters.Add(
					"@IdWEB_FF_Cabecera",
					SqlDbType.UniqueIdentifier
					).Value = param.IdCabecera;

				// =====================================
				// FILTRO
				// =====================================
				if (
					param.Filtro == null
					||
					param.Filtro.Trim() == ""
					)
				{
					cmd.Parameters.Add(
						"@Filtro",
						SqlDbType.NVarChar,
						200
						).Value = DBNull.Value;
				}
				else
				{
					cmd.Parameters.Add(
						"@Filtro",
						SqlDbType.NVarChar,
						200
						).Value = param.Filtro;
				}

				// =====================================
				// AREA
				// =====================================
				if (
					param.Area == null
					||
					param.Area.Trim() == ""
					)
				{
					cmd.Parameters.Add(
						"@Area",
						SqlDbType.NVarChar,
						50
						).Value = DBNull.Value;
				}
				else
				{
					cmd.Parameters.Add(
						"@Area",
						SqlDbType.NVarChar,
						50
						).Value = param.Area;
				}

				// =====================================
				// NUEVO PARAMETRO ESTADO
				// =====================================
				cmd.Parameters.Add(
					"@idEstado",
					SqlDbType.Int
					).Value = param.idEstado;

				conn.Open();

				reader = cmd.ExecuteReader();

				while (reader.Read())
				{
					FamilyFriendsReporteModel item =
						FamilyFriendsReporteModel.Load(reader);

					lista.Add(item);
				}
			}
			catch (Exception ex)
			{
				throw new Exception(
					"Error al ejecutar FamilyFriends_ReporteCon_V6: "
					+ ex.Message);
			}
			finally
			{
				if (reader != null)
					reader.Close();

				if (
					conn != null
					&&
					conn.State == ConnectionState.Open
					)
				{
					conn.Close();
				}

				if (cmd != null)
					cmd.Dispose();
			}

			return lista;
		}

		public ArrayList ObtenerReporteDetalle(FamilyFriendsReporteFiltroModel param)
		{
			ArrayList lista = new ArrayList();
			SqlConnection conn = null;
			SqlCommand cmd = null;
			SqlDataReader reader = null;

			try
			{
				conn = new SqlConnection(connectionString);

				// NUEVO SP (CORRECTO)
				cmd = new SqlCommand("WEB_FF_Invitado_Listar_V7", conn);
				cmd.CommandType = CommandType.StoredProcedure;

				// @IdWEB_FF_Cabecera (OBLIGATORIO)
				cmd.Parameters.Add("@IdWEB_FF_Cabecera", SqlDbType.UniqueIdentifier).Value =
					param.IdCabecera;

				// @Filtro
				if (param.Filtro == null || param.Filtro.Trim() == "")
					cmd.Parameters.Add("@Filtro", SqlDbType.NVarChar, 200).Value = DBNull.Value;
				else
					cmd.Parameters.Add("@Filtro", SqlDbType.NVarChar, 200).Value = param.Filtro;

				// @Area
				if (param.Area == null || param.Area.Trim() == "")
					cmd.Parameters.Add("@Area", SqlDbType.NVarChar, 50).Value = DBNull.Value;
				else
					cmd.Parameters.Add("@Area", SqlDbType.NVarChar, 50).Value = param.Area;

				// @idEstado
				cmd.Parameters.Add("@idEstado", SqlDbType.Int).Value =
					param.idEstado;

				conn.Open();
				reader = cmd.ExecuteReader();

				while (reader.Read())
				{
					// SIGUE USANDO TU MODELO ACTUAL
					FamilyFriendsInvitadoDetalleModel item =
						FamilyFriendsInvitadoDetalleModel.Load(reader);

					lista.Add(item);
				}
			}
			catch (Exception ex)
			{
				throw new Exception("Error al obtener el reporte Family & Friends (Invitados - Detalle): " + ex.Message);
			}
			finally
			{
				if (reader != null)
					reader.Close();

				if (conn != null && conn.State == ConnectionState.Open)
					conn.Close();

				if (cmd != null)
					cmd.Dispose();
			}

			return lista;
		}
		public ArrayList ObtenerEstados()
		{
			ArrayList lista = new ArrayList();

			SqlConnection conn = null;
			SqlCommand cmd = null;
			SqlDataReader reader = null;

			try
			{
				conn = new SqlConnection(connectionString);

				cmd = new SqlCommand(
					"SELECT idEstado, NombreEstado " +
					"FROM WEB_FF_Estado " +
					"WHERE Activo = 1 " +
					"ORDER BY idEstado",
					conn);

				cmd.CommandType = CommandType.Text;

				conn.Open();

				reader = cmd.ExecuteReader();

				while (reader.Read())
				{
					FamilyFriendsEstadCombooModel item =
						new FamilyFriendsEstadCombooModel();

					item.idEstado =
						Convert.ToInt32(reader["idEstado"]);

					item.NombreEstado =
						reader["NombreEstado"].ToString();

					lista.Add(item);
				}
			}
			catch (Exception ex)
			{
				throw new Exception(
					"Error al obtener estados Family & Friends: "
					+ ex.Message);
			}
			finally
			{
				if (reader != null)
					reader.Close();

				if (conn != null &&
					conn.State == ConnectionState.Open)
					conn.Close();

				if (cmd != null)
					cmd.Dispose();
			}

			return lista;
		}

		public ArrayList ObtenerCabeceras()
		{
			ArrayList lista = new ArrayList();
			SqlConnection conn = null;
			SqlCommand cmd = null;
			SqlDataReader reader = null;

			try
			{
				conn = new SqlConnection(connectionString);

				cmd = new SqlCommand("WEB_FF_Cabecera_Listar", conn);
				cmd.CommandType = CommandType.StoredProcedure;

				conn.Open();
				reader = cmd.ExecuteReader();

				while (reader.Read())
				{
					FamilyFriendsCabeceraModel item =
						FamilyFriendsCabeceraModel.Load(reader);

					lista.Add(item);
				}
			}
			catch (Exception ex)
			{
				throw new Exception("Error al obtener cabeceras Family & Friends: " + ex.Message);
			}
			finally
			{
				if (reader != null)
					reader.Close();

				if (conn != null && conn.State == ConnectionState.Open)
					conn.Close();

				if (cmd != null)
					cmd.Dispose();
			}

			return lista;
		}

	}
}
