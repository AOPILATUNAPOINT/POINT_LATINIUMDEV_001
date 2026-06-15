using System;
using System.Collections;
using System.Data;
using System.Data.SqlClient;
using Latinium.Models;
using Latinium.Utils;

namespace Latinium.Services
{
	public class RequerimientoDepartamentoRolService
	{
		public RequerimientoDepartamentoRolService()
		{
			// Constructor vacío, si no necesitas inyectar nada.
		}

		private string connectionString;

		public RequerimientoDepartamentoRolService(string connectionString)
		{
			this.connectionString = connectionString;
		}

		// Listar todos los grupos disponibles
		public ArrayList ObtenerTodosLosGrupos()
		{
			ArrayList grupos = new ArrayList();

			SqlConnection conn = new SqlConnection(connectionString);
			SqlCommand cmd = new SqlCommand("SELECT idRolesWeb, Nombre FROM RolesWeb", conn);

			try
			{
				conn.Open();
				SqlDataReader reader = cmd.ExecuteReader();

				while (reader.Read())
				{
					RolesWeb grupo = new RolesWeb();
					grupo.IdRolesWeb = Convert.ToInt32(reader["idRolesWeb"]);
					grupo.Nombre = reader["Nombre"].ToString();
					grupos.Add(grupo);
				}

				reader.Close();
			}
			finally
			{
				conn.Close();
			}

			return grupos;
		}

		public ArrayList ObtenerAsociacionesPorDepartamento(int idDepartamento)
		{
			ArrayList lista = new ArrayList();

			using (SqlConnection conn = new SqlConnection(connectionString))
			{
				string query = @"
			SELECT 
				r.idRequerimientoDepartamentoRol, 
				ISNULL(r.idRequerimientoDepartamento, @idDepartamento) AS idRequerimientoDepartamento,
				rw.IdRolesWeb AS idGrupo, 
				rw.Nombre AS NombreGrupo,
				ISNULL(r.Estado, 0) AS Estado
			FROM 
				RolesWeb rw
			LEFT JOIN 
				RequerimientoDepartamentoRol r 
				ON rw.IdRolesWeb = r.idGrupo AND r.idRequerimientoDepartamento = @idDepartamento";

				SqlCommand cmd = new SqlCommand(query, conn);
				cmd.Parameters.Add(new SqlParameter("@idDepartamento", SqlDbType.Int)).Value = idDepartamento;

				try
				{
					conn.Open();
					SqlDataReader reader = cmd.ExecuteReader();

					while (reader.Read())
					{
						RequerimientoDepartamentoRol item = new RequerimientoDepartamentoRol();
						if (reader["idRequerimientoDepartamentoRol"] != DBNull.Value)
							item.IdRequerimientoDepartamentoRol = Convert.ToInt32(reader["idRequerimientoDepartamentoRol"]);
						else
							item.IdRequerimientoDepartamentoRol = -1; // o 0, o algún valor especial que represente "nuevo"
						item.IdRequerimientoDepartamento = Convert.ToInt32(reader["idRequerimientoDepartamento"]);
						item.IdGrupo = Convert.ToInt32(reader["idGrupo"]);
						item.NombreGrupo = reader["NombreGrupo"].ToString();
						item.Estado = Convert.ToBoolean(reader["Estado"]);
						lista.Add(item);
					}
					reader.Close();
				}
				finally
				{
					conn.Close();
				}
			}
			return lista;
		}

		public void GuardarRoles(ArrayList listaRoles)
		{
			foreach (RequerimientoDepartamentoRol rol in listaRoles)
			{
				GuardarRol(rol);
			}
		}

		private void GuardarRol(RequerimientoDepartamentoRol rol)
		{
			using (SqlConnection conn = new SqlConnection(connectionString))
			{
				string query = "IF EXISTS (SELECT 1 FROM RequerimientoDepartamentoRol WHERE IdRequerimientoDepartamento = @IdRequerimientoDepartamento AND IdGrupo = @IdGrupo) " +
					"BEGIN " +
					"   UPDATE RequerimientoDepartamentoRol SET Estado = @Estado WHERE IdRequerimientoDepartamento = @IdRequerimientoDepartamento AND IdGrupo = @IdGrupo " +
					"END " +
					"ELSE " +
					"BEGIN " +
					"   INSERT INTO RequerimientoDepartamentoRol (IdRequerimientoDepartamento, IdGrupo, Estado) " +
					"   VALUES (@IdRequerimientoDepartamento, @IdGrupo, @Estado) " +
					"END";

				SqlCommand cmd = new SqlCommand(query, conn);
				cmd.Parameters.Add("@IdRequerimientoDepartamento", SqlDbType.Int).Value = rol.IdRequerimientoDepartamento;
				cmd.Parameters.Add("@IdGrupo", SqlDbType.Int).Value = rol.IdGrupo;
				cmd.Parameters.Add("@Estado", SqlDbType.Bit).Value = rol.Estado;

				conn.Open();
				cmd.ExecuteNonQuery();
			}
		}

		// Asignar grupo a un departamento
		public void AsignarGrupo(int idDepartamento, int idGrupo)
		{
			SqlConnection conn = new SqlConnection(connectionString);
			SqlCommand cmd = new SqlCommand(
				"INSERT INTO RequerimientoDepartamentoRol (idRequerimientoDepartamento, idGrupo, Estado) " +
				"VALUES (@idDepartamento, @idGrupo, 1)", conn);

			// --- Código corregido para .NET 1.1 ---
			SqlParameter paramDepto = new SqlParameter("@idDepartamento", SqlDbType.Int);
			paramDepto.Value = idDepartamento;
			cmd.Parameters.Add(paramDepto);

			SqlParameter paramGrupo = new SqlParameter("@idGrupo", SqlDbType.Int);
			paramGrupo.Value = idGrupo;
			cmd.Parameters.Add(paramGrupo);
			// --------------------------------------

			try
			{
				conn.Open();
				cmd.ExecuteNonQuery();
			}
			finally
			{
				conn.Close();
			}
		}

		// Cambiar estado de relación
		public void CambiarEstado(int idRelacion, bool nuevoEstado)
		{
			SqlConnection conn = new SqlConnection(connectionString);
			SqlCommand cmd = new SqlCommand(
				"UPDATE RequerimientoDepartamentoRol SET Estado = @estado WHERE idRequerimientoDepartamentoRol = @id", conn);

			// --- Código corregido para .NET 1.1 ---
			SqlParameter paramEstado = new SqlParameter("@estado", SqlDbType.Bit);
			paramEstado.Value = nuevoEstado;
			cmd.Parameters.Add(paramEstado);

			SqlParameter paramId = new SqlParameter("@id", SqlDbType.Int);
			paramId.Value = idRelacion;
			cmd.Parameters.Add(paramId);
			// --------------------------------------

			try
			{
				conn.Open();
				cmd.ExecuteNonQuery();
			}
			finally
			{
				conn.Close();
			}
		}

		// Eliminar relación
		public void EliminarRelacion(int idRelacion)
		{
			SqlConnection conn = new SqlConnection(connectionString);
			SqlCommand cmd = new SqlCommand("DELETE FROM RequerimientoDepartamentoRol WHERE idRequerimientoDepartamentoRol = @id", conn);

			// --- Código corregido para .NET 1.1 ---
			SqlParameter param = new SqlParameter("@id", SqlDbType.Int);
			param.Value = idRelacion;
			cmd.Parameters.Add(param);
			// --------------------------------------

			try
			{
				conn.Open();
				cmd.ExecuteNonQuery();
			}
			finally
			{
				conn.Close();
			}
		}

		public bool TieneAcceso(int idGrupo, int idRequerimientoDepartamento)
		{
			string query = @"SELECT 1 FROM RequerimientoDepartamentoRol 
                           WHERE idGrupo = @idGrupo 
                           AND idRequerimientoDepartamento = @idDepartamento 
                           AND Estado = 1";
            
			DataTable dt = null;
			SqlParameter[] parameters = null;
            
			try
			{
				// Creación de parámetros al estilo .NET 1.1
				parameters = new SqlParameter[2];
                
				parameters[0] = new SqlParameter("@idGrupo", SqlDbType.Int);
				parameters[0].Value = idGrupo;
                
				parameters[1] = new SqlParameter("@idDepartamento", SqlDbType.Int);
				parameters[1].Value = idRequerimientoDepartamento;

				dt = DBHelper.ExecuteQuery(query, parameters);

				return dt != null && dt.Rows.Count > 0;
			}
			catch (Exception ex)
			{
				// Puedes loggear el error aquí si es necesario
				throw new ApplicationException("Error al verificar acceso al departamento", ex);
			}
			finally
			{
				// Liberación explícita de recursos
				if (dt != null)
					dt.Dispose();
			}
		}

		// Método adicional para obtener todos los accesos de un grupo
		public DataTable ObtenerAccesosPorGrupo(int idGrupo)
		{
			string query = @"SELECT idRequerimientoDepartamento 
                           FROM RequerimientoDepartamentoRol 
                           WHERE idGrupo = @idGrupo AND Estado = 1";
            
			SqlParameter[] parameters = null;
            
			try
			{
				parameters = new SqlParameter[1];
				parameters[0] = new SqlParameter("@idGrupo", SqlDbType.Int);
				parameters[0].Value = idGrupo;

				return DBHelper.ExecuteQuery(query, parameters);
			}
			catch (Exception ex)
			{
				throw new ApplicationException("Error al obtener accesos por grupo", ex);
			}
		}

		public DataTable ObtenerTodosLosDepartamentos()
		{
			DataTable tabla = new DataTable();

			string query = "SELECT idRequerimientoDepartamento, Descripcion FROM RequerimientoDepartamento WHERE Estado = 1";

			try
			{
				using (SqlConnection conn = new SqlConnection(connectionString))
				{
					SqlDataAdapter da = new SqlDataAdapter(query, conn);
					da.Fill(tabla);
				}
			}
			catch (Exception ex)
			{
				System.Diagnostics.Debug.WriteLine("Error en ObtenerTodosLosDepartamentos: " + ex.Message);
			}

			return tabla;
		}

		public bool EsUsuarioPermitidoEnSoloJefes(int idUsuario)
		{
			SqlConnection conn = null;
			SqlCommand cmd = null;

			try
			{
				conn = new SqlConnection(this.connectionString);
				cmd = new SqlCommand("RequerimientoUsuarioSoloJefesVal", conn);
				cmd.CommandType = CommandType.StoredProcedure;

				// Parámetro idUsuario
				SqlParameter p1 = new SqlParameter();
				p1.ParameterName = "@idUsuario";
				p1.SqlDbType = SqlDbType.Int;
				p1.Value = idUsuario;
				cmd.Parameters.Add(p1);

				// Parámetro idEmpresa
				SqlParameter p2 = new SqlParameter();
				p2.ParameterName = "@idEmpresa";
				p2.SqlDbType = SqlDbType.Int;
				p2.Value = Convert.ToInt32(MenuLatinium.stIdDB);
				cmd.Parameters.Add(p2);

				conn.Open();
				object result = cmd.ExecuteScalar();

				return result != null;  // Si devuelve algo, el usuario tiene acceso
			}
			finally
			{
				if (cmd != null) cmd.Dispose();
				if (conn != null) conn.Close();
			}
		}



	}
}