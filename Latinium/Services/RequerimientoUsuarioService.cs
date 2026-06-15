using System;
using System.Data;
using System.Data.SqlClient;
using Latinium.Utils;

namespace Latinium.Services
{
	/// <summary>
	/// Servicio para manejar operaciones relacionadas con RequerimientoUsuario
	/// </summary>
	public class RequerimientoUsuarioService
	{
		public RequerimientoUsuarioService()
		{
			// Constructor - inicialización si es necesaria
		}

		public int ObtenerGrupoPorUsuario(int idUsuario)
		{
			string query = "SELECT idGrupo FROM RequerimientoUsuario WHERE idUsuario = @idUsuario";
			DataTable dt = null;
			SqlParameter[] parameters = null;
            
			try
			{
				parameters = new SqlParameter[1];
				parameters[0] = new SqlParameter("@idUsuario", SqlDbType.Int);
				parameters[0].Value = idUsuario;

				dt = DBHelper.ExecuteQuery(query, parameters);

				if (dt != null && dt.Rows.Count > 0)
				{
					if (dt.Rows[0]["idGrupo"] != DBNull.Value)
					{
						return Convert.ToInt32(dt.Rows[0]["idGrupo"]);
					}
				}
				return -1; // Usuario no encontrado o grupo nulo
			}
			catch (Exception ex)
			{
				throw new ApplicationException("Error al obtener grupo por usuario", ex);
			}
			finally
			{
				if (dt != null)
					dt.Dispose();
			}
		}

		public DataTable ObtenerUsuario(int idUsuario)
		{
			string query = "SELECT * FROM RequerimientoUsuario WHERE idUsuario = @idUsuario";
			SqlParameter[] parameters = null;
            
			try
			{
				parameters = new SqlParameter[1];
				parameters[0] = new SqlParameter("@idUsuario", SqlDbType.Int);
				parameters[0].Value = idUsuario;

				return DBHelper.ExecuteQuery(query, parameters);
			}
			catch (Exception ex)
			{
				throw new ApplicationException("Error al obtener datos del usuario", ex);
			}
		}

		public bool ActualizarUsuario(int idUsuario, string nombreCompleto, int idGrupo, bool activo)
		{
			string query = @"UPDATE RequerimientoUsuario 
                             SET NombreCompleto = @NombreCompleto, 
                                 idGrupo = @idGrupo, 
                                 Activo = @Activo 
                             WHERE idUsuario = @idUsuario";
            
			SqlParameter[] parameters = new SqlParameter[4];
            
			try
			{
				parameters[0] = new SqlParameter("@NombreCompleto", SqlDbType.NVarChar, 150);
				parameters[0].Value = nombreCompleto;
                
				parameters[1] = new SqlParameter("@idGrupo", SqlDbType.Int);
				parameters[1].Value = idGrupo;
                
				parameters[2] = new SqlParameter("@Activo", SqlDbType.Bit);
				parameters[2].Value = activo;
                
				parameters[3] = new SqlParameter("@idUsuario", SqlDbType.Int);
				parameters[3].Value = idUsuario;

				int result = DBHelper.ExecuteNonQuery(query, parameters);
				return result > 0;
			}
			catch (Exception ex)
			{
				throw new ApplicationException("Error al actualizar usuario", ex);
			}
		}

		public bool InsertarUsuario(int idUsuario, string nombreCompleto, int idGrupo, bool activo)
		{
			string query = @"INSERT INTO RequerimientoUsuario 
                           (idUsuario, NombreCompleto, idGrupo, Activo) 
                           VALUES (@idUsuario, @NombreCompleto, @idGrupo, @Activo)";
            
			SqlParameter[] parameters = new SqlParameter[4];
            
			try
			{
				parameters[0] = new SqlParameter("@idUsuario", SqlDbType.Int);
				parameters[0].Value = idUsuario;
                
				parameters[1] = new SqlParameter("@NombreCompleto", SqlDbType.NVarChar, 150);
				parameters[1].Value = nombreCompleto;
                
				parameters[2] = new SqlParameter("@idGrupo", SqlDbType.Int);
				parameters[2].Value = idGrupo;
                
				parameters[3] = new SqlParameter("@Activo", SqlDbType.Bit);
				parameters[3].Value = activo;

				int result = DBHelper.ExecuteNonQuery(query, parameters);
				return result > 0;
			}
			catch (Exception ex)
			{
				throw new ApplicationException("Error al insertar usuario", ex);
			}
		}
	}
}