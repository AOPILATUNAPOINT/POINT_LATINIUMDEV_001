using System;
using System.Data;
using System.Data.SqlClient;
using Latinium.Models.Requerimiento.TurnoSoporte;
using System.Collections;

namespace Latinium.Services.Requerimiento.TurnoSoporte
{
	public class RequerimientoTurnoSoporteService
	{
		private SqlConnection _cn;

		public RequerimientoTurnoSoporteService(SqlConnection conexion)
		{
			_cn = conexion;
		}
		// ================== OBTENER TODOS LOS TURNOS ==================
		public ArrayList ObtenerTodosTurnos(DateTime desde, DateTime hasta, string filtro)
		{
			// ===============================
			// LÓGICA DE FECHAS (SERVICE)
			// ===============================

			// Si no vienen fechas inicio del mes hasta hoy
			if (desde == DateTime.MinValue)
				desde = new DateTime(DateTime.Now.Year, DateTime.Now.Month, 1);

			if (hasta == DateTime.MinValue)
				hasta = DateTime.Now;

			// Normalización CORRECTA
			desde = desde.Date; // 00:00:00
			hasta = hasta.Date.AddDays(1).AddSeconds(-1); // 23:59:59

			if (_cn.State != ConnectionState.Open)
				_cn.Open();

			string sql = @"
        SELECT 
            t.idRequerimientoTurnoSoporte,
            t.idRequerimientoDepartamento,
            t.idRequerimientoPerTecnico,
            t.idRequerimientoTipo,
            t.FechaDesde,
            t.FechaHasta,
            t.Prioridad,
            t.Activo,
            d.Descripcion AS Departamento,
            tec.Nombre AS Tecnico,
            tt.Descripcion AS TipoTurno
        FROM RequerimientoTurnoSoporte t
        INNER JOIN RequerimientoDepartamento d 
            ON t.idRequerimientoDepartamento = d.idRequerimientoDepartamento
        INNER JOIN RequerimientoPerTecnico tec 
            ON t.idRequerimientoPerTecnico = tec.idRequerimientoPerTecnico
        INNER JOIN RequerimientoTipo tt 
            ON t.idRequerimientoTipo = tt.idRequerimientoTipo
        WHERE t.FechaDesde >= @FechaDesde
          AND t.FechaHasta <= @FechaHasta
          AND (@Filtro IS NULL 
               OR d.Descripcion LIKE '%' + @Filtro + '%' 
               OR tec.Nombre LIKE '%' + @Filtro + '%' 
               OR tt.Descripcion LIKE '%' + @Filtro + '%')
        ORDER BY t.FechaDesde DESC";

			SqlCommand cmd = new SqlCommand(sql, _cn);
			cmd.CommandType = CommandType.Text;

			// Parámetros FECHAS (tipados)
			cmd.Parameters.Add("@FechaDesde", SqlDbType.DateTime).Value = desde;
			cmd.Parameters.Add("@FechaHasta", SqlDbType.DateTime).Value = hasta;

			SqlParameter pFiltro = new SqlParameter("@Filtro", SqlDbType.VarChar, 100);
			if (filtro != null && filtro != "")
			{
				pFiltro.Value = filtro;
			}
			else
			{
				pFiltro.Value = DBNull.Value;
			}

			cmd.Parameters.Add(pFiltro);

			ArrayList lista = new ArrayList();
			SqlDataReader dr = null;

			try
			{
				dr = cmd.ExecuteReader();
				while (dr.Read())
				{
					RequerimientoTurnoSoporteModel m = new RequerimientoTurnoSoporteModel();

					m.idRequerimientoTurnoSoporte = Convert.ToInt32(dr["idRequerimientoTurnoSoporte"]);
					m.IdRequerimientoDepartamento = Convert.ToInt32(dr["idRequerimientoDepartamento"]);
					m.IdRequerimientoPerTecnico = Convert.ToInt32(dr["idRequerimientoPerTecnico"]);
					m.IdRequerimientoTipo = Convert.ToInt32(dr["idRequerimientoTipo"]);
					m.FechaDesde = Convert.ToDateTime(dr["FechaDesde"]);
					m.FechaHasta = Convert.ToDateTime(dr["FechaHasta"]);
					m.Prioridad = Convert.ToInt32(dr["Prioridad"]);
					m.Activo = Convert.ToBoolean(dr["Activo"]);

					m.NombreDepartamento = dr["Departamento"].ToString();
					m.NombreTecnico = dr["Tecnico"].ToString();
					m.TipoTurnoDescripcion = dr["TipoTurno"].ToString();

					lista.Add(m);
				}
			}
			finally
			{
				if (dr != null) dr.Close();
				cmd.Dispose();
				if (_cn.State == ConnectionState.Open)
					_cn.Close();
			}

			return lista;
		}

		// Sobrecarga sin parámetros
		public ArrayList ObtenerTodosTurnos()
		{
			return ObtenerTodosTurnos(DateTime.MinValue, DateTime.MinValue, "");
		}

		public void Guardar(RequerimientoTurnoSoporteModel t)
		{
			Validar(t);

			if (_cn.State != ConnectionState.Open)
				_cn.Open();

			using (SqlCommand cmd = new SqlCommand("RequerimientoTurnoSoporteInsUpd", _cn))
			{
				cmd.CommandType = CommandType.StoredProcedure;

				cmd.Parameters.Add("@idRequerimientoTurnoSoporte", SqlDbType.Int)
					.Value = t.idRequerimientoTurnoSoporte;

				cmd.Parameters.Add("@idRequerimientoPerTecnico", SqlDbType.Int)
					.Value = t.IdRequerimientoPerTecnico;

				cmd.Parameters.Add("@idRequerimientoDepartamento", SqlDbType.Int)
					.Value = t.IdRequerimientoDepartamento;

				cmd.Parameters.Add("@idRequerimientoTipo", SqlDbType.Int)
					.Value = t.IdRequerimientoTipo;

				cmd.Parameters.Add("@FechaDesde", SqlDbType.DateTime)
					.Value = t.FechaDesde;

				cmd.Parameters.Add("@FechaHasta", SqlDbType.DateTime)
					.Value = t.FechaHasta;

				cmd.Parameters.Add("@Prioridad", SqlDbType.Int)
					.Value = t.Prioridad;

				cmd.Parameters.Add("@Activo", SqlDbType.Bit)
					.Value = t.Activo;

				cmd.ExecuteNonQuery();
			}

			_cn.Close();
		}

		public RequerimientoTurnoSoporteModel ObtenerPorId(int id)
		{
			RequerimientoTurnoSoporteModel m = null;

			if (_cn.State != ConnectionState.Open)
				_cn.Open();

			using (SqlCommand cmd = new SqlCommand("RequerimientoTurnoSoporteCon", _cn))
			{
				cmd.CommandType = CommandType.StoredProcedure;
				cmd.Parameters.Add("@idRequerimientoTurnoSoporte", SqlDbType.Int)
					.Value = id;

				using (SqlDataReader dr =
								 cmd.ExecuteReader(CommandBehavior.CloseConnection))
				{
					if (dr.Read())
					{
						m = new RequerimientoTurnoSoporteModel();
						m.idRequerimientoTurnoSoporte = Convert.ToInt32(dr["idRequerimientoTurnoSoporte"]);
						m.IdRequerimientoPerTecnico =
							Convert.ToInt32(dr["idRequerimientoPerTecnico"]);
						m.IdRequerimientoDepartamento =
							Convert.ToInt32(dr["idRequerimientoDepartamento"]);
						m.IdRequerimientoTipo =
							Convert.ToInt32(dr["idRequerimientoTipo"]);
						m.FechaDesde =
							Convert.ToDateTime(dr["FechaDesde"]);
						m.FechaHasta =
							Convert.ToDateTime(dr["FechaHasta"]);
						m.Prioridad =
							Convert.ToInt32(dr["Prioridad"]);
						m.Activo =
							Convert.ToBoolean(dr["Activo"]);
					}
				}
			}

			return m;
		}

		// ================== COMBOS ==================

		public DataTable ObtenerDepartamentos()
		{
			string sql =
				"SELECT idRequerimientoDepartamento, Descripcion " +
				"FROM RequerimientoDepartamento " +
				"WHERE Estado = 1 ORDER BY Descripcion";

			return EjecutarConsulta(sql);
		}

		public DataTable ObtenerTecnicosPorDepartamento(int idDepartamento)
		{
			string sql =
				"SELECT DISTINCT t.idRequerimientoPerTecnico, t.Nombre " +
				"FROM RequerimientoPerTecnico t " +
				"INNER JOIN RequerimientoUsuarioDepartamento ud " +
				" ON t.idUsuario = ud.idUsuario " +
				"WHERE ud.idRequerimientoDepartamento = @idDepartamento " +
				"AND t.Estado = 1 AND t.HaceSoporte = 1 " +
				"ORDER BY t.Nombre";

			SqlDataAdapter da = new SqlDataAdapter(sql, _cn);
			da.SelectCommand.Parameters.Add("@idDepartamento", SqlDbType.Int)
				.Value = idDepartamento;

			DataTable dt = new DataTable();
			da.Fill(dt);
			return dt;
		}


		public DataTable ObtenerTiposPorDepartamento(int idDepartamento)
		{
			string sql =
				"SELECT idRequerimientoTipo, Descripcion " +
				"FROM RequerimientoTipo " +
				"WHERE Estado = 1 " +
				"AND idRequerimientoDepartamento = @idDepartamento " +
				"ORDER BY Descripcion";

			SqlDataAdapter da = new SqlDataAdapter(sql, _cn);
			da.SelectCommand.Parameters.Add("@idDepartamento", SqlDbType.Int)
				.Value = idDepartamento;

			DataTable dt = new DataTable();
			da.Fill(dt);
			return dt;
		}


		// ================== VALIDACIÓN ==================

		private void Validar(RequerimientoTurnoSoporteModel m)
		{

			if (m.IdRequerimientoDepartamento <= 0)
				throw new ApplicationException("Departamento obligatorio");

			if (m.IdRequerimientoPerTecnico <= 0)
				throw new ApplicationException("Técnico obligatorio");

			if (m.IdRequerimientoTipo <= 0)
				throw new ApplicationException("Tipo obligatorio");

			if (m.FechaDesde > m.FechaHasta)
				throw new ApplicationException("Fecha Desde no puede ser mayor que Hasta");
		}

		// ================== HELPERS ==================

		private DataTable EjecutarConsulta(string sql)
		{
			SqlDataAdapter da = new SqlDataAdapter(sql, _cn);
			DataTable dt = new DataTable();
			da.Fill(dt);
			return dt;
		}

		public void InhabilitarTurno(int idTurno)
		{
			try
			{
				// Abrir conexión si no está abierta
				if (_cn.State != ConnectionState.Open)
					_cn.Open();

				// Verificar si el turno existe y está activo
				SqlCommand cmdCheck = new SqlCommand(
					"SELECT COUNT(*) FROM RequerimientoTurnoSoporte WHERE idRequerimientoTurnoSoporte = @id AND Activo = 1",
					_cn);
				SqlParameter pCheck = new SqlParameter("@id", SqlDbType.Int);
				pCheck.Value = idTurno;
				cmdCheck.Parameters.Add(pCheck);

				int existe = (int)cmdCheck.ExecuteScalar();
				cmdCheck.Dispose();

				if (existe == 0)
				{
					// No existe o ya está inhabilitado, salir
					return;
				}

				// Inhabilitar el turno
				SqlCommand cmd = new SqlCommand(
					"UPDATE RequerimientoTurnoSoporte SET Activo = 0 WHERE idRequerimientoTurnoSoporte = @id",
					_cn);
				SqlParameter pUpdate = new SqlParameter("@id", SqlDbType.Int);
				pUpdate.Value = idTurno;
				cmd.Parameters.Add(pUpdate);

				cmd.ExecuteNonQuery();
				cmd.Dispose();
			}
			finally
			{
				if (_cn.State == ConnectionState.Open)
					_cn.Close();
			}
		}

	}
}
