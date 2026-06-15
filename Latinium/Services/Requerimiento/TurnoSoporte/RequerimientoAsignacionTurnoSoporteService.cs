using System;
using System.Data;
using System.Data.SqlClient;
using Latinium.Models.Requerimiento;

namespace Latinium.Services.Requerimiento
{
	public class RequerimientoAsignacionTurnoSoporteService
	{
		private SqlConnection _cn;

		public RequerimientoAsignacionTurnoSoporteService(SqlConnection cn)
		{
			_cn = cn;
		}

		public int ObtenerResponsableAsignado(
			int idResponsable,
			int idTipoRequerimiento,
			int idRequerimientoDepartamento
			)
		{
			if (_cn.State != ConnectionState.Open)
				_cn.Open();

			SqlCommand cmd = new SqlCommand(
				"RequerimientoAsignacionTurnoSoporte",
				_cn
				);

			cmd.CommandType = CommandType.StoredProcedure;

			cmd.Parameters.Add("@idResponsable", SqlDbType.Int).Value = idResponsable;
			cmd.Parameters.Add("@idTipoRequerimiento", SqlDbType.Int).Value = idTipoRequerimiento;
			cmd.Parameters.Add("@idRequerimientoDepartamento", SqlDbType.Int).Value = idRequerimientoDepartamento;

			int responsableFinal = idResponsable; // fallback seguro
			SqlDataReader dr = null;

			try
			{
				dr = cmd.ExecuteReader();
				if (dr.Read())
				{
					responsableFinal = Convert.ToInt32(dr["idResponsableAsignado"]);
				}
			}
			finally
			{
				if (dr != null) dr.Close();
				cmd.Dispose();
				if (_cn.State == ConnectionState.Open)
					_cn.Close();
			}

			return responsableFinal;
		}

		public DataTable ListarTecnicos(int idDepartamento, int idTipoRequerimiento)
		{
			DataTable dt = new DataTable();

			string query = @"
        SELECT DISTINCT
            r.idPersonal,
            r.Nombre
        FROM RequerimientoPerTecnico r
        WHERE r.idRequerimientoDepartamento = @idDepartamento
          AND r.idRequerimientoRol = 2
          AND r.Estado = 1
          AND r.HaceSoporte = 1
        ORDER BY r.Nombre
    ";

			if (_cn.State != ConnectionState.Open)
				_cn.Open();

			SqlCommand cmd = new SqlCommand(query, _cn);
			cmd.CommandType = CommandType.Text;
			cmd.Parameters.Add("@idDepartamento", SqlDbType.Int).Value = idDepartamento;

			SqlDataAdapter da = new SqlDataAdapter(cmd);
			da.Fill(dt);

			cmd.Dispose();

			if (_cn.State == ConnectionState.Open)
				_cn.Close();

			return dt;
		}


	}
}
