using System;
using System.Data.SqlClient;
using Latinium.Models.Versionamiento;

namespace Latinium.Services.Versionamiento
{
	public class VersionamientoService
	{
		private string _connectionString;

		public VersionamientoService(string connectionString)
		{
			_connectionString = connectionString;
		}

		public VersionamientoModel ObtenerVersionamiento()
		{
			VersionamientoModel modelo = new VersionamientoModel();

			try
			{
				// Consulta correcta: obtiene el valor del campo IntervaloVersionamiento
				string sql = "SELECT IntervaloVersionamiento FROM Seteo";

				SqlConnection cn = new SqlConnection(_connectionString);
				SqlCommand cmd = new SqlCommand(sql, cn);

				cn.Open();
				object result = cmd.ExecuteScalar();
				cn.Close();

				if (result != null && result != DBNull.Value)
				{
					int intervalo = Convert.ToInt32(result);

					if (intervalo > 0 && intervalo < 120)
					{
						modelo.Intervalo = intervalo;
					}
				}
			}
			catch
			{
				// Manejo silencioso: se mantiene el valor por defecto en el model (5)
			}

			return modelo;
		}
	}
}
