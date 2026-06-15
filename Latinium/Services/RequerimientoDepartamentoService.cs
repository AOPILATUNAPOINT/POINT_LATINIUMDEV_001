using System;
using System.Collections; // Necesario para ArrayList en .NET 1.1
using System.Data;
using System.Data.SqlClient;
using Latinium.Models;

namespace Latinium.Services
{
	/// <summary>
	/// Proporciona servicios para la gestión de departamentos.
	/// </summary>
	public class RequerimientoDepartamentoService
	{
		private string connectionString;

		public RequerimientoDepartamentoService(string connectionString)
		{
			this.connectionString = connectionString;
		}

		// Listar todos los departamentos activos
		public ArrayList ListarDepartamentosActivos()
		{
			ArrayList listaDepartamentos = new ArrayList();

			SqlConnection conn = new SqlConnection(connectionString);
			SqlCommand cmd = new SqlCommand(
				"SELECT idRequerimientoDepartamento, Descripcion FROM RequerimientoDepartamento WHERE Estado = 1 ORDER BY Descripcion",
				conn
				);

			try
			{
				conn.Open();
				SqlDataReader reader = cmd.ExecuteReader();

				while (reader.Read())
				{
					RequerimientoDepartamento departamento = new RequerimientoDepartamento();
					departamento.IdRequerimientoDepartamento = Convert.ToInt32(reader["idRequerimientoDepartamento"]);
					departamento.Descripcion = reader["Descripcion"].ToString();
					listaDepartamentos.Add(departamento);
				}

				reader.Close();
			}
			finally
			{
				conn.Close();
			}

			return listaDepartamentos;
		}
	}
}