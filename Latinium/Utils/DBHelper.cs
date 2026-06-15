using System;
using System.Data;
using System.Data.SqlClient;

namespace Latinium.Utils
{
	public class DBHelper
	{
		#region CONEXIONES

		private static SqlConnection GetConnection()
		{
			return new SqlConnection(frmRequerimiento.sconexionPoint);
		}

		private static SqlConnection GetConnection(string connectionString)
		{
			if (connectionString == null || connectionString == "")
				return GetConnection();

			return new SqlConnection(connectionString);
		}

		#endregion

		#region ExecuteDataTable

		// .NET 1.1
		public static DataTable ExecuteQuery(string query, SqlParameter[] parameters)
		{
			return ExecuteQuery(query, parameters, null);
		}

		public static DataTable ExecuteQuery(string query, SqlParameter[] parameters, string connectionString)
		{
			DataTable dt = new DataTable();
			SqlConnection conn = null;
			SqlCommand cmd = null;
			SqlDataAdapter adapter = null;

			try
			{
				// Aquí usamos tu GetConnection
				if (connectionString == null)
					conn = GetConnection();          // conexión por defecto
				else
					conn = GetConnection(connectionString); // conexión personalizada

				cmd = new SqlCommand(query, conn);

				if (parameters != null)
				{
					for (int i = 0; i < parameters.Length; i++)
					{
						cmd.Parameters.Add(parameters[i]);
					}
				}

				adapter = new SqlDataAdapter(cmd);
				adapter.Fill(dt);
			}
			finally
			{
				if (adapter != null) adapter.Dispose();
				if (cmd != null) cmd.Dispose();
				if (conn != null) conn.Dispose();
			}

			return dt;
		}

		public static DataTable ExecuteDataTable(
			string query,
			SqlParameter[] parameters,
			string connectionString)
		{
			DataTable dt = new DataTable();
			SqlConnection conn = GetConnection(connectionString);
			SqlCommand cmd = new SqlCommand(query, conn);
			SqlDataAdapter da = new SqlDataAdapter(cmd);

			try
			{
				cmd.CommandType = CommandType.Text;

				if (parameters != null)
				{
					for (int i = 0; i < parameters.Length; i++)
						cmd.Parameters.Add(parameters[i]);
				}

				conn.Open();
				da.Fill(dt);
			}
			finally
			{
				if (conn.State == ConnectionState.Open)
					conn.Close();
			}

			return dt;
		}

		#endregion

		#region ExecuteNonQuery

		public static int ExecuteNonQuery(string query, SqlParameter[] parameters)
		{
			return ExecuteNonQuery(query, parameters, null);
		}

		public static int ExecuteNonQuery(string query, SqlParameter[] parameters, string connectionString)
		{
			SqlConnection conn = null;
			SqlCommand cmd = null;
			int rowsAffected = 0;

			try
			{
				// Obtener la conexión: por defecto o personalizada
				if (connectionString == null)
					conn = GetConnection();
				else
					conn = GetConnection(connectionString);

				cmd = new SqlCommand(query, conn);

				// Agregar parámetros uno por uno
				if (parameters != null)
				{
					for (int i = 0; i < parameters.Length; i++)
					{
						cmd.Parameters.Add(parameters[i]);
					}
				}

				conn.Open();
				rowsAffected = cmd.ExecuteNonQuery();
			}
			finally
			{
				// Liberación de recursos
				if (cmd != null) cmd.Dispose();
				if (conn != null) conn.Dispose();
			}

			return rowsAffected;
		}
		#endregion

		public static int ExecuteScalarInt(string query, SqlParameter[] parametros, string connectionString)
		{
			int resultado = 0;

			using (SqlConnection conn = new SqlConnection(connectionString))
			{
				using (SqlCommand cmd = new SqlCommand(query, conn))
				{
					cmd.CommandType = CommandType.Text;

					if (parametros != null)
					{
						for (int i = 0; i < parametros.Length; i++)
						{
							cmd.Parameters.Add(parametros[i]);
						}
					}

					conn.Open();

					object value = cmd.ExecuteScalar();

					if (value != null && value != DBNull.Value)
					{
						resultado = Convert.ToInt32(value);
					}
				}
			}

			return resultado;
		}
	}
}
