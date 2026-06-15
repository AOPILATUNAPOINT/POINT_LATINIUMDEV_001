using System;
using System.Data;
using System.Data.SqlClient;
using System.Collections; // Para ArrayList
using Latinium.Models.Pinpad; // Para PinpadDiferido

namespace Latinium.Services.Pinpad
{
	/// <summary>
	/// Servicio para manejar diferidos del PINPAD
	/// </summary>
	public class PinpadDiferidoService
	{
		private string _connectionString;

		public PinpadDiferidoService(string connectionString)
		{
			_connectionString = connectionString;
		}

		// Obtener diferidos por red, incluyendo planes y meses de gracia
		public ArrayList ObtenerDiferidosPorRed(string idRed)
		{
			ArrayList diferidos = new ArrayList();

			string query = @"
			SELECT d.idPinpad_Diferidos, d.TipoDiferido, d.Descripcion, p.Plazo, p.Gracia
			FROM Pinpad_Diferidos d
			LEFT JOIN Pinpad_DiferidoPlanes p
					ON d.idPinpad_Redes = p.idPinpad_Redes 
					AND d.idPinpad_Diferidos = p.idPinpad_Diferidos
			WHERE d.idPinpad_Redes = @idRed
				AND d.Activo = 1
				AND (p.Activo = 1 OR p.Activo IS NULL)
			ORDER BY d.idPinpad_Diferidos, p.Plazo";

			using (SqlConnection cn = new SqlConnection(_connectionString))
			using (SqlCommand cmd = new SqlCommand(query, cn))
			{
				cmd.Parameters.Add("@idRed", SqlDbType.Char, 2).Value = idRed.PadLeft(2, '0');

				cn.Open();
				SqlDataReader dr = cmd.ExecuteReader();
				while (dr.Read())
				{
					PinpadDiferido d = new PinpadDiferido();
					d.IdDiferido = dr["idPinpad_Diferidos"].ToString();
					d.IdRed = idRed;
					d.TipoDiferido = dr["TipoDiferido"].ToString();
					d.Descripcion = dr["Descripcion"].ToString();

					// Si hay planes
					d.Plazo = dr["Plazo"] != DBNull.Value ? Convert.ToInt32(dr["Plazo"]) : 0;
					d.Gracia = dr["Gracia"] != DBNull.Value ? Convert.ToInt32(dr["Gracia"]) : 0;

					diferidos.Add(d);
				}
			}

			return diferidos;
		}
	}
}