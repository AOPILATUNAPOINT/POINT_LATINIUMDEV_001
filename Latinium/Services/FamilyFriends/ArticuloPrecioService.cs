using System;
using System.Data;
using System.Data.SqlClient;

namespace Latinium.Services.FamilyFriends
{
	public class ArticuloPrecioService
	{
		private string _cn;

		public ArticuloPrecioService(string cn)
		{
			_cn = cn;
		}

		// LISTAR PRECIOS
		public DataTable ObtenerPrecios()
		{
			SqlConnection conn = new SqlConnection(_cn);
			DataTable dt = new DataTable();

			try
			{
				conn.Open();

				string sql = @"
				SELECT 
					idArticulo,
					Codigo,
					Articulo,
					Precio8,
					Precio9
				FROM Articulo";

				SqlDataAdapter da = new SqlDataAdapter(sql, conn);
				da.Fill(dt);
			}
			catch (Exception)
			{
				throw;
			}
			finally
			{
				if (conn != null)
					conn.Close();
			}

			return dt;
		}

		// ACTUALIZAR POR ID
		public void ActualizarPrecio(int idArticulo, decimal precio8, decimal precio9)
		{
			SqlConnection conn = new SqlConnection(_cn);

			try
			{
				conn.Open();

				string sql = @"
				UPDATE Articulo
				SET Precio8 = @Precio8,
					Precio9 = @Precio9
				WHERE idArticulo = @idArticulo";

				SqlCommand cmd = new SqlCommand(sql, conn);

				cmd.Parameters.Add("@Precio8", SqlDbType.Decimal).Value = precio8;
				cmd.Parameters.Add("@Precio9", SqlDbType.Decimal).Value = precio9;
				cmd.Parameters.Add("@idArticulo", SqlDbType.Int).Value = idArticulo;

				cmd.ExecuteNonQuery();
			}
			catch (Exception)
			{
				throw;
			}
			finally
			{
				if (conn != null)
					conn.Close();
			}
		}

		// ACTUALIZAR POR CODIGO
		public void ActualizarPorCodigo(string codigo, decimal p8, decimal p9)
		{
			SqlConnection conn = new SqlConnection(_cn);

			try
			{
				conn.Open();

				string sql = @"
				UPDATE Articulo
				SET Precio8 = @Precio8,
					Precio9 = @Precio9
				WHERE Codigo = @Codigo";

				SqlCommand cmd = new SqlCommand(sql, conn);

				cmd.Parameters.Add("@Precio8", SqlDbType.Decimal).Value = p8;
				cmd.Parameters.Add("@Precio9", SqlDbType.Decimal).Value = p9;
				cmd.Parameters.Add("@Codigo", SqlDbType.VarChar).Value = codigo;

				cmd.ExecuteNonQuery();
			}
			catch (Exception)
			{
				throw;
			}
			finally
			{
				if (conn != null)
					conn.Close();
			}
		}
	}
}