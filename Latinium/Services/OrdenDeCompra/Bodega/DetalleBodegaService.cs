//using System;
//using System.Collections;
//using Latinium.Models.OrdenDeCompra.Bodega;
//using C1.Data;
//using System.Data;
//using System.Data.SqlClient;
//
//namespace Latinium.Services.OrdenDeCompra.Bodega
//{
//	public class DetalleBodegaService
//	{
//		public DetalleBodegaService()
//		{
//		}
//
//		public bool ValidarDistribucion(
//			Infragistics.Win.UltraWinGrid.UltraGridRow row,
//			out string mensaje)
//		{
//			mensaje = "";
//
//			if (row.Tag == null)
//			{
//				mensaje = "No ha distribuido bodegas.";
//				return false;
//			}
//
//			ArrayList lista = row.Tag as ArrayList;
//			if (lista == null || lista.Count == 0)
//			{
//				mensaje = "No ha distribuido bodegas.";
//				return false;
//			}
//
//			int totalBodegas = 0;
//
//			foreach (DetalleBodegaModel item in lista)
//			{
//				totalBodegas += item.Cantidad;
//			}
//
//			int cantidadDetalle = 0;
//
//			if (row.Cells["Cantidad"].Value != DBNull.Value)
//				cantidadDetalle = Convert.ToInt32(row.Cells["Cantidad"].Value);
//
//			// OBTENER DATOS DEL ARTÍCULO
//			string codigo = "";
//			string nombre = "";
//
//			if (row.Cells.Exists("Codigo") && row.Cells["Codigo"].Value != DBNull.Value)
//				codigo = row.Cells["Codigo"].Value.ToString();
//
//			if (row.Cells.Exists("Articulo") && row.Cells["Articulo"].Value != DBNull.Value)
//				nombre = row.Cells["Articulo"].Value.ToString();
//
//			string articulo = codigo;
//			if (nombre.Length > 0)
//				articulo += " - " + nombre;
//
//			if (totalBodegas != cantidadDetalle)
//			{
//				mensaje = "El artículo [" + articulo + "] tiene error:\n\n" +
//					"Suma de bodegas: " + totalBodegas + "\n" +
//					"Cantidad: " + cantidadDetalle;
//
//				return false;
//			}
//
//			return true;
//		}
//
//		public DataTable ObtenerBodegas(C1.Data.C1DataSet cds, int idDetOrdenCompra)
//		{
//			DataTable dt = new DataTable();
//
//			dt.Columns.Add("Bodega", typeof(int));
//			dt.Columns.Add("Nombre", typeof(string));
//			dt.Columns.Add("Cantidad", typeof(int));
//
//			string sql = @"
//    SELECT 
//        B.Bodega,
//        B.Nombre,
//        ISNULL(D.Cantidad, 0) AS Cantidad
//    FROM Bodega B
//    LEFT JOIN OrdenDeCompraBodegaDetalle D
//        ON D.Bodega = B.Bodega
//       AND D.IdDetOrdenDeCompra = " + idDetOrdenCompra + @"
//       AND D.Activo = 1
//    ORDER BY B.Nombre";
//
//			System.Data.SqlClient.SqlDataReader dr =
//				Funcion.rEscalarSQL(cds, sql, true);
//
//			while (dr.Read())
//			{
//				DataRow row = dt.NewRow();
//
//				row["Bodega"] = dr["Bodega"];
//				row["Nombre"] = dr["Nombre"];
//				row["Cantidad"] = dr["Cantidad"];
//
//				dt.Rows.Add(row);
//			}
//
//			dr.Close();
//
//			return dt;
//		}
//
//		public void GuardarDesdeTag(
//			int idDetOrdenDeCompra,
//			Infragistics.Win.UltraWinGrid.UltraGridRow row,
//			System.Data.SqlClient.SqlCommand cmd)
//		{
//			if (row.Tag == null) return;
//
//			ArrayList lista = row.Tag as ArrayList;
//			if (lista == null) return;
//
//			foreach (DetalleBodegaModel item in lista)
//			{
//				if (item.Cantidad <= 0) continue;
//
//				string sql = "EXEC GuardaDetOrdenCompraBodega " +
//					idDetOrdenDeCompra + "," +
//					item.Bodega + "," +
//					item.Cantidad;
//
//				cmd.CommandText = sql;
//				cmd.ExecuteNonQuery();
//			}
//		}
//
//		public bool AnularBodegasPorOrden(C1DataSet cds, int idOrden)
//		{
//			try
//			{
//				string sql = @"
//				UPDATE b
//				SET 
//					b.Activo = 0,
//					b.FechaAnulacion = GETDATE(),
//					b.UsuarioAnulacion = SUSER_NAME(),
//					b.EstacionAnulacion = HOST_NAME()
//				FROM OrdenDeCompraBodegaDetalle b
//				INNER JOIN DetOrdenDeCompra d 
//					ON d.IdDetOrdenDeCompra = b.IdDetOrdenDeCompra
//				WHERE d.IdOrdenDeCompra = " + idOrden + @"
//					AND b.Activo = 1";
//
//				return Funcion.EjecutaSQL(cds, sql);
//			}
//			catch
//			{
//				return false;
//			}
//		}
//
//		public ArrayList TransformarBodegas(Hashtable data)
//		{
//			ArrayList lista = new ArrayList();
//
//			int idArticulo = 0;
//
//			if (data.ContainsKey("idArticulo"))
//				idArticulo = Convert.ToInt32(data["idArticulo"]);
//
//			foreach (DictionaryEntry item in data)
//			{
//				string key = item.Key.ToString();
//
//				if (key.StartsWith("BOD_"))
//				{
//					int idBodega = int.Parse(key.Replace("BOD_", ""));
//					int cantidad = Convert.ToInt32(item.Value);
//
//					if (cantidad > 0)
//					{
//						DetalleBodegaModel obj = new DetalleBodegaModel();
//						obj.idArticulo = idArticulo;
//						obj.Bodega = idBodega;
//						obj.Cantidad = cantidad;
//
//						lista.Add(obj);
//					}
//				}
//			}
//
//			return lista;
//		}
//	}
//}

using System;
using System.Collections;
using System.Data;
using System.Data.SqlClient;
using Latinium.Models.OrdenDeCompra.Bodega;
using Infragistics.Win.UltraWinGrid;
using C1.Data;

namespace Latinium.Services.OrdenDeCompra.Bodega
{
	public class DetalleBodegaService
	{
		public DetalleBodegaService()
		{
		}

		// NUEVO: Preparar datos para el formulario
		public BodegaResponse PrepararBodegas(BodegaRequest request)
		{
			BodegaResponse response = new BodegaResponse();

			DataTable bodegas = ObtenerBodegas(
				request.CdsSeteoF,
				request.IdDetOrdenCompra
				);

			// Mezclar datos temporales
			if (request.Temporales != null)
			{
				foreach (DetalleBodegaModel item in request.Temporales)
				{
					foreach (DataRow dr in bodegas.Rows)
					{
						if (dr["Bodega"] != DBNull.Value &&
							Convert.ToInt32(dr["Bodega"]) == item.Bodega)
						{
							dr["Cantidad"] = item.Cantidad;
						}
					}
				}
			}

			response.DataBodegas = bodegas;
			response.CantidadTotal = request.Cantidad;
			response.SoloLectura = request.SoloLectura;
			response.Articulo = request.Articulo;

			return response;
		}

		// NUEVO: Calcular total
		public int CalcularTotal(ArrayList lista)
		{
			int total = 0;

			if (lista != null)
			{
				foreach (DetalleBodegaModel item in lista)
				{
					total += item.Cantidad;
				}
			}

			return total;
		}

		// =========================
		// TU CÓDIGO ORIGINAL (CASI IGUAL)
		// =========================

		public bool ValidarDistribucion(UltraGridRow row, out string mensaje)
		{
			mensaje = "";

			if (row.Tag == null)
			{
				mensaje = "No ha distribuido bodegas.";
				return false;
			}

			ArrayList lista = row.Tag as ArrayList;
			if (lista == null || lista.Count == 0)
			{
				mensaje = "No ha distribuido bodegas.";
				return false;
			}

			int totalBodegas = CalcularTotal(lista);

			int cantidadDetalle = 0;

			if (row.Cells["Cantidad"].Value != DBNull.Value)
				cantidadDetalle = Convert.ToInt32(row.Cells["Cantidad"].Value);

			string codigo = "";
			string nombre = "";

			if (row.Cells.Exists("Codigo") && row.Cells["Codigo"].Value != DBNull.Value)
				codigo = row.Cells["Codigo"].Value.ToString();

			if (row.Cells.Exists("Articulo") && row.Cells["Articulo"].Value != DBNull.Value)
				nombre = row.Cells["Articulo"].Value.ToString();

			string articulo = codigo;
			if (nombre.Length > 0)
				articulo += " - " + nombre;

			if (totalBodegas != cantidadDetalle)
			{
				mensaje = "El artículo [" + articulo + "] tiene error:\n\n" +
					"Suma de bodegas: " + totalBodegas + "\n" +
					"Cantidad: " + cantidadDetalle;

				return false;
			}

			return true;
		}

//		public DataTable ObtenerBodegas(C1DataSet cds, int idDetOrdenCompra)
//		{
//			DataTable dt = new DataTable();
//
//			dt.Columns.Add("Bodega", typeof(int));
//			dt.Columns.Add("Nombre", typeof(string));
//			dt.Columns.Add("Cantidad", typeof(int));
//
//			string sql = @"
//            SELECT 
//                B.Bodega,
//                B.Nombre,
//                ISNULL(D.Cantidad, 0) AS Cantidad
//            FROM Bodega B
//            LEFT JOIN OrdenDeCompraBodegaDetalle D
//                ON D.Bodega = B.Bodega
//               AND D.IdDetOrdenDeCompra = " + idDetOrdenCompra + @"
//               AND D.Activo = 1
//            ORDER BY B.Nombre";
//
//			SqlDataReader dr = Funcion.rEscalarSQL(cds, sql, true);
//
//			while (dr.Read())
//			{
//				DataRow row = dt.NewRow();
//
//				row["Bodega"] = dr["Bodega"];
//				row["Nombre"] = dr["Nombre"];
//				row["Cantidad"] = dr["Cantidad"];
//
//				dt.Rows.Add(row);
//			}
//
//			dr.Close();
//
//			return dt;
//		}

		public DataTable ObtenerBodegas(C1DataSet cds, int idDetOrdenCompra)
		{
			DataTable dtFinal = new DataTable();

			dtFinal.Columns.Add("Bodega", typeof(int));
			dtFinal.Columns.Add("Nombre", typeof(string));
			dtFinal.Columns.Add("Cantidad", typeof(int));

			// 1. Traer bodegas filtradas
			string sSQL = "Exec ListaBodegas " + MenuLatinium.IdUsuario + ", 2";
			DataTable dtBodegas = FuncionesProcedimientos.dtGeneral(sSQL);

			// 2. Traer cantidades
			string sqlDetalle = "SELECT Bodega, ISNULL(Cantidad,0) AS Cantidad " +
				"FROM OrdenDeCompraBodegaDetalle " +
				"WHERE IdDetOrdenDeCompra = " + idDetOrdenCompra + " " +
				"AND Activo = 1";

			DataTable dtDetalle = FuncionesProcedimientos.dtGeneral(sqlDetalle);

			// 3. Recorrer bodegas
			for (int i = 0; i < dtBodegas.Rows.Count; i++)
			{
				DataRow b = dtBodegas.Rows[i];
				DataRow row = dtFinal.NewRow();

				row["Bodega"] = Convert.ToInt32(b["Bodega"]);
				row["Nombre"] = b["Nombre"].ToString();

				int cantidad = 0;

				// 4. Buscar manualmente (más compatible que Select)
				for (int j = 0; j < dtDetalle.Rows.Count; j++)
				{
					if (Convert.ToInt32(dtDetalle.Rows[j]["Bodega"]) == Convert.ToInt32(b["Bodega"]))
					{
						cantidad = Convert.ToInt32(dtDetalle.Rows[j]["Cantidad"]);
						break;
					}
				}

				row["Cantidad"] = cantidad;

				dtFinal.Rows.Add(row);
			}

			return dtFinal;
		}

		public void GuardarDesdeTag(int idDetOrdenDeCompra, UltraGridRow row, SqlCommand cmd)
		{
			if (row.Tag == null) return;

			ArrayList lista = row.Tag as ArrayList;
			if (lista == null) return;

			foreach (DetalleBodegaModel item in lista)
			{
				if (item.Cantidad <= 0) continue;

				string sql = "EXEC GuardaDetOrdenCompraBodega " +
					idDetOrdenDeCompra + "," +
					item.Bodega + "," +
					item.Cantidad;

				cmd.CommandText = sql;
				cmd.ExecuteNonQuery();
			}
		}

		public bool AnularBodegasPorOrden(C1DataSet cds, int idOrden)
		{
			try
			{
				string sql = @"
                UPDATE b
                SET 
                    b.FechaAnulacion = GETDATE(),
                    b.UsuarioAnulacion = SUSER_NAME(),
                    b.EstacionAnulacion = HOST_NAME()
                FROM OrdenDeCompraBodegaDetalle b
                INNER JOIN DetOrdenDeCompra d 
                    ON d.IdDetOrdenDeCompra = b.IdDetOrdenDeCompra
                WHERE d.IdOrdenDeCompra = " + idOrden + @"
                    AND b.Activo = 1";
				return Funcion.EjecutaSQL(cds, sql);
			}
			catch
			{
				return false;
			}
		}

		public ArrayList TransformarBodegas(Hashtable data)
		{
			ArrayList lista = new ArrayList();

			foreach (DictionaryEntry item in data)
			{
				string key = item.Key.ToString();

				if (key.StartsWith("BOD_"))
				{
					int idBodega = int.Parse(key.Replace("BOD_", ""));
					int cantidad = Convert.ToInt32(item.Value);

					if (cantidad > 0)
					{
						DetalleBodegaModel obj = new DetalleBodegaModel();
						obj.Bodega = idBodega;
						obj.Cantidad = cantidad;

						lista.Add(obj);
					}
				}
			}

			return lista;
		}
	}
}