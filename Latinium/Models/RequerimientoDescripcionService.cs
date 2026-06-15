using System;
using System.Collections;
using System.Data;
using System.Data.SqlClient;
using Latinium.Models;
using Latinium.Utils;

namespace Latinium.Services
{
	/// <summary>
	/// Descripción breve de RequerimientoService.
	/// </summary>
	public class RequerimientoDescripcionService
	{

		public RequerimientoDescripcionService()
		{
			//
			// TODO: agregar aquí la lógica del constructor
			//
		}

		public static ArrayList ObtenerRequerimientosPorTipo(int idTipo)
		{
			SqlParameter[] parametros = new SqlParameter[]
		{
			new SqlParameter("@idRequerimientoTipo", idTipo)
		};

			// Ahora llamas al nuevo método que sabe cómo ejecutar un procedimiento almacenado.
			DataTable dt = DBHelper.ExecuteStoredProcedure("RequerimientoDescripcionPorTipoCon", parametros);
    
			ArrayList lista = new ArrayList();

			foreach (DataRow row in dt.Rows)
			{
				RequerimientoDescripcion req = new RequerimientoDescripcion();
				req.IdRequerimientoDescripcion = Convert.ToInt32(row["idRequerimientoDescripcion"]);
				req.Descripcion = row["Descripcion"].ToString();
				lista.Add(req);
			}

			return lista;
		}
	}
}