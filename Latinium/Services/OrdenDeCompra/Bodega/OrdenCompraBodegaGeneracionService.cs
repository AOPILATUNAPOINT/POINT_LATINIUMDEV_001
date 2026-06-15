using System;
using System.Data;
using System.Collections;
using System.Data.SqlClient;

using Latinium.Models.OrdenDeCompra.Bodega;

namespace Latinium.Services.OrdenDeCompra.Bodega
{
	public class OrdenCompraBodegaGeneracionService
	{
		private C1.Data.C1DataSet cdsSeteoF;

		public OrdenCompraBodegaGeneracionService(
			C1.Data.C1DataSet _cds
			)
		{
			cdsSeteoF = _cds;
		}

		public DataTable Consultar(
			int idDetOrdenDeCompra
			)
		{
			SqlConnection oConexion =
				new SqlConnection(
				"Password=Bl45o6$9;Persist Security Info=True;User ID=" +
				MenuLatinium.stUsuario +
				";Initial Catalog=" +
				MenuLatinium.stDirFacturacion +
				";Data Source=" +
				MenuLatinium.stDirServidor +
				"; max pool size=1024"
				);

			oConexion.Open();

			try
			{
				return Consultar(
					idDetOrdenDeCompra,
					oConexion,
					null
					);
			}
			finally
			{
				oConexion.Close();
			}
		}

		public DataTable Consultar(
			int idDetOrdenDeCompra,
			SqlConnection oConexion,
			SqlTransaction oTransaction
			)
		{
			DataTable dt =
				new DataTable();

			SqlCommand oCmd =
				new SqlCommand();

			oCmd.Connection =
				oConexion;

			if (oTransaction != null)
			{
				oCmd.Transaction =
					oTransaction;
			}

			oCmd.CommandType =
				CommandType.Text;

			oCmd.CommandTimeout =
				120;

			oCmd.CommandText =
				"Exec ConsultaOrdenCompraGeneracionBodega " +
				idDetOrdenDeCompra;

			SqlDataAdapter da =
				new SqlDataAdapter();

			da.SelectCommand =
				oCmd;

			da.Fill(dt);

			return dt;
		}

		public DataTable ConsultarBodegaOrden(
			int IdOrdenDeCompra
			)
		{
			string sSQL =
				string.Format(
				"Exec ConsultaOrdenCompraGeneracionBodegaOrden {0}",
				IdOrdenDeCompra
				);

			return FuncionesProcedimientos.dtGeneral(
				sSQL
				);
		}

		public int ValidarCantidad(
			int valor,
			int pendiente
			)
		{
			if (valor < 0)
				return 0;

			if (valor > pendiente)
				return pendiente;

			return valor;
		}

		public DataTable ObtenerDetalleOrden(
			int idOrdenDeCompra,
			SqlConnection oConexion,
			SqlTransaction oTransaction
			)
		{
			DataTable dt = new DataTable();

			SqlCommand cmd = new SqlCommand();
			cmd.Connection = oConexion;

			if (oTransaction != null)
				cmd.Transaction = oTransaction;

			cmd.CommandType = CommandType.Text;
			cmd.CommandTimeout = 120;

			cmd.CommandText =
				"SELECT IdDetOrdenDeCompra, IdArticulo, Costo, Notas " +
				"FROM DetOrdenDeCompra " +
				"WHERE IdOrdenDeCompra = " + idOrdenDeCompra;

			SqlDataAdapter da = new SqlDataAdapter(cmd);
			da.Fill(dt);

			return dt;
		}

		public void GuardarTemporalMasivo(
			DataTable dt
			)
		{
			foreach (DataRow row in dt.Rows)
			{
				int id =
					Convert.ToInt32(
					row["idOrdenDeCompraBodegaDetalle"]
					);

				int cantidad = 0;

				if (row["CantidadTemporal"] != DBNull.Value)
				{
					cantidad =
						Convert.ToInt32(
						row["CantidadTemporal"]
						);
				}

				string sSQL =
					string.Format(
					"UPDATE OrdenDeCompraBodegaDetalle " +
					"SET CantidadTemporal = {0} " +
					"WHERE idOrdenDeCompraBodegaDetalle = {1}",
					cantidad,
					id
					);

				Funcion.EjecutaSQL(
					cdsSeteoF,
					sSQL
					);
			}
		}
	}
}