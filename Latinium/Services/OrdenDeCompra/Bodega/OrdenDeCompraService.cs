using System;
using System.Collections;
using C1.Data;
using System.Data;
using System.Data.SqlClient;

namespace Latinium.Services.OrdenDeCompra.Bodega
{
	public class OrdenDeCompraService
	{
		private C1DataSet cdsSeteoF;

		public OrdenDeCompraService(C1DataSet _cdsSeteoF)
		{
			cdsSeteoF = _cdsSeteoF;
		}

		/// <summary>
		/// Elimina uno o varios detalles de orden de compra
		/// </summary>
		public void EliminarDetalles(ArrayList idsDetCompra)
		{
			if (idsDetCompra == null || idsDetCompra.Count == 0)
				return;

			int i;
			for (i = 0; i < idsDetCompra.Count; i++)
			{
				int idDetCompra = Convert.ToInt32(idsDetCompra[i]);

				ValidarEliminacion(idDetCompra);
				EliminarBodegas(idDetCompra);
				EliminarDetalle(idDetCompra);
			}
		}

		public void ActualizarFacturacionDetalle(
			SqlConnection conn,
			SqlTransaction tx,
			int idDetCompra,
			string usuario)
		{
			string sql = "EXEC dbo.OrdenDeCompraActualizarFacturacion @IdDetCompra, @Usuario";

			using (SqlCommand cmd = new SqlCommand(sql, conn, tx))
			{
				cmd.Parameters.Add("@IdDetCompra", SqlDbType.Int).Value = idDetCompra;
				cmd.Parameters.Add("@Usuario", SqlDbType.VarChar, 50).Value = usuario;

				cmd.ExecuteNonQuery();
			}
		}

		/// <summary>
		/// Valida si se puede eliminar
		/// </summary>
		private void ValidarEliminacion(int idDetCompra)
		{
			string sValida =
				"SELECT CASE WHEN ISNULL(SaldoUIO,0)=0 AND ISNULL(SaldoGYE,0)=0 THEN 0 ELSE 1 END " +
				"FROM DetOrdenDeCompra WHERE IdDetOrdenDeCompra = " + idDetCompra;

			int okEliminar = Funcion.iEscalarSQL(cdsSeteoF, sValida, false);

			if (okEliminar != 0)
			{
				throw new Exception("No se puede eliminar el ítem. Existen órdenes de compra generadas o saldo pendiente.");
			}
		}

		public void GuardarDetalleOrdenCompra(
			SqlConnection conn,
			SqlTransaction tx,
			int idCompra,
			int idArticulo,
			int cantidad,
			decimal costo,
			decimal subtotal,
			int bodega,
			int idProyecto,
			decimal impuesto)
		{
			SqlCommand cmd = new SqlCommand(
				"Exec GuardaDetalleOrdenDeCompra " +
				"0," +              // @idDetCompra
				"@idCompra," +
				"@idArticulo," +
				"0," +              // @idUnidad
				"@cantidad," +
				"0," +              // @CantAereo
				"@costo," +
				"@impuesto," +      // @Impuesto
				"0," +              // @DescuentoArt
				"0," +              // @DescuentoPorc
				"@subtotal," +
				"@idProyecto," +    // @idProyecto
				"0," +              // @idImportacion
				"@bodega," +
				"0," +              // @Signo
				"0," +              // @Tipo
				"0," +              // @EstadoSerial
				"0," +              // @Courier
				"0," +              // @Peso
				"0," +              // @Arancel
				"0," +              // @idTipoCalculo
				"0," +              // @DescNC
				"0," +              // @Valores
				"0",                // @color
				conn,
				tx);

			// ======================
			// PARAMETROS
			// ======================

			cmd.Parameters.Add("@impuesto", SqlDbType.Decimal).Value = impuesto;
			SqlParameter pImpuesto = cmd.Parameters["@impuesto"];
			pImpuesto.Precision = 18;
			pImpuesto.Scale = 0;

			cmd.Parameters.Add("@idProyecto", SqlDbType.Int).Value = idProyecto;
			cmd.Parameters.Add("@idCompra", SqlDbType.Int).Value = idCompra;
			cmd.Parameters.Add("@idArticulo", SqlDbType.Int).Value = idArticulo;

			cmd.Parameters.Add("@cantidad", SqlDbType.Decimal).Value = cantidad;
			SqlParameter pCantidad = cmd.Parameters["@cantidad"];
			pCantidad.Precision = 18;
			pCantidad.Scale = 0;

			cmd.Parameters.Add("@costo", SqlDbType.Decimal).Value = costo;
			SqlParameter pCosto = cmd.Parameters["@costo"];
			pCosto.Precision = 18;
			pCosto.Scale = 5;

			cmd.Parameters.Add("@subtotal", SqlDbType.Decimal).Value = subtotal;
			SqlParameter pSubtotal = cmd.Parameters["@subtotal"];
			pSubtotal.Precision = 18;
			pSubtotal.Scale = 4;

			cmd.Parameters.Add("@bodega", SqlDbType.Int).Value = bodega;

			// ejecutar dentro de la transacción
			cmd.ExecuteNonQuery();
			cmd.Dispose();
		}

		/// <summary>
		/// Inicializa control de facturación en DetCompra
		/// </summary>
		public void InicializarControlFacturacion(
			SqlConnection conn,
			SqlTransaction tx,
			int idCompra,
			int idArticulo,
			int bodega,
			int cantidad)
		{
			SqlCommand cmd = new SqlCommand(
				@"UPDATE DetCompra
          SET 
              EstadoFacturacion = 0,
              CantidadFacturada = 0,
              CantidadPorFacturar = 0
          WHERE idCompra = @idCompra
            AND IdArticulo = @idArticulo
            AND Bodega = @bodega",
				conn,
				tx);

			cmd.Parameters.Add("@idCompra", SqlDbType.Int).Value = idCompra;
			cmd.Parameters.Add("@idArticulo", SqlDbType.Int).Value = idArticulo;
			cmd.Parameters.Add("@bodega", SqlDbType.Int).Value = bodega;

			cmd.ExecuteNonQuery();
			cmd.Dispose();
		}

		/// <summary>
		/// Elimina registros de bodegas (tabla hija)
		/// </summary>
		private void EliminarBodegas(int idDetCompra)
		{
			string sDelBodega =
				"DELETE FROM OrdenDeCompraBodegaDetalle WHERE IdDetOrdenDeCompra = " + idDetCompra;

			Funcion.EjecutaSQL(cdsSeteoF, sDelBodega);
		}

		/// <summary>
		/// Elimina el detalle principal (tabla padre)
		/// </summary>
		private void EliminarDetalle(int idDetCompra)
		{
			string sDel =
				"DELETE FROM DetOrdenDeCompra WHERE IdDetOrdenDeCompra = " + idDetCompra;

			Funcion.EjecutaSQL(cdsSeteoF, sDel);
		}

		/// <summary>
		/// ENVÍA UNA ORDEN A POR FACTURAR
		/// </summary>
		public void EnviarPorFacturar(
			int idCompra,
			string usuario)
		{
			SqlConnection conn = null;

			try
			{
				//---------------------------------------------------------
				// CONEXIÓN
				//---------------------------------------------------------
				conn =
					new SqlConnection(
					MenuLatinium.sconexionEmpresa);

				conn.Open();

				//---------------------------------------------------------
				// STORED PROCEDURE
				//---------------------------------------------------------
				SqlCommand cmd =
					new SqlCommand(
					"OrdenDeCompraEnviarPorFacturar",
					conn);

				cmd.CommandType =
					CommandType.StoredProcedure;

				//---------------------------------------------------------
				// PARÁMETROS
				//---------------------------------------------------------
				cmd.Parameters.Add(
					"@idCompra",
					SqlDbType.Int).Value = idCompra;

				cmd.Parameters.Add(
					"@Usuario",
					SqlDbType.VarChar).Value = usuario;

				//---------------------------------------------------------
				// TIMEOUT
				//---------------------------------------------------------
				cmd.CommandTimeout = 0;

				//---------------------------------------------------------
				// EJECUTA
				//---------------------------------------------------------
				cmd.ExecuteNonQuery();

				cmd.Dispose();
			}
			catch
			{
				throw;
			}
			finally
			{
				//---------------------------------------------------------
				// CERRAR CONEXIÓN
				//---------------------------------------------------------
				if (conn != null)
				{
					if (conn.State == ConnectionState.Open)
						conn.Close();

					conn.Dispose();
				}
			}
		}
		
		public void EnviarRecibidoBodega(
			int idCompra,
			string usuario)
		{
			SqlConnection conn = null;

			try
			{
				//---------------------------------------------------------
				// CONEXIÓN
				//---------------------------------------------------------
				conn =
					new SqlConnection(
					MenuLatinium.sconexionEmpresa);

				conn.Open();

				//---------------------------------------------------------
				// STORED PROCEDURE
				//---------------------------------------------------------
				SqlCommand cmd =
					new SqlCommand(
					"OrdenDeCompraEnviarRecibidoBodega",
					conn);

				cmd.CommandType =
					CommandType.StoredProcedure;

				//---------------------------------------------------------
				// PARÁMETROS
				//---------------------------------------------------------
				cmd.Parameters.Add(
					"@idCompra",
					SqlDbType.Int).Value = idCompra;

				cmd.Parameters.Add(
					"@Usuario",
					SqlDbType.VarChar).Value = usuario;

				//---------------------------------------------------------
				// TIMEOUT
				//---------------------------------------------------------
				cmd.CommandTimeout = 0;

				//---------------------------------------------------------
				// EJECUTA
				//---------------------------------------------------------
				cmd.ExecuteNonQuery();

				cmd.Dispose();
			}
			catch
			{
				throw;
			}
			finally
			{
				//---------------------------------------------------------
				// CERRAR CONEXIÓN
				//---------------------------------------------------------
				if (conn != null)
				{
					if (conn.State == ConnectionState.Open)
						conn.Close();

					conn.Dispose();
				}
			}
		}

	}
}