using System;
using System.Data;
using System.Data.SqlClient;

namespace Latinium
{
	/// <summary>
	/// Descripción breve de FacturacionProductos.
	/// </summary>
	public class FacturacionProductos
	{
		public FacturacionProductos()
		{
			//
			// TODO: agregar aquí la lógica del constructor
			//
		}

		static public int iCuentaArticulo(string sArticulo, int idTipoFactura, int iBodega, C1.Data.C1DataSet ds)
		{									
			return Funcion.iEscalarSQL(ds, string.Format("Exec VentaArticuloCuenta '{0}', {1}, {2}, 0, 0", sArticulo, idTipoFactura, iBodega), true);
		}

		#region Tipo de Articulo
		static public int iPromocion(int idArticulo, int idTipoPrecio, int iBodega, int idBanco, int idTarjeta, int idPlan, int idPlazo, C1.Data.C1DataSet ds)
		{
			return Funcion.iEscalarSQL(ds, string.Format("Select dbo.ValidaPromocionEspecifica ({0}, {1}, {2}, {3}, {4}, {5}, {6})", 
				idArticulo, idTipoPrecio, iBodega, idBanco, idTarjeta, idPlan, idPlazo));		 
		}

		static public int iPromocionValor(int iBodega, int idArticulo, int idTipoPrecio, decimal dTotal, C1.Data.C1DataSet ds)
		{
			return Funcion.iEscalarSQL(ds, string.Format("Select dbo.ValidaPromocionPorValor ({0}, {1}, {2}, {3}, 0, 0, 0, 0)", 
				iBodega, idArticulo, idTipoPrecio , dTotal));
		}

		static public int iPromocionAdicional(int iBodega, int idArticulo, int idTipoPrecio, C1.Data.C1DataSet ds)
		{
			return Funcion.iEscalarSQL(ds, string.Format("Select dbo.ValidaPromocionAdicional ({0}, {1}, {2}, 0, 0, 0, 0)", 
				iBodega, idArticulo, idTipoPrecio));		 
		}

		static public int idCombo(int idArticulo, C1.Data.C1DataSet ds)
		{
			return Funcion.iEscalarSQL(ds, string.Format("Select Count(*) From Combo Where idArticuloCombo = {0} And Estado = 0", idArticulo));
		}

		static public int idComboSubGrupo(int idArticulo, C1.Data.C1DataSet ds)
		{
			return Funcion.iEscalarSQL(ds, string.Format("Exec ComboPorSubGrupo {0}", idArticulo));
		}
		#endregion Tipo de Articulo

		static public void ArticuloIndividual(Infragistics.Win.UltraWinGrid.UltraGrid uGrid, int iFila, int idArticulo, int idFormaPago, int idTipoPrecio, int iBodega, int idBanco, int idTarjeta, int idPlan, DateTime dtFecha)
		{
			SqlCommand oCmdActualiza = Funcion.Coneccion().CreateCommand();
			oCmdActualiza.Connection = Funcion.Coneccion();
			oCmdActualiza.CommandType = CommandType.Text;
			oCmdActualiza.CommandTimeout = 0;

			string sSQL = string.Format("Exec CargaInformacionArticulo {0}, {1}, {2}, {3}, {4}, {5}, {6}, '{7}'", 
				idArticulo, idFormaPago, idTipoPrecio, iBodega, idBanco, idTarjeta, idPlan, dtFecha.ToString("yyyyMMdd"));			
			oCmdActualiza.CommandText = sSQL;
			SqlDataAdapter oDADetalle = new SqlDataAdapter(oCmdActualiza);
			DataTable oDTDetalle = new DataTable();
			oDADetalle.Fill(oDTDetalle);
						
			CargaArticulos(uGrid, oDTDetalle, iFila);
		}


		static public void CargaArticulos(Infragistics.Win.UltraWinGrid.UltraGrid uGrid, DataTable oDT, int iFila)
		{
			
		}

		static public void Articulo_Individual()
		{
					
		}
	}
}
