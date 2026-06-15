using System;

namespace Latinium.Models.OrdenDeCompra.Bodega
{
	public class OrdenCompraBodegaGeneracionModel
	{
		public int IdOrdenDeCompraBodegaDetalle;
		public int IdDetOrdenDeCompra;

		public int Bodega;
		public string NombreBodega;

		public int IdArticulo;
		public string Articulo;

		public int CantidadAsignada;
		public int CantidadGenerada;

		public int SaldoPendiente;

		public int CantidadTemporal;

		public int EstadoProceso;

		public bool Sel;
	}
}