using System;

namespace Latinium.Models.OrdenDeCompra.Bodega
{
	public class DetalleBodegaModel
	{
		private int _bodega;
		private int _cantidad;

		public int Bodega
		{
			get { return _bodega; }
			set { _bodega = value; }
		}

		public int Cantidad
		{
			get { return _cantidad; }
			set { _cantidad = value; }
		}
	}
}