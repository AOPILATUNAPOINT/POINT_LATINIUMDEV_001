using System;
using System.Data;

namespace Latinium.Models.OrdenDeCompra.Bodega
{
	/// <summary>
	/// Respuesta de bodegas
	/// </summary>
	public class BodegaResponse
	{
		private int _cantidadTotal;
		private bool _soloLectura;
		private string _articulo;
		private DataTable _dataBodegas;

		public int CantidadTotal
		{
			get { return _cantidadTotal; }
			set { _cantidadTotal = value; }
		}

		public bool SoloLectura
		{
			get { return _soloLectura; }
			set { _soloLectura = value; }
		}

		public string Articulo
		{
			get { return _articulo; }
			set { _articulo = value; }
		}

		public DataTable DataBodegas
		{
			get { return _dataBodegas; }
			set { _dataBodegas = value; }
		}
	}
}