using System;
using System.Collections;
using C1.Data;

namespace Latinium.Models.OrdenDeCompra.Bodega
{
	/// <summary>
	/// Request para manejar bodegas
	/// </summary>
	public class BodegaRequest
	{
		private int _cantidad;
		private int _idDetOrdenCompra;
		private bool _soloLectura;
		private string _articulo;
		private ArrayList _temporales;
		private C1DataSet _cdsSeteoF;

		public BodegaRequest()
		{
			_cantidad = 0;
			_idDetOrdenCompra = 0;
			_soloLectura = false;
			_articulo = "";
			_temporales = new ArrayList();
			_cdsSeteoF = null;
		}

		public int Cantidad
		{
			get { return _cantidad; }
			set { _cantidad = value; }
		}

		public int IdDetOrdenCompra
		{
			get { return _idDetOrdenCompra; }
			set { _idDetOrdenCompra = value; }
		}

		public bool SoloLectura
		{
			get { return _soloLectura; }
			set { _soloLectura = value; }
		}

		public string Articulo
		{
			get { return _articulo == null ? "" : _articulo; }
			set { _articulo = value; }
		}

		public ArrayList Temporales
		{
			get { return _temporales; }
			set { _temporales = value == null ? new ArrayList() : value; }
		}

		public C1DataSet CdsSeteoF
		{
			get { return _cdsSeteoF; }
			set { _cdsSeteoF = value; }
		}

		public bool EsValido()
		{
			if (_idDetOrdenCompra <= 0)
				return false;

			if (_cdsSeteoF == null)
				return false;

			return true;
		}
	}
}