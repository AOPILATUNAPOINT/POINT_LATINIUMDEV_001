using System;

namespace Latinium.Models
{
	/// <summary>
	/// Descripción breve de LecturaTarjetaRequest.
	/// </summary>
	public class LecturaTarjetaRequest
	{
		private string tipoMensaje;
		private decimal montoTotal;
		private string numeroCaja;
		private string numeroComercio;

		public LecturaTarjetaRequest()
		{
			tipoMensaje = "LT";  // Inicialización en constructor
		}

		public string TipoMensaje
		{
			get { return tipoMensaje; }
			set { tipoMensaje = value; }
		}

		public decimal MontoTotal
		{
			get { return montoTotal; }
			set { montoTotal = value; }
		}

		public string NumeroCaja
		{
			get { return numeroCaja; }
			set { numeroCaja = value; }
		}

		public string NumeroComercio
		{
			get { return numeroComercio; }
			set { numeroComercio = value; }
		}
	}
}
