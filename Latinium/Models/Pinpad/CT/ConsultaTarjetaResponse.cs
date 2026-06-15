using System;

namespace Latinium.Models.Pinpad
{
	/// <summary>
	/// Descripción breve de ConsultaTarjetaResponse.
	/// </summary>
	public class ConsultaTarjetaResponse
	{
		public bool Success;
		public string DataAscii;
		public string ErrorMessage;
		public string TipoMensaje;
		public string CodigoRespuesta;
		public string CodigoRespuestaDescripcion;
		public string NumeroTarjeta;
		public string Bin;
		public string FechaVencimiento;
		public string MensajeRespuesta;
		public string Filler;
		public string TipoTarjetaDescripcion;
	}
}
