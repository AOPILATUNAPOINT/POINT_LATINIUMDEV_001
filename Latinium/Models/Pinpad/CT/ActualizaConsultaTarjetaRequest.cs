using System;

namespace Latinium.Models.Pinpad
{
	public class ActualizaConsultaTarjetaRequest
	{
		// =========================
		// CAMPOS PRIVADOS
		// =========================
		public Guid idPinpad_LecturaConsultaTarjeta;

		public string TipoMensaje;
		public string CodigoRespuesta;
		public string NumeroTarjeta;
		public string Bin;
		public string FechaVencimiento;
		public string MensajeRespuesta;
		public string Filler;
		public string TramaRecibida;
		public DateTime FechaRespuesta;

		// =========================
		
	}
}