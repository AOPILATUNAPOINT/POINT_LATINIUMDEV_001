using System;

namespace Latinium.Models.Pinpad
{
	public class ActualizaLecturaTarjetaRequest
	{
		// =========================
		// CAMPOS PRIVADOS
		// =========================
		public Guid idPinpad_LecturaConsultaTarjeta;

		public string TipoMensaje;
		public string CodigoRespuesta;
		public string RedCorriente;
		public string RedDiferido;
		public string NumeroTarjetaTruncado;
		public string FechaVencimiento;
		public string NumeroTarjetaEncriptado;
		public string MensajeRespuesta;
		public string Filler;
		public string TramaRecibida;
		public DateTime FechaRespuesta;

		// =========================
		
	}
}