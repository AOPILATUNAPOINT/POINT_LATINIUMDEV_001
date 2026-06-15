using System;

namespace Latinium.Models.Pinpad
{
	/// <summary>
	/// Respuesta de lectura de tarjeta (LT)
	/// </summary>
	public class LecturaTarjetaResponse
	{
		public bool Success;
		public string DataAscii;
		public string ErrorMessage;

		// Campos originales
		public string TipoMensaje;
		public string CodigoRespuesta;
		public string CodigoRespuestaDescripcion;
		public string RedCorriente;
		public string RedCorrienteDescripcion;
		public string RedDiferido;
		public string RedDiferidoDescripcion;

		// Nuevos campos (interpretados)
		public string NumeroTarjetaTruncado;
		public string FechaVencimiento;
		public string NumeroTarjetaEncriptado;
		public string MensajeRespuesta;
		public string Filler;

	}
}