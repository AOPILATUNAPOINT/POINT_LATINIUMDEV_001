using System;

namespace Latinium.Models.Pinpad
{
	public class PinpadResponse
	{
		// =========================
		// BASE (respuesta servicio)
		// =========================
		public bool Success;
		public string DataAscii;
		public string ErrorMessage;

		public string Mensaje;
		public string Codigo;

		// =========================
		// DATOS DE TRAMA
		// =========================
		public string CodigoRedCorriente;
		public string CodigoRedDiferido;
		public string NumeroTarjetaTruncado;
		public string Bin;

		// =========================
		// INTERPRETADOS (DEL PARSER)
		// =========================
		public string RedCorrienteDescripcion;
		public string RedDiferidoDescripcion;

		public string TipoTarjetaDescripcion;
		public string TipoFinanciamiento;
		public string EstadoTransaccion;

		// =========================
		// DETALLE (OBJETO PARSER)
		// =========================
		public LtDetalle Detalle;
	}
}