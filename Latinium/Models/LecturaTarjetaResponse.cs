using System;

namespace Latinium.Models
{
	public class LecturaTarjetaResponse
	{
		private string tipoMensaje;
		private string codigoRespuesta;
		private string numeroTarjetaTruncado;
		private string fechaVencimiento;
		private string numeroTarjetaEncriptado;
		private string mensajeRespuesta;
		private string codigoRedCorriente;
		private string codigoRedDiferido;
		private string filler;  // <-- Aquí lo agregamos

		public string TipoMensaje
		{
			get { return tipoMensaje; }
			set { tipoMensaje = value; }
		}

		public string CodigoRespuesta
		{
			get { return codigoRespuesta; }
			set { codigoRespuesta = value; }
		}

		public string NumeroTarjetaTruncado
		{
			get { return numeroTarjetaTruncado; }
			set { numeroTarjetaTruncado = value; }
		}

		public string FechaVencimiento
		{
			get { return fechaVencimiento; }
			set { fechaVencimiento = value; }
		}

		public string NumeroTarjetaEncriptado
		{
			get { return numeroTarjetaEncriptado; }
			set { numeroTarjetaEncriptado = value; }
		}

		public string MensajeRespuesta
		{
			get { return mensajeRespuesta; }
			set { mensajeRespuesta = value; }
		}

		public string CodigoRedCorriente
		{
			get { return codigoRedCorriente; }
			set { codigoRedCorriente = value; }
		}

		public string CodigoRedDiferido
		{
			get { return codigoRedDiferido; }
			set { codigoRedDiferido = value; }
		}

		public string Filler  // <-- Propiedad nueva
		{
			get { return filler; }
			set { filler = value; }
		}
	}
}
