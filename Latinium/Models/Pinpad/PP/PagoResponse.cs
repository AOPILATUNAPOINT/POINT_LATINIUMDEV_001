using System;

namespace Latinium.Models.Pinpad
{
	/// <summary>
	/// Descripción breve de PagoResponse.
	/// </summary>
	// Clase de respuesta para el proceso de pago en .NET 1.1
	public class PagoResponse
	{
		private bool _success;
		private string _codigo;
		private string _mensaje;
		private string _errorMessage;

		// Constructor por defecto
		public PagoResponse()
		{
			_success = false;
			_codigo = "";
			_mensaje = "";
			_errorMessage = "";
		}

		// Propiedad Success
		public bool Success
		{
			get { return _success; }
			set { _success = value; }
		}

		// Propiedad Codigo
		public string Codigo
		{
			get { return _codigo; }
			set { _codigo = value; }
		}

		// Propiedad Mensaje
		public string Mensaje
		{
			get { return _mensaje; }
			set { _mensaje = value; }
		}

		// Propiedad ErrorMessage
		public string ErrorMessage
		{
			get { return _errorMessage; }
			set { _errorMessage = value; }
		}
	}
}
