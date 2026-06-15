using System;

namespace Latinium.Models
{
	/// <summary>
	/// Modelo para la respuesta de la API.
	/// </summary>
	public class ApiResponseModel
	{
		private int status;
		private string message;
		private bool data;

		public ApiResponseModel()
		{
			// Puedes inicializar valores por defecto si lo deseas
			status = 0;
			message = string.Empty;
			data = false;
		}

		public int Status
		{
			get { return status; }
			set { status = value; }
		}

		public string Message
		{
			get { return message; }
			set { message = value; }
		}

		public bool Data
		{
			get { return data; }
			set { data = value; }
		}
	}
}
