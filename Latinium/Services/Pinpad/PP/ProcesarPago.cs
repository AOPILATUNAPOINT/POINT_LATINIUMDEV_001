using System;
using System.IO;
using System.Net;
using System.Text;

namespace Latinium.Services.Pinpad.PP
{
	/// <summary>
	/// Cliente para consumir el servicio PinPad API desde ERP (.NET 1.1)
	/// </summary>
	public class ProcesarPago
	{
		private string _urlApi;

		public ProcesarPago(string urlApi)
		{
			_urlApi = urlApi;
		}

		/// <summary>
		/// Ejecuta el pago contra el servicio PinPad
		/// </summary>
		public string Ejecutar(decimal monto, decimal baseIVA, decimal base0, decimal iva, decimal propina,
			string tipoTransaccion, string usuario, string numeroCaja, int idCompra)
		{
			try
			{
				string json = ConstruirJson(monto, baseIVA, base0, iva, propina,
					tipoTransaccion, usuario, numeroCaja, idCompra);

				byte[] data = Encoding.UTF8.GetBytes(json);

				HttpWebRequest request = (HttpWebRequest)WebRequest.Create(_urlApi);
				request.Method = "POST";
				request.ContentType = "application/json";
				request.ContentLength = data.Length;
				request.Timeout = 60000; // 60 segundos

				using (Stream stream = request.GetRequestStream())
				{
					stream.Write(data, 0, data.Length);
				}

				using (HttpWebResponse response = (HttpWebResponse)request.GetResponse())
				using (StreamReader reader = new StreamReader(response.GetResponseStream()))
				{
					string result = reader.ReadToEnd();
					return result;
				}
			}
			catch (WebException ex)
			{
				if (ex.Response != null)
				{
					using (StreamReader reader = new StreamReader(ex.Response.GetResponseStream()))
					{
						string error = reader.ReadToEnd();
						return "Error API PinPad: " + error;
					}
				}

				return "Error de conexión con PinPad: " + ex.Message;
			}
			catch (Exception ex)
			{
				return "Error inesperado: " + ex.Message;
			}
		}

		public string EjecutarLT(decimal monto, string deviceId)
		{
			try
			{
				// JSON MANUAL compatible .NET 1.1
				string json = "{"
					+ "\"Monto\":" + monto.ToString(System.Globalization.CultureInfo.InvariantCulture) + ","
					+ "\"DeviceId\":\"" + deviceId + "\""
					+ "}";

				byte[] data = Encoding.UTF8.GetBytes(json);

				HttpWebRequest request = (HttpWebRequest)WebRequest.Create(_urlApi);
				request.Method = "POST";
				request.ContentType = "application/json";
				request.ContentLength = data.Length;
				request.Timeout = 60000;

				using (Stream stream = request.GetRequestStream())
				{
					stream.Write(data, 0, data.Length);
				}

				using (HttpWebResponse response = (HttpWebResponse)request.GetResponse())
				using (StreamReader reader = new StreamReader(response.GetResponseStream()))
				{
					return reader.ReadToEnd();
				}
			}
			catch (WebException ex)
			{
				if (ex.Response != null)
				{
					using (StreamReader reader = new StreamReader(ex.Response.GetResponseStream()))
					{
						return "Error API PinPad: " + reader.ReadToEnd();
					}
				}

				return "Error conexión: " + ex.Message;
			}
			catch (Exception ex)
			{
				return "Error: " + ex.Message;
			}
		}
		/// <summary>
		/// Construye el JSON manualmente (compatible con .NET 1.1)
		/// </summary>
		private string ConstruirJson(decimal monto, decimal baseIVA, decimal base0, decimal iva,
			decimal propina, string tipoTransaccion,
			string usuario, string numeroCaja, int idCompra)
		{
			return "{"
				+ "\"Monto\":" + monto.ToString(System.Globalization.CultureInfo.InvariantCulture) + ","
				+ "\"BaseIVA\":" + baseIVA.ToString(System.Globalization.CultureInfo.InvariantCulture) + ","
				+ "\"Base0\":" + base0.ToString(System.Globalization.CultureInfo.InvariantCulture) + ","
				+ "\"IVA\":" + iva.ToString(System.Globalization.CultureInfo.InvariantCulture) + ","
				+ "\"Propina\":" + propina.ToString(System.Globalization.CultureInfo.InvariantCulture) + ","
				+ "\"TipoTransaccion\":\"" + tipoTransaccion + "\","
				+ "\"Usuario\":\"" + usuario + "\","
				+ "\"NumeroCaja\":\"" + numeroCaja + "\","
				+ "\"IdCompra\":" + idCompra +
				"}";
		}
	}
}