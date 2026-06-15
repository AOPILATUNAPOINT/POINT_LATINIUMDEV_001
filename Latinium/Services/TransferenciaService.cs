using System;
using System.IO;
using System.Net;
using Latinium.Models;

namespace Latinium.Services
{
	public class TransferenciaService
	{
		private string baseUrl;

		public TransferenciaService()
		{
			this.baseUrl = ConfigService.ObtenerValor("apiEcommerce");

			// Si no hay URL, solo logueamos
			if (this.baseUrl == null || this.baseUrl.Length == 0)
			{
				Log("No se encontró la URL base del servicio en WEB_GlobalConfig.", "WARNING");
			}
		}

		/// <summary>
		/// Ejecuta el procesamiento de la transferencia con log detallado
		/// </summary>
		public void EjecutarProcesamiento(string codigoTransf)
		{
			Log("Iniciando procesamiento de transferencia: " + codigoTransf, "INFO");

			try
			{
				ApiResponseModel resultado = ProcesarTransferencia(codigoTransf);

				if (resultado != null && resultado.Data)
				{
					Log("Transferencia procesada correctamente: " + codigoTransf, "SUCCESS");
				}
				else
				{
					Log("Error funcional en transferencia: " + codigoTransf +
						" -> " + (resultado != null ? resultado.Message : "Respuesta nula"), "ERROR");
				}
			}
			catch (Exception ex)
			{
				Log("Error técnico en procesamiento: " + ex.Message, "ERROR");
			}

			Log("Finalizado procesamiento de transferencia: " + codigoTransf, "INFO");
		}

		/// <summary>
		/// Llama al servicio externo para procesar la transferencia
		/// </summary>
		public ApiResponseModel ProcesarTransferencia(string codigoTransf)
		{
			ApiResponseModel result = new ApiResponseModel();

			if (this.baseUrl == null || this.baseUrl.Length == 0)
			{
				result.Status = 0;
				result.Message = "No hay URL configurada para el servicio.";
				result.Data = false;
				Log("No se puede procesar la transferencia " + codigoTransf + " porque no hay URL configurada.", "WARNING");
				return result;
			}

			try
			{
				string urlServicio = this.baseUrl + "transferencia/procesar-transferencia?codigoTransf=" + codigoTransf;
				Log("Llamando a API: " + urlServicio, "INFO");

				HttpWebRequest request = (HttpWebRequest)WebRequest.Create(urlServicio);
				request.Method = "GET";
				request.Timeout = 30000; // 30 segundos

				HttpWebResponse response = (HttpWebResponse)request.GetResponse();
				StreamReader reader = new StreamReader(response.GetResponseStream());
				string responseBody = reader.ReadToEnd();
				reader.Close();
				response.Close();

				Log("Respuesta recibida (resumida): " + 
					(responseBody.Length > 500 ? responseBody.Substring(0, 500) + "..." : responseBody), "INFO");

				if (response.StatusCode == HttpStatusCode.OK)
				{
					result.Status = GetJsonIntValue(responseBody, "Status");
					result.Message = GetJsonStringValue(responseBody, "Message");
					result.Data = GetJsonBoolValue(responseBody, "Data");

					Log("Resultado parseado: Status=" + result.Status + ", Data=" + result.Data + ", Message=" + result.Message, 
						result.Data ? "SUCCESS" : "ERROR");
				}
				else
				{
					result.Status = (int)response.StatusCode;
					result.Message = "Error en la respuesta del servicio";
					result.Data = false;

					Log("HTTP ERROR: " + result.Status, "ERROR");
				}
			}
			catch (Exception ex)
			{
				result.Status = 500;
				result.Message = "Error al llamar al servicio: " + ex.Message;
				result.Data = false;

				Log("EXCEPCIÓN: " + ex.Message, "ERROR");
			}

			return result;
		}

		/// <summary>
		/// Logging centralizado
		/// </summary>
		private void Log(string mensaje, string tipo)
		{
			try
			{
				string tipoLog = Funcion.slogtipoinformacion;

				if (tipo == "SUCCESS")
					tipoLog = Funcion.slogtipoexito;
				else if (tipo == "ERROR")
					tipoLog = Funcion.slogtipoerror;
				else if (tipo == "WARNING")
					tipoLog = Funcion.slogtipoadvertencia; // Si existe, si no usar info

				Funcion.LogSistema(
					frmRequerimiento.sconexionPoint,
					"TransferenciaService.cs",
					mensaje,
					"EJECUCIÓN API SERVICE DE ECOMMERCE",
					tipoLog
					);
			}
			catch
			{
				// No interrumpir si falla el log
			}
		}

		#region Parseo JSON simple
		private int GetJsonIntValue(string json, string key)
		{
			string pattern = "\"" + key + "\":";
			int start = json.IndexOf(pattern);
			if (start >= 0)
			{
				start += pattern.Length;
				int end = json.IndexOfAny(new char[] { ',', '}' }, start);
				string value = json.Substring(start, end - start).Trim();
				try { return int.Parse(value); } 
				catch { return 0; }
			}
			return 0;
		}

		private string GetJsonStringValue(string json, string key)
		{
			string pattern = "\"" + key + "\":\"";
			int start = json.IndexOf(pattern);
			if (start >= 0)
			{
				start += pattern.Length;
				int end = json.IndexOf("\"", start);
				if (end >= 0)
					return json.Substring(start, end - start);
			}
			return string.Empty;
		}

		private bool GetJsonBoolValue(string json, string key)
		{
			string pattern = "\"" + key + "\":";
			int start = json.IndexOf(pattern);
			if (start >= 0)
			{
				start += pattern.Length;
				int end = json.IndexOfAny(new char[] { ',', '}' }, start);
				string value = json.Substring(start, end - start).Trim().ToLower();
				return value == "true";
			}
			return false;
		}
		#endregion
	}
}
