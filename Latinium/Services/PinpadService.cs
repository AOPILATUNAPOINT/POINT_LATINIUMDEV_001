using System;
using Latinium.Models;
using Latinium.Utils;
using System.Net.Sockets;

namespace Latinium.Services
{
	public class PinpadService
	{
		private string ipPinpad = "192.168.0.125";  // IP del Pinpad
		private int puertoPinpad = 6500;             // Puerto TCP/IP del Pinpad
		private bool modoSimulado;
		private const string CODIGO_OK = "00";
		private const string CODIGO_ERROR_TRAMA = "01";
		private const string CODIGO_ERROR_CONEXION = "02";
		private const string CODIGO_ERROR_HASH = "03";
		private const string CODIGO_ERROR_GENERAL = "99";

		public bool ModoSimulado
		{
			get { return modoSimulado; }
			set { modoSimulado = value; }
		}

		public LecturaTarjetaResponse LeerTarjeta(LecturaTarjetaRequest request)
		{
			// Crear objeto de respuesta
			LecturaTarjetaResponse respuesta = new LecturaTarjetaResponse();
			try
			{

			decimal monto = request.MontoTotal;

			if (ModoSimulado)
			{
				respuesta.CodigoRespuesta = CODIGO_OK;
				respuesta.MensajeRespuesta = "AUTORIZADO - MODO SIMULADO";
				respuesta.CodigoRedCorriente = "2";
				respuesta.CodigoRedDiferido = "2";
				respuesta.NumeroTarjetaTruncado = "1234********5678";
				respuesta.FechaVencimiento = "1225";
				respuesta.NumeroTarjetaEncriptado = "ENCRIPTADO_SIMULADO";
				// Si tienes propiedad Filler en tu modelo
				respuesta.Filler = "";  
				return respuesta;  // Regresamos inmediatamente respuesta simulada
			}

			// Preparar datos para la trama
			string[] datos = new string[2];
			datos[0] = "LT"; // Tipo de mensaje Lectura Tarjeta

			// Aquí formateas monto a 12 dígitos con ceros a la izquierda
			datos[1] = monto.ToString("000000000000");

			// Construir trama con cifrado y seguridad
			byte[] tramaEnviar = ProcessData.getFinalData(datos);
			bool conectado = Conexion.Connectar(ipPinpad, puertoPinpad, tramaEnviar, 60000);

			if (!conectado)
			{
				respuesta.CodigoRespuesta = CODIGO_ERROR_CONEXION;
				respuesta.MensajeRespuesta = "Error de conexión al Pinpad";
				return respuesta;
			}

			// Intentar conectar y enviar trama
			if (conectado)
			{
				byte[] tramaRecibida = Conexion.getDataRecived();

				if (!CifradoTrama.EsTramaValida(tramaRecibida)) { 
					respuesta.CodigoRespuesta = CODIGO_ERROR_TRAMA;
					respuesta.MensajeRespuesta = "Error: trama recibida inválida o incompleta";
					return respuesta;
				}

				// Convertir trama a texto para procesar
				string respuestaHex = ProcessData.byte2hex(tramaRecibida);
				string respuestaAscii = ProcessData.hex2AsciiStr(respuestaHex);

				// Validar hash recibido
				string hashRecibido = respuestaAscii.Substring(respuestaAscii.Length - 32);

				if (!CifradoTrama.validateHash(hashRecibido))
				{
					respuesta.CodigoRespuesta = CODIGO_ERROR_HASH;
					respuesta.MensajeRespuesta = "Error: hash inválido";
					return respuesta;
				}

				// Aquí parseas la respuestaAscii para llenar los campos del objeto respuesta
				// Ejemplo ficticio, debes implementar según protocolo:
				respuesta.CodigoRespuesta = CODIGO_OK;
				respuesta.MensajeRespuesta = "Autorizado";

				// ... llena el resto de propiedades del objeto respuesta
			}
			else
			{
				respuesta.CodigoRespuesta = "Error conexión";
				respuesta.MensajeRespuesta = "No se pudo conectar al Pinpad";
			}

			return respuesta;
			}
			catch (Exception ex)
			{
				respuesta.CodigoRespuesta = CODIGO_ERROR_GENERAL;
				respuesta.MensajeRespuesta = "Error inesperado: " + ex.Message;
				return respuesta;
			}
		}

		public LecturaTarjetaResponse ProcesarRespuesta(string tramaRespuesta)
		{
			LecturaTarjetaResponse respuesta = new LecturaTarjetaResponse();

			if (tramaRespuesta == null || tramaRespuesta == "" || tramaRespuesta.Length < 40)
			{
				respuesta.CodigoRespuesta = CODIGO_ERROR_TRAMA; // Error en trama
				respuesta.MensajeRespuesta = "Trama de respuesta inválida";
				return respuesta;
			}

			try
			{
				// Extraer campos de la trama (posición y longitud según protocolo)
				respuesta.TipoMensaje = tramaRespuesta.Substring(0, 2); // "LT"
				respuesta.CodigoRespuesta = tramaRespuesta.Substring(2, 2);
				respuesta.CodigoRedCorriente = tramaRespuesta.Substring(4, 1);
				respuesta.CodigoRedDiferido = tramaRespuesta.Substring(5, 1);
				respuesta.NumeroTarjetaTruncado = tramaRespuesta.Substring(6, 25).Trim();
				respuesta.FechaVencimiento = tramaRespuesta.Substring(31, 4);
				respuesta.NumeroTarjetaEncriptado = tramaRespuesta.Substring(35, 64).Trim();
				respuesta.MensajeRespuesta = tramaRespuesta.Substring(99, 20).Trim();
				// Ignorar filler por ahora

				return respuesta;
			}
			catch (Exception ex)
			{
				respuesta.CodigoRespuesta = CODIGO_ERROR_CONEXION; // Error procesamiento
				respuesta.MensajeRespuesta = "Error procesando respuesta: " + ex.Message;
				return respuesta;
			}
		}

		public LecturaTarjetaResponse EnviarSolicitudLectura(LecturaTarjetaRequest request)
		{
			// Aquí se debería construir la trama según protocolo (no mostrado completo)
			string tramaSolicitud = ConstruirTramaSolicitud(request);

			byte[] datosEnviar = System.Text.Encoding.ASCII.GetBytes(tramaSolicitud);

			try
			{
				using (TcpClient client = new TcpClient())
				{
					client.Connect(ipPinpad, puertoPinpad);
					NetworkStream stream = client.GetStream();

					// Enviar datos
					stream.Write(datosEnviar, 0, datosEnviar.Length);

					// Buffer para recibir respuesta
					byte[] buffer = new byte[256];
					int bytesLeidos = stream.Read(buffer, 0, buffer.Length);

					// Convertir bytes recibidos a string
					string respuestaStr = System.Text.Encoding.ASCII.GetString(buffer, 0, bytesLeidos);

					// Procesar la respuesta recibida (usar método anterior)
					return ProcesarRespuesta(respuestaStr);
				}
			}
			catch (Exception ex)
			{
				LecturaTarjetaResponse response = new LecturaTarjetaResponse();
				response.CodigoRespuesta = CODIGO_ERROR_CONEXION;
				response.MensajeRespuesta = "Error de conexión o comunicación: " + ex.Message;
				return response;
			}
		}

		private string ConstruirTramaSolicitud(LecturaTarjetaRequest request)
		{
			// Formatear los campos según protocolo
			string tipoMensaje = request.TipoMensaje.PadLeft(2, '0'); // ej. "LT"
			string montoFormateado = ((int)(request.MontoTotal * 100)).ToString().PadLeft(12, '0'); // 123.45 => "000000012345"
			string numeroCaja = request.NumeroCaja.PadLeft(4, '0');         // ej. "0001"
			string numeroComercio = request.NumeroComercio.PadLeft(4, '0'); // ej. "0010"

			// Construir mensaje base sin ETX ni LRC
			string mensaje = tipoMensaje + montoFormateado + numeroCaja + numeroComercio;

			// Agregar ETX (código ASCII 0x03)
			char etx = (char)0x03;
			mensaje += etx;

			// Calcular LRC (XOR de todos los caracteres desde el inicio hasta ETX)
			byte lrc = CalcularLRC(mensaje);
			mensaje += (char)lrc;

			return mensaje;
		}

		private byte CalcularLRC(string input)
		{
			byte lrc = 0;
			foreach (char c in input)
			{
				lrc ^= (byte)c;
			}
			return lrc;
		}


	}
}
