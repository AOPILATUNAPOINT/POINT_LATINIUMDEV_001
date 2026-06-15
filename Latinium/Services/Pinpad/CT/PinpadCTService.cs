using System;
using Latinium.Models.Pinpad;

namespace Latinium.Services.Pinpad
{
	/// <summary>
	/// Servicio de Consulta de Tarjeta (CT)
	/// </summary>
	public class PinpadCTService : PinpadBaseService
	{
		public ConsultaTarjetaResponse Procesar(ConsultaTarjetaRequest request)
		{
			string tallaHex;
			string trama = ArmarTrama(out tallaHex);
			Guid id = _repo.InsertarLectura(
				request.IdCompra,
				request.MontoTotal,
				"CT",
				trama,
				tallaHex
				);
			string json = _tcp.Enviar(trama);
			ConsultaTarjetaResponse resp = Parsear(json);
			if (!resp.Success)
				return resp;
			ProcesarRespuesta(id, resp);
			return resp;
		}

		private string ArmarTrama(out string tallaHex)
		{
			string cuerpo = "CT0000";
			tallaHex = CalcularTalla(cuerpo);
			string cuerpoFinal = cuerpo.Substring(0, 2) + tallaHex + cuerpo.Substring(6);
			return tallaHex + cuerpoFinal;
		}

		// ESTE ES EL MÉTODO QUE TE FALTABA
		private ConsultaTarjetaResponse Parsear(string json)
		{
			ConsultaTarjetaResponse resp = new ConsultaTarjetaResponse();
			if (json == null || json.Trim() == "")
			{
				resp.Success = false;
				resp.ErrorMessage = "Respuesta vacía";
				return resp;
			}
			resp.Success = ExtraerBool(json, "Success");
			resp.DataAscii = ExtraerValor(json, "DataAscii");
			resp.ErrorMessage = ExtraerValor(json, "ErrorMessage");

			resp.TipoMensaje = ExtraerValor(json, "TipoMensaje");
			resp.CodigoRespuesta = ExtraerValor(json, "CodigoRespuesta");
			resp.CodigoRespuestaDescripcion = ExtraerValor(json, "CodigoRespuestaDescripcion");
			resp.NumeroTarjeta = ExtraerValor(json, "NumeroTarjeta");
			resp.Bin = ExtraerValor(json, "Bin");
			resp.FechaVencimiento = ExtraerValor(json, "FechaVencimiento");
			resp.MensajeRespuesta = ExtraerValor(json, "MensajeRespuesta");
			resp.Filler = ExtraerValor(json, "Filler");
			resp.TipoTarjetaDescripcion = ExtraerValor(json, "TipoTarjetaDescripcion");

			return resp;
		}

		// PROCESO DE GUARDADO (igual que ya tenías antes)
		private void ProcesarRespuesta(Guid id, ConsultaTarjetaResponse resp)
		{
			string dataAscii = resp.DataAscii;
			if (dataAscii == null || dataAscii.Length < 125)
				throw new Exception("Respuesta incompleta");
			
			ActualizaConsultaTarjetaRequest lt = new ActualizaConsultaTarjetaRequest();

			lt.idPinpad_LecturaConsultaTarjeta = id;
			lt.TipoMensaje = resp.TipoMensaje;
			lt.CodigoRespuesta = resp.CodigoRespuesta;
			lt.NumeroTarjeta = resp.NumeroTarjeta;
			lt.Bin = resp.Bin;
			lt.FechaVencimiento = resp.FechaVencimiento;
			lt.MensajeRespuesta = resp.MensajeRespuesta;
			lt.Filler= resp.Filler;
			
			lt.TramaRecibida = resp.DataAscii;
			lt.FechaRespuesta = DateTime.Now;
			
			_repo.ActualizarRespuestaCT(lt);

		}
	}
}