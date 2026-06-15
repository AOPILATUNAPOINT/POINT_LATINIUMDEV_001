using System;
using Latinium.Models.Pinpad;

namespace Latinium.Services.Pinpad
{
	/// <summary>
	/// Servicio para Lectura de Tarjeta (LT)
	/// </summary>
	public class PinpadLTService : PinpadBaseService
	{
		public LecturaTarjetaResponse Procesar(LecturaTarjetaRequest request)
		{
			//Guid id = _repo.InsertarLectura(request.IdCompra, request.MontoTotal, "LT");

			string tallaHex;
			string trama = ArmarTrama(request.MontoTotal, out tallaHex);

//			_repo.ActualizarEnvio(id, "LT", trama, tallaHex);

			Guid id = _repo.InsertarLectura(
				request.IdCompra,
				request.MontoTotal,
				"LT",
				trama,
				tallaHex
				);

			string json = _tcp.Enviar(trama);

			LecturaTarjetaResponse resp = Parsear(json);

			if (!resp.Success)
				return resp;

			ProcesarRespuesta(id, resp);

			return resp;
		}

		private string ArmarTrama(decimal monto, out string tallaHex)
		{
			string cuerpo = "LT0000" + FormatearMonto(monto);

			tallaHex = CalcularTalla(cuerpo);

			string cuerpoFinal = cuerpo.Substring(0, 2) + tallaHex + cuerpo.Substring(6);

			return tallaHex + cuerpoFinal;
		}

		private LecturaTarjetaResponse Parsear(string json)
		{
			LecturaTarjetaResponse resp = new LecturaTarjetaResponse();
			
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
			resp.RedCorriente = ExtraerValor(json, "RedCorriente");
			resp.RedCorrienteDescripcion = ExtraerValor(json, "RedCorrienteDescripcion");
			resp.RedDiferido = ExtraerValor(json, "RedDiferido");
			resp.RedDiferidoDescripcion = ExtraerValor(json, "RedDiferidoDescripcion");
			resp.NumeroTarjetaTruncado = ExtraerValor(json, "NumeroTarjetaTruncado");
			resp.FechaVencimiento = ExtraerValor(json, "FechaVencimiento");
			resp.NumeroTarjetaEncriptado = ExtraerValor(json, "NumeroTarjetaEncriptado");
			resp.MensajeRespuesta = ExtraerValor(json, "MensajeRespuesta");
			resp.Filler = ExtraerValor(json, "Filler");
			
			return resp;
		}

		private void ProcesarRespuesta(Guid id, LecturaTarjetaResponse resp)
		{
			if (resp.DataAscii == "" || resp.DataAscii  == null || resp.DataAscii.Length < 120)
				throw new Exception("Respuesta incompleta");
			
			ActualizaLecturaTarjetaRequest lt = new ActualizaLecturaTarjetaRequest();
			
			lt.idPinpad_LecturaConsultaTarjeta = id;
			lt.TipoMensaje = resp.TipoMensaje;
			lt.CodigoRespuesta = resp.CodigoRespuesta;
			lt.RedCorriente = resp.RedCorriente;
			lt.RedDiferido = resp.RedDiferido;
			lt.NumeroTarjetaTruncado = resp.NumeroTarjetaTruncado;
			lt.FechaVencimiento = resp.FechaVencimiento;
			lt.NumeroTarjetaEncriptado = resp.NumeroTarjetaEncriptado;
			lt.MensajeRespuesta = resp.MensajeRespuesta;
			lt.Filler= resp.Filler;
			
			lt.TramaRecibida = resp.DataAscii;
			lt.FechaRespuesta = DateTime.Now;
			
			_repo.ActualizarRespuestaLT(lt);

		}
	}
}