using System;
using Latinium.Models.Pinpad;
using Latinium.Models.Pinpad.PP;

namespace Latinium.Services.Pinpad
{
	/// <summary>
	/// Servicio de Pago (PP)
	/// </summary>
	public class PinpadPPService : PinpadBaseService
	{
		public ProcesarPagoResponse Procesar(ProcesoPagoRequest request)
		{
			// ===== 1. ARMAR TRAMA =====
			string tallaHex;
			string trama = ArmarTrama(request, out tallaHex);

			request.Trama = trama;
			request.Talla = tallaHex;

			// ===== 2. INSERTAR TODO DE UNA =====
			Guid id = _repo.InsertarPPCompleto(request);

			// ===== 3. ENVIAR =====
			string json = _tcp.Enviar(trama);

			ProcesarPagoResponse resp = Parsear(json);

			if (resp == null)
				throw new Exception("Error parseando respuesta del pinpad");

			if (resp.Success)
				ProcesarRespuesta(id, resp);

			return resp;
		}

		// =========================
		// ARMAR TRAMA
		// =========================
		private string ArmarTrama(ProcesoPagoRequest r, out string tallaHex)
		{
			string cuerpo = ConstruirCuerpoPP(r);

			tallaHex = CalcularTalla(cuerpo);

			string cuerpoFinal = cuerpo.Substring(0, 2) + tallaHex + cuerpo.Substring(6);

			return tallaHex + cuerpoFinal;
		}

		// =========================
		// PARSEAR RESPUESTA JSON
		// =========================
		private ProcesarPagoResponse Parsear(string json)
		{
			ProcesarPagoResponse resp = new ProcesarPagoResponse();

			if (json == null || json.Trim() == "")
			{
				resp.Success = false;
				resp.ErrorMessage = "Respuesta vacía";
				return resp;
			}

			// =========================
			// CAMPOS BASE
			// =========================
			resp.Success = ExtraerBool(json, "Success");
			resp.DataAscii = ExtraerValor(json, "DataAscii");
			resp.ErrorMessage = ExtraerValor(json, "ErrorMessage");

			// =========================
			// CAMPOS DEL SERVICE (YA VIENEN PARSEADOS)
			// =========================
			resp.TipoMensaje = ExtraerValor(json, "TipoMensaje");
			resp.CodigoRespuestaMensaje = ExtraerValor(json, "CodigoRespuestaMensaje");
			resp.CodigoRed = ExtraerValor(json, "CodigoRed");
			resp.CodigoRespuestaAutorizador = ExtraerValor(json, "CodigoRespuestaAutorizador");
			resp.MensajeRespuesta = ExtraerValor(json, "MensajeRespuesta");

			resp.SecuencialTransaccion = ParseInt(ExtraerValor(json, "SecuencialTransaccion"));
			resp.NumeroLote = ParseInt(ExtraerValor(json, "NumeroLote"));

			resp.HoraTransaccion = ExtraerValor(json, "HoraTransaccion");
			resp.FechaTransaccion = ExtraerValor(json, "FechaTransaccion");
			resp.NumeroAutorizacion = ExtraerValor(json, "NumeroAutorizacion");

			resp.TerminalId = ExtraerValor(json, "TerminalId");
			resp.MerchantId = ExtraerValor(json, "MerchantId");

			resp.ValorInteres = ParseDecimal(ExtraerValor(json, "ValorInteres"));
			resp.MensajePublicidad = ExtraerValor(json, "MensajePublicidad");

			resp.CodigoBancoAdquiriente = ExtraerValor(json, "CodigoBancoAdquiriente");
			resp.NombreBancoEmisor = ExtraerValor(json, "NombreBancoEmisor");

			resp.GrupoTarjeta = ExtraerValor(json, "GrupoTarjeta");
			resp.ModoLectura = ExtraerValor(json, "ModoLectura");
			resp.NombreTarjetahabiente = ExtraerValor(json, "NombreTarjetahabiente");

			resp.MontoFijo = ParseDecimal(ExtraerValor(json, "MontoFijo"));

			// =========================
			// EMV / TARJETA
			// =========================
			resp.ApplicationLabel = ExtraerValor(json, "ApplicationLabel");
			resp.AID = ExtraerValor(json, "AID");
			resp.Criptograma = ExtraerValor(json, "Criptograma");
			resp.VerificacionPin = ExtraerValor(json, "VerificacionPin");
			resp.ARQC = ExtraerValor(json, "ARQC");
			resp.TVR = ExtraerValor(json, "TVR");
			resp.TSI = ExtraerValor(json, "TSI");

			resp.NumeroTarjetaTruncado = ExtraerValor(json, "NumeroTarjetaTruncado");
			resp.FechaVencimiento = ExtraerValor(json, "FechaVencimiento");
			resp.NumeroTarjetaEncriptado = ExtraerValor(json, "NumeroTarjetaEncriptado");

			// =========================
			// OTROS
			// =========================
			resp.Filler = ExtraerValor(json, "Filler");

			return resp;
		}
		// =========================
		// PROCESAR RESPUESTA ASCII
		// =========================
		private void ProcesarRespuesta(Guid id, ProcesarPagoResponse resp)
		{
			if (resp.DataAscii == null || resp.DataAscii.Length < 50)
				throw new Exception("Respuesta PP incompleta");

			// Ya no necesitas re-mapear a otro response
			// Solo construyes el request para persistencia

			ActualizarProcesarPagoRequest rq = new ActualizarProcesarPagoRequest();

			rq.idPinpad_ProcesoPago = id;

			rq.TipoMensaje = resp.TipoMensaje;
			rq.CodigoRespuestaMensaje = resp.CodigoRespuestaMensaje;
			rq.CodigoRed = resp.CodigoRed;
			rq.CodigoRespuestaAutorizador = resp.CodigoRespuestaAutorizador;

			rq.MensajeRespuesta = resp.MensajeRespuesta;

			rq.SecuencialTransaccion = resp.SecuencialTransaccion;
			rq.NumeroLote = resp.NumeroLote;

			rq.HoraTransaccion = resp.HoraTransaccion;
			rq.FechaTransaccion = resp.FechaTransaccion;

			rq.NumeroAutorizacion = resp.NumeroAutorizacion;
			rq.TerminalId = resp.TerminalId;
			rq.MerchantId = resp.MerchantId;

			rq.ValorInteres = resp.ValorInteres;
			rq.MensajePublicidad = resp.MensajePublicidad;

			rq.CodigoBancoAdquiriente = resp.CodigoBancoAdquiriente;
			rq.NombreBancoEmisor = resp.NombreBancoEmisor;

			rq.GrupoTarjeta = resp.GrupoTarjeta;
			rq.ModoLectura = resp.ModoLectura;

			rq.NombreTarjetahabiente = resp.NombreTarjetahabiente;

			rq.MontoFijo = resp.MontoFijo;

			rq.ApplicationLabel = resp.ApplicationLabel;
			rq.AID = resp.AID;

			rq.Criptograma = resp.Criptograma;
			rq.VerificacionPin = resp.VerificacionPin;
			rq.ARQC = resp.ARQC;

			rq.TVR = resp.TVR;
			rq.TSI = resp.TSI;

			rq.NumeroTarjetaTruncado = resp.NumeroTarjetaTruncado;
			rq.FechaVencimiento = resp.FechaVencimiento;
			rq.NumeroTarjetaEncriptado = resp.NumeroTarjetaEncriptado;
			rq.Filler = resp.Filler;
			rq.TramaRecibida = resp.DataAscii;
			rq.FechaRespuesta = DateTime.Now;

			_repo.GuardarProcesarPago(rq);
		}
		// =========================
		// CONSTRUIR CUERPO PP
		// =========================
		private string ConstruirCuerpoPP(ProcesoPagoRequest r)
		{
			string tipoMensaje = "PP";
			string talla = "0000";

			string tipoTransaccion = r.TipoTransaccion;
			string red = r.CodigoRed;

			string codDiferido = "00";
			string plazo = "  ";
			string gracia = "  ";

			if (tipoTransaccion == "02")
			{
				if (r.CodigoDiferido != null && r.CodigoDiferido != "")
					codDiferido = r.CodigoDiferido.PadLeft(2, '0');

				if (r.PlazoDiferido > 0)
					plazo = r.PlazoDiferido.ToString().PadLeft(2, '0');
				else
					plazo = "00";

				if (r.MesesGracia > 0)
					gracia = r.MesesGracia.ToString().PadLeft(2, '0');
				else
					gracia = "00";
			}

			string filler1 = " ";

			string montoTotal = FormatearMonto(r.MontoTotal);
			string montoIVA = FormatearMonto(r.BaseIVA);
			string monto0 = FormatearMonto(r.Base0);
			string iva = FormatearMonto(r.IVA);

			string servicio = new string(' ', 12);

			string propina;
			if (r.Propina > 0)
				propina = FormatearMonto(r.Propina);
			else
				propina = new string(' ', 12);

			string montoFijo = new string(' ', 12);
			string secuencial = "      ";

			DateTime now = DateTime.Now;
			string hora = now.Hour.ToString("00") +
				now.Minute.ToString("00") +
				now.Second.ToString("00");

			string fecha = now.Year.ToString("0000") +
				now.Month.ToString("00") +
				now.Day.ToString("00");

			string autorizacion = "      ";

			string mid = (r.MID != null)
				? r.MID.PadRight(15, ' ')
				: new string(' ', 15);

			string tid = (r.TID != null)
				? r.NumeroCaja.PadRight(8, ' ')
				: new string(' ', 8);

			string cid = (r.CID != null)
				? r.NumeroCaja.PadRight(15, ' ')
				: new string(' ', 15);

			string fillerFinal = new string(' ', 20);

			string cuerpo =
				tipoMensaje +
				talla +
				tipoTransaccion +
				red +
				codDiferido +
				plazo +
				gracia +
				filler1 +
				montoTotal +
				montoIVA +
				monto0 +
				iva +
				servicio +
				propina +
				montoFijo +
				secuencial +
				hora +
				fecha +
				autorizacion +
				mid +
				tid +
				cid +
				fillerFinal;

			return cuerpo;
		}

		// =========================
		// HELPERS
		// =========================
		private int ParseInt(string value)
		{
			try { return int.Parse(value.Trim()); }
			catch { return 0; }
		}

		private decimal ParseDecimal(string value)
		{
			try { return decimal.Parse(value.Trim()); }
			catch { return 0; }
		}
	}
}