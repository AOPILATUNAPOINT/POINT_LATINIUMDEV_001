//using System;
//using Latinium.Models.Pinpad;
//
//namespace Latinium.Services.Pinpad
//{
//	public class PinpadService
//	{
//		private PinpadTcpService _tcp = new PinpadTcpService();
//		private PinpadRepository _repo = new PinpadRepository();
//
//		public LecturaTarjetaResponse LeerTarjeta(LecturaTarjetaRequest request)
//		{
//			LecturaTarjetaResponse respuesta = new LecturaTarjetaResponse();
//			try
//			{
//				LecturaTarjetaResponse respLT = EjecutarLT(request);
//				respuesta = ConstruirRespuestaLT(respLT);
//			}
//			catch (Exception ex)
//			{
//				respuesta.CodigoRespuesta = "99";
//				respuesta.MensajeRespuesta = ex.Message;
//			}
//			return respuesta;
//		}
//		
//
//		public ConsultaTarjetaResponse ConsultaTarjeta(ConsultaTarjetaRequest request)
//		{
//			ConsultaTarjetaResponse respuesta = new ConsultaTarjetaResponse();
//			try
//			{
//				respuesta = EjecutarCT(request);
//			}
//			catch (Exception ex)
//			{
//				respuesta.CodigoRespuesta = "99";
//				respuesta.MensajeRespuesta = ex.Message;
//			}
//			return respuesta;
//		}
//		
//
//		private LecturaTarjetaResponse EjecutarLT(LecturaTarjetaRequest request)
//		{
//			LecturaTarjetaResponse resp = ProcesarLecturaLT(request.IdCompra, request.MontoTotal, "LT");
//			if (resp.CodigoRespuesta != "00")
//			{
//				throw new Exception(resp.MensajeRespuesta);
//			}
//			return resp;
//		}
//
//		private LecturaTarjetaResponse ConstruirRespuestaLT(LecturaTarjetaResponse lt)
//		{
//			LecturaTarjetaResponse respuesta = new LecturaTarjetaResponse();
//			// ===== BASE =====
//			respuesta.CodigoRespuesta = lt.CodigoRespuesta;
//			respuesta.RedCorriente = lt.RedCorriente;
//			respuesta.RedDiferido = lt.RedDiferido;
//			respuesta.NumeroTarjetaTruncado = lt.NumeroTarjetaTruncado;
//			// SOLO COPIAR LO QUE YA VIENE DEL SERVICIO
//			respuesta.FechaVencimiento = lt.FechaVencimiento;
//			respuesta.PanEncriptado = lt.PanEncriptado;
//			respuesta.MensajeRespuesta = lt.MensajeRespuesta;
//			respuesta.Filler = lt.Filler;
//			return respuesta;
//		}
//
//		private ConsultaTarjetaResponse ConstruirRespuestaCT(ConsultaTarjetaResponse ct)
//		{
//			ConsultaTarjetaResponse respuesta = new ConsultaTarjetaResponse();
//			respuesta.Bin = ct.Bin;
//			respuesta.CodigoRespuesta = ct.CodigoRespuesta;
//			return respuesta;
//		}
//
//		private ConsultaTarjetaResponse EjecutarCT(ConsultaTarjetaRequest request)
//		{
//			ConsultaTarjetaResponse resp = ProcesarLecturaCT(request.IdCompra, request.MontoTotal, "CT");
//			if (resp.CodigoRespuesta != "00")
//			{
//				throw new Exception(resp.MensajeRespuesta);
//			}
//			return resp;
//		}
//
////		public PinpadResponse ProcesarLectura(int idCompra, decimal monto, string tipo)
////		{
////			Guid id = _repo.InsertarLectura(idCompra, monto, tipo);
////			string tallaHex;
////			string trama = ArmarTrama(tipo, monto, out tallaHex);
////			_repo.ActualizarEnvio(id, tipo, trama, tallaHex);
////			string json = _tcp.Enviar(trama);
////			PinpadResponse resp = ParsearJson(json);
////			if (resp == null)
////				throw new Exception("Error parseando respuesta del pinpad");
////			if (!resp.Success)
////			{
////				return resp;
////			}
////			// PROCESAR TRAMA ASCII (CLAVE)
////			if (tipo == "LT")
////			{
////				ProcesarRespuestaLT(id, resp.DataAscii, resp);
////			}
////			return resp;
////		}
//
//		public LecturaTarjetaResponse ProcesarLecturaLT(int idCompra, decimal monto, string tipo)
//		{
//			Guid id = _repo.InsertarLectura(idCompra, monto, tipo);
//			string tallaHex;
//			string trama = ArmarTrama(tipo, monto, out tallaHex);
//			_repo.ActualizarEnvio(id, tipo, trama, tallaHex);
//			string json = _tcp.Enviar(trama);
//			LecturaTarjetaResponse resp = ParsearJsonLT(json);
//			if (resp == null)
//				throw new Exception("Error parseando respuesta del pinpad");
//			if (!resp.Success)
//			{
//				return resp;
//			}
//			// PROCESAR TRAMA ASCII (CLAVE)
//			if (tipo == "LT")
//			{
//				ProcesarRespuestaLT(id, resp.DataAscii, resp);
//			}
//			return resp;
//		}
//
//		public ConsultaTarjetaResponse ProcesarLecturaCT(int idCompra, decimal monto, string tipo)
//		{
//			Guid id = _repo.InsertarLectura(idCompra, monto, tipo);
//			string tallaHex;
//			string trama = ArmarTrama(tipo, monto, out tallaHex);
//			_repo.ActualizarEnvio(id, tipo, trama, tallaHex);
//			string json = _tcp.Enviar(trama);
//			ConsultaTarjetaResponse resp = ParsearJsonCT(json);
//			if (resp == null)
//				throw new Exception("Error parseando respuesta del pinpad");
//			if (!resp.Success)
//			{
//				return resp;
//			}
//			// PROCESAR TRAMA ASCII (CLAVE)
//			if (tipo == "CT")
//			{
//				ProcesarRespuestaCT(id, resp.DataAscii, resp);
//			}
//			return resp;
//		}
//
//		public SimulacionResponse ParseResponse(string json)
//		{
//			SimulacionResponse res = new SimulacionResponse();
//
//			res.Success = json.IndexOf("\"Success\":true") > -1;
//
//			res.DataAscii = ExtractValue(json, "DataAscii");
//			res.ErrorMessage = ExtractValue(json, "ErrorMessage");
//
//			string detalleJson = ExtractObject(json, "Detalle");
//
//			if (detalleJson != "")
//			{
//				LtDetalle det = new LtDetalle();
//
//				det.TipoFinanciamiento = ExtractValue(detalleJson, "Tipo");
//				det.EstadoTransaccion = ExtractValue(detalleJson, "CodigoRespuesta");
//
//				res.Detalle = det;
//			}
//
//			return res;
//		}
//
//		private string ExtractValue(string json, string key)
//		{
//			string pattern = "\"" + key + "\":\"";
//			int start = json.IndexOf(pattern);
//			if (start < 0) return "";
//
//			start += pattern.Length;
//			int end = json.IndexOf("\"", start);
//
//			return json.Substring(start, end - start);
//		}
//
//		private string ExtractObject(string json, string key)
//		{
//			string pattern = "\"" + key + "\":{";
//			int start = json.IndexOf(pattern);
//			if (start < 0) return "";
//
//			start += pattern.Length - 1;
//
//			int braceCount = 0;
//
//			for (int i = start; i < json.Length; i++)
//			{
//				if (json[i] == '{') braceCount++;
//				if (json[i] == '}')
//				{
//					if (braceCount == 0)
//						return json.Substring(start, i - start + 1);
//
//					braceCount--;
//				}
//			}
//
//			return "";
//		}
//
//		private string ArmarTrama(string tipo, decimal monto, out string tallaHex)
//		{
//			string cuerpo;
//
//			switch (tipo)
//			{
//				case "LT":
//					cuerpo = ConstruirCuerpoLT(monto);
//					break;
//
//				case "CT":
//					cuerpo = ConstruirCuerpoCT();
//					break;
//
//				default:
//					throw new ArgumentException("Tipo no válido. Solo se permite 'LT' o 'CT'.");
//			}
//
//		
//
//			tallaHex = CalcularTalla(cuerpo);
//
//			// REEMPLAZA LA TALLA INTERNA (posiciones 2-6)
//			string cuerpoFinal = cuerpo.Substring(0, 2) + tallaHex + cuerpo.Substring(6);
//
//			return tallaHex + cuerpoFinal;
//		}
//
//		private string ConstruirCuerpoLT(decimal monto)
//		{
//			string tipoMensaje = "LT";
//			string talla = "0000";
//			string montoStr = FormatearMonto(monto);
//
//			return tipoMensaje + talla + montoStr;
//		}
//
//		private string FormatearMonto(decimal monto)
//		{
//			long montoCentavos = (long)(monto * 100);
//			return montoCentavos.ToString().PadLeft(12, '0');
//		}
//
//		private string ConstruirCuerpoCT()
//		{
//			string tipoMensaje = "CT";
//			string talla = "0000";
//
//			return tipoMensaje + talla;
//		}
//
//		private const int HEADER_SIZE = 32;
//
//		private string CalcularTalla(string cuerpo)
//		{
//			int longitud = cuerpo.Length + HEADER_SIZE;
//			return longitud.ToString("X4");
//		}
//
//		private PinpadResponse ParsearJson(string json)
//		{
//			PinpadResponse resp = new PinpadResponse();
//
//			if (json == null || json.Trim() == "")
//			{
//				resp.Success = false;
//				resp.ErrorMessage = "Respuesta vacía";
//				return resp;
//			}
//
//			resp.Success = ExtraerBool(json, "Success");
//			resp.DataAscii = ExtraerValor(json, "DataAscii");
//			resp.ErrorMessage = ExtraerValor(json, "ErrorMessage");
//			resp.Codigo = ExtraerValor(json, "Codigo");
//			resp.Mensaje = ExtraerValor(json, "Mensaje");
//			resp.NumeroTarjetaTruncado = ExtraerValor(json, "NumeroTarjetaTruncado");
//			resp.Bin = ExtraerValor(json, "Bin");
//			resp.RedCorrienteDescripcion = ExtraerValor(json, "RedCorrienteDescripcion");
//			resp.RedDiferidoDescripcion = ExtraerValor(json, "RedDiferidoDescripcion");
//			resp.TipoTarjetaDescripcion = ExtraerValor(json, "TipoTarjetaDescripcion");
//
//			return resp;
//		}
//
//		private LecturaTarjetaResponse ParsearJsonLT(string json)
//		{
//			LecturaTarjetaResponse resp = new LecturaTarjetaResponse();
//
//			if (json == null || json.Trim() == "")
//			{
//				resp.Success = false;
//				resp.ErrorMessage = "Respuesta vacía";
//				return resp;
//			}
//
//			resp.Success = ExtraerBool(json, "Success");
//			resp.DataAscii = ExtraerValor(json, "DataAscii");
//			resp.ErrorMessage = ExtraerValor(json, "ErrorMessage");
//
//			resp.TipoMensaje = ExtraerValor(json, "TipoMensaje");
//			resp.CodigoRespuesta = ExtraerValor(json, "CodigoRespuesta");
//			resp.CodigoRespuestaDescripcion = ExtraerValor(json, "CodigoRespuestaDescripcion");
//			resp.RedCorriente = ExtraerValor(json, "RedCorriente");
//			resp.RedCorrienteDescripcion = ExtraerValor(json, "RedCorrienteDescripcion");
//			resp.RedDiferido = ExtraerValor(json, "RedDiferido");
//			resp.RedDiferidoDescripcion = ExtraerValor(json, "RedDiferidoDescripcion");
//			resp.NumeroTarjetaTruncado = ExtraerValor(json, "NumeroTarjetaTruncado");
//			resp.FechaVencimiento = ExtraerValor(json, "FechaVencimiento");
//			resp.PanEncriptado = ExtraerValor(json, "PanEncriptado");
//			resp.MensajeRespuesta = ExtraerValor(json, "MensajeRespuesta");
//			resp.Filler = ExtraerValor(json, "Filler");
//
//			return resp;
//		}
//
//		private ConsultaTarjetaResponse ParsearJsonCT(string json)
//		{
//			ConsultaTarjetaResponse resp = new ConsultaTarjetaResponse();
//
//			if (json == null || json.Trim() == "")
//			{
//				resp.Success = false;
//				resp.ErrorMessage = "Respuesta vacía";
//				return resp;
//			}
//
//			resp.Success = ExtraerBool(json, "Success");
//			resp.DataAscii = ExtraerValor(json, "DataAscii");
//			resp.ErrorMessage = ExtraerValor(json, "ErrorMessage");
//			resp.TipoMensaje = ExtraerValor(json, "TipoMensaje");
//			resp.CodigoRespuesta = ExtraerValor(json, "CodigoRespuesta");
//			resp.CodigoRespuestaDescripcion = ExtraerValor(json, "CodigoRespuestaDescripcion");
//			resp.NumeroTarjeta = ExtraerValor(json, "NumeroTarjeta");
//			resp.Bin = ExtraerValor(json, "Bin");
//			resp.FechaVencimiento = ExtraerValor(json, "FechaVencimiento");
//			resp.MensajeRespuesta = ExtraerValor(json, "MensajeRespuesta");
//			resp.Filler = ExtraerValor(json, "Filler");
//			resp.TipoTarjetaDescripcion = ExtraerValor(json, "TipoTarjetaDescripcion");
//
//			return resp;
//		}
//
//		// reutilizas tu parser actual
//		private string ExtraerValor(string json, string campo)
//		{
//			string patron = "\"" + campo + "\":";
//
//			int pos = json.IndexOf(patron);
//
//			if (pos < 0)
//				return "";
//
//			pos += patron.Length;
//
//			// Saltar espacios
//			while (pos < json.Length && (json[pos] == ' ' || json[pos] == '\"'))
//				pos++;
//
//			int fin = json.IndexOf("\"", pos);
//
//			if (fin < 0)
//				return "";
//
//			return json.Substring(pos, fin - pos);
//		}
//
//
//		private int ExtraerEntero(string json, string campo)
//		{
//			string patron = "\"" + campo + "\":";
//			int pos = json.IndexOf(patron);
//			if (pos < 0)
//				return 0;
//			pos += patron.Length;
//			// Saltar espacios
//			while (pos < json.Length && json[pos] == ' ')
//				pos++;
//			// Buscar fin del número
//			int fin = json.IndexOfAny(new char[] { ',', '}' }, pos);
//			if (fin < 0)
//				fin = json.Length;
//			string valor = json.Substring(pos, fin - pos).Trim();
//			try
//			{
//				return int.Parse(valor);
//			}
//			catch
//			{
//				return 0;
//			}
//		}
//
//		private bool ExtraerBool(string json, string campo)
//		{
//			string patron = "\"" + campo + "\":";
//
//			int pos = json.IndexOf(patron);
//			if (pos < 0)
//				return false;
//
//			pos += patron.Length;
//
//			// Saltar espacios
//			while (pos < json.Length && json[pos] == ' ')
//				pos++;
//
//			// Buscar fin (coma o cierre de objeto)
//			int fin = json.IndexOfAny(new char[] { ',', '}' }, pos);
//			if (fin < 0)
//				fin = json.Length;
//
//			string valor = json.Substring(pos, fin - pos).Trim().ToLower();
//
//			return valor == "true";
//		}
//
//		private void ProcesarRespuestaLT(Guid id, string dataAscii, LecturaTarjetaResponse resp)
//		{
//			if (dataAscii == "" || dataAscii == null || dataAscii.Length < 120)
//				throw new Exception("Respuesta incompleta");
//
//			ActualizaLecturaTarjetaRequest lt = new ActualizaLecturaTarjetaRequest();
//
//			lt.idPinpad_LecturaConsultaTarjeta = id;
//			lt.TipoMensaje = resp.TipoMensaje;
//			lt.CodigoRespuesta = resp.CodigoRespuesta;
//			lt.RedCorriente = resp.RedCorriente;
//			lt.RedDiferido = resp.RedDiferido;
//			lt.NumeroTarjetaTruncado = resp.NumeroTarjetaTruncado;
//			lt.FechaVencimiento = resp.FechaVencimiento;
//			lt.NumeroTarjetaTruncado = resp.NumeroTarjetaTruncado;
//			lt.MensajeRespuesta = resp.MensajeRespuesta;
//			lt.Filler= resp.Filler;
//
//			lt.TramaRecibida = dataAscii;
//			lt.FechaRespuesta = DateTime.Now;
//
//			_repo.ActualizarRespuestaLT(lt);
//
//		}
//		
//		public PinpadResponse ProcesarPagoPP(ProcesoPagoRequest request)
//		{
//			// Registrar en repositorio
//			Guid id = _repo.InsertarLecturaPP(request);
//
//			// Armar trama
//			string tallaHex;
//			string trama = ArmarTramaPP(request, out tallaHex);
//
//			// Guardar envío
//			_repo.ActualizarEnvioPP(id, "PP", trama, tallaHex);
//
//			// Enviar
//			string json = _tcp.Enviar(trama);
//
//			PinpadResponse resp = ParsearJson(json);
//
//			if (resp == null)
//				throw new Exception("Error parseando respuesta del pinpad");
//
//			if (resp.Success)
//			{
//				ProcesarRespuestaPP(id, resp.DataAscii, resp);
//			}
//			else
//			{
//				// _repo.ActualizarError(id, resp.ErrorMessage);
//			}
//
//			return resp;
//		}
//
//		private string ArmarTramaPP(ProcesoPagoRequest request, out string tallaHex)
//		{
//			string cuerpo = ConstruirCuerpoPP(request);
//
//			tallaHex = CalcularTalla(cuerpo);
//
//			// Igual que LT
//			string cuerpoFinal = cuerpo.Substring(0, 2) + tallaHex + cuerpo.Substring(6);
//
//			return tallaHex + cuerpoFinal;
//		}
//
//		private string ConstruirCuerpoPP(ProcesoPagoRequest r)
//		{
//			string tipoMensaje = "PP";
//			string talla = "0000";
//
//			string tipoTransaccion = r.TipoTransaccion;
//			string red = r.CodigoRed;
//
//			// ===== DIFERIDO vs CORRIENTE =====
//			string codDiferido;
//			string plazo;
//			string gracia;
//
//			if (tipoTransaccion == "02") // DIFERIDO
//			{
//				codDiferido = (r.CodigoDiferido != null && r.CodigoDiferido != "")
//					? r.CodigoDiferido.PadLeft(2, '0')
//					: "00";
//
//				plazo = (r.PlazoDiferido > 0)
//					? r.PlazoDiferido.ToString().PadLeft(2, '0')
//					: "00";
//
//				gracia = (r.MesesGracia > 0)
//					? r.MesesGracia.ToString().PadLeft(2, '0')
//					: "00";
//			}
//			else // CORRIENTE
//			{
//				codDiferido = "00";
//				plazo = "  ";   // espacios según estándar
//				gracia = "  ";  // espacios según estándar
//			}
//
//			string filler1 = " ";
//
//			// ===== MONTOS =====
//			string montoTotal = FormatearMonto(r.MontoTotal);
//			string montoIVA = FormatearMonto(r.BaseIVA);
//			string monto0 = FormatearMonto(r.Base0);
//			string iva = FormatearMonto(r.IVA);
//
//			string servicio = new string(' ', 12);
//
//			string propina = (r.Propina > 0)
//				? FormatearMonto(r.Propina)
//				: new string(' ', 12);
//
//			string montoFijo = new string(' ', 12);
//			string secuencial = "      ";
//
//			DateTime now = DateTime.Now;
//			string hora = now.ToString("HHmmss");
//			string fecha = now.ToString("yyyyMMdd");
//
//			string autorizacion = "      ";
//
//			string mid = r.NumeroComercio.PadRight(15, ' ');
//			string tid = r.NumeroCaja.PadRight(8, ' ');
//			string cid = r.NumeroCaja.PadRight(15, ' ');
//
//			string fillerFinal = new string(' ', 20);
//
//			return
//				tipoMensaje +
//				talla +
//				tipoTransaccion +
//				red +
//				codDiferido +
//				plazo +
//				gracia +
//				filler1 +
//				montoTotal +
//				montoIVA +
//				monto0 +
//				iva +
//				servicio +
//				propina +
//				montoFijo +
//				secuencial +
//				hora +
//				fecha +
//				autorizacion +
//				mid +
//				tid +
//				cid +
//				fillerFinal;
//		}
//
//		private void ProcesarRespuestaPP(Guid id, string dataAscii, PinpadResponse resp)
//		{
//			if (dataAscii == null || dataAscii.Length < 100)
//				throw new Exception("Respuesta PP incompleta");
//
//			int i = 0;
//
//			string tipo = dataAscii.Substring(i, 2); i += 2;
//			string codigo = dataAscii.Substring(i, 2); i += 2;
//
//			// Ajusta según tu formato real
//			string autorizacion = dataAscii.Substring(i, 6); i += 6;
//			string mensaje = dataAscii.Substring(i, 20).Trim(); i += 20;
//
//			string estado = "RESPONDIDO";
//			string tramaRecibida = dataAscii;
//
//			// Guarda en repo (puedes crear método específico)
//			_repo.ActualizarRespuestaCT(
//				id,
//				codigo,
//				"", // tarjeta
//				"", // bin
//				"",
//				mensaje,
//				"",
//				null,
//				tramaRecibida,
//				estado,
//				DateTime.MinValue,
//				DateTime.Now
//				);
//
//			resp.Codigo = codigo;
//			resp.Mensaje = mensaje;
//		}
//
//
//		private void ProcesarRespuestaCT(Guid id, string dataAscii, ConsultaTarjetaResponse resp)
//		{
//			// Validación mínima
//			if (dataAscii == null || dataAscii.Length < 125)
//				throw new Exception("Respuesta incompleta");
//
//			int i = 0;
//
//			// Leer campos de la trama según esquema CT
//			string tipo = dataAscii.Substring(i, 2); i += 2;
//			string codigo = dataAscii.Substring(i, 2); i += 2;
//			string tarjeta = dataAscii.Substring(i, 64).Trim(); i += 64;
//			string bin = dataAscii.Substring(i, 8); i += 8;
//			string fecha = dataAscii.Substring(i, 4); i += 4;
//			string mensaje = dataAscii.Substring(i, 20).Trim(); i += 20;
//			string filler = dataAscii.Substring(i, 27); i += 27;
//
//			string estado = "RESPONDIDO";
//			string tramaRecibida = dataAscii; // toda la trama recibida
//
//			// Actualizar repositorio CT (sin parámetros de LT)
//			_repo.ActualizarRespuestaCT(
//				id,
//				codigo,
//				tarjeta,
//				bin,
//				fecha,
//				mensaje,
//				filler,
//				null,          // TramaEnviada no aplica
//				tramaRecibida, // TramaRecibida
//				estado,
//				DateTime.MinValue, // FechaEnvio no se actualiza
//				DateTime.Now       // FechaRespuesta actual
//				);
//
//			// Llenar todos los campos del response
//			resp.CodigoRespuesta = codigo;
//			resp.MensajeRespuesta = mensaje;
//			resp.NumeroTarjeta = tarjeta; // opcional, si quieres tenerlo también en CT
//			resp.Bin = bin;                        // esto es lo importante para tu UI
//			resp.CodigoRespuesta = codigo;
//		}
//
//		public PagoResponse ProcesarPagoModulo(ProcesoPagoRequest request)
//		{
//			PagoResponse respuesta = new PagoResponse();
//
//			try
//			{
//				// ===== Validar monto =====
//				if (request.MontoTotal <= 0)
//					throw new Exception("Monto total inválido");
//
//				// ===== Generar trama y talla INTERNAMENTE =====
//				string tallaHex;
//				string trama = ArmarTramaPP(request, out tallaHex);
//
//				// (Opcional) guardar en request si lo necesitas luego
//				request.Trama = trama;
//				request.Talla = tallaHex;
//
//				// ===== Insertar en BD =====
//				_repo.InsertarLecturaPP(request);
//
//				// ===== Llamar al servicio =====
//				PinpadResponse respService = ProcesarPagoPP(request);
//
//				// ===== Validar respuesta =====
//				if (!respService.Success)
//					throw new Exception(respService.ErrorMessage);
//
//				// ===== Respuesta final =====
//				respuesta.Success = true;
//				respuesta.Codigo = respService.Codigo;
//				respuesta.Mensaje = respService.Mensaje;
//				respuesta.ErrorMessage = "";
//			}
//			catch (Exception ex)
//			{
//				respuesta.Success = false;
//				respuesta.ErrorMessage = ex.Message;
//				respuesta.Codigo = "99";
//				respuesta.Mensaje = "ERROR";
//			}
//
//			return respuesta;
//		}
//		/// <summary>
//		/// Crea la trama de la transacción de pago según el requerimiento
//		/// </summary>
//		private string CrearTramaPP(ProcesoPagoRequest request)
//		{
//			// Variables para la trama
//			string trama = "";
//
//			// ===== Tamaño de la trama: se calcula al final + 32 bytes de la llave =====
//			string talla = "0000"; // placeholder temporal
//
//			// ===== Tipo de mensaje =====
//			trama += "PP"; // 2 AN
//
//			// ===== Tipo de transacción =====
//			trama += request.TipoTransaccion.PadLeft(2, '0'); // 2 AN, "01", "02", etc.
//
//			// ===== Código de Red =====
//			trama += request.CodigoRed; // 1 N
//
//			// ===== Código de diferido =====
//			trama += request.CodigoDiferido.PadLeft(2, '0'); // 2 N
//
//			// ===== Plazo Diferido =====
//			trama += request.PlazoDiferido.ToString().PadLeft(2, '0'); // 2 N
//
//			// ===== Meses de gracia =====
//			trama += request.MesesGracia.ToString().PadLeft(2, '0'); // 2 N
//
//			// ===== Filler 1 N =====
//			trama += " "; // 1 N
//
//			// ===== Monto total 12 N =====
//			trama += FormatearMonto(request.MontoTotal, 12);
//
//			// ===== Base IVA 12 N =====
//			trama += FormatearMonto(request.BaseIVA, 12);
//
//			// ===== Base 0 12 N =====
//			trama += FormatearMonto(request.Base0, 12);
//
//			// ===== IVA 12 N =====
//			trama += FormatearMonto(request.IVA, 12);
//
//			// ===== Propina 12 N =====
//			trama += FormatearMonto(request.Propina, 12);
//
//			// ===== Campos adicionales fijos / blanks =====
//			trama += new string(' ', 6 + 8 + 6); // Secuencial, Fecha, Hora (ejemplo, ajustar si aplica)
//			trama += request.NumeroComercio.PadRight(15, ' '); // MID 15 AN
//			trama += request.NumeroCaja.PadRight(8, ' ');       // TID 8 AN
//			trama += request.NumeroCaja.PadRight(15, ' ');      // CID 15 AN
//			trama += new string(' ', 20);                       // Filler 20 AN
//
//			// ===== Calcular talla real de la trama =====
//			int longitudTrama = trama.Length + 32; // +32 bytes de la llave
//			talla = longitudTrama.ToString("X4");  // Convertir a HEX 4 dígitos
//
//			// ===== Reemplazar talla al inicio de la trama =====
//			trama = talla + trama;
//
//			return trama;
//		}
//
//		/// <summary>
//		/// Formatea un monto decimal a 12 dígitos numéricos: 10 enteros + 2 decimales, ceros a la izquierda
//		/// </summary>
//		private string FormatearMonto(decimal monto, int totalLength)
//		{
//			// Multiplicamos por 100 para mover decimales
//			long montoInt = (long)Math.Round(monto * 100, 0);
//			return montoInt.ToString().PadLeft(totalLength, '0');
//		}
//	}
//}
using System;
using Latinium.Models.Pinpad;
using Latinium.Models.Pinpad.PP;

namespace Latinium.Services.Pinpad
{
	public class PinpadService
	{
		private PinpadLTService _lt = new PinpadLTService();
		private PinpadCTService _ct = new PinpadCTService();
		private PinpadPPService _pp = new PinpadPPService();

		public LecturaTarjetaResponse LeerTarjeta(LecturaTarjetaRequest request)
		{
			return _lt.Procesar(request);
		}

		public ConsultaTarjetaResponse ConsultaTarjeta(ConsultaTarjetaRequest request)
		{
			return _ct.Procesar(request);
		}

		public ProcesarPagoResponse ProcesarPago(ProcesoPagoRequest request)
		{
			return _pp.Procesar(request);
		}
	}
}