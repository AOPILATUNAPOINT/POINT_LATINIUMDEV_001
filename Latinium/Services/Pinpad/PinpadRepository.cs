using System;
using System.Data;
using System.Data.SqlClient;
using System.Collections; // <- para ArrayList
using Latinium.Models.Pinpad;
using Latinium.Models.Pinpad.PP;

namespace Latinium.Services.Pinpad
{
	/// <summary>
	/// Descripción breve de PinpadRepository.
	/// </summary>
	public class PinpadRepository
	{
		private string _cn = MenuLatinium.sconexionEmpresa;
		public Guid InsertarLectura(int idCompra, decimal monto, string tipoMensaje, string trama, string tallaHex)
		{
			Guid id = Guid.NewGuid();

			SqlConnection cn = new SqlConnection(_cn);
			cn.Open();

			try
			{
				SqlCommand cmd = new SqlCommand(
					"INSERT INTO Pinpad_LecturaConsultaTarjeta (" +
					"idPinpad_LecturaConsultaTarjeta, " +
					"IdCompra, " +
					"MontoTotal, " +
					"TipoMensaje, " +
					"TramaEnviada, " +
					"Talla, " +
					"FechaEnvio" +
					") VALUES (" +
					"@id, @idCompra, @monto, @tipo, @trama, @talla, GETDATE()" +
					")", cn);

				// Parámetros
				SqlParameter p1 = new SqlParameter("@id", SqlDbType.UniqueIdentifier);
				p1.Value = id;
				cmd.Parameters.Add(p1);

				SqlParameter p2 = new SqlParameter("@idCompra", SqlDbType.Int);
				p2.Value = idCompra;
				cmd.Parameters.Add(p2);

				SqlParameter p3 = new SqlParameter("@monto", SqlDbType.Decimal);
				p3.Value = monto;
				cmd.Parameters.Add(p3);

				SqlParameter p4 = new SqlParameter("@tipo", SqlDbType.VarChar, 2);
				p4.Value = tipoMensaje;
				cmd.Parameters.Add(p4);

				SqlParameter p5 = new SqlParameter("@trama", SqlDbType.VarChar);
				p5.Value = trama;
				cmd.Parameters.Add(p5);

				SqlParameter p6 = new SqlParameter("@talla", SqlDbType.VarChar, 4);
				p6.Value = tallaHex;
				cmd.Parameters.Add(p6);

				cmd.ExecuteNonQuery();
			}
			finally
			{
				cn.Close();
			}

			return id;
		}

		public Guid InsertarLecturaPP(ProcesoPagoRequest request)
		{

			// Validación manual (compatible con .NET 1.1)
			if (request.Trama == null || request.Trama == "" ||
				request.Talla == null || request.Talla == "")
			{
				throw new ArgumentException("El request debe contener la Trama y la Talla");
			}

			using (SqlConnection cn = new SqlConnection(_cn))
			{
				cn.Open();
				InsertarPP(cn, request);
			}

			return request.idPinpad_ProcesoPago;
		}
	
		public void InsertarPP(SqlConnection cn, ProcesoPagoRequest request)
		{
			try
			{
				SqlCommand cmd = new SqlCommand(@"
            INSERT INTO Pinpad_ProcesoPago
            (
                idPinpadProcesoPago,
                IdCompra,
                TipoMensaje,
                TipoTransaccion,
                CodigoRed,
                CodigoDiferido,
                PlazoDiferido,
                MesesGracia,
                NumeroCaja,
                MontoTotal,
                BaseIVA,
                Base0,
                IVA,
                Propina,
                Trama,
                Talla
            )
            VALUES
            (
                @id,
                @idCompra,
                @tipoMensaje,
                @tipoTransaccion,
                @codigoRed,
                @codigoDiferido,
                @plazoDiferido,
                @mesesGracia,
                @numeroCaja,
                @montoTotal,
                @baseIVA,
                @base0,
                @iva,
                @propina,
                @trama,
                @talla
            )", cn);

				// ===== ID =====
				cmd.Parameters.Add("@id", SqlDbType.UniqueIdentifier).Value = request.idPinpad_ProcesoPago;

				// ===== IDENTIFICACIÓN =====
				if (request.idCompra > 0)
					cmd.Parameters.Add("@idCompra", SqlDbType.Int).Value = request.idCompra;
				else
					cmd.Parameters.Add("@idCompra", SqlDbType.Int).Value = DBNull.Value;

				// ===== TIPO =====
				cmd.Parameters.Add("@tipoMensaje", SqlDbType.Char, 2).Value = request.TipoMensaje;
				cmd.Parameters.Add("@tipoTransaccion", SqlDbType.Char, 2).Value = request.TipoTransaccion;
				cmd.Parameters.Add("@codigoRed", SqlDbType.Char, 1).Value = request.CodigoRed;

				// ===== DIFERIDO =====
				if (request.CodigoDiferido != null)
					cmd.Parameters.Add("@codigoDiferido", SqlDbType.Char, 2).Value = request.CodigoDiferido;
				else
					cmd.Parameters.Add("@codigoDiferido", SqlDbType.Char, 2).Value = DBNull.Value;

				cmd.Parameters.Add("@plazoDiferido", SqlDbType.Char, 2).Value = request.PlazoDiferido.ToString("00");
				cmd.Parameters.Add("@mesesGracia", SqlDbType.Char, 2).Value = request.MesesGracia.ToString("00");

				// ===== CAJA / COMERCIO =====
				cmd.Parameters.Add("@numeroCaja", SqlDbType.Char, 10).Value = request.NumeroCaja;

				// ===== MONTO =====
				cmd.Parameters.Add("@montoTotal", SqlDbType.Decimal).Value = request.MontoTotal;
				cmd.Parameters.Add("@baseIVA", SqlDbType.Decimal).Value = request.BaseIVA;
				cmd.Parameters.Add("@base0", SqlDbType.Decimal).Value = request.Base0;
				cmd.Parameters.Add("@iva", SqlDbType.Decimal).Value = request.IVA;
				cmd.Parameters.Add("@propina", SqlDbType.Decimal).Value = request.Propina;

				// ===== TRAMA =====
				if (request.Trama != null)
					cmd.Parameters.Add("@trama", SqlDbType.VarChar).Value = request.Trama;
				else
					cmd.Parameters.Add("@trama", SqlDbType.VarChar).Value = DBNull.Value;

				if (request.Talla != null)
					cmd.Parameters.Add("@talla", SqlDbType.Char, 4).Value = request.Talla;
				else
					cmd.Parameters.Add("@talla", SqlDbType.Char, 4).Value = DBNull.Value;

				cmd.ExecuteNonQuery();
			}
			catch (Exception ex)
			{
				throw new Exception("Error insertando proceso de pago en Pinpad_ProcesoPago.", ex);
			}
		}

		public void ActualizarEnvioPP(Guid id, string tipoMensaje, string trama, string tallaHex)
		{
			SqlConnection cn = new SqlConnection(_cn);
			cn.Open();

			try
			{
				if (tipoMensaje != "PP")
				{
					throw new ArgumentException("TipoMensaje no válido. Solo se permite 'PP'.");
				}

				SqlCommand cmd = new SqlCommand(@"
            UPDATE Pinpad_ProcesoPago
            SET 
                Trama = @trama,
                Talla = @talla
            WHERE idPinpadProcesoPago = @id", cn);

				// ===== Parámetros (.NET 1.1) =====

				SqlParameter p1 = new SqlParameter("@trama", SqlDbType.VarChar);
				if (trama != null)
					p1.Value = trama;
				else
					p1.Value = DBNull.Value;

				cmd.Parameters.Add(p1);

				SqlParameter p2 = new SqlParameter("@talla", SqlDbType.Char, 4);
				if (tallaHex != null)
					p2.Value = tallaHex;
				else
					p2.Value = DBNull.Value;

				cmd.Parameters.Add(p2);

				SqlParameter p3 = new SqlParameter("@id", SqlDbType.UniqueIdentifier);
				p3.Value = id;
				cmd.Parameters.Add(p3);

				// Ejecutar
				cmd.ExecuteNonQuery();
			}
			finally
			{
				cn.Close();
			}
		}

		public void ActualizarRespuestaLT(ActualizaLecturaTarjetaRequest lt)
		{
			SqlConnection cn = new SqlConnection(_cn);

			try
			{
				cn.Open();

				ArrayList setList = new ArrayList();

				// =========================
				// CAMPOS OBLIGATORIOS
				// =========================
				setList.Add("CodigoRespuesta = @CodigoRespuesta");
				setList.Add("CodigoRedCorriente = @CodigoRedCorriente");
				setList.Add("CodigoRedDiferido = @CodigoRedDiferido");
				setList.Add("NumeroTarjetaTruncado = @NumeroTarjetaTruncado");
				setList.Add("FechaVencimiento = @FechaVencimiento");
				setList.Add("NumeroTarjetaEncriptado = @NumeroTarjetaEncriptado");
				setList.Add("MensajeRespuesta = @MensajeRespuesta");

				// =========================
				// CAMPOS OPCIONALES
				// =========================
				if (lt.Filler != null && lt.Filler.Length > 0)
					setList.Add("Filler = @Filler");

				if (lt.TramaRecibida != null && lt.TramaRecibida.Length > 0)
					setList.Add("TramaRecibida = @TramaRecibida");

				if (lt.FechaRespuesta != DateTime.MinValue)
					setList.Add("FechaRespuesta = @FechaRespuesta");

				// =========================
				// ARMAR SET
				// =========================
				string setClause = "";
				for (int i = 0; i < setList.Count; i++)
				{
					if (i > 0) setClause += ", ";
					setClause += setList[i].ToString();
				}

				SqlCommand cmd = new SqlCommand(
					"UPDATE Pinpad_LecturaConsultaTarjeta SET " + setClause +
					" WHERE idPinpad_LecturaConsultaTarjeta = @idPinpad_LecturaConsultaTarjeta",
					cn
					);

				// =========================
				// PARÁMETROS OBLIGATORIOS
				// =========================
				cmd.Parameters.Add(new SqlParameter("@CodigoRespuesta", SqlDbType.Char, 2)).Value =
					(lt.CodigoRespuesta == null) ? (object)DBNull.Value : lt.CodigoRespuesta;

				cmd.Parameters.Add(new SqlParameter("@CodigoRedCorriente", SqlDbType.Char, 1)).Value =
					(lt.RedCorriente == null) ? (object)DBNull.Value : lt.RedCorriente;

				cmd.Parameters.Add(new SqlParameter("@CodigoRedDiferido", SqlDbType.Char, 1)).Value =
					(lt.RedDiferido == null) ? (object)DBNull.Value : lt.RedDiferido;

				cmd.Parameters.Add(new SqlParameter("@NumeroTarjetaTruncado", SqlDbType.VarChar, 25)).Value =
					(lt.NumeroTarjetaTruncado == null) ? (object)DBNull.Value : lt.NumeroTarjetaTruncado;

				cmd.Parameters.Add(new SqlParameter("@FechaVencimiento", SqlDbType.Char, 4)).Value =
					(lt.FechaVencimiento == null || lt.FechaVencimiento == "")
					? (object)DBNull.Value : lt.FechaVencimiento;

				cmd.Parameters.Add(new SqlParameter("@NumeroTarjetaEncriptado", SqlDbType.VarChar, 128)).Value =
					(lt.NumeroTarjetaEncriptado == null) ? (object)DBNull.Value : lt.NumeroTarjetaEncriptado;

				cmd.Parameters.Add(new SqlParameter("@MensajeRespuesta", SqlDbType.VarChar, 50)).Value =
					(lt.MensajeRespuesta == null) ? (object)DBNull.Value : lt.MensajeRespuesta;

				// =========================
				// PARÁMETROS OPCIONALES
				// =========================
				if (lt.TramaRecibida != null && lt.TramaRecibida.Length > 0)
				{
					cmd.Parameters.Add(new SqlParameter("@TramaRecibida", SqlDbType.VarChar, 4000)).Value = lt.TramaRecibida;
				}

				if (lt.FechaRespuesta != DateTime.MinValue)
				{
					cmd.Parameters.Add(new SqlParameter("@FechaRespuesta", SqlDbType.DateTime)).Value = lt.FechaRespuesta;
				}

				// =========================
				// ID (IMPORTANTE)
				// =========================
				cmd.Parameters.Add(new SqlParameter("@idPinpad_LecturaConsultaTarjeta", SqlDbType.UniqueIdentifier)).Value = lt.idPinpad_LecturaConsultaTarjeta;

				cmd.ExecuteNonQuery();
			}
			finally
			{
				if (cn != null && cn.State != ConnectionState.Closed)
				{
					cn.Close();
				}
			}
		}

		//		public void ActualizarRespuestaCT(
		//			Guid id,
		//			string codigo,
		//			string tarjeta,
		//			string bin,
		//			string fecha,
		//			string mensaje,
		//			string filler,
		//			string tramaEnviada,
		//			string tramaRecibida,
		//			string estado,
		//			DateTime fechaEnvio,
		//			DateTime fechaRespuesta
		//			)
		//		{
		//			SqlConnection cn = new SqlConnection(_cn);
		//			cn.Open();
		//
		//			try
		//			{
		//				ArrayList setList = new ArrayList();
		//				setList.Add("CodigoRespuesta = @codigo");
		//				setList.Add("NumeroTarjetaTruncado = @tarjeta");
		//				setList.Add("BIN = @bin");
		//				setList.Add("FechaVencimiento = @fecha");
		//				setList.Add("MensajeRespuesta = @mensaje");
		//
		//				if (filler != null && filler.Length > 0) setList.Add("Filler = @filler");
		//				if (tramaRecibida != null && tramaRecibida.Length > 0) setList.Add("TramaRecibida = @tramaRecibida");
		//				if (fechaRespuesta != DateTime.MinValue) setList.Add("FechaRespuesta = @fechaRespuesta");
		//
		//				string setClause = string.Join(", ", (string[])setList.ToArray(typeof(string)));
		//				SqlCommand cmd = new SqlCommand("UPDATE Pinpad_LecturaConsultaTarjeta SET " + setClause + " WHERE idPinpad_LecturaConsultaTarjeta = @id", cn);
		//
		//				// Parámetros
		//				cmd.Parameters.Add(new SqlParameter("@codigo", SqlDbType.Char, 2)).Value = codigo;
		//				cmd.Parameters.Add(new SqlParameter("@tarjeta", SqlDbType.VarChar, 25)).Value = tarjeta;
		//				cmd.Parameters.Add(new SqlParameter("@bin", SqlDbType.Char, 8)).Value = bin;
		//				cmd.Parameters.Add(new SqlParameter("@fecha", SqlDbType.Char, 4)).Value = fecha;
		//				cmd.Parameters.Add(new SqlParameter("@mensaje", SqlDbType.VarChar, 50)).Value = mensaje;
		//
		//				if (filler != null && filler.Length > 0) cmd.Parameters.Add(new SqlParameter("@filler", SqlDbType.VarChar, 27)).Value = filler;
		//				if (tramaRecibida != null && tramaRecibida.Length > 0) cmd.Parameters.Add(new SqlParameter("@tramaRecibida", SqlDbType.VarChar, 4000)).Value = tramaRecibida;
		//				if (fechaRespuesta != DateTime.MinValue) cmd.Parameters.Add(new SqlParameter("@fechaRespuesta", SqlDbType.DateTime)).Value = fechaRespuesta;
		//
		//				cmd.Parameters.Add(new SqlParameter("@id", SqlDbType.UniqueIdentifier)).Value = id;
		//
		//				cmd.ExecuteNonQuery();
		//			}
		//			finally
		//			{
		//				cn.Close();
		//			}
		//		}

		public Guid InsertarPPCompleto(ProcesoPagoRequest request)
		{
			SqlConnection cn = new SqlConnection(_cn);

			try
			{
				cn.Open();

				SqlCommand cmd = new SqlCommand(@"
INSERT INTO Pinpad_ProcesoPago
(
    idPinpad_ProcesoPago,
    idCompra,
    TipoMensaje,
    TipoTransaccion,
    CodigoRed,
    CodigoDiferido,
    PlazoDiferido,
    MesesGracia,
    NumeroCaja,
    MontoTotal,
    BaseIVA,
    Base0,
    IVA,
    Propina,
    Trama,
    Talla,
    SecuencialTransaccion,
    HoraTransaccion,
    FechaTransaccion,
    NumeroAutorizacion,
    MID,
    TID,
    CID,
    Usuario
)
VALUES
(
    @id,
    @idCompra,
    @tipoMensaje,
    @tipoTransaccion,
    @codigoRed,
    @codigoDiferido,
    @plazoDiferido,
    @mesesGracia,
    @numeroCaja,
    @montoTotal,
    @baseIVA,
    @base0,
    @iva,
    @propina,
    @trama,
    @talla,
    @secuencial,
    @hora,
    @fecha,
    @autorizacion,
    @mid,
    @tid,
    @cid,
    @usuario
)", cn);

				// ===== ID =====
				if (request.idPinpad_ProcesoPago == Guid.Empty)
					request.idPinpad_ProcesoPago = Guid.NewGuid();

				cmd.Parameters.Add("@id", SqlDbType.UniqueIdentifier).Value = request.idPinpad_ProcesoPago;

				// ===== COMPRA =====
				if (request.idCompra > 0)
					cmd.Parameters.Add("@idCompra", SqlDbType.Int).Value = request.idCompra;
				else
					cmd.Parameters.Add("@idCompra", SqlDbType.Int).Value = DBNull.Value;

				// ===== TIPO =====
				cmd.Parameters.Add("@tipoMensaje", SqlDbType.Char, 2).Value = request.TipoMensaje;
				cmd.Parameters.Add("@tipoTransaccion", SqlDbType.Char, 2).Value = request.TipoTransaccion;
				cmd.Parameters.Add("@codigoRed", SqlDbType.Char, 1).Value = request.CodigoRed;

				// ===== DIFERIDO =====
				if (request.CodigoDiferido != null && request.CodigoDiferido != "")
					cmd.Parameters.Add("@codigoDiferido", SqlDbType.Char, 2).Value = request.CodigoDiferido;
				else
					cmd.Parameters.Add("@codigoDiferido", SqlDbType.Char, 2).Value = DBNull.Value;

				cmd.Parameters.Add("@plazoDiferido", SqlDbType.Char, 2).Value = request.PlazoDiferido.ToString("00");
				cmd.Parameters.Add("@mesesGracia", SqlDbType.Char, 2).Value = request.MesesGracia.ToString("00");

				// ===== CAJA =====
				cmd.Parameters.Add("@numeroCaja", SqlDbType.Char, 10).Value = request.NumeroCaja;

				// ===== MONTOS =====
				cmd.Parameters.Add("@montoTotal", SqlDbType.Decimal).Value = request.MontoTotal;
				cmd.Parameters.Add("@baseIVA", SqlDbType.Decimal).Value = request.BaseIVA;
				cmd.Parameters.Add("@base0", SqlDbType.Decimal).Value = request.Base0;
				cmd.Parameters.Add("@iva", SqlDbType.Decimal).Value = request.IVA;
				cmd.Parameters.Add("@propina", SqlDbType.Decimal).Value = request.Propina;

				// ===== TRAMA =====
				cmd.Parameters.Add("@trama", SqlDbType.VarChar).Value = request.Trama;
				cmd.Parameters.Add("@talla", SqlDbType.Char, 4).Value = request.Talla;

				// ===== FECHA Y HORA =====
				DateTime now = DateTime.Now;

				string hora = now.Hour.ToString("00") +
					now.Minute.ToString("00") +
					now.Second.ToString("00");

				string fecha = now.Year.ToString("0000") +
					now.Month.ToString("00") +
					now.Day.ToString("00");

				cmd.Parameters.Add("@hora", SqlDbType.Char, 6).Value = hora;
				cmd.Parameters.Add("@fecha", SqlDbType.Char, 8).Value = fecha;

				// ===== OTROS =====
				cmd.Parameters.Add("@secuencial", SqlDbType.VarChar, 6).Value = "";
				cmd.Parameters.Add("@autorizacion", SqlDbType.VarChar, 6).Value = "";

				// ===== PINPAD =====
				cmd.Parameters.Add("@mid", SqlDbType.Char, 15).Value = request.MID;
				cmd.Parameters.Add("@tid", SqlDbType.Char, 8).Value = request.TID;
				cmd.Parameters.Add("@cid", SqlDbType.Char, 15).Value = request.CID;

				// ===== AUDITORÍA =====
				cmd.Parameters.Add("@usuario", SqlDbType.VarChar, 50).Value = request.Usuario;

				cmd.ExecuteNonQuery();
			}
			finally
			{
				cn.Close();
			}

			return request.idPinpad_ProcesoPago;
		}



		public void GuardarProcesarPago(ActualizarProcesarPagoRequest rq)
		{
			SqlConnection conn = new SqlConnection(_cn);

			try
			{
				conn.Open();

				string sql = @"
UPDATE Pinpad_ProcesoPago
SET
    TerminalId = @TerminalId,
    MerchantId = @MerchantId,
    SecuencialTransaccion = @SecuencialTransaccion,
    NumeroLote = @NumeroLote,
    NumeroAutorizacion = @NumeroAutorizacion,
    FechaTransaccion = @FechaTransaccion,
    HoraTransaccion = @HoraTransaccion,
    CodigoRespuestaMensaje = @CodigoRespuestaMensaje,
    CodigoRespuestaAutorizador = @CodigoRespuestaAutorizador,
    MensajeRespuesta = @MensajeRespuesta,
    CodigoRed = @CodigoRed,
    NumeroTarjetaTruncado = @NumeroTarjetaTruncado,
    FechaVencimiento = @FechaVencimiento,
    NombreTarjetahabiente = @NombreTarjetahabiente,
    NumeroTarjetaEncriptado = @NumeroTarjetaEncriptado,
    MontoFijo = @MontoFijo,
    ValorInteres = @ValorInteres,
    MensajePublicidad = @MensajePublicidad,
    CodigoBancoAdquiriente = @CodigoBancoAdquiriente,
    NombreBancoEmisor = @NombreBancoEmisor,
    AID = @AID,
    ApplicationLabel = @ApplicationLabel,
    Criptograma = @Criptograma,
    ARQC = @ARQC,
    TVR = @TVR,
    TSI = @TSI,
    GrupoTarjeta = @GrupoTarjeta,
    ModoLectura = @ModoLectura,
    VerificacionPin = @VerificacionPin,
    Filler = @Filler,
    TramaRecibida = @TramaRecibida,
    FechaRespuesta = @FechaRespuesta,
    Success = @Success,
    ErrorMessage = @ErrorMessage
WHERE
    idPinpad_ProcesoPago = @id
";

				SqlCommand cmd = new SqlCommand(sql, conn);

				// =========================
				// PARAMETROS
				// =========================

				cmd.Parameters.Add("@id", SqlDbType.UniqueIdentifier).Value = rq.idPinpad_ProcesoPago;

				SqlParameter pTerminal = cmd.Parameters.Add("@TerminalId", SqlDbType.VarChar);
				pTerminal.Value = (rq.TerminalId != null) ? (object)rq.TerminalId : DBNull.Value;

				SqlParameter pMerchant = cmd.Parameters.Add("@MerchantId", SqlDbType.VarChar);
				pMerchant.Value = (rq.MerchantId != null) ? (object)rq.MerchantId : DBNull.Value;

				cmd.Parameters.Add("@SecuencialTransaccion", SqlDbType.Int).Value = rq.SecuencialTransaccion;
				cmd.Parameters.Add("@NumeroLote", SqlDbType.Int).Value = rq.NumeroLote;

				SqlParameter pAut = cmd.Parameters.Add("@NumeroAutorizacion", SqlDbType.VarChar);
				pAut.Value = (rq.NumeroAutorizacion != null) ? (object)rq.NumeroAutorizacion : DBNull.Value;

				SqlParameter pFechaTrx = cmd.Parameters.Add("@FechaTransaccion", SqlDbType.VarChar);
				pFechaTrx.Value = (rq.FechaTransaccion != null) ? (object)rq.FechaTransaccion : DBNull.Value;

				SqlParameter pHoraTrx = cmd.Parameters.Add("@HoraTransaccion", SqlDbType.VarChar);
				pHoraTrx.Value = (rq.HoraTransaccion != null) ? (object)rq.HoraTransaccion : DBNull.Value;

				SqlParameter pCodMsg = cmd.Parameters.Add("@CodigoRespuestaMensaje", SqlDbType.VarChar);
				pCodMsg.Value = (rq.CodigoRespuestaMensaje != null) ? (object)rq.CodigoRespuestaMensaje : DBNull.Value;

				SqlParameter pCodAut = cmd.Parameters.Add("@CodigoRespuestaAutorizador", SqlDbType.VarChar);
				pCodAut.Value = (rq.CodigoRespuestaAutorizador != null) ? (object)rq.CodigoRespuestaAutorizador : DBNull.Value;

				SqlParameter pMsg = cmd.Parameters.Add("@MensajeRespuesta", SqlDbType.VarChar);
				pMsg.Value = (rq.MensajeRespuesta != null) ? (object)rq.MensajeRespuesta : DBNull.Value;

				SqlParameter pRed = cmd.Parameters.Add("@CodigoRed", SqlDbType.VarChar);
				pRed.Value = (rq.CodigoRed != null) ? (object)rq.CodigoRed : DBNull.Value;

				SqlParameter pTarjeta = cmd.Parameters.Add("@NumeroTarjetaTruncado", SqlDbType.VarChar);
				pTarjeta.Value = (rq.NumeroTarjetaTruncado != null) ? (object)rq.NumeroTarjetaTruncado : DBNull.Value;

				SqlParameter pVenc = cmd.Parameters.Add("@FechaVencimiento", SqlDbType.VarChar);
				pVenc.Value = (rq.FechaVencimiento != null) ? (object)rq.FechaVencimiento : DBNull.Value;

				SqlParameter pNombre = cmd.Parameters.Add("@NombreTarjetahabiente", SqlDbType.VarChar);
				pNombre.Value = (rq.NombreTarjetahabiente != null) ? (object)rq.NombreTarjetahabiente : DBNull.Value;

				SqlParameter pTarjetaEnc = cmd.Parameters.Add("@NumeroTarjetaEncriptado", SqlDbType.VarChar);
				pTarjetaEnc.Value = (rq.NumeroTarjetaEncriptado != null) ? (object)rq.NumeroTarjetaEncriptado : DBNull.Value;

				cmd.Parameters.Add("@MontoFijo", SqlDbType.Decimal).Value = rq.MontoFijo;
				cmd.Parameters.Add("@ValorInteres", SqlDbType.Decimal).Value = rq.ValorInteres;

				SqlParameter pMsgPub = cmd.Parameters.Add("@MensajePublicidad", SqlDbType.VarChar);
				pMsgPub.Value = (rq.MensajePublicidad != null) ? (object)rq.MensajePublicidad : DBNull.Value;

				SqlParameter pBancoAdq = cmd.Parameters.Add("@CodigoBancoAdquiriente", SqlDbType.VarChar);
				pBancoAdq.Value = (rq.CodigoBancoAdquiriente != null) ? (object)rq.CodigoBancoAdquiriente : DBNull.Value;

				SqlParameter pBancoEmi = cmd.Parameters.Add("@NombreBancoEmisor", SqlDbType.VarChar);
				pBancoEmi.Value = (rq.NombreBancoEmisor != null) ? (object)rq.NombreBancoEmisor : DBNull.Value;

				SqlParameter pAID = cmd.Parameters.Add("@AID", SqlDbType.VarChar);
				pAID.Value = (rq.AID != null) ? (object)rq.AID : DBNull.Value;

				SqlParameter pApp = cmd.Parameters.Add("@ApplicationLabel", SqlDbType.VarChar);
				pApp.Value = (rq.ApplicationLabel != null) ? (object)rq.ApplicationLabel : DBNull.Value;

				SqlParameter pCripto = cmd.Parameters.Add("@Criptograma", SqlDbType.VarChar);
				pCripto.Value = (rq.Criptograma != null) ? (object)rq.Criptograma : DBNull.Value;

				SqlParameter pARQC = cmd.Parameters.Add("@ARQC", SqlDbType.VarChar);
				pARQC.Value = (rq.ARQC != null) ? (object)rq.ARQC : DBNull.Value;

				SqlParameter pTVR = cmd.Parameters.Add("@TVR", SqlDbType.VarChar);
				pTVR.Value = (rq.TVR != null) ? (object)rq.TVR : DBNull.Value;

				SqlParameter pTSI = cmd.Parameters.Add("@TSI", SqlDbType.VarChar);
				pTSI.Value = (rq.TSI != null) ? (object)rq.TSI : DBNull.Value;

				SqlParameter pGrupo = cmd.Parameters.Add("@GrupoTarjeta", SqlDbType.VarChar);
				pGrupo.Value = (rq.GrupoTarjeta != null) ? (object)rq.GrupoTarjeta : DBNull.Value;

				SqlParameter pModo = cmd.Parameters.Add("@ModoLectura", SqlDbType.VarChar);
				pModo.Value = (rq.ModoLectura != null) ? (object)rq.ModoLectura : DBNull.Value;

				SqlParameter pPin = cmd.Parameters.Add("@VerificacionPin", SqlDbType.VarChar);
				pPin.Value = (rq.VerificacionPin != null) ? (object)rq.VerificacionPin : DBNull.Value;

				SqlParameter pFiller = cmd.Parameters.Add("@Filler", SqlDbType.VarChar);
				pFiller.Value = (rq.Filler != null) ? (object)rq.Filler : DBNull.Value;

				SqlParameter pTrama = cmd.Parameters.Add("@TramaRecibida", SqlDbType.VarChar);
				pTrama.Value = (rq.TramaRecibida != null) ? (object)rq.TramaRecibida : DBNull.Value;

				cmd.Parameters.Add("@FechaRespuesta", SqlDbType.DateTime).Value = rq.FechaRespuesta;
				cmd.Parameters.Add("@Success", SqlDbType.Bit).Value = rq.Success;

				SqlParameter pError = cmd.Parameters.Add("@ErrorMessage", SqlDbType.VarChar);
				pError.Value = (rq.ErrorMessage != null) ? (object)rq.ErrorMessage : DBNull.Value;

				// Ejecutar UPDATE
				cmd.ExecuteNonQuery();
			}
			finally
			{
				if (conn != null)
				{
					if (conn.State == ConnectionState.Open)
						conn.Close();

					conn.Dispose();
				}
			}
		}

		public void ActualizarRespuestaCT(ActualizaConsultaTarjetaRequest ct)
		{
			SqlConnection cn = new SqlConnection(_cn);
			cn.Open();

			try
			{
				ArrayList setList = new ArrayList();

				// ===== CAMPOS BASE =====
				setList.Add("TipoMensaje = @TipoMensaje");
				setList.Add("CodigoRespuesta = @CodigoRespuesta");
				setList.Add("NumeroTarjeta = @NumeroTarjeta");
				setList.Add("Bin = @Bin");
				setList.Add("FechaVencimiento = @FechaVencimiento");
				setList.Add("MensajeRespuesta = @MensajeRespuesta");

				// ===== OPCIONALES =====
				if (ct.Filler != null && ct.Filler.Length > 0)
					setList.Add("Filler = @Filler");

				if (ct.TramaRecibida != null && ct.TramaRecibida.Length > 0)
					setList.Add("TramaRecibida = @TramaRecibida");

				if (ct.FechaRespuesta != DateTime.MinValue)
					setList.Add("FechaRespuesta = @FechaRespuesta");

				string setClause = string.Join(", ", (string[])setList.ToArray(typeof(string)));

				SqlCommand cmd = new SqlCommand(
					"UPDATE Pinpad_LecturaConsultaTarjeta SET " + setClause +
					" WHERE idPinpad_LecturaConsultaTarjeta = @idPinpad_LecturaConsultaTarjeta", cn);

				// ===== PARÁMETROS =====
				cmd.Parameters.Add(new SqlParameter("@TipoMensaje", SqlDbType.Char, 2)).Value = ct.TipoMensaje;
				cmd.Parameters.Add(new SqlParameter("@CodigoRespuesta", SqlDbType.Char, 2)).Value = ct.CodigoRespuesta;
				cmd.Parameters.Add(new SqlParameter("@NumeroTarjeta", SqlDbType.VarChar, 25)).Value = ct.NumeroTarjeta;
				cmd.Parameters.Add(new SqlParameter("@Bin", SqlDbType.Char, 8)).Value = ct.Bin;
				cmd.Parameters.Add(new SqlParameter("@FechaVencimiento", SqlDbType.Char, 4)).Value = ct.FechaVencimiento;
				cmd.Parameters.Add(new SqlParameter("@MensajeRespuesta", SqlDbType.VarChar, 50)).Value = ct.MensajeRespuesta;

				if (ct.Filler != null && ct.Filler.Length > 0)
					cmd.Parameters.Add(new SqlParameter("@Filler", SqlDbType.VarChar, 27)).Value = ct.Filler;

				if (ct.TramaRecibida != null && ct.TramaRecibida.Length > 0)
					cmd.Parameters.Add(new SqlParameter("@TramaRecibida", SqlDbType.VarChar, 4000)).Value = ct.TramaRecibida;

				if (ct.FechaRespuesta != DateTime.MinValue)
					cmd.Parameters.Add(new SqlParameter("@FechaRespuesta", SqlDbType.DateTime)).Value = ct.FechaRespuesta;

				cmd.Parameters.Add(new SqlParameter("@idPinpad_LecturaConsultaTarjeta", SqlDbType.UniqueIdentifier))
					.Value = ct.idPinpad_LecturaConsultaTarjeta;

				cmd.ExecuteNonQuery();
			}
			finally
			{
				cn.Close();
			}
		}
	}
}
