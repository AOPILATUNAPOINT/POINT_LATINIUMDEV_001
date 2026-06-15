using System;

namespace Latinium.Models.Pinpad.PP
{
	/// <summary>
	/// Datos del procesamiento de pago que se guardarán en base de datos
	/// </summary>
	public class ActualizarProcesarPagoRequest
	{
		// =========================
		// CAMPOS PRIVADOS
		// =========================

		private Guid _idPinpad_ProcesoPago;

		private string _tipoMensaje;
		private string _terminalId;
		private string _merchantId;

		private int _secuencialTransaccion;
		private int _numeroLote;
		private string _numeroAutorizacion;

		private string _fechaTransaccion;
		private string _horaTransaccion;

		private string _codigoRespuestaMensaje;
		private string _codigoRespuestaAutorizador;
		private string _mensajeRespuesta;

		private string _codigoRed;

		private string _numeroTarjetaTruncado;
		private string _fechaVencimiento;
		private string _nombreTarjetahabiente;
		private string _numeroTarjetaEncriptado;

		private decimal _montoFijo;
		private decimal _valorInteres;

		private string _mensajePublicidad;
		private string _codigoBancoAdquiriente;
		private string _nombreBancoEmisor;

		private string _aid;
		private string _applicationLabel;
		private string _criptograma;
		private string _arqc;
		private string _tvr;
		private string _tsi;

		private string _grupoTarjeta;
		private string _modoLectura;
		private string _verificacionPin;

		private string _filler;
		private string _tramaRecibida;

		private DateTime _fechaRespuesta;

		private bool _success;
		private string _errorMessage;

		// =========================
		// PROPIEDADES
		// =========================

		public Guid idPinpad_ProcesoPago
		{
			get { return _idPinpad_ProcesoPago; }
			set { _idPinpad_ProcesoPago = value; }
		}

		public string TipoMensaje
		{
			get { return _tipoMensaje; }
			set { _tipoMensaje = value; }
		}

		public string TerminalId
		{
			get { return _terminalId; }
			set { _terminalId = value; }
		}

		public string MerchantId
		{
			get { return _merchantId; }
			set { _merchantId = value; }
		}

		public int SecuencialTransaccion
		{
			get { return _secuencialTransaccion; }
			set { _secuencialTransaccion = value; }
		}

		public int NumeroLote
		{
			get { return _numeroLote; }
			set { _numeroLote = value; }
		}

		public string NumeroAutorizacion
		{
			get { return _numeroAutorizacion; }
			set { _numeroAutorizacion = value; }
		}

		public string FechaTransaccion
		{
			get { return _fechaTransaccion; }
			set { _fechaTransaccion = value; }
		}

		public string HoraTransaccion
		{
			get { return _horaTransaccion; }
			set { _horaTransaccion = value; }
		}

		public string CodigoRespuestaMensaje
		{
			get { return _codigoRespuestaMensaje; }
			set { _codigoRespuestaMensaje = value; }
		}

		public string CodigoRespuestaAutorizador
		{
			get { return _codigoRespuestaAutorizador; }
			set { _codigoRespuestaAutorizador = value; }
		}

		public string MensajeRespuesta
		{
			get { return _mensajeRespuesta; }
			set { _mensajeRespuesta = value; }
		}

		public string CodigoRed
		{
			get { return _codigoRed; }
			set { _codigoRed = value; }
		}

		public string NumeroTarjetaTruncado
		{
			get { return _numeroTarjetaTruncado; }
			set { _numeroTarjetaTruncado = value; }
		}

		public string FechaVencimiento
		{
			get { return _fechaVencimiento; }
			set { _fechaVencimiento = value; }
		}

		public string NombreTarjetahabiente
		{
			get { return _nombreTarjetahabiente; }
			set { _nombreTarjetahabiente = value; }
		}

		public string NumeroTarjetaEncriptado
		{
			get { return _numeroTarjetaEncriptado; }
			set { _numeroTarjetaEncriptado = value; }
		}

		public decimal MontoFijo
		{
			get { return _montoFijo; }
			set { _montoFijo = value; }
		}

		public decimal ValorInteres
		{
			get { return _valorInteres; }
			set { _valorInteres = value; }
		}

		public string MensajePublicidad
		{
			get { return _mensajePublicidad; }
			set { _mensajePublicidad = value; }
		}

		public string CodigoBancoAdquiriente
		{
			get { return _codigoBancoAdquiriente; }
			set { _codigoBancoAdquiriente = value; }
		}

		public string NombreBancoEmisor
		{
			get { return _nombreBancoEmisor; }
			set { _nombreBancoEmisor = value; }
		}

		public string AID
		{
			get { return _aid; }
			set { _aid = value; }
		}

		public string ApplicationLabel
		{
			get { return _applicationLabel; }
			set { _applicationLabel = value; }
		}

		public string Criptograma
		{
			get { return _criptograma; }
			set { _criptograma = value; }
		}

		public string ARQC
		{
			get { return _arqc; }
			set { _arqc = value; }
		}

		public string TVR
		{
			get { return _tvr; }
			set { _tvr = value; }
		}

		public string TSI
		{
			get { return _tsi; }
			set { _tsi = value; }
		}

		public string GrupoTarjeta
		{
			get { return _grupoTarjeta; }
			set { _grupoTarjeta = value; }
		}

		public string ModoLectura
		{
			get { return _modoLectura; }
			set { _modoLectura = value; }
		}

		public string VerificacionPin
		{
			get { return _verificacionPin; }
			set { _verificacionPin = value; }
		}

		public string Filler
		{
			get { return _filler; }
			set { _filler = value; }
		}

		public string TramaRecibida
		{
			get { return _tramaRecibida; }
			set { _tramaRecibida = value; }
		}

		public DateTime FechaRespuesta
		{
			get { return _fechaRespuesta; }
			set { _fechaRespuesta = value; }
		}

		public bool Success
		{
			get { return _success; }
			set { _success = value; }
		}

		public string ErrorMessage
		{
			get { return _errorMessage; }
			set { _errorMessage = value; }
		}
	}
}