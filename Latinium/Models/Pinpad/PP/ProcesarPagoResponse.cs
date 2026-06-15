using System;

namespace Latinium.Models.Pinpad.PP
{
	/// <summary>
	/// Respuesta base genérica
	/// </summary>
	public class BaseResponse
	{
		private bool _success;
		private string _dataAscii;
		private string _errorMessage;

		public bool Success
		{
			get { return _success; }
			set { _success = value; }
		}

		public string DataAscii
		{
			get { return _dataAscii; }
			set { _dataAscii = value; }
		}

		public string ErrorMessage
		{
			get { return _errorMessage; }
			set { _errorMessage = value; }
		}
	}

	/// <summary>
	/// Respuesta del proceso de pago Pinpad (alineada al detalle)
	/// </summary>
	public class ProcesarPagoResponse : BaseResponse
	{
		// ===== HEADER =====
		private string _tipoMensaje;
		private string _codigoRespuestaMensaje;

		// ===== RESPUESTA =====
		private string _codigoRed;
		private string _codigoRespuestaAutorizador;
		private string _mensajeRespuesta;

		// ===== TRANSACCIÓN =====
		private int _secuencialTransaccion;
		private int _numeroLote;
		private string _horaTransaccion;
		private string _fechaTransaccion;
		private string _numeroAutorizacion;

		// ===== TERMINAL / COMERCIO =====
		private string _terminalId;
		private string _merchantId;

		// ===== FINANCIAMIENTO =====
		private decimal _valorInteres;
		private string _mensajePublicidad;

		// ===== BANCOS =====
		private string _codigoBancoAdquiriente;
		private string _nombreBancoEmisor;

		// ===== TARJETA =====
		private string _grupoTarjeta;
		private string _modoLectura;
		private string _nombreTarjetahabiente;

		// ===== GASOLINERAS / OTROS =====
		private decimal _montoFijo;

		// ===== EMV =====
		private string _applicationLabel;
		private string _aid;
		private string _criptograma;
		private string _verificacionPin;
		private string _arqc;
		private string _tvr;
		private string _tsi;

		// ===== TARJETA DATOS =====
		private string _numeroTarjetaTruncado;
		private string _fechaVencimiento;
		private string _numeroTarjetaEncriptado;

		// ===== OTROS =====
		private string _filler;

		// ================= PROPIEDADES =================

		public string TipoMensaje
		{
			get { return _tipoMensaje; }
			set { _tipoMensaje = value; }
		}

		public string CodigoRespuestaMensaje
		{
			get { return _codigoRespuestaMensaje; }
			set { _codigoRespuestaMensaje = value; }
		}

		public string CodigoRed
		{
			get { return _codigoRed; }
			set { _codigoRed = value; }
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

		public string HoraTransaccion
		{
			get { return _horaTransaccion; }
			set { _horaTransaccion = value; }
		}

		public string FechaTransaccion
		{
			get { return _fechaTransaccion; }
			set { _fechaTransaccion = value; }
		}

		public string NumeroAutorizacion
		{
			get { return _numeroAutorizacion; }
			set { _numeroAutorizacion = value; }
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

		public string NombreTarjetahabiente
		{
			get { return _nombreTarjetahabiente; }
			set { _nombreTarjetahabiente = value; }
		}

		public decimal MontoFijo
		{
			get { return _montoFijo; }
			set { _montoFijo = value; }
		}

		public string ApplicationLabel
		{
			get { return _applicationLabel; }
			set { _applicationLabel = value; }
		}

		public string AID
		{
			get { return _aid; }
			set { _aid = value; }
		}

		public string Criptograma
		{
			get { return _criptograma; }
			set { _criptograma = value; }
		}

		public string VerificacionPin
		{
			get { return _verificacionPin; }
			set { _verificacionPin = value; }
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

		public string NumeroTarjetaEncriptado
		{
			get { return _numeroTarjetaEncriptado; }
			set { _numeroTarjetaEncriptado = value; }
		}

		public string Filler
		{
			get { return _filler; }
			set { _filler = value; }
		}
	}
}