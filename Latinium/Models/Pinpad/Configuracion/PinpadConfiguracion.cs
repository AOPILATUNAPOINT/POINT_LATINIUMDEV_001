using System;

namespace Latinium.Models.Pinpad.Configuracion
{
	/// <summary>
	/// Modelo actualizado según tabla Pinpad_Configuracion
	/// </summary>
	public class PinpadConfiguracion
	{
		private Guid _idPinpadConfiguracion;

		private string _deviceId;
		private string _nombrePC;

		private int _bodega;
		private string _numeroCaja;

		private string _ipPinPad;
		private int _puertoPinPad;
		private int _timeoutPinPad;

		private string _mid;
		private string _tid;
		private string _cid;

		private bool _activo;

		private DateTime _fechaSistema;
		private DateTime _fechaModificacion;

		private string _usuario;
		private string _usuarioModificacion;

		private string _nombreBodega;

		// =========================
		// PROPIEDADES
		// =========================

		public string NombreBodega
		{
			get { return _nombreBodega; }
			set { _nombreBodega = value; }
		}

		public Guid IdPinpadConfiguracion
		{
			get { return _idPinpadConfiguracion; }
			set { _idPinpadConfiguracion = value; }
		}

		public string DeviceId
		{
			get { return _deviceId; }
			set { _deviceId = value; }
		}

		public string NombrePC
		{
			get { return _nombrePC; }
			set { _nombrePC = value; }
		}

		public int Bodega
		{
			get { return _bodega; }
			set { _bodega = value; }
		}

		public string NumeroCaja
		{
			get { return _numeroCaja; }
			set { _numeroCaja = value; }
		}

		public string IpPinPad
		{
			get { return _ipPinPad; }
			set { _ipPinPad = value; }
		}

		public int PuertoPinPad
		{
			get { return _puertoPinPad; }
			set { _puertoPinPad = value; }
		}

		public int TimeoutPinPad
		{
			get { return _timeoutPinPad; }
			set { _timeoutPinPad = value; }
		}

		public string MID
		{
			get { return _mid; }
			set { _mid = value; }
		}

		public string TID
		{
			get { return _tid; }
			set { _tid = value; }
		}

		public string CID
		{
			get { return _cid; }
			set { _cid = value; }
		}

		public bool Activo
		{
			get { return _activo; }
			set { _activo = value; }
		}

		public DateTime FechaSistema
		{
			get { return _fechaSistema; }
			set { _fechaSistema = value; }
		}

		public DateTime FechaModificacion
		{
			get { return _fechaModificacion; }
			set { _fechaModificacion = value; }
		}

		public string Usuario
		{
			get { return _usuario; }
			set { _usuario = value; }
		}

		public string UsuarioModificacion
		{
			get { return _usuarioModificacion; }
			set { _usuarioModificacion = value; }
		}
	}
}