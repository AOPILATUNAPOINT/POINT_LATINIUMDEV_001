using System;

namespace Latinium.Models.Pinpad
{
	public class ProcesoPagoRequest
	{
		// ===== CAMPOS =====

		private Guid _idPinpad_ProcesoPago;
		private int _idCompra;

		private string tipoMensaje;
		private string tipoTransaccion;
		private string codigoRed;
		private string codigoDiferido;

		private int plazoDiferido;
		private int mesesGracia;

		private string numeroCaja;

		private decimal montoTotal;
		private decimal baseIVA;
		private decimal base0;
		private decimal iva;
		private decimal propina;

		private string trama;
		private string talla;

		private string secuencialTransaccion;
		private string horaTransaccion;
		private string fechaTransaccion;
		private string numeroAutorizacion;

		private string mid;
		private string tid;
		private string cid;

		private DateTime fechaSistema;
		private string usuario;
		private string estacion;

		// ===== PROPIEDADES =====

		public Guid idPinpad_ProcesoPago
		{
			get { return _idPinpad_ProcesoPago; }
			set { _idPinpad_ProcesoPago = value; }
		}

		public int idCompra
		{
			get { return _idCompra; }
			set { _idCompra = value; }
		}

		public string TipoMensaje
		{
			get { return tipoMensaje; }
			set { tipoMensaje = value; }
		}

		public string TipoTransaccion
		{
			get { return tipoTransaccion; }
			set { tipoTransaccion = value; }
		}

		public string CodigoRed
		{
			get { return codigoRed; }
			set { codigoRed = value; }
		}

		public string CodigoDiferido
		{
			get { return codigoDiferido; }
			set { codigoDiferido = value; }
		}

		public int PlazoDiferido
		{
			get { return plazoDiferido; }
			set { plazoDiferido = value; }
		}

		public int MesesGracia
		{
			get { return mesesGracia; }
			set { mesesGracia = value; }
		}

		public string NumeroCaja
		{
			get { return numeroCaja; }
			set { numeroCaja = value; }
		}

		public decimal MontoTotal
		{
			get { return montoTotal; }
			set { montoTotal = value; }
		}

		public decimal BaseIVA
		{
			get { return baseIVA; }
			set { baseIVA = value; }
		}

		public decimal Base0
		{
			get { return base0; }
			set { base0 = value; }
		}

		public decimal IVA
		{
			get { return iva; }
			set { iva = value; }
		}

		public decimal Propina
		{
			get { return propina; }
			set { propina = value; }
		}

		public string Trama
		{
			get { return trama; }
			set { trama = value; }
		}

		public string Talla
		{
			get { return talla; }
			set { talla = value; }
		}

		public string SecuencialTransaccion
		{
			get { return secuencialTransaccion; }
			set { secuencialTransaccion = value; }
		}

		public string HoraTransaccion
		{
			get { return horaTransaccion; }
			set { horaTransaccion = value; }
		}

		public string FechaTransaccion
		{
			get { return fechaTransaccion; }
			set { fechaTransaccion = value; }
		}

		public string NumeroAutorizacion
		{
			get { return numeroAutorizacion; }
			set { numeroAutorizacion = value; }
		}

		public string MID
		{
			get { return mid; }
			set { mid = value; }
		}

		public string TID
		{
			get { return tid; }
			set { tid = value; }
		}

		public string CID
		{
			get { return cid; }
			set { cid = value; }
		}

		public DateTime FechaSistema
		{
			get { return fechaSistema; }
			set { fechaSistema = value; }
		}

		public string Usuario
		{
			get { return usuario; }
			set { usuario = value; }
		}

		public string Estacion
		{
			get { return estacion; }
			set { estacion = value; }
		}
	}
}