using System;
using System.Data;

namespace Latinium.Models
{
	/// <summary>
	/// Modelo para enviar datos al procedimiento WEB_FF_CompraGuardar (compatible con .NET 1.1)
	/// </summary>
	public class FamilyFriendsCompraModel
	{
		// ================================
		// Campos privados
		// ================================
		private int idCompra;
		private Guid idWEB_FF_Invitado;
		private int estado;
		private int idTipoPrecio;
		private int idFormaPago;
		private int idCliente;
		private int idNotaCredito;

		// ================================
		// Propiedades públicas (solo las que recibe el SP)
		// ================================
		public int IdCompra
		{
			get { return idCompra; }
			set { idCompra = value; }
		}

		public Guid IdWEB_FF_Invitado
		{
			get { return idWEB_FF_Invitado; }
			set { idWEB_FF_Invitado = value; }
		}

		public int Estado
		{
			get { return estado; }
			set { estado = value; }
		}

		public int IdTipoPrecio
		{
			get { return idTipoPrecio; }
			set { idTipoPrecio = value; }
		}

		public int IdFormaPago
		{
			get { return idFormaPago; }
			set { idFormaPago = value; }
		}

		public int IdCliente
		{
			get { return idCliente; }
			set { idCliente = value; }
		}

		public int IdNotaCredito
		{
			get { return idNotaCredito; }
			set { idNotaCredito = value; }
		}

		// ================================
		// Constructor
		// ================================
		public FamilyFriendsCompraModel()
		{
			idCompra = 0;
			idWEB_FF_Invitado = Guid.Empty;
			estado = 1;
			idTipoPrecio = 0;
			idFormaPago = 0;
			idCliente = 0;
			idNotaCredito = 0;
		}
	}
}
