using System;

namespace Latinium.Models
{
	/// <summary>
	/// Modelo de verificación Family & Friends (compatible con .NET Framework 1.1)
	/// </summary>
	public class FamilyFriendsVerificacionModel
	{
		// Campos privados
		private bool esFF;
		private int idTipoPrecio;
		private string mensajeSP;
		private int idCliente;
		private Guid idWEB_FF_Invitado;

		// NUEVO: campo privado para rechazo
		private bool rechazoPromocion;

		// Propiedades públicas
		public bool EsFF
		{
			get { return esFF; }
			set { esFF = value; }
		}

		public int IdTipoPrecio
		{
			get { return idTipoPrecio; }
			set { idTipoPrecio = value; }
		}

		public string MensajeSP
		{
			get { return mensajeSP; }
			set { mensajeSP = value; }
		}

		public int IdCliente
		{
			get { return idCliente; }
			set { idCliente = value; }
		}

		public Guid IdWEB_FF_Invitado
		{
			get { return idWEB_FF_Invitado; }
			set { idWEB_FF_Invitado = value; }
		}

		// NUEVA PROPIEDAD COMPLETA
		public bool RechazoPromocion
		{
			get { return rechazoPromocion; }
			set { rechazoPromocion = value; }
		}

		// Constructor
		public FamilyFriendsVerificacionModel()
		{
			esFF = false;
			idTipoPrecio = 0;
			idCliente = 0;
			mensajeSP = "";
			idWEB_FF_Invitado = Guid.Empty;
			rechazoPromocion = false;  // por defecto no rechazó
		}
	}
}
