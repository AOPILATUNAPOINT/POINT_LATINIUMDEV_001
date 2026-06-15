using System;
using System.Data;

namespace Latinium.Models
{
	/// <summary>
	/// Estado actual de la verificación Family & Friends (compatible con .NET Framework 1.1)
	/// </summary>
	public class FamilyFriendsEstadoModel
	{
		// Campos privados
		private bool esFF;
		private int idCliente;
		private Guid idWEB_FF_Invitado;
		private bool rechazoPromocion;
		private int idTipoPrecio;
		private string mensajeEstadoFF;
		private bool existeRegistro;

		// Propiedades públicas
		public bool EsFF
		{
			get { return esFF; }
			set { esFF = value; }
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

		public bool RechazoPromocion
		{
			get { return rechazoPromocion; }
			set { rechazoPromocion = value; }
		}

		public int IdTipoPrecio
		{
			get { return idTipoPrecio; }
			set { idTipoPrecio = value; }
		}

		public string MensajeEstadoFF
		{
			get { return mensajeEstadoFF; }
			set { mensajeEstadoFF = value; }
		}

		public bool ExisteRegistro
		{
			get { return existeRegistro; }
			set { existeRegistro = value; }
		}

		// Constructor
		public FamilyFriendsEstadoModel()
		{
			esFF = false;
			idCliente = 0;
			idWEB_FF_Invitado = Guid.Empty;
			rechazoPromocion = false;
			idTipoPrecio = 0;
			mensajeEstadoFF = string.Empty;
			existeRegistro = false;
		}

		public void Load(IDataReader dr)
		{
			int index;

			index = dr.GetOrdinal("EsFF");
			if (!dr.IsDBNull(index))
				esFF = dr.GetBoolean(index);

			index = dr.GetOrdinal("IdCliente");
			if (!dr.IsDBNull(index))
				idCliente = dr.GetInt32(index);

			index = dr.GetOrdinal("IdWEB_FF_Invitado");
			if (!dr.IsDBNull(index))
				idWEB_FF_Invitado = dr.GetGuid(index);

			index = dr.GetOrdinal("RechazoPromocion");
			if (!dr.IsDBNull(index))
				rechazoPromocion = dr.GetBoolean(index);

			index = dr.GetOrdinal("IdTipoPrecio");
			if (!dr.IsDBNull(index))
				idTipoPrecio = dr.GetInt32(index);

			index = dr.GetOrdinal("MensajeEstadoFF");
			if (!dr.IsDBNull(index))
				mensajeEstadoFF = dr.GetString(index);

			//  Este campo NO viene de la BD, lo controla el service
			existeRegistro = true;
		}
	}
}
