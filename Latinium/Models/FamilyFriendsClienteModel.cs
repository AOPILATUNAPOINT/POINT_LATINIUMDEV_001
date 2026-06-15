using System;
using System.Data;

namespace Latinium.Models
{
	/// <summary>
	/// Modelo para información del cliente Family & Friends (compatible con .NET Framework 1.1)
	/// </summary>
	public class FamilyFriendsClienteModel
	{
		// Campos privados
		private string cedula;
		private int idCliente;
		private bool existeInvitado;
		private DateTime fechaInicioEvento;
		private DateTime fechaFinEvento;
		private int facturasRegistradas;
		private bool puedeVerPrecio;
		private string mensaje;

		private int idFormaPago;
		private int tipoInvitado;
		private int idTipoPrecio; 

		// *** NUEVO ***
		private Guid idWEB_FF_Invitado;

		// Constructor
		public FamilyFriendsClienteModel()
		{
			this.cedula = string.Empty;
			this.idCliente = 0;
			this.existeInvitado = false;
			this.fechaInicioEvento = DateTime.MinValue;
			this.fechaFinEvento = DateTime.MinValue;
			this.facturasRegistradas = 0;
			this.puedeVerPrecio = false;
			this.mensaje = string.Empty;

			this.idFormaPago = 0;
			this.tipoInvitado = 0;
			this.idTipoPrecio = 0;

			// Inicialización del nuevo GUID
			this.idWEB_FF_Invitado = Guid.Empty;
		}

		// Propiedades
		public string Cedula
		{
			get { return cedula; }
			set { cedula = value; }
		}

		public int IdCliente
		{
			get { return idCliente; }
			set { idCliente = value; }
		}

		public bool ExisteInvitado
		{
			get { return existeInvitado; }
			set { existeInvitado = value; }
		}

		public DateTime FechaInicioEvento
		{
			get { return fechaInicioEvento; }
			set { fechaInicioEvento = value; }
		}

		public DateTime FechaFinEvento
		{
			get { return fechaFinEvento; }
			set { fechaFinEvento = value; }
		}

		public int FacturasRegistradas
		{
			get { return facturasRegistradas; }
			set { facturasRegistradas = value; }
		}

		public bool PuedeVerPrecio
		{
			get { return puedeVerPrecio; }
			set { puedeVerPrecio = value; }
		}

		public string Mensaje
		{
			get { return mensaje; }
			set { mensaje = value; }
		}

		public int IdFormaPago
		{
			get { return idFormaPago; }
			set { idFormaPago = value; }
		}

		public int TipoInvitado
		{
			get { return tipoInvitado; }
			set { tipoInvitado = value; }
		}

		public int IdTipoPrecio
		{
			get { return idTipoPrecio; }
			set { idTipoPrecio = value; }
		}

		// *** NUEVA PROPIEDAD ***
		public Guid IdWEB_FF_Invitado
		{
			get { return idWEB_FF_Invitado; }
			set { idWEB_FF_Invitado = value; }
		}

		// Cargar desde DataReader
		public static FamilyFriendsClienteModel Load(IDataReader reader)
		{
			FamilyFriendsClienteModel cliente = new FamilyFriendsClienteModel();

			if (reader["Cedula"] != DBNull.Value)
				cliente.Cedula = Convert.ToString(reader["Cedula"]);

			if (reader["idCliente"] != DBNull.Value)
				cliente.IdCliente = Convert.ToInt32(reader["idCliente"]);

			if (reader["ExisteInvitado"] != DBNull.Value)
				cliente.ExisteInvitado = Convert.ToBoolean(reader["ExisteInvitado"]);

			if (reader["FechaInicioEvento"] != DBNull.Value)
				cliente.FechaInicioEvento = Convert.ToDateTime(reader["FechaInicioEvento"]);

			if (reader["FechaFinEvento"] != DBNull.Value)
				cliente.FechaFinEvento = Convert.ToDateTime(reader["FechaFinEvento"]);

			if (reader["FacturasRegistradas"] != DBNull.Value)
				cliente.FacturasRegistradas = Convert.ToInt32(reader["FacturasRegistradas"]);

			if (reader["PuedeVerPrecio"] != DBNull.Value)
				cliente.PuedeVerPrecio = Convert.ToBoolean(reader["PuedeVerPrecio"]);

			if (reader["Mensaje"] != DBNull.Value)
				cliente.Mensaje = Convert.ToString(reader["Mensaje"]);

			if (reader["idFormaPago"] != DBNull.Value)
				cliente.IdFormaPago = Convert.ToInt32(reader["idFormaPago"]);

			if (reader["TipoInvitado"] != DBNull.Value)
				cliente.TipoInvitado = Convert.ToInt32(reader["TipoInvitado"]);

			if (reader["IdTipoPrecio"] != DBNull.Value)
				cliente.IdTipoPrecio = Convert.ToInt32(reader["IdTipoPrecio"]);

			// *** NUEVO GUID ***
			if (reader["idWEB_FF_Invitado"] != DBNull.Value)
				cliente.IdWEB_FF_Invitado = (Guid)reader["idWEB_FF_Invitado"];

			return cliente;
		}
	}
}
