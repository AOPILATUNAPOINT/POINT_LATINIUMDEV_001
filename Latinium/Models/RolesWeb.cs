using System;

namespace Latinium.Models
{
	/// <summary>
	/// Descripción breve de RolesWeb.
	/// </summary>
	public class RolesWeb
	{
		private int idRolesWeb;
		private string nombre;
		private DateTime fechaSistema;
		private string estacion;
		private string usuario;

		public int IdRolesWeb
		{
			get { return IdRolesWeb; }
			set { IdRolesWeb = value; }
		}

		public string Nombre
		{
			get { return nombre; }
			set { nombre = value; }
		}

		public DateTime FechaSistema
		{
			get { return fechaSistema; }
			set { fechaSistema = value; }
		}

		public string Estacion
		{
			get { return estacion; }
			set { estacion = value; }
		}

		public string Usuario
		{
			get { return usuario; }
			set { usuario = value; }
		}

		public RolesWeb()
		{
			// Constructor vacío
		}
	}
}
