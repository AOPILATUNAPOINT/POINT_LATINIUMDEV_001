using System;
using System.Data;  // Esta directiva es necesaria para IDataReader

namespace Latinium.Models
{
	/// <summary>
	/// Model for RequerimientoUsuario table
	/// </summary>
	public class RequerimientoUsuario
	{
		// Private fields (lowercase)
		private int idUsuario;
		private int idGrupo;
		private string nombreCompleto;
		private bool activo;

		// Constructors
		public RequerimientoUsuario()
		{
		}

		public RequerimientoUsuario(int idUsuario, int idGrupo, string nombreCompleto, bool activo)
		{
			this.idUsuario = idUsuario;
			this.idGrupo = idGrupo;
			this.nombreCompleto = nombreCompleto;
			this.activo = activo;
		}

		// Properties
		public int IdUsuario
		{
			get { return idUsuario; }
			set { idUsuario = value; }
		}

		public int IdGrupo
		{
			get { return idGrupo; }
			set { idGrupo = value; }
		}

		public string NombreCompleto
		{
			get { return nombreCompleto; }
			set { nombreCompleto = value; }
		}

		public bool Activo
		{
			get { return activo; }
			set { activo = value; }
		}

		// Data Access Method
		public static RequerimientoUsuario Load(IDataReader reader)
		{
			RequerimientoUsuario usuario = new RequerimientoUsuario();
            
			usuario.IdUsuario = Convert.ToInt32(reader["idUsuario"]);
            
			if (reader["idGrupo"] != DBNull.Value)
				usuario.IdGrupo = Convert.ToInt32(reader["idGrupo"]);
            
			if (reader["NombreCompleto"] != DBNull.Value)
				usuario.NombreCompleto = Convert.ToString(reader["NombreCompleto"]);
            
			if (reader["Activo"] != DBNull.Value)
				usuario.Activo = Convert.ToBoolean(reader["Activo"]);
            
			return usuario;
		}
	}
}
