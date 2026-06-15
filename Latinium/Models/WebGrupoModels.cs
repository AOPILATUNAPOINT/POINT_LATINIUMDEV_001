using System;
using System.Data.SqlClient;

namespace Latinium.Models
{
	/// <summary>
	/// Modelo para la entidad WEB_Grupos
	/// </summary>
	[Serializable]
	public class WebGrupoModels
	{
		private int _idWEB_Grupos;
		private string _descripcion;

		public int IdWEB_Grupos
		{
			get { return _idWEB_Grupos; }
			set { _idWEB_Grupos = value; }
		}

		public string Descripcion
		{
			get { return _descripcion; }
			set { _descripcion = value; }
		}

		public WebGrupoModels()
		{
			_descripcion = string.Empty;
		}

		/// <summary>
		/// Carga el modelo desde un SqlDataReader
		/// </summary>
		public void Cargar(SqlDataReader dr)
		{
			if (!dr.IsDBNull(dr.GetOrdinal("idWEB_Grupos")))
				_idWEB_Grupos = dr.GetInt32(dr.GetOrdinal("idWEB_Grupos"));

			if (!dr.IsDBNull(dr.GetOrdinal("Descripcion")))
				_descripcion = dr.GetString(dr.GetOrdinal("Descripcion"));
		}
	}
}
