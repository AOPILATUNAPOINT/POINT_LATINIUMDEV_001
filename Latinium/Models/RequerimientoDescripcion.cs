using System;

namespace Latinium.Models
{
	public class RequerimientoDescripcion
	{
		// Campos privados para almacenar los valores
		private int _idRequerimientoDescripcion;
		private string _descripcion;

		// Propiedad para IdRequerimientoDescripcion
		public int IdRequerimientoDescripcion
		{
			get { return _idRequerimientoDescripcion; }
			set { _idRequerimientoDescripcion = value; }
		}

		// Propiedad para Descripcion
		public string Descripcion
		{
			get { return _descripcion; }
			set { _descripcion = value; }
		}
	}
}
