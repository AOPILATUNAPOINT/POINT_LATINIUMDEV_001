using System;

namespace Latinium.Models.Requerimiento
{
	public class CentroCostoModel
	{
		private int _id;
		private string _nombre;

		public int Id
		{
			get { return _id; }
			set { _id = value; }
		}

		public string Nombre
		{
			get { return _nombre; }
			set { _nombre = value; }
		}
	}
}
