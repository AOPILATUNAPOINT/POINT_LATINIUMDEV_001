using System;

namespace Latinium.Models.Requerimiento
{
	/// <summary>
	/// Modelo en memoria para selección de menú del requerimiento
	/// </summary>
	public class RequerimientoMenuSeguridadModel
	{
		private int _id_seg;
		private string _Modulo;
		private string _Pantalla;
		private string _SubPantalla;

			public int id_seg
			{
				get { return _id_seg; }
				set { _id_seg = value; }
			}

		public string Modulo
		{
			get { return _Modulo; }
			set { _Modulo = value; }
		}

		public string Pantalla
		{
			get { return _Pantalla; }
			set { _Pantalla = value; }
		}

		public string SubPantalla
		{
			get { return _SubPantalla; }
			set { _SubPantalla = value; }
		}
	}
}
