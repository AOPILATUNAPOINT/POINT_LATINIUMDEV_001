using System;

namespace Latinium.Models.FamilyFriends
{
	public class ArticuloPrecio
	{
		private int _idArticulo;
		private string _codigo;
		private decimal _precio8;
		private decimal _precio9;

		public int IdArticulo
		{
			get { return _idArticulo; }
			set { _idArticulo = value; }
		}

		public string Codigo
		{
			get { return _codigo; }
			set { _codigo = value; }
		}

		public decimal Precio8
		{
			get { return _precio8; }
			set { _precio8 = value; }
		}

		public decimal Precio9
		{
			get { return _precio9; }
			set { _precio9 = value; }
		}
	}
}