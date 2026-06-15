using System;

namespace Latinium.Models.FamilyFriends.Estado
{
	/// <summary>
	/// Descripción breve de FamilyFriendsEstadCombooModel.
	/// </summary>
	public class FamilyFriendsEstadCombooModel
	{
			private int IdEstado;
		private string nombreEstado;

		public int idEstado
		{
			get { return IdEstado; }
			set { IdEstado = value; }
		}

		public string NombreEstado
		{
			get { return nombreEstado; }
			set { nombreEstado = value; }
		}
	}
}
