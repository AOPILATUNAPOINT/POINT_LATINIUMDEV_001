using System;
using System.Data;

namespace Latinium.Models
{
	/// <summary>
	/// Modelo simple de cabecera Family & Friends
	/// Compatible con .NET Framework 1.1
	/// </summary>
	public class FamilyFriendsCabeceraModel
	{
		public Guid IdCabecera;
		public string Titulo;

		public FamilyFriendsCabeceraModel()
		{
		}

		public static FamilyFriendsCabeceraModel Load(IDataReader reader)
		{
			FamilyFriendsCabeceraModel m =
				new FamilyFriendsCabeceraModel();

			m.IdCabecera =
				new Guid(reader["idWEB_FF_Cabecera"].ToString());

			m.Titulo =
				reader["Titulo"].ToString();

			return m;
		}
	}
}