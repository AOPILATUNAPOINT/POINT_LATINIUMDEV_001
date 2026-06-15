using System;

namespace Latinium.Models
{
	/// <summary>
	/// Modelo de parámetros para el reporte Family & Friends V4
	/// Compatible con .NET Framework 1.1
	/// </summary>
	public class FamilyFriendsReporteFiltroModel
	{
		public Guid IdCabecera;
		public string Filtro;
		public string Area;

		// NUEVO PARÁMETRO
		public int idEstado;

		/// <summary>
		/// Constructor vacío
		/// </summary>
		public FamilyFriendsReporteFiltroModel()
		{
			this.idEstado = 0;
		}

		/// <summary>
		/// Constructor con parámetros
		/// </summary>
		public FamilyFriendsReporteFiltroModel(
			Guid idCabecera,
			string filtro,
			string area,
			int idEstado
			)
		{
			this.IdCabecera = idCabecera;
			this.Filtro = filtro;
			this.Area = area;
			this.idEstado = idEstado;
		}
	}
}