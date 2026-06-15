using System;

namespace Latinium.Models.Requerimiento.TurnoSoporte
{
	/// <summary>
	/// Modelo de Turno de Soporte
	/// Compatible .NET 1.1
	/// </summary>
	public class RequerimientoTurnoSoporteModel
	{
		// ===== CAMPOS =====
		public int idRequerimientoTurnoSoporte;
		public int IdRequerimientoPerTecnico;
		public int IdRequerimientoDepartamento;
		public int IdRequerimientoTipo;
		public DateTime FechaDesde;
		public DateTime FechaHasta;
		public int Prioridad;
		public bool Activo;
		public string NombreDepartamento;
		public string NombreTecnico;
		public string TipoTurnoDescripcion;

		// ===== CONSTRUCTOR =====
		public RequerimientoTurnoSoporteModel()
		{
			idRequerimientoTurnoSoporte = 0;
			Prioridad = 0;
			Activo = true;
			FechaDesde = DateTime.Today;
			FechaHasta = DateTime.Today;
		}
	}
}
