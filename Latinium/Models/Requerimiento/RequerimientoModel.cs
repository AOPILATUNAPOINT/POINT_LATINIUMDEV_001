using System;

namespace Latinium.Models.Requerimiento
{
	/// <summary>
	/// Descripción breve de RequerimientoModel.
	/// </summary>
	public class RequerimientoModel
	{
		// CAMPOS
		public int IdRequerimiento;
		public int IdEmpresa;
		public string Empresa;

		public string Numero;

		public string Titulo;
		public string Descripcion;
		public string NotaInterna;

		public int IdEstado;

		public int IdResponsable;
		public string Responsable;

		public int IdAsignado;
		public string Asignado;

		public int IdTipoRequerimiento;

		public int IdDepartamento;
		public string NombreDepartamento;

		public int IdBodega;
		public string NombreBodega;

		public int IdCargo;
		public string NombreCargo;

		public int IdCiudad;
		public string NombreCiudad;

		public int IdPantalla;
		public string Modulo;
		public string Pantalla;
		public string SubPantalla;

		public DateTime Fecha;

		public int IdRequerimientoDepartamento;

		public int IdRequerimientoPadre;

		public string UltimaNota;
		
		public DateTime FechaUltimaNota;
		public string UsuarioRespondeNota;

		// CONSTRUCTOR
		public RequerimientoModel()
		{
		}
	}
}