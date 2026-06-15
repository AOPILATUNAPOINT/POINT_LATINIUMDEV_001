using System;
using System.Collections;

namespace Latinium.Services
{
	/// <summary>
	/// Contrato del servicio de requerimientos
	/// Compatible con .NET 1.1
	/// </summary>
	public interface IRequerimientoService
	{
		bool ExisteRequerimiento(string numeroRequerimiento);
		ArrayList ObtenerCentrosCostoActivos();
	}
}
