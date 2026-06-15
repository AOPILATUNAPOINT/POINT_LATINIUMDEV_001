using System;
using System.Data.SqlClient;

namespace Latinium.Services
{
	public class ConfigService
	{
		// Constructor privado: evita instanciar la clase
		private ConfigService() { }

		public static string ObtenerValor(string clave)
		{
			string valor = "";

			try
			{
				string safeClave = clave.Replace("'", "''");
				string sQuery = "SELECT ValorStr FROM WEB_GlobalConfig WHERE Activo = 1 AND Nombre = '" + safeClave + "'";
				valor = Funcion.sEjecucion(MenuLatinium.sconexionEmpresa, sQuery);

				//if (valor == null || valor.Length == 0)
					//throw new Exception("No se encontró configuración activa para '" + clave + "'.");

			}
			catch (Exception ex)
			{
				string sms = ex.Message;
				//throw new Exception("Error obteniendo la configuración '" + clave + "': " + ex.Message);
			}

			return valor;
		}
	}
}
