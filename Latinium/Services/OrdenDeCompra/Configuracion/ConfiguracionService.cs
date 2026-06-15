using System;

namespace Latinium.Services.OrdenDeCompra.Configuracion
{
	/// <summary>
	/// Descripción breve de ConfiguracionService.
	/// </summary>
	public class ConfiguracionService
	{
		private C1.Data.C1DataSet cdsSeteoF;

		public ConfiguracionService(
			C1.Data.C1DataSet cds
			)
		{
			cdsSeteoF = cds;
		}

		public bool ObtenerValorBool(
			string clave,
			bool valorDefault
			)
		{
			try
			{
				string valor =
					ObtenerValor(
					clave);

				if (
					valor == null ||
					valor.Trim() == ""
					)
				{
					return valorDefault;
				}

				valor = valor.Trim();

				return (
					valor == "1" ||
					valor.ToLower() == "true"
					);
			}
			catch
			{
				return valorDefault;
			}
		}

		// =====================================================
		// OBTENER CONFIGURACION STRING
		// =====================================================

		public string ObtenerValor(
			string clave
			)
		{
			string sql =
				string.Format(
				@"SELECT TOP 1 Valor
				  FROM OrdenDeCompraConfiguracion
				  WHERE Clave = '{0}'
				  AND Activo = 1",
				clave.Replace("'", "''")
				);

			object valor =
				Funcion.sEscalarSQL(
				cdsSeteoF,
				sql
				);

			if (valor == null)
				return "";

			return valor.ToString();
		}

		// =====================================================
		// URL BASE API
		// =====================================================

		public string ObtenerUrlApi()
		{
			return ObtenerValor(
				"urlApiOrdenCompra");
		}
	}
}
