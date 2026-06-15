using System;
using System.Data;

namespace Latinium.Models
{
	/// <summary>
	/// Modelo para el reporte de Family & Friends (versión resumen).
	/// Contiene información del colaborador, sus métricas de invitados, facturas y notas de crédito.
	/// Compatible con .NET Framework 1.1.
	/// </summary>
	public class FamilyFriendsReporteModel
	{
		// Campos privados
		private string colaborador;
		private string empresaEmpleado;
		private string areaNomina;
		private string departamentoNomina;
		private string cargoNomina;
		private string identificacion;
		private string nombre_Bodega;
		private int totalInvitados;
		private int totalFacturas;
		private decimal totalMontoFacturas;
		private int totalNotasCreditoCount;
		private decimal totalNotasCredito;
		private decimal totalNeto;


		// Nuevos campos
		private decimal totalCosto;
		private decimal margenUtilidad;

		// Constructor por defecto
		public FamilyFriendsReporteModel()
		{
			this.colaborador = string.Empty;
			this.empresaEmpleado = string.Empty;
			this.areaNomina = string.Empty;
			this.departamentoNomina = string.Empty;
			this.cargoNomina = string.Empty;
			this.identificacion = string.Empty;
			this.nombre_Bodega = string.Empty;
			this.totalInvitados = 0;
			this.totalFacturas = 0;
			this.totalMontoFacturas = 0m;
			this.totalNotasCreditoCount = 0;
			this.totalNotasCredito = 0m;
			this.totalNeto = 0m;

			// Nuevos inicializados
			this.totalCosto = 0m;
			this.margenUtilidad = 0m;
		}

		// Constructor completo
		public FamilyFriendsReporteModel(
			string colaborador,
			string empresaEmpleado,
			string areaNomina,
			string departamentoNomina,
			string cargoNomina,
			string identificacion,
			string nombre_Bodega,
			int totalInvitados,
			int totalFacturas,
			decimal totalMontoFacturas,
			int totalNotasCreditoCount,
			decimal totalNotasCredito,
			decimal totalNeto,
			decimal totalCosto,
			decimal margenUtilidad)
		{
			this.colaborador = colaborador;
			this.empresaEmpleado = empresaEmpleado;
			this.areaNomina = areaNomina;
			this.departamentoNomina = departamentoNomina;
			this.cargoNomina = cargoNomina;
			this.identificacion = identificacion;
			this.nombre_Bodega = nombre_Bodega;
			this.totalInvitados = totalInvitados;
			this.totalFacturas = totalFacturas;
			this.totalMontoFacturas = totalMontoFacturas;
			this.totalNotasCreditoCount = totalNotasCreditoCount;
			this.totalNotasCredito = totalNotasCredito;
			this.totalNeto = totalNeto;

			// Nuevos
			this.totalCosto = totalCosto;
			this.margenUtilidad = margenUtilidad;
		}

		// Propiedades públicas
		public string Colaborador { get { return colaborador; } set { colaborador = value; } }
		public string EmpresaEmpleado { get { return empresaEmpleado; } set { empresaEmpleado = value; } }
		public string AreaNomina { get { return areaNomina; } set { areaNomina = value; } }
		public string DepartamentoNomina { get { return departamentoNomina; } set { departamentoNomina = value; } }
		public string CargoNomina { get { return cargoNomina; } set { cargoNomina = value; } }
		public string Identificacion { get { return identificacion; } set { identificacion = value; } }
		public string Nombre_Bodega { get { return nombre_Bodega; } set { nombre_Bodega = value; } }

		public int TotalInvitados { get { return totalInvitados; } set { totalInvitados = value; } }
		public int TotalFacturas { get { return totalFacturas; } set { totalFacturas = value; } }
		public decimal TotalMontoFacturas { get { return totalMontoFacturas; } set { totalMontoFacturas = value; } }
		public int TotalNotasCreditoCount { get { return totalNotasCreditoCount; } set { totalNotasCreditoCount = value; } }
		public decimal TotalNotasCredito { get { return totalNotasCredito; } set { totalNotasCredito = value; } }
		public decimal TotalNeto { get { return totalNeto; } set { totalNeto = value; } }

		// Propiedades nuevas
		public decimal TotalCosto { get { return totalCosto; } set { totalCosto = value; } }
		public decimal MargenUtilidad { get { return margenUtilidad; } set { margenUtilidad = value; } }

		// Método para cargar datos desde un DataReader
		public static FamilyFriendsReporteModel Load(IDataReader reader)
		{
			FamilyFriendsReporteModel reporte = new FamilyFriendsReporteModel();

			if (reader["Colaborador"] != DBNull.Value)
				reporte.Colaborador = Convert.ToString(reader["Colaborador"]);

			if (reader["EmpresaEmpleado"] != DBNull.Value)
				reporte.EmpresaEmpleado = Convert.ToString(reader["EmpresaEmpleado"]);

			if (reader["AreaNomina"] != DBNull.Value)
				reporte.AreaNomina = Convert.ToString(reader["AreaNomina"]);

			if (reader["DepartamentoNomina"] != DBNull.Value)
				reporte.DepartamentoNomina = Convert.ToString(reader["DepartamentoNomina"]);

			if (reader["CargoNomina"] != DBNull.Value)
				reporte.CargoNomina = Convert.ToString(reader["CargoNomina"]);

			if (reader["Identificacion"] != DBNull.Value)
				reporte.Identificacion = Convert.ToString(reader["Identificacion"]);

			if (reader["Nombre_Bodega"] != DBNull.Value)
				reporte.Nombre_Bodega = Convert.ToString(reader["Nombre_Bodega"]);

			if (reader["TotalInvitados"] != DBNull.Value)
				reporte.TotalInvitados = Convert.ToInt32(reader["TotalInvitados"]);

			if (reader["TotalFacturas"] != DBNull.Value)
				reporte.TotalFacturas = Convert.ToInt32(reader["TotalFacturas"]);

			if (reader["TotalMontoFacturas"] != DBNull.Value)
				reporte.TotalMontoFacturas = Convert.ToDecimal(reader["TotalMontoFacturas"]);

			if (reader["TotalNotasCreditoCount"] != DBNull.Value)
				reporte.TotalNotasCreditoCount = Convert.ToInt32(reader["TotalNotasCreditoCount"]);

			if (reader["TotalNotasCredito"] != DBNull.Value)
				reporte.TotalNotasCredito = Convert.ToDecimal(reader["TotalNotasCredito"]);

			if (reader["TotalNeto"] != DBNull.Value)
				reporte.TotalNeto = Convert.ToDecimal(reader["TotalNeto"]);

			// Nuevos campos
			if (reader["TotalCosto"] != DBNull.Value)
				reporte.TotalCosto = Convert.ToDecimal(reader["TotalCosto"]);

			if (reader["MargenUtilidad"] != DBNull.Value)
				reporte.MargenUtilidad = Convert.ToDecimal(reader["MargenUtilidad"]);

			return reporte;
		}
	}
}
