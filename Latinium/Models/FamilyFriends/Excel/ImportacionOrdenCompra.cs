using System;
using System.Collections;

namespace Latinium.Models.FamilyFriends.Excel
{
	[Serializable]
	public class RootImportacion
	{
		public bool ok;
		public string mensaje;
		public DataImportacion data;
		public ArrayList errores;
	}

	[Serializable]
	public class DataImportacion
	{
		public ProveedorImportacion proveedor;
		public DateTime fecha;
		public string notas;
		public ArrayList detalle;
	}

	[Serializable]
	public class ProveedorImportacion
	{
		public int id;
		public string nombre;
		public string ruc;
	}

	[Serializable]
	public class DetalleImportacion
	{
		public string codigo;
		public int cantidad;
		public decimal precio;
		public decimal subtotal;
		public ArrayList bodegas;
	}

	[Serializable]
	public class BodegaImportacion
	{
		public int id;
		public string codigo;
		public int valor;
	}
}
