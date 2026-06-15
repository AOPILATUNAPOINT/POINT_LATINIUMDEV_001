using System;
using System.Collections;

namespace Latinium.Models.OrdenDeCompra.Importacion
{
	public class ImportacionXmlResult
	{
		public bool Ok;
		public string Mensaje;
		public int IdProveedor;
		public string RucProveedor;
		public string NombreProveedor;
		public DateTime Fecha;
		public string Notas;

		// NUEVO CAMPO
		public string Consignacion;

		public ArrayList Detalles = new ArrayList(); // Lista de ImportacionDetalle
	}

	public class ImportacionDetalle
	{
		public string Codigo;
		public int Cantidad;
		public decimal Precio;
		public decimal Subtotal;
		public ArrayList Bodegas = new ArrayList(); // Lista de DetalleBodegaModel
	}
}