using System;
using System.Data;

namespace Latinium.Models
{
	public class FamilyFriendsInvitadoDetalleModel
	{
		// -----------------------------
		// Campos privados
		// -----------------------------
		private Guid idInvitado;
		private int idCompra;

		private string invitado;
		private string cedula;
		private string email;
		private string celular;

		private string areaNomina;
		private string patrocinadorNombre;
		private string patrocinadorCedula;

		private string tipoDocumento;
		private string numeroDocumento;
		private DateTime fechaDocumento;

		private string ff;
		private string articulo;
		private string codigo;

		private decimal cantidad;
		private decimal descuentoPorc;
		private decimal subTotal;

		private decimal costoTotalUltimo;
		private decimal margenUtilidad;

		private decimal total;   // NUEVO

		private string nombreBodega;
		private string formaPago;

		// -----------------------------
		// Constructor
		// -----------------------------
		public FamilyFriendsInvitadoDetalleModel()
		{
			idInvitado = Guid.Empty;
			idCompra = 0;

			invitado = string.Empty;
			cedula = string.Empty;
			email = string.Empty;
			celular = string.Empty;

			areaNomina = string.Empty;
			patrocinadorNombre = string.Empty;
			patrocinadorCedula = string.Empty;

			tipoDocumento = string.Empty;
			numeroDocumento = string.Empty;
			fechaDocumento = DateTime.MinValue;

			ff = string.Empty;
			articulo = string.Empty;
			codigo = string.Empty;

			cantidad = 0m;
			descuentoPorc = 0m;
			subTotal = 0m;

			costoTotalUltimo = 0m;
			margenUtilidad = 0m;

			nombreBodega = string.Empty;
			formaPago = string.Empty;
		}

		// -----------------------------
		// Propiedades
		// -----------------------------
		public Guid IdInvitado { get { return idInvitado; } set { idInvitado = value; } }
		public int IdCompra { get { return idCompra; } set { idCompra = value; } }

		public string Invitado { get { return invitado; } set { invitado = value; } }
		public string Cedula { get { return cedula; } set { cedula = value; } }
		public string Email { get { return email; } set { email = value; } }
		public string Celular { get { return celular; } set { celular = value; } }

		public string AreaNomina { get { return areaNomina; } set { areaNomina = value; } }
		public string PatrocinadorNombre { get { return patrocinadorNombre; } set { patrocinadorNombre = value; } }
		public string PatrocinadorCedula { get { return patrocinadorCedula; } set { patrocinadorCedula = value; } }

		public string TipoDocumento { get { return tipoDocumento; } set { tipoDocumento = value; } }
		public string NumeroDocumento { get { return numeroDocumento; } set { numeroDocumento = value; } }
		public DateTime FechaDocumento { get { return fechaDocumento; } set { fechaDocumento = value; } }

		public string FF { get { return ff; } set { ff = value; } }
		public string Articulo { get { return articulo; } set { articulo = value; } }
		public string Codigo { get { return codigo; } set { codigo = value; } }

		public decimal Cantidad { get { return cantidad; } set { cantidad = value; } }
		public decimal DescuentoPorc { get { return descuentoPorc; } set { descuentoPorc = value; } }
		public decimal SubTotal { get { return subTotal; } set { subTotal = value; } }

		public decimal CostoTotalUltimo { get { return costoTotalUltimo; } set { costoTotalUltimo = value; } }
		public decimal MargenUtilidad { get { return margenUtilidad; } set { margenUtilidad = value; } }

		public decimal Total { get { return total; } set { total = value; } } // NUEVO

		public string NombreBodega { get { return nombreBodega; } set { nombreBodega = value; } }
		public string FormaPago { get { return formaPago; } set { formaPago = value; } }

		// -----------------------------
		// Load desde DataReader
		// -----------------------------
		public static FamilyFriendsInvitadoDetalleModel Load(IDataReader reader)
		{
			FamilyFriendsInvitadoDetalleModel m = new FamilyFriendsInvitadoDetalleModel();

			if (reader["idWEB_FF_Invitado"] != DBNull.Value)
				m.IdInvitado = new Guid(reader["idWEB_FF_Invitado"].ToString());

			if (reader["idCompra"] != DBNull.Value)
				m.IdCompra = Convert.ToInt32(reader["idCompra"]);

			if (reader["Invitado"] != DBNull.Value)
				m.Invitado = Convert.ToString(reader["Invitado"]);

			if (reader["Cedula"] != DBNull.Value)
				m.Cedula = Convert.ToString(reader["Cedula"]);

			if (reader["Email"] != DBNull.Value)
				m.Email = Convert.ToString(reader["Email"]);

			if (reader["Celular"] != DBNull.Value)
				m.Celular = Convert.ToString(reader["Celular"]);

			if (reader["AreaNomina"] != DBNull.Value)
				m.AreaNomina = Convert.ToString(reader["AreaNomina"]);

			if (reader["PatrocinadorNombre"] != DBNull.Value)
				m.PatrocinadorNombre = Convert.ToString(reader["PatrocinadorNombre"]);

			if (reader["PatrocinadorCedula"] != DBNull.Value)
				m.PatrocinadorCedula = Convert.ToString(reader["PatrocinadorCedula"]);

			if (reader["TipoDocumento"] != DBNull.Value)
				m.TipoDocumento = Convert.ToString(reader["TipoDocumento"]);

			if (reader["NumeroDocumento"] != DBNull.Value)
				m.NumeroDocumento = Convert.ToString(reader["NumeroDocumento"]);

			if (reader["FechaDocumento"] != DBNull.Value)
				m.FechaDocumento = Convert.ToDateTime(reader["FechaDocumento"]);

			if (reader["FF"] != DBNull.Value)
				m.FF = Convert.ToString(reader["FF"]);

			if (reader["Articulo"] != DBNull.Value)
				m.Articulo = Convert.ToString(reader["Articulo"]);

			if (reader["Codigo"] != DBNull.Value)
				m.Codigo = Convert.ToString(reader["Codigo"]);

			if (reader["Cantidad"] != DBNull.Value)
				m.Cantidad = Convert.ToDecimal(reader["Cantidad"]);

			if (reader["DescuentoPorc"] != DBNull.Value)
				m.DescuentoPorc = Convert.ToDecimal(reader["DescuentoPorc"]);

			if (reader["SubTotal"] != DBNull.Value)
				m.SubTotal = Convert.ToDecimal(reader["SubTotal"]);

			if (reader["CostoTotalUltimo"] != DBNull.Value)
				m.CostoTotalUltimo = Convert.ToDecimal(reader["CostoTotalUltimo"]);

			if (reader["MargenUtilidad"] != DBNull.Value)
				m.MargenUtilidad = Convert.ToDecimal(reader["MargenUtilidad"]);

			if (reader["NombreBodega"] != DBNull.Value)
				m.NombreBodega = Convert.ToString(reader["NombreBodega"]);

			if (reader["FormaPago"] != DBNull.Value)
				m.FormaPago = Convert.ToString(reader["FormaPago"]);

			return m;
		}
	}
}