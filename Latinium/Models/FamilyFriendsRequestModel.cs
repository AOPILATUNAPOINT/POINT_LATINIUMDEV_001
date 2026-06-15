using System;
using System.Windows.Forms;

namespace Latinium.Models
{
	/// <summary>
	/// Modelo para solicitud Family & Friends (compatible con .NET Framework 1.1)
	/// Maneja cambios para decidir si volver a pedir confirmación.
	/// </summary>
	public class FamilyFriendsRequestModel
	{
		// ================================
		// Campos privados (valores actuales)
		// ================================
		private object valorFormaPago;
		private string rucCliente;
		private Form owner;

		// ===========================================
		// Campos privados (valores ya confirmados)
		// ===========================================
		private object ultimoValorFormaPagoConfirmado;
		private string ultimoRucClienteConfirmado;

		// ================================
		// Nuevo campo para indicar rechazo
		// ================================
		private bool fueRechazado;

		// ================================
		// Propiedades públicas
		// ================================
		public object ValorFormaPago
		{
			get { return valorFormaPago; }
			set { valorFormaPago = value; }
		}

		public string RucCliente
		{
			get { return rucCliente; }
			set { rucCliente = value; }
		}

		public Form Owner
		{
			get { return owner; }
			set { owner = value; }
		}

		public bool FueRechazado
		{
			get { return fueRechazado; }
			set { fueRechazado = value; }
		}

		// ================================
		// Constructor
		// ================================
		public FamilyFriendsRequestModel()
		{
			valorFormaPago = null;
			rucCliente = "";

			// valores confirmados inicialmente vacíos
			ultimoValorFormaPagoConfirmado = null;
			ultimoRucClienteConfirmado = "";

			owner = null;
			fueRechazado = false;
		}

		// ==================================================
		// Determina si se debe volver a pedir confirmación
		// ==================================================
		public bool DebePreguntarNuevamente()
		{
			// Si ya fue rechazado, no volver a preguntar
			if (fueRechazado)
				return false;

			// Si cambia la forma de pago => preguntar
			if (ultimoValorFormaPagoConfirmado == null ||
				!ultimoValorFormaPagoConfirmado.Equals(valorFormaPago))
			{
				return true;
			}

			// Si cambia el RUC => preguntar
			if (ultimoRucClienteConfirmado == null ||
				!ultimoRucClienteConfirmado.Equals(rucCliente))
			{
				return true;
			}

			// No hubo cambios
			return false;
		}

		// ==================================================
		// Guarda los valores confirmados después de preguntar
		// ==================================================
		public void MarcarComoConfirmado()
		{
			ultimoValorFormaPagoConfirmado = valorFormaPago;
			ultimoRucClienteConfirmado = rucCliente;
			fueRechazado = false;
		}

		// ==================================================
		// Marca como rechazado
		// ==================================================
		public void MarcarComoRechazado()
		{
			fueRechazado = true;
			// Opcional: también puedes actualizar los valores confirmados si quieres
			ultimoValorFormaPagoConfirmado = valorFormaPago;
			ultimoRucClienteConfirmado = rucCliente;
		}
	}
}
