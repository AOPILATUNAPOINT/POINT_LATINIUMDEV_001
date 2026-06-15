using System;

namespace Latinium.Models.Pinpad
{
	/// <summary>
	/// Modelo de diferidos para PINPAD
	/// </summary>
	public class PinpadDiferido
	{
		// Identificadores
		public string IdDiferido;     // '01', '02', etc.
		public string IdRed;          // '01', '02', '03'

		// Información
		public string TipoDiferido;   // Ej: "Sin Interés", "Con Interés"
		public string Descripcion;    // Opcional

		// Datos para transacción
		public string Codigo;         // Código que envías al PINPAD
		public int Plazo;             // Número de cuotas
		public int Gracia;            // Meses de gracia

		// Texto que se muestra en el ComboBox
		public override string ToString()
		{
			string texto = "";

			// Tipo (ej: Sin interés / Con interés)
			if (TipoDiferido != null && TipoDiferido != "")
				texto += TipoDiferido;
			else
				texto += "Diferido";

			// Plazo
			texto += " (" + Plazo + " meses";

			// Gracia (solo si aplica)
			if (Gracia > 0)
				texto += " | " + Gracia + " meses de gracia";

			texto += ")";

			return texto;
		}
	}
}