using System;

namespace Latinium.Utils
{
	public class CifradoTrama
	{
		// Este método solo valida el hash que ya fue extraído
		public static bool validateHash(string hashRecibido)
		{
			if (hashRecibido == null || hashRecibido == "")
				return false;

			return hashRecibido.Length == 32;
		}


		// Verifica que la trama tenga al menos 40 bytes
		public static bool EsTramaValida(byte[] trama)
		{
			return trama != null && trama.Length >= 40;
		}
	}
}
