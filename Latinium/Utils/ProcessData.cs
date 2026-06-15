using System;
using System.Text;

namespace Latinium.Utils
{
	public class ProcessData
	{
		public static byte[] getFinalData(string[] datos)
		{
			// Concatenar todos los elementos del arreglo en un solo string
			string trama = string.Empty;
			for (int i = 0; i < datos.Length; i++)
			{
				trama += datos[i];
			}

			// Convertir el string concatenado a un arreglo de bytes ASCII
			byte[] resultado = Encoding.ASCII.GetBytes(trama);

			return resultado;
		}

		public static string byte2hex(byte[] bytes)
		{
			StringBuilder sb = new StringBuilder();
			for (int i = 0; i < bytes.Length; i++)
			{
				sb.Append(bytes[i].ToString("X2")); // Convierte byte a hex en mayúscula, dos dígitos
			}
			return sb.ToString();
		}

		public static string hex2AsciiStr(string hex)
		{
			if (hex == null || hex.Length == 0 || hex.Length % 2 != 0)
				return "";

			StringBuilder sb = new StringBuilder();

			for (int i = 0; i < hex.Length; i += 2)
			{
				string hs = hex.Substring(i, 2);
				int value = Convert.ToInt32(hs, 16);
				sb.Append((char)value);
			}

			return sb.ToString();
		}
	}
}

