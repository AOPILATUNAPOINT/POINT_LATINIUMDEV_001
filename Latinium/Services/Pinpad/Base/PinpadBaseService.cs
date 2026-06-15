using System;

namespace Latinium.Services.Pinpad
{
	/// <summary>
	/// Descripción breve de PinpadBaseService.
	/// </summary>
	public class PinpadBaseService
	{
		protected PinpadTcpService _tcp = new PinpadTcpService();
		protected PinpadRepository _repo = new PinpadRepository();

		protected string CalcularTalla(string cuerpo)
		{
			int longitud = cuerpo.Length + 32;
			return longitud.ToString("X4");
		}

		protected string FormatearMonto(decimal monto)
		{
			long montoCentavos = (long)(monto * 100);
			return montoCentavos.ToString().PadLeft(12, '0');
		}

		protected string ExtraerValor(string json, string campo)
		{
			string patron = "\"" + campo + "\":";
			int pos = json.IndexOf(patron);

			if (pos < 0)
				return "";

			pos += patron.Length;

			while (pos < json.Length && (json[pos] == ' ' || json[pos] == '\"'))
				pos++;

			int fin = json.IndexOf("\"", pos);
			if (fin < 0)
				return "";

			return json.Substring(pos, fin - pos);
		}

		protected bool ExtraerBool(string json, string campo)
		{
			string patron = "\"" + campo + "\":";
			int pos = json.IndexOf(patron);

			if (pos < 0)
				return false;

			pos += patron.Length;

			while (pos < json.Length && json[pos] == ' ')
				pos++;

			int fin = json.IndexOfAny(new char[] { ',', '}' }, pos);
			if (fin < 0)
				fin = json.Length;

			string valor = json.Substring(pos, fin - pos).Trim().ToLower();

			return valor == "true";
		}
	}
}
