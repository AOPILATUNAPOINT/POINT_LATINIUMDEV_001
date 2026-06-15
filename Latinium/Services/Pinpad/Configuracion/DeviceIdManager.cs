using System;
using System.IO;

namespace Latinium.Services.Pinpad.Configuracion
{
	public class DeviceIdManager
	{
		private static string ObtenerRuta()
		{
			string path = ConfigManager.Get("RutaDeviceId");

			if (path != null && path != "")
				return path;

			return @"C:\Latinium\device.id";
		}

		public static string ObtenerDeviceId()
		{
			string path = ObtenerRuta();

			if (!File.Exists(path))
				return null;

			string value = "";

			StreamReader sr = null;

			try
			{
				sr = new StreamReader(path);
				value = sr.ReadToEnd().Trim();
			}
			catch
			{
				return null;
			}
			finally
			{
				if (sr != null)
					sr.Close();
			}

			try
			{
				Guid g = new Guid(value);
				return g.ToString();
			}
			catch
			{
				return null;
			}
		}

		public static void GenerarSiNoExiste()
		{
			string path = ObtenerRuta();

			if (!File.Exists(path))
			{
				string id = Guid.NewGuid().ToString();

				string dir = Path.GetDirectoryName(path);
				if (!Directory.Exists(dir))
					Directory.CreateDirectory(dir);

				StreamWriter sw = null;

				try
				{
					sw = new StreamWriter(path, false);
					sw.Write(id);
				}
				finally
				{
					if (sw != null)
						sw.Close();
				}
			}
		}
	}
}