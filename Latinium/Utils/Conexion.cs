using System;
using System.Net.Sockets;
using System.Threading;

namespace Latinium.Utils
{
	public class Conexion
	{
		private static byte[] dataRecived;
		private static TcpClient client;
		private static bool connected;
		private static string ip;
		private static int puerto;
		private static byte[] dataSend;

		public static bool Connectar(string ipParam, int puertoParam, byte[] dataSendParam, int timeout)
		{
			dataRecived = new byte[0];
			ip = ipParam;
			puerto = puertoParam;
			dataSend = dataSendParam;
			connected = false;
			client = null;

			Thread connectThread = new Thread(new ThreadStart(ConnectThreadMethod));
			connectThread.Start();

			bool completed = connectThread.Join(timeout);
			if (!completed)
			{
				// Timeout: cerrar conexión si existe
				try
				{
					if (client != null)
					{
						client.Close();
					}
				}
				catch { }
				return false; // No se conectó a tiempo
			}

			if (!connected)
				return false; // No se conectó

			try
			{
				client.ReceiveTimeout = timeout;
				client.SendTimeout = timeout;

				NetworkStream stream = client.GetStream();

				// Enviar datos
				stream.Write(dataSend, 0, dataSend.Length);

				// Leer datos
				byte[] buffer = new byte[1024];
				int bytesRead = stream.Read(buffer, 0, buffer.Length);

				// Guardar sólo bytes recibidos
				dataRecived = new byte[bytesRead];
				Array.Copy(buffer, dataRecived, bytesRead);

				stream.Close();
				client.Close();

				return true; // Éxito
			}
			catch
			{
				dataRecived = new byte[0];
				try
				{
					if (client != null)
					{
						client.Close();
					}
				}
				catch { }
				return false; // Error al enviar/recibir
			}
		}

		private static void ConnectThreadMethod()
		{
			try
			{
				client = new TcpClient(ip, puerto);
				connected = true;
			}
			catch
			{
				connected = false;
			}
		}

		public static byte[] getDataRecived()
		{
			return dataRecived;
		}
	}
}
