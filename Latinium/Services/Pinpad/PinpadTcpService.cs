//using System;
//using System.Net.Sockets;
//using System.Text;
//using System.IO;
//using System.Threading;
//
//namespace Latinium.Services.Pinpad
//{
//	public class PinpadTcpService
//	{
//		private string _ip = "192.168.2.21";
//		private int _puerto = 9998;
//
//		private TcpClient _client = null;
//		private NetworkStream _stream = null;
//
//		// =========================
//		// CONEXION
//		// =========================
//		private void Conectar()
//		{
//			Console.WriteLine("Conectando a " + _ip + ":" + _puerto);
//
//			_client = new TcpClient();
//			_client.ReceiveTimeout = 15000;
//			_client.SendTimeout = 15000;
//
//			_client.Connect(_ip, _puerto);
//
//			_stream = _client.GetStream();
//
//			if (_stream == null)
//				throw new Exception("No se pudo obtener el NetworkStream");
//
//			Console.WriteLine("Conexion establecida");
//		}
//
//		// =========================
//		// ASEGURAR CONEXION
//		// =========================
//		private void AsegurarConexion()
//		{
//			try
//			{
//				if (_stream == null || _client == null)
//				{
//					Conectar();
//				}
//			}
//			catch
//			{
//				CerrarConexion();
//				Conectar();
//			}
//		}
//
//		// =========================
//		// ENVIO Y RECEPCION
//		// =========================
//		public string Enviar(string trama)
//		{
//			MemoryStream ms = null;
//
//			try
//			{
//				Console.WriteLine("====================================");
//				Console.WriteLine("INICIANDO ENVIO AL PINPAD");
//
//				AsegurarConexion();
//
//				byte[] data = Encoding.ASCII.GetBytes(trama);
//
//				Console.WriteLine("Trama enviada:");
//				Console.WriteLine(trama);
//
//				_stream.Write(data, 0, data.Length);
//				_stream.Flush();
//
//				Console.WriteLine("Trama enviada correctamente");
//
//				ms = new MemoryStream();
//				byte[] buffer = new byte[1024];
//
//				DateTime inicio = DateTime.Now;
//
//				while (true)
//				{
//					try
//					{
//						if (_stream.DataAvailable)
//						{
//							int bytesRead = _stream.Read(buffer, 0, buffer.Length);
//
//							if (bytesRead > 0)
//							{
//								ms.Write(buffer, 0, bytesRead);
//								Console.WriteLine("Bytes recibidos: " + bytesRead);
//							}
//						}
//						else
//						{
//							Thread.Sleep(50);
//						}
//					}
//					catch
//					{
//						Console.WriteLine("Error en lectura, reconectando...");
//						CerrarConexion();
//						Conectar();
//						throw new Exception("Conexion interrumpida durante lectura");
//					}
//
//					// Timeout
//					if ((DateTime.Now - inicio).TotalMilliseconds > 15000)
//					{
//						Console.WriteLine("Timeout esperando respuesta");
//						break;
//					}
//
//					// Detectar ETX
//					string parcial = Encoding.ASCII.GetString(ms.ToArray());
//
//					if (parcial.IndexOf("\u0003") >= 0)
//					{
//						Console.WriteLine("Fin de trama detectado");
//						break;
//					}
//				}
//
//				if (ms.Length == 0)
//					throw new Exception("No se recibio respuesta del pinpad");
//
//				string response = Encoding.ASCII.GetString(ms.ToArray());
//
//				response = response
//					.Replace("\0", "")
//					.Replace("\u0002", "")
//					.Replace("\u0003", "")
//					.Trim();
//
//				Console.WriteLine("RESPUESTA FINAL:");
//				Console.WriteLine(response);
//				Console.WriteLine("====================================");
//
//				return response;
//			}
//			catch (Exception ex)
//			{
//				Console.WriteLine("ERROR EN PINPAD:");
//				Console.WriteLine(ex.ToString());
//
//				return "{ \"Success\": false, \"DataAscii\": \"\", \"ErrorMessage\": \"" +
//					ex.Message.Replace("\"", "") + "\" }";
//			}
//			finally
//			{
//				try { if (ms != null) ms.Close(); } 
//				catch { }
//			}
//		}
//
//		// =========================
//		// CIERRE
//		// =========================
//		public void CerrarConexion()
//		{
//			try
//			{
//				if (_stream != null)
//				{
//					_stream.Close();
//					_stream = null;
//				}
//
//				if (_client != null)
//				{
//					_client.Close();
//					_client = null;
//				}
//
//				Console.WriteLine("Conexion cerrada");
//			}
//			catch (Exception ex)
//			{
//				Console.WriteLine("Error cerrando conexion: " + ex.Message);
//			}
//		}
//	}
//}

using System;
using System.Net.Sockets;
using System.Text;

public class PinpadTcpService
{
	private string _ip = "192.168.2.21";
	private int _puerto = 9998;

	public string Enviar(string mensaje)
	{
		TcpClient client = new TcpClient();
		client.ReceiveTimeout = 15000;
		client.SendTimeout = 15000;

		client.Connect(_ip, _puerto);

		NetworkStream stream = client.GetStream();

		byte[] data = Encoding.ASCII.GetBytes(mensaje);
		byte[] length = BitConverter.GetBytes(data.Length);

		stream.Write(length, 0, 4);
		stream.Write(data, 0, data.Length);

		byte[] lenBuffer = new byte[4];
		ReadExact(stream, lenBuffer, 4);
		int size = BitConverter.ToInt32(lenBuffer, 0);

		byte[] buffer = new byte[size];
		ReadExact(stream, buffer, size);
		string response = Encoding.UTF8.GetString(buffer);
		// Limpia unicode y escapes
		// 1. Quitar comillas externas PRIMERO
		if (response.StartsWith("\"") && response.EndsWith("\""))
		{
			response = response.Substring(1, response.Length - 2);
		}
		// 2. Luego limpiar escapes
		response = System.Text.RegularExpressions.Regex.Unescape(response);

		stream.Close();
		client.Close();

		return response;
	}

	private void ReadExact(NetworkStream stream, byte[] buffer, int size)
	{
		int read = 0;

		while (read < size)
		{
			int r = stream.Read(buffer, read, size - read);

			if (r == 0)
				throw new Exception("Conexión cerrada");

			read += r;
		}
	}
}