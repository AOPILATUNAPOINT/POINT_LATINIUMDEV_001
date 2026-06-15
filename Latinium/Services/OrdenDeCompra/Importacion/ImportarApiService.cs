using System;
using System.Net;
using System.IO;
using System.Text;
using System.Xml;
using System.Collections;
using Latinium.Models.OrdenDeCompra.Importacion;
using Latinium.Models.OrdenDeCompra.Bodega;
using Latinium.Services.OrdenDeCompra.Configuracion;

namespace Latinium.Services.OrdenDeCompra.Importacion
{
	public class ImportarApiService
	{
		private C1.Data.C1DataSet cdsSeteoF;

		public ImportarApiService(
			C1.Data.C1DataSet cds
			)
		{
			cdsSeteoF = cds;
		}
		// =========================================================
		// 1. ORQUESTADOR PRINCIPAL
		// =========================================================
		public ImportacionXmlResult ProcesarArchivo(string base64File, int idUsuario)
		{
			ImportacionXmlResult resultado = new ImportacionXmlResult();

			HttpWebResponse response = EnviarRequest(base64File, idUsuario);

			if (response == null)
			{
				resultado.Ok = false;
				resultado.Mensaje = "No hubo respuesta del servidor";
				return resultado;
			}

			string xmlResponse = LeerXml(response);

			XmlDocument doc = new XmlDocument();
			doc.LoadXml(xmlResponse);

			XmlNode root = doc.SelectSingleNode("/root");

			if (root == null)
			{
				resultado.Ok = false;
				resultado.Mensaje = xmlResponse;
				return resultado;
			}

			// =========================
			// ERROR O OK
			// =========================
			if (EsError(root))
			{
				MapearError(root, resultado);
				return resultado;
			}

			XmlNode ordenNode = root.SelectSingleNode("OrdenCompra");

			if (ordenNode == null)
			{
				resultado.Ok = false;
				resultado.Mensaje = xmlResponse;
				return resultado;
			}

			MapearOk(ordenNode, resultado);

			return resultado;
		}

		// =========================================================
		// 3. LEER RESPUESTA
		// =========================================================
		private string LeerXml(HttpWebResponse response)
		{
			StreamReader reader = new StreamReader(response.GetResponseStream(), Encoding.UTF8);
			string xml = reader.ReadToEnd();
			reader.Close();
			return xml;
		}

		// =========================================================
		// 4. VALIDAR ERROR
		// =========================================================
		private bool EsError(XmlNode root)
		{
			return root.SelectSingleNode("OrdenCompra") == null;
		}

		// =========================================================
		// 5. MAPEAR ERROR (SIN STRINGBUILDER .NET MODERNO)
		// =========================================================
		private void MapearError(XmlNode root, ImportacionXmlResult resultado)
		{
			XmlNode msg = root.SelectSingleNode("mensaje");

			string mensaje = (msg != null) ? msg.InnerText : "Error desconocido";

			XmlNodeList errores = root.SelectNodes("errores/error");

			if (errores != null && errores.Count > 0)
			{
				IEnumerator en = errores.GetEnumerator();

				while (en.MoveNext())
				{
					XmlNode err = (XmlNode)en.Current;

					if (err != null && err.InnerText != "")
					{
						mensaje = mensaje + "\r\n" + err.InnerText;
					}
				}
			}

			resultado.Ok = false;
			resultado.Mensaje = mensaje;
		}

		private HttpWebResponse EnviarRequest(
			string base64File,
			int idUsuario
			)
		{
			try
			{
				// =========================
				// CONFIGURACION API
				// =========================

				ConfiguracionService config =
					new ConfiguracionService(
					cdsSeteoF);

				string url =
					config.ObtenerUrlApi() +
					"/api/v1/orden-compra/importar-xml";

				// =========================
				// REQUEST
				// =========================

				HttpWebRequest request =
					(HttpWebRequest)WebRequest.Create(
					url);

				request.Method = "POST";
				request.ContentType = "application/json";

				string json =
					"{ \"file\": \"" + base64File + "\", " +
					"\"idUsuario\": " + idUsuario + " }";

				byte[] data =
					Encoding.UTF8.GetBytes(
					json);

				request.ContentLength =
					data.Length;

				Stream stream =
					request.GetRequestStream();

				stream.Write(
					data,
					0,
					data.Length);

				stream.Close();

				return
					(HttpWebResponse)
					request.GetResponse();
			}
			catch (WebException ex)
			{
				if (ex.Response != null)
					return (HttpWebResponse)ex.Response;

				return null;
			}
		}

		// =========================================================
		// 6. MAPEAR OK
		// =========================================================
		private void MapearOk(XmlNode ordenNode, ImportacionXmlResult resultado)
		{
			XmlNode okNode = ordenNode.SelectSingleNode("ok");
			resultado.Ok = (okNode != null && okNode.InnerText.ToLower() == "true");

			XmlNode msgNode = ordenNode.SelectSingleNode("mensaje");
			resultado.Mensaje = (msgNode != null) ? msgNode.InnerText : "";

			if (!resultado.Ok)
				return;

			// =========================
			// CABECERA
			// =========================
			XmlNode prov = ordenNode.SelectSingleNode("Cabecera/Proveedor");

			if (prov != null)
			{
				resultado.IdProveedor = Convert.ToInt32(prov.SelectSingleNode("id").InnerText);
				resultado.RucProveedor = prov.SelectSingleNode("ruc").InnerText;
				resultado.NombreProveedor = prov.SelectSingleNode("nombre").InnerText;
			}

			XmlNode fecha = ordenNode.SelectSingleNode("Cabecera/fecha");

			if (fecha != null && fecha.InnerText != "")
			{
				resultado.Fecha = DateTime.Parse(fecha.InnerText);
			}

			XmlNode notas = ordenNode.SelectSingleNode("Cabecera/notas");

			if (notas != null)
			{
				resultado.Notas = notas.InnerText;
			}

			XmlNode consignacion = ordenNode.SelectSingleNode("Cabecera/consignacion");

			if (consignacion != null)
			{
				resultado.Consignacion = consignacion.InnerText;
			}
			else
			{
				resultado.Consignacion = "NO";
			}

			// =========================
			// DETALLE
			// =========================
			XmlNodeList items = ordenNode.SelectNodes("Detalle/item");

			IEnumerator iEnum = items.GetEnumerator();

			while (iEnum.MoveNext())
			{
				XmlNode item = (XmlNode)iEnum.Current;

				ImportacionDetalle det = new ImportacionDetalle();

				det.Codigo = item.SelectSingleNode("codigo").InnerText;
				det.Cantidad = Convert.ToInt32(item.SelectSingleNode("cantidad").InnerText);
				det.Precio = Convert.ToDecimal(item.SelectSingleNode("precio").InnerText);
				det.Subtotal = Convert.ToDecimal(item.SelectSingleNode("subtotal").InnerText);

				XmlNodeList bodegas = item.SelectNodes("bodegas/bodega");

				IEnumerator bEnum = bodegas.GetEnumerator();

				while (bEnum.MoveNext())
				{
					XmlNode b = (XmlNode)bEnum.Current;

					DetalleBodegaModel bm = new DetalleBodegaModel();
					bm.Bodega = Convert.ToInt32(b.SelectSingleNode("id").InnerText);
					bm.Cantidad = Convert.ToInt32(b.SelectSingleNode("valor").InnerText);

					det.Bodegas.Add(bm);
				}

				resultado.Detalles.Add(det);
			}
		}
	}
}