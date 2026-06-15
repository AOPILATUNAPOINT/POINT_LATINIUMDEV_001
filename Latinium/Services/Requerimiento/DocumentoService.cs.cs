using System;
using System.Collections;
using System.IO;
using System.Data;
using System.Data.SqlClient;
using Latinium.Models.Requerimiento;

namespace Latinium.Services.Requerimiento
{
	/// <summary>
	/// Servicio para manejar documentos de requerimientos
	/// Compatible con .NET 1.1
	/// </summary>
	public class DocumentoService
	{
		private ArrayList listaDocumentos;
		private string usuarioActual;
		private string conexion;
		private string dirFacturacion;

		public DocumentoService()
		{
			listaDocumentos = new ArrayList();
		}

		public DocumentoService(string usuario, string conexionString, string dirFacturacion)
		{
			this.usuarioActual = usuario;
			this.conexion = conexionString;
			this.dirFacturacion = dirFacturacion;
			listaDocumentos = new ArrayList();
		}

		#region Manejo de documentos en memoria

		public bool AgregarDocumento(
			string rutaArchivo,
			string cadenaConexion,
			int idConfiguracionDocumento,
			int idConfiguracionTamano,
			out string mensaje)
		{
			mensaje = "";
			try
			{
				if (rutaArchivo == null)
				{
					mensaje = "La ruta del archivo es inválida.";
					return false;
				}

				// Validar tamaño y configuración
				FileInfo info = new FileInfo(rutaArchivo);
				long tamano = info.Length;

				DataSet dsConfiguracion = Funcion.dsEjecucion(
					cadenaConexion,
					string.Format(
					"EXEC RequerimientoConfgVal {0}, {1}, {2}",
					idConfiguracionDocumento,
					tamano,
					idConfiguracionTamano
					)
					);

				bool resultado = bool.Parse(dsConfiguracion.Tables[0].Rows[0]["bResultado"].ToString());
				mensaje = dsConfiguracion.Tables[0].Rows[0]["sMensaje"].ToString();
				return resultado;
			}
			catch (Exception ex)
			{
				mensaje = "Error al validar documento: " + ex.Message;
				return false;
			}
		}

		public ArrayList ObtenerDocumentos()
		{
			return listaDocumentos;
		}

		public bool EliminarDocumento(DocumentoModel doc, out string mensaje)
		{
			mensaje = "";
			try
			{
				if (listaDocumentos.Contains(doc))
				{
					listaDocumentos.Remove(doc);
					mensaje = "Documento eliminado.";
					return true;
				}
				mensaje = "Documento no encontrado.";
				return false;
			}
			catch (Exception ex)
			{
				mensaje = "Error al eliminar documento: " + ex.Message;
				return false;
			}
		}

		public bool VerDocumento(DocumentoModel doc, out string mensaje)
		{
			mensaje = "";
			try
			{
				if (File.Exists(doc.txtDirectorioarchivo))
				{
					System.Diagnostics.Process.Start(doc.txtDirectorioarchivo);
					return true;
				}
				mensaje = "El archivo no existe.";
				return false;
			}
			catch (Exception ex)
			{
				mensaje = "Error al abrir documento: " + ex.Message;
				return false;
			}
		}

		#endregion

		#region Guardar documentos en disco y BD

		/// <summary>
		/// Guarda documentos en carpeta física y en la base de datos
		/// </summary>
//		public bool GuardarDocumentos(
//			SqlCommand oCmd,
//			int idSeguimiento,
//			string numeroRequerimiento,
//			string sidRequerimientoDepartamento,
//			ArrayList documentos,
//			DateTime fechaAsignacion)
//		{
//			try
//			{
//				for (int i = 0; i < documentos.Count; i++)
//				{
//					DocumentoModel doc = (DocumentoModel)documentos[i];
//					string dtFecha = fechaAsignacion.ToString("yyyyMMdd");
//
//					// Carpeta destino
//					string carpetaDestino = "\\\\192.168.1.10\\Proyectos\\RequerimientoDocumento\\" 
//						+ dirFacturacion + "\\" + sidRequerimientoDepartamento + "\\" + numeroRequerimiento;
//
//					string carpetaRelativa = dirFacturacion + "\\" + sidRequerimientoDepartamento + "\\" + numeroRequerimiento;
//
//					if (!Directory.Exists(carpetaDestino))
//						Directory.CreateDirectory(carpetaDestino);
//
//					// Generar nombre final del documento
//					oCmd.Parameters.Clear();
//					oCmd.CommandText = "RequermientoNumeracion";
//					oCmd.CommandType = CommandType.StoredProcedure;
//					oCmd.Parameters.Add("@idTipoFactura", SqlDbType.Int).Value = 89;
//
//					string nombreDoc = oCmd.ExecuteScalar().ToString().Replace(" ", "_");
//					string extension = Path.GetExtension(doc.txtDirectorioarchivo);
//					string nombreFinal = nombreDoc + extension;
//
//					string rutaArchivoFinal = Path.Combine(carpetaDestino, nombreFinal);
//
//					// Validar existencia
//					if (File.Exists(rutaArchivoFinal))
//						throw new Exception("El archivo '" + nombreFinal + "' ya existe.");
//
//					// Copiar archivo
//					File.Copy(doc.txtDirectorioarchivo, rutaArchivoFinal, false);
//
//					// Actualizar modelo
//					doc.sCarpeta = rutaArchivoFinal;
//					doc.sUrlNube = Path.Combine(carpetaRelativa.Replace(" ", "_"), nombreFinal);
//					doc.Nombre = nombreFinal;
//					doc.idSeguimiento = idSeguimiento;
//					doc.Fecha = fechaAsignacion;
//
//					// Insertar en BD
//					oCmd.Parameters.Clear();
//					oCmd.CommandText = "RequerimientoDocumentoNuevo";
//					oCmd.CommandType = CommandType.StoredProcedure;
//
//					oCmd.Parameters.Add("@Usuario", SqlDbType.VarChar, 50).Value = usuarioActual;
//					oCmd.Parameters.Add("@Nuevo", SqlDbType.Int).Value = 0;
//					oCmd.Parameters.Add("@idResponsable", SqlDbType.Int).Value = doc.idResponsable;
//					oCmd.Parameters.Add("@idAsignado", SqlDbType.Int).Value = doc.idAsignado;
//					oCmd.Parameters.Add("@txtDirectorioarchivo", SqlDbType.VarChar, 500).Value = doc.txtDirectorioarchivo;
//					oCmd.Parameters.Add("@sCarpeta", SqlDbType.VarChar, 500).Value = doc.sCarpeta;
//					oCmd.Parameters.Add("@sUrlNube", SqlDbType.VarChar, 500).Value = doc.sUrlNube;
//					oCmd.Parameters.Add("@idSeguimiento", SqlDbType.Int).Value = doc.idSeguimiento;
//					oCmd.Parameters.Add("@idRequerimientoDocumento", SqlDbType.Int).Value = doc.idRequerimientoDocumento;
//					oCmd.Parameters.Add("@Fecha", SqlDbType.VarChar, 10).Value = dtFecha;
//
//					oCmd.ExecuteNonQuery();
//				}
//
//				return true;
//			}
//			catch (Exception ex)
//			{
//				Funcion.LogSistema(conexion, "DocumentoService", "Error guardando documentos: " + ex.GetType() + " " + ex.Message, "", Funcion.slogtipoerror);
//				throw;
//			}
//		}

		public bool GuardarDocumentos(
			SqlCommand oCmd,
			int idSeguimiento,
			string numeroRequerimiento,
			string sidRequerimientoDepartamento,
			ArrayList documentos,
			DateTime fechaAsignacion)
		{
			try
			{
				string rutaBase = ObtenerRutaBase(oCmd);

				for (int i = 0; i < documentos.Count; i++)
				{
					DocumentoModel doc = (DocumentoModel)documentos[i];

					string carpetaDestino = ConstruirCarpetaDestino(
						rutaBase,
						dirFacturacion,
						sidRequerimientoDepartamento,
						numeroRequerimiento
						);

					string carpetaRelativa = Path.Combine(
						Path.Combine(dirFacturacion, sidRequerimientoDepartamento),
						numeroRequerimiento
						);

					CrearCarpetaSiNoExiste(carpetaDestino);

					string nombreFinal = GenerarNombreDocumento(oCmd, doc);

					string rutaArchivoFinal = Path.Combine(carpetaDestino, nombreFinal);

					ValidarArchivoExiste(rutaArchivoFinal, nombreFinal);

					File.Copy(doc.txtDirectorioarchivo, rutaArchivoFinal, false);

					ActualizarModelo(doc, rutaArchivoFinal, carpetaRelativa, nombreFinal, idSeguimiento, fechaAsignacion);

					InsertarDocumentoBD(oCmd, doc, fechaAsignacion);
				}

				return true;
			}
			catch (Exception ex)
			{
				Funcion.LogSistema(conexion, "DocumentoService",
					"Error guardando documentos: " + ex.GetType() + " " + ex.Message,
					"", Funcion.slogtipoerror);
				throw;
			}
		}

		private string ObtenerRutaBase(SqlCommand oCmd)
		{
			oCmd.Parameters.Clear();
			oCmd.CommandText = @"
        SELECT Excepcion 
        FROM RequerimientoConfiguracion 
        WHERE Descripcion = 'RUTA_BASE_DOCUMENTOS' 
        AND Estado = 1";

			oCmd.CommandType = CommandType.Text;

			object result = oCmd.ExecuteScalar();

			if (result == null)
				throw new Exception("No se encontró la ruta base en configuración.");

			return result.ToString();
		}

		private string ConstruirCarpetaDestino(string rutaBase, string dirFacturacion, string departamento, string numeroReq)
		{
			return Path.Combine(
				Path.Combine(
				Path.Combine(rutaBase, dirFacturacion),
				departamento
				),
				numeroReq
				);
		}

		private void CrearCarpetaSiNoExiste(string ruta)
		{
			if (!Directory.Exists(ruta))
				Directory.CreateDirectory(ruta);
		}

		private string GenerarNombreDocumento(SqlCommand oCmd, DocumentoModel doc)
		{
			oCmd.Parameters.Clear();
			oCmd.CommandText = "RequermientoNumeracion";
			oCmd.CommandType = CommandType.StoredProcedure;
			oCmd.Parameters.Add("@idTipoFactura", SqlDbType.Int).Value = 89;

			string nombreDoc = oCmd.ExecuteScalar().ToString().Replace(" ", "_");
			string extension = Path.GetExtension(doc.txtDirectorioarchivo);

			return nombreDoc + extension;
		}

		private void ValidarArchivoExiste(string ruta, string nombre)
		{
			if (File.Exists(ruta))
				throw new Exception("El archivo '" + nombre + "' ya existe.");
		}

		private void ActualizarModelo(
			DocumentoModel doc,
			string rutaFinal,
			string carpetaRelativa,
			string nombreFinal,
			int idSeguimiento,
			DateTime fecha)
		{
			doc.sCarpeta = rutaFinal;
			doc.sUrlNube = Path.Combine(carpetaRelativa.Replace(" ", "_"), nombreFinal);
			doc.Nombre = nombreFinal;
			doc.idSeguimiento = idSeguimiento;
			doc.Fecha = fecha;
		}

		private void InsertarDocumentoBD(SqlCommand oCmd, DocumentoModel doc, DateTime fechaAsignacion)
		{
			string dtFecha = fechaAsignacion.ToString("yyyyMMdd");

			oCmd.Parameters.Clear();
			oCmd.CommandText = "RequerimientoDocumentoNuevo";
			oCmd.CommandType = CommandType.StoredProcedure;

			oCmd.Parameters.Add("@Usuario", SqlDbType.VarChar, 50).Value = usuarioActual;
			oCmd.Parameters.Add("@Nuevo", SqlDbType.Int).Value = 0;
			oCmd.Parameters.Add("@idResponsable", SqlDbType.Int).Value = doc.idResponsable;
			oCmd.Parameters.Add("@idAsignado", SqlDbType.Int).Value = doc.idAsignado;
			oCmd.Parameters.Add("@txtDirectorioarchivo", SqlDbType.VarChar, 500).Value = doc.txtDirectorioarchivo;
			oCmd.Parameters.Add("@sCarpeta", SqlDbType.VarChar, 500).Value = doc.sCarpeta;
			oCmd.Parameters.Add("@sUrlNube", SqlDbType.VarChar, 500).Value = doc.sUrlNube;
			oCmd.Parameters.Add("@idSeguimiento", SqlDbType.Int).Value = doc.idSeguimiento;
			oCmd.Parameters.Add("@idRequerimientoDocumento", SqlDbType.Int).Value = doc.idRequerimientoDocumento;
			oCmd.Parameters.Add("@Fecha", SqlDbType.VarChar, 10).Value = dtFecha;

			oCmd.ExecuteNonQuery();
		}

		#endregion
	}
}
