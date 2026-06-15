using System;
using System.Data;
using System.Data.SqlClient;
using Latinium.Models;
using System.Windows.Forms;

namespace Latinium.Services
{
	public class FamilyFriendsService
	{
		private string connectionString;
		public static int EstadoFFCompraTemporal = 1;
		public static int EstadoFFCompraContabilizada = 2;
		public static int EstadoFFCompraNotaCredito = 3;
		public static int EstadoFFCompraNoAcepta = 4;

		public FamilyFriendsService(string connectionString)
		{
			this.connectionString = connectionString;
		}

		public FamilyFriendsClienteModel BuscarClienteFamilyFriends(string cedula, int idFormaPago)
		{
			FamilyFriendsClienteModel result = new FamilyFriendsClienteModel();
			SqlConnection connection = null;
			SqlCommand command = null;
			SqlDataReader reader = null;

			// ===== LOG INICIO
			Funcion.LogSistema(
				frmRequerimiento.sconexionInicio,
				"FamilyFriendsService.BuscarClienteFamilyFriends",
				"INICIO | Cedula=[" + cedula + "] IdFormaPago=[" + idFormaPago + "]",
				"FamilyFriends",
				Funcion.slogtipoinformacion
				);

			// ===== VALIDACIONES (solo log)
			if (cedula == null || cedula.Trim().Length == 0)
			{
				Funcion.LogSistema(
					frmRequerimiento.sconexionInicio,
					"FamilyFriendsService.BuscarClienteFamilyFriends",
					"VALIDACIÓN: Cedula vacía o NULL",
					"FamilyFriends",
					Funcion.slogtipoadvertencia
					);
			}

			if (idFormaPago <= 0)
			{
				Funcion.LogSistema(
					frmRequerimiento.sconexionInicio,
					"FamilyFriendsService.BuscarClienteFamilyFriends",
					"VALIDACIÓN: IdFormaPago inválido = " + idFormaPago,
					"FamilyFriends",
					Funcion.slogtipoadvertencia
					);
			}

			try
			{
				connection = new SqlConnection(this.connectionString);
				command = new SqlCommand("BuscaCliente_FamilyFriends", connection);
				command.CommandType = CommandType.StoredProcedure;

				// Parámetros
				SqlParameter pCedula = new SqlParameter("@Cedula", SqlDbType.NVarChar, 40);
				pCedula.Value = (cedula != null) ? (object)cedula : DBNull.Value;
				command.Parameters.Add(pCedula);

				SqlParameter pFormaPago = new SqlParameter("@IdFormaPago", SqlDbType.Int);
				pFormaPago.Value = idFormaPago;
				command.Parameters.Add(pFormaPago);

				Funcion.LogSistema(
					frmRequerimiento.sconexionInicio,
					"FamilyFriendsService.BuscarClienteFamilyFriends",
					"Ejecutando SP BuscaCliente_FamilyFriends",
					"FamilyFriends",
					Funcion.slogtipoinformacion
					);

				connection.Open();
				reader = command.ExecuteReader();

				if (reader.Read())
				{
					// ===== CARGA COMPLETA DEL MODELO
					result = FamilyFriendsClienteModel.Load(reader);

					// ===== LOG DETALLADO DE SALIDA (TODO lo que retorna)
					Funcion.LogSistema(
						frmRequerimiento.sconexionInicio,
						"FamilyFriendsService.BuscarClienteFamilyFriends",
						"SP OK | " +
						"Cedula=" + result.Cedula +
						" | IdCliente=" + result.IdCliente +
						" | ExisteInvitado=" + result.ExisteInvitado +
						" | FechaInicioEvento=" + result.FechaInicioEvento +
						" | FechaFinEvento=" + result.FechaFinEvento +
						" | FacturasRegistradas=" + result.FacturasRegistradas +
						" | PuedeVerPrecio=" + result.PuedeVerPrecio +
						" | IdFormaPago=" + result.IdFormaPago +
						" | TipoInvitado=" + result.TipoInvitado +
						" | IdTipoPrecio=" + result.IdTipoPrecio +
						" | IdWEB_FF_Invitado=" + result.IdWEB_FF_Invitado +
						" | Mensaje=" + result.Mensaje,
						"FamilyFriends",
						Funcion.slogtipoinformacion
						);
				}
				else
				{
					// ===== SIN REGISTROS
					result = new FamilyFriendsClienteModel();

					Funcion.LogSistema(
						frmRequerimiento.sconexionInicio,
						"FamilyFriendsService.BuscarClienteFamilyFriends",
						"SP sin registros",
						"FamilyFriends",
						Funcion.slogtipoinformacion
						);
				}
			}
			catch (Exception ex)
			{
				// ===== ERROR CONTROLADO (NO rompe flujo)
				Funcion.LogSistema(
					frmRequerimiento.sconexionInicio,
					"FamilyFriendsService.BuscarClienteFamilyFriends",
					"EXCEPTION: " + ex.GetType().ToString() + " | " + ex.Message,
					"FamilyFriends",
					Funcion.slogtipoerror
					);

				result = new FamilyFriendsClienteModel();
				result.Mensaje = "Error interno al consultar Family & Friends.";
			}
			finally
			{
				if (reader != null && !reader.IsClosed)
					reader.Close();

				if (connection != null && connection.State == ConnectionState.Open)
					connection.Close();

				Funcion.LogSistema(
					frmRequerimiento.sconexionInicio,
					"FamilyFriendsService.BuscarClienteFamilyFriends",
					"FIN | Retorno garantizado del modelo",
					"FamilyFriends",
					Funcion.slogtipoinformacion
					);
			}

			// ===== SIEMPRE retorna TODO el modelo
			return result;
		}

		public FamilyFriendsVerificacionModel VerificarFF(FamilyFriendsRequestModel req)
		{
			FamilyFriendsVerificacionModel resultado = new FamilyFriendsVerificacionModel();

			// Identificador para los logs
			string nombreClase = this.GetType().Name;
			string nombreMetodo = "VerificarFF";

			try
			{
				// ===========================
				// Log de entrada
				// ===========================
				string ultimoRucConfirmado = (string)req.GetType()
					.GetField("ultimoRucClienteConfirmado", System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Instance)
					.GetValue(req);

				object ultimoValorFPConfirmado = req.GetType()
					.GetField("ultimoValorFormaPagoConfirmado", System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Instance)
					.GetValue(req);

				string logEntrada = string.Format(
					"Inicio {0} | RucCliente: {1}, ValorFormaPago: {2}, UltimoRucClienteConfirmado: {3}, UltimoValorFormaPagoConfirmado: {4}, FueRechazado: {5}",
					nombreMetodo,
					req.RucCliente,
					req.ValorFormaPago != null ? req.ValorFormaPago.ToString() : "null",
					ultimoRucConfirmado,
					ultimoValorFPConfirmado != null ? ultimoValorFPConfirmado.ToString() : "null",
					req.FueRechazado
					);

				Funcion.LogSistema(frmRequerimiento.sconexionInicio, nombreClase + ": " + nombreMetodo, logEntrada, nombreMetodo, Funcion.slogtipoinformacion);

				// ===========================
				// Validación mínima
				// ===========================
				if (req == null || req.RucCliente == null)
				{
					resultado.MensajeSP = "Datos insuficientes para la verificación.";
					Funcion.LogSistema(frmRequerimiento.sconexionInicio, nombreClase + ": " + nombreMetodo,
						"Validación fallida: req es null o RucCliente es null", nombreMetodo, Funcion.slogtipoinformacion);
					return resultado;
				}

				// ===========================
				// 1) Obtener idFormaPagoFF
				// ===========================
				int idFormaPagoFF = 0;
				if (req.ValorFormaPago != null && req.ValorFormaPago != DBNull.Value)
				{
					try
					{
						idFormaPagoFF = Convert.ToInt32(req.ValorFormaPago);
					}
					catch
					{
						idFormaPagoFF = 0;
					}
				}

				Funcion.LogSistema(frmRequerimiento.sconexionInicio, nombreClase + ": " + nombreMetodo,
					"idFormaPagoFF calculado: " + idFormaPagoFF.ToString(), nombreMetodo, Funcion.slogtipoinformacion);

				// ===========================
				// 2) Llamar al SP
				// ===========================
				FamilyFriendsClienteModel clienteFF = this.BuscarClienteFamilyFriends(req.RucCliente, idFormaPagoFF);

				if (clienteFF == null)
				{
					resultado.MensajeSP = "No se pudo obtener información del cliente.";
					Funcion.LogSistema(frmRequerimiento.sconexionInicio, nombreClase + ": " + nombreMetodo,
						"SP retorno null", nombreMetodo, Funcion.slogtipoinformacion);
					return resultado;
				}

				// Log de todos los campos del clienteFF
				Funcion.LogSistema(frmRequerimiento.sconexionInicio, nombreClase + ": " + nombreMetodo,
					string.Format(
					"ClienteFF obtenido | Cedula: {0}, IdCliente: {1}, ExisteInvitado: {2}, FechaInicioEvento: {3}, FechaFinEvento: {4}, FacturasRegistradas: {5}, PuedeVerPrecio: {6}, Mensaje: {7}, IdFormaPago: {8}, TipoInvitado: {9}, IdTipoPrecio: {10}, IdWEB_FF_Invitado: {11}",
					clienteFF.Cedula,
					clienteFF.IdCliente,
					clienteFF.ExisteInvitado,
					clienteFF.FechaInicioEvento,
					clienteFF.FechaFinEvento,
					clienteFF.FacturasRegistradas,
					clienteFF.PuedeVerPrecio,
					clienteFF.Mensaje,
					clienteFF.IdFormaPago,
					clienteFF.TipoInvitado,
					clienteFF.IdTipoPrecio,
					clienteFF.IdWEB_FF_Invitado
					),
					nombreMetodo,
					Funcion.slogtipoinformacion
					);

				resultado.MensajeSP = (clienteFF.Mensaje != null) ? clienteFF.Mensaje : "";
				resultado.IdWEB_FF_Invitado = clienteFF.IdWEB_FF_Invitado;

				if (!clienteFF.PuedeVerPrecio)
				{
					Funcion.LogSistema(frmRequerimiento.sconexionInicio, nombreClase + ": " + nombreMetodo,
						"Cliente no puede ver precio, retorno inmediato", nombreMetodo, Funcion.slogtipoinformacion);
					return resultado;
				}

				resultado.IdCliente = clienteFF.IdCliente;

				// ===========================
				// 3) Si ya se respondió antes
				// ===========================
				if (!req.DebePreguntarNuevamente())
				{
					if (req.FueRechazado)
					{
						resultado.EsFF = false;
						resultado.RechazoPromocion = true;
						Funcion.LogSistema(frmRequerimiento.sconexionInicio, nombreClase + ": " + nombreMetodo,
							"No se debe preguntar nuevamente y fue rechazado", nombreMetodo, Funcion.slogtipoinformacion);
						return resultado;
					}

					resultado.EsFF = true;
					resultado.IdTipoPrecio = clienteFF.IdTipoPrecio;
					resultado.RechazoPromocion = false;
					Funcion.LogSistema(frmRequerimiento.sconexionInicio, nombreClase + ": " + nombreMetodo,
						"No se debe preguntar nuevamente y fue aceptado", nombreMetodo, Funcion.slogtipoinformacion);
					return resultado;
				}

				// ===========================
				// 4) Mostrar modal
				// ===========================
				Funcion.LogSistema(frmRequerimiento.sconexionInicio, nombreClase + ": " + nombreMetodo,
					"Mostrando modal de confirmación", nombreMetodo, Funcion.slogtipoinformacion);

				frmFamilyFriendsConfirm modal = new frmFamilyFriendsConfirm(
					"El Cliente está registrado en FAMILY & FRIENDS.\n\n¿Desea continuar con esta promoción?"
					);

				DialogResult resp = modal.ShowDialog(req.Owner);

				// ===========================
				// 5) Procesar respuesta
				// ===========================
				if (resp == DialogResult.OK && modal.Confirmado)
				{
					resultado.EsFF = true;
					resultado.IdTipoPrecio = clienteFF.IdTipoPrecio;
					req.MarcarComoConfirmado();
					resultado.RechazoPromocion = false;

					Funcion.LogSistema(frmRequerimiento.sconexionInicio, nombreClase + ": " + nombreMetodo,
						"Modal confirmado, cliente acepta promoción", nombreMetodo, Funcion.slogtipoinformacion);

					return resultado;
				}

				req.MarcarComoRechazado();
				resultado.RechazoPromocion = true;
				Funcion.LogSistema(frmRequerimiento.sconexionInicio, nombreClase + ": " + nombreMetodo,
					"Modal no confirmado, cliente rechaza promoción", nombreMetodo, Funcion.slogtipoinformacion);

				return resultado;
			}
			catch (Exception ex)
			{
				Funcion.LogSistema(frmRequerimiento.sconexionInicio, nombreClase + ": " + nombreMetodo,
					"Excepción en VerificarFF: " + ex.Message, nombreMetodo, Funcion.slogtipoerror);
				throw;
			}
		}
		/// <summary>
		/// Guarda o actualiza una compra usando el SP WEB_FF_CompraGuardar
		/// y retorna el idWEB_FF_Compra (GUID)
		/// </summary>
		public Guid GuardarCompra(FamilyFriendsCompraModel model)
		{
			string nombreClase = this.GetType().Name;
			string nombreMetodo = "GuardarCompra";

			// ============================
			// Validación previa con logs
			// ============================
			if (model == null)
			{
				Funcion.LogSistema(frmRequerimiento.sconexionInicio, nombreClase + ": " + nombreMetodo,
					"El modelo recibido es nulo.", nombreMetodo, Funcion.slogtipoinformacion);
				throw new ArgumentNullException("model", "El modelo recibido es nulo.");
			}

			if (model.IdWEB_FF_Invitado == Guid.Empty)
			{
				Funcion.LogSistema(frmRequerimiento.sconexionInicio, nombreClase + ": " + nombreMetodo,
					"El IdWEB_FF_Invitado es obligatorio.", nombreMetodo, Funcion.slogtipoinformacion);
				throw new ArgumentException("El invitado es obligatorio.", "IdWEB_FF_Invitado");
			}

			if (model.IdCliente <= 0)
			{
				Funcion.LogSistema(frmRequerimiento.sconexionInicio, nombreClase + ": " + nombreMetodo,
					"El IdCliente es obligatorio y debe ser mayor a cero.", nombreMetodo, Funcion.slogtipoinformacion);
				throw new ArgumentException("El idCliente es obligatorio y debe ser mayor a cero.", "IdCliente");
			}

			if (model.IdTipoPrecio <= 0)
			{
				Funcion.LogSistema(frmRequerimiento.sconexionInicio, nombreClase + ": " + nombreMetodo,
					"El IdTipoPrecio es obligatorio.", nombreMetodo, Funcion.slogtipoinformacion);
				throw new ArgumentException("El idTipoPrecio es obligatorio.", "IdTipoPrecio");
			}

			if (model.IdFormaPago <= 0)
			{
				Funcion.LogSistema(frmRequerimiento.sconexionInicio, nombreClase + ": " + nombreMetodo,
					"El IdFormaPago es obligatorio.", nombreMetodo, Funcion.slogtipoinformacion);
				throw new ArgumentException("El idFormaPago es obligatorio.", "IdFormaPago");
			}

			// ============================
			// Lógica original con logs
			// ============================
			SqlConnection cn = new SqlConnection(connectionString);
			SqlCommand cmd = new SqlCommand("WEB_FF_CompraGuardar", cn);
			cmd.CommandType = CommandType.StoredProcedure;

			cmd.Parameters.Add(new SqlParameter("@idCompra", SqlDbType.Int)).Value = model.IdCompra;
			cmd.Parameters.Add(new SqlParameter("@idWEB_FF_Invitado", SqlDbType.UniqueIdentifier)).Value = model.IdWEB_FF_Invitado;
			cmd.Parameters.Add(new SqlParameter("@Estado", SqlDbType.Int)).Value = model.Estado;
			cmd.Parameters.Add(new SqlParameter("@idTipoPrecio", SqlDbType.Int)).Value = model.IdTipoPrecio;
			cmd.Parameters.Add(new SqlParameter("@idFormaPago", SqlDbType.Int)).Value = model.IdFormaPago;
			cmd.Parameters.Add(new SqlParameter("@idCliente", SqlDbType.Int)).Value = model.IdCliente;

			SqlParameter pNota = new SqlParameter("@idNotaCredito", SqlDbType.Int);
			pNota.Value = model.IdNotaCredito > 0 ? (object)model.IdNotaCredito : DBNull.Value;
			cmd.Parameters.Add(pNota);

			// Log de parámetros
			Funcion.LogSistema(frmRequerimiento.sconexionInicio, nombreClase + ": " + nombreMetodo,
				string.Format("Parámetros SP | idCompra: {0}, IdWEB_FF_Invitado: {1}, Estado: {2}, IdTipoPrecio: {3}, IdFormaPago: {4}, IdCliente: {5}, IdNotaCredito: {6}",
				model.IdCompra,
				model.IdWEB_FF_Invitado,
				model.Estado,
				model.IdTipoPrecio,
				model.IdFormaPago,
				model.IdCliente,
				model.IdNotaCredito),
				nombreMetodo, Funcion.slogtipoinformacion);

			Guid idResultado = Guid.Empty;

			try
			{
				cn.Open();
				SqlDataReader reader = cmd.ExecuteReader();

				if (reader.Read())
				{
					if (reader["idWEB_FF_Compra"] != DBNull.Value)
					{
						idResultado = (Guid)reader["idWEB_FF_Compra"];
					}
				}

				reader.Close();

				Funcion.LogSistema(frmRequerimiento.sconexionInicio, nombreClase + ": " + nombreMetodo,
					"Resultado obtenido: idWEB_FF_Compra = " + idResultado.ToString(),
					nombreMetodo, Funcion.slogtipoinformacion);
			}
			catch (Exception ex)
			{
				Funcion.LogSistema(frmRequerimiento.sconexionInicio, nombreClase + ": " + nombreMetodo,
					"Excepción al ejecutar WEB_FF_CompraGuardar: " + ex.Message,
					nombreMetodo, Funcion.slogtipoerror);
				throw new Exception("Error al ejecutar WEB_FF_CompraGuardar: " + ex.Message, ex);
			}
			finally
			{
				if (cn.State != ConnectionState.Closed)
					cn.Close();
			}

			return idResultado;
		}

		/// <summary>
		/// Actualiza únicamente la Nota de Crédito de una compra FF
		/// usando el SP WEB_FF_CompraGuardar
		/// </summary>
		public Guid ActualizarNotaCreditoFF(int idCompra, int idNotaCredito)
		{
			string nombreClase = this.GetType().Name;
			string nombreMetodo = "ActualizarNotaCreditoFF";

			// ============================
			// Validación previa con logs
			// ============================
			if (idCompra <= 0)
			{
				Funcion.LogSistema(frmRequerimiento.sconexionInicio, nombreClase + ": " + nombreMetodo,
					"idCompra inválido: " + idCompra, nombreMetodo, Funcion.slogtipoinformacion);
				throw new ArgumentException("idCompra inválido.", "idCompra");
			}

			if (idNotaCredito <= 0)
			{
				Funcion.LogSistema(frmRequerimiento.sconexionInicio, nombreClase + ": " + nombreMetodo,
					"idNotaCredito inválido: " + idNotaCredito, nombreMetodo, Funcion.slogtipoinformacion);
				throw new ArgumentException("idNotaCredito inválido.", "idNotaCredito");
			}

			Guid idResultado = Guid.Empty;

			try
			{
				using (SqlConnection cn = new SqlConnection(connectionString))
				using (SqlCommand cmd = new SqlCommand("WEB_FF_CompraGuardar", cn))
				{
					cmd.CommandType = CommandType.StoredProcedure;

					cmd.Parameters.Add("@idCompra", SqlDbType.Int).Value = idCompra;
					cmd.Parameters.Add("@idWEB_FF_Invitado", SqlDbType.UniqueIdentifier).Value = Guid.Empty;
					cmd.Parameters.Add("@Estado", SqlDbType.Int).Value = FamilyFriendsService.EstadoFFCompraNotaCredito; // 3
					cmd.Parameters.Add("@idTipoPrecio", SqlDbType.Int).Value = 0;
					cmd.Parameters.Add("@idFormaPago", SqlDbType.Int).Value = 0;
					cmd.Parameters.Add("@idCliente", SqlDbType.Int).Value = 0;
					cmd.Parameters.Add("@idNotaCredito", SqlDbType.Int).Value = idNotaCredito;

					// Log de parámetros enviados
					Funcion.LogSistema(frmRequerimiento.sconexionInicio, nombreClase + ": " + nombreMetodo,
						string.Format(
						"Parámetros SP | idCompra: {0}, idWEB_FF_Invitado: {1}, Estado: {2}, idTipoPrecio: {3}, idFormaPago: {4}, idCliente: {5}, idNotaCredito: {6}",
						idCompra,
						Guid.Empty,
						FamilyFriendsService.EstadoFFCompraNotaCredito,
						0, 0, 0,
						idNotaCredito),
						nombreMetodo, Funcion.slogtipoinformacion);

					cn.Open();

					using (SqlDataReader dr = cmd.ExecuteReader())
					{
						if (dr.Read() && dr["idWEB_FF_Compra"] != DBNull.Value)
							idResultado = (Guid)dr["idWEB_FF_Compra"];
					}

					// Log del resultado
					Funcion.LogSistema(frmRequerimiento.sconexionInicio, nombreClase + ": " + nombreMetodo,
						"Resultado obtenido: idWEB_FF_Compra = " + idResultado.ToString(),
						nombreMetodo, Funcion.slogtipoinformacion);
				}
			}
			catch (Exception ex)
			{
				Funcion.LogSistema(frmRequerimiento.sconexionInicio, nombreClase + ": " + nombreMetodo,
					"Excepción en ActualizarNotaCreditoFF: " + ex.Message,
					nombreMetodo, Funcion.slogtipoerror);
				throw;
			}

			return idResultado;
		}

		public FamilyFriendsEstadoModel ObtenerEstadoFF(int idCompra)
		{
			string nombreClase = this.GetType().Name;
			string nombreMetodo = "ObtenerEstadoFF";

			FamilyFriendsEstadoModel modelo = new FamilyFriendsEstadoModel();

			// ============================
			// Log de entrada
			// ============================
			Funcion.LogSistema(frmRequerimiento.sconexionInicio, nombreClase + ": " + nombreMetodo,
				"Inicio ObtenerEstadoFF | idCompra: " + idCompra,
				nombreMetodo, Funcion.slogtipoinformacion);

			SqlConnection cn = new SqlConnection(connectionString);
			SqlCommand cmd = new SqlCommand("WEB_FF_EstadoCompra", cn);
			cmd.CommandType = CommandType.StoredProcedure;

			// Parámetro de entrada
			SqlParameter p1 = new SqlParameter("@idCompra", SqlDbType.Int);
			p1.Value = idCompra;
			cmd.Parameters.Add(p1);

			try
			{
				cn.Open();
				SqlDataReader dr = cmd.ExecuteReader();

				if (dr.Read())
				{
					modelo.Load(dr);           // Cargas por nombre desde el reader
					modelo.ExisteRegistro = true;

					// Log de todos los campos cargados
					System.Text.StringBuilder sb = new System.Text.StringBuilder();
					sb.Append("EstadoFF obtenido: ");
					foreach (System.Data.IDataRecord record in new System.Data.IDataRecord[] { dr })
					{
						for (int i = 0; i < dr.FieldCount; i++)
						{
							sb.AppendFormat("{0}: {1}, ", dr.GetName(i), dr[i] != DBNull.Value ? dr[i].ToString() : "null");
						}
					}
					Funcion.LogSistema(frmRequerimiento.sconexionInicio, nombreClase + ": " + nombreMetodo,
						sb.ToString().TrimEnd(' ', ','),
						nombreMetodo, Funcion.slogtipoinformacion);
				}
				else
				{
					modelo.ExisteRegistro = false;
					Funcion.LogSistema(frmRequerimiento.sconexionInicio, nombreClase + ": " + nombreMetodo,
						"No se encontró registro para idCompra: " + idCompra,
						nombreMetodo, Funcion.slogtipoinformacion);
				}

				dr.Close();
			}
			catch (Exception ex)
			{
				Funcion.LogSistema(frmRequerimiento.sconexionInicio, nombreClase + ": " + nombreMetodo,
					"Excepción en ObtenerEstadoFF: " + ex.Message,
					nombreMetodo, Funcion.slogtipoerror);
				throw;
			}
			finally
			{
				if (cn.State != ConnectionState.Closed)
					cn.Close();
			}

			return modelo;
		}
	}
}
