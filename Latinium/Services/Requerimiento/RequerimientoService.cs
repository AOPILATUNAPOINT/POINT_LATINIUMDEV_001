using System;
using System.Data;
using System.Data.SqlClient;
using System.Collections;
using Latinium.Models.Requerimiento;
using Latinium.Utils;

namespace Latinium.Services.Requerimiento
{
	public class RequerimientoService : IRequerimientoService
	{
		private string _cadenaConexion;

		public RequerimientoService(string cadenaConexion)
		{
			_cadenaConexion = cadenaConexion;
		}

		// =========================================
		// CREAR REQUERIMIENTO
		// =========================================
		public int CrearRequerimiento(SqlCommand cmd, RequerimientoModel model)
		{
			int idSeguimiento = 0;

			cmd.Parameters.Clear();
			cmd.CommandText = "RequerimientoCrear";
			cmd.CommandType = CommandType.StoredProcedure;

			cmd.Parameters.Add("@idEmpresa", SqlDbType.Int).Value = model.IdEmpresa;
			cmd.Parameters.Add("@Descripcion", SqlDbType.VarChar).Value = model.Descripcion;
			cmd.Parameters.Add("@idEstado", SqlDbType.Int).Value = model.IdEstado;
			cmd.Parameters.Add("@FechaTentativaEntrega", SqlDbType.DateTime).Value = DateTime.Now;
			cmd.Parameters.Add("@NotaInterna", SqlDbType.VarChar).Value = model.NotaInterna;
			cmd.Parameters.Add("@idPersonal", SqlDbType.Int).Value = model.IdResponsable;
			cmd.Parameters.Add("@Titulo", SqlDbType.VarChar, 50).Value = model.Titulo;
			cmd.Parameters.Add("@idRequerimientoTipo", SqlDbType.Int).Value = model.IdTipoRequerimiento;
			cmd.Parameters.Add("@idAsignado", SqlDbType.Int).Value = model.IdAsignado;
			cmd.Parameters.Add("@Numero", SqlDbType.VarChar, 50).Value = model.Numero;
			cmd.Parameters.Add("@Fecha", SqlDbType.DateTime).Value = DateTime.Now;
			cmd.Parameters.Add("@Responsable", SqlDbType.VarChar, 500).Value = model.Responsable;
			cmd.Parameters.Add("@Asignado", SqlDbType.VarChar, 500).Value = model.Asignado;
			cmd.Parameters.Add("@Empresa", SqlDbType.VarChar, 500).Value = model.Empresa;
			cmd.Parameters.Add("@idRequerimientoDepartamento", SqlDbType.Int).Value = model.IdRequerimientoDepartamento;

			cmd.Parameters.Add("@idPantalla", SqlDbType.Int).Value =
				(model.IdPantalla > 0) ? (object)model.IdPantalla : DBNull.Value;

			cmd.Parameters.Add("@Modulo", SqlDbType.VarChar, 500).Value =
				(model.Modulo == null || model.Modulo == "" ? (object)DBNull.Value : model.Modulo);

			cmd.Parameters.Add("@Pantalla", SqlDbType.VarChar, 500).Value =
				(model.Pantalla == null || model.Pantalla == "" ? (object)DBNull.Value : model.Pantalla);

			cmd.Parameters.Add("@SubPantalla", SqlDbType.VarChar, 500).Value =
				(model.SubPantalla == null || model.SubPantalla == "" ? (object)DBNull.Value : model.SubPantalla);

			cmd.Parameters.Add("@idRequerimientoPadre", SqlDbType.Int).Value = model.IdRequerimientoPadre;

			SqlParameter pOut = new SqlParameter("@idSeguimiento", SqlDbType.Int);
			pOut.Direction = ParameterDirection.Output;
			cmd.Parameters.Add(pOut);

			cmd.ExecuteNonQuery();

			if (pOut.Value != null && pOut.Value != DBNull.Value)
				idSeguimiento = Convert.ToInt32(pOut.Value);

			return idSeguimiento;
		}

		// =========================================
		// OBTENER REQUERIMIENTO
		// =========================================
		// =========================================
		// OBTENER REQUERIMIENTO SIN BLOQUEO DE TRANSACCION
		// =========================================
		public RequerimientoModel ObtenerTicket(SqlCommand oCmd, int idRequerimiento)
		{
			RequerimientoModel model = null;

			// Reutilizamos la conexión del comando
			oCmd.CommandText = "RequerimientoConXId";
			oCmd.CommandType = CommandType.StoredProcedure;

			// Limpiamos parámetros anteriores
			oCmd.Parameters.Clear();
			oCmd.Parameters.Add("@idRequerimiento", SqlDbType.Int).Value = idRequerimiento;

			using (SqlDataReader dr = oCmd.ExecuteReader())
			{
				if (dr.Read())
				{
					model = new RequerimientoModel();

					try { model.IdEmpresa = Convert.ToInt32(dr["idEmpresa"]); } 
					catch {}
					try { model.Empresa = dr["Empresa"].ToString(); } 
					catch {}
					try { model.Titulo = dr["Titulo"].ToString(); } 
					catch {}
					try { model.Descripcion = dr["Descripcion"].ToString(); } 
					catch {}
					try { model.IdEstado = Convert.ToInt32(dr["idEstado"]); } 
					catch {}
					try { model.IdResponsable = Convert.ToInt32(dr["idResponsable"]); } 
					catch {}
					try { model.Responsable = dr["Responsable"].ToString(); } 
					catch {}
					try { model.IdAsignado = Convert.ToInt32(dr["idAsignado"]); } 
					catch {}
					try { model.Asignado = dr["Asignado"].ToString(); } 
					catch {}
					try { model.IdTipoRequerimiento = Convert.ToInt32(dr["idTipoRequerimiento"]); } 
					catch {}
					try { model.IdRequerimientoDepartamento = Convert.ToInt32(dr["idDepartamento"]); } 
					catch {}
					try { model.NombreDepartamento = dr["NombreDepartamento"].ToString(); } 
					catch {}
					try { model.IdBodega = Convert.ToInt32(dr["idBodega"]); } 
					catch {}
					try { model.NombreBodega = dr["NombreBodega"].ToString(); } 
					catch {}
					try { model.IdCargo = Convert.ToInt32(dr["idCargo"]); } 
					catch {}
					try { model.NombreCargo = dr["NombreCargo"].ToString(); } 
					catch {}
					try { model.IdCiudad = Convert.ToInt32(dr["idCiudad"]); } 
					catch {}
					try { model.NombreCiudad = dr["NombreCiudad"].ToString(); } 
					catch {}
					try { model.IdPantalla = Convert.ToInt32(dr["idPantalla"]); } 
					catch {}
					try { model.Modulo = dr["Modulo"].ToString(); } 
					catch {}
					try { model.Pantalla = dr["Pantalla"].ToString(); } 
					catch {}
					try { model.SubPantalla = dr["SubPantalla"].ToString(); } 
					catch {}
					try { model.Numero = dr["Numero"].ToString(); } 
					catch {}
					try { model.UltimaNota = dr["UltimaNota"].ToString(); } 
					catch {}
					try { model.FechaUltimaNota = Convert.ToDateTime(dr["FechaUltimaNota"]); } 
					catch {}
					try { model.UsuarioRespondeNota = dr["UsuarioRespondeNota"].ToString(); } 
					catch {}
					try { model.IdRequerimiento = Convert.ToInt32(dr["idRequerimiento"]); } 
					catch {}
					try { model.IdRequerimientoPadre = Convert.ToInt32(dr["idRequerimiento"]); } 
					catch {}

					model.Fecha = DateTime.Now;
				}
			}

			return model;
		}
		// Fecha automática (GETDATE)
		private DateTime _fechaAsignacion = DateTime.Now;

		#region Constructores

		/// <summary>
		/// Usa la conexión por defecto
		/// </summary>
		public RequerimientoService()
		{
			_cadenaConexion = null;
		}

		#endregion

		#region Centros de Costos / Proyectos

		public ArrayList ObtenerCentrosCostoActivos()
		{
			ArrayList lista = new ArrayList();

			try
			{
				string sql = 
					"SELECT p.idProyecto, p.Nombre " +
					"FROM proyecto p " +
					"WHERE p.FechaCierre IS NULL " +
					"ORDER BY p.Nombre";

				DataTable dt = DBHelper.ExecuteDataTable(
					sql,
					null,
					_cadenaConexion
					);

				for (int i = 0; i < dt.Rows.Count; i++)
				{
					DataRow row = dt.Rows[i];

					CentroCostoModel model = new CentroCostoModel();
					model.Id = Convert.ToInt32(row["idProyecto"]);
					model.Nombre = row["Nombre"].ToString();

					lista.Add(model);
				}
			}
			catch (Exception ex)
			{
				throw new Exception("Error al obtener centros de costo", ex);
			}

			return lista;
		}


		#endregion

		#region Validaciones

		public bool ExisteRequerimiento(string numeroRequerimiento)
		{
			try
			{
				SqlParameter[] parametros =
				{
					new SqlParameter("@Numero", numeroRequerimiento)
				};

				DataTable dt = DBHelper.ExecuteDataTable(
					"EXEC RequerimientoExisteCon @Numero",
					parametros,
					_cadenaConexion
					);

				return dt.Rows.Count > 0;
			}
			catch (Exception ex)
			{
				throw new Exception("Error al validar requerimiento", ex);
			}
		}

		#endregion

		#region Guardar Requerimiento

		public void CrearRequerimiento(
			string numero,
			int idProyecto,
			int idUsuario
			)
		{
			try
			{
				SqlParameter[] parametros =
				{
					new SqlParameter("@Numero", numero),
					new SqlParameter("@IdProyecto", idProyecto),
					new SqlParameter("@IdUsuario", idUsuario),
					new SqlParameter("@FechaAsignacion", DateTime.Now)
				};

				DBHelper.ExecuteNonQuery(
					"EXEC RequerimientoIns @Numero, @IdProyecto, @IdUsuario, @FechaAsignacion",
					parametros,
					_cadenaConexion
					);
			}
			catch (Exception ex)
			{
				throw new Exception("Error al crear requerimiento", ex);
			}
		}

		#endregion

		// -----------------------------------------
		// Obtener idUsuario por código de Nomina
		// -----------------------------------------
		public int ObtenerIdUsuarioPorCodigoNomina(string codigoNomina)
		{
			if(codigoNomina == null || codigoNomina == "")
				return 0;

			string sQuery = "SELECT TOP 1 idUsuario FROM Usuario WHERE Nombre = '" 
				+ codigoNomina.Replace("'", "''") + "'";
			return Funcion.iEjecucion(_cadenaConexion, sQuery, false);
		}
		// -----------------------------------------
		// Obtener proyecto asignado por usuario
		// -----------------------------------------
		public int ObtenerProyectoPorUsuario(int idUsuario)
		{
			if (idUsuario == 0)
				return 0;

			string sQuery =
				"DECLARE @idProyecto INT " +
				"SELECT TOP 1 @idProyecto = n.idProyecto " +
				"FROM Usuario u " +
				"INNER JOIN Nomina n ON n.Codigo = u.Nombre " +
				"WHERE u.idUsuario = @idUsuario " +
				"IF @idProyecto IS NULL " +
				"BEGIN " +
				"   SELECT TOP 1 @idProyecto = p.idProyecto " +
				"   FROM Bodega b " +
				"   INNER JOIN UsuariosBodegas ub ON ub.Bodega = b.Bodega " +
				"   LEFT JOIN Proyecto p ON p.idProyecto = b.idProyecto " +
				"   WHERE ub.idUsuario = @idUsuario " +
				"   AND b.Factura = 1 " +
				"   AND b.Remates = 0 " +
				"END " +
				"SELECT ISNULL(@idProyecto,0)";

			SqlParameter[] parametros =
	{
		new SqlParameter("@idUsuario", idUsuario)
	};

			return DBHelper.ExecuteScalarInt(sQuery, parametros, _cadenaConexion);
		}

		public int ObtenerProyectoNominaPorPersonal(int idPersonal)
		{
			if (idPersonal == 0)
				return 0;

			string sQuery =
				"SELECT TOP 1 ISNULL(idProyecto,0) " +
				"FROM Nomina " +
				"WHERE idPersonal = @idPersonal";

			SqlParameter[] parametros =
			{
				new SqlParameter("@idPersonal", idPersonal)
			};

			return DBHelper.ExecuteScalarInt(sQuery, parametros, _cadenaConexion);
		}

		public ArrayList ObtenerCentrosCostoTodos()
		{
			ArrayList lista = new ArrayList();

			try
			{
				string sql =
					"SELECT p.idProyecto, p.Nombre, p.FechaCierre " +
					"FROM Proyecto p " +
					"ORDER BY p.Nombre";

				DataTable dt = DBHelper.ExecuteDataTable(
					sql,
					null,
					_cadenaConexion
					);

				for (int i = 0; i < dt.Rows.Count; i++)
				{
					DataRow row = dt.Rows[i];

					CentroCostoModel model = new CentroCostoModel();

					model.Id = Convert.ToInt32(row["idProyecto"]);

					string nombre = row["Nombre"].ToString();

					// Si está cerrado lo marcamos
					if (row["FechaCierre"] != DBNull.Value)
					{
						nombre = nombre + " (CERRADO)";
					}

					model.Nombre = nombre;

					lista.Add(model);
				}
			}
			catch (Exception ex)
			{
				throw new Exception("Error al obtener centros de costo", ex);
			}

			return lista;
		}

		// =========================================
		// OBTENER NUMERACION DE REQUERIMIENTO (SIN COMMIT)
		// =========================================
		// Nueva función para generar numeración usando un comando existente
		public string ObtenerNumeracion(SqlCommand oCmdActualiza, int idRequerimientoDepartamento)
		{
			// =============================
			// 1) Obtener idTipoFactura
			// =============================
			int idTipoFactura = 0;

			oCmdActualiza.Parameters.Clear();
			oCmdActualiza.CommandText = "RequerimientoDepartamentoModel"; // SP
			oCmdActualiza.CommandType = CommandType.StoredProcedure;

			oCmdActualiza.Parameters.Add("@NombreColumna", SqlDbType.VarChar, 50).Value = "idTipoFactura";
			oCmdActualiza.Parameters.Add("@id", SqlDbType.Int).Value = idRequerimientoDepartamento;

			object result = oCmdActualiza.ExecuteScalar();

			if (result == null || result == DBNull.Value)
				return null;

			idTipoFactura = Convert.ToInt32(result);

			// =============================
			// 2) Generar numeración
			// =============================
			oCmdActualiza.Parameters.Clear();
			oCmdActualiza.CommandText = "RequermientoNumeracion"; // SP
			oCmdActualiza.CommandType = CommandType.StoredProcedure;

			oCmdActualiza.Parameters.Add("@idTipoFactura", SqlDbType.Int).Value = idTipoFactura;

			object num = oCmdActualiza.ExecuteScalar();

			if (num == null || num == DBNull.Value)
				return null;

			return Convert.ToString(num);
		}

		public int CrearSeguimiento(SqlCommand cmd, SeguimientoModel model)
		{
			cmd.CommandText = "RequerimientoSeguimientoCrear";
			cmd.CommandType = CommandType.StoredProcedure;
			cmd.Parameters.Clear();

			cmd.Parameters.Add(new SqlParameter("@Notas", model.Notas));
			cmd.Parameters.Add(new SqlParameter("@idPersonal", model.idPersonal));
			cmd.Parameters.Add(new SqlParameter("@idRequerimiento", model.idRequerimiento));
			cmd.Parameters.Add(new SqlParameter("@idEstado", model.idEstado));
			cmd.Parameters.Add(new SqlParameter("@idAsignado", model.idAsignado));

			cmd.Parameters.Add(new SqlParameter("@NotaInterna", model.NotaInterna == null ? "" : model.NotaInterna));
			cmd.Parameters.Add(new SqlParameter("@UsuarioRespondeNota", model.UsuarioRespondeNota == null ? "" : model.UsuarioRespondeNota));

			// Como NO hay nullable, usa validación manual
			cmd.Parameters.Add(new SqlParameter("@idRequerimientoUrgencia", model.idRequerimientoUrgencia == 0 ? (object)DBNull.Value : model.idRequerimientoUrgencia));
			cmd.Parameters.Add(new SqlParameter("@idRequerimientoComplejidad", model.idRequerimientoComplejidad == 0 ? (object)DBNull.Value : model.idRequerimientoComplejidad));

			cmd.Parameters.Add(new SqlParameter("@FechaEntregaAdm", model.FechaEntregaAdm == DateTime.MinValue ? (object)DBNull.Value : model.FechaEntregaAdm));
			cmd.Parameters.Add(new SqlParameter("@FechaEntregaTec", model.FechaEntregaTec == DateTime.MinValue ? (object)DBNull.Value : model.FechaEntregaTec));

			cmd.Parameters.Add(new SqlParameter("@idRequerimientoDepartamento", model.idRequerimientoDepartamento));
			cmd.Parameters.Add(new SqlParameter("@idRequerimientoTipo", model.idRequerimientoTipo));

			cmd.Parameters.Add(new SqlParameter("@NombreCompletoCol", model.NombreCompletoCol == null ? (object)DBNull.Value : model.NombreCompletoCol));
			cmd.Parameters.Add(new SqlParameter("@FechaIngresoCol", model.FechaIngresoCol == DateTime.MinValue ? (object)DBNull.Value : model.FechaIngresoCol));
			cmd.Parameters.Add(new SqlParameter("@CargoAsignadoCol", model.CargoAsignadoCol == null ? (object)DBNull.Value : model.CargoAsignadoCol));

			return Convert.ToInt32(cmd.ExecuteScalar());
		}
	}
}