//using System;
//using System.Collections;
//using System.Data;
//using System.Data.SqlClient;
//using Latinium.Models.Pinpad.Configuracion;
//
//namespace Latinium.Services.Pinpad.Configuracion
//{
//	public class PinpadConfiguracionRepository
//	{
//		private string _connectionString;
//
//		public PinpadConfiguracionRepository(string connectionString)
//		{
//			_connectionString = connectionString;
//		}
//
//		// =========================
//		// CREATE
//		// =========================
////		public void Insertar(PinpadConfiguracion cfg)
////		{
////			SqlConnection cn = new SqlConnection(_connectionString);
////
////			try
////			{
////				string sql = @"
////				INSERT INTO Pinpad_Configuracion
////				(idPinpad_Configuracion, idUsuario, NumeroCaja, MID, TID, CID, Activo, Usuario)
////				VALUES
////				(@Id, @IdUsuario, @NumeroCaja, @MID, @TID, @CID, @Activo, @Usuario)";
////
////				SqlCommand cmd = new SqlCommand(sql, cn);
////
////				cmd.Parameters.Add(new SqlParameter("@Id", Guid.NewGuid()));
////				cmd.Parameters.Add(new SqlParameter("@IdUsuario", cfg.IdUsuario));
////				cmd.Parameters.Add(new SqlParameter("@NumeroCaja", cfg.NumeroCaja));
////				cmd.Parameters.Add(new SqlParameter("@MID", cfg.MID));
////				cmd.Parameters.Add(new SqlParameter("@TID", cfg.TID));
////				cmd.Parameters.Add(new SqlParameter("@CID", cfg.CID));
////				cmd.Parameters.Add(new SqlParameter("@Activo", cfg.Activo));
////
////				if (cfg.Usuario != null)
////					cmd.Parameters.Add(new SqlParameter("@Usuario", cfg.Usuario));
////				else
////					cmd.Parameters.Add(new SqlParameter("@Usuario", DBNull.Value));
////
////				cn.Open();
////				cmd.ExecuteNonQuery();
////			}
////			finally
////			{
////				cn.Close();
////			}
////		}
//
//		public void Insertar(PinpadConfiguracion cfg)
//		{
//			SqlConnection cn = new SqlConnection(_connectionString);
//
//			try
//			{
//				string sql = @"
//		INSERT INTO Pinpad_Configuracion
//		(
//			idPinpad_Configuracion,
//			DeviceId,
//			NombrePC,
//			Bodega,
//			NumeroCaja,
//			IpPinPad,
//			PuertoPinPad,
//			TimeoutPinPad,
//			MID,
//			TID,
//			CID,
//			Activo,
//			FechaSistema,
//			Usuario
//		)
//		VALUES
//		(
//			@Id,
//			@DeviceId,
//			@NombrePC,
//			@Bodega,
//			@NumeroCaja,
//			@IpPinPad,
//			@PuertoPinPad,
//			@TimeoutPinPad,
//			@MID,
//			@TID,
//			@CID,
//			@Activo,
//			GETDATE(),
//			@Usuario
//		)";
//
//				SqlCommand cmd = new SqlCommand(sql, cn);
//
//				cmd.Parameters.Add(new SqlParameter("@Id", Guid.NewGuid()));
//				cmd.Parameters.Add(new SqlParameter("@DeviceId", cfg.DeviceId));
//				cmd.Parameters.Add(new SqlParameter("@NombrePC", cfg.NombrePC));
//				cmd.Parameters.Add(new SqlParameter("@Bodega", cfg.Bodega));
//				cmd.Parameters.Add(new SqlParameter("@NumeroCaja", cfg.NumeroCaja));
//
//				cmd.Parameters.Add(new SqlParameter("@IpPinPad", cfg.IpPinPad));
//				cmd.Parameters.Add(new SqlParameter("@PuertoPinPad", cfg.PuertoPinPad));
//				cmd.Parameters.Add(new SqlParameter("@TimeoutPinPad", cfg.TimeoutPinPad));
//
//				cmd.Parameters.Add(new SqlParameter("@MID", cfg.MID));
//				cmd.Parameters.Add(new SqlParameter("@TID", cfg.TID));
//				cmd.Parameters.Add(new SqlParameter("@CID", cfg.CID));
//
//				cmd.Parameters.Add(new SqlParameter("@Activo", cfg.Activo));
//				cmd.Parameters.Add(new SqlParameter("@Usuario", cfg.Usuario ?? (object)DBNull.Value));
//
//				cn.Open();
//				cmd.ExecuteNonQuery();
//			}
//			finally
//			{
//				cn.Close();
//			}
//		}
//
//		// =========================
//		// READ (LISTAR)
//		// =========================
//		public ArrayList Listar()
//		{
//			ArrayList lista = new ArrayList();
//			SqlConnection cn = new SqlConnection(_connectionString);
//
//			try
//			{
//				string sql = "SELECT * FROM Pinpad_Configuracion where Activo = 1";
//				SqlCommand cmd = new SqlCommand(sql, cn);
//
//				cn.Open();
//				SqlDataReader dr = cmd.ExecuteReader();
//
//				while (dr.Read())
//				{
//					PinpadConfiguracion cfg = new PinpadConfiguracion();
//
//					cfg.IdPinpadConfiguracion = new Guid(dr["idPinpad_Configuracion"].ToString());
//					cfg.IdUsuario = Convert.ToInt32(dr["idUsuario"]);
//					cfg.NumeroCaja = dr["NumeroCaja"].ToString();
//					cfg.MID = dr["MID"].ToString();
//					cfg.TID = dr["TID"].ToString();
//					cfg.CID = dr["CID"].ToString();
//					cfg.Activo = Convert.ToBoolean(dr["Activo"]);
//					cfg.FechaSistema = Convert.ToDateTime(dr["FechaSistema"]);
//					cfg.Usuario = dr["Usuario"].ToString();
//
//					lista.Add(cfg);
//				}
//
//				dr.Close();
//			}
//			finally
//			{
//				cn.Close();
//			}
//
//			return lista;
//		}
//
//		// =========================
//		// READ (POR ID)
//		// =========================
//		public PinpadConfiguracion ObtenerPorId(Guid id)
//		{
//			PinpadConfiguracion cfg = null;
//			SqlConnection cn = new SqlConnection(_connectionString);
//
//			try
//			{
//				string sql = "SELECT * FROM Pinpad_Configuracion WHERE idPinpad_Configuracion = @Id";
//				SqlCommand cmd = new SqlCommand(sql, cn);
//				cmd.Parameters.Add(new SqlParameter("@Id", id));
//
//				cn.Open();
//				SqlDataReader dr = cmd.ExecuteReader();
//
//				if (dr.Read())
//				{
//					cfg = new PinpadConfiguracion();
//
//					cfg.IdPinpadConfiguracion = new Guid(dr["idPinpad_Configuracion"].ToString());
//					cfg.IdUsuario = Convert.ToInt32(dr["idUsuario"]);
//					cfg.NumeroCaja = dr["NumeroCaja"].ToString();
//					cfg.MID = dr["MID"].ToString();
//					cfg.TID = dr["TID"].ToString();
//					cfg.CID = dr["CID"].ToString();
//					cfg.Activo = Convert.ToBoolean(dr["Activo"]);
//					cfg.FechaSistema = Convert.ToDateTime(dr["FechaSistema"]);
//					cfg.Usuario = dr["Usuario"].ToString();
//				}
//
//				dr.Close();
//			}
//			finally
//			{
//				cn.Close();
//			}
//
//			return cfg;
//		}
//
//		// =========================
//		// UPDATE
//		// =========================
//		public void Actualizar(PinpadConfiguracion cfg)
//		{
//			SqlConnection cn = new SqlConnection(_connectionString);
//
//			try
//			{
//				string sql = @"
//				UPDATE Pinpad_Configuracion
//				SET 
//					idUsuario = @IdUsuario,
//					NumeroCaja = @NumeroCaja,
//					MID = @MID,
//					TID = @TID,
//					CID = @CID,
//					Activo = @Activo,
//					FechaModificacion = GETDATE(),
//					UsuarioModificacion = @Usuario
//				WHERE idPinpad_Configuracion = @Id";
//
//				SqlCommand cmd = new SqlCommand(sql, cn);
//
//				cmd.Parameters.Add(new SqlParameter("@Id", cfg.IdPinpadConfiguracion));
//				cmd.Parameters.Add(new SqlParameter("@IdUsuario", cfg.IdUsuario));
//				cmd.Parameters.Add(new SqlParameter("@NumeroCaja", cfg.NumeroCaja));
//				cmd.Parameters.Add(new SqlParameter("@MID", cfg.MID));
//				cmd.Parameters.Add(new SqlParameter("@TID", cfg.TID));
//				cmd.Parameters.Add(new SqlParameter("@CID", cfg.CID));
//				cmd.Parameters.Add(new SqlParameter("@Activo", cfg.Activo));
//
//				if (cfg.Usuario != null)
//					cmd.Parameters.Add(new SqlParameter("@Usuario", cfg.Usuario));
//				else
//					cmd.Parameters.Add(new SqlParameter("@Usuario", DBNull.Value));
//
//				cn.Open();
//				cmd.ExecuteNonQuery();
//			}
//			finally
//			{
//				cn.Close();
//			}
//		}
//
//		public PinpadConfiguracion ObtenerConfiguracionPorUsuario(int idUsuario)
//		{
//			SqlConnection cn = new SqlConnection(_connectionString);
//			PinpadConfiguracion cfg = null;
//
//			try
//			{
//				string sql = "SELECT TOP 1 * " +
//					"FROM Pinpad_Configuracion " +
//					"WHERE IdUsuario = @IdUsuario AND Activo = 1";
//
//				SqlCommand cmd = new SqlCommand(sql, cn);
//				cmd.Parameters.Add(new SqlParameter("@IdUsuario", SqlDbType.Int)).Value = idUsuario;
//
//				cn.Open();
//				SqlDataReader dr = cmd.ExecuteReader();
//
//				if (dr.Read())
//				{
//					cfg = new PinpadConfiguracion();
//
//					cfg.IdPinpadConfiguracion = (Guid)dr["idPinpad_Configuracion"];
//					cfg.IdUsuario = (int)dr["IdUsuario"];
//					cfg.NumeroCaja = dr["NumeroCaja"].ToString();
//					cfg.MID = dr["MID"].ToString();
//					cfg.TID = dr["TID"].ToString();
//					cfg.CID = dr["CID"].ToString();
//					cfg.Activo = (bool)dr["Activo"];
//				}
//
//				dr.Close();
//			}
//			finally
//			{
//				cn.Close();
//			}
//
//			return cfg;
//		}
//
//		// =========================
//		// DELETE
//		// =========================
//		public void Eliminar(Guid id)
//		{
//			SqlConnection cn = new SqlConnection(_connectionString);
//
//			try
//			{
//				string sql = "UPDATE Pinpad_Configuracion " +
//					"SET Activo = 0, UsuarioModificacion = @Usuario, FechaModificacion = GETDATE() " +
//					"WHERE idPinpad_Configuracion = @Id";
//
//				SqlCommand cmd = new SqlCommand(sql, cn);
//				cmd.Parameters.Add(new SqlParameter("@Id", id));
//				cmd.Parameters.Add(new SqlParameter("@Usuario", MenuLatinium.stUsuario));
//
//				cn.Open();
//				cmd.ExecuteNonQuery();
//			}
//			finally
//			{
//				cn.Close();
//			}
//		}
//
//		public bool ExisteDevice(string deviceId)
//		{
//			SqlConnection cn = new SqlConnection(_connectionString);
//
//			try
//			{
//				string sql = "SELECT COUNT(*) FROM Pinpad_Configuracion WHERE DeviceId = @DeviceId";
//
//				SqlCommand cmd = new SqlCommand(sql, cn);
//				cmd.Parameters.Add(new SqlParameter("@DeviceId", deviceId));
//
//				cn.Open();
//
//				int count = (int)cmd.ExecuteScalar();
//
//				return count > 0;
//			}
//			finally
//			{
//				cn.Close();
//			}
//		}
//
//		public PinpadConfiguracion ObtenerPorDeviceId(string deviceId)
//		{
//			SqlConnection cn = new SqlConnection(_connectionString);
//			PinpadConfiguracion cfg = null;
//
//			try
//			{
//				string sql = @"
//		SELECT TOP 1 *
//		FROM Pinpad_Configuracion
//		WHERE DeviceId = @DeviceId
//		AND Activo = 1";
//
//				SqlCommand cmd = new SqlCommand(sql, cn);
//				cmd.Parameters.Add(new SqlParameter("@DeviceId", deviceId));
//
//				cn.Open();
//				SqlDataReader dr = cmd.ExecuteReader();
//
//				if (dr.Read())
//				{
//					cfg = new PinpadConfiguracion();
//
//					cfg.IdPinpadConfiguracion = (Guid)dr["idPinpad_Configuracion"];
//					cfg.DeviceId = dr["DeviceId"].ToString();
//					cfg.NombrePC = dr["NombrePC"].ToString();
//
//					cfg.Bodega = Convert.ToInt32(dr["Bodega"]);
//					cfg.NumeroCaja = dr["NumeroCaja"].ToString();
//
//					cfg.IpPinPad = dr["IpPinPad"].ToString();
//					cfg.PuertoPinPad = Convert.ToInt32(dr["PuertoPinPad"]);
//					cfg.TimeoutPinPad = Convert.ToInt32(dr["TimeoutPinPad"]);
//
//					cfg.MID = dr["MID"].ToString();
//					cfg.TID = dr["TID"].ToString();
//					cfg.CID = dr["CID"].ToString();
//
//					cfg.Activo = Convert.ToBoolean(dr["Activo"]);
//				}
//
//				dr.Close();
//			}
//			finally
//			{
//				cn.Close();
//			}
//
//			return cfg;
//		}
//
//		public PinpadConfiguracion GetPinpadActual(string deviceId)
//		{
//			return ObtenerPorDeviceId(deviceId);
//		}
//
//		public PinpadConfiguracion ResolverConfiguracionActual(string deviceId, int bodega)
//		{
//			SqlConnection cn = new SqlConnection(_connectionString);
//
//			try
//			{
//				string sql = @"
//            SELECT TOP 1 *
//            FROM Pinpad_Configuracion
//            WHERE DeviceId = @DeviceId
//            AND Bodega = @Bodega
//            AND Activo = 1";
//
//				SqlCommand cmd = new SqlCommand(sql, cn);
//
//				cmd.Parameters.Add(new SqlParameter("@DeviceId", SqlDbType.VarChar, 50));
//				cmd.Parameters["@DeviceId"].Value = deviceId;
//
//				cmd.Parameters.Add(new SqlParameter("@Bodega", SqlDbType.Int));
//				cmd.Parameters["@Bodega"].Value = bodega;
//
//				cn.Open();
//
//				SqlDataReader dr = cmd.ExecuteReader();
//
//				if (!dr.Read())
//				{
//					dr.Close();
//					return null;
//				}
//
//				PinpadConfiguracion config = new PinpadConfiguracion();
//
//				config.IdPinpadConfiguracion = new Guid(dr["idPinpad_Configuracion"].ToString());
//				config.DeviceId = dr["DeviceId"].ToString();
//				config.NombrePC = dr["NombrePC"].ToString();
//				config.Bodega = Convert.ToInt32(dr["Bodega"]);
//				config.NumeroCaja = dr["NumeroCaja"].ToString();
//				config.IpPinPad = dr["IpPinPad"].ToString();
//				config.PuertoPinPad = Convert.ToInt32(dr["PuertoPinPad"]);
//				config.TimeoutPinPad = Convert.ToInt32(dr["TimeoutPinPad"]);
//				config.MID = dr["MID"].ToString();
//				config.TID = dr["TID"].ToString();
//				config.CID = dr["CID"].ToString();
//				config.Activo = Convert.ToBoolean(dr["Activo"]);
//
//				dr.Close();
//				return config;
//			}
//			finally
//			{
//				if (cn != null)
//				{
//					cn.Close();
//					cn.Dispose();
//				}
//			}
//		}
//
//	}
//}

using System;
using System.Collections;
using System.Data;
using System.Data.SqlClient;
using Latinium.Models.Pinpad.Configuracion;

namespace Latinium.Services.Pinpad.Configuracion
{
	public class PinpadConfiguracionRepository
	{
		private string _connectionString;

		public PinpadConfiguracionRepository(string connectionString)
		{
			_connectionString = connectionString;
		}

		// =========================
		// CREATE
		// =========================
		public void Insertar(PinpadConfiguracion cfg)
		{
			SqlConnection cn = new SqlConnection(_connectionString);

			try
			{
				string sql = @"
		INSERT INTO Pinpad_Configuracion
		(
			idPinpad_Configuracion,
			DeviceId,
			NombrePC,
			Bodega,
			NumeroCaja,
			IpPinPad,
			PuertoPinPad,
			TimeoutPinPad,
			MID,
			TID,
			CID,
			Activo,
			FechaSistema,
			Usuario
		)
		VALUES
		(
			@Id,
			@DeviceId,
			@NombrePC,
			@Bodega,
			@NumeroCaja,
			@IpPinPad,
			@PuertoPinPad,
			@TimeoutPinPad,
			@MID,
			@TID,
			@CID,
			@Activo,
			GETDATE(),
			@Usuario
		)";

				SqlCommand cmd = new SqlCommand(sql, cn);

				cmd.Parameters.Add(new SqlParameter("@Id", Guid.NewGuid()));
				cmd.Parameters.Add(new SqlParameter("@DeviceId", cfg.DeviceId));
				cmd.Parameters.Add(new SqlParameter("@NombrePC", cfg.NombrePC));
				cmd.Parameters.Add(new SqlParameter("@Bodega", cfg.Bodega));
				cmd.Parameters.Add(new SqlParameter("@NumeroCaja", cfg.NumeroCaja));

				cmd.Parameters.Add(new SqlParameter("@IpPinPad", cfg.IpPinPad));
				cmd.Parameters.Add(new SqlParameter("@PuertoPinPad", cfg.PuertoPinPad));
				cmd.Parameters.Add(new SqlParameter("@TimeoutPinPad", cfg.TimeoutPinPad));

				cmd.Parameters.Add(new SqlParameter("@MID", cfg.MID));
				cmd.Parameters.Add(new SqlParameter("@TID", cfg.TID));
				cmd.Parameters.Add(new SqlParameter("@CID", cfg.CID));

				cmd.Parameters.Add(new SqlParameter("@Activo", cfg.Activo));

				object usuario = DBNull.Value;
				if (cfg.Usuario != null)
					usuario = cfg.Usuario;

				cmd.Parameters.Add(new SqlParameter("@Usuario", usuario));

				cn.Open();
				cmd.ExecuteNonQuery();
			}
			catch (Exception ex)
			{
				// SOLO IMPRIME ERROR
				Console.WriteLine("ERROR INSERTAR PINPAD: " + ex.Message);
			}
			finally
			{
				cn.Close();
			}
		}
		// =========================
		// LISTAR
		// =========================
		public ArrayList Listar()
		{
			ArrayList lista = new ArrayList();
			SqlConnection cn = new SqlConnection(_connectionString);

			try
			{
				string sql = @"
        SELECT 
            pc.idPinpad_Configuracion,
            pc.DeviceId,
            pc.NombrePC,
            pc.Bodega,
            b.Nombre AS NombreBodega,
            pc.NumeroCaja,
            pc.IpPinPad,
            pc.PuertoPinPad,
            pc.TimeoutPinPad,
            pc.MID,
            pc.TID,
            pc.CID,
            pc.Activo
        FROM Pinpad_Configuracion pc
        LEFT JOIN Bodega b ON pc.Bodega = b.Bodega
        WHERE pc.Activo = 1";

				SqlCommand cmd = new SqlCommand(sql, cn);

				cn.Open();
				SqlDataReader dr = cmd.ExecuteReader();

				while (dr.Read())
				{
					PinpadConfiguracion cfg = new PinpadConfiguracion();

					cfg.IdPinpadConfiguracion = (Guid)dr["idPinpad_Configuracion"];
					cfg.DeviceId = dr["DeviceId"].ToString();
					cfg.NombrePC = dr["NombrePC"].ToString();

					cfg.Bodega = Convert.ToInt32(dr["Bodega"]);

					// NUEVO CAMPO (NO OLVIDES AGREGARLO AL MODELO)
					cfg.NombreBodega = dr["NombreBodega"].ToString();

					cfg.NumeroCaja = dr["NumeroCaja"].ToString();
					cfg.IpPinPad = dr["IpPinPad"].ToString();
					cfg.PuertoPinPad = Convert.ToInt32(dr["PuertoPinPad"]);
					cfg.TimeoutPinPad = Convert.ToInt32(dr["TimeoutPinPad"]);

					cfg.MID = dr["MID"].ToString();
					cfg.TID = dr["TID"].ToString();
					cfg.CID = dr["CID"].ToString();

					cfg.Activo = Convert.ToBoolean(dr["Activo"]);

					lista.Add(cfg);
				}

				dr.Close();
			}
			finally
			{
				cn.Close();
			}

			return lista;
		}

		// =========================
		// POR ID
		// =========================
		public PinpadConfiguracion ObtenerPorId(Guid id)
		{
			PinpadConfiguracion cfg = null;
			SqlConnection cn = new SqlConnection(_connectionString);

			try
			{
				string sql = "SELECT * FROM Pinpad_Configuracion WHERE idPinpad_Configuracion = @Id";
				SqlCommand cmd = new SqlCommand(sql, cn);
				cmd.Parameters.Add(new SqlParameter("@Id", id));

				cn.Open();
				SqlDataReader dr = cmd.ExecuteReader();

				if (dr.Read())
				{
					cfg = new PinpadConfiguracion();

					cfg.IdPinpadConfiguracion = (Guid)dr["idPinpad_Configuracion"];
					cfg.DeviceId = dr["DeviceId"].ToString();
					cfg.NombrePC = dr["NombrePC"].ToString();
					cfg.Bodega = Convert.ToInt32(dr["Bodega"]);
					cfg.NumeroCaja = dr["NumeroCaja"].ToString();

					cfg.IpPinPad = dr["IpPinPad"].ToString();
					cfg.PuertoPinPad = Convert.ToInt32(dr["PuertoPinPad"]);
					cfg.TimeoutPinPad = Convert.ToInt32(dr["TimeoutPinPad"]);

					cfg.MID = dr["MID"].ToString();
					cfg.TID = dr["TID"].ToString();
					cfg.CID = dr["CID"].ToString();

					cfg.Activo = Convert.ToBoolean(dr["Activo"]);
				}

				dr.Close();
			}
			finally
			{
				cn.Close();
			}

			return cfg;
		}

		// =========================
		// UPDATE
		// =========================
		public void Actualizar(PinpadConfiguracion cfg)
		{
			SqlConnection cn = new SqlConnection(_connectionString);

			try
			{
				string sql = @"
				UPDATE Pinpad_Configuracion
				SET 
					DeviceId = @DeviceId,
					NombrePC = @NombrePC,
					Bodega = @Bodega,
					NumeroCaja = @NumeroCaja,
					IpPinPad = @IpPinPad,
					PuertoPinPad = @PuertoPinPad,
					TimeoutPinPad = @TimeoutPinPad,
					MID = @MID,
					TID = @TID,
					CID = @CID,
					Activo = @Activo,
					FechaModificacion = GETDATE(),
					UsuarioModificacion = @Usuario
				WHERE idPinpad_Configuracion = @Id";

				SqlCommand cmd = new SqlCommand(sql, cn);

				cmd.Parameters.Add(new SqlParameter("@Id", cfg.IdPinpadConfiguracion));
				cmd.Parameters.Add(new SqlParameter("@DeviceId", cfg.DeviceId));
				cmd.Parameters.Add(new SqlParameter("@NombrePC", cfg.NombrePC));
				cmd.Parameters.Add(new SqlParameter("@Bodega", cfg.Bodega));
				cmd.Parameters.Add(new SqlParameter("@NumeroCaja", cfg.NumeroCaja));

				cmd.Parameters.Add(new SqlParameter("@IpPinPad", cfg.IpPinPad));
				cmd.Parameters.Add(new SqlParameter("@PuertoPinPad", cfg.PuertoPinPad));
				cmd.Parameters.Add(new SqlParameter("@TimeoutPinPad", cfg.TimeoutPinPad));

				cmd.Parameters.Add(new SqlParameter("@MID", cfg.MID));
				cmd.Parameters.Add(new SqlParameter("@TID", cfg.TID));
				cmd.Parameters.Add(new SqlParameter("@CID", cfg.CID));

				cmd.Parameters.Add(new SqlParameter("@Activo", cfg.Activo));
				if (cfg.UsuarioModificacion != null)
				{
					cmd.Parameters.Add(new SqlParameter("@Usuario", cfg.UsuarioModificacion));
				}
				else
				{
					cmd.Parameters.Add(new SqlParameter("@Usuario", DBNull.Value));
				}

				cn.Open();
				cmd.ExecuteNonQuery();
			}
			finally
			{
				cn.Close();
			}
		}

		// =========================
		// DELETE (SOFT)
		// =========================
		public void Eliminar(Guid id)
		{
			SqlConnection cn = new SqlConnection(_connectionString);

			try
			{
				string sql = @"
				UPDATE Pinpad_Configuracion
				SET Activo = 0,
					UsuarioModificacion = @Usuario,
					FechaModificacion = GETDATE()
				WHERE idPinpad_Configuracion = @Id";

				SqlCommand cmd = new SqlCommand(sql, cn);
				cmd.Parameters.Add(new SqlParameter("@Id", id));
				cmd.Parameters.Add(new SqlParameter("@Usuario", MenuLatinium.stUsuario));

				cn.Open();
				cmd.ExecuteNonQuery();
			}
			finally
			{
				cn.Close();
			}
		}



		public PinpadConfiguracion ObtenerPorDeviceId(string deviceId)
		{
			SqlConnection cn = new SqlConnection(_connectionString);
			PinpadConfiguracion cfg = null;

			try
			{
				string sql = @"
				SELECT TOP 1 *
				FROM Pinpad_Configuracion
				WHERE DeviceId = @DeviceId
				AND Activo = 1";

				SqlCommand cmd = new SqlCommand(sql, cn);
				cmd.Parameters.Add(new SqlParameter("@DeviceId", deviceId));

				cn.Open();
				SqlDataReader dr = cmd.ExecuteReader();

				if (dr.Read())
				{
					cfg = new PinpadConfiguracion();

					cfg.IdPinpadConfiguracion = (Guid)dr["idPinpad_Configuracion"];
					cfg.DeviceId = dr["DeviceId"].ToString();
					cfg.NombrePC = dr["NombrePC"].ToString();
					cfg.Bodega = Convert.ToInt32(dr["Bodega"]);
					cfg.NumeroCaja = dr["NumeroCaja"].ToString();

					cfg.IpPinPad = dr["IpPinPad"].ToString();
					cfg.PuertoPinPad = Convert.ToInt32(dr["PuertoPinPad"]);
					cfg.TimeoutPinPad = Convert.ToInt32(dr["TimeoutPinPad"]);

					cfg.MID = dr["MID"].ToString();
					cfg.TID = dr["TID"].ToString();
					cfg.CID = dr["CID"].ToString();

					cfg.Activo = Convert.ToBoolean(dr["Activo"]);
				}

				dr.Close();
			}
			finally
			{
				cn.Close();
			}

			return cfg;
		}

		public PinpadConfiguracion ResolverConfiguracionActual(string deviceId, int bodega)
		{
			return ObtenerPorDeviceId(deviceId);
		}

		public int ObtenerSiguienteNumeroCajaGlobal()
		{
			SqlConnection cn = new SqlConnection(_connectionString);

			try
			{
				cn.Open();

				string sql = @"
				SELECT ISNULL(MAX(CAST(SUBSTRING(NumeroCaja, 6, 10) AS INT)), 0) + 1
				FROM Pinpad_Configuracion
				WHERE NumeroCaja LIKE 'CAJA-%'";

				SqlCommand cmd = new SqlCommand(sql, cn);

				object result = cmd.ExecuteScalar();

				return Convert.ToInt32(result);
			}
			finally
			{
				cn.Close();
			}
		}
	}
}