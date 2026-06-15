using System;
using Latinium.Models.Pinpad.Configuracion;
using System.Data.SqlClient;

namespace Latinium.Services.Pinpad.Configuracion
{
	/// <summary>
	/// Descripción breve de PinpadBootstrapService.
	/// </summary>
	public class PinpadBootstrapService
	{
		private PinpadConfiguracionRepository _repo;

		public PinpadBootstrapService(PinpadConfiguracionRepository repo)
		{
			_repo = repo;
		}

		public PinpadConfiguracion Inicializar(int bodega, string nombrePC)
		{
			// 1. Config global debe estar cargada
			ConfigManager.LoadConfig();

			// 2. generar device si no existe
			DeviceIdManager.GenerarSiNoExiste();

			string deviceId = DeviceIdManager.ObtenerDeviceId();

			if (deviceId == null)
				throw new Exception("DeviceId inválido");

			// 3. validar si existe config máquina
			PinpadConfiguracion cfg = _repo.ObtenerPorDeviceId(deviceId);

			if (cfg != null)
				return cfg;

			// 4. crear config automática usando CONFIG GLOBAL
			cfg = new PinpadConfiguracion();

			cfg.IdPinpadConfiguracion = Guid.NewGuid();
			cfg.DeviceId = deviceId;
			cfg.NombrePC = nombrePC;
			cfg.Bodega = bodega;

			cfg.NumeroCaja = GenerarNombreCaja();

			// USANDO CONFIG GLOBAL (NO QUEMADO)
			cfg.IpPinPad = "127.0.0.1";
			cfg.PuertoPinPad = ConfigManager.GetInt("PuertoDefaultPinpad");
			cfg.TimeoutPinPad = ConfigManager.GetInt("TimeoutPinpad");

			cfg.MID = "000000";
			cfg.TID = "000000";
			cfg.CID = "000000";

			cfg.Activo = true;
			cfg.Usuario = Environment.UserName;

			_repo.Insertar(cfg);

			return cfg;
		}

		private string GenerarNombreCaja()
		{
			int correlativo = _repo.ObtenerSiguienteNumeroCajaGlobal();

			return string.Format("CAJA-{0:0000}", correlativo);
		}
	}
}
