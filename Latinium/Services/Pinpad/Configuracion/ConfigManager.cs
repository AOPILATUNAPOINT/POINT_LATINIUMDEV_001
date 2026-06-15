using System;
using System.Data;
using System.Collections;
using System.Data.SqlClient;

namespace Latinium.Services.Pinpad.Configuracion
{
	public class ConfigManager
	{
		private static Hashtable _configCache = new Hashtable();

		// PON TU CONNECTION STRING AQUÍ O CÁRGALA DESDE CONFIG
		private static string _connectionString = MenuLatinium.sconexionEmpresa;

		public static void SetConnectionString(string cs)
		{
			_connectionString = cs;
		}

		public static void LoadConfig()
		{
			_configCache.Clear();

			SqlConnection cn = new SqlConnection(_connectionString);
			SqlCommand cmd = new SqlCommand(
				"SELECT Clave, Valor FROM PinPad_Sistema WHERE Activo = 1", cn);

			try
			{
				cn.Open();

				SqlDataReader dr = cmd.ExecuteReader();

				while (dr.Read())
				{
					string clave = dr["Clave"].ToString();
					string valor = dr["Valor"].ToString();

					if (!_configCache.ContainsKey(clave))
						_configCache.Add(clave, valor);
					else
						_configCache[clave] = valor;
				}

				dr.Close();
			}
			catch (Exception ex)
			{
				throw new Exception("Error cargando configuración PINPAD: " + ex.Message);
			}
			finally
			{
				if (cn.State == ConnectionState.Open)
					cn.Close();
			}
		}

		public static string Get(string clave)
		{
			if (clave == null || clave == "")
				return "";

			if (_configCache.ContainsKey(clave))
				return _configCache[clave].ToString();

			return "";
		}

		public static int GetInt(string clave)
		{
			string val = Get(clave);

			if (val == null || val == "")
				return 0;

			try
			{
				return Convert.ToInt32(val);
			}
			catch
			{
				return 0;
			}
		}

		public static bool GetBool(string clave)
		{
			string val = Get(clave);

			if (val == "1" || val.ToLower() == "true")
				return true;

			return false;
		}
	}
}