using System;
using System.Collections;
using System.Data;
using System.Data.SqlClient;
using Latinium.Models;
using C1.Data;

namespace Latinium.Services
{
	/// <summary>
	/// Servicio para gestión de grupos web
	/// </summary>
	public class WEBGrupoService
	{
		private C1DataSet _dsSeteo;

		public WEBGrupoService(C1DataSet dsSeteo)
		{
			_dsSeteo = dsSeteo;
		}

		public ArrayList ListarGrupos()
		{
			ArrayList lista = new ArrayList();

			_dsSeteo.Schema.Connections[0].Open();
			C1.Data.SchemaObjects.Connection miConexion =
				_dsSeteo.Schema.Connections[0];

			SqlConnection cn =
				new SqlConnection(miConexion.ConnectionString);

			SqlCommand cmd =
				new SqlCommand("WEB_GruposCon", cn);

			cmd.CommandType = CommandType.StoredProcedure;

			try
			{
				cn.Open();
				SqlDataReader dr = cmd.ExecuteReader();

				while (dr.Read())
				{
					WebGrupoModels grupo = new WebGrupoModels();
					grupo.Cargar(dr);
					lista.Add(grupo);
				}

				dr.Close();
			}
			finally
			{
				cn.Close();
			}

			return lista;
		}
	}
}
