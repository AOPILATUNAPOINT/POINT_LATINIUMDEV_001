using System;
using System.Data;
using Latinium.Services;
using System.Collections;

namespace Latinium.Services
{
	public class RequerimientoAccesoService // Quitamos el 'static'
	{
		public static bool UsuarioTieneAcceso(int idUsuario, int idRequerimientoDepartamento)
		{
			int idGrupo = -1;
            
			try
			{
				// Instancias creadas dentro del método
				RequerimientoUsuarioService usuarioService = new RequerimientoUsuarioService();
				RequerimientoDepartamentoRolService deptoRolService = new RequerimientoDepartamentoRolService(frmRequerimiento.sconexionPoint);

				idGrupo = usuarioService.ObtenerGrupoPorUsuario(idUsuario);
                
				if (idGrupo == -1)
					return false;

				return deptoRolService.TieneAcceso(idGrupo, idRequerimientoDepartamento);
			}
			catch (Exception ex)
			{
				System.Diagnostics.Debug.WriteLine("Error verificando acceso: " + ex.Message);
				return false;
			}
		}

		public static ArrayList ObtenerTicketerasPermitidas(int idUsuario)
		{
			ArrayList ticketerasPermitidas = new ArrayList();

			try
			{
				RequerimientoUsuarioService usuarioService = new RequerimientoUsuarioService();
				RequerimientoDepartamentoRolService deptoRolService = new RequerimientoDepartamentoRolService(frmRequerimiento.sconexionPoint);

				// Obtener el grupo del usuario
				int idGrupo = usuarioService.ObtenerGrupoPorUsuario(idUsuario);
				if (idGrupo == -1)
					return ticketerasPermitidas;

				// Obtener todos los departamentos habilitados
				DataTable todosLosDepartamentos = deptoRolService.ObtenerTodosLosDepartamentos();

				// Si es administrador (grupo 1), devolver todas las ticketeras
				if (idGrupo == frmRequerimiento.idGrupoAdministrador)
				{
					foreach (DataRow fila in todosLosDepartamentos.Rows)
						ticketerasPermitidas.Add(fila);

					return ticketerasPermitidas;
				}

				// Si no es administrador, validar ticketeras
				foreach (DataRow fila in todosLosDepartamentos.Rows)
				{
					int idRequerimientoDepartamento = Convert.ToInt32(fila["idRequerimientoDepartamento"]);

					// Validar acceso por grupo
					bool tieneAcceso = deptoRolService.TieneAcceso(idGrupo, idRequerimientoDepartamento);

					// Validación especial: si es el departamento con ID 1, validar que esté en la tabla de jefes
					if (idRequerimientoDepartamento == 1)
					{
						bool esJefePermitido = deptoRolService.EsUsuarioPermitidoEnSoloJefes(idUsuario);
						if (esJefePermitido)
							ticketerasPermitidas.Add(fila);
					}
					else
					{
						// Otros departamentos, igual que antes
						if (tieneAcceso)
							ticketerasPermitidas.Add(fila);
					}
				}
			}
			catch (Exception ex)
			{
				System.Diagnostics.Debug.WriteLine("Error al obtener ticketeras permitidas: " + ex.Message);
			}

			return ticketerasPermitidas;
		}

	}
}
