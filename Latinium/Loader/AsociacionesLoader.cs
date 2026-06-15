using System;
using System.Collections;
using System.Windows.Forms;
using Latinium.Models;
using Latinium.Services;

public class AsociacionesLoader
{
	public delegate void MostrarCheckMaestroDelegate(bool mostrar);

	public static void CargarAsociaciones(int idDepartamento, DataGrid dg, string connectionString, Label lblMensaje, MostrarCheckMaestroDelegate mostrarCheckMaestroCallback)
	{
		try
		{
			RequerimientoDepartamentoRolService service = new RequerimientoDepartamentoRolService(connectionString);
			ArrayList asociaciones = service.ObtenerAsociacionesPorDepartamento(idDepartamento);

			if (asociaciones.Count > 0)
			{
				dg.DataSource = asociaciones;
				ConfigurarEstilos(dg);
				lblMensaje.Text = "";
				if (mostrarCheckMaestroCallback != null)
				{
					mostrarCheckMaestroCallback(true);
				}
			}
			else
			{
				dg.DataSource = null;
				lblMensaje.Text = "No se encontraron asociaciones para este departamento.";
				if (mostrarCheckMaestroCallback != null) mostrarCheckMaestroCallback(false);
			}
		}
		catch (Exception ex)
		{
			lblMensaje.Text = "Error al cargar las asociaciones: " + ex.Message;
			if (mostrarCheckMaestroCallback != null) mostrarCheckMaestroCallback(false);
		}
	}

	private static void ConfigurarEstilos(DataGrid dg)
	{
		DataGridTableStyle style = new DataGridTableStyle();
		style.MappingName = "ArrayList"; // Para ArrayList necesitas este MappingName

		DataGridTextBoxColumn colNombreGrupo = new DataGridTextBoxColumn();
		colNombreGrupo.MappingName = "NombreGrupo";
		colNombreGrupo.HeaderText = "Grupo";
		colNombreGrupo.Width = dg.Width - 100;
		colNombreGrupo.ReadOnly = true;
		style.GridColumnStyles.Add(colNombreGrupo);

		DataGridBoolColumn colEstado = new DataGridBoolColumn();
		colEstado.MappingName = "Estado";
		colEstado.HeaderText = "Acceso";
		colEstado.Width = 75;
		style.GridColumnStyles.Add(colEstado);

		dg.TableStyles.Clear();
		dg.TableStyles.Add(style);
	}
}
