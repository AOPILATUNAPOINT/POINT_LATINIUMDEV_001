using System;
using System.Collections;
using System.Windows.Forms;
using Latinium.Models;
using Latinium.Services;

namespace Latinium.Loaders
{
	public class DepartamentoLoader
	{
		public static void CargarDepartamentos(ComboBox combo, string connectionString, Label lblMensaje)
		{
			try
			{
				RequerimientoDepartamentoService service = new RequerimientoDepartamentoService(connectionString);
				ArrayList departamentos = service.ListarDepartamentosActivos();

				combo.Items.Clear();

				// Item por defecto
				RequerimientoDepartamento defaultItem = new RequerimientoDepartamento();
				defaultItem.Descripcion = "- Seleccione -";
				defaultItem.IdRequerimientoDepartamento = 0;
				combo.Items.Add(defaultItem);

				foreach (RequerimientoDepartamento depto in departamentos)
				{
					combo.Items.Add(depto);
				}

				// En .NET 1.1 no existen DisplayMember ni ValueMember en ComboBox clásico
				// Debes sobrescribir ToString() en la clase RequerimientoDepartamento para mostrar texto
			}
			catch (Exception ex)
			{
				lblMensaje.Text = "Error al cargar los departamentos: " + ex.Message;
			}
		}
	}
}
