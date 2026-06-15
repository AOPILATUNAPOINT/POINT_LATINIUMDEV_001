using System;
using System.Collections;
using Latinium.Models.Requerimiento;
using System.Data;

namespace Latinium.Services.Requerimiento
{
	public class RequerimientoMenuMemoriaService
	{
		private ArrayList _items;

		public RequerimientoMenuMemoriaService()
		{
			_items = new ArrayList();
		}

		public void ProcesarSeleccion(
			int idSeg,
			bool seleccionado,
			string modulo,
			string pantalla,
			string subPantalla)
		{
			RequerimientoMenuSeguridadModel existente = Buscar(idSeg);

			if (seleccionado)
			{
				if (existente == null)
				{
					RequerimientoMenuSeguridadModel item =
						new RequerimientoMenuSeguridadModel();

					item.id_seg = idSeg;
					item.Modulo = modulo;
					item.Pantalla = pantalla;
					item.SubPantalla = subPantalla;

					_items.Add(item);
				}
			}
			else
			{
				if (existente != null)
				{
					_items.Remove(existente);
				}
			}
		}

		private RequerimientoMenuSeguridadModel Buscar(int idSeg)
		{
			foreach (RequerimientoMenuSeguridadModel item in _items)
			{
				if (item.id_seg == idSeg)
				{
					return item;
				}
			}
			return null;
		}

		public ArrayList ObtenerTodo()
		{
			return _items;
		}

		public void Limpiar()
		{
			_items.Clear();
		}

		public DataTable GenerarDataTable()
		{
			DataTable dt = new DataTable("RequerimientoMenu");

			dt.Columns.Add("id_seg", typeof(int));
			dt.Columns.Add("Modulo", typeof(string));
			dt.Columns.Add("Pantalla", typeof(string));
			dt.Columns.Add("SubPantalla", typeof(string));

			foreach (RequerimientoMenuSeguridadModel item in _items)
			{
				DataRow row = dt.NewRow();

				row["id_seg"] = item.id_seg;
				row["Modulo"] = item.Modulo;
				row["Pantalla"] = item.Pantalla;
				row["SubPantalla"] = item.SubPantalla;

				dt.Rows.Add(row);
			}

			return dt;
		}

		public bool Existe(int idSeg)
		{
			return Buscar(idSeg) != null;
		}
	}
}
