using System;
using System.Data;
using System.Runtime.InteropServices;

namespace Latinium.Services.OrdenDeCompra.Excel
{
	public class OrdenDeCompraExcel
	{
		public void GenerarExcel(DataSet ds, string rutaArchivo)
		{
			if (ds == null || ds.Tables.Count < 2)
				throw new Exception("DataSet inválido");

			object xlApp = null;
			object wb = null;
			object wsOrden = null;
			object wsProv = null;

			try
			{
				// =========================
				// INICIAR EXCEL (COM PURO)
				// =========================
				Type tExcel = Type.GetTypeFromProgID("Excel.Application");
				xlApp = Activator.CreateInstance(tExcel);

				tExcel.InvokeMember("Visible", System.Reflection.BindingFlags.SetProperty, null, xlApp, new object[] { false });

				object workbooks = tExcel.InvokeMember("Workbooks", System.Reflection.BindingFlags.GetProperty, null, xlApp, null);
				wb = workbooks.GetType().InvokeMember("Add", System.Reflection.BindingFlags.InvokeMethod, null, workbooks, null);

				object sheets = wb.GetType().InvokeMember("Worksheets", System.Reflection.BindingFlags.GetProperty, null, wb, null);

				// =====================================================
				// HOJA 1 - ORDEN
				// =====================================================
				wsOrden = sheets.GetType().InvokeMember("Item", System.Reflection.BindingFlags.GetProperty, null, sheets, new object[] { 1 });
				wsOrden.GetType().InvokeMember("Name", System.Reflection.BindingFlags.SetProperty, null, wsOrden, new object[] { "ORDEN" });

				DataRow cab = ds.Tables[0].Rows[0];

				SetCell(wsOrden, 1, 1, "ORDEN DE COMPRA");

				// proveedor en celda B3
				SetCell(wsOrden, 3, 1, "Proveedor");
				SetCell(wsOrden, 3, 2, cab["Proveedor"]);

				// detalle inicio fila 6
				int fila = 6;

				foreach (DataRow r in ds.Tables[1].Rows)
				{
					SetCell(wsOrden, fila, 1, r["Articulo"]);
					SetCell(wsOrden, fila, 2, r["Descripcion"]);
					SetCell(wsOrden, fila, 3, r["Cantidad"]);
					SetCell(wsOrden, fila, 4, r["Precio"]);
					fila++;
				}

				// =====================================================
				// HOJA 2 - PROVEEDORES
				// =====================================================
				wsProv = sheets.GetType().InvokeMember("Add", System.Reflection.BindingFlags.InvokeMethod, null, sheets, null);
				wsProv.GetType().InvokeMember("Name", System.Reflection.BindingFlags.SetProperty, null, wsProv, new object[] { "PROVEEDORES" });

				DataTable dtProv = CargarProveedores();

				SetCell(wsProv, 1, 1, "ID");
				SetCell(wsProv, 1, 2, "NOMBRE");

				int i = 2;

				foreach (DataRow r in dtProv.Rows)
				{
					SetCell(wsProv, i, 1, r["IdCliente"]);
					SetCell(wsProv, i, 2, r["Nombre"]);
					i++;
				}

				// =====================================================
				// COMBO REAL (DATA VALIDATION)
				// =====================================================
				object range = GetRange(wsOrden, "B3");
				object validation = range.GetType().InvokeMember("Validation", System.Reflection.BindingFlags.GetProperty, null, range, null);

				validation.GetType().InvokeMember("Delete", System.Reflection.BindingFlags.InvokeMethod, null, validation, null);

				string formula = "=PROVEEDORES!$B$2:$B$" + (dtProv.Rows.Count + 1);

				validation.GetType().InvokeMember(
					"Add",
					System.Reflection.BindingFlags.InvokeMethod,
					null,
					validation,
					new object[] { 3, 1, 1, formula }
					);

				validation.GetType().InvokeMember("InCellDropdown", System.Reflection.BindingFlags.SetProperty, null, validation, new object[] { true });

				// =====================================================
				// GUARDAR
				// =====================================================
				wb.GetType().InvokeMember("SaveAs", System.Reflection.BindingFlags.InvokeMethod, null, wb, new object[] { rutaArchivo });

				wb.GetType().InvokeMember("Close", System.Reflection.BindingFlags.InvokeMethod, null, wb, null);
				xlApp.GetType().InvokeMember("Quit", System.Reflection.BindingFlags.InvokeMethod, null, xlApp, null);
			}
			finally
			{
				Release(xlApp);
				Release(wb);
				Release(wsOrden);
				Release(wsProv);
			}
		}

		// =========================
		// HELPERS .NET 1.1
		// =========================

		private void SetCell(object sheet, int row, int col, object value)
		{
			object cells = sheet.GetType().InvokeMember("Cells", System.Reflection.BindingFlags.GetProperty, null, sheet, new object[] { row, col });

			cells.GetType().InvokeMember("Value", System.Reflection.BindingFlags.SetProperty, null, cells, new object[] { value });
		}

		private object GetRange(object sheet, string address)
		{
			return sheet.GetType().InvokeMember("Range", System.Reflection.BindingFlags.GetProperty, null, sheet, new object[] { address });
		}

		private void Release(object obj)
		{
			try
			{
				if (obj != null)
					Marshal.ReleaseComObject(obj);
			}
			catch { }
		}

		// =========================
		// PROVEEDORES
		// =========================
		public DataTable CargarProveedores()
		{
			DataTable dt = new DataTable();
			dt.Columns.Add("IdCliente", typeof(int));
			dt.Columns.Add("Nombre", typeof(string));

			DataTable dtTemp = FuncionesProcedimientos.dtGeneral(
				"SELECT IdCliente, Nombre FROM Cliente WHERE Proveedor = 1 ORDER BY Nombre"
				);

			foreach (DataRow r in dtTemp.Rows)
			{
				DataRow row = dt.NewRow();
				row["IdCliente"] = r["IdCliente"];
				row["Nombre"] = r["Nombre"];
				dt.Rows.Add(row);
			}

			return dt;
		}
	}
}