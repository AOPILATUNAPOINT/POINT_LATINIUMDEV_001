using System;
using System.Data;
using System.Data.SqlClient;
using System.Windows.Forms;
using Infragistics;
using Infragistics.Win.UltraWinGrid;
using System.Globalization;
using System.Data.OleDb;
using System.Data.Odbc;

namespace Latinium
{
	/// <summary>
	/// Descripción breve de FuncionesInventario.
	/// </summary>
	public class FuncionesInventario
	{
		public FuncionesInventario()
		{
			//
			// TODO: agregar aquí la lógica del constructor
			//
		}

		static public void ActualizaPosicionArticulosSeriales(System.EventArgs e, Infragistics.Win.UltraWinGrid.UltraGrid uGrid, Infragistics.Win.UltraWinGrid.UltraGrid uGridSeriales)
		{		
			int iPosicion = 1;
				
			foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in uGrid.Rows.All)
			{
				if (uGridSeriales.Rows.Count > 0)
				{
					foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr1 in uGridSeriales.Rows.All)
					{
						if (((int)dr.Cells["idArticulo"].Value == (int)dr1.Cells["idArticulo"].Value) && ((int)dr.Cells["Posicion"].Value == (int)dr1.Cells["Posicion"].Value))
						{								
							dr1.Cells["Posicion"].Value = iPosicion;								
						}						
					}						
				}
				else
				{
					dr.Cells["Posicion"].Value = iPosicion;
					iPosicion++;						
				}

				dr.Cells["Posicion"].Value = iPosicion;
				iPosicion++;					
			}	
		}

		static public bool vSerialesIngresadasGridAbreInventario(Infragistics.Win.UltraWinGrid.UltraGrid uGridArt, Infragistics.Win.UltraWinGrid.UltraGrid uGrid, bool bNuevo)
		{			
			if (!bNuevo && uGridArt.ActiveRow.Cells["idArticulo"].Value != System.DBNull.Value) return false;

			if (uGrid.Rows.Count == 0) return true;
			
			if (uGridArt.ActiveRow.Cells["idArticulo"].Value == System.DBNull.Value) return true;

			int iConteo = 0;

			foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in uGrid.Rows.All)
			{
				if ((int)uGridArt.ActiveRow.Cells["idArticulo"].Value == (int)dr.Cells["idArticulo"].Value 
					&& (int)uGridArt.ActiveRow.Cells["Posicion"].Value == (int)dr.Cells["Posicion"].Value)
					iConteo++;				
			}

			if (iConteo > 0)
			{
				MessageBox.Show("Ya ha ingresado las seriales para este articulo.", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				return false;
			}

			return true;
		}

		static public bool vSerialesIngresadasCantidad(Infragistics.Win.UltraWinGrid.CellEventArgs e, Infragistics.Win.UltraWinGrid.UltraGrid uGrid)
		{			
			if (uGrid.Rows.Count == 0) return true;
			
			if (e.Cell.Row.Cells["idArticulo"].Value == System.DBNull.Value) return false;

			int iConteo = 0;

			foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in uGrid.Rows.All)
			{
				if ((int)e.Cell.Row.Cells["idArticulo"].Value == (int)dr.Cells["idArticulo"].Value 
					&& (int)e.Cell.Row.Cells["Posicion"].Value == (int)dr.Cells["Posicion"].Value)
					iConteo++;				
			}

			if (iConteo > (int)e.Cell.Row.Cells["Cantidad"].Value)
			{
				e.Cell.Value = iConteo;
				MessageBox.Show(string.Format("Ha ingresado {0} seriales para este articulo", iConteo), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				return false;
			}

			return true;
		}

		static public bool vSerialesIngresadasGrid(Infragistics.Win.UltraWinGrid.BeforeCellUpdateEventArgs e, Infragistics.Win.UltraWinGrid.UltraGrid uGrid, bool bNuevo)
		{			
			if (!bNuevo && e.Cell.Row.Cells["idArticulo"].Value != System.DBNull.Value) return false;
			
			if (uGrid.Rows.Count == 0) return true;
			
			if (e.Cell.Row.Cells["idArticulo"].Value == System.DBNull.Value) return true;

			int iConteo = 0;

			foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in uGrid.Rows.All)
			{
				if ((int)e.Cell.Row.Cells["idArticulo"].Value == (int)dr.Cells["idArticulo"].Value 
					&& (int)e.Cell.Row.Cells["Posicion"].Value == (int)dr.Cells["Posicion"].Value)
					iConteo++;				
			}

			if (iConteo > 0)
			{
				MessageBox.Show("Ya ha ingresado las seriales para este articulo.", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				return false;
			}

			return true;
		}

		static public bool ValidaSerialesIngresadasModificaProducto(Infragistics.Win.UltraWinGrid.BeforeCellUpdateEventArgs e, C1.Data.C1DataSet ds)
		{
			if (e.Cell.Column.ToString() == "Codigo")
			{
				int idDetCompra = 0;
				if (e.Cell.Row.Cells["idDetCompra"].Value != System.DBNull.Value) idDetCompra = (int)e.Cell.Row.Cells["idDetCompra"].Value;
				int idArticulo = 0;
				if (e.Cell.Row.Cells["idArticulo"].Value != System.DBNull.Value) idArticulo = (int)e.Cell.Row.Cells["idArticulo"].Value;
								
				string sSQL = string.Format("Exec ValidaConteoSeriales {0}, {1}", idDetCompra, idArticulo);
				int iConteo = Funcion.iEscalarSQL(ds, sSQL);
				if (iConteo > 0)
				{
					MessageBox.Show("Este Articulo ya tiene las Seriales Registradas no puede Modificarlo", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
					e.Cancel = true;
					return false;
				}
			}

			return true;
		}

		static public bool ValidaSerialesIngresadasModificaCantidad(Infragistics.Win.UltraWinGrid.CellEventArgs e, C1.Data.C1DataSet ds)
		{
			int idDetCompra = 0;
			if (e.Cell.Row.Cells["idDetCompra"].Value != System.DBNull.Value) idDetCompra = (int)e.Cell.Row.Cells["idDetCompra"].Value;
			int idArticulo = 0;
			if (e.Cell.Row.Cells["idArticulo"].Value != System.DBNull.Value) idArticulo = (int)e.Cell.Row.Cells["idArticulo"].Value;

			string sSQL = string.Format("Exec ValidaConteoSeriales {0}, {1}", idDetCompra, idArticulo);
			int iCantidad = Funcion.iEscalarSQL(ds, sSQL);		
 
			if ((int)e.Cell.Row.Cells["Cantidad"].Value < iCantidad)
			{
				e.Cell.Row.Cells["Cantidad"].Value = (int)e.Cell.OriginalValue;

				MessageBox.Show(string.Format("La Cantidad del Artículo '{0}' No Puede Ser Menor a {1} seriales registradas en este documento", 
					e.Cell.Row.Cells["Codigo"].Value.ToString().Trim(), iCantidad), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);						

				return false;
			}
			return true;	
		}

		static public bool ValidaConteoSeriales(Infragistics.Win.UltraWinGrid.UltraGrid uGrid, C1.Data.C1DataSet ds)
		{			
			int idDetCompra = 0;
			int idArticulo = 0;

			if (uGrid.ActiveRow.Cells["idDetCompra"].Value != System.DBNull.Value) idDetCompra = (int)uGrid.ActiveRow.Cells["idDetCompra"].Value;					
			if (uGrid.ActiveRow.Cells["idArticulo"].Value != System.DBNull.Value) idArticulo = (int)uGrid.ActiveRow.Cells["idArticulo"].Value;

			string sSQL = string.Format("Exec ValidaConteoSeriales {0}, {1}", idDetCompra, idArticulo);			
			if (Funcion.iEscalarSQL(ds, sSQL) > 0)
			{
				MessageBox.Show("Este Articulo ya tiene las Seriales Registradas no puede Modificarlo\n\nConsulte con el Administrador", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				return false;
			}

			return true;
		}

		static public bool ValidaBodegaEnInventario(int iBodega, DateTime dtFecha, C1.Data.C1DataSet ds)
		{			
			int iCont = Funcion.iEscalarSQL(ds, string.Format("Select COUNT(*) From CierreConteoInventario Where Bodega = {0} And CONVERT(DATE, Fecha) Between '20161001' And '{1}' And Estado = 1", 
				iBodega, dtFecha.ToString("yyyyMMdd")));

			if (iCont > 0)
			{				
				MessageBox.Show("El local esta en Inventario no puede crear Transferencia", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);												
				return false;
			}
			return true;
		}

		static public int ValidaSerialesIngresadasParaAnular(int idCompra, int idTipoFactura, C1.Data.C1DataSet ds)
		{
			string sSQL = string.Format("Exec ValidaSerialesIngresadasParaAnular {0}, {1}", idCompra, idTipoFactura);
			return Funcion.iEscalarSQL(ds, sSQL, true);
		}

		static public bool ValidaBodegaInventarios(Infragistics.Win.UltraWinGrid.UltraGrid uGrid, int iBodega, DateTime dtFecha, C1.Data.C1DataSet ds)
		{
			int iCont = 0;
			
			foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in uGrid.Rows.All)			
				if (Funcion.bEscalarSQL(ds, string.Format("Declare @Res Bit = 0 If Exists(Select idArticulo From Articulo Where idTipoGrupo In (1, 6) And idArticulo = {0}) Set @Res = 1 Select @Res ", (int)dr.Cells["idArticulo"].Value), true)) iCont++;				
			
			if (iCont > 0)
			{
				if (Funcion.iEscalarSQL(ds, string.Format("Select COUNT(*) From CierreConteoInventario Where Bodega = {0} And CONVERT(Date, Fecha) > '{1}'", 
					iBodega, dtFecha.ToString("yyyyMMdd"))) > 0)
				{
					MessageBox.Show("Existen conteos de inventario posteriores a la fecha del documento", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);																
					return false;
				}
			}

			return true;
		}

		static public decimal FacturacionValorDYD(Infragistics.Win.UltraWinGrid.UltraGrid uGrid, int idFormaPago, int iCuotas)
		{
			decimal dPorcentaje = 0.00m;
			decimal dBase = 0.00m;
		
			decimal dValorDD = 0.00m;

			if (idFormaPago == 3) dPorcentaje = 10m;
	
			if (idFormaPago == 9) 
			{
				if (iCuotas <= 18) dPorcentaje = 10m;	
				if (iCuotas > 18) dPorcentaje = 12m;	
			}
			
			if (uGrid.Rows.Count > 0)
				foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in uGrid.Rows.All)
					if (dr.Cells["idArticuloGE"].Value != System.DBNull.Value)
						if ((int)dr.Cells["idArticuloGE"].Value == 0)				
							if ((int)dr.Cells["Tipo"].Value != 8)									
								dBase = dBase + Convert.ToDecimal(dr.Cells["Subtotal"].Value) - Convert.ToDecimal(dr.Cells["DescuentoArt"].Value);
			
			dValorDD = (dBase * dPorcentaje)/100.00m;

			return Decimal.Round(dValorDD, 2);
		}

		static public void FacturacionCargaDYD (Infragistics.Win.UltraWinGrid.UltraGrid uGrid, decimal dValor, decimal dIVA, C1.Data.C1DataSet ds)
		{
			int iFila = uGrid.Rows.Count - 1;
			uGrid.DisplayLayout.Bands[0].AddNew();	
			iFila++;

			string sSQLInf = "Select a.idArticulo, a.Codigo, a.Articulo, Precio4 From Articulo a Inner Join ArticuloGrupo ag On ag.idGrupoArticulo = a.idGrupoArticulo Where Seguro = 1 And Discontinuo = 0";
			SqlDataReader drA = Funcion.rEscalarSQL(ds, sSQLInf, true);
			drA.Read();											
			string sCodigo = drA.GetString(1);
			int idArticulo = drA.GetInt32(0);
			string sArticulo = drA.GetString(2);
			if (drA.GetDecimal(3) > 0) dValor = drA.GetDecimal(3); 
			drA.Close();
			
			uGrid.Rows[iFila].Cells["idDetCompra"].Value = 0;
			uGrid.Rows[iFila].Cells["idOrigen"].Value = 0;
			uGrid.Rows[iFila].Cells["idUnidad"].Value = 1;			
			uGrid.Rows[iFila].Cells["idCombo"].Value = 0;
			uGrid.Rows[iFila].Cells["Tipo"].Value = 8;
			uGrid.Rows[iFila].Cells["Codigo"].Value = sCodigo;			
			uGrid.Rows[iFila].Cells["Articulo"].Value = sArticulo;
			uGrid.Rows[iFila].Cells["idArticulo"].Value = idArticulo;
			uGrid.Rows[iFila].Cells["Cantidad"].Value = 1;
			uGrid.Rows[iFila].Cells["Precio"].Value = dValor;			
			uGrid.Rows[iFila].Cells["Impuesto"].Value = dIVA;
			uGrid.Rows[iFila].Cells["DescuentoArt"].Value = 0.00m;
			uGrid.Rows[iFila].Cells["DescuentoPorc"].Value = 0.00m;			
			uGrid.Rows[iFila].Cells["Entregado"].Value = false;
			uGrid.Rows[iFila].Cells["idArticuloGE"].Value = 0;
			uGrid.Rows[iFila].Cells["Principal"].Value = false;			
			uGrid.Rows[iFila].Cells["Bloqueado"].Value = false;
			uGrid.Rows[iFila].Cells["EstadoSerial"].Value = false;
			uGrid.Rows[iFila].Cells["Posicion"].Value = uGrid.Rows.Count;	
			uGrid.Rows[iFila].Cells["Jornadas"].Value = 0.00m;
		}

		static public void FacturacionCargaArticulos(object sender, Infragistics.Win.UltraWinGrid.CellEventArgs e, Infragistics.Win.UltraWinGrid.UltraGrid uGrid, int idFormaPago, int idTipoPrecio, int iBodega, int iBodega1, 
			int idBanco, int idTarjeta, int idPlan, int idPlazo, DateTime dtFecha, int idTipoFactura, bool bFiscal, bool bCargaCombos, bool bCombosPromociones, bool bGEx, int iCuotas, int idEntidadFinanciera, 
			int idCliente, int iNumCompras, int idSerial, bool bCupon, bool bConsignacion, int idOrigen, C1.Data.C1DataSet ds)
		{				
			int idArticulo = 0;
			int iCuenta = 0;				
			int idCombo = 0;
			int iTipo = 0;
			int iFila = 0;

			bool bRemates = false;
			
			if (iBodega1 > 0)
				bRemates = Funcion.bEscalarSQL(ds, string.Format("Select Remates From Bodega Where Bodega = {0}", iBodega1), true);

			if (e.Cell.Row.Cells["Codigo"].Value.ToString().Length == 0) e.Cell.Row.Cells["Codigo"].Value = e.Cell.OriginalValue.ToString();						
			
			#region Retorna idArticulo
			string sSQL = string.Format("Exec ConsultaArticulos '{0}'", e.Cell.Row.Cells["Codigo"].Value.ToString());//, idTipoFactura, iBodega);
			iCuenta = Funcion.iEscalarSQL(ds, sSQL, true);
						
			if (iCuenta != 1)
			{ 
				using (BuscaArticulos miBusqueda = new BuscaArticulos(e.Cell.Row.Cells["Codigo"].Value.ToString(), iBodega1, idTipoFactura))
				{
					if (DialogResult.OK == miBusqueda.ShowDialog())
					{
						idArticulo = miBusqueda.idArticulo;
					}
					else
					{
						e.Cell.Row.Cells["idArticulo"].Value = System.DBNull.Value;
						e.Cell.Row.Cells["Precio"].Value = 0.00m;
						e.Cell.Row.Cells["Codigo"].Value = "";
						e.Cell.Row.Cells["Articulo"].Value = "";
						if (idTipoFactura != 8 && idTipoFactura != 9)
						{
							e.Cell.Row.Cells["idUnidad"].Value = 0;
							e.Cell.Row.Cells["Impuesto"].Value = 0.00m;
							e.Cell.Row.Cells["DescuentoArt"].Value = 0.00m;
							e.Cell.Row.Cells["DescuentoPorc"].Value = 0.00m;
						}

						return;						
					}
				}
			}			
			else 
			{
				sSQL = string.Format("If Exists(Select idArticulo From Articulo a Where Codigo = '{0}' And ISNULL((Select GE from ArticuloGrupo Where idGrupoArticulo = a.idGrupoArticulo), 0) = 0) Select idArticulo From Articulo a Where Codigo = '{0}' And ISNULL((Select GE from ArticuloGrupo Where idGrupoArticulo = a.idGrupoArticulo), 0) = 0 Else Select 0", e.Cell.Row.Cells["Codigo"].Value.ToString());//, iBodega, idTipoFactura);
				idArticulo = Funcion.iEscalarSQL(ds, sSQL, true);

				if (idArticulo == 0)
				{
					MessageBox.Show("Articulo no existe", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
					return;
				}
				
				if (idTipoFactura == 36 && bConsignacion)
				{
					if (!Funcion.bEscalarSQL(ds, string.Format("Select PropioConsignado From Articulo Where idArticulo = {0}", idArticulo), true))
					{
						MessageBox.Show("Articulo no esta asignado para Consignación", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
						return;
					}
				}
			}
			#endregion Retorna idArticulo

			bool bDYD = Funcion.bEscalarSQL(ds, string.Format("Select ag.Seguro From Articulo a Inner Join ArticuloGrupo ag On ag.idGrupoArticulo = a.idGrupoArticulo Where idArticulo = {0}", idArticulo), true);

			#region Valida Cupon Promocional
			bool bCuponPromocional = false;
			
			if (idTipoFactura == 1)
			{
				if (Funcion.bEscalarSQL(ds, string.Format("Select CuponPromocional From Articulo Where idArticulo = {0}", idArticulo), true))
				{
					using (frmCuponWeb CuponWeb = new frmCuponWeb(idArticulo, true))
					{				
						if (DialogResult.OK == CuponWeb.ShowDialog())
						{	
							bCuponPromocional = true;
						}
						else bCuponPromocional = false;
					}
				}
			}
			#endregion Valida Cupon Promocional

			#region Valida y carga combos promocionales
			bool bComboPromocion = Funcion.bEscalarSQL(ds, string.Format("Declare @Res Bit = 0 If Exists(Select idPromocion From CompraPromocion Where idArticulo = {0}) Set @Res = 1 Select @Res", idArticulo), true);

			if (!bRemates)
			{				
				if (bCombosPromociones && bComboPromocion)
				{
					int idTipoPreciovalida = idTipoPrecio;
					if (idFormaPago == 1) idTipoPreciovalida = 2;
					if (idFormaPago == 13) idTipoPreciovalida = 3;
					if (idFormaPago == 19) idTipoPreciovalida = 3;
					if (idFormaPago == 11) idTipoPreciovalida = 90;

					idCombo = Funcion.iEscalarSQL(ds, string.Format("Select dbo.ValidaPromocionEspecifica ({0}, {1}, {2}, {3}, {4}, {5}, {6}, {7}, {8})", 
						idArticulo, idTipoPreciovalida, iBodega, idBanco, idTarjeta, idPlan, idPlazo, iCuotas, Convert.ToInt32(bCupon)));

					if (idCombo > 0)
					{					
						iTipo = 2;
						iFila = e.Cell.Row.Index;

						CargaArticulosCombo(uGrid, string.Format("Exec CargaArticulosPromocion {0}, {1}, {2}, {3}, {4}, '{5}', {6}", idCombo, idTipoPrecio, idFormaPago, iBodega1, idTipoFactura, dtFecha.ToString("yyyyMMdd"), idOrigen), 
							idCombo, idTipoFactura, 2, iTipo, iFila, ds);
					}
					else
					{
						MessageBox.Show("Promoción no cumple con los parametros necesarios", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
						iTipo = -1;
					}				
				}
			}
			#endregion Valida y carga combos promocionales

			#region Valida y Carga Combos
			if (bCargaCombos && !bComboPromocion)
			{
				int idMarca = Funcion.iEscalarSQL(ds, string.Format("Select ISNULL(idMarca, 0) From Articulo Where idArticulo = {0}", idArticulo));

				SqlDataReader drValidaCombo = Funcion.rEscalarSQL(ds, string.Format("Exec ComboPorSubGrupo {0}, 0, {1}", idArticulo, idMarca), true);
				drValidaCombo.Read();
				if (drValidaCombo.HasRows)
				{
					idCombo = drValidaCombo.GetInt32(0);
					iTipo = drValidaCombo.GetInt32(1);
				}
				drValidaCombo.Close();
				
				iFila = e.Cell.Row.Index;
			}

			if (!bComboPromocion && idCombo > 0)
			{
				if (DialogResult.No == MessageBox.Show(string.Format("Este Artículo esta asociado a un combo\n\n¿Desea cargar los artículos del combo?"), "Point Technology", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2))
				{
					idCombo = 0;
					iTipo = 0;
				}
			}

			if (iTipo == 0)				
			{
				int iTipoArt = 0;
				if (idCombo > 0) iTipoArt = 1;
				bool bObsequio = false;

				if (iTipoArt == 0 && 
					Funcion.bEscalarSQL(ds, string.Format("Select Obsequio From Articulo Where idArticulo = {0}", idArticulo), true) && 
					Funcion.bEscalarSQL(ds, string.Format("Select Credito From Bodega Where Bodega = {0}", iBodega), true) && (idTipoFactura == 1 || idTipoFactura == 43)) bObsequio = true;
				
				if (!bObsequio)// && !bCuponPromocional)
					CargaArticuloIndividual(e, idArticulo, idFormaPago, idTipoPrecio, iBodega, iBodega1, idBanco, idTarjeta, idPlan, dtFecha, idTipoFactura, bFiscal, idCombo, iTipoArt, idCliente, iNumCompras, idSerial, bCupon, bRemates, bCuponPromocional, bConsignacion, ds);
				else 
				{
					e.Cell.Row.Cells["Codigo"].Value = "";
					MessageBox.Show("Producto de obsequio no puede facturar sin su respectivo combo o promocion", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);					
				}
			}
						
			if (idCombo > 0 && !bComboPromocion)
			{				
				CargaArticulosCombo(uGrid, string.Format("Exec ArticulosCombos {0}, {1}, {2}, {3}, {4}, {5}, '{6}', {7}, 0, {8}", idFormaPago, idCombo, idTipoPrecio, iBodega, iBodega1, idTipoFactura, dtFecha.ToString("yyyyMMdd"), idCliente, bCuponPromocional), 
					idCombo, idTipoFactura, 1, iTipo, iFila, ds);
			}
			#endregion Valida y Carga Combos			

			if (!bRemates)
			{
				#region Promociones Por Valor - Adicional - Opcional
				if (!bDYD)
				{
					if (bCombosPromociones)
					{
						CargaArticulosPromocionOpcional(uGrid, dtFecha, idTipoFactura, iBodega, iBodega1, idArticulo, idFormaPago, idTipoPrecio, idBanco, idTarjeta, idPlan, idPlazo, uGrid.Rows.Count - 1, iCuotas, idCliente, bCupon, idOrigen, ds);

						CargaArticulosPromocionPorValor(uGrid, dtFecha, iBodega, iBodega1, idArticulo, idFormaPago, idTipoPrecio, idBanco, idTarjeta, idPlan, idPlazo, uGrid.Rows.Count - 1, iCuotas, bCupon, idOrigen, ds);

						CargaArticulosPromocionAdicional(uGrid, dtFecha, idTipoFactura, iBodega, iBodega1, idArticulo, idFormaPago, idTipoPrecio, idBanco, idTarjeta, idPlan, idPlazo, uGrid.Rows.Count - 1, iCuotas, bCupon, idOrigen, ds);
					}
				}
				#endregion Promociones Por Valor y Adicional - Opcional

				#region Garantia Extendida
				if (bGEx)
				{
					if (idCombo == 0) 
						if (!Funcion.bEscalarSQL(ds, string.Format("Select GarantiaExtendida From Articulo Where idArticulo = {0}", idArticulo), true)) return;

					if (Funcion.bEscalarSQL(ds, "Select GarantiaExtendida From SeteoF", true))
						GarantiaExtendida(uGrid, dtFecha, iCuotas, idEntidadFinanciera, idCombo, uGrid.Rows.Count - 1, ds);	
		
					if (Funcion.bEscalarSQL(ds, "Select GarantiaExtendidaPlus From SeteoF", true))
						GarantiaExtendidaPlus(uGrid, dtFecha, iCuotas, idTipoPrecio, idEntidadFinanciera, idCombo, uGrid.Rows.Count - 1, ds);	
				}
				#endregion Garantia Extendida			
			}
		}

		static private void CargaArticulosPromocionOpcional(Infragistics.Win.UltraWinGrid.UltraGrid uGrid, DateTime dtFecha, int idTipoFactura, int iBodega, int iBodega1, int idArticulo, int idFormaPago, int idTipoPrecio, int idBanco, int idTarjeta, int idPlan, int idPlazo, int iFila, int iCuotas, int idCliente, bool bCupon, int idOrigen, C1.Data.C1DataSet ds)
		{
			#region Promocion Opcional
			int idTipoPreciovalida = idTipoPrecio;
			if (idFormaPago == 1) idTipoPreciovalida = 2;
			if (idFormaPago == 11) idTipoPreciovalida = 90;

			int idPromocionOpcional = Funcion.iEscalarSQL(ds, string.Format("Select dbo.ValidaPromocionOpcional ({0}, {1}, {2}, {3}, {4}, {5}, {6}, {7}, {8})", 
				iBodega, idArticulo, idTipoPreciovalida, idBanco, idTarjeta, idPlan, idPlazo, iCuotas, Convert.ToInt32(bCupon)));

			int idCombo = 0;
			int iTipo = 0;
			int idMarca = Funcion.iEscalarSQL(ds, string.Format("Select ISNULL(idMarca, 0) From Articulo Where idArticulo = {0}", idArticulo));

			SqlDataReader drValidaCombo = Funcion.rEscalarSQL(ds, string.Format("Exec ComboPorSubGrupo {0}, 1, {1}", idArticulo, idMarca), true);
			drValidaCombo.Read();
			if (drValidaCombo.HasRows)
			{
				idCombo = drValidaCombo.GetInt32(0);
				iTipo = drValidaCombo.GetInt32(1);
			}
			drValidaCombo.Close();					

			if (idPromocionOpcional > 0 || idCombo > 0)
			{					
				using (frmFacturacionListaPromocionOpcional FLPO = new frmFacturacionListaPromocionOpcional(idPromocionOpcional, idTipoPrecio, idFormaPago, iBodega, iBodega1, idTipoFactura, dtFecha, idCombo, idCliente, bCupon, idOrigen))
				{				
					if (DialogResult.OK == FLPO.ShowDialog())
					{
						#region Carga Articulos Promociones
						if (FLPO.ultraGrid1.Rows.Count > 0)
						{								
//							if (iFila == 0) 
//							{
//								uGrid.DisplayLayout.Bands[0].AddNew();	
//								iFila++;
//							}

							if (idCombo > 0) uGrid.ActiveRow.Cells["idCombo"].Value = idCombo;
							if (idPromocionOpcional > 0) uGrid.ActiveRow.Cells["idCombo"].Value = idPromocionOpcional;

							foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in FLPO.ultraGrid1.Rows.All)
							{
								if ((bool)dr.Cells["Sel"].Value)
								{			
									uGrid.DisplayLayout.Bands[0].AddNew();	
									iFila++;
						
									uGrid.Rows[iFila].Cells["Codigo"].Value = dr.Cells["Codigo"].Value;
									uGrid.Rows[iFila].Cells["idArticulo"].Value = (int)dr.Cells["idArticulo"].Value;
									uGrid.Rows[iFila].Cells["Articulo"].Value = dr.Cells["Articulo"].Value;
									uGrid.Rows[iFila].Cells["idUnidad"].Value = 1;
									uGrid.Rows[iFila].Cells["Cantidad"].Value = 1;
									uGrid.Rows[iFila].Cells["Precio"].Value = Convert.ToDecimal(dr.Cells["Precio"].Value);
									uGrid.Rows[iFila].Cells["Impuesto"].Value = Convert.ToDecimal(dr.Cells["Impuesto"].Value);
									uGrid.Rows[iFila].Cells["DescuentoArt"].Value = Convert.ToDecimal(dr.Cells["DescuentoArt"].Value);
									uGrid.Rows[iFila].Cells["DescuentoPorc"].Value = Convert.ToDecimal(dr.Cells["DescuentoPorc"].Value);
									uGrid.Rows[iFila].Cells["Tipo"].Value = 6;
									uGrid.Rows[iFila].Cells["Entregado"].Value = false;
									uGrid.Rows[iFila].Cells["Principal"].Value = (bool)dr.Cells["Principal"].Value;
									uGrid.Rows[iFila].Cells["idCombo"].Value = idPromocionOpcional;									
									uGrid.Rows[iFila].Cells["idArticuloGE"].Value = 0;
									
//									iFila = iFila + 1;
//									uGrid.Rows.Band.AddNew();								
								}
							}
						}
						#endregion Carga Articulos Promociones

						#region Carga Articulos Opcionales de Combos
						if (FLPO.ultraGrid2.Rows.Count > 0)
						{								
//							if (iFila == 0) 
//							{
//								uGrid.DisplayLayout.Bands[0].AddNew();	
//								iFila++;
//							}

							foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in FLPO.ultraGrid2.Rows.All)
							{
								if ((bool)dr.Cells["Sel"].Value)
								{					
									uGrid.DisplayLayout.Bands[0].AddNew();	
									iFila++;
				
									uGrid.Rows[iFila].Cells["Codigo"].Value = dr.Cells["Codigo"].Value;
									uGrid.Rows[iFila].Cells["idArticulo"].Value = (int)dr.Cells["idArticulo"].Value;																				
									uGrid.Rows[iFila].Cells["Articulo"].Value = dr.Cells["Articulo"].Value;
									uGrid.Rows[iFila].Cells["idUnidad"].Value = 1;
									uGrid.Rows[iFila].Cells["Cantidad"].Value = 1;
									uGrid.Rows[iFila].Cells["Precio"].Value = Convert.ToDecimal(dr.Cells["Precio"].Value);
									uGrid.Rows[iFila].Cells["Impuesto"].Value = Convert.ToDecimal(dr.Cells["Impuesto"].Value);
									uGrid.Rows[iFila].Cells["DescuentoArt"].Value = Convert.ToDecimal(dr.Cells["DescuentoArt"].Value);
									uGrid.Rows[iFila].Cells["DescuentoPorc"].Value = Convert.ToDecimal(dr.Cells["DescuentoPorc"].Value);
									uGrid.Rows[iFila].Cells["Tipo"].Value = 7;
									uGrid.Rows[iFila].Cells["Entregado"].Value = false;
									uGrid.Rows[iFila].Cells["Principal"].Value = (bool)dr.Cells["Principal"].Value;
									uGrid.Rows[iFila].Cells["idCombo"].Value = idCombo;									
									uGrid.Rows[iFila].Cells["idArticuloGE"].Value = 0;
									
//									iFila = iFila + 1;
//									uGrid.Rows.Band.AddNew();								
								}
							}
						}
						#endregion Carga Articulos Opcionales de Combos
					}
				}
			}
			#endregion Promocion Opcional
		}

		static private void CargaArticulosPromocionPorValor(Infragistics.Win.UltraWinGrid.UltraGrid uGrid, DateTime dtFecha, int iBodega, int iBodega1, int idArticulo, int idFormaPago, int idTipoPrecio, int idBanco, int idTarjeta, int idPlan, int idPlazo, int iFila, int iCuotas, bool bCupon, int idOrigen, C1.Data.C1DataSet ds)
		{
			#region Promocion Por Valor
//			if (iFila == 0) 
//			{
//				uGrid.DisplayLayout.Bands[0].AddNew();	
//				iFila++;
//			}

			#region Total factura
			decimal dTotalFactura = 0.00m;
			decimal dSubtotalfactura = 0.00m;
			decimal dIVA = Funcion.decEscalarSQL(ds, string.Format("Select dbo.RetornaImpuestoIVAVigente('{0}', 1)", dtFecha.ToString("yyyyMMdd")));
			
			foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in uGrid.Rows.All)	
				dSubtotalfactura = dSubtotalfactura + (((int)dr.Cells["Cantidad"].Value * Convert.ToDecimal(dr.Cells["Precio"].Value)) - Convert.ToDecimal(dr.Cells["DescuentoArt"].Value)); 
						
			dTotalFactura = dSubtotalfactura + (dSubtotalfactura * dIVA/100.00m);
			#endregion Total factura

			int idTipoPreciovalida = idTipoPrecio;
			if (idFormaPago == 1) idTipoPreciovalida = 2;
			if (idFormaPago == 11) idTipoPreciovalida = 90;

			int idPromocionValor = Funcion.iEscalarSQL(ds, string.Format("Select dbo.ValidaPromocionPorValor ({0}, {1}, {2}, {3}, {4}, {5}, {6}, {7}, {8}, {9})", 
				iBodega, idArticulo, idTipoPreciovalida, dTotalFactura, idBanco, idTarjeta, idPlan, idPlazo, iCuotas, Convert.ToInt32(bCupon)));

			if (idPromocionValor > 0)
			{					
				#region Valida Tipo de Promocion
				string sTipoPromocion = "";				
				if (uGrid.Rows.Count > 0)
				{									
					foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in uGrid.Rows.All)
					{
						if ((bool)dr.Cells["Principal"].Value && (int)dr.Cells["Tipo"].Value == 2 && (int)dr.Cells["idCombo"].Value > 0)
						{
							sTipoPromocion = Funcion.sEscalarSQL(ds, string.Format("Select Tipo From CompraPromocion Where idPromocion = {0}", (int)dr.Cells["idCombo"].Value));
							continue;
						}
					}					
				}
				#endregion Valida Tipo de Promocion
				
				string sSQLProm = string.Format("Exec CargaArticulosPromocionPorValor {0}, {1}, {2}, {3}, {4}, '{5}', {6}", 
					idPromocionValor, idTipoPrecio, idFormaPago, iBodega1, dTotalFactura, dtFecha.ToString("yyyyMMdd"), idOrigen);
				SqlDataReader drC = Funcion.rEscalarSQL(ds, sSQLProm, true);

				while (drC.Read())
				{
					#region Variables
					int idArtProm = drC.GetInt32(0);
					string sCodProm = drC.GetString(1); 
					string sArtProm = drC.GetString(2);
					int iUniProm = drC.GetInt32(3);
					decimal dPrecProm = drC.GetDecimal(4);
					decimal dImpProm = drC.GetDecimal(5);
					decimal dDesArtProm = drC.GetDecimal(6);
					decimal dDesPorProm = drC.GetDecimal(7);									
					bool bPrincipal = drC.GetBoolean(8);
					bool bAplicaBaratazos = drC.GetBoolean(10);
					bool bAplicaCombos = drC.GetBoolean(11);
					bool bAplicaPromociones = drC.GetBoolean(12);
					bool bValArtCombo = false;
					#endregion Variables

					#region Valida Tipo de Promocion y repetidos
					if (sTipoPromocion.Length > 0)
					{
						if (sTipoPromocion == "BARATAZO" && bAplicaBaratazos) bValArtCombo = true;
						if (sTipoPromocion == "COMBO" && bAplicaCombos) bValArtCombo = true;
						if (sTipoPromocion == "PROMOCIÓN" && bAplicaPromociones) bValArtCombo = true;
					}
					else bValArtCombo = true;

					foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in uGrid.Rows.All)
					{
						int idArtPromCarg = 0;
						if (dr.Cells["idArticulo"].Value != System.DBNull.Value) idArtPromCarg = (int)dr.Cells["idArticulo"].Value;
						if (idArtProm == idArtPromCarg && (int)dr.Cells["Tipo"].Value == 3) bValArtCombo = false;
					}
					#endregion Valida Tipo de Promocion y repetidos

					#region Asigna valores en el grid
					if (bValArtCombo)
					{	
						uGrid.DisplayLayout.Bands[0].AddNew();	
						iFila++;

						uGrid.Rows[iFila].Cells["Codigo"].Value = sCodProm;
						uGrid.Rows[iFila].Cells["idArticulo"].Value = idArtProm;						
						uGrid.Rows[iFila].Cells["Articulo"].Value = sArtProm;
						uGrid.Rows[iFila].Cells["idUnidad"].Value = iUniProm;
						uGrid.Rows[iFila].Cells["Cantidad"].Value = 1;
						uGrid.Rows[iFila].Cells["Precio"].Value = dPrecProm;
						uGrid.Rows[iFila].Cells["Impuesto"].Value = dImpProm;
						uGrid.Rows[iFila].Cells["DescuentoArt"].Value = dDesArtProm;
						uGrid.Rows[iFila].Cells["DescuentoPorc"].Value = dDesPorProm;
						uGrid.Rows[iFila].Cells["Principal"].Value = bPrincipal;
						uGrid.Rows[iFila].Cells["idCombo"].Value = idPromocionValor;
						uGrid.Rows[iFila].Cells["Tipo"].Value = 3;

//						uGrid.DisplayLayout.Bands[0].AddNew();
//						iFila = iFila + 1;
					}
					#endregion Asigna valores en el grid
				} 
				drC.Close();														
			}
			#endregion Promocion Por Valor
		}

		static private void CargaArticulosPromocionAdicional(Infragistics.Win.UltraWinGrid.UltraGrid uGrid, DateTime dtFecha, int idTipoFactura, int iBodega, int iBodega1, int idArticulo, int idFormaPago, int idTipoPrecio, int idBanco, int idTarjeta, int idPlan, int idPlazo, int iFila, int iCuotas, bool bCupon, int idOrigen, C1.Data.C1DataSet ds)
		{
			#region Promocion Adicional
//			if (iFila == 0) 
//			{
//				uGrid.DisplayLayout.Bands[0].AddNew();	
//				iFila++;
//			}

			int idTipoPreciovalida = idTipoPrecio;
			if (idFormaPago == 1) idTipoPreciovalida = 2;
			if (idFormaPago == 11) idTipoPreciovalida = 90;

			int idPromocionAdicional = Funcion.iEscalarSQL(ds, string.Format("Select dbo.ValidaPromocionAdicional ({0}, {1}, {2}, {3}, {4}, {5}, {6}, {7}, {8})", 
				iBodega, idArticulo, idTipoPreciovalida, idBanco, idTarjeta, idPlan, idPlazo, iCuotas, Convert.ToInt32(bCupon)));

			if (idPromocionAdicional > 0)
			{	
				#region Productos Promociones
				string sSQLProm = string.Format("Exec CargaArticulosPromocion {0}, {1}, {2}, {3}, {4}, '{5}', {6}", 
					idPromocionAdicional, idTipoPrecio, idFormaPago, iBodega1, idTipoFactura, dtFecha.ToString("yyyyMMdd"), idOrigen);
				SqlDataReader drC = Funcion.rEscalarSQL(ds, sSQLProm, true);
																	
				while (drC.Read())
				{
					#region Variables
					int idArtProm = drC.GetInt32(0);
					string sCodProm = drC.GetString(1); 
					string sArtProm = drC.GetString(2);
					int iUniProm = drC.GetInt32(3);
					int iCantidad = drC.GetInt32(4);
					decimal dPrecProm = drC.GetDecimal(5);
					decimal dImpProm = drC.GetDecimal(6);
					decimal dDesArtProm = drC.GetDecimal(7);
					decimal dDesPorProm = drC.GetDecimal(8);									
					bool bPrincipal = drC.GetBoolean(9);
					bool bValArtCombo = true;
					#endregion Variables

					#region Valida Tipo de Promocion y repetidos
					foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in uGrid.Rows.All)
					{
						int idArtPromCarg = 0;
						if (dr.Cells["idArticulo"].Value != System.DBNull.Value) idArtPromCarg = (int)dr.Cells["idArticulo"].Value;
						if (idArtProm == idArtPromCarg && (int)dr.Cells["Tipo"].Value == 4) bValArtCombo = false;
					}
					#endregion Valida Tipo de Promocion y repetidos

					#region Asigna valores en el grid
					if (bValArtCombo)
					{						
						uGrid.DisplayLayout.Bands[0].AddNew();	
						iFila++;

						uGrid.Rows[iFila].Cells["Codigo"].Value = sCodProm;
						uGrid.Rows[iFila].Cells["idArticulo"].Value = idArtProm;						
						uGrid.Rows[iFila].Cells["Articulo"].Value = sArtProm;
						uGrid.Rows[iFila].Cells["idUnidad"].Value = iUniProm;
						uGrid.Rows[iFila].Cells["Cantidad"].Value = iCantidad;
						uGrid.Rows[iFila].Cells["Precio"].Value = dPrecProm;
						uGrid.Rows[iFila].Cells["Impuesto"].Value = dImpProm;
						uGrid.Rows[iFila].Cells["DescuentoArt"].Value = dDesArtProm;
						uGrid.Rows[iFila].Cells["DescuentoPorc"].Value = dDesPorProm;
						uGrid.Rows[iFila].Cells["Principal"].Value = bPrincipal;
						uGrid.Rows[iFila].Cells["idCombo"].Value = idPromocionAdicional;
						uGrid.Rows[iFila].Cells["Tipo"].Value = 4;

//						iFila = iFila + 1;
//
//						uGrid.DisplayLayout.Bands[0].AddNew();
					}
					#endregion Asigna valores en el grid
				} 
				drC.Close();						
				#endregion Productos Promociones				
			}					
			#endregion Promocion Adicional
		}

		static public void CargaArticulosPromocionCupon(Infragistics.Win.UltraWinGrid.UltraGrid uGrid, DateTime dtFecha, int idTipoFactura, int iBodega, int iBodega1, int idArticulo, int idFormaPago, int idTipoPrecio, int idBanco, int idTarjeta, int idPlan, int idPlazo, int iFila, int iCuotas, int idOrigen, C1.Data.C1DataSet ds)
		{
			#region Promocion Cupones
			int idTipoPreciovalida = idTipoPrecio;
			if (idFormaPago == 1) idTipoPreciovalida = 2;
			if (idFormaPago == 11) idTipoPreciovalida = 90;

			int idPromocionCupon = Funcion.iEscalarSQL(ds, string.Format("Select dbo.ValidaPromocionCupon ({0}, {1}, {2}, {3}, {4}, {5}, {6}, {7})", 
				iBodega, idArticulo, idTipoPreciovalida, idBanco, idTarjeta, idPlan, idPlazo, iCuotas));

			if (idPromocionCupon > 0)
			{	
				#region Productos Promociones
				string sSQLProm = string.Format("Exec CargaArticulosPromocion {0}, {1}, {2}, {3}, {4}, '{5}', {6}", 
					idPromocionCupon, idTipoPrecio, idFormaPago, iBodega1, idTipoFactura, dtFecha.ToString("yyyyMMdd"), idOrigen);
				SqlDataReader drC = Funcion.rEscalarSQL(ds, sSQLProm, true);
																	
				while (drC.Read())
				{
					#region Variables
					int idArtProm = drC.GetInt32(0);
					string sCodProm = drC.GetString(1); 
					string sArtProm = drC.GetString(2);
					int iUniProm = drC.GetInt32(3);
					int iCantidad = drC.GetInt32(4);
					decimal dPrecProm = drC.GetDecimal(5);
					decimal dImpProm = drC.GetDecimal(6);
					decimal dDesArtProm = drC.GetDecimal(7);
					decimal dDesPorProm = drC.GetDecimal(8);									
					bool bPrincipal = drC.GetBoolean(9);
					bool bValArtCombo = true;
					#endregion Variables

					#region Valida Tipo de Promocion y repetidos
					foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in uGrid.Rows.All)
					{
						int idArtPromCarg = 0;
						if (dr.Cells["idArticulo"].Value != System.DBNull.Value) idArtPromCarg = (int)dr.Cells["idArticulo"].Value;
						if (idArtProm == idArtPromCarg && (int)dr.Cells["Tipo"].Value == 4) bValArtCombo = false;
					}
					#endregion Valida Tipo de Promocion y repetidos

					#region Asigna valores en el grid
					if (bValArtCombo)
					{						
						uGrid.DisplayLayout.Bands[0].AddNew();	
						iFila++;

						uGrid.Rows[iFila].Cells["Codigo"].Value = sCodProm;
						uGrid.Rows[iFila].Cells["idArticulo"].Value = idArtProm;						
						uGrid.Rows[iFila].Cells["Articulo"].Value = sArtProm;
						uGrid.Rows[iFila].Cells["idUnidad"].Value = iUniProm;
						uGrid.Rows[iFila].Cells["Cantidad"].Value = 1;
						uGrid.Rows[iFila].Cells["Precio"].Value = dPrecProm;
						uGrid.Rows[iFila].Cells["Impuesto"].Value = dImpProm;
						uGrid.Rows[iFila].Cells["DescuentoArt"].Value = dDesArtProm;
						uGrid.Rows[iFila].Cells["DescuentoPorc"].Value = dDesPorProm;
						uGrid.Rows[iFila].Cells["Principal"].Value = bPrincipal;
						uGrid.Rows[iFila].Cells["idCombo"].Value = idPromocionCupon;
						uGrid.Rows[iFila].Cells["Tipo"].Value = 8;
					}
					#endregion Asigna valores en el grid
				} 
				drC.Close();						
				#endregion Productos Promociones				
			}					
			#endregion Promocion Cupones
		}

		static private void CargaArticuloIndividual (Infragistics.Win.UltraWinGrid.CellEventArgs e, int idArticulo, int idFormaPago, int idTipoPrecio, int iBodega, int iBodega1,
			int idBanco, int idTarjeta, int idPlan, DateTime dtFecha, int idTipoFactura, bool bFiscal, int idCombo, int iTipoArticulo, int idCliente, int iNumCompras, 
			int idSerial, bool bCupon, bool bRemates, bool bCuponPromocional, bool bConsignado, C1.Data.C1DataSet ds)
		{
			#region Carga Producto
			string sCodigo = "";
			decimal dPrecio = 0.00m;
			int idUnidad = 0;
			decimal dImpuesto = 0.00m;
			decimal dDescuentoValor = 0.00m;
			decimal dDescuentoPorc = 0.00m;
			
			
			string sSQL = string.Format("Exec CargaInformacionArticulo {0}, {1}, {2}, {3}, {4}, {5}, {6}, '{7}', {8}, {9}, {10}, {11}, {12}, {13}, {14}, {15}, {16}", 
				idArticulo, idFormaPago, idTipoPrecio, iBodega, idBanco, idTarjeta, idPlan, dtFecha.ToString("yyyyMMdd"), idTipoFactura, bFiscal, idCliente, iNumCompras, idSerial, Convert.ToInt32(bCupon), Convert.ToInt32(bRemates), bCuponPromocional, bConsignado);
			SqlDataReader dr = Funcion.rEscalarSQL(ds, sSQL, true);
			dr.Read();					
			if (dr.HasRows)
			{								
				sCodigo = dr.GetString(1);
				e.Cell.Row.Cells["Articulo"].Value = dr.GetString(2);

				dPrecio = dr.GetDecimal(4);
				idUnidad = dr.GetInt32(3);
				dImpuesto = dr.GetDecimal(5);
				dDescuentoValor = dr.GetDecimal(6);
				dDescuentoPorc = dr.GetDecimal(7);
			}
			dr.Close();	

			e.Cell.Row.Cells["Codigo"].Value = sCodigo;
			e.Cell.Row.Cells["idArticulo"].Value = idArticulo;
			if (idTipoFactura != 42 && idTipoFactura != 54) e.Cell.Row.Cells["Precio"].Value = dPrecio;

			if (idTipoFactura != 8 && idTipoFactura != 9 && idTipoFactura != 36 && idTipoFactura != 42 && idTipoFactura != 54)
			{
				e.Cell.Row.Cells["idUnidad"].Value = idUnidad;
				e.Cell.Row.Cells["Impuesto"].Value = dImpuesto;				
				e.Cell.Row.Cells["Tipo"].Value = iTipoArticulo;
				e.Cell.Row.Cells["DescuentoArt"].Value = dDescuentoValor;
				e.Cell.Row.Cells["DescuentoPorc"].Value = dDescuentoPorc;	
			}

			if (idTipoFactura == 1 || idTipoFactura == 43) 
			{
				if (idCombo > 0) e.Cell.Row.Cells["Principal"].Value = true; else e.Cell.Row.Cells["Principal"].Value = false;
				e.Cell.Row.Cells["Bloqueado"].Value = false;
				e.Cell.Row.Cells["idCombo"].Value = idCombo;
				e.Cell.Row.Cells["Entregado"].Value = false;
			}
			#endregion Carga Producto		
		}

		static private void GarantiaExtendida(Infragistics.Win.UltraWinGrid.UltraGrid uGrid, DateTime dtFecha, int iCuotas, int idEntidadFinanciera, int idCombo, int iFila, C1.Data.C1DataSet ds)
		{
			string sIdArticulos = "";							
			if (Funcion.bEscalarSQL(ds, "Select GarantiaExtendida From SeteoF", true))
			{
				#region Valida si existen articulos para garantia extendida
				foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in uGrid.Rows.All)
				{
					if (dr.Cells["idArticulo"].Value != System.DBNull.Value)
					{
						if (sIdArticulos.Length == 0) sIdArticulos = dr.Cells["idArticulo"].Value.ToString();
						else if (sIdArticulos.Length > 0) sIdArticulos = sIdArticulos + ", " + dr.Cells["idArticulo"].Value.ToString();				
					}
				}
				#endregion Valida si existen articulos para garantia extendida

				#region Abre pantalla de Garantia Extendida
				if (sIdArticulos.Length > 0)
				{
					if (Funcion.iEscalarSQL(ds, string.Format("Exec ValidaArticulosGarantiaExtendida '{0}'", sIdArticulos)) > 0)
					{
						using (frmProductosGarantiaExtendida GE = new frmProductosGarantiaExtendida(uGrid, iCuotas, dtFecha, idEntidadFinanciera))
						{				
							if (DialogResult.OK == GE.ShowDialog())
							{
								#region Carga Articulo de Garantia Extendida
								if (GE.ultraGrid1.Rows.Count > 0)
								{
									decimal dIVA = Funcion.decEscalarSQL(ds, string.Format("Select dbo.RetornaImpuestoIVAVigente('{0}', 1)", dtFecha.ToString("yyyyMMdd")));

									//int iFila = 0;
									

									foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in GE.ultraGrid1.Rows.All)
									{										
										if ((bool)dr.Cells["Sel"].Value)
										{
//											if (idCombo == 0) 
//											{
											
											uGrid.DisplayLayout.Bands[0].AddNew();	
											iFila++;
												//iFila = uGrid.Rows.Count - 1; 
												//uGrid.DisplayLayout.Bands[0].AddNew();	
//											}
//											else iFila = uGrid.Rows.Count - 1;

											string sSQLInf = "Select a.idArticulo, a.Codigo, a.Articulo From Articulo a Inner Join ArticuloGrupo ag On ag.idGrupoArticulo = a.idGrupoArticulo Where GE = 1";									
											SqlDataReader drA = Funcion.rEscalarSQL(ds, sSQLInf, true);
											drA.Read();
											
											uGrid.Rows[iFila].Cells["Codigo"].Value = drA.GetString(1);
											uGrid.Rows[iFila].Cells["idArticulo"].Value = drA.GetInt32(0);
											uGrid.Rows[iFila].Cells["Articulo"].Value = drA.GetString(2);
											uGrid.Rows[iFila].Cells["idUnidad"].Value = 1;
											uGrid.Rows[iFila].Cells["Cantidad"].Value = 1;
											uGrid.Rows[iFila].Cells["Precio"].Value = dr.Cells["Precio"].Value;
											uGrid.Rows[iFila].Cells["Impuesto"].Value = dIVA;
											uGrid.Rows[iFila].Cells["DescuentoArt"].Value = 0.00m;
											uGrid.Rows[iFila].Cells["DescuentoPorc"].Value = 0.00m;
											uGrid.Rows[iFila].Cells["Tipo"].Value = 5;
											uGrid.Rows[iFila].Cells["Entregado"].Value = false;
											uGrid.Rows[iFila].Cells["idArticuloGE"].Value = dr.Cells["idArticulo"].Value;
											drA.Close();						

											//iFila = iFila + 1;
											//uGrid.Rows.Band.AddNew();
										}
									}
								}
								#endregion Carga Articulo de Garantia Extendida
							}
						}
					}
				}
				#endregion Abre pantalla de Garantia Extendida
			}
		}

		static private void GarantiaExtendidaPlus(Infragistics.Win.UltraWinGrid.UltraGrid uGrid, DateTime dtFecha, int iCuotas, int idTipoPrecio, int idEntidadFinanciera, int idCombo, int iFila, C1.Data.C1DataSet ds)
		{
			string sIdArticulos = "";							
			
			#region Agrupa IdArticulos
			foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in uGrid.Rows.All)
			{
				if (dr.Cells["idArticulo"].Value != System.DBNull.Value)
				{
					if (sIdArticulos.Length == 0) sIdArticulos = dr.Cells["idArticulo"].Value.ToString();
					else if (sIdArticulos.Length > 0) sIdArticulos = sIdArticulos + ", " + dr.Cells["idArticulo"].Value.ToString();				
				}
			}
			#endregion Agrupa IdArticulos

			#region Abre pantalla de Garantia Extendida
			if (sIdArticulos.Length > 0)
			{
				if (Funcion.iEscalarSQL(ds, string.Format("Exec ValidaArticulosGarantiaExtendidaYV '{0}'", sIdArticulos)) > 0)
				{
					using (frmGarantiaExtendida GE = new frmGarantiaExtendida(uGrid, iCuotas, dtFecha, idEntidadFinanciera, idTipoPrecio))
					{				
						if (DialogResult.OK == GE.ShowDialog())
						{
							#region Carga Articulo de Garantia Extendida
							if (GE.ultraGrid1.Rows.Count > 0)
							{
								decimal dIVA = Funcion.decEscalarSQL(ds, string.Format("Select dbo.RetornaImpuestoIVAVigente('{0}', 1)", dtFecha.ToString("yyyyMMdd")));
									
								foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in GE.ultraGrid1.Rows.All)
								{										
									if ((bool)dr.Cells["Sel"].Value)
									{								
											
										uGrid.DisplayLayout.Bands[0].AddNew();	
										iFila++;
											
										int idArticuloGExCP = 4497;
										if ((int)dr.Cells["TipoProducto"].Value == 2) idArticuloGExCP = 8724;
										if ((int)dr.Cells["TipoProducto"].Value == 3) idArticuloGExCP = 8723;

										string sSQLInf = string.Format("Select a.idArticulo, a.Codigo, a.Articulo From Articulo a Where idArticulo = {0}", idArticuloGExCP);
										SqlDataReader drA = Funcion.rEscalarSQL(ds, sSQLInf, true);
										drA.Read();
											
										uGrid.Rows[iFila].Cells["Codigo"].Value = drA.GetString(1);
										uGrid.Rows[iFila].Cells["idArticulo"].Value = drA.GetInt32(0);
										uGrid.Rows[iFila].Cells["Articulo"].Value = drA.GetString(2);
										uGrid.Rows[iFila].Cells["idUnidad"].Value = 1;
										uGrid.Rows[iFila].Cells["Cantidad"].Value = 1;
										uGrid.Rows[iFila].Cells["Precio"].Value = dr.Cells["Precio"].Value;
										uGrid.Rows[iFila].Cells["Impuesto"].Value = dIVA;
										uGrid.Rows[iFila].Cells["DescuentoArt"].Value = 0.00m;
										uGrid.Rows[iFila].Cells["DescuentoPorc"].Value = 0.00m;
										uGrid.Rows[iFila].Cells["Tipo"].Value = 5;
										uGrid.Rows[iFila].Cells["Entregado"].Value = false;
										uGrid.Rows[iFila].Cells["idArticuloGE"].Value = dr.Cells["idArticulo"].Value;
										drA.Close();						
									}
								}
							}
							#endregion Carga Articulo de Garantia Extendida
						}
					}
				}
			}
			#endregion Abre pantalla de Garantia Extendida
		}

		static private void CargaArticulosCombo(Infragistics.Win.UltraWinGrid.UltraGrid uGrid, string sSQL, int idCombo, int idTipoFactura, int iTipoArticulo, int iTipo, int iFila, C1.Data.C1DataSet ds)
		{
			if (iTipo == 0) 
			{
				uGrid.DisplayLayout.Bands[0].AddNew();	
				iFila++;
			}

			#region Carga Articulos en el Grid
			int iFilasCombo = Funcion.dvProcedimiento(ds, sSQL).Table.Rows.Count;
			int iFilasCargadas = 0; 

			foreach(DataRow oDRCombos in Funcion.dvProcedimiento(ds, sSQL).Table.Rows)
			{
//				uGrid.DisplayLayout.Bands[0].AddNew();	
//				iFila++;

				uGrid.Rows[iFila].Cells["Codigo"].Value = oDRCombos["Codigo"].ToString();
				uGrid.Rows[iFila].Cells["idArticulo"].Value = (int)oDRCombos["idArticulo"];				
				uGrid.Rows[iFila].Cells["Articulo"].Value = oDRCombos["Articulo"].ToString();				
				uGrid.Rows[iFila].Cells["Cantidad"].Value = oDRCombos["Cantidad"].ToString();
				if (idTipoFactura != 42) uGrid.Rows[iFila].Cells["Precio"].Value = Convert.ToDecimal(oDRCombos["Precio"]);

				if (idTipoFactura != 8 && idTipoFactura != 9 && idTipoFactura != 42)
				{
					uGrid.Rows[iFila].Cells["idUnidad"].Value = (int)oDRCombos["idUnidad"];
					uGrid.Rows[iFila].Cells["Impuesto"].Value = Convert.ToDecimal(oDRCombos["Iva"]);
					uGrid.Rows[iFila].Cells["DescuentoArt"].Value = Convert.ToDecimal(oDRCombos["ValDescuento"]);
					uGrid.Rows[iFila].Cells["DescuentoPorc"].Value = Convert.ToDecimal(oDRCombos["Descuento"]);
					uGrid.Rows[iFila].Cells["Principal"].Value = (bool)oDRCombos["Principal"];
					uGrid.Rows[iFila].Cells["Tipo"].Value = iTipoArticulo;
					uGrid.Rows[iFila].Cells["Entregado"].Value = false;
					uGrid.Rows[iFila].Cells["Bloqueado"].Value = (bool)oDRCombos["Bloqueado"];
					uGrid.Rows[iFila].Cells["idCombo"].Value = idCombo;
				}

				iFilasCargadas++;
				if (iFilasCargadas < iFilasCombo) uGrid.DisplayLayout.Bands[0].AddNew();	
				iFila++;
			}			
			#endregion Carga Articulos en el Grid
		}

		static public bool ValidaCamposGridDecimal(object sender, Infragistics.Win.UltraWinGrid.CellEventArgs e, string sColumna)
		{			
			if (e.Cell.Row.Cells[sColumna].Value == System.DBNull.Value) 
			{
				e.Cell.Row.Cells[sColumna].Value = Convert.ToDecimal(e.Cell.OriginalValue);
				return false;
			}

			if (e.Cell.Row.Cells["idArticulo"].Value == System.DBNull.Value)
			{							
				if (Convert.ToDecimal(e.Cell.Row.Cells[sColumna].Value) > 0.00m)
				{							
					e.Cell.Value = 0.00m;
					MessageBox.Show("Seleccione un Articulo antes de Ingresar un valor", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
					return false;
				}
			}		

			return true;
		}

		static public void ValidaCamposGridEnteros(object sender, Infragistics.Win.UltraWinGrid.CellEventArgs e, string sColumna)
		{			
			if (e.Cell.Row.Cells[sColumna].Value == System.DBNull.Value) e.Cell.Row.Cells[sColumna].Value = (int)e.Cell.OriginalValue;

			if (e.Cell.Row.Cells["idArticulo"].Value == System.DBNull.Value)
			{							
				if ((int)e.Cell.Row.Cells[sColumna].Value > 1)
				{							
					e.Cell.Value = 0;
					MessageBox.Show("Seleccione un Articulo", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				}
			}
			else if (sColumna == "Cantidad")
			{
				if ((int)e.Cell.Row.Cells[sColumna].Value == 0)
				{							
					e.Cell.Value = 1;
					MessageBox.Show("La Cantidad no puede ser 0", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				}
			}
		}

		static public void AbreBusquedaArticulos(Infragistics.Win.UltraWinGrid.UltraGrid uGrid, int iBodega, int idTipoFactura, int idFormaPago, DateTime dtFecha, bool bConsignado)
		{			
			using (BuscaArticulos Busqueda = new BuscaArticulos(uGrid.ActiveRow.Cells["Codigo"].Value.ToString(), iBodega, idTipoFactura, idFormaPago, dtFecha, bConsignado))
			{
				if (DialogResult.OK == Busqueda.ShowDialog())
				{									
					uGrid.ActiveRow.Cells["Codigo"].Value = Busqueda.sCodigo;
				}
			}			
		}

		static public void AbreBusquedaArticulosCombos(Infragistics.Win.UltraWinGrid.UltraGrid uGrid, string sColumna, int iBodega, int idTipoFactura, int idFormaPago, DateTime dtFecha)
		{			
			using (BuscaArticulos Busqueda = new BuscaArticulos(uGrid.ActiveRow.Cells[sColumna].Value.ToString(), iBodega, idTipoFactura, idFormaPago, dtFecha, false))
			{
				if (DialogResult.OK == Busqueda.ShowDialog())
				{									
					uGrid.ActiveRow.Cells[sColumna].Value = Busqueda.sCodigo;
				}
			}			
		}

		static public void ValidaCostoEIBodega (object sender, Infragistics.Win.UltraWinGrid.BeforeCellUpdateEventArgs e, C1.Data.C1DataSet ds)
		{			
			if (e.Cell.Row.Cells["idArticulo"].Value != System.DBNull.Value)
			{
				if (Funcion.bEscalarSQL(ds, string.Format("Declare @Res Bit = 0 If Exists(Select idUltimoCosto From UltimoCosto Where idArticulo = {0}) Set @Res = 1 Select @Res", (int)e.Cell.Row.Cells["idArticulo"].Value), true))
				{
					MessageBox.Show(string.Format("El articulo {0} tiene costo registrado en Compras", e.Cell.Row.Cells["Codigo"].Value.ToString()), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
					e.Cancel = true;
					return;
				}
			}			
		}

		static public bool CambioColorSerialesIngresadas (Infragistics.Win.UltraWinGrid.UltraGrid uGrid, Infragistics.Win.UltraWinGrid.UltraGrid uGridSeriales)
		{
			foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in uGridSeriales.Rows.All)
			{				
				if ((int)dr.Cells["idArticulo"].Value == (int)uGrid.ActiveRow.Cells["idArticulo"].Value && 
					(int)dr.Cells["Posicion"].Value == (int)uGrid.ActiveRow.Cells["Posicion"].Value) 
				{
					MessageBox.Show(string.Format("El articulo {0} tiene serial registrada", uGrid.ActiveRow.Cells["Codigo"].Value.ToString()), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);					
					return false;
				}
			}

			return true;
		}

		static public void EliminaSerialesRemates (Infragistics.Win.UltraWinGrid.BeforeRowsDeletedEventArgs e, Infragistics.Win.UltraWinGrid.UltraGrid uGrid, int idTipoFactura, int idCompra, int iTipo, C1.Data.C1DataSet ds)
		{
			for (int i=0; i<e.Rows.Length; i++)
			{
				foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in uGrid.Rows.All)
				{				
					if ((int)dr.Cells["idArticulo"].Value == (int)e.Rows[i].Cells["idArticulo"].Value && (int)dr.Cells["Posicion"].Value == (int)e.Rows[i].Cells["Posicion"].Value) 
					{
						if ((int)e.Rows[i].Cells["idDetCompra"].Value == 0) dr.Delete();
						else
						{
							string sSQL = string.Format("Exec BorrarSeriales {0}, {1}, {2}, {3}, {4}, {5}, {6}", 
								idTipoFactura, (int)dr.Cells["idSerial"].Value, 
								(int)dr.Cells["idDetalleSerial"].Value, (int)dr.Cells["idOrigen"].Value, 
								(int)dr.Cells["idDetCompra"].Value, idCompra, iTipo);
							Funcion.EjecutaSQL(ds, sSQL, true);
						}
					}
				}
			}
		}

		static public bool ValidaEliminaFilasGrids(object sender, Infragistics.Win.UltraWinGrid.BeforeRowsDeletedEventArgs e, int idTipoFactura, Infragistics.Win.UltraWinGrid.UltraGrid uGrid, C1.Data.C1DataSet ds)
		{			
			if (uGrid.Rows.Count > 0)
			{
				MessageBox.Show(string.Format("Existen {0} seriales registradas, \n\nElimine primero las seriales", uGrid.Rows.Count), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				e.Cancel = true;
				return false;
			}

			for (int i=0; i<e.Rows.Length; i++)
			{
				#region Valida Seriales Ingreso de Bodega - Venta
				if (idTipoFactura == 1 || idTipoFactura == 9)
				{					
					int iCantSeriales = 0;
					if (Funcion.bEscalarSQL(ds, string.Format("Select ISNULL(ManejaSerial, 0) From Articulo Where idArticulo = {0}", (int)e.Rows[i].Cells["idArticulo"].Value), true))
					{
						foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in uGrid.Rows.All)
						{
							if ((int)dr.Cells["idArticulo"].Value == (int)e.Rows[i].Cells["idArticulo"].Value && (int)dr.Cells["Posicion"].Value == (int)e.Rows[i].Cells["Posicion"].Value) iCantSeriales++;
						}
					}			
			          
					if (iCantSeriales > 0)
					{
						MessageBox.Show(string.Format("Este Artículo tiene {0} seriales registradas, \n\nElimine primero las seriales", iCantSeriales), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
						e.Cancel = true;						
						return false;
					}
				}
				#endregion Valida Seriales Ingreso de Bodega - Venta
			}

			return true;
		}

		static public bool EliminaFilasGrids(System.Windows.Forms.Button oButton, object sender, Infragistics.Win.UltraWinGrid.BeforeRowsDeletedEventArgs e, C1.Data.C1DataSet ds)
		{
			if (oButton.Enabled && DialogResult.Yes ==	MessageBox.Show("¿Desea Borrar Líneas Seleccionadas?", "Confirmación", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2))
			{
				#region Elimina Articulo
				for (int i=0; i<e.Rows.Length; i++)
				{					
					if ((int)e.Rows[i].Cells["idDetCompra"].Value > 0) 
					{
						oButton.Enabled = false;
						string sSQL = string.Format("Delete From DetCompra Where idDetCompra = '{0}'", (int)e.Rows[i].Cells["idDetCompra"].Value);
						Funcion.EjecutaSQL(ds, sSQL, true);
						e.DisplayPromptMsg = false;
					}
					else e.DisplayPromptMsg = false;					
				}
				#endregion Elimina Articulo

				return true;
			}
			else return false;      
		}

		static public void CargaGridSerial(Infragistics.Win.UltraWinGrid.UltraGrid uGrid, bool bModificaSerial, int idSerial, int idOrigen, int idDetalleSerial, int idArticulo, string sSerial, int iEstado, int iPosicion)
		{
			uGrid.Rows.Band.AddNew();
			uGrid.ActiveRow.Cells["idSerial"].Value = idSerial;
			uGrid.ActiveRow.Cells["idOrigen"].Value = idOrigen;
			if (!bModificaSerial)
			{
				uGrid.ActiveRow.Cells["idDetalleSerial"].Value = idDetalleSerial;			
				uGrid.ActiveRow.Cells["idArticulo"].Value = idArticulo;
			}
			uGrid.ActiveRow.Cells["Serial"].Value = sSerial;
			uGrid.ActiveRow.Cells["Estado"].Value = iEstado;
			uGrid.ActiveRow.Cells["Posicion"].Value = iPosicion;
		}

		static public void EliminaFilasGridsSeriales(object sender, Infragistics.Win.UltraWinGrid.BeforeRowsDeletedEventArgs e, int idTipoFactura, int idCompra, int iTipo, C1.Data.C1DataSet ds)
		{
			#region Validacion
			for (int i=0; i<e.Rows.Length; i++)
			{	
				if ((int)e.Rows[i].Cells["idDetalleSerial"].Value > 0)
				{
					string sSQL = string.Format("Exec ValSerialBorrar {0}, {1}, {2}", 
						(int)e.Rows[i].Cells["idSerial"].Value, idTipoFactura, (int)e.Rows[i].Cells["idDetCompra"].Value);

					if (!Funcion.bEscalarSQL(ds, sSQL, true))
					{
						MessageBox.Show(string.Format("No puede Eliminar esta serial '{0}' ya ha sido utilizada", e.Rows[i].Cells["Serial"].Value.ToString()), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
						e.Cancel = true;
						return;
					}
				}
			}
			#endregion Validacion

			#region Elimina
			if (DialogResult.Yes ==	MessageBox.Show("¿Esta seguro de borrar las seriales seleccionadas?", "Point Technology", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2))
			{				
				for (int i=0; i < e.Rows.Length; i++)
				{	
					if ((int)e.Rows[i].Cells["idDetalleSerial"].Value > 0)
					{
						string sSQL = string.Format("Exec BorrarSeriales {0}, {1}, {2}, {3}, {4}, {5}, {6}", 
							idTipoFactura, (int)e.Rows[i].Cells["idSerial"].Value, 
							(int)e.Rows[i].Cells["idDetalleSerial"].Value, (int)e.Rows[i].Cells["idOrigen"].Value, 
							(int)e.Rows[i].Cells["idDetCompra"].Value, idCompra, iTipo);
						Funcion.EjecutaSQL(ds, sSQL, true);

						e.DisplayPromptMsg = false;
					}
					else e.DisplayPromptMsg = false;
				}					
			}
			else e.Cancel = true;
			#endregion Elimina
		}
	}
}
