using System;
using System.Data;
using System.Data.SqlClient;
using System.Windows.Forms;
using Infragistics;
using Infragistics.Win.UltraWinGrid;
using System.Globalization;
using System.Data.OleDb;
using System.Data.Odbc;
using DevExpress.XtraEditors;
using CrystalDecisions;
using CrystalDecisions.CrystalReports.Engine;
using CrystalDecisions.ReportSource;
using CrystalDecisions.Shared;

namespace Latinium
{
	/// <summary>
	/// Descripción breve de Facturacion.
	/// </summary>
	public class Facturacion
	{
		public Facturacion()
		{
			//
			// TODO: agregar aquí la lógica del constructor
			//
		}
		
		#region Cupones Web
		static public void ActualizaEstadoCuponWeb (int idCuponesWeb, int Estado, int idCompra, C1.Data.C1DataSet ds)
		{
			Funcion.EjecutaSQL(ds, string.Format("Exec ActualizaEstadoCuponWeb {0}, {1}, {2}", idCuponesWeb, Estado, idCompra));	
		}

		static public int iPromocionMail(C1.Data.C1DataSet ds)
		{							
			int idMail = 0;
						
			using (frmCorreosMarketing CM = new frmCorreosMarketing())
			{				
				if (DialogResult.OK == CM.ShowDialog())
				{						
					idMail = CM.idCorreo;
				}			
			}			

			return idMail;
		}

		static public string[] sCuponWeb(C1.Data.C1DataSet ds)
		{			
			string[] sParametros = new string[2];
			
			int idCuponWeb = 0;
			decimal dPorcentaje = 0.00m;

			sParametros[0] = idCuponWeb.ToString();
			sParametros[1] = dPorcentaje.ToString();

			if (Funcion.bEscalarSQL(ds, "Select CuponWeb From SeteoF", true))
			{
				using (frmCuponWeb CuponWeb = new frmCuponWeb())
				{				
					if (DialogResult.OK == CuponWeb.ShowDialog())
					{						
						idCuponWeb = CuponWeb.idCuponWeb;
						dPorcentaje = Funcion.decEscalarSQL(ds, "Select PorcentajeCuponWeb From SeteoF");
						
						ActualizaEstadoCuponWeb(idCuponWeb, 1, 0, ds);						

						sParametros[0] = idCuponWeb.ToString();
						sParametros[1] = dPorcentaje.ToString();
					}			
				}
			}

			return sParametros;
		}

		static public string[] sEditarCuponWeb(int idCompra, C1.Data.C1DataSet ds)
		{			
			string[] sParametros = new string[2];
			
			int idCuponWeb = 0;
			decimal dPorcentaje = 0.00m;

			sParametros[0] = idCuponWeb.ToString();
			sParametros[1] = dPorcentaje.ToString();

			if (Funcion.bEscalarSQL(ds, "Select CuponWeb From SeteoF", true))
			{
				idCuponWeb = Funcion.iEscalarSQL(ds, string.Format("Declare @id Int = 0 If Exists(Select idCuponesWeb From CuponesWeb Where Estado = 2 And idCompra = {0}) Select @id = idCuponesWeb From CuponesWeb Where Estado = 2 And idCompra = {0} Select @id", idCompra));
				dPorcentaje = Funcion.decEscalarSQL(ds, "Select PorcentajeCuponWeb From SeteoF");
				
				sParametros[0] = idCuponWeb.ToString();
				sParametros[1] = dPorcentaje.ToString();
			}			

			return sParametros;
		}
		#endregion Cupones Web	

		static public void IniciaValoresColumnas(object sender, Infragistics.Win.UltraWinGrid.RowEventArgs e, Infragistics.Win.UltraWinGrid.UltraGrid uGrid, int idTipoFactura, int iBodega)
		{
			if (idTipoFactura == 1 || idTipoFactura == 2 || idTipoFactura == 3 || idTipoFactura == 5 || idTipoFactura == 13 || idTipoFactura == 43 || idTipoFactura == 44) e.Row.Cells["idOrigen"].Value = 0;				

			e.Row.Cells["idDetCompra"].Value = 0;				
			e.Row.Cells["idCompra"].Value = 0;						
			e.Row.Cells["Precio"].Value = 0.00m;
			e.Row.Cells["Subtotal"].Value = 0.00m;			
			e.Row.Cells["Cantidad"].Value = 1;	

			if (idTipoFactura == 1 || idTipoFactura == 3 || idTipoFactura == 5 || idTipoFactura == 13 || idTipoFactura == 43 || idTipoFactura == 44) e.Row.Cells["Jornadas"].Value = 0.00m;

			if (idTipoFactura == 1 || idTipoFactura == 3 || idTipoFactura == 5 || idTipoFactura == 9 || 
				idTipoFactura == 8 || idTipoFactura == 13 || idTipoFactura == 43 || idTipoFactura == 44) e.Row.Cells["Posicion"].Value = uGrid.Rows.Count;	
			      
			if (idTipoFactura != 8 && idTipoFactura != 9 && idTipoFactura != 36)
			{
				e.Row.Cells["Impuesto"].Value = 0.00m;
				e.Row.Cells["DescuentoArt"].Value = 0.0000m;
				e.Row.Cells["DescuentoPorc"].Value = 0.0000m;
				e.Row.Cells["Subtotal"].Value = 0.00m;
			}
			
			if (idTipoFactura == 1 || idTipoFactura == 3 || idTipoFactura == 5 || idTipoFactura == 13|| idTipoFactura == 43 || idTipoFactura == 44)
			{				
				e.Row.Cells["EstadoSerial"].Value = false;
				e.Row.Cells["Tipo"].Value = 0;
				e.Row.Cells["Entregado"].Value = false;
				e.Row.Cells["Principal"].Value = false;
				e.Row.Cells["Bloqueado"].Value = false;
				e.Row.Cells["idCombo"].Value = 0;				
				e.Row.Cells["idArticuloGE"].Value = 0;				
			}		

			if (idTipoFactura == 2)
			{
				e.Row.Cells["CantAereo"].Value = 0;			
				e.Row.Cells["Valores"].Value = 0.00m;
				e.Row.Cells["DescNC"].Value = 0.00m;
								
				e.Row.Cells["Tipo"].Value = 1;
				e.Row.Cells["EstadoSerial"].Value = false;

				e.Row.Cells["Peso"].Value = 0.00m;
				e.Row.Cells["Courier"].Value = 0.00m;				
				e.Row.Cells["Arancel"].Value = 0.00m;				
			}
		}				

		static public bool bValidaDescuentos(Infragistics.Win.UltraWinGrid.CellEventArgs e, Infragistics.Win.UltraWinGrid.UltraGrid uGrid, int iBodega, bool bDescuentoActivo, decimal dClavePorcDscto, 
			int idTipoFactura, bool bAcceso, int idTipoPrecio, int idBanco, int idTarjeta, int idPlan, int idPlazo, decimal dPorcentajeCuponWeb, decimal dDsctoReiterativo, bool bCupon, int idMail, bool bEmpleado, C1.Data.C1DataSet ds)
		{			
			if (e.Cell.Row.Cells["idArticulo"].Value == System.DBNull.Value) return false;
			if (e.Cell.Row.Cells["DescuentoPorc"].Value == System.DBNull.Value) return false;
				
			if (!bAcceso || idTipoPrecio == 94 || idTipoPrecio == 95) 
			{
				if (e.Cell.OriginalValue != null)
				{
					try
					{
						e.Cell.Value = Convert.ToDecimal(e.Cell.OriginalValue);
						MessageBox.Show("No tiene Acceso a Modificar Descuentos", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Information);	
					}
					catch
					{
						// Si no se puede convertir, asigna 0 o deja el valor original
						e.Cell.Value = 0m; // o e.Cell.OriginalValue
					}
				}
				else
				{
					e.Cell.Value = 0m; // Valor por defecto si es null
				}			
				return false;
			}
		
			if (!bDescuentoActivo && (idTipoFactura == 13 || idTipoFactura == 44))
			{
				if (e.Cell.OriginalValue != System.DBNull.Value) e.Cell.Value = (decimal)e.Cell.OriginalValue; else e.Cell.Value = 0.00m;
				MessageBox.Show("No esta autorizado a modificar descuentos", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Information);
				return false;
			}
			
			if ((decimal)e.Cell.Row.Cells["DescuentoPorc"].Value > 100.00m)
			{								
				e.Cell.Value = (decimal)e.Cell.OriginalValue;
				MessageBox.Show("El Porcentaje de Descuento Maximo es 100%", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				return false;
			}

			if (!bDescuentoActivo)
			{
				#region Sin Clave de Descuento
				if (!bEmpleado)
				{
					if (Funcion.iEscalarSQL(ds, string.Format("Select COUNT(*) From DetallePromocion dp Inner Join CompraPromocion cp On cp.idPromocion = dp.idPromocion Where dp.idArticulo = {0} And Promocional = 1 And Estado = 'ACTIVO'", (int)e.Cell.Row.Cells["idArticulo"].Value)) > 0)
					{
						e.Cell.Value = (decimal)e.Cell.OriginalValue;
						MessageBox.Show("Articulo Promocional No puede Modificar Descuento", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
						return false;
					}
				
					if ((int)e.Cell.Row.Cells["Tipo"].Value >= 1 && !(bool)e.Cell.Row.Cells["Principal"].Value)
					{
						e.Cell.Value = (decimal)e.Cell.OriginalValue;
						MessageBox.Show("Articulo Promocional No puede Modificar Descuento", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
						return false;
					}
				}

				int idArticulo = (int)e.Cell.Row.Cells["idArticulo"].Value;
				if ((int)e.Cell.Row.Cells["idCombo"].Value > 0 && !(bool)e.Cell.Row.Cells["Principal"].Value && (int)e.Cell.Row.Cells["Tipo"].Value == 1)
				{
					foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in uGrid.Rows.All)
					{
						if ((bool)dr.Cells["Principal"].Value && ((int)e.Cell.Row.Cells["idCombo"].Value == (int)dr.Cells["idCombo"].Value)) 
						{
							idArticulo = (int)dr.Cells["idArticulo"].Value;
							continue;
						}
					}
				}

				decimal dDescuento = 0.00m;
				if ((int)e.Cell.Row.Cells["Tipo"].Value == 2 && (bool)e.Cell.Row.Cells["Principal"].Value)
				{
					dDescuento = Funcion.decEscalarSQL(ds, string.Format("Exec RetornaDescuentoPromociones {0}, {1}, {2}", idArticulo, idTipoPrecio, (int)e.Cell.Row.Cells["idCombo"].Value));
				}
				else
				{
					dDescuento = Funcion.decEscalarSQL(ds, string.Format("Select dbo.ValidaLimiteDescuentoPromocion({0}, {1}, {2}, {3}, {4}, {5}, {6}, {7}, {8})",
						iBodega, idArticulo, idTipoPrecio, (decimal)e.Cell.Row.Cells["DescuentoPorc"].Value, idBanco, idTarjeta, idPlan, idPlazo, Convert.ToInt32(bCupon)));
				}
				
				decimal dSumaDescuento = Funcion.decEscalarSQL(ds, string.Format("Select dbo.ValidaSumaDescuentoPromocion({0}, {1}, {2}, {3}, {4}, {5}, {6}, {7})",
					iBodega, idArticulo, idTipoPrecio, idBanco, idTarjeta, idPlan, idPlazo, Convert.ToInt32(bCupon)));
				
				decimal dDescuentoPermitido = dDescuento + dSumaDescuento;			

				if ((int)e.Cell.Row.Cells["Tipo"].Value == 0 || (bool)e.Cell.Row.Cells["Principal"].Value) 
					dDescuentoPermitido = dDescuentoPermitido + dPorcentajeCuponWeb;

				dDescuentoPermitido = dDescuentoPermitido + dDsctoReiterativo;

				if (dDescuentoPermitido > 100.00m) dDescuentoPermitido = 100.00m;

				//if (dDsctoReiterativo > dDescuentoPermitido) dDescuentoPermitido = dDsctoReiterativo;

				if (idMail > 0)
				{
					if (idTipoPrecio == 4) dDescuentoPermitido = 5.00m;
					else dDescuentoPermitido = 3.00m;
				}

				if ((decimal)e.Cell.Row.Cells["DescuentoPorc"].Value > dDescuentoPermitido)
				{
					e.Cell.Value = (decimal)e.Cell.OriginalValue;
					MessageBox.Show("El Descuento debe estar en el Rango Permitido", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);				
					return false;
				}
				#endregion Sin Clave de Descuento
			}			
			else
			{
				#region Con Clave de Descuento
				decimal dDescuentoTotal = 0.00m;
//				if (idTipoFactura == 1 || idTipoFactura == 43) dDescuentoTotal = dClavePorcDscto + Convert.ToDecimal(e.Cell.OriginalValue);
//				else 
//				{
				if (e.Cell.OriginalValue != System.DBNull.Value) dDescuentoTotal = dClavePorcDscto; else dDescuentoTotal = 0.00m;
//				}
				
				if ((int)e.Cell.Row.Cells["Tipo"].Value == 0 || (bool)e.Cell.Row.Cells["Principal"].Value) 
					dDescuentoTotal = dDescuentoTotal + dPorcentajeCuponWeb;

				if ((decimal)e.Cell.Row.Cells["DescuentoPorc"].Value > dDescuentoTotal)
				{
					e.Cell.Value = (decimal)e.Cell.OriginalValue;
					MessageBox.Show("El Descuento no puede ser mayor al Permitido", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);				
					return false;
				}
				#endregion Con Clave de Descuento
			}	

			return true;
		}

		static public bool bValidaImpuestos(Infragistics.Win.UltraWinGrid.CellEventArgs e, DateTime dtFecha, int iFiscal, C1.Data.C1DataSet ds)
		{
			if ((decimal)e.Cell.Row.Cells["Impuesto"].Value > 0.00m)
			{
				decimal dPorcIVA = Funcion.decEscalarSQL(ds, string.Format("Select dbo.RetornaImpuestoIVAVigente('{0}', {1})", dtFecha.ToString("yyyyMMdd"), iFiscal));

				if ((decimal)e.Cell.Row.Cells["Impuesto"].Value != dPorcIVA)
				{
					e.Cell.Value = dPorcIVA;
					MessageBox.Show(string.Format("El Porcentaje de I.V.A. debe ser 0% o {0}%", dPorcIVA), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
					return false;
				}
			}

			return true;
		}
		
		static public bool bValidaModificacionPrecios(Infragistics.Win.UltraWinGrid.CellEventArgs e, bool bAcceso, int idTipoFactura, int idTipoPrecio, int idCliente, C1.Data.C1DataSet ds)
		{
			if (e.Cell.Row.Cells["idArticulo"].Value == DBNull.Value) return true;
			if (e.Cell.Row.Cells["Tipo"].Value == DBNull.Value) return true;
			if (e.Cell.Row.Cells["idCombo"].Value == DBNull.Value) return true;
			
			if (!bAcceso) 
			{
				e.Cell.Value = (decimal)e.Cell.OriginalValue;
				MessageBox.Show("No tiene Acceso a Modificar Precios", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Information);				
				return false;
			}

			if ((int)e.Cell.Row.Cells["Tipo"].Value == 5) 
			{
				e.Cell.Row.Cells["Precio"].Value = (decimal)e.Cell.OriginalValue;
				MessageBox.Show("En Garantia Extendida no puede modificar precio", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				return false;
			}

			string sSQL = string.Format("Exec ValidaPreciosFacturacionClientes {0}, {1}, {2}, {3}, {4}, {5}", 
				(int)e.Cell.Row.Cells["idArticulo"].Value, idTipoFactura, idTipoPrecio, (int)e.Cell.Row.Cells["Tipo"].Value, 
				(int)e.Cell.Row.Cells["idCombo"].Value, idCliente);
			decimal dPrecio = Funcion.decEscalarSQL(ds, sSQL);

			if ((decimal) e.Cell.Row.Cells["Precio"].Value < dPrecio) 
			{
				e.Cell.Row.Cells["Precio"].Value = e.Cell.OriginalValue;//dPrecio;
				MessageBox.Show(string.Format("No puede poner precio menor a {0}", dPrecio), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				return false;
			}
			
			return true;
		}

		static public bool ValidaEdicionAnulacion(int idCompra, Infragistics.Win.UltraWinGrid.UltraGrid uGrid, bool bAcceso, int iEstado, int idTipoFactura, int iBodega, DateTime dtFecha, 
			bool bEdicionAnulacion, bool bAnulado, int iTipo, C1.Data.C1DataSet ds,
			bool bImpreso, bool bAccesoProcesado, bool bAccesoAprobado, bool bAccesoConSeriales, string sOperacion, int idMotivo, bool bFiscal)
		{	
			string sAccion = "Editar";
			if (bEdicionAnulacion) sAccion = "Anular";

			if (bAnulado)
			{
				MessageBox.Show("El documento esta anulado", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);						
				return false;
			}

			if (!bAcceso)
			{
				MessageBox.Show("No tiene acceso a modificar este documento", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);						
				return false;
			}
			
			if (!bEdicionAnulacion && bImpreso && (idTipoFactura == 1  || (idTipoFactura == 5 && bFiscal) || idTipoFactura == 43))
			{
				MessageBox.Show("No puede editar el documento esta impreso", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);				
				return false;
			}

			if (iEstado == 4 && !bEdicionAnulacion)
			{
				MessageBox.Show("El documento fue Rechazado", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);								
				return false;
			}

			if (!bEdicionAnulacion && iEstado == 5 && (idTipoFactura == 5 || idTipoFactura == 43) && !bAccesoAprobado)
			{
				MessageBox.Show("El documento esta aprobado", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);								
				return false;
			}

			if (idTipoFactura == 13)
			{
				if (iEstado == 3 && !bAccesoProcesado)
				{
					MessageBox.Show("El Documento esta PROCESADO", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);				
					return false;
				}	

				if (iEstado == 5 && !bAccesoAprobado)
				{
					MessageBox.Show("El Documento esta APROBADO", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
					return false;
				}					
			}

			if (idTipoFactura == 5)
			{
				if ((iEstado == 5 || iEstado == 7) && !bAccesoAprobado)
				{
					MessageBox.Show("El Documento esta APROBADO", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
					return false;
				}
				
				if (iEstado == 8 && !bAccesoConSeriales)
				{
					MessageBox.Show("El Documento esta LIQUIDADO", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);				
					return false;
				}

				if (!bEdicionAnulacion && idMotivo == 6 && Funcion.iEscalarSQL(ds, string.Format("Select COUNT(*) from AutorizacionDePagos Where Estado = 1 And idTipo = 14 And idOrigen = {0}", idCompra)) > 0)
				{
					MessageBox.Show("No puede editar el documento tiene pagos autorizados", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);				
					return false;
				}

				if (bEdicionAnulacion && idMotivo == 6 && Funcion.iEscalarSQL(ds, string.Format("Select COUNT(*) from AutorizacionDePagos Where Estado = 2 And idTipo = 14 And idOrigen = {0}", idCompra)) > 0)
				{
					MessageBox.Show("No puede anular el documento esta cancelado", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);				
					return false;
				}

				if (!bEdicionAnulacion && idTipoFactura == 5 && Funcion.bEscalarSQL(ds, string.Format("Exec NDCValidaRegistroIRM {0}", idCompra), true))
				{
					MessageBox.Show("El documento proviene de Ingreso de Mercaderia de Retiro", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);				
					return false;
				}
			}

//			if (!bEdicionAnulacion && iEstado == 8)
//			{
//				MessageBox.Show("El Documento esta CANCELADO", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
//				return false;
//			}

			if (iEstado == 9)
			{
				MessageBox.Show("El Documento esta FACTURADO", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				return false;
			}		

			if (Funcion.ValidaBloqueo(idCompra, 1, ds)) return false;

			if (Funcion.ValidaPeriodo(dtFecha, ds, sOperacion))
			{				
				MessageBox.Show(string.Format("No puede modificar documentos en esta fecha, El período '{0}' esta cerrado,\n\nConsulte al Administrador", dtFecha.ToString("MMMM").ToUpper()), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);									
				return false;
			}

			if (iTipo > 0 && (idTipoFactura == 8 || idTipoFactura == 9))
			{
				string sMensaje = "TRANSFERENCIA";
				if (iTipo == 2) sMensaje = "PRODUCCION";
				if (iTipo == 3) sMensaje = "INGRESO REMATES";
				if (iTipo == 4) sMensaje = "AJUSTE RETORNO MERCADERIA";

				MessageBox.Show("No puede modificar registros ingresados desde " + sMensaje, "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);					
				return false;
			}

			if (idTipoFactura == 5 || idTipoFactura == 8 || idTipoFactura == 9 || idTipoFactura == 13 || idTipoFactura == 44)			
				if (!FuncionesInventario.ValidaBodegaEnInventario(iBodega, dtFecha, ds)) return false;
			

			if (!FuncionesInventario.ValidaBodegaInventarios(uGrid, iBodega, dtFecha, ds)) return false;

			if (idTipoFactura == 2 || idTipoFactura == 13)
			{
				if (Funcion.ValidaBloqueoERMEnTransacciones(idCompra, ds)) return false;

				if (bEdicionAnulacion)
					if (FuncionesInventario.ValidaSerialesIngresadasParaAnular(idCompra, idTipoFactura, ds) > 0 && !bAccesoConSeriales) return false;				
			}
						
			return true;
		}

		static public bool AnulacionCompraVenta (int idCompra, int idTipoFactura, C1.Data.C1DataSet ds)
		{
			bool bRes = false;

			if (DialogResult.Yes == MessageBox.Show("Esta seguro de anular este documento", "Point Technology", MessageBoxButtons.YesNo, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button2)) 
			{
				using (frmJustAnulaVenta Nota = new frmJustAnulaVenta(20))
				{
					if (DialogResult.OK == Nota.ShowDialog())
					{
						try
						{							
							string sSQL = string.Format("Exec AnulacionCompraVenta {0}, {1}, '{2}'", 
								idCompra, idTipoFactura, Nota.txtNotas.Text.ToString());
							Funcion.sEscalarSQL(ds, sSQL);

							bRes = true;
						}
						catch(Exception Exc)
						{
							MessageBox.Show(string.Format("Error al anular: {0}", Exc), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);							
						}
					}
				}				
			}

			return bRes;
		}		

		static public void CambiaColorProducto(Infragistics.Win.UltraWinGrid.UltraGrid uGrid, int idTipoFactura, int iBodega, C1.Data.C1DataSet ds)
		{
			if ((int)uGrid.ActiveRow.Cells["Tipo"].Value > 0)
			{
				if (Funcion.iEscalarSQL(ds, string.Format("Exec ValidaColoresArticulosPromocion {0}", (int)uGrid.ActiveRow.Cells["idArticulo"].Value)) <= 1)
				{
					MessageBox.Show(string.Format("El Articulo {0} solo tiene un color", uGrid.ActiveRow.Cells["Codigo"].Value.ToString().Trim()), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Information);
					return;
				}

				if (idTipoFactura == 1)
				{
					string sSQL = string.Format("Exec ListaArticulosPromocionPorColor {0}, {1}", (int)uGrid.ActiveRow.Cells["idArticulo"].Value, iBodega);
					if (Funcion.dvProcedimiento(ds, sSQL).Count == 0)
					{
						MessageBox.Show("Este Articulo no tiene Existencia en otros Colores", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Information);				
						return;
					}
				}

				using (frmColoresProductos Colores = new frmColoresProductos((int)uGrid.ActiveRow.Cells["idArticulo"].Value, iBodega))
				{				
					if (DialogResult.OK == Colores.ShowDialog())
					{					
						uGrid.ActiveRow.Cells["idArticulo"].Value = (int)Colores.ultraGrid1.ActiveRow.Cells["idArticulo"].Value;
						uGrid.ActiveRow.Cells["Codigo"].Value = Colores.ultraGrid1.ActiveRow.Cells["Codigo"].Value.ToString();
						uGrid.ActiveRow.Cells["Articulo"].Value = Colores.ultraGrid1.ActiveRow.Cells["Articulo"].Value.ToString();					
					}
				}
			}
		}

		static public void Impresion(int idCompra, bool bAnulado, int iEstado, bool bAcceso, bool bAccesoReimprimir, bool bFiscal, bool bImpreso, int idTipoFactura, decimal dMonto, int iCuotas, int idCliente, string sFecha, DateTime dtFecha, int idEntidadFinanciera, string sRuc, C1.Data.C1DataSet ds)
		{
			if (bAnulado)
			{
				MessageBox.Show("El documento esta anulado", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);						
				return;
			}

			if (!bAcceso)
			{
				MessageBox.Show("No tiene acceso a Imprimir", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);						
				return;
			}

			if (!bAccesoReimprimir && bImpreso && bFiscal && (idTipoFactura == 1 || idTipoFactura == 5 || idTipoFactura == 43))
			{
				MessageBox.Show("No tiene acceso a Re Imprimir", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);						
				return;
			}

			if (idTipoFactura == 5 && iEstado == 1)
			{
				MessageBox.Show("El documento no esta aprobado", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);						
				return;
			}

			if (Funcion.ValidaBloqueo(idCompra, 1, ds)) return;

			if (idTipoFactura != 43)
			{
				string sReporte = Funcion.sEscalarSQL(ds, string.Format("Select ISNULL(ReporteImpresion, '') From CompraNumero Where idTipoFactura = {0}", idTipoFactura));

				if (sReporte.Length == 0)
				{
					MessageBox.Show("Este tipo de documento no tiene reporte predefinido para impresión\n\nComuniquese con Sistemas", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);						
					return;
				}

				string stFiltro = "{Compra.idCompra} = " + idCompra.ToString();						
				Reporte Reporte = new Reporte(sReporte, stFiltro);
				Reporte.ShowDialog();

				Funcion.EjecutaSQL(ds, string.Format("Update Compra Set Impreso = 1 Where idCompra = {0}", idCompra));

				return;
			}
		
			if (idTipoFactura == 43)
			{
				DocsPrefactura Dctos = new DocsPrefactura (idCompra, dMonto, iCuotas, idCliente, sFecha, dtFecha, idEntidadFinanciera, sRuc);
				Dctos.ShowDialog();	

				frmDocsPreFacturaEnBlanco DctosBlanco = new frmDocsPreFacturaEnBlanco (idCompra, dMonto, iCuotas, idCliente, sFecha, dtFecha, idEntidadFinanciera, sRuc);
				DctosBlanco.ShowDialog();
			}
		}

		static public void ImpresionRecibos(int idAnticipo, int iTipo, C1.Data.C1DataSet ds)
		{
			string sReporte = "";

			ParameterFields paramFields = new ParameterFields ();

			ParameterField paramFieldIdAnticipo = new ParameterField ();
			ParameterDiscreteValue discreteValIdAnticipo = new ParameterDiscreteValue ();
			paramFieldIdAnticipo.ParameterFieldName = "@idAnticipo";
			discreteValIdAnticipo.Value = idAnticipo; 
			paramFieldIdAnticipo.CurrentValues.Add (discreteValIdAnticipo);
			paramFields.Add (paramFieldIdAnticipo);
			
			sReporte = Funcion.sEscalarSQL(ds, string.Format("Select Reporte From TiposDeAnticipos Where Tipo = {0}", iTipo));

//			if (iTipo == 1 || iTipo == 6 || iTipo == 11) sReporte = "ReporteAnticipo.rpt";
//			else if (iTipo == 2 || iTipo == 14) sReporte = "ReciboCrediPOINT.rpt";
//			else if (iTipo == 3) sReporte = "AbonoCXC.rpt";
//			else if (iTipo == 4 || iTipo == 5 || iTipo == 7 || iTipo == 8 || iTipo == 12) sReporte = "ReporteCobranzaSolidario.rpt";

			Reporte Reporte = new Reporte(sReporte, "");
			Reporte.crVista.ParameterFieldInfo = paramFields;
			Reporte.Show();

			Funcion.EjecutaSQL(ds, string.Format("Update Anticipos Set Impreso = 1 Where idAnticipo = {0}", idAnticipo));			
		}

		static private void EliminaGarantiaExtendida(Infragistics.Win.UltraWinGrid.UltraGrid uGrid, int idArticulo, bool bNuevo, C1.Data.C1DataSet ds)
		{
			#region Validaciones
			int iCont = 0;
			int iContGE = 0;

			foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr1 in uGrid.Rows.All)
			{
				if (dr1.Cells["idArticulo"].Value != System.DBNull.Value)
					if (idArticulo == (int)dr1.Cells["idArticulo"].Value) iCont++;
			}

			foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr2 in uGrid.Rows.All)
			{
				if (dr2.Cells["idArticulo"].Value != System.DBNull.Value)
					if (idArticulo == (int)dr2.Cells["idArticuloGE"].Value) iContGE++;
			}
			#endregion Validaciones

			#region Elimina Filas
			if (iCont >= iContGE)
			{
				foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr3 in uGrid.Rows.All)
				{
					if (iContGE <= iCont)
					{
						if (dr3.Cells["idArticulo"].Value != System.DBNull.Value)
							if (idArticulo == (int)dr3.Cells["idArticuloGE"].Value) 
							{
								if ((int)dr3.Cells["idDetCompra"].Value > 0)									
								{								
									string sSQLBorraLinea = string.Format("Delete From DetCompra Where idDetCompra = {0}", (int)dr3.Cells["idDetCompra"].Value);
									Funcion.EjecutaSQL(ds, sSQLBorraLinea, true);
									dr3.Delete();
								}
								else dr3.Delete();
								iContGE++;								
							}
					}				
				}						
			}			
			#endregion Elimina Filas
		}

		static private void EliminaMatriculaVehicular(Infragistics.Win.UltraWinGrid.UltraGrid uGrid, int idArticulo, bool bNuevo, C1.Data.C1DataSet ds)
		{
			#region Validaciones
			int iCont = 0;
			int iContGE = 0;

			foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr1 in uGrid.Rows.All)
			{
				if (dr1.Cells["idArticulo"].Value != System.DBNull.Value)
					if (idArticulo == (int)dr1.Cells["idArticulo"].Value) iCont++;
			}

			foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr2 in uGrid.Rows.All)
			{
				if (dr2.Cells["idArticulo"].Value != System.DBNull.Value)
					if (idArticulo == (int)dr2.Cells["idArticuloGE"].Value) iContGE++;
			}
			#endregion Validaciones

			#region Elimina Filas
			if (iCont >= iContGE)
			{
				foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr3 in uGrid.Rows.All)
				{
					if (iContGE <= iCont)
					{
						if (dr3.Cells["idArticulo"].Value != System.DBNull.Value)
							if (idArticulo == (int)dr3.Cells["idArticuloGE"].Value) 
							{
								if ((int)dr3.Cells["idDetCompra"].Value > 0)									
								{								
									string sSQLBorraLinea = string.Format("Delete From DetCompra Where idDetCompra = {0}", (int)dr3.Cells["idDetCompra"].Value);
									Funcion.EjecutaSQL(ds, sSQLBorraLinea, true);
									dr3.Delete();
								}
								else dr3.Delete();
								iContGE++;								
							}
					}				
				}						
			}			
			#endregion Elimina Filas

		}

		static private void EliminaPromocionPorValor(Infragistics.Win.UltraWinGrid.UltraGrid uGrid, int iBodega, int idArticulo, int idTipoPrecio, int idFormaPago, decimal dTotal, int idBanco, int idTarjeta, int idPlazo, int idPlan, bool bNuevo, int iCuotas, bool bCupon, C1.Data.C1DataSet ds)
		{		
			int idTipoPreciovalida = idTipoPrecio;
			if (idFormaPago == 1) idTipoPreciovalida = 2;
			if (idFormaPago == 11) idTipoPreciovalida = 90;

			int idPromocionValor = Funcion.iEscalarSQL(ds, string.Format("Select dbo.ValidaPromocionPorValor ({0}, {1}, {2}, {3}, {4}, {5}, {6}, {7}, {8}, {9})", 
				iBodega, idArticulo, idTipoPreciovalida, dTotal, idBanco, idTarjeta, idPlan, idPlazo, iCuotas, Convert.ToInt32(bCupon)));
			
			if (idPromocionValor > 0)
			{
				decimal dSubtotalLineal = 0.00m;
				decimal dSubtotal = 0.00m;
				decimal dTotalElim = 0.00m;

				foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in uGrid.Rows.All)
				{
					if (dr.Cells["idArticulo"].Value != System.DBNull.Value)
					{		
						if ((int)dr.Cells["idCombo"].Value != idPromocionValor) //if ((int)dr.Cells["idArticulo"].Value != idArticulo) 
						{
							dSubtotalLineal = (Convert.ToDecimal(dr.Cells["Precio"].Value) * (int)dr.Cells["Cantidad"].Value) - Convert.ToDecimal(dr.Cells["DescuentoArt"].Value);
							dSubtotal = dSubtotal + dSubtotalLineal;
						}

						dTotalElim = dSubtotal * 1.14m;

						int iTipo = 0;
						if (dr.Cells["Tipo"].Value != System.DBNull.Value) 
							if (dr.Cells["Tipo"].Value.ToString().Length > 0) 
								iTipo = (int)dr.Cells["Tipo"].Value;

						if (iTipo == 3)
						{
							string sSQLValElim = string.Format("Exec ValidaEliminaArticuloPromocionPorValor {0}, {1}, {2}", idPromocionValor, (int)dr.Cells["idArticulo"].Value, dTotalElim);
							if (Funcion.iEscalarSQL(ds, sSQLValElim) > 0)
							{
								if ((int)dr.Cells["idDetCompra"].Value > 0)									
								{
									string sSQLBorraLinea = string.Format("Delete From DetCompra Where idDetCompra = {0}", (int)dr.Cells["idDetCompra"].Value);
									Funcion.EjecutaSQL(ds, sSQLBorraLinea, true);
									dr.Delete();
								}	
								else dr.Delete();
							}
						}
					}
				}
			}
		}

		static private void EliminaPromocionAdicional(Infragistics.Win.UltraWinGrid.UltraGrid uGrid, int iBodega, int idArticulo, int idTipoPrecio, int idFormaPago, int idBanco, int idTarjeta, int idPlazo, int idPlan, bool bNuevo, int iCuotas, bool bCupon, C1.Data.C1DataSet ds)
		{
			int idTipoPreciovalida = idTipoPrecio;
			if (idFormaPago == 1) idTipoPreciovalida = 2;
			if (idFormaPago == 11) idTipoPreciovalida = 90;

			int idPromocionAdicional = Funcion.iEscalarSQL(ds, string.Format("Select dbo.ValidaPromocionAdicional ({0}, {1}, {2}, {3}, {4}, {5}, {6}, {7}, {8})", 
				iBodega, idArticulo, idTipoPreciovalida, idBanco, idTarjeta, idPlan, idPlazo, iCuotas, Convert.ToInt32(bCupon)));

			if (idPromocionAdicional > 0)
			{
				foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in uGrid.Rows.All)
				{
					if (dr.Cells["idArticulo"].Value != System.DBNull.Value)
					{											
						int iTipo = 0;
						if (dr.Cells["Tipo"].Value != System.DBNull.Value) 
							if (dr.Cells["Tipo"].Value.ToString().Length > 0) 
								iTipo = (int)dr.Cells["Tipo"].Value;
													
						if (iTipo == 4)
						{
							if (Funcion.iEscalarSQL(ds, string.Format("Exec ValidaEliminaArticuloPromocionAdicional {0}, {1}", idPromocionAdicional, (int)dr.Cells["idArticulo"].Value)) > 0)
							{
								if ((int)dr.Cells["idDetCompra"].Value > 0)									
								{						
									string sSQLBorraLinea = string.Format("Delete From DetCompra Where idDetCompra = {0}", (int)dr.Cells["idDetCompra"].Value);
									Funcion.EjecutaSQL(ds, sSQLBorraLinea, true);
									dr.Delete();
								}
								else dr.Delete();
							}
						}																							
					}									
				}
			}
		}

		static private void EliminaPromocionOpcional(Infragistics.Win.UltraWinGrid.UltraGrid uGrid, int iBodega, int idArticulo, int idTipoPrecio, int idFormaPago, int idBanco, int idTarjeta, int idPlazo, int idPlan, bool bNuevo, int iCuotas, bool bCupon, C1.Data.C1DataSet ds)
		{
			#region Promocion
			int idTipoPreciovalida = idTipoPrecio;
			if (idFormaPago == 1) idTipoPreciovalida = 2;
			if (idFormaPago == 11) idTipoPreciovalida = 90;

			int idPromocionOpcional = Funcion.iEscalarSQL(ds, string.Format("Select dbo.ValidaPromocionOpcional ({0}, {1}, {2}, {3}, {4}, {5}, {6}, {7}, {8})", 
				iBodega, idArticulo, idTipoPreciovalida, idBanco, idTarjeta, idPlan, idPlazo, iCuotas, Convert.ToInt32(bCupon)));

			if (idPromocionOpcional > 0)
			{
				foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in uGrid.Rows.All)
				{
					if (dr.Cells["idArticulo"].Value != System.DBNull.Value)
					{											
						int iTipo = 0;
						if (dr.Cells["Tipo"].Value != System.DBNull.Value) 
							if (dr.Cells["Tipo"].Value.ToString().Length > 0) 
								iTipo = (int)dr.Cells["Tipo"].Value;
													
						if (iTipo == 6)
						{
							if (Funcion.iEscalarSQL(ds, string.Format("Exec ValidaEliminaArticuloPromocionAdicional {0}, {1}", idPromocionOpcional, (int)dr.Cells["idArticulo"].Value)) > 0)
							{
								if ((int)dr.Cells["idDetCompra"].Value > 0)									
								{						
									string sSQLBorraLinea = string.Format("Delete From DetCompra Where idDetCompra = {0}", (int)dr.Cells["idDetCompra"].Value);
									Funcion.EjecutaSQL(ds, sSQLBorraLinea, true);
									dr.Delete();
								}
								else dr.Delete();
							}
						}																							
					}									
				}
			}
			#endregion Promocion

			EliminaCombos(uGrid, idArticulo, 7, true, ds);
		}

		static private void EliminaCombos(Infragistics.Win.UltraWinGrid.UltraGrid uGrid, int idArticulo, int iTipoE, bool bOpcional, C1.Data.C1DataSet ds)
		{
			#region Combo	
			int idCombo = 0;
			int idMarca = Funcion.iEscalarSQL(ds, string.Format("Select ISNULL(idMarca, 0) From Articulo Where idArticulo = {0}", idArticulo));
			SqlDataReader drValidaCombo = Funcion.rEscalarSQL(ds, string.Format("Exec ComboPorSubGrupo {0}, {1}, {2}", idArticulo, bOpcional, idMarca), true);
			drValidaCombo.Read();
			if (drValidaCombo.HasRows)
			{
				idCombo = drValidaCombo.GetInt32(0);				
			}
			drValidaCombo.Close();
			
			if (idCombo > 0)
			{
				foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in uGrid.Rows.All)
				{
					if (dr.Cells["idArticulo"].Value != System.DBNull.Value)
					{											
						int iTipo = 0;
						if (dr.Cells["Tipo"].Value != System.DBNull.Value) 
							if (dr.Cells["Tipo"].Value.ToString().Length > 0) 
								iTipo = (int)dr.Cells["Tipo"].Value;
													
						if (iTipo == iTipoE)
						{
							string sSQLVal = string.Format("Exec ValidaEliminaArticuloComboOpcional {0}, {1}", idCombo, (int)dr.Cells["idArticulo"].Value);
							if (Funcion.iEscalarSQL(ds, sSQLVal) > 0)
							{
								if ((int)dr.Cells["idDetCompra"].Value > 0)									
								{						
									string sSQLBorraLinea = string.Format("Delete From DetCompra Where idDetCompra = {0}", (int)dr.Cells["idDetCompra"].Value);
									Funcion.EjecutaSQL(ds, sSQLBorraLinea, true);
									dr.Delete();
								}
								else dr.Delete();
							}
						}																							
					}									
				}
			}
			#endregion Combo
		}

		static public bool FacturacionEliminaFilas(Infragistics.Win.UltraWinGrid.UltraGrid uGrid, Infragistics.Win.UltraWinGrid.UltraGrid uGridSeriales, Infragistics.Win.UltraWinGrid.BeforeRowsDeletedEventArgs e, 
			System.Windows.Forms.Button oButton, 
			int idTipoFactura, bool bNuevo, bool bImpreso, bool BEliArtCmbProm, bool BEliArt, int iBodega, int idTipoPrecio, int idFormaPago, decimal dTotal, int idBanco, int idTarjeta, int idPlazo, int idPlan, int iCuotas,
			bool bCupon, bool bRemates, C1.Data.C1DataSet ds)
		{
			#region Validacion
			if (bImpreso)
			{
				MessageBox.Show("No puede eliminar filas de documentos impresos", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				return false;
			}
				
			if (!(bool)uGrid.ActiveRow.Cells["Principal"].Value && !BEliArtCmbProm)
			{
				MessageBox.Show("No puede Eliminar el articulo principal de un combo o promoción", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				return false;
			}
				
			if ((int)uGrid.ActiveRow.Cells["Tipo"].Value != 0 && (int)uGrid.ActiveRow.Cells["Tipo"].Value != 5 && !BEliArtCmbProm)
			{
				MessageBox.Show("No puede eliminar un producto de un combo o promoción\n\n Consulte con el Administrador", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				return false;
			}	
				
			if ((bool)uGrid.ActiveRow.Cells["Bloqueado"].Value)
			{
				MessageBox.Show("Este Articulo no puede ser Eliminado", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				return false;
			}
			
			if ((int)uGrid.ActiveRow.Cells["Tipo"].Value == 8)
			{
				MessageBox.Show("El seguro D&FI no puede ser Eliminado", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				return false;
			}

			if ((int)uGrid.ActiveRow.Cells["Tipo"].Value == 11)
			{
				MessageBox.Show("La Matricula Vehicular no puede ser eliminado", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				return false;
			}
			#endregion Validacion				

			bool bRes = false;

			if (DialogResult.Yes ==	MessageBox.Show("¿Desea Borrar las Líneas Seleccionadas?", "Point Technology", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2)) 
			{
				int iTipo = (int)uGrid.ActiveRow.Cells["Tipo"].Value;
				bool bPrincipal = (bool)uGrid.ActiveRow.Cells["Principal"].Value;

				if (iTipo == 0 || iTipo == 3 || iTipo == 4 || iTipo == 5 || iTipo == 11 )
				{					
					if (iTipo == 0)
					{
						EliminaPromocionPorValor(uGrid, iBodega, (int)uGrid.ActiveRow.Cells["idArticulo"].Value, idTipoPrecio, idFormaPago, dTotal, idBanco, idTarjeta, idPlazo, idPlan, bNuevo, iCuotas, bCupon, ds);
						EliminaPromocionAdicional(uGrid, iBodega, (int)uGrid.ActiveRow.Cells["idArticulo"].Value, idTipoPrecio, idFormaPago, idBanco, idTarjeta, idPlazo, idPlan, bNuevo, iCuotas, bCupon, ds);
						EliminaPromocionOpcional(uGrid, iBodega, (int)uGrid.ActiveRow.Cells["idArticulo"].Value, idTipoPrecio, idFormaPago, idBanco, idTarjeta, idPlazo, idPlan, bNuevo, iCuotas, bCupon, ds);
						EliminaCombos(uGrid, (int)uGrid.ActiveRow.Cells["idArticulo"].Value, 1, false, ds);
						
						if (uGrid.ActiveRow.Cells["idArticulo"].Value != System.DBNull.Value)
							if (uGrid.ActiveRow.Cells["idArticulo"].Value.ToString().Length > 0)
								EliminaCombos(uGrid, (int)uGrid.ActiveRow.Cells["idArticulo"].Value, 2, false, ds);

						if (uGrid.ActiveRow.Cells["idArticulo"].Value != System.DBNull.Value)
							if (uGrid.ActiveRow.Cells["idArticulo"].Value.ToString().Length > 0)
								EliminaCombos(uGrid, (int)uGrid.ActiveRow.Cells["idArticulo"].Value, 7, true, ds);
						
						if (bRemates) FuncionesInventario.EliminaSerialesRemates(e, uGridSeriales, idTipoFactura, (int)uGrid.ActiveRow.Cells["idCompra"].Value, 0, ds);

						if (iTipo == 11) EliminaMatriculaVehicular(uGrid, (int)uGrid.ActiveRow.Cells["idArticulo"].Value, bNuevo, ds);
					}

					//if (Funcion.bEscalarSQL(ds, string.Format("Select GarantiaExtendida From Articulo Where idArticulo = {0}", (int)uGrid.ActiveRow.Cells["idArticulo"].Value), true))
						EliminaGarantiaExtendida(uGrid, (int)uGrid.ActiveRow.Cells["idArticulo"].Value, bNuevo, ds);

						for (int i=0; i<e.Rows.Length; i++)
						{
							if ((int)e.Rows[i].Cells["idDetCompra"].Value > 0) 
							{
								oButton.Enabled = false;
								string sSQLBorraLinea = string.Format("Delete From DetCompra Where idDetCompra = {0}", (int)e.Rows[i].Cells["idDetCompra"].Value);
								Funcion.EjecutaSQL(ds, sSQLBorraLinea, true);	
							}
						}

					bRes = true;
				}	
				else if ((iTipo == 1 || iTipo == 2) && !bPrincipal)
				{					
					EliminaPromocionPorValor(uGrid, iBodega, (int)uGrid.ActiveRow.Cells["idArticulo"].Value, idTipoPrecio, idFormaPago, dTotal, idBanco, idTarjeta, idPlazo, idPlan, bNuevo, iCuotas, bCupon, ds);
					EliminaPromocionAdicional(uGrid, iBodega, (int)uGrid.ActiveRow.Cells["idArticulo"].Value, idTipoPrecio, idFormaPago, idBanco, idTarjeta, idPlazo, idPlan, bNuevo, iCuotas, bCupon, ds);
					EliminaPromocionOpcional(uGrid, iBodega, (int)uGrid.ActiveRow.Cells["idArticulo"].Value, idTipoPrecio, idFormaPago, idBanco, idTarjeta, idPlazo, idPlan, bNuevo, iCuotas, bCupon, ds);
					EliminaCombos(uGrid, (int)uGrid.ActiveRow.Cells["idArticulo"].Value, 1, false, ds);
					
					if (uGrid.ActiveRow.Cells["idArticulo"].Value != System.DBNull.Value)
						if (uGrid.ActiveRow.Cells["idArticulo"].Value.ToString().Length > 0)
							EliminaCombos(uGrid, (int)uGrid.ActiveRow.Cells["idArticulo"].Value, 2, false, ds);

					if (uGrid.ActiveRow.Cells["idArticulo"].Value != System.DBNull.Value)
						if (uGrid.ActiveRow.Cells["idArticulo"].Value.ToString().Length > 0)
							EliminaCombos(uGrid, (int)uGrid.ActiveRow.Cells["idArticulo"].Value, 7, true, ds);

					if (bRemates) FuncionesInventario.EliminaSerialesRemates(e, uGridSeriales, idTipoFactura, (int)uGrid.ActiveRow.Cells["idCompra"].Value, 0, ds);

					//if (Funcion.bEscalarSQL(ds, string.Format("Select GarantiaExtendida From Articulo Where idArticulo = {0}", (int)uGrid.ActiveRow.Cells["idArticulo"].Value), true))
						EliminaGarantiaExtendida(uGrid, (int)uGrid.ActiveRow.Cells["idArticulo"].Value, bNuevo, ds);

						for (int i=0; i<e.Rows.Length; i++)
						{								
							if ((int)e.Rows[i].Cells["idDetCompra"].Value > 0) 
							{
								oButton.Enabled = false;
								string sSQLBorraLinea = string.Format("Delete From DetCompra Where idDetCompra = {0}", (int)e.Rows[i].Cells["idDetCompra"].Value);
								Funcion.EjecutaSQL(ds, sSQLBorraLinea, true);	
							}
						}

					bRes = true;
				}
				else if ((iTipo == 1 || iTipo == 2) && bPrincipal)
				{
					int idCombo = (int)uGrid.ActiveRow.Cells["idCombo"].Value;
					int idArticulo = (int)uGrid.ActiveRow.Cells["idArticulo"].Value;

					foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in uGrid.Rows.All)
					{
						if (dr.Cells["idCombo"].Value.ToString().Length > 0)
						{							
							if ((int)dr.Cells["idCombo"].Value == idCombo)
							{									
								EliminaPromocionPorValor(uGrid, iBodega, (int)dr.Cells["idArticulo"].Value, idTipoPrecio, idFormaPago, dTotal, idBanco, idTarjeta, idPlazo, idPlan, bNuevo, iCuotas, bCupon, ds);
								EliminaPromocionAdicional(uGrid, iBodega, (int)dr.Cells["idArticulo"].Value, idTipoPrecio, idFormaPago, idBanco, idTarjeta, idPlazo, idPlan, bNuevo, iCuotas, bCupon, ds);								
								EliminaPromocionOpcional(uGrid, iBodega, (int)dr.Cells["idArticulo"].Value, idTipoPrecio, idFormaPago, idBanco, idTarjeta, idPlazo, idPlan, bNuevo, iCuotas, bCupon, ds);
								EliminaCombos(uGrid, (int)dr.Cells["idArticulo"].Value, 1, false, ds);

								if (dr.Cells["idArticulo"].Value != System.DBNull.Value)
									if (dr.Cells["idArticulo"].Value.ToString().Length > 0)
										EliminaCombos(uGrid, (int)dr.Cells["idArticulo"].Value, 2, false, ds);

								if (dr.Cells["idArticulo"].Value != System.DBNull.Value)
									if (dr.Cells["idArticulo"].Value.ToString().Length > 0)
										EliminaCombos(uGrid, (int)dr.Cells["idArticulo"].Value, 7, true, ds);

								if (bRemates) FuncionesInventario.EliminaSerialesRemates(e, uGridSeriales, idTipoFactura, (int)uGrid.ActiveRow.Cells["idCompra"].Value, 0, ds);

//								if (Funcion.bEscalarSQL(ds, string.Format("Select GarantiaExtendida From Articulo Where idArticulo = {0}", (int)dr.Cells["idArticulo"].Value), true))
//									EliminaGarantiaExtendida(uGrid, (int)dr.Cells["idArticulo"].Value, bNuevo, ds);

								//if (Funcion.bEscalarSQL(ds, string.Format("Select GarantiaExtendida From Articulo Where idArticulo = {0}", idArticulo), true))
									EliminaGarantiaExtendida(uGrid, idArticulo, bNuevo, ds);

									int getXidDetCompra = 0;
									try{ getXidDetCompra = (int)dr.Cells["idDetCompra"].Value;} 
									catch(Exception ex){string sms = ex.Message;}
									if(getXidDetCompra > 0)
									{
										oButton.Enabled = false;
										string sSQLBorraLinea = string.Format("Delete From DetCompra Where idDetCompra = {0}", (int)dr.Cells["idDetCompra"].Value);
										Funcion.EjecutaSQL(ds, sSQLBorraLinea, true);
										dr.Delete();
									}
									else dr.Delete();						
							}
						}
					}

					bRes = false;
				}
			}
			else bRes = false;					

			return bRes;
		}
		
		static public decimal dPorcentajeIVAVigente(bool bFiscal, bool bValidaOrigen, int idOrigen, DateTime dtFecha, C1.Data.C1DataSet ds)
		{
			decimal dPorcentaje = 0.00m;

			int iFiscal = 0;
			if (bFiscal) iFiscal = 1;

			if (bValidaOrigen && idOrigen > 0)
				dtFecha = Funcion.dtEscalarSQL(ds, string.Format("Select Fecha, Iva From Compra Where idCompra = {0}", idOrigen));
					
			dPorcentaje = Funcion.decEscalarSQL(ds, string.Format("Select dbo.RetornaImpuestoIVAVigente('{0}', {1})", 
				dtFecha.ToString("yyyyMMdd"), iFiscal));

			return dPorcentaje;
		}

		static public int GrabaFacturacion (
			int idCompra, int idOrigen, int idSucursal, int idTipoFactura, int idComprobante, int iEstado, bool bConsignacion,
			int iContadoCredito, int idFormaPago, int iBodega, int idProyecto, int idVendedor, string sNumero, string sComprobante, 
			int idTipoRuc, int idCliente, string sRUC, string sNombre, DateTime Fecha, string sSerie, string sAutorizacion, DateTime FechaCaducidad,
			bool bEntregaDom, DateTime FechaEntrega, DateTime HoraEntrega, int idDireccion,	string sNotas,
			bool bCrediPoint, decimal dTotalCuotas, int iNumCuotas, decimal dCuotaEntrada, DateTime FechaPrimerPago, 
			decimal dSubTotalExcento, decimal dDescuento0, decimal dSubTotalIva, decimal dDescuento, decimal dIva, decimal dTotal, 
			bool bFecPrimPag, bool bFecCad, bool bFecEnt, bool bHorEnt, bool bBorrar,
			bool bFiscal, int iBodegaDestino, int idMotivo, string sObservacion, decimal dSaldoNotaDeCredito, 
			decimal dInteres, decimal dCuotaAsignada, decimal dCuotaDisponible, decimal dCuotaUsada, 
			bool bAutorizaEntrada, decimal dValorCuota, int idEntidadFinanciera, 
			int idBanco, int idTarjeta, int idPlan, int idPlazo, string sNumeroDeCredito, bool bVerificador,
			int iTipoVenta, string sNumTransaccion, string sDireccionEntrega, int idCourier, string sNumGuia, decimal dValorEntrega, int idProyectoZOOM, string sClienteNombre, decimal dJornadas,
			bool bNuevo, bool bNumeracionAutomatica, int idBloqueaTransacciones, Infragistics.Win.UltraWinGrid.UltraGrid uGrid,
			bool bSeriales, int idTipoOrigen, Infragistics.Win.UltraWinGrid.UltraGrid uGridSeriales, bool bDescuentoActivo)
		{			
			using (SqlConnection oConexion = new SqlConnection("Password=Bl45o6$9;Persist Security Info=True;User ID= " + MenuLatinium.stUsuario + ";Initial Catalog=" + MenuLatinium.stDirFacturacion + ";Data Source=" + MenuLatinium.stDirServidor + ";Max Pool Size = 1024;"))
			{
				oConexion.Open();

				SqlTransaction oTransaction = null;
				SqlCommand oCmdActualiza = oConexion.CreateCommand();
				oCmdActualiza.Connection = oConexion;
				oCmdActualiza.CommandType = CommandType.Text;
				oCmdActualiza.CommandTimeout = 120;
						
				try
				{							
					oTransaction = oConexion.BeginTransaction();
					oCmdActualiza.Transaction = oTransaction;

					#region Numeracion
					if (bNuevo && bNumeracionAutomatica && iEstado != 6)
					{
						string sSQLNumero = string.Format("Exec NumeracionLocales {0}, {1}, {2}", idTipoFactura, iBodega, bFiscal);
						oCmdActualiza.CommandText = sSQLNumero;
						sNumero = oCmdActualiza.ExecuteScalar().ToString();
					}
					#endregion Numeracion
					
					#region Maestro
					string sSQL = string.Format("Exec GrabaMaestroFactura {0}, {1}, {2}, {3}, {4}, {5}, {6}, {7}, {8}, {9}, {10}, {11}, '{12}', '{13}', {14}, {15}, '{16}', '{17}', '{18}', '{19}', '{20}', '{21}', {22}, '{23}', '{24}', {25}, '{26}', {27}, {28}, {29}, {30}, '{31}', {32}, {33}, {34}, {35}, {36}, {37}, {38}, {39}, {40}, {41}, {42}, {43}, {44}, {45}, '{46}', {47}, {48}, {49}, {50}, {51}, {52}, {53}, {54}, {55}, {56}, {57}, {58}, '{59}', {60}, {61}, '{62}', '{63}', {64}, '{65}', {66}, {67}, '{68}', {69}", 
						idCompra, idOrigen, idSucursal, idTipoFactura, idComprobante, iEstado, bConsignacion,
						iContadoCredito, idFormaPago, iBodega, idProyecto, idVendedor, sNumero, sComprobante, 
						idTipoRuc, idCliente, sRUC, sNombre, Fecha.ToString("yyyyMMdd"), sSerie, sAutorizacion, FechaCaducidad.ToString("yyyyMMdd"),
						bEntregaDom, FechaEntrega.ToString("yyyyMMdd"), HoraEntrega.ToString("yyyyMMdd HH:mm"), idDireccion, sNotas,
						bCrediPoint, dTotalCuotas, iNumCuotas, dCuotaEntrada, FechaPrimerPago.ToString("yyyyMMdd"), 
						dSubTotalExcento, dDescuento0, dSubTotalIva, dDescuento, dIva, dTotal, 
						bFecPrimPag, bFecCad, bFecEnt, bHorEnt, bBorrar,
						bFiscal, iBodegaDestino, idMotivo, sObservacion, dSaldoNotaDeCredito, 
						dInteres, dCuotaAsignada, dCuotaDisponible, dCuotaUsada, 
						bAutorizaEntrada, dValorCuota, idEntidadFinanciera, 
						idBanco, idTarjeta, idPlan, idPlazo, sNumeroDeCredito, bVerificador,
						iTipoVenta, sNumTransaccion, sDireccionEntrega, idCourier, sNumGuia, dValorEntrega, idProyectoZOOM, sClienteNombre, dJornadas);
					oCmdActualiza.CommandText = sSQL;
					idCompra = (int)oCmdActualiza.ExecuteScalar();
					#endregion Maestro

					if (!bBorrar)
					{	
						#region Detalle
						foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in uGrid.Rows.All)
						{							
							string sSQLDetalle = string.Format("Exec GuardaDetalleFactura {0}, {1}, {2}, {3}, {4}, '{5}', '{6}', {7}, {8}, {9}, {10}, {11}, {12}, {13}, {14}, {15}, {16}, {17}, {18}, {19}, {20}, {21}",
								(int)dr.Cells["idDetCompra"].Value, (int)dr.Cells["idOrigen"].Value, idCompra,
								(int)dr.Cells["idArticulo"].Value, (int)dr.Cells["idUnidad"].Value, dr.Cells["Codigo"].Value.ToString(), dr.Cells["Articulo"].Value,
								(int)dr.Cells["Cantidad"].Value, (decimal)dr.Cells["Precio"].Value, (decimal)dr.Cells["Impuesto"].Value, 
								(decimal)dr.Cells["DescuentoArt"].Value, (decimal)dr.Cells["DescuentoPorc"].Value, (decimal)dr.Cells["Subtotal"].Value,
								(bool)dr.Cells["EstadoSerial"].Value, (int)dr.Cells["Tipo"].Value, (bool)dr.Cells["Entregado"].Value, (int)dr.Cells["Posicion"].Value, (bool)dr.Cells["Principal"].Value, 
								(bool)dr.Cells["Bloqueado"].Value, (int)dr.Cells["idCombo"].Value, (int)dr.Cells["idArticuloGE"].Value, (decimal)dr.Cells["Jornadas"].Value);
							oCmdActualiza.CommandText = sSQLDetalle;
							oCmdActualiza.ExecuteNonQuery();
						}
						#endregion Detalle	
						
						string sSQLActDetalle = string.Format("Exec ConsultaDetalleFactura {0}", idCompra);
						oCmdActualiza.CommandText = sSQLActDetalle;
						SqlDataAdapter oDADetalle = new SqlDataAdapter(oCmdActualiza);
						DataTable oDTDetalle = new DataTable();
						oDADetalle.Fill(oDTDetalle);						
						uGrid.DataSource = oDTDetalle;						

						#region Seriales
						#region Notas de Credito
						if (idTipoFactura == 5)
						{
							if (idOrigen > 0)
							{
								foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in uGrid.Rows.All)
								{
									string sSQLMovSerial = string.Format("Exec MovimientoSerial {0}, {1}, {2}, {3}", 
										(int)dr.Cells["idOrigen"].Value, (int)dr.Cells["idDetCompra"].Value, (int)dr.Cells["idArticulo"].Value, iBodegaDestino);
									oCmdActualiza.CommandText = sSQLMovSerial;
									oCmdActualiza.ExecuteNonQuery();
								}
							}
						}
						#endregion Notas de Credito
						#region Facturas
						if (idTipoFactura == 1 || idTipoFactura == 43)
						{
							if (bSeriales)
							{	
								#region Seriales desde Pedido - Reserva
								if (idTipoOrigen == 13 || idTipoOrigen == 11)
								{							
									if (bNuevo)
									{								
										foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in uGridSeriales.Rows.All)
										{
											string sSQLMovSerial = string.Format("Exec MovimientoSerial {0}, {1}, {2}, 0, {3}", 
												(int)dr.Cells["idOrigen"].Value, (int)dr.Cells["idDetCompra"].Value, (int)dr.Cells["idArticulo"].Value, (int)dr.Cells["Posicion"].Value);
											oCmdActualiza.CommandText = sSQLMovSerial;
											oCmdActualiza.ExecuteNonQuery();										
										}								

										string sSQLActSeriales = string.Format("Exec RefrescaSerialVenta {0}", idCompra);
										oCmdActualiza.CommandText = sSQLActSeriales;
										SqlDataAdapter oDASeriales = new SqlDataAdapter(oCmdActualiza);
										DataTable oDTSeriales = new DataTable();
										oDASeriales.Fill(oDTSeriales);
										uGridSeriales.DataSource = oDTSeriales;
									}							
								}
								#endregion Seriales desde Pedido - Reserva

								#region Grupos de seriales
								if (idTipoOrigen == 0 || idTipoOrigen == 43)
								{
									foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in uGrid.Rows.All)
									{
										if (dr.Cells["idArticulo"].Value != DBNull.Value && (int) dr.Cells["idArticulo"].Value > 0)
										{
											int IdArticulo = (int) dr.Cells["idArticulo"].Value;
											int IdDetCompra = (int) dr.Cells["idDetCompra"].Value;										
											int IPosicion = (int) dr.Cells["Posicion"].Value; 					

											foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr1 in uGridSeriales.Rows.All)
											{
												int IdArtSerial = (int) dr1.Cells["idArticulo"].Value;
												int IEstado = (int) dr1.Cells["Estado"].Value;
											
												if (IdArticulo == IdArtSerial && IPosicion == (int) dr1.Cells["Posicion"].Value)
												{													
													if (IEstado == 4 || IEstado == 2 || IEstado == 28)
													{
														int idSerial = 0;
														int idOrigenS = 0;

														if (dr1.Cells["idSerial"].Value != DBNull.Value) idSerial = (int)dr1.Cells["idSerial"].Value;						
														if (dr1.Cells["idOrigen"].Value != DBNull.Value) idOrigenS = (int)dr1.Cells["idOrigen"].Value;

														string sSQLGDetalleSeriales = string.Format("Exec GuardaDetalleSeriales {0}, {1}, {2}, {3}, '{4}', 1, '', 0, 0, {5}, {6}", 
															idSerial, IdDetCompra, idOrigenS, iBodegaDestino, dr1.Cells["Observaciones"].Value.ToString(), (int)dr1.Cells["idDetalleSerial"].Value, (int)dr1.Cells["Posicion"].Value);
														oCmdActualiza.CommandText = sSQLGDetalleSeriales;
														oCmdActualiza.ExecuteNonQuery();
																
														string sSQLActEstSerial = string.Format("Update Seriales Set Estado = {0}, Bodega = {2} Where idSerial = {1}", 4, idSerial, iBodegaDestino);
														oCmdActualiza.CommandText = sSQLActEstSerial;																
														oCmdActualiza.ExecuteNonQuery();

														string sSQLConteo = string.Format("Exec ActualizaVendidoConteo '{0}', {1}, '{2}', {3}", 
															Fecha.ToString("yyyyMMdd"), iBodegaDestino, dr1.Cells["Serial"].Value.ToString(), (int)dr1.Cells["idArticulo"].Value);
														oCmdActualiza.CommandText = sSQLConteo;
														oCmdActualiza.ExecuteNonQuery();
													}
												}
											}
										}
									}

									string sSQLActSeriales = string.Format("Exec RefrescaSerialVenta {0}", idCompra);
									oCmdActualiza.CommandText = sSQLActSeriales;
									SqlDataAdapter oDASeriales = new SqlDataAdapter(oCmdActualiza);
									DataTable oDTSeriales = new DataTable();
									oDASeriales.Fill(oDTSeriales);
									uGridSeriales.DataSource = oDTSeriales;
								}
								#endregion Grupos de seriales
							}
						}
						#endregion Facturas
						#endregion Seriales

						#region Inventario
						if (idTipoFactura != 1)
						{	
							oCmdActualiza.CommandText = string.Format("Exec GuardaMovimientosInventario {0}, {1}", idCompra, idTipoFactura);
							oCmdActualiza.ExecuteNonQuery();
						}					
						#endregion Inventario

						#region Asiento
						string sSQLAs = string.Format("Select Asiento From CompraNumero Where idTipoFactura = {0}", idTipoFactura);
						oCmdActualiza.CommandText = sSQLAs;
						bool bAsiento = (bool)oCmdActualiza.ExecuteScalar();
						if (bAsiento && iEstado != 6)
						{
							string sSQLNA = string.Format("Select ISNULL(AsientoProcedimiento, '') From CompraNumero Where idTipoFactura = {0}", idTipoFactura);
							oCmdActualiza.CommandText = sSQLNA;
							string spNombre = oCmdActualiza.ExecuteScalar().ToString();

							string sSQLAsiento = string.Format("Exec {0} {1}", spNombre, idCompra);
							oCmdActualiza.CommandText = sSQLAsiento;
							oCmdActualiza.ExecuteNonQuery();
						}
						#endregion Asiento

						#region Autorización de Pagos Devolucion de Dinero
						if (idTipoFactura == 5 && idMotivo == 6)
						{	
							string sSQLAP = string.Format("Exec CreaAutorizacionesDePago 14, {0}, '{1}', '{2}', {3}, 2, '1 DE 1', 'DEVOLUCION DE DINERO CON NOTA DE CREDITO', 0", 
								idCompra, DateTime.Today.ToString("yyyyMMdd"), DateTime.Today.ToString("yyyyMMdd"), Math.Round(dTotal, 2));
							oCmdActualiza.CommandText = sSQLAP;
							oCmdActualiza.ExecuteNonQuery();
						}
						else
						{
							if (!bNuevo) 
							{
								string sSQLEAP = string.Format("If (Select COUNT(*) From AutorizacionDePagos Where idOrigen = {0} And idTipo = 14) > 0 Delete From AutorizacionDePagos Where idOrigen = {0} And idTipo = 14", idCompra);
								oCmdActualiza.CommandText = sSQLEAP;
								oCmdActualiza.ExecuteNonQuery();
							}
						}
						#endregion Autorización de Pagos Devolucion de Dinero

						#region Bloqueos
						if (idBloqueaTransacciones > 0)
						{
							string sSQLBloqueo = string.Format("Exec RegistraBloqueoTransacciones {0}, 0, {1}, 1", idBloqueaTransacciones, idCompra);
							oCmdActualiza.CommandText = sSQLBloqueo;
							oCmdActualiza.ExecuteNonQuery();
						}

						if (bNuevo && idOrigen > 0)
						{
							string sSQLBloqueoOrigen = string.Format("Update BloqueoTransacciones Set Bloqueado = 0, FechaDesbloqueo = GETDATE() Where idOrigen = {0} And Origen = 1", idOrigen);
							oCmdActualiza.CommandText = sSQLBloqueoOrigen;
							oCmdActualiza.ExecuteNonQuery();
						}
						#endregion Bloqueos

						#region Audita Descuento
						if (bDescuentoActivo)
						{
							string sMensaje = "";
											
							int iContador = 1;
						
							foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in uGrid.Rows.All)
							{
								if (iContador > 1) sMensaje = " - ";
						
								sMensaje = sMensaje + " Codigo : " + dr.Cells["Codigo"].Value.ToString() + " % Dscto : " + dr.Cells["DescuentoPorc"].Value.ToString();
						
								iContador++;
							}
						
							string sDocumento = "FACTURA";
							if (idTipoFactura == 13) sDocumento = "PEDIDO";
							if (idTipoFactura == 43) sDocumento = "PREFACTURA";
							if (idTipoFactura == 44) sDocumento = "CONSIGNACION";
						
							string sSQLCD = string.Format("Exec RegistraAuditoriaClaveDesCuento {0}, '{1}', {2}, 3, '{3}'", MenuLatinium.IdUsuario, sDocumento, idCompra, sMensaje);
							oCmdActualiza.CommandText = sSQLCD;
							oCmdActualiza.ExecuteNonQuery();						
						}
						#endregion Audita Descuento
					}
	
					oTransaction.Commit();

					oConexion.Close();
					
					return idCompra;
				}
				catch (Exception ex)
				{
					try
					{
						oTransaction.Rollback();

						MessageBox.Show(string.Format("Commit Exception Type : {0} {1}", ex.GetType(), ex.Message), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);						

						return 0;
					}
					catch (Exception ex2)
					{
						MessageBox.Show(string.Format("Rollback Exception Type: {0} {1}", ex2.GetType(), ex2.Message), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);													

						return 0;
					}
				}
				finally
				{
					oConexion.Close();					
				}

				return idCompra;
			}			
		}	

		static public string Modulo11(string number)
		{
			int pivote = 2;
			int longitudCadena = number.Length;
			int cantidadTotal = 0;
			int b = 1;

			for (int i = 0; i < longitudCadena; i++)
			{
				if (pivote == 8) pivote = 2;

				int temporal = int.Parse(number.Substring(i, b));

				temporal *= pivote;

				pivote++;

				cantidadTotal += temporal;
			}

			cantidadTotal = 11 - cantidadTotal % 11;

			if (cantidadTotal == 11) cantidadTotal = 0;
			else if (cantidadTotal == 10) cantidadTotal = 1;

			return cantidadTotal.ToString();
		}

		static public string invertirCadena(String cadena)
		{
			String cadenaInvertida = "";
			for (int x = cadena.Length - 1; x >= 0; x--)
				cadenaInvertida = cadenaInvertida + cadena.Substring(x, 1);

			return cadenaInvertida;
		}

		static public void ClaveAcceso(int idDocumento, int idTipoFactura)
		{
//			oCmdActualiza.CommandText = string.Format("Exec FE_ClaveDeAcceso {0}, {1}", idDocumento, idTipoFactura);;
//			string sClaveAcceso = oCmdActualiza.ExecuteScalar().ToString();
//
//			string cadInvertida = Facturacion.invertirCadena(sClaveAcceso.Substring(0, 48));
//			string ValM11 = Facturacion.Modulo11(cadInvertida).ToString().Trim();
//			string sClaveSRI = sClaveAcceso.ToString().Substring(0, 48) + ValM11;
//
//			oCmdActualiza.CommandText = string.Format("Update Compra Set AutFactura = '{0}' Where idCompra = {1}", sClaveSRI, idDocumento);
//			oCmdActualiza.ExecuteNonQuery();
		}
	}
}
