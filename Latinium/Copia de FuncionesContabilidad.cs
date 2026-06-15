using System;
using System.Data;
using System.Data.SqlClient;
using System.Windows.Forms;
using Infragistics;
using Infragistics.Win.UltraWinGrid;
using System.Globalization;
using System.Data.OleDb;
using System.Data.Odbc;
using CrystalDecisions.Shared;

namespace Latinium
{
	/// <summary>
	/// Descripción breve de FuncionesContabilidad.
	/// </summary>
	public class FuncionesContabilidad
	{
		public FuncionesContabilidad()
		{
			//
			// TODO: agregar aquí la lógica del constructor
			//
		}
		
		static public void ImprimeAsiento(int idAsiento)
		{			
			string strFiltro = "{Asiento.idAsiento} = " + idAsiento.ToString();
			Reporte Reporte = new Reporte("Asientos.rpt", strFiltro);			
			Reporte.Show();
		}

		static public void AsientoDocumentosCompra(int idAsiento, int idCompra, int idTipoFactura, bool bCrearAsiento, bool bAnulado, DateTime dtFecha, C1.Data.C1DataSet ds)
		{
			#region Valida documento crea Asiento
			string sSQL = string.Format("Select Asiento, ISNULL(AsientoProcedimiento, '') From CompraNumero Where idTipoFactura = {0}", idTipoFactura);
			SqlDataReader dr = Funcion.rEscalarSQL(ds, sSQL, true);
			dr.Read();
			bool bAsiento = dr.GetBoolean(0);
			string spNombre = dr.GetString(1);
			dr.Close();

			if (!bAsiento)
			{
				MessageBox.Show("Este documento no genera Asiento", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				return;
			}
			
			if (spNombre.Length == 0)
			{
				MessageBox.Show("No ha especificado el nombre del procedimiento que genera asiento para este tipo de documento\n\ncomuniquese con Sistemas", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				return;
			}
			#endregion Valida documento crea Asiento

			if (bAnulado)
			{
				MessageBox.Show("El documento esta anulado", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);						
				return;
			}

			if (!bCrearAsiento)
			{
				MessageBox.Show("No tiene acceso a generar asiento de este documento", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);						
				return;
			}

			if (Funcion.ValidaBloqueo(idCompra, 1, ds)) return;

			if (Funcion.ValidaPeriodo(dtFecha, ds, "Bodega"))
			{				
				MessageBox.Show(string.Format("No puede modificar documentos en esta fecha, El período '{0}' esta cerrado,\n\nConsulte al Administrador", dtFecha.ToString("MMMM").ToUpper()), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);									
				return;
			}

			if (idAsiento > 0 && DialogResult.No == MessageBox.Show("El Asiento de este documento ya existe\n\n\n Desea sobrescribirlo", "Point Technology", MessageBoxButtons.YesNo, MessageBoxIcon.Question)) return;
 
			string sSQLAsiento = string.Format("Exec {0} {1}", spNombre, idCompra);
			Funcion.EjecutaSQL(ds, sSQLAsiento, true);

			MessageBox.Show("Asiento generado con éxito", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Information);			
		}

		static public void ValidaVistaAsientoDesdeDocumentos(int idAsiento, bool bAnulado, bool bCrearAsiento, C1.Data.C1DataSet ds)
		{
			if (bAnulado)
			{
				MessageBox.Show("El documento esta anulado", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);						
				return;
			}

			if (!bCrearAsiento)
			{
				MessageBox.Show("No tiene acceso a ver el asiento de este documento", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);						
				return;
			}

			if (idAsiento == 0) 
			{
				MessageBox.Show("No Existe un Asiento Creado de esta Transacción", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				return;
			}
			
			if (!Funcion.bEscalarSQL(ds, string.Format("Declare @Res Bit If Exists(Select idAsiento From Asiento Where idAsiento = {0}) Set @Res = 1 Else Set @Res = 0 Select @Res", idAsiento), true))
			{
				MessageBox.Show("El Asiento de este Ingreso fue Eliminado", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				return;
			}

			if (Funcion.bEscalarSQL(ds, string.Format("Declare @Res Bit If Exists(Select idAsiento From Asiento Where idAsiento = {0} And Borrar = 1) Set @Res = 1 Else Set @Res = 0 Select @Res", idAsiento), true))
			{
				MessageBox.Show("El Asiento de este Documento esta Anulado", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				return;
			}
			
			frmAsientos Asiento = new frmAsientos(idAsiento);			
			Asiento.Show();			
		}

		static public bool vbValidaDocumentoBancario(int idCuenta, Infragistics.Win.UltraWinEditors.UltraTextEditor uTXT, int idTipo, int idAsiento, C1.Data.C1DataSet ds)
		{
			if (Funcion.bEscalarSQL(ds, string.Format("Exec ValidaNumeroDocumento {0}, '{1}', {2}, {3}", idCuenta, uTXT.Text.ToString(), idTipo, idAsiento), true))
			{
				MessageBox.Show("El Número de Documento ya existe", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				uTXT.Focus();
				return false;
			}

			return true;
		}

		static public bool ValidaAsientoEnConciliacion(int idAsiento, DateTime dtFecha, C1.Data.C1DataSet ds)
		{
			if (Funcion.iEscalarSQL(ds, string.Format("Exec ValidaAsientoEnConciliacion {0}", idAsiento)) > 0)
			{
				MessageBox.Show(string.Format("El Documento esta registrado en la Conciliacion bancaria del periodo '{0}'", dtFecha.ToString("MM/yyyy")), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);				
				return false;
			}

			return true;
		}

		static public bool vBancosValidaEstadoConciliacion (int idCuenta, DateTime dtFecha, C1.Data.C1DataSet ds)
		{
			int iEstado = Funcion.iEscalarSQL(ds, string.Format("Exec BancosValidaEstadoConciliacion {0}, '{1}'", idCuenta, dtFecha.ToString("yyyyMMdd")));
			
			if (iEstado == 1)
			{
				MessageBox.Show("El periodo bancario de la cuenta seleccionada esta Cerrado", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				return false;
			}

			return true;
		}

		static public bool bEstadoCaja (int iBodega, DateTime dtFecha, int idCompra, int idTipoFactura, C1.Data.C1DataSet ds)
		{
			int iEstado = Funcion.iEscalarSQL(ds, string.Format("Exec ValidaEstadoCaja {0}, '{1}', {2}, {3}", 
				iBodega, dtFecha.ToString("yyyyMMdd"), idCompra, idTipoFactura));

			if (iEstado == 2)
			{
				MessageBox.Show("La caja de este documento esta cerrada\n\nConsulte con el administrador", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				return false;
			}

			return true;
		}

		static public bool bValidaEdicionAnulacionAsientos (Infragistics.Win.UltraWinGrid.UltraCombo uCmb, int iAccion, C1.Data.C1DataSet ds)
		{
			int idOrigenAsiento = 0;
			if (uCmb.ActiveRow != null) idOrigenAsiento = (int)uCmb.Value;

			bool bRes = Funcion.bEscalarSQL(ds, string.Format("Exec ValidaAsientoTB {0}, {1}", idOrigenAsiento, iAccion), true);
			
			string sAccion = "Editar";
			if (iAccion == 3) sAccion = "Anular";
			if (!bRes)
			{
				MessageBox.Show("No puede " + sAccion + " :\n\nEste registro fue creado Automaticamente desde :" + uCmb.Text.ToString(), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				return false;
			}

			return true;
		}

		static public bool bValidaAnticipoDuplicado (Infragistics.Win.UltraWinGrid.UltraCombo uCmb, Infragistics.Win.UltraWinGrid.UltraCombo uCmbPersonal, int idRolPrestamo, C1.Data.C1DataSet ds)
		{
			int iCont = Funcion.iEscalarSQL(ds, string.Format("Exec ValidaAnticipoDuplicado {0}, {1}, {2}", (int)uCmb.Value, (int)uCmbPersonal.Value, idRolPrestamo), true);
			
			if (iCont > 0)
			{
				MessageBox.Show(string.Format("'{0}' tiene {1} '{2}' sin procesar", uCmbPersonal.Text.ToString(), iCont, uCmb.Text.ToString()),  "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				uCmbPersonal.Focus();
				return false;
			}

			return true;
		}

		static public bool bValidaMovimientoBancarioConciliado (Infragistics.Win.UltraWinGrid.UltraCombo uCmb, Infragistics.Win.UltraWinGrid.UltraCombo uCmbPersonal, C1.Data.C1DataSet ds)
		{
			int iCont = Funcion.iEscalarSQL(ds, string.Format("Exec ValidaAnticipoDuplicado {0}, {1}", (int)uCmb.Value, (int)uCmbPersonal.Value), true);
			
			if (iCont > 0)
			{
				MessageBox.Show(string.Format("'{0}' tiene {1} '{2}' sin procesar", iCont, uCmb.Text.ToString()),  "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				uCmbPersonal.Focus();
				return false;
			}

			return true;
		}
		
		static public bool bValidaConciliado(int idAsiento, int idCuenta, C1.Data.C1DataSet ds)
		{			
			string sNumCB = Funcion.sEscalarSQL(ds, string.Format("Exec ValidaDepositoCajaConciliado {0}, {1}", idAsiento, idCuenta));
			if (sNumCB.Length > 0)
			{
				MessageBox.Show(string.Format("El movimiento se encuentra registrado en la Conciliacion N. '{0}'", sNumCB), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				return false;
			}

			return true;
		}
	}
}
