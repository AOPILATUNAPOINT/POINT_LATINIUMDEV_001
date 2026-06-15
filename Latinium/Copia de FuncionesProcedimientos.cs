using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Data;
using System.Data.SqlClient;
using System.Windows.Forms;
using Infragistics;
using Infragistics.Win.UltraWinGrid;
using System.Globalization;
using System.Data.OleDb;
using System.Data.Odbc;
using System.IO;
using CDO;
using System.Net;
using System.Web;
using System.Web.Mail;
using CrystalDecisions;
using CrystalDecisions.CrystalReports.Engine;
using CrystalDecisions.ReportSource;
using CrystalDecisions.Shared;
using System.Runtime.InteropServices;
using System.Diagnostics;

namespace Latinium
{
	/// <summary>
	/// Descripción breve de FuncionesProcedimientos.
	/// </summary>
	public class FuncionesProcedimientos
	{
		public FuncionesProcedimientos()
		{
			
		}

		static public bool bExportaPDF_PA(string sReporte, int idCompra, string Archivo)
		{
			try
			{
				ReportDocument Report = new ReportDocument();				
					
				Report.FileName = MenuLatinium.stDirReportes + sReporte;
								
				ConnectionInfo crConnectionInfo = new ConnectionInfo ();
			{
				string strBaseDato = MenuLatinium.stDirNumero;
				string strUsuario = MenuLatinium.stDirConta;
				string strServidor = MenuLatinium.stDirServidor;

				crConnectionInfo.ServerName = strServidor;
				crConnectionInfo.DatabaseName = strBaseDato; 
				crConnectionInfo.UserID = "infoelect"; 
				crConnectionInfo.Password = "Bl45o6$9"; 
			}
			
				TableLogOnInfo crTableLogOnInfo; 
				Database crDatabase = Report.Database; 
				Tables crTables = crDatabase.Tables; 
				Table crTable;

				for (int i=0; i< crTables.Count; i++)
				{
				{
					crTable = crTables [i]; 
					crTableLogOnInfo = crTable.LogOnInfo; 
					crTableLogOnInfo.ConnectionInfo = crConnectionInfo; 
					crTable.ApplyLogOnInfo(crTableLogOnInfo);
				}
				}

				Report.SetParameterValue("@idCompra", idCompra);
				Report.ExportToDisk(CrystalDecisions.Shared.ExportFormatType.PortableDocFormat, MenuLatinium.stDirInicio + @"\" + Archivo);

				return true;
			}
			catch (Exception ex)
			{
				string stError = ex.Message + "\n"; if (ex.InnerException.Message != null) stError += ex.InnerException.Message;
				MessageBox.Show(stError, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);

				return false;
			}
		}	

		static public void Envio_Mail(string sContactos, string sAsunto, string sArchivo, string sMensaje, C1.Data.C1DataSet ds)
		{
			/*	
			try
			{
				#region Datos de servidor de correo saliente
				string sSQL = string.Format("Select ISNULL(Email, ''), ISNULL(ServidorCorreo, ''), ISNULL(UsuarioCorreo, ''), ISNULL(ClaveCorreo, ''), ISNULL(PuertoCorreo, ''), ISNULL(SSLCorreo, 0) From Seteo");
				SqlDataReader dr = Funcion.rEscalarSQL(ds, sSQL, true);
				dr.Read();
				string sEMail = dr.GetString(0);
				string sServidorCorreo = dr.GetString(1);
				string sUsuarioCorreo = dr.GetString(2);
				string sClaveCorreo = dr.GetString(3);
				string sPuertoCorreo = dr.GetString(4);
				bool bSSLCorreo = dr.GetBoolean(5);
				dr.Close();
				#endregion Datos de servidor de correo saliente

				#region Validacion
				if (sEMail.Trim().Length == 0)
				{
					MessageBox.Show("Ingrese EMail, en Archivo - Propiedades - Generales Pestaña: Básico", "Información", MessageBoxButtons.OK, MessageBoxIcon.Information);
					return;
				}
				if (sServidorCorreo.Trim().Length == 0)
				{
					MessageBox.Show("Ingrese Servidor de Correo Saliente Ej. smtp.correo.yahoo.es, en Archivo Propiedades Generales Complemento, en Archivo - Propiedades - Generales Pestaña: Básico", "Información", MessageBoxButtons.OK, MessageBoxIcon.Information);
					return;
				}
				if (sUsuarioCorreo.Trim().Length == 0)
				{
					MessageBox.Show("Ingrese Usuario, en Archivo Propiedades Generales en la Pestaña con Correo Electronico", "Información", MessageBoxButtons.OK, MessageBoxIcon.Information);
					return;
				}			
				if (sClaveCorreo.Trim().Length == 0)
				{
					MessageBox.Show("Ingrese Contraseña, en Archivo - Propiedades - Generales Pestaña: Correo Electrónico", "Información", MessageBoxButtons.OK, MessageBoxIcon.Information);
					return;
				}
				if (sPuertoCorreo.Trim().Length == 0)
				{
					MessageBox.Show("Ingrese Puerto de Correo, en Archivo Propiedades Generales en la Pestaña con Correo Electronico", "Información", MessageBoxButtons.OK, MessageBoxIcon.Information);
					return;
				}
				#endregion Validacion

				EASendMail.SmtpMail oMail = new EASendMail.SmtpMail("TryIt");
				EASendMail.SmtpClient oSmtp = new EASendMail.SmtpClient();

				EASendMail.SmtpServer oServer = new EASendMail.SmtpServer("192.168.1.254");//sServidorCorreo);//

				oServer.User = "dpozo@point.com.ec";//sUsuarioCorreo; //
				oServer.Password = "dpozo2017$";//sClaveCorreo; //

				oServer.Port = 25; //int.Parse(sPuertoCorreo);
			
				oServer.ConnectType = SmtpConnectType.ConnectSSLAuto;

				oMail.From = "latinium@point.com.ec";//sUsuarioCorreo; //

				oMail.To = sContactos;

				oMail.Subject = sAsunto;

				oMail.TextBody = sMensaje;
				
//				oMail.AddAttachment(sArchivo);

				oSmtp.SendMail(oServer, oMail);				

				MessageBox.Show("E-Mail enviado correctamente", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Information);
			}
			catch (Exception ep)
			{
				MessageBox.Show(ep.Message, "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Error);
			}

					
		
			
			#region Configuracion de datos de Email			
			MailMessage eMail = new MailMessage();			
			eMail.To = sContactos;
			eMail.From = sEMail;//"diego-pozo@hotmail.com";//
			eMail.Subject = sAsunto;
			eMail.Body = sMensaje;
			eMail.Priority = MailPriority.Normal;
			eMail.BodyFormat = MailFormat.Text;			//o MailFormat.Html
			#endregion Configuracion de Email
			
			#region Activar Autenticación
						eMail.Fields.Add ("http://schemas.microsoft.com/cdo/configuration/smtpserver", sServidorCorreo);
						eMail.Fields.Add("http://schemas.microsoft.com/cdo/configuration/smtpserverport", sPuertoCorreo);														
						eMail.Fields.Add("http://schemas.microsoft.com/cdo/configuration/sendusing", 2);
						eMail.Fields.Add("http://schemas.microsoft.com/cdo/configuration/sendusername", sUsuarioCorreo);// Nombre de usuario
						eMail.Fields.Add("http://schemas.microsoft.com/cdo/configuration/sendpassword", sClaveCorreo);  // Contraseña
						eMail.Fields.Add("http://schemas.microsoft.com/cdo/configuration/smtpauthenticate", 1);				
						eMail.Fields.Add("http://schemas.microsoft.com/cdo/configuration/smtpusessl", bSSLCorreo);
						eMail.Fields.Add("http://schemas.microsoft.com/cdo/configuration/smtpconnectiontimeout", 600);						
			#endregion Activar Autenticación

			#region Adjunta archivo y envia
			SmtpMail.SmtpServer = sServidorCorreo;

			try
			{
				MailAttachment Adjunto = new MailAttachment(string.Format(@"{0}\{1}", MenuLatinium.stDirInicio, sArchivo));
				eMail.Attachments.Add(Adjunto);
				SmtpMail.Send(eMail); 
				MessageBox.Show("Correo Electronico Enviado", "Información", MessageBoxButtons.OK, MessageBoxIcon.Information);
			}
			catch (Exception ex)
			{				
				MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
			}			
			#endregion Adjunta archivo y envia	
			*/
		}

		static public void EstadoControles (Form oForm, bool bEstado, int iAccion)
		{
			foreach(Control oCtrl in oForm.Controls)
			{					
				#region Estado
				if (oCtrl.Name.Substring(0, 3) != "btn" && oCtrl.Name.Substring(0, 3) != "utc")
				{
					if (iAccion > 2) oCtrl.Enabled = bEstado;
					else if (iAccion == 1)
					{
						if (oCtrl.Name.Substring(3, 1) == "N") oCtrl.Enabled = bEstado;
					}
					else if (iAccion == 2)
					{
						if (oCtrl.Name.Substring(4, 1) == "E") oCtrl.Enabled = bEstado;
					}
				}
				#endregion Estado

				#region Valores de Controles
				if (iAccion == 4)
				{
					if (oCtrl is Infragistics.Win.UltraWinEditors.UltraNumericEditor)
					{					
						Infragistics.Win.UltraWinEditors.UltraNumericEditor oUNE = (Infragistics.Win.UltraWinEditors.UltraNumericEditor)oCtrl;					
						oUNE.Value = 0;
					}

					if (oCtrl is Infragistics.Win.UltraWinGrid.UltraCombo)
					{
						Infragistics.Win.UltraWinGrid.UltraCombo oCmb = (Infragistics.Win.UltraWinGrid.UltraCombo)oCtrl;
						oCmb.Value = System.DBNull.Value;
					}
				
					if (oCtrl is Infragistics.Win.UltraWinEditors.UltraTextEditor)
					{
						oCtrl.Text = "";
					}

					if (oCtrl is Infragistics.Win.UltraWinEditors.UltraCheckEditor)
					{					
						Infragistics.Win.UltraWinEditors.UltraCheckEditor oCHK = (Infragistics.Win.UltraWinEditors.UltraCheckEditor)oCtrl;
						oCHK.Checked = false;
					}
				}
				#endregion Valores de Controles
			}
		}

		static public void EstadoControlesTabControl (Infragistics.Win.UltraWinTabControl.UltraTabControl UTab, bool bEstado, int iAccion)
		{
			foreach(Infragistics.Win.UltraWinTabControl.UltraTab oTab in UTab.Tabs)
			{
				foreach(Control oCtrl in oTab.TabPage.Controls)
				{
					#region Estado
					if (oCtrl.Name.Substring(0, 3) != "utc")
					{
						if (iAccion > 2) oCtrl.Enabled = bEstado;
						else if (iAccion == 1)
						{
							if (oCtrl.Name.Substring(3, 1) == "N") oCtrl.Enabled = bEstado;
						}
						else if (iAccion == 2)
						{
							if (oCtrl.Name.Substring(4, 1) == "E") oCtrl.Enabled = bEstado;
						}
					}
					#endregion Estado
				
					#region Valores de Controles
					if (iAccion == 4)
					{
						if (oCtrl is Infragistics.Win.UltraWinEditors.UltraNumericEditor)
						{					
							Infragistics.Win.UltraWinEditors.UltraNumericEditor oUNE = (Infragistics.Win.UltraWinEditors.UltraNumericEditor)oCtrl;					
							oUNE.Value = 0;
						}

						if (oCtrl is Infragistics.Win.UltraWinGrid.UltraCombo)
						{
							Infragistics.Win.UltraWinGrid.UltraCombo oCmb = (Infragistics.Win.UltraWinGrid.UltraCombo)oCtrl;
							oCmb.Value = System.DBNull.Value;
						}
				
						if (oCtrl is Infragistics.Win.UltraWinEditors.UltraTextEditor)
						{
							oCtrl.Text = "";
						}

						if (oCtrl is Infragistics.Win.UltraWinEditors.UltraCheckEditor)
						{					
							Infragistics.Win.UltraWinEditors.UltraCheckEditor oCHK = (Infragistics.Win.UltraWinEditors.UltraCheckEditor)oCtrl;
							oCHK.Checked = false;
						}

						if (oCtrl is Infragistics.Win.UltraWinEditors.UltraOptionSet)
						{					
							Infragistics.Win.UltraWinEditors.UltraOptionSet oUOS = (Infragistics.Win.UltraWinEditors.UltraOptionSet)oCtrl;					
							oUOS.Value = System.DBNull.Value;
						}

						if (oCtrl is Infragistics.Win.UltraWinSchedule.UltraCalendarCombo)
						{					
							Infragistics.Win.UltraWinSchedule.UltraCalendarCombo oUCC = (Infragistics.Win.UltraWinSchedule.UltraCalendarCombo)oCtrl;					
							oUCC.Value = System.DBNull.Value;
						}
					}
					#endregion Valores de Controles
				}
			}
		}

		static public void DesplazamientoGrids(object sender, System.Windows.Forms.KeyEventArgs e, Infragistics.Win.UltraWinGrid.UltraGrid uGrid)
		{
			if (uGrid.ActiveCell != null)
			{
				try
				{				
					switch(e.KeyValue)
					{							
							#region Flechas
						case (int) Keys.Up:
							//							if (uGrid.ActiveCell.DroppedDown) return;

							uGrid.PerformAction(UltraGridAction.ExitEditMode, false, false);
							uGrid.PerformAction(UltraGridAction.AboveCell, false, false);
							e.Handled = true;
							uGrid.PerformAction(UltraGridAction.EnterEditMode, false, false);
							break;
						case (int) Keys.Left:
							uGrid.PerformAction(UltraGridAction.ExitEditMode, false, false);
							uGrid.PerformAction(UltraGridAction.PrevCellByTab, false, false);
							e.Handled = true;
							uGrid.PerformAction(UltraGridAction.EnterEditMode, false, false);
							break;
						case (int) Keys.Down:
							//							if (uGrid.ActiveCell.DroppedDown) return;
			
							uGrid.PerformAction(UltraGridAction.ExitEditMode, false, false);
							uGrid.PerformAction(UltraGridAction.BelowCell, false, false);
							e.Handled = true;
							uGrid.PerformAction(UltraGridAction.EnterEditMode, false, false);
							break;
							#endregion Flechas
							#region Enter
						case (int) Keys.Enter:
							uGrid.PerformAction(UltraGridAction.ExitEditMode, false, false);
							uGrid.PerformAction(UltraGridAction.NextCellByTab, false, false);
							e.Handled = true;
							uGrid.PerformAction(UltraGridAction.EnterEditMode, false, false);							
							break;
							#endregion Enter
					}
				}
				catch(System.Exception ex)
				{
					MessageBox.Show(ex.Message, "Error del sistema.");
				}
			}
		}

		static public void FormatoGrid(Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs e, string sColumna, int iDecimales)
		{
			#region Formato numerico
			CultureInfo cultureENUS = CultureInfo.CreateSpecificCulture("ja-jp");
			string stMoneda = "#,##0.";
			string stMascara = "nnn,nnn,nnn.";
			string stCeros = "0000000000";
			string stNumero = "nnnnnnnnnn";
			string stFormato = stMoneda + stCeros.Substring(0, iDecimales);
			string stInput = stMascara + stNumero.Substring(0, iDecimales);
					
			e.Layout.Bands[0].Columns[sColumna].Format = stFormato; 
			e.Layout.Bands[0].Columns[sColumna].MaskInput = stInput;
			#endregion Formato numerico
		}

		static public void EstadoGrids(bool bEstado, Infragistics.Win.UltraWinGrid.UltraGrid uGrid)
		{
			uGrid.Enabled = true;

			if (bEstado)
			{
				uGrid.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.True;
				uGrid.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.True;
			}
			else
			{
				uGrid.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.False;
				uGrid.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.False;
			}
		}		

		static public DataSet ImportaDeExcel(string sArchivo, string sComando, string sNombreHoja)
		{
			OleDbConnection oConexion = new OleDbConnection();
			OleDbCommand oCmd = new OleDbCommand();
			OleDbDataAdapter oDA = new OleDbDataAdapter();
			DataSet oDS = new DataSet();

			try
			{				
				oConexion.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + sArchivo + "; Extended Properties= Excel 8.0;";
				oCmd.CommandText = string.Format(sComando);

				oCmd.Connection = oConexion;
				oDA.SelectCommand = oCmd;
				oDA.Fill(oDS, sNombreHoja);				
			}
			catch (Exception ex)
			{
				try
				{
					MessageBox.Show(string.Format("Commit Exception Type : {0} {1}", ex.GetType(), ex.Message), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Error);				
				}
				catch (Exception ex2)
				{
					MessageBox.Show(string.Format("Rollback Exception Type: {0} {1}", ex2.GetType(), ex2.Message), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Error);												
				}
			}

			return oDS;
		}

		static public decimal CalculaMeta(int iBodega, int iMes, int iAño, int idCargo, int iNumeroDias, int iTipoMeta, C1.Data.C1DataSet ds)
		{
			decimal dMeta = 0.00m;

			if (Funcion.bEscalarSQL(ds, string.Format("Select Meta From Com_CargosDeVentas Where idCargo = {0}", idCargo), true))
			{
				string sSQL = string.Format("Select dbo.Com_RetornaMetaPersonalLocalesPorCargo({0}, {1}, {2}, {3}, {4})", 
					iAño, iMes, iBodega, idCargo, iTipoMeta);
				decimal dValorMeta = Funcion.decEscalarSQL(ds, sSQL);
			
				int iDiasMes = DateTime.DaysInMonth(iAño, iMes);
		
				if (iDiasMes == iNumeroDias) dMeta = dValorMeta;
				else dMeta = (dValorMeta/iDiasMes) * iNumeroDias;
			}					

			return Math.Round(dMeta, 2);
		}

		static public decimal CalculaMetaGE(int idEstado, string sPeriodo, int iBodega, int idCargo, DateTime dtDesde, int iNumeroDias, C1.Data.C1DataSet ds)
		{
			decimal dMetaMensual = 0.00m;

			if (Funcion.bEscalarSQL(ds, string.Format("Select CalculaMeta From Com_Estado Where idCom_Estado = {0}", idEstado), true))
			{
				if (Funcion.bEscalarSQL(ds, string.Format("Select Meta From Com_CargosDeVentas Where idCargo = {0}", idCargo), true))
				{
					int iAño = int.Parse(sPeriodo.Substring(3, 4));
					int iMes = int.Parse(sPeriodo.Substring(0, 2));

					string sSQL = string.Format("Select dbo.Com_RetornaMetaPorCargoGE({0}, {1}, {2}, {3})", iAño, iMes, iBodega, idCargo);
					decimal dValorMeta = Funcion.decEscalarSQL(ds, sSQL);

					int iDiasMes = DateTime.DaysInMonth(dtDesde.Year, dtDesde.Month);
					
					if (iDiasMes == iNumeroDias) dMetaMensual = dValorMeta;
					else dMetaMensual = (dValorMeta/iDiasMes) * iNumeroDias;
				}
			}			

			return Math.Round(dMetaMensual, 2);
		}

		static public decimal CalculaMetaCarpa(int idEstado, string sPeriodo, int iBodega, int idCargo, DateTime dtDesde, int iNumeroDias, C1.Data.C1DataSet ds)
		{
			decimal dMetaMensual = 0.00m;

			if (Funcion.bEscalarSQL(ds, string.Format("Select Carpa From Bodega Where Bodega = {0}", iBodega), true))
			{			
				if (Funcion.bEscalarSQL(ds, string.Format("Select CalculaMeta From Com_Estado Where idCom_Estado = {0}", idEstado), true))
				{
					if (Funcion.bEscalarSQL(ds, string.Format("Select Meta From Com_CargosDeVentas Where idCargo = {0}", idCargo), true))
					{
						int iAño = int.Parse(sPeriodo.Substring(3, 4));
						int iMes = int.Parse(sPeriodo.Substring(0, 2));

						string sSQL = string.Format("Select dbo.Com_RetornaMetaPorCargoCarpa({0}, {1}, {2}, {3})", iAño, iMes, iBodega, idCargo);
						decimal dValorMeta = Funcion.decEscalarSQL(ds, sSQL);

						int iDiasMes = DateTime.DaysInMonth(dtDesde.Year, dtDesde.Month);
					
						if (iDiasMes == iNumeroDias) dMetaMensual = dValorMeta;
						else dMetaMensual = (dValorMeta/iDiasMes) * iNumeroDias;
					}
				}
			}			

			return Math.Round(dMetaMensual, 2);
		}

		static public decimal dRetornaMetaPorLocal(int iPeriodo, int iMes, int iBodega, int idCargo, C1.Data.C1DataSet ds)
		{
			return Funcion.decEscalarSQL(ds, string.Format("Select dbo.Com_RetornaMetaPorCargo({0}, {1}, {2}, {3})", iPeriodo, iMes, iBodega, idCargo));
		}

		static public int CalculaDias(DateTime dtDesde, DateTime dtHasta, C1.Data.C1DataSet ds)
		{
			int iNumeroDias = 0;

			iNumeroDias = Funcion.iEscalarSQL(ds, string.Format("Select DateDiff(DAY, '{0}', '{1}') + 1", dtDesde.ToString("yyyyMMdd"), dtHasta.ToString("yyyyMMdd")));

			return iNumeroDias;
		}

		static public int iUltimoDiaDelMes(int iMes, int iAño)
		{			
			return DateTime.DaysInMonth(iAño, iMes);
		}

		static public bool ValidaPeriodoCerrado(string sPeriodo, C1.Data.C1DataSet ds)
		{
			bool bRes = Funcion.bEscalarSQL(ds, string.Format("Exec Com_ValidaPeriodoCerrado '{0}'", sPeriodo), true);

			return bRes;
		}

		static public DataView Lista_Vendedores(DateTime dtFecha, int iBodega, C1.Data.C1DataSet ds)
		{
			return Funcion.dvProcedimiento(ds, string.Format("Exec FacturacionListaVendedores '{0}', {1}, {2}", dtFecha.ToString("yyyyMMdd"), iBodega, MenuLatinium.iNivel));	
		}

		static public DataView Lista_Choferes(DateTime dtFecha, int iBodega, C1.Data.C1DataSet ds)
		{
			return Funcion.dvProcedimiento(ds, string.Format("Exec ListaChoferesGuiasDeResmision '{0}', {1}", dtFecha.ToString("yyyyMMdd"), iBodega));	
		}

		static public DataView Lista_VendedoresPlanifCarpa(DateTime dtFecha, int iBodega, C1.Data.C1DataSet ds)
		{
			return Funcion.dvProcedimiento(ds, string.Format("Exec ListaVendedoresPlanifCarpa '{0}', {1}", dtFecha.ToString("yyyyMMdd"), iBodega));	
		}

		static public string AbreListaClientesProveedores(bool bClienteProveedor, int idTipoFactura)
		{			
			string sIdentificacion = "";
			
			using (frmListaClientesProveedores LCP = new frmListaClientesProveedores(bClienteProveedor, idTipoFactura))
			{
				if (DialogResult.OK == LCP.ShowDialog())
				{
					sIdentificacion = LCP.uGridClientesProveedores.ActiveRow.Cells["Ruc"].Value.ToString();
				}
			}

			return sIdentificacion;
		}

		static public string AbreFacturacionClientes(int idTipoRuc, string sIdentificacion, bool bCancelar)
		{			
			string sRIdentificacion = "";
			
			using (frmFacturacionRegistroDeClientes FRC = new frmFacturacionRegistroDeClientes(idTipoRuc, sIdentificacion, bCancelar))
			{
				if (DialogResult.OK == FRC.ShowDialog())
				{
					sRIdentificacion = FRC.txtIdentificacion.Text.ToString();
				}
			}

			return sRIdentificacion;
		}
 /**Incrementado por angélica */
		static public string AbreFacturacionClientes1(bool BClienteProveedor,int IOrigen, int IdClienteC)
		{			
			string sIdentificacion = "";
			
			using(frmClientesProveedores FRC1 = new frmClientesProveedores(false,1,IdClienteC))
			{
				if (DialogResult.OK == FRC1.ShowDialog())
				{
					sIdentificacion = FRC1.txtRuc.Text.ToString();
				}
			}

			return sIdentificacion;
		}

		static public string[] RetornaParametrosClientesProveedores (bool bClienteProveedor, int idTipoFactura, string sIdentificacion, C1.Data.C1DataSet ds)
		{			
			string[] sParametros = new string[13];
			
			string sSQL = string.Format("Exec RetornaParametrosClientesProveedores {0}, {1}, '{2}'", bClienteProveedor, idTipoFactura, sIdentificacion);
			SqlDataReader dr = Funcion.rEscalarSQL(ds, sSQL, true);
			dr.Read();
			if (dr.HasRows)
			{
				sParametros[0] = dr.GetValue(0).ToString();
				sParametros[1] = dr.GetValue(1).ToString();
				sParametros[2] = dr.GetValue(2).ToString();
				sParametros[3] = dr.GetValue(3).ToString();
				sParametros[4] = dr.GetValue(4).ToString();
				sParametros[5] = dr.GetValue(5).ToString();
				sParametros[6] = dr.GetValue(6).ToString();
				sParametros[7] = dr.GetValue(7).ToString();
				sParametros[8] = dr.GetValue(8).ToString();
				sParametros[9] = dr.GetValue(9).ToString();
				sParametros[10] = dr.GetValue(10).ToString();
				sParametros[11] = dr.GetValue(11).ToString();
				sParametros[12] = dr.GetValue(12).ToString();
			}
			dr.Close();
			
			return sParametros;
		}	
					
		static public DataView Lista_TiposDeIdentificacion(int iBodega, int idTipoFactura, C1.Data.C1DataSet ds)
		{
			return Funcion.dvProcedimiento(ds, string.Format("Exec FacturacionListaTiposDeIdentificacion {0}, {1}", iBodega, idTipoFactura));	
		}

		static public void FechasDocumentos (Infragistics.Win.UltraWinSchedule.UltraCalendarCombo dtFecha)
		{
			dtFecha.Value = MenuLatinium.dtFechaSistema.Date;

			dtFecha.CalendarInfo.MaxDate = MenuLatinium.dtFechaSistema.Date;
		}

		static public void RetornaFechaServidor (Infragistics.Win.UltraWinSchedule.UltraCalendarCombo dtFecha, C1.Data.C1DataSet ds, bool bHora)
		{
			DateTime dtFechaServ = Funcion.dtEscalarSQL(ds, "Exec RetornaFechaServidor");

			if (!bHora) dtFecha.Value = dtFechaServ.Date;
			if (bHora) dtFecha.Value = dtFechaServ;

			dtFecha.CalendarInfo.MaxDate = dtFechaServ.Date;
		} 

		static public DateTime dtRetornaFechaServidor (C1.Data.C1DataSet ds)
		{
			return Funcion.dtEscalarSQL(ds, "Exec RetornaFechaServidor");
		}

		static public DataView ListaDireccionesClientes (int idCliente, C1.Data.C1DataSet ds)
		{
			return Funcion.dvProcedimiento(ds, string.Format("Select idDreccion, Direccion From ClientesDirecciones Where idCliente = {0}", idCliente));
		}

		static public void ListaBodegas (Infragistics.Win.UltraWinGrid.UltraCombo uCombo, int idUsuario, int idTipoFactura, DateTime dtFecha, bool bRecibeConsignacion, C1.Data.C1DataSet ds)
		{
			uCombo.DataSource = Funcion.dvProcedimiento(ds, string.Format("Exec FacturacionListaBodegas {0}, {1}, '{2}', {3}", idUsuario, idTipoFactura, dtFecha.ToString("yyyyMMdd"), bRecibeConsignacion));
		}

		static public void ListaImportacionesActivas (Infragistics.Win.UltraWinGrid.UltraCombo uCombo, int idImportacion, C1.Data.C1DataSet ds)
		{
			uCombo.DataSource = Funcion.dvProcedimiento(ds, string.Format("Exec ListaImportacionesActivas {0}", idImportacion));
		}

		static public void ListaPersonalFacturas (Infragistics.Win.UltraWinGrid.UltraCombo uCombo, DateTime dtFecha, int idCargo, C1.Data.C1DataSet ds)
		{
			uCombo.DataSource = Funcion.dvProcedimiento(ds, string.Format("Exec ListaPersonalFacturas '{0}', {1}", dtFecha.ToString("yyyyMMdd"), idCargo));
		}

		static public int NuevoRegistroVentas (int idTipoFactura)
		{
			int iFormaPago = 0;
			using (frmFormasDePago FormaDePago = new frmFormasDePago(idTipoFactura, false, false, 0))
			{				
				if (DialogResult.OK == FormaDePago.ShowDialog())
				{
					iFormaPago = (int) FormaDePago.ultraGrid1.ActiveRow.Cells["idFormaPago"].Value;
				}
				return iFormaPago;
			}
		}
				
		static public void GuardaDetalleHorariosDeCarpa(Infragistics.Win.UltraWinGrid.UltraGrid uGrid, int idHorariosDeCarpa, int iTipo, C1.Data.C1DataSet ds)//(int idDetalleHorariosDeCarpa, int idHorariosDeCarpa, int idPersonal, int Semana, int Lunes, int Martes, int Miercoles, int Jueves, int Viernes, int Sabado, int Domingo, string sSector, C1.Data.C1DataSet ds)
		{
			int idPersonal = 0;
			decimal dPresupuesto = 0.00m;
			string sSector = "";
			string sColumna = "idDetallePlanificacionVentas";

			int iLunes = 0;
			int iMartes = 0;
			int iMiercoles = 0;
			int iJueves = 0;
			int iViernes = 0;
			int iSabado = 0;
			int iDomingo = 0;

			foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in uGrid.Rows.All)
			{
				iLunes = 0;
				iMartes = 0;
				iMiercoles = 0;
				iJueves = 0;
				iViernes = 0;
				iSabado = 0;
				iDomingo = 0;

				if (iTipo == 1)
				{
					idPersonal = (int)dr.Cells["idPersonal"].Value;
					dPresupuesto = Convert.ToDecimal(dr.Cells["Presupuesto"].Value);
					sSector = dr.Cells["Sector"].Value.ToString();
					sColumna = "idDetalleHorariosDeCarpa";
				}

				if (dr.Cells["Lunes"].Value != System.DBNull.Value) iLunes = (int)dr.Cells["Lunes"].Value;
				if (dr.Cells["Martes"].Value != System.DBNull.Value) iMartes = (int)dr.Cells["Martes"].Value;
				if (dr.Cells["Miercoles"].Value != System.DBNull.Value) iMiercoles = (int)dr.Cells["Miercoles"].Value;
				if (dr.Cells["Jueves"].Value != System.DBNull.Value) iJueves = (int)dr.Cells["Jueves"].Value;
				if (dr.Cells["Viernes"].Value != System.DBNull.Value) iViernes = (int)dr.Cells["Viernes"].Value;
				if (dr.Cells["Sabado"].Value != System.DBNull.Value) iSabado = (int)dr.Cells["Sabado"].Value;
				if (dr.Cells["Domingo"].Value != System.DBNull.Value) iDomingo = (int)dr.Cells["Domingo"].Value;

				string sSQL = string.Format("Exec GuardaDetalleHorariosCarpa {0}, {1}, {2}, {3}, {4}, {5}, {6}, {7}, {8}, {9}, {10}, {11}, '{12}', {13}",
					(int)dr.Cells[sColumna].Value, idHorariosDeCarpa, idPersonal, (int)dr.Cells["Semana"].Value, 
					iLunes, iMartes, iMiercoles, iJueves, iViernes, iSabado, iDomingo, dPresupuesto, sSector, iTipo);
				Funcion.EjecutaSQL(ds, sSQL, true);	
			}
		}

		static public void ConsultaDetalleHorariosDeCarpa (Infragistics.Win.UltraWinGrid.UltraGrid uGrid, int idHorariosDeCarpa, int iSemana, int iTipo, C1.Data.C1DataSet ds)
		{
			string sSQL = string.Format("Exec ConsultaDetalleHorariosCarpa {0}, {1}, {2}", idHorariosDeCarpa, iSemana, iTipo);
			uGrid.DataSource = Funcion.dvProcedimiento(ds, sSQL);
		}

		static public void IniciaValoresColumnasDetalleHorariosDeCarpa(object sender, Infragistics.Win.UltraWinGrid.RowEventArgs e, Infragistics.Win.UltraWinGrid.UltraGrid uGrid, int iSemana, int iTipo)
		{
			string sColumna = "idDetallePlanificacionVentas";
			string sColumna1 = "idPlanificacionVentas";

			if (iTipo == 1) 
			{
				sColumna = "idDetalleHorariosDeCarpa";
				sColumna1 = "idHorariosDeCarpa";

				e.Row.Cells["idPersonal"].Value = 0;
				e.Row.Cells["Presupuesto"].Value = 0.00m;
			}

			e.Row.Cells[sColumna].Value = 0;
			e.Row.Cells[sColumna1].Value = 0;
			e.Row.Cells["Semana"].Value = iSemana;
			e.Row.Cells["Lunes"].Value = 0;
			e.Row.Cells["Martes"].Value = 0;
			e.Row.Cells["Miercoles"].Value = 0;
			e.Row.Cells["Jueves"].Value = 0;
			e.Row.Cells["Viernes"].Value = 0;
			e.Row.Cells["Sabado"].Value = 0;
			e.Row.Cells["Domingo"].Value = 0;			
		}

		static public void MostrarOcultar(Infragistics.Win.UltraWinGrid.UltraGrid uGrid, bool bEstado)
		{
			for (int j = 3; j <= 9; j++)
			{
				uGrid.DisplayLayout.Bands[0].Columns[j].Hidden = bEstado;
				uGrid.DisplayLayout.Bands[0].Columns[j].Header.Caption = uGrid.DisplayLayout.Bands[0].Columns[j].Key.ToString();
			}		
		}
		
		static public void dsGeneral(string sSQL, Infragistics.Win.UltraWinGrid.UltraGrid uGrid)
		{
			using (SqlConnection oConexion = new SqlConnection("Password=Bl45o6$9;Persist Security Info=True;User ID= 'infoelect';Initial Catalog=" + MenuLatinium.stDirFacturacion + ";Data Source=" + MenuLatinium.stDirServidor + ";Max Pool Size = 1024;"))
			{
				try
				{
					SqlCommand oCmdActualiza = oConexion.CreateCommand();
					SqlDataAdapter oDADetalle = new SqlDataAdapter(oCmdActualiza);
					DataSet oDSDetalle = new DataSet();
				
					oCmdActualiza.Connection = oConexion;					
					oCmdActualiza.CommandTimeout = 120;						
					oCmdActualiza.CommandText = sSQL;		
									
					oDADetalle.Fill(oDSDetalle);
					
					uGrid.DataSource = oDSDetalle.Tables[0];					
				}
				catch(Exception e) 
				{
					MessageBox.Show(string.Format("Error de Consulta : {0}", e.ToString()));
				}
			}
		}

		static public DataTable dtGeneral(string sSQL)
		{
			using (SqlConnection oConexion = new SqlConnection("Password=Bl45o6$9;Persist Security Info=True;User ID= " + MenuLatinium.stUsuario + ";Initial Catalog=" + MenuLatinium.stDirFacturacion + ";Data Source=" + MenuLatinium.stDirServidor + "; max pool size=1024"))
			{					
				SqlCommand oCmdActualiza = oConexion.CreateCommand();
				SqlDataAdapter oDADetalle = new SqlDataAdapter(oCmdActualiza);
				DataTable oDTDetalle = new DataTable();

				try
				{
					oCmdActualiza.Connection = oConexion;					
					oCmdActualiza.CommandTimeout = 120;				
					oCmdActualiza.CommandText = sSQL;										
					oDADetalle.Fill(oDTDetalle);					
				}
				catch(Exception e) 
				{				
					MessageBox.Show(string.Format("Error de Consulta : {0}", e.ToString()));
				}

				return oDTDetalle;
			}
		}

		static public void EliminaFilasGridNomina(Infragistics.Win.UltraWinGrid.BeforeRowsDeletedEventArgs e, int iNumero, string sColumna, bool bElimina, C1.Data.C1DataSet ds)
		{
			e.DisplayPromptMsg = false;
			for (int i=0; i<e.Rows.Length; i++)
			{
				if (!bElimina && (int)e.Rows[i].Cells[sColumna].Value > 0)
				{
					MessageBox.Show("No tiene acceso a eliminar", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
					e.Cancel = true;
					return;
				}
			
				if (iNumero == 8 || iNumero == 9)
				{
					if ((int)e.Rows[i].Cells["Estado"].Value == 1)
					{
						MessageBox.Show("No puede eliminar ya se aplico en Rol de Pagos", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
						e.Cancel = true;
						return;
					}
				
					if (Funcion.bEscalarSQL(ds, string.Format("Exec NominaValidaRolPagado '{0}'", Convert.ToDateTime(e.Rows[i].Cells["Fecha"].Value).ToString("yyyyMMdd")), true))
					{
						MessageBox.Show("No puede eliminar la falta, ya esta generado el Rol Del Periodo", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
						e.Cancel = true;
						return;
					}
				}
			}

			if (DialogResult.Yes ==	MessageBox.Show("¿Esta Seguro de Borrar este registro?", "Confirmación", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2)) 
			{
				for (int i=0; i<e.Rows.Length; i++)
				{
					if ((int)e.Rows[i].Cells[sColumna].Value > 0)
					{
						string sSQL = string.Format("Exec NominaEliminaRegistrosTablas {0}, {1}", iNumero, (int)e.Rows[i].Cells[sColumna].Value);
						Funcion.EjecutaSQL(ds, sSQL, true);
					}
				}
			}
			else e.Cancel = true;
		}
		
		static public void ExportaExcel (Infragistics.Win.UltraWinGrid.UltraGrid uGrid)
		{
			if (uGrid.Rows.Count == 0)
			{
				MessageBox.Show("No hay Filas para Exportar", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);				
				return;
			}

			SaveFileDialog sArchivo = new SaveFileDialog();
			sArchivo.Filter = "Excel (*.xls)|*.xls";

			Infragistics.Win.UltraWinGrid.ExcelExport.UltraGridExcelExporter oExportar = new Infragistics.Win.UltraWinGrid.ExcelExport.UltraGridExcelExporter();

			if(sArchivo.ShowDialog() == System.Windows.Forms.DialogResult.OK && sArchivo.FileName.Length > 0) 
			{				
				oExportar.Export(uGrid, sArchivo.FileName);
				System.Diagnostics.Process.Start(sArchivo.FileName);
			}
		}

		static public string sMensajeAprobado(int idPlanificacionVentas, int iTipo, C1.Data.C1DataSet ds, int iSemana)
		{
			string sMensaje = "";
			if (Funcion.iEscalarSQL(ds, string.Format("Exec RetornaIdASHC {0}, {1}, {2}", iSemana, idPlanificacionVentas, iTipo)) > 0)
				sMensaje = " - APROBADO";

			return sMensaje;
		}

		static public void AbreVistaMensajes(string sMensaje)
		{
			if (sMensaje.Length == 0)
			{
				MessageBox.Show("No existe texto para mostrar", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				return;
			}

			frmVistaMensajes FVM = new frmVistaMensajes(sMensaje);
			FVM.ShowDialog();
		}

		static public void NominaFoto(string sUbicacion, string sCedula, PictureBox pbFoto)
		{
			string sArchivo = sUbicacion + sCedula + ".jpg";

//			if (!File.Exists(sArchivo))
//			{
//				MessageBox.Show("Archivo de foto no existe", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
//				return;
//			}

			if (File.Exists(sArchivo))
				pbFoto.Image = Image.FromFile(sArchivo);
		}

		static public bool CVValidaCapturadoDepositado (int idCompra, int iOrigen, C1.Data.C1DataSet ds)
		{
			if (Funcion.iEscalarSQL(ds, string.Format("Exec CVValidaCapturadoDepositado {0}, {1}", idCompra, iOrigen)) > 0)
			{
				MessageBox.Show("No puede anular este documento\n\nExisten vouchers de tarjeta de credito/debito capturados", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				return false;
			}

			return true;
		}

		static public void RegistraBloqueoTransacciones (int idBloqueaTransacciones, bool bBloqueado, int idOrigen, int iOrigen, C1.Data.C1DataSet ds)
		{//@idBloqueoTransacciones Int, @Bloqueado Bit, @idOrigen Int, @Origen Int
			string sSQL = string.Format("Exec RegistraBloqueoTransacciones {0}, {1}, {2}, {3}", idBloqueaTransacciones, bBloqueado, idOrigen, iOrigen);
			Funcion.EjecutaSQL(ds, sSQL);
		}

		static public int iDiaSemana (DateTime dtFecha, C1.Data.C1DataSet ds)
		{
			return Funcion.iEscalarSQL(ds, string.Format("Select DATEPART(DW, '{0}')", dtFecha.ToString("yyyyMMdd")));
		}
	}
}
