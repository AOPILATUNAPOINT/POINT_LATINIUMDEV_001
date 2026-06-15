


using System;
using System.Globalization;
using System.Data;
using System.Data.SqlClient;
using C1.Data;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using Infragistics.Win.UltraWinGrid;
using System.Web.Mail;
using CrystalDecisions.Shared;
using System.Runtime.InteropServices;
using System.Diagnostics;
using System.IO;
using CrystalDecisions;
using CrystalDecisions.CrystalReports.Engine;
using CrystalDecisions.ReportSource;
using System.Net;
using System.Web;
using System.Xml;
using System.Xml.Schema;
using CrystalDecisions.Shared;

using Infragistics.Win.UltraWinToolbars;

namespace Latinium
{
	/// <summary>
	/// Descripción breve de frmImpresiondocumentosDeCredito.
	/// </summary>
	public class frmImpresiondocumentosDeCredito : System.Windows.Forms.Form
	{
		private Infragistics.Win.UltraWinEditors.UltraOptionSet optEstado;
		public DevExpress.XtraEditors.SimpleButton btCancelar;
		public DevExpress.XtraEditors.SimpleButton btGrabar;
		private C1.Data.C1DataSet cdsSeteoF;
		/// <summary>
		/// Variable del diseñador requerida.
		/// </summary>
		private System.ComponentModel.Container components = null;

		public frmImpresiondocumentosDeCredito()
		{
			//
			// Necesario para admitir el Diseñador de Windows Forms
			//
			InitializeComponent();

			//
			// TODO: agregar código de constructor después de llamar a InitializeComponent
			//
		}

		/// <summary>
		/// Limpiar los recursos que se estén utilizando.
		/// </summary>
		protected override void Dispose( bool disposing )
		{
			if( disposing )
			{
				if(components != null)
				{
					components.Dispose();
				}
			}
			base.Dispose( disposing );
		}

		#region Código generado por el Diseñador de Windows Forms
		/// <summary>
		/// Método necesario para admitir el Diseñador. No se puede modificar
		/// el contenido del método con el editor de código.
		/// </summary>
		private void InitializeComponent()
		{
			Infragistics.Win.Appearance appearance1 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance2 = new Infragistics.Win.Appearance();
			Infragistics.Win.ValueListItem valueListItem1 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.ValueListItem valueListItem2 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.ValueListItem valueListItem3 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.ValueListItem valueListItem4 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.ValueListItem valueListItem5 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.ValueListItem valueListItem6 = new Infragistics.Win.ValueListItem();
			this.optEstado = new Infragistics.Win.UltraWinEditors.UltraOptionSet();
			this.btCancelar = new DevExpress.XtraEditors.SimpleButton();
			this.btGrabar = new DevExpress.XtraEditors.SimpleButton();
			this.cdsSeteoF = new C1.Data.C1DataSet();
			((System.ComponentModel.ISupportInitialize)(this.optEstado)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).BeginInit();
			this.SuspendLayout();
			// 
			// optEstado
			// 
			appearance1.ForeColorDisabled = System.Drawing.Color.Black;
			this.optEstado.Appearance = appearance1;
			this.optEstado.BorderStyle = Infragistics.Win.UIElementBorderStyle.Solid;
			this.optEstado.CheckedIndex = 0;
			this.optEstado.ItemAppearance = appearance2;
			this.optEstado.ItemOrigin = new System.Drawing.Point(10, 0);
			valueListItem1.DataValue = 0;
			valueListItem1.DisplayText = "Pagaré  en Blanco";
			valueListItem2.DataValue = 1;
			valueListItem2.DisplayText = "Tabla De Amortización En Blanco";
			valueListItem3.DataValue = 2;
			valueListItem3.DisplayText = "Gastos de cobranza Pichincha en Blanco";
			valueListItem4.DataValue = 5;
			valueListItem4.DisplayText = "Gastos de cobranza en Blanco Solidario";
			valueListItem5.DataValue = 3;
			valueListItem5.DisplayText = "Seguro D&D en Blanco";
			valueListItem6.DataValue = 4;
			valueListItem6.DisplayText = "Compromiso Lugar de Pago";
			this.optEstado.Items.Add(valueListItem1);
			this.optEstado.Items.Add(valueListItem2);
			this.optEstado.Items.Add(valueListItem3);
			this.optEstado.Items.Add(valueListItem4);
			this.optEstado.Items.Add(valueListItem5);
			this.optEstado.Items.Add(valueListItem6);
			this.optEstado.ItemSpacingVertical = 8;
			this.optEstado.Location = new System.Drawing.Point(8, 8);
			this.optEstado.Name = "optEstado";
			this.optEstado.Size = new System.Drawing.Size(248, 144);
			this.optEstado.TabIndex = 960;
			this.optEstado.Text = "Pagaré  en Blanco";
			// 
			// btCancelar
			// 
			this.btCancelar.Anchor = System.Windows.Forms.AnchorStyles.None;
			this.btCancelar.Location = new System.Drawing.Point(140, 160);
			this.btCancelar.Name = "btCancelar";
			this.btCancelar.Size = new System.Drawing.Size(88, 24);
			this.btCancelar.TabIndex = 962;
			this.btCancelar.Text = "&Cancelar";
			// 
			// btGrabar
			// 
			this.btGrabar.Anchor = System.Windows.Forms.AnchorStyles.None;
			this.btGrabar.Location = new System.Drawing.Point(36, 160);
			this.btGrabar.Name = "btGrabar";
			this.btGrabar.Size = new System.Drawing.Size(88, 24);
			this.btGrabar.TabIndex = 961;
			this.btGrabar.Text = "&Imprimir";
			this.btGrabar.Click += new System.EventHandler(this.btGrabar_Click);
			// 
			// cdsSeteoF
			// 
			this.cdsSeteoF.BindingContextCtrl = this;
			this.cdsSeteoF.DataLibrary = "LibFacturacion";
			this.cdsSeteoF.DataLibraryUrl = "";
			this.cdsSeteoF.DataSetDef = "dsSeteoF";
			this.cdsSeteoF.Locale = new System.Globalization.CultureInfo("es-EC");
			this.cdsSeteoF.Name = "cdsSeteoF";
			this.cdsSeteoF.SchemaClassName = "LibFacturacion.DataClass";
			this.cdsSeteoF.SchemaDef = null;
			this.cdsSeteoF.BeforeFill += new C1.Data.FillEventHandler(this.cdsSeteoF_BeforeFill);
			// 
			// frmImpresiondocumentosDeCredito
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 13);
			this.ClientSize = new System.Drawing.Size(264, 190);
			this.Controls.Add(this.btCancelar);
			this.Controls.Add(this.btGrabar);
			this.Controls.Add(this.optEstado);
			this.KeyPreview = true;
			this.MaximizeBox = false;
			this.MinimizeBox = false;
			this.Name = "frmImpresiondocumentosDeCredito";
			this.Text = "Documentación";
			this.KeyDown += new System.Windows.Forms.KeyEventHandler(this.frmImpresiondocumentosDeCredito_KeyDown);
			this.Load += new System.EventHandler(this.frmImpresiondocumentosDeCredito_Load);
			((System.ComponentModel.ISupportInitialize)(this.optEstado)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).EndInit();
			this.ResumeLayout(false);

		}
		#endregion

		private void frmImpresiondocumentosDeCredito_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
				if (e.KeyCode == Keys.Escape) this.Close();
		}

		private void btGrabar_Click(object sender, System.EventArgs e)
		{
			DateTime FHasta = DateTime.Today ;
			using (frmImprimeRolesDePago IRP = new frmImprimeRolesDePago(75))
			{
				if (DialogResult.OK == IRP.ShowDialog())
				{
					int iCont = 0;

					if ((int)this.optEstado.Value == 0)
					{						
						#region Impresion
						string sReporte = "PagareConDatos_2022.rpt";
						//							if (MenuLatinium.stDirFacturacion == "LOTESERVICIO") sReporte = "RolIndividualPeony.rpt";
						ReportDocument Report1 = new ReportDocument();
						Report1.FileName = MenuLatinium.stDirReportes + sReporte;//"RolesIndividuales.rpt";
								
						ConnectionInfo crConnectionInfo = new ConnectionInfo (); 
					{
						string strBaseDato = MenuLatinium.stDirNumero;
						string strUsuario = MenuLatinium.stDirConta;
						string strServidor = MenuLatinium.stDirServidor;
						string strImpresora = MenuLatinium.ImpresoraNombre;

						crConnectionInfo.ServerName = strServidor;
						crConnectionInfo.DatabaseName = strBaseDato; 
						crConnectionInfo.UserID = "infoelect"; 
						crConnectionInfo.Password = "Bl45o6$9"; 
					}
			
						TableLogOnInfo crTableLogOnInfo; 
						Database crDatabase = Report1.Database; 
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
						Report1.PrintOptions.PrinterName = MenuLatinium.ImpresoraNombre;
						for (int i=0; i< (int) IRP.txtHasta.Value; i++)
						{
							Report1.PrintToPrinter(1, false, 0, 0);					
						}
						iCont++;	
						#endregion Impresion							
					}

					if( (int)this.optEstado.Value == 1 )
					{
						#region Impresion
						string sReporte = "TablaAmortizacionDatos.rpt";						
						ReportDocument Report1 = new ReportDocument();
						Report1.FileName = MenuLatinium.stDirReportes + sReporte;//"RolesIndividuales.rpt";
								
						ConnectionInfo crConnectionInfo = new ConnectionInfo (); 
					{
						string strBaseDato = MenuLatinium.stDirNumero;
						string strUsuario = MenuLatinium.stDirConta;
						string strServidor = MenuLatinium.stDirServidor;
						string strImpresora = MenuLatinium.ImpresoraNombre;

						crConnectionInfo.ServerName = strServidor;
						crConnectionInfo.DatabaseName = strBaseDato; 
						crConnectionInfo.UserID = "infoelect"; 
						crConnectionInfo.Password = "Bl45o6$9"; 
					}
			
						TableLogOnInfo crTableLogOnInfo; 
						Database crDatabase = Report1.Database; 
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
						Report1.PrintOptions.PrinterName = MenuLatinium.ImpresoraNombre;
						for (int i=0; i< (int) IRP.txtHasta.Value; i++)
						{
							Report1.PrintToPrinter(1, false, 0, 0);					
						}
						iCont++;	
						#endregion Impresion	
					}

					if ((int)this.optEstado.Value == 2)
					{						
						#region Impresion
						string sReporte = "GestiondeCobranzaDatos.rpt";
						//							if (MenuLatinium.stDirFacturacion == "LOTESERVICIO") sReporte = "RolIndividualPeony.rpt";
						ReportDocument Report1 = new ReportDocument();
						Report1.FileName = MenuLatinium.stDirReportes + sReporte;//"RolesIndividuales.rpt";
								
						ConnectionInfo crConnectionInfo = new ConnectionInfo (); 
					{
						string strBaseDato = MenuLatinium.stDirNumero;
						string strUsuario = MenuLatinium.stDirConta;
						string strServidor = MenuLatinium.stDirServidor;
						string strImpresora = MenuLatinium.ImpresoraNombre;

						crConnectionInfo.ServerName = strServidor;
						crConnectionInfo.DatabaseName = strBaseDato; 
						crConnectionInfo.UserID = "infoelect"; 
						crConnectionInfo.Password = "Bl45o6$9"; 
					}
			
						TableLogOnInfo crTableLogOnInfo; 
						Database crDatabase = Report1.Database; 
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
						Report1.PrintOptions.PrinterName = MenuLatinium.ImpresoraNombre;
						for (int i=0; i< (int) IRP.txtHasta.Value; i++)
						{
					
							
							Report1.PrintToPrinter(1, false, 0, 0);
					
						}
						iCont++;	
						#endregion Impresion							
					}

					if ((int)this.optEstado.Value == 5)
					{						
						#region Impresion

						string sReporte = "GestiondeCobranzaSolidarioDatos.rpt";
						//							if (MenuLatinium.stDirFacturacion == "LOTESERVICIO") sReporte = "RolIndividualPeony.rpt";
						ReportDocument Report1 = new ReportDocument();
						Report1.FileName = MenuLatinium.stDirReportes + sReporte;//"RolesIndividuales.rpt";
								
						ConnectionInfo crConnectionInfo = new ConnectionInfo (); 
					{
						string strBaseDato = MenuLatinium.stDirNumero;
						string strUsuario = MenuLatinium.stDirConta;
						string strServidor = MenuLatinium.stDirServidor;
						string strImpresora = MenuLatinium.ImpresoraNombre;

						crConnectionInfo.ServerName = strServidor;
						crConnectionInfo.DatabaseName = strBaseDato; 
						crConnectionInfo.UserID = "infoelect"; 
						crConnectionInfo.Password = "Bl45o6$9"; 
					}
			
						TableLogOnInfo crTableLogOnInfo; 
						Database crDatabase = Report1.Database; 
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
						Report1.PrintOptions.PrinterName = MenuLatinium.ImpresoraNombre;
						for (int i=0; i< (int) IRP.txtHasta.Value; i++)
						{
					
							
							Report1.PrintToPrinter(1, false, 0, 0);
					
						}
						iCont++;	
						#endregion Impresion	
					}

					if ((int)this.optEstado.Value == 3)
					{						
						#region Impresion
						string sReporte = "DatosSeguroDesempleoDatos.rpt";
						//							if (MenuLatinium.stDirFacturacion == "LOTESERVICIO") sReporte = "RolIndividualPeony.rpt";
						ReportDocument Report1 = new ReportDocument();
						Report1.FileName = MenuLatinium.stDirReportes + sReporte;//"RolesIndividuales.rpt";
								
						ConnectionInfo crConnectionInfo = new ConnectionInfo (); 
					{
						string strBaseDato = MenuLatinium.stDirNumero;
						string strUsuario = MenuLatinium.stDirConta;
						string strServidor = MenuLatinium.stDirServidor;
						string strImpresora = MenuLatinium.ImpresoraNombre;

						crConnectionInfo.ServerName = strServidor;
						crConnectionInfo.DatabaseName = strBaseDato; 
						crConnectionInfo.UserID = "infoelect"; 
						crConnectionInfo.Password = "Bl45o6$9"; 
					}
			
						TableLogOnInfo crTableLogOnInfo; 
						Database crDatabase = Report1.Database; 
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
						Report1.PrintOptions.PrinterName = MenuLatinium.ImpresoraNombre;
						for (int i=0; i< (int) IRP.txtHasta.Value; i++)
						{
					
							
							Report1.PrintToPrinter(1, false, 0, 0);
					
						}
						iCont++;	
						#endregion Impresion					
					}
          
					if ((int)this.optEstado.Value == 4)
					{						
						#region Impresion
						string sReporte = "CompromisoLugarDePagoBlancoV2.rpt";
						//							if (MenuLatinium.stDirFacturacion == "LOTESERVICIO") sReporte = "RolIndividualPeony.rpt";
						ReportDocument Report1 = new ReportDocument();
						Report1.FileName = MenuLatinium.stDirReportes + sReporte;//"RolesIndividuales.rpt";
								
						ConnectionInfo crConnectionInfo = new ConnectionInfo (); 
					{
						string strBaseDato = MenuLatinium.stDirNumero;
						string strUsuario = MenuLatinium.stDirConta;
						string strServidor = MenuLatinium.stDirServidor;
						string strImpresora = MenuLatinium.ImpresoraNombre;

						crConnectionInfo.ServerName = strServidor;
						crConnectionInfo.DatabaseName = strBaseDato; 
						crConnectionInfo.UserID = "infoelect"; 
						crConnectionInfo.Password = "Bl45o6$9"; 
					}
			
						TableLogOnInfo crTableLogOnInfo; 
						Database crDatabase = Report1.Database; 
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
						Report1.PrintOptions.PrinterName = MenuLatinium.ImpresoraNombre;
						for (int i=0; i< (int) IRP.txtHasta.Value; i++)
						{
					
							
							Report1.PrintToPrinter(1, false, 0, 0);
					
						}
						iCont++;	
						#endregion Impresion							
					}					
				}
			}
		}

		private void frmImpresiondocumentosDeCredito_Load(object sender, System.EventArgs e)
		{
		}

		private void cdsSeteoF_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			cdsSeteoF.Schema.Connections[0].ConnectionString = miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);	
	
		}
	}
}
