using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using DevExpress.XtraEditors;
using CrystalDecisions;
using CrystalDecisions.CrystalReports.Engine;
using CrystalDecisions.ReportSource;
using CrystalDecisions.Shared;
using System.IO;
using System.Data;
using System.Data.SqlClient;
using System.Diagnostics;


namespace Latinium
{
	/// <summary>
	/// Summary description for DocsPrefactura.
	/// </summary>
	public class DocsPrefactura : DevExpress.XtraEditors.XtraForm
	{
		private Infragistics.Win.UltraWinEditors.UltraCheckEditor chkPagare;
		private Infragistics.Win.UltraWinEditors.UltraCheckEditor chkAmortizaciones;
		private Infragistics.Win.UltraWinEditors.UltraCheckEditor chkGastosDeCobranza;
		public DevExpress.XtraEditors.SimpleButton btGrabar;
		/// <summary>
		/// Required designer variable.
		/// </summary>
		private System.ComponentModel.Container components = null;

		int IdCompra = 0;
		decimal DCapital = 0;
		int ICuotas = 0;
		int IdCliente = 0;
		int IdEntFinanciera = 0;
		public DevExpress.XtraEditors.SimpleButton btCancelar;
		string Fecha = "";
		private C1.Data.C1DataSet cdsSeteoF;
		private System.Windows.Forms.ContextMenu contextMenu1;
		private System.Windows.Forms.MenuItem mnuImpresora;
		private System.Windows.Forms.MenuItem mnuPantalla;
		DateTime FHasta;
		private System.Windows.Forms.GroupBox groupBox1;
		private Infragistics.Win.UltraWinEditors.UltraCheckEditor chkCompromisoDePago;
		private Infragistics.Win.UltraWinEditors.UltraCheckEditor chkActa;
		private Infragistics.Win.UltraWinEditors.UltraCheckEditor chkDatos;
		string sRUC = "";
		string sDocumentoLugarPago = "Cre_CompromisoLugarDePagoActualizado.rpt";
		public DocsPrefactura(int idCompra, decimal dCapital, int iCuotas, int idCliente, string fecha, DateTime fHasta, int idEntFinanciera, string SRUC)
		{
			//
			// Required for Windows Form Designer support
			//
			InitializeComponent();

			IdCompra = idCompra;
			DCapital = dCapital;
			ICuotas = iCuotas;
			IdCliente = idCliente;
			Fecha = fecha;
			FHasta = fHasta;
			IdEntFinanciera = idEntFinanciera;
			sRUC = SRUC;
			//
			// TODO: Add any constructor code after InitializeComponent call
			//
		}

		/// <summary>
		/// Clean up any resources being used.
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

		#region Windows Form Designer generated code
		/// <summary>
		/// Required method for Designer support - do not modify
		/// the contents of this method with the code editor.
		/// </summary>
		private void InitializeComponent()
		{
			Infragistics.Win.Appearance appearance1 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance2 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance3 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance4 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance5 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance6 = new Infragistics.Win.Appearance();
			this.chkPagare = new Infragistics.Win.UltraWinEditors.UltraCheckEditor();
			this.chkAmortizaciones = new Infragistics.Win.UltraWinEditors.UltraCheckEditor();
			this.chkGastosDeCobranza = new Infragistics.Win.UltraWinEditors.UltraCheckEditor();
			this.btGrabar = new DevExpress.XtraEditors.SimpleButton();
			this.btCancelar = new DevExpress.XtraEditors.SimpleButton();
			this.cdsSeteoF = new C1.Data.C1DataSet();
			this.contextMenu1 = new System.Windows.Forms.ContextMenu();
			this.mnuImpresora = new System.Windows.Forms.MenuItem();
			this.mnuPantalla = new System.Windows.Forms.MenuItem();
			this.chkCompromisoDePago = new Infragistics.Win.UltraWinEditors.UltraCheckEditor();
			this.groupBox1 = new System.Windows.Forms.GroupBox();
			this.chkActa = new Infragistics.Win.UltraWinEditors.UltraCheckEditor();
			this.chkDatos = new Infragistics.Win.UltraWinEditors.UltraCheckEditor();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).BeginInit();
			this.SuspendLayout();
			// 
			// chkPagare
			// 
			appearance1.ForeColorDisabled = System.Drawing.Color.Black;
			this.chkPagare.Appearance = appearance1;
			this.chkPagare.BackColor = System.Drawing.Color.Transparent;
			this.chkPagare.Checked = true;
			this.chkPagare.CheckState = System.Windows.Forms.CheckState.Checked;
			this.chkPagare.Location = new System.Drawing.Point(16, 8);
			this.chkPagare.Name = "chkPagare";
			this.chkPagare.Size = new System.Drawing.Size(96, 22);
			this.chkPagare.TabIndex = 144;
			this.chkPagare.Text = "Pagaré";
			// 
			// chkAmortizaciones
			// 
			appearance2.ForeColorDisabled = System.Drawing.Color.Black;
			this.chkAmortizaciones.Appearance = appearance2;
			this.chkAmortizaciones.BackColor = System.Drawing.Color.Transparent;
			this.chkAmortizaciones.Checked = true;
			this.chkAmortizaciones.CheckState = System.Windows.Forms.CheckState.Checked;
			this.chkAmortizaciones.Location = new System.Drawing.Point(16, 48);
			this.chkAmortizaciones.Name = "chkAmortizaciones";
			this.chkAmortizaciones.Size = new System.Drawing.Size(96, 22);
			this.chkAmortizaciones.TabIndex = 145;
			this.chkAmortizaciones.Text = "Amortizaciones";
			this.chkAmortizaciones.CheckedChanged += new System.EventHandler(this.chkAmortizaciones_CheckedChanged);
			// 
			// chkGastosDeCobranza
			// 
			appearance3.ForeColorDisabled = System.Drawing.Color.Black;
			this.chkGastosDeCobranza.Appearance = appearance3;
			this.chkGastosDeCobranza.BackColor = System.Drawing.Color.Transparent;
			this.chkGastosDeCobranza.Checked = true;
			this.chkGastosDeCobranza.CheckState = System.Windows.Forms.CheckState.Checked;
			this.chkGastosDeCobranza.Location = new System.Drawing.Point(16, 88);
			this.chkGastosDeCobranza.Name = "chkGastosDeCobranza";
			this.chkGastosDeCobranza.Size = new System.Drawing.Size(120, 22);
			this.chkGastosDeCobranza.TabIndex = 146;
			this.chkGastosDeCobranza.Text = "Gastos de cobranza";
			// 
			// btGrabar
			// 
			this.btGrabar.Anchor = System.Windows.Forms.AnchorStyles.None;
			this.btGrabar.Location = new System.Drawing.Point(13, 272);
			this.btGrabar.Name = "btGrabar";
			this.btGrabar.Size = new System.Drawing.Size(88, 24);
			this.btGrabar.TabIndex = 168;
			this.btGrabar.Text = "&Imprimir";
			this.btGrabar.Click += new System.EventHandler(this.btGrabar_Click);
			// 
			// btCancelar
			// 
			this.btCancelar.Anchor = System.Windows.Forms.AnchorStyles.None;
			this.btCancelar.Location = new System.Drawing.Point(125, 272);
			this.btCancelar.Name = "btCancelar";
			this.btCancelar.Size = new System.Drawing.Size(88, 24);
			this.btCancelar.TabIndex = 169;
			this.btCancelar.Text = "&Cancelar";
			this.btCancelar.Click += new System.EventHandler(this.btCancelar_Click);
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
			// contextMenu1
			// 
			this.contextMenu1.MenuItems.AddRange(new System.Windows.Forms.MenuItem[] {
																																								 this.mnuImpresora,
																																								 this.mnuPantalla});
			// 
			// mnuImpresora
			// 
			this.mnuImpresora.Index = 0;
			this.mnuImpresora.Text = "Impresora";
			this.mnuImpresora.Click += new System.EventHandler(this.mnuImpresora_Click);
			// 
			// mnuPantalla
			// 
			this.mnuPantalla.Index = 1;
			this.mnuPantalla.Text = "Pantalla";
			this.mnuPantalla.Click += new System.EventHandler(this.mnuPantalla_Click);
			// 
			// chkCompromisoDePago
			// 
			appearance4.ForeColorDisabled = System.Drawing.Color.Black;
			this.chkCompromisoDePago.Appearance = appearance4;
			this.chkCompromisoDePago.BackColor = System.Drawing.Color.Transparent;
			this.chkCompromisoDePago.Checked = true;
			this.chkCompromisoDePago.CheckState = System.Windows.Forms.CheckState.Checked;
			this.chkCompromisoDePago.Location = new System.Drawing.Point(16, 128);
			this.chkCompromisoDePago.Name = "chkCompromisoDePago";
			this.chkCompromisoDePago.Size = new System.Drawing.Size(160, 22);
			this.chkCompromisoDePago.TabIndex = 170;
			this.chkCompromisoDePago.Text = "Compromiso Lugar de Pago";
			// 
			// groupBox1
			// 
			this.groupBox1.Location = new System.Drawing.Point(0, 248);
			this.groupBox1.Name = "groupBox1";
			this.groupBox1.Size = new System.Drawing.Size(224, 8);
			this.groupBox1.TabIndex = 171;
			this.groupBox1.TabStop = false;
			// 
			// chkActa
			// 
			appearance5.ForeColorDisabled = System.Drawing.Color.Black;
			this.chkActa.Appearance = appearance5;
			this.chkActa.BackColor = System.Drawing.Color.Transparent;
			this.chkActa.Checked = true;
			this.chkActa.CheckState = System.Windows.Forms.CheckState.Checked;
			this.chkActa.Location = new System.Drawing.Point(16, 168);
			this.chkActa.Name = "chkActa";
			this.chkActa.Size = new System.Drawing.Size(160, 22);
			this.chkActa.TabIndex = 172;
			this.chkActa.Text = "Acta Entrega Recepción de Documentos";
			// 
			// chkDatos
			// 
			appearance6.ForeColorDisabled = System.Drawing.Color.Black;
			this.chkDatos.Appearance = appearance6;
			this.chkDatos.BackColor = System.Drawing.Color.Transparent;
			this.chkDatos.Checked = true;
			this.chkDatos.CheckState = System.Windows.Forms.CheckState.Checked;
			this.chkDatos.Location = new System.Drawing.Point(16, 216);
			this.chkDatos.Name = "chkDatos";
			this.chkDatos.Size = new System.Drawing.Size(160, 22);
			this.chkDatos.TabIndex = 173;
			this.chkDatos.Text = "Autorización de uso de Datos Personales";
			// 
			// DocsPrefactura
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 14);
			this.ClientSize = new System.Drawing.Size(226, 304);
			this.ControlBox = false;
			this.Controls.Add(this.chkDatos);
			this.Controls.Add(this.chkActa);
			this.Controls.Add(this.groupBox1);
			this.Controls.Add(this.chkCompromisoDePago);
			this.Controls.Add(this.btCancelar);
			this.Controls.Add(this.btGrabar);
			this.Controls.Add(this.chkGastosDeCobranza);
			this.Controls.Add(this.chkAmortizaciones);
			this.Controls.Add(this.chkPagare);
			this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle;
			this.MaximizeBox = false;
			this.MinimizeBox = false;
			this.Name = "DocsPrefactura";
			this.ShowInTaskbar = false;
			this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
			this.Text = "Documentos";
			this.KeyDown += new System.Windows.Forms.KeyEventHandler(this.DocsPrefactura_KeyDown);
			this.Load += new System.EventHandler(this.DocsPrefactura_Load);
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).EndInit();
			this.ResumeLayout(false);

		}
		#endregion

		private void btGrabar_Click(object sender, System.EventArgs e)
		{
			this.contextMenu1.Show(this.btGrabar, new Point(1, 25));
		}

		private void DocsPrefactura_Load(object sender, System.EventArgs e)
		{
			//int idTipoFactura = Funcion.iEscalarSQL(cdsSeteoF, string.Format("Select idTipoFactura From Compra Where idCompra = {0}", IdCompra));

//			int idEntidadFinanciera = Funcion.iEscalarSQL(cdsSeteoF, string.Format("Exec RetornaIdEFSolCredito '{0}'", sRUC));
//			
//			Funcion.EjecutaSQL(cdsSeteoF, string.Format("Update Compra Set idEntidadFinanciera = {0} Where idCompra = {1}", idEntidadFinanciera, IdCompra));
			string carpeta_1 ="";
			carpeta_1 = @"C:\Latinium\Reportes\Cre_Consentimiento_1.rpt";
			if(IdEntFinanciera==14) 
			{
					this.chkGastosDeCobranza.Checked=false;
				  this.chkGastosDeCobranza.Enabled=false;
			}

		}

		private void btCancelar_Click(object sender, System.EventArgs e)
		{
			this.Close();
		}

		private void cdsSeteoF_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			cdsSeteoF.Schema.Connections[0].ConnectionString = miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);		
		}

		private void mnuImpresora_Click(object sender, System.EventArgs e)
		{
			this.Cursor = Cursors.WaitCursor;

			int iReporte = IdEntFinanciera;

			if (this.chkPagare.Checked)
			{
				ReportDocument Report = new ReportDocument();
				
				string sPagare = "Pagare.rpt";
				if (iReporte == 2) sPagare = "Informe11.rpt";

				Report.FileName = MenuLatinium.stDirReportes + sPagare;
								
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

				Report.SetParameterValue("@idCompra", IdCompra);
				Report.SetParameterValue("@FechaGenera", FHasta);				
				Report.PrintToPrinter(1, false, 0, 0);
			}

			if (this.chkAmortizaciones.Checked)
			{
				ReportDocument Report1 = new ReportDocument();
				Report1.FileName = MenuLatinium.stDirReportes + "Informe12.rpt";
								
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

				Report1.SetParameterValue("@Capital", DCapital);
				Report1.SetParameterValue("@Cuotas", ICuotas);
				Report1.SetParameterValue("@IdCliente", IdCliente);
				Report1.SetParameterValue("@Fecha", Fecha);
				Report1.SetParameterValue("@FechaGenera", FHasta);
				Report1.SetParameterValue("@idEntidadFinanciera", IdEntFinanciera);
				
				Report1.PrintToPrinter(1, false, 0, 0);
			}

			if (this.chkGastosDeCobranza.Checked)
			{
				string sReporte = "";
				
				// Funcion.iEscalarSQL(cdsSeteoF, string.Format("Select dbo.TipoDeEntidad({0})", IdCompra));

				if (iReporte == 1) sReporte = "GestionDeCobranza.rpt";
				else if (iReporte == 4) sReporte = "GestionDeCobranza.rpt";
				else if (iReporte == 2) sReporte = "GestionDeCobranzaSolidario.rpt";
				else if (iReporte == 15) sReporte = "GestionDeCobranza.rpt";
				else if (iReporte == 16) sReporte = "GestionDeCobranza.rpt";

				ReportDocument Report2 = new ReportDocument();
				Report2.FileName = MenuLatinium.stDirReportes + sReporte;
								
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
				Database crDatabase = Report2.Database; 
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
				
				Report2.SetParameterValue("@idCompra", IdCompra);
				//if (iReporte != 2) 
					Report2.SetParameterValue("@Refinanciacion", 0);
				Report2.PrintToPrinter(1, false, 0, 0);
			}
			
			#region compromiso de pago
			if (this.chkCompromisoDePago.Checked)
			{
				ReportDocument Report3 = new ReportDocument();
				//				Report3.FileName = MenuLatinium.stDirReportes + "Cre_CompromisoLugarDePago.rpt";
				Report3.FileName = MenuLatinium.stDirReportes + "Cre_CompromisoLugarDePagoActualizado.rpt";			
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
				Database crDatabase = Report3.Database; 
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
				
				Report3.SetParameterValue("@idCompra", IdCompra);
				Report3.PrintToPrinter(1, false, 0, 0);
			}
			#endregion compromiso de pago

			#region documentos
			if (this.chkActa.Checked)
			{
				ReportDocument Report3 = new ReportDocument();
				//				Report3.FileName = MenuLatinium.stDirReportes + "Cre_CompromisoLugarDePago.rpt";
				Report3.FileName = MenuLatinium.stDirReportes + "Cre_ActaDeEntrega.rpt";			
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
				Database crDatabase = Report3.Database; 
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
				
				Report3.SetParameterValue("@idCompra", IdCompra);
				Report3.PrintToPrinter(1, false, 0, 0);
			}
			#endregion documentos

			string sArchivo = @"C:\Latinium\Reportes\Cre_Consentimiento_1.rpt";

			if (File.Exists(sArchivo))
			{	

				#region datos
				if (this.chkDatos.Checked)
				{
					ReportDocument Report3 = new ReportDocument();
					//				Report3.FileName = MenuLatinium.stDirReportes + "Cre_CompromisoLugarDePago.rpt";
					Report3.FileName = MenuLatinium.stDirReportes + "Cre_Consentimiento_1.rpt";			
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
					Database crDatabase = Report3.Database; 
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
				
					Report3.SetParameterValue("@idCompra", IdCompra);
					Report3.PrintToPrinter(1, false, 0, 0);
				}
				#endregion datos
			}


			Funcion.EjecutaSQL(cdsSeteoF, string.Format("Update Compra Set Impreso = Case idTipoFactura When 43 Then 1 When 1 Then Impreso End Where idCompra = {0}", IdCompra), true);

			this.btCancelar.Text = "&Cerrar";			
//			this.btGrabar.Enabled = false;

			this.Cursor = Cursors.Default;
		}

		private void mnuPantalla_Click(object sender, System.EventArgs e)
		{
			this.Cursor = Cursors.WaitCursor;

			int iReporte = IdEntFinanciera;

			if (this.chkPagare.Checked)
			{				
				string sPagare = "Pagare.rpt";
				if (iReporte == 2) sPagare = "Informe11.rpt";

				ParameterFields paramFields = new ParameterFields ();
				
				ParameterField paramFieldCompra = new ParameterField ();
				ParameterDiscreteValue ValIdCompra = new ParameterDiscreteValue ();
				paramFieldCompra.ParameterFieldName = "@idCompra";
				ValIdCompra.Value = int.Parse(IdCompra.ToString());
				paramFieldCompra.CurrentValues.Add (ValIdCompra);
				paramFields.Add (paramFieldCompra);

				ParameterField paramFieldFecha = new ParameterField ();
				ParameterDiscreteValue ValIdFecha = new ParameterDiscreteValue ();
				paramFieldFecha.ParameterFieldName = "@FechaGenera";
				ValIdFecha.Value = FHasta;
				paramFieldFecha.CurrentValues.Add (ValIdFecha);
				paramFields.Add (paramFieldFecha);
				
				Reporte miRep = new Reporte(sPagare, "");
				miRep.MdiParent = this.MdiParent;
				miRep.Titulo("Págare");
				miRep.crVista.ParameterFieldInfo = paramFields;
				miRep.ShowDialog();
			}

			if (this.chkAmortizaciones.Checked)
			{
				ParameterFields paramFields = new ParameterFields ();
				
				ParameterField pfCapital = new ParameterField ();
				ParameterDiscreteValue ValCapital = new ParameterDiscreteValue ();
				pfCapital.ParameterFieldName = "@Capital";
				ValCapital.Value = DCapital;
				pfCapital.CurrentValues.Add (ValCapital);
				paramFields.Add (pfCapital);
				
				ParameterField pfCuotas = new ParameterField ();
				ParameterDiscreteValue ValCuotas = new ParameterDiscreteValue ();
				pfCuotas.ParameterFieldName = "@Cuotas";
				ValCuotas.Value = ICuotas;
				pfCuotas.CurrentValues.Add (ValCuotas);
				paramFields.Add (pfCuotas);
				
				ParameterField pfIdCliente = new ParameterField ();
				ParameterDiscreteValue ValIdCliente = new ParameterDiscreteValue ();
				pfIdCliente.ParameterFieldName = "@IdCliente";
				ValIdCliente.Value = IdCliente;
				pfIdCliente.CurrentValues.Add (ValIdCliente);
				paramFields.Add (pfIdCliente);
				
				ParameterField pfFecha = new ParameterField ();
				ParameterDiscreteValue ValFecha = new ParameterDiscreteValue ();
				pfFecha.ParameterFieldName = "@Fecha";
				ValFecha.Value = Fecha;
				pfFecha.CurrentValues.Add (ValFecha);
				paramFields.Add (pfFecha);
				
				ParameterField pfFechaGenera = new ParameterField ();
				ParameterDiscreteValue ValFechaGenera = new ParameterDiscreteValue ();
				pfFechaGenera.ParameterFieldName = "@FechaGenera";
				ValFechaGenera.Value = FHasta;
				pfFechaGenera.CurrentValues.Add (ValFechaGenera);
				paramFields.Add (pfFechaGenera);

				ParameterField pfEF = new ParameterField ();
				ParameterDiscreteValue ValEF = new ParameterDiscreteValue ();
				pfEF.ParameterFieldName = "@idEntidadFinanciera";
				ValEF.Value = IdEntFinanciera;
				pfEF.CurrentValues.Add (ValEF);
				paramFields.Add (pfEF);
				
				Reporte miRepor = new Reporte("Informe12.rpt", "");
				miRepor.MdiParent = this.MdiParent;
				miRepor.Titulo("Págare");
				miRepor.crVista.ParameterFieldInfo = paramFields;
				miRepor.ShowDialog();
			}

			if (this.chkGastosDeCobranza.Checked)
			{
				string sReporte = "";
				
				// Funcion.iEscalarSQL(cdsSeteoF, string.Format("Select dbo.TipoDeEntidad({0})", IdCompra));

				if (iReporte == 1) sReporte = "GestionDeCobranza.rpt";
				if (iReporte == 4) sReporte = "GestionDeCobranza.rpt";
				if (iReporte == 15) sReporte = "GestionDeCobranza.rpt";
				if (iReporte == 16) sReporte = "GestionDeCobranza.rpt";
				else if (iReporte == 2) sReporte = "GestionDeCobranzaSolidario.rpt";

				ParameterFields paramFields = new ParameterFields ();
				
				ParameterField paramFieldCompra = new ParameterField ();
				ParameterDiscreteValue ValIdCompra = new ParameterDiscreteValue ();
				paramFieldCompra.ParameterFieldName = "@idCompra";
				ValIdCompra.Value = int.Parse(IdCompra.ToString());
				paramFieldCompra.CurrentValues.Add (ValIdCompra);
				paramFields.Add (paramFieldCompra);

				int bRef = 0;
				ParameterField paramFieldRef = new ParameterField ();
				ParameterDiscreteValue ValIdRef = new ParameterDiscreteValue ();
				paramFieldRef.ParameterFieldName = "@Refinanciacion";
				ValIdRef.Value = int.Parse(bRef.ToString());;
				paramFieldRef.CurrentValues.Add (ValIdRef);
				paramFields.Add (paramFieldRef);

				Reporte miReporte = new Reporte(sReporte, "");
				miReporte.MdiParent = MdiParent;
				miReporte.Titulo("Gastos de Cobranza");
				miReporte.crVista.ParameterFieldInfo = paramFields;
				miReporte.ShowDialog();
			}	

			if (this.chkCompromisoDePago.Checked)
			{					
				ParameterFields paramFields = new ParameterFields ();
				
				ParameterField paramFieldCompra = new ParameterField ();
				ParameterDiscreteValue ValIdCompra = new ParameterDiscreteValue ();
				paramFieldCompra.ParameterFieldName = "@idCompra";
				ValIdCompra.Value = int.Parse(IdCompra.ToString());
				paramFieldCompra.CurrentValues.Add (ValIdCompra);
				paramFields.Add (paramFieldCompra);		
				
				
				Reporte miRep = new Reporte("Cre_CompromisoLugarDePagoActualizado.rpt", "");
//				Reporte miRep = new Reporte("Cre_CompromisoLugarDePago.rpt", "");
				miRep.MdiParent = this.MdiParent;
				miRep.Titulo("Compromiso Lugar de Pago");
				miRep.crVista.ParameterFieldInfo = paramFields;
				miRep.ShowDialog();
			}

			#region Acta
			if (this.chkActa.Checked)
			{					
				ParameterFields paramFields = new ParameterFields ();
				
				ParameterField paramFieldCompra = new ParameterField ();
				ParameterDiscreteValue ValIdCompra = new ParameterDiscreteValue ();
				paramFieldCompra.ParameterFieldName = "@idCompra";
				ValIdCompra.Value = int.Parse(IdCompra.ToString());
				paramFieldCompra.CurrentValues.Add (ValIdCompra);
				paramFields.Add (paramFieldCompra);		
				
				
				Reporte miRep = new Reporte("Cre_ActaDeEntrega.rpt", "");
				//				Reporte miRep = new Reporte("Cre_CompromisoLugarDePago.rpt", "");
				miRep.MdiParent = this.MdiParent;
				miRep.Titulo("Acta Entrega Recepción de Documentos");
				miRep.crVista.ParameterFieldInfo = paramFields;
				miRep.ShowDialog();
			}

			
			#endregion Acta

			string sArchivo = @"C:\Latinium\Reportes\Cre_Consentimiento_1.rpt";

			if (File.Exists(sArchivo))
			{	
				#region Datos
				if (this.chkDatos.Checked)
				{					
					ParameterFields paramFields = new ParameterFields ();
				
					ParameterField paramFieldCompra = new ParameterField ();
					ParameterDiscreteValue ValIdCompra = new ParameterDiscreteValue ();
					paramFieldCompra.ParameterFieldName = "@idCompra";
					ValIdCompra.Value = int.Parse(IdCompra.ToString());
					paramFieldCompra.CurrentValues.Add (ValIdCompra);
					paramFields.Add (paramFieldCompra);		
				
				
					Reporte miRep = new Reporte("Cre_Consentimiento_1.rpt", "");
					//				Reporte miRep = new Reporte("Cre_CompromisoLugarDePago.rpt", "");
					miRep.MdiParent = this.MdiParent;
					miRep.Titulo("Autorización de uso de Datos Personales");
					miRep.crVista.ParameterFieldInfo = paramFields;
					miRep.ShowDialog();
				}

			
				#endregion Datos
			}
			
			Funcion.EjecutaSQL(cdsSeteoF, string.Format("Update Compra Set Impreso = Case idTipoFactura When 43 Then 1 When 1 Then Impreso End Where idCompra = {0}", IdCompra), true);

			this.btCancelar.Text = "&Cerrar";			
			//this.btGrabar.Enabled = false;
			this.Close();
			this.Cursor = Cursors.Default;
		}

		private void DocsPrefactura_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Escape) this.Close();
		}

		private void chkAmortizaciones_CheckedChanged(object sender, System.EventArgs e)
		{
		
		}
	}
}

