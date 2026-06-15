using System;
using System.IO;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using CrystalDecisions.CrystalReports.Engine;
using CrystalDecisions.ReportSource;
using CrystalDecisions.Shared;
using System.Drawing.Printing;

namespace Latinium
{
	/// <summary>
	/// Descripción breve de Reportes.
	/// </summary>
	public class Reporte : DevExpress.XtraEditors.XtraForm
	{
		public CrystalDecisions.Windows.Forms.CrystalReportViewer crVista;
		/// <summary>
		/// Variable del diseñador requerida.
		/// </summary>
		private System.ComponentModel.Container components = null;
		public ReportDocument oRpt;
		private Infragistics.Win.Misc.UltraButton btSalir;
		private Infragistics.Win.Misc.UltraButton btImprime;
		private Infragistics.Win.UltraWinEditors.UltraComboEditor cmbSalida;
		private Infragistics.Win.Misc.UltraButton btExcel;

		private int iTipoPapel = 3;
		/// <summary>
		/// Define tamano de papel
		/// </summary>
		/// <param name="iTipo"> 1 A4, 2 Carta, 3 Default</param>
		public void TipoPapel(int iTipo)
		{
			iTipoPapel = iTipo;
		}

		private void ReporteNombre(string strReporte)
		{
			string strBase = MenuLatinium.stDirReportes;

			if (strBase.Trim().ToUpper().EndsWith("DEBUG") || strBase.Trim().ToUpper().EndsWith("RELEASE")) 
			{
				strBase = MenuLatinium.stDirInicio + @"\Reportes";
				strReporte = MenuLatinium.stDirInicio + @"\Reportes\" + strReporte;
			}
			else
				strReporte = strBase + strReporte;			
			oRpt = new ReportDocument();
			
			if (!strReporte.ToUpper().EndsWith(".RPT")) strReporte += ".rpt";
			if (!File.Exists(@strReporte))
			{
				MessageBox.Show("Archivo no existe: " + strReporte);
				return;
			}

			try
			{
				oRpt.Load(strReporte);
			}
			catch(System.Exception ex)
			{ 
				MessageBox.Show(ex.Message, "Carga de Reporte: " + strReporte);
				return;
			}
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
			Database crDatabase = oRpt.Database; 
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
			this.crVista.ReportSource = oRpt;
		}

		private void ReporteOrden(string strOrden)
		{
			FieldDefinition FieldDef;
			if (strOrden.Length > 0) 
			{
				string[] stOr;
				int i=2;
				stOr = strOrden.Split(new char[]{','},5);
				foreach(string st in stOr) 
				{
					try
					{
						FieldDef = oRpt.Database.Tables[0].Fields[st];
						oRpt.DataDefinition.SortFields[i].Field = FieldDef;
						i++;
					}
					catch(Exception e)
					{
						MessageBox.Show("Orden "+e.Message);
					}
				}
			}
		}
		private void ReporteGrupo(string strGrupo, int iTabla)
		{
				try
			{
					if (strGrupo.Length > 0) 
					{
						FieldDefinition FieldDef;
						FieldDef = oRpt.Database.Tables[iTabla].Fields [strGrupo];
						oRpt.DataDefinition.Groups[0].ConditionField = FieldDef;
					}
			}
			catch(Exception e)
			{
				MessageBox.Show("Grupo "+e.Message);
			}
		}


		public Reporte(string strReporte, string strFiltro, string strOrden, string strGrupo, int iTabla)
		{
			InitializeComponent();
			ReporteNombre(strReporte);			
			this.crVista.SelectionFormula = strFiltro;
			ReporteGrupo(strGrupo, iTabla);
			ReporteOrden(strOrden);
		}
		public Reporte(string strReporte, string strFiltro, string strOrden)
		{
			InitializeComponent();
			ReporteNombre(strReporte);			
			this.crVista.SelectionFormula = strFiltro;
			ReporteOrden(strOrden);
		}

		string sNombrePapel = "";
		bool bNombrePapel = false;

		public Reporte(string strReporte, string strFiltro, string SNombrePapel, bool BNombrePapel)
		{
			InitializeComponent();

			ReporteNombre(strReporte);
			this.crVista.SelectionFormula = strFiltro;
			sNombrePapel = SNombrePapel;
			bNombrePapel = BNombrePapel;
		}

		public Reporte(string strReporte, string strFiltro)
		{
			InitializeComponent();

			ReporteNombre(strReporte);
			this.crVista.SelectionFormula = strFiltro;
		}

		public Reporte(string strReporte)
		{
			InitializeComponent();
			ReporteNombre(strReporte);
		}

		public Reporte()
		{
			InitializeComponent();
			oRpt = new ReportDocument();
			
			oRpt.Load(this.crVista.ReportSource.ToString());   
			this.crVista.ReportSource = oRpt;
//			oRpt.PrintToPrinter(1, false, 0, 0);
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

		#region Windows Form Designer generated code
		/// <summary>
		/// Método necesario para admitir el Diseñador, no se puede modificar
		/// el contenido del método con el editor de código.
		/// </summary>
		private void InitializeComponent()
		{
			Infragistics.Win.Appearance appearance1 = new Infragistics.Win.Appearance();
			System.Resources.ResourceManager resources = new System.Resources.ResourceManager(typeof(Reporte));
			Infragistics.Win.Appearance appearance2 = new Infragistics.Win.Appearance();
			Infragistics.Win.ValueListItem valueListItem1 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.ValueListItem valueListItem2 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.ValueListItem valueListItem3 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.ValueListItem valueListItem4 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.ValueListItem valueListItem5 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.ValueListItem valueListItem6 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.ValueListItem valueListItem7 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.ValueListItem valueListItem8 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.Appearance appearance3 = new Infragistics.Win.Appearance();
			this.crVista = new CrystalDecisions.Windows.Forms.CrystalReportViewer();
			this.btSalir = new Infragistics.Win.Misc.UltraButton();
			this.btImprime = new Infragistics.Win.Misc.UltraButton();
			this.cmbSalida = new Infragistics.Win.UltraWinEditors.UltraComboEditor();
			this.btExcel = new Infragistics.Win.Misc.UltraButton();
			((System.ComponentModel.ISupportInitialize)(this.cmbSalida)).BeginInit();
			this.SuspendLayout();
			// 
			// crVista
			// 
			this.crVista.ActiveViewIndex = -1;
			this.crVista.DisplayBackgroundEdge = false;
			this.crVista.DisplayGroupTree = false;
			this.crVista.Dock = System.Windows.Forms.DockStyle.Fill;
			this.crVista.EnableDrillDown = false;
			this.crVista.Location = new System.Drawing.Point(0, 0);
			this.crVista.Name = "crVista";
			this.crVista.ReportSource = "C:\\corpam\\CronosH\\Reportes\\Socios.rpt";
			this.crVista.ShowRefreshButton = false;
			this.crVista.Size = new System.Drawing.Size(776, 414);
			this.crVista.TabIndex = 0;
			this.crVista.Load += new System.EventHandler(this.crVista_Load);
			// 
			// btSalir
			// 
			appearance1.Image = ((object)(resources.GetObject("appearance1.Image")));
			this.btSalir.Appearance = appearance1;
			this.btSalir.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			this.btSalir.DialogResult = System.Windows.Forms.DialogResult.Cancel;
			this.btSalir.Location = new System.Drawing.Point(568, 0);
			this.btSalir.Name = "btSalir";
			this.btSalir.Size = new System.Drawing.Size(56, 25);
			this.btSalir.TabIndex = 1;
			this.btSalir.Text = "&Salir";
			this.btSalir.Click += new System.EventHandler(this.btSalir_Click);
			// 
			// btImprime
			// 
			appearance2.Image = ((object)(resources.GetObject("appearance2.Image")));
			this.btImprime.Appearance = appearance2;
			this.btImprime.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			this.btImprime.DialogResult = System.Windows.Forms.DialogResult.Cancel;
			this.btImprime.Location = new System.Drawing.Point(480, 0);
			this.btImprime.Name = "btImprime";
			this.btImprime.Size = new System.Drawing.Size(24, 25);
			this.btImprime.TabIndex = 2;
			this.btImprime.Text = "&Imprime";
			this.btImprime.Visible = false;
			this.btImprime.Click += new System.EventHandler(this.btImprime_Click);
			this.btImprime.MouseUp += new System.Windows.Forms.MouseEventHandler(this.btImprime_MouseUp);
			// 
			// cmbSalida
			// 
			this.cmbSalida.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbSalida.DropDownStyle = Infragistics.Win.DropDownStyle.DropDownList;
			valueListItem1.DataValue = 0;
			valueListItem1.DisplayText = "Imp";
			valueListItem2.DataValue = 1;
			valueListItem2.DisplayText = "Xls";
			valueListItem3.DataValue = 2;
			valueListItem3.DisplayText = "Xls1";
			valueListItem4.DataValue = 3;
			valueListItem4.DisplayText = "Pdf";
			valueListItem5.DataValue = 4;
			valueListItem5.DisplayText = "Rtf";
			valueListItem6.DataValue = 5;
			valueListItem6.DisplayText = "Html";
			valueListItem7.DataValue = 6;
			valueListItem7.DisplayText = "Doc";
			valueListItem8.DataValue = 7;
			valueListItem8.DisplayText = "Text";
			this.cmbSalida.Items.Add(valueListItem1);
			this.cmbSalida.Items.Add(valueListItem2);
			this.cmbSalida.Items.Add(valueListItem3);
			this.cmbSalida.Items.Add(valueListItem4);
			this.cmbSalida.Items.Add(valueListItem5);
			this.cmbSalida.Items.Add(valueListItem6);
			this.cmbSalida.Items.Add(valueListItem7);
			this.cmbSalida.Items.Add(valueListItem8);
			this.cmbSalida.Location = new System.Drawing.Point(424, 0);
			this.cmbSalida.Name = "cmbSalida";
			this.cmbSalida.Size = new System.Drawing.Size(56, 22);
			this.cmbSalida.TabIndex = 3;
			// 
			// btExcel
			// 
			appearance3.Image = ((object)(resources.GetObject("appearance3.Image")));
			this.btExcel.Appearance = appearance3;
			this.btExcel.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			this.btExcel.DialogResult = System.Windows.Forms.DialogResult.Cancel;
			this.btExcel.Location = new System.Drawing.Point(504, 0);
			this.btExcel.Name = "btExcel";
			this.btExcel.Size = new System.Drawing.Size(56, 25);
			this.btExcel.TabIndex = 4;
			this.btExcel.Text = "&Excel";
			this.btExcel.Click += new System.EventHandler(this.btExcel_Click);
			this.btExcel.MouseUp += new System.Windows.Forms.MouseEventHandler(this.btExcel_MouseUp);
			// 
			// Reporte
			// 
			this.AcceptButton = this.btImprime;
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 14);
			this.CancelButton = this.btSalir;
			this.ClientSize = new System.Drawing.Size(776, 414);
			this.Controls.Add(this.btExcel);
			this.Controls.Add(this.cmbSalida);
			this.Controls.Add(this.btImprime);
			this.Controls.Add(this.btSalir);
			this.Controls.Add(this.crVista);
			this.Name = "Reporte";
			this.ShowInTaskbar = false;
			this.Text = "Reportes";
			this.Closing += new System.ComponentModel.CancelEventHandler(this.Reporte_Closing);
			this.Load += new System.EventHandler(this.Reporte_Load);
			((System.ComponentModel.ISupportInitialize)(this.cmbSalida)).EndInit();
			this.ResumeLayout(false);

		}
		#endregion
		public void Periodo(string strPeriodo)
		{
			Section section;
			TextObject textObject=null;

			try
			{
				section = oRpt.ReportDefinition.Sections["ReportHeaderSection1"];
				textObject = section.ReportObjects["pcPeriodo"] as TextObject;
			}
			catch
			{
				try
				{
					section = oRpt.ReportDefinition.Sections["PageHeaderSection1"];
					textObject = section.ReportObjects["pcPeriodo"] as TextObject;
				}
				catch(System.Exception ex)
				{
					Console.WriteLine(ex.Message + ". Funcion: Periodo");
				}
			}

			if (textObject!= null)
			{
				textObject.Text = strPeriodo;
//				textObject.Color = Color.Blue;
			}
		}

		public void Titulo(string strTitulo)
		{
			Section section;
			TextObject textObject=null;

			try
			{
				section = oRpt.ReportDefinition.Sections["ReportHeaderSection1"];
				textObject = section.ReportObjects["pcTitulo"] as TextObject;
			}
			catch
			{
				try
				{
					section = oRpt.ReportDefinition.Sections["PageHeaderSection1"];
					textObject = section.ReportObjects["pcTitulo"] as TextObject;
				}
				catch(System.Exception ex)
				{
					Console.WriteLine(ex.Message + ". Funcion: Titulo");
				}
			}

			if (textObject!= null)
			{
				textObject.Text = strTitulo;
			}
		}

		public void Resumen()
		{
			try
			{
				Section section;

				section = oRpt.ReportDefinition.Sections["DetailSection1"];
				section.SectionFormat.EnableSuppress = true;
			}
			catch{}
		}
		public void Resumen(int i)
		{
			Section section;

			try
			{
				section = oRpt.ReportDefinition.Sections["GroupFooterSection1"];
				section.SectionFormat.EnableSuppress = true;
			}
			catch
			{
				try
				{
					section = oRpt.ReportDefinition.Sections["DetailSection1"];
					section.SectionFormat.EnableSuppress = true;
				}
				catch{}
			}
		}

		public void SuprimeNivel(int iNivel)
		{
			Section section;

			section = oRpt.ReportDefinition.Sections[iNivel];
			section.SectionFormat.EnableSuppress = true;
		}

		public void Refresca()
		{
			this.crVista.RefreshReport();
		}

		public int iNumeroCopias = 1;

		public void ExportaPDFPedido(string sNombre)
		{
			try
			{
				// Define the full path where the file will be saved
				string fullPath = @"\\192.168.1.10\Proyectos\Latinium\" + sNombre;
        
				// Get the directory from the full path
				string directory = Path.GetDirectoryName(fullPath);

				// Check if the directory exists
				if (!Directory.Exists(directory))
				{
					// Create the directory if it does not exist
					Directory.CreateDirectory(directory);
				}

				// Set the record selection formula (assumed to be previously set)
				oRpt.RecordSelectionFormula = crVista.SelectionFormula;

				// Export the report to PDF
				oRpt.ExportToDisk(ExportFormatType.PortableDocFormat, fullPath);

				// Clean up
				oRpt = null;
				crVista.Dispose();
				crVista = null;
			}
			catch (Exception ex)
			{
				// Handle exceptions (e.g., log the error, rethrow, show a message to the user, etc.)
				Console.WriteLine("Error exporting PDF: " + ex.Message);
				// Optionally, rethrow or handle the exception as needed
				throw;
			}
		}


		public void ExportaPDF(string sNombre)
		{
			try
			{
				// Define the full path where the file will be saved
				string fullPath = @"\\192.168.1.10\Proyectos\Latinium\" + sNombre;
        
				// Get the directory from the full path
				string directory = Path.GetDirectoryName(fullPath);

				// Check if the directory exists
				if (!Directory.Exists(directory))
				{
					// Create the directory if it does not exist
					Directory.CreateDirectory(directory);
				}

				// Set the record selection formula (assumed to be previously set)
				oRpt.RecordSelectionFormula = crVista.SelectionFormula;

				// Export the report to PDF
				oRpt.ExportToDisk(ExportFormatType.PortableDocFormat, fullPath);

				// Clean up
				oRpt = null;
				crVista.Dispose();
				crVista = null;
			}
			catch (Exception ex)
			{
				// Handle exceptions (e.g., log the error, rethrow, show a message to the user, etc.)
				Console.WriteLine("Error exporting PDF: " + ex.Message);
				// Optionally, rethrow or handle the exception as needed
				throw;
			}
		}

		public void ExportaRoles(string sNombre, string sCarpeta)
		{		
			oRpt.RecordSelectionFormula = crVista.SelectionFormula;
			oRpt.ExportToDisk(CrystalDecisions.Shared.ExportFormatType.PortableDocFormat, @""+sCarpeta+@"\" + sNombre);// @"C:\Latinium\OrcCmp" + sNombre); // @"\\192.168.1.186\Latinium\OrdenesDeCompra" + sNombre); //MenuLatinium.stDirInicio + sNombre);
			oRpt = null;
			crVista.Dispose();
			crVista = null;
		}

		public void ExportaPDF1(string sNombre)
		{		
			oRpt.RecordSelectionFormula = crVista.SelectionFormula;
			oRpt.ExportToDisk(CrystalDecisions.Shared.ExportFormatType.PortableDocFormat, @"\\192.168.1.186\Proyectos\" + sNombre);// @"C:\Latinium\OrcCmp" + sNombre); // @"\\192.168.1.186\Latinium\OrdenesDeCompra" + sNombre); //MenuLatinium.stDirInicio + sNombre);
			oRpt = null;
			crVista.Dispose();
			crVista = null;
		}
	/*	
		public int GetPapersizeID(string PrinterName, string PaperSizeName) 
		{
			System.Drawing.Printing.PrintDocument pdprint = new System.Drawing.Printing.PrintDocument();
			int PaperSizeID = 0;
			string ppname = "";
			pdprint.PrinterSettings.PrinterName = PrinterName;
			for(int i= 0; i< pdprint.PrinterSettings.PaperSizes.Count; i++)
			{
				int rawKind;
				ppname = PaperSizeName;
				if(pdprint.PrinterSettings.PaperSizes[i].PaperName == ppname)
				{
					rawKind = Convert.ToInt32(pdprint.PrinterSettings.PaperSizes[i].RawKind);					
					PaperSizeID = rawKind;
					break;
				}
			}
			return PaperSizeID;
		}
*/
		public void ImprimeRep(int iImp)
		{
			try
			{
				if (iImp == 1)
				{
					oRpt.RecordSelectionFormula = crVista.SelectionFormula;
					string stImpresora = MenuLatinium.ImpresoraNombre;
//					MessageBox.Show(stImpresora);
//					if (Funcion.iEscalarSQL(cds, "Select Count(*) From SeteoG Where Codigo = 'FPVISEC' And Activado = 1") > 0)
//						stImpresora = Funcion.sEscalarSQL("Select Nombre From SeteoG Where Codigo = 'FPVISEC' And Activado = 1");
																		
					oRpt.PrintToPrinter(iNumeroCopias, false, 1, 1);
					oRpt.Dispose();
					oRpt = null;
					crVista.Dispose();
					crVista = null;					
				}
				else if (iImp == 2)
				{
					oRpt.RecordSelectionFormula= crVista.SelectionFormula;
					oRpt.ExportToDisk(CrystalDecisions.Shared.ExportFormatType.PortableDocFormat, "Correo.pdf");
					oRpt = null;
					crVista.Dispose();
					crVista = null;
				}
				else
				{
					this.crVista.RefreshReport();
					this.crVista.PrintReport();
//					this.crVista.PrintReport();
				}
			}
			catch (System.Exception ex)
			{
				MessageBox.Show(ex.Message, "Impresion Directa");
			}
		}

		public void ImprimeRep()
		{
			this.crVista.RefreshReport();
			this.crVista.PrintReport();
		}
		private Acceso miAcceso;

		private void crVista_Load(object sender, System.EventArgs e)
		{
		
		}

		private void Reporte_Load(object sender, System.EventArgs e)
		{

//			miAcceso = new Acceso();
//			if (miAcceso.BExportar) this.crVista.ShowExportButton= true;

			cmbSalida.Value = 0;
		}

		private void Reporte_Closing(object sender, System.ComponentModel.CancelEventArgs e)
		{
			if (oRpt != null)
			{
				oRpt.Dispose();
				oRpt = null;
			}
			if (crVista != null)
			{
				crVista.Dispose();
				crVista = null;
			}
		}

		private void btSalir_Click(object sender, System.EventArgs e)
		{
			Close();
		}

		private void btImprime_Click(object sender, System.EventArgs e)
		{
			Cursor = Cursors.WaitCursor;
			switch ((int) cmbSalida.Value)
			{
				case 0:
					ImprimeRep(1);
					break;
				case 1:
					oRpt.ExportToDisk(CrystalDecisions.Shared.ExportFormatType.Excel, "Salida.xls");
					break;
				case 2:
					oRpt.ExportToDisk(CrystalDecisions.Shared.ExportFormatType.ExcelRecord, "Salida.xls");
					break;
				case 3:
					oRpt.ExportToDisk(CrystalDecisions.Shared.ExportFormatType.PortableDocFormat, "Salida.pdf");
					break;
				case 4:
					oRpt.ExportToDisk(CrystalDecisions.Shared.ExportFormatType.RichText, "Salida.rtf");
					break;
				case 5:
//					oRpt.ExportToHttpResponse(CrystalDecisions.Shared.ExportFormatType.HTML32, System.Web.HttpResponse, false, "correo.rtf");
					oRpt.ExportToDisk(CrystalDecisions.Shared.ExportFormatType.HTML32, "Salida.html");
					break;
				case 6:
					oRpt.ExportToDisk(CrystalDecisions.Shared.ExportFormatType.WordForWindows, "Salida.doc");
					break;
				case 7:
					oRpt.ExportToDisk(CrystalDecisions.Shared.ExportFormatType.Text, "Salida.txt");
					break;
			}
			Cursor = Cursors.Default;
			Close();
		}

		private void Exporta()
		{
			Cursor = Cursors.WaitCursor;
			switch ((int) cmbSalida.Value)
			{
				case 0:
					oRpt.ExportToDisk(CrystalDecisions.Shared.ExportFormatType.PortableDocFormat, "correo.pdf");
					break;
				case 1:
					oRpt.ExportToDisk(CrystalDecisions.Shared.ExportFormatType.Excel, "correo.xls");
					break;
				case 2:
					oRpt.ExportToDisk(CrystalDecisions.Shared.ExportFormatType.ExcelRecord, "correo.xls");
					break;
				case 3:
					oRpt.ExportToDisk(CrystalDecisions.Shared.ExportFormatType.HTML32, "correo.html");
					break;
				case 4:
					oRpt.ExportToDisk(CrystalDecisions.Shared.ExportFormatType.WordForWindows, "correo.doc");
					break;
			}
			Cursor = Cursors.Default;
		}

		private void btImprime_MouseUp(object sender, System.Windows.Forms.MouseEventArgs e)
		{
			if (e.Button == MouseButtons.Right) Exporta();
		}

		private void btExcel_Click(object sender, System.EventArgs e)
		{
			ExportaExcel(1);
		}

		private void ExportaExcel(int iTipo)
		{
			SaveFileDialog saveFileDialog1 = new SaveFileDialog();
			saveFileDialog1.Filter = "Excel|*.xls";
			saveFileDialog1.Title = "Archivo a Exportar";
			saveFileDialog1.ShowDialog();

			// If the file name is not an empty string open it for saving.
			if(saveFileDialog1.FileName == "") return;
			// Declare variables and get the export options.
			ExportOptions exportOpts = new ExportOptions();
			ExcelFormatOptions excelFormatOpts = new ExcelFormatOptions ();
			DiskFileDestinationOptions diskOpts = new DiskFileDestinationOptions();
			exportOpts = oRpt.ExportOptions;
			
			// Set the excel format options.
			excelFormatOpts.ExcelUseConstantColumnWidth = true;
			excelFormatOpts.ExcelTabHasColumnHeadings = true;
			if (iTipo == 1)
				exportOpts.ExportFormatType = ExportFormatType.ExcelRecord;
			else
				exportOpts.ExportFormatType = ExportFormatType.Excel;

			exportOpts.FormatOptions = excelFormatOpts;
			
			// Set the disk file options and export.
			exportOpts.ExportDestinationType = ExportDestinationType.DiskFile;
			diskOpts.DiskFileName = saveFileDialog1.FileName;
			exportOpts.DestinationOptions = diskOpts;
			
			oRpt.Export ();

			if (DialogResult.Yes == MessageBox.Show(
				"¿Desea Ver Archivo Generado?", 
				"Información", MessageBoxButtons.YesNo, MessageBoxIcon.Question))
				System.Diagnostics.Process.Start(saveFileDialog1.FileName);
		}

		private void btExcel_MouseUp(object sender, System.Windows.Forms.MouseEventArgs e)
		{
			if (e.Button != MouseButtons.Right) return;
			ExportaExcel(2);		
		}
	}
}
