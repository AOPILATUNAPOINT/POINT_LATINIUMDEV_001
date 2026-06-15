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

namespace Latinium
{
	/// <summary>
	/// Descripción breve de frmPagoFioCredito.
	/// </summary>
	public class frmPagoFioCredito : System.Windows.Forms.Form
	{
		private System.Windows.Forms.Label label6;
		public Infragistics.Win.UltraWinSchedule.UltraCalendarCombo dtFecha;
		public Infragistics.Win.UltraWinEditors.UltraOptionSet optTransporte;
		public Infragistics.Win.UltraWinEditors.UltraTextEditor txtOrigen;
		private System.Windows.Forms.Label label1;
		private System.Windows.Forms.Label label3;
		public Infragistics.Win.UltraWinEditors.UltraNumericEditor txtValor;
		private System.Windows.Forms.Label label2;
		public System.Windows.Forms.Button btnGuardar;
		private System.Windows.Forms.Button btnSalir;
		public Infragistics.Win.UltraWinEditors.UltraNumericEditor txtvalorPendiente;
		private System.Windows.Forms.Button btnNube;
		private C1.Data.C1DataSet cdsSeteoF;
		/// <summary>
		/// Variable del diseñador requerida.
		/// </summary>
		private System.ComponentModel.Container components = null;

		public frmPagoFioCredito( )
		{
			//
			// Necesario para admitir el Diseñador de Windows Forms
			//
			InitializeComponent();

			//
			// TODO: agregar código de constructor después de llamar a InitializeComponent
			//
		}
		int idFioCredito = 0;
		public frmPagoFioCredito(int IdFioCredito )
		{
			//
			// Necesario para admitir el Diseñador de Windows Forms
			//
			idFioCredito = IdFioCredito ;
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
			Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton dateButton1 = new Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton();
			System.Resources.ResourceManager resources = new System.Resources.ResourceManager(typeof(frmPagoFioCredito));
			Infragistics.Win.Appearance appearance2 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance3 = new Infragistics.Win.Appearance();
			Infragistics.Win.ValueListItem valueListItem1 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.ValueListItem valueListItem2 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.Appearance appearance4 = new Infragistics.Win.Appearance();
			this.label6 = new System.Windows.Forms.Label();
			this.dtFecha = new Infragistics.Win.UltraWinSchedule.UltraCalendarCombo();
			this.optTransporte = new Infragistics.Win.UltraWinEditors.UltraOptionSet();
			this.txtOrigen = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.label1 = new System.Windows.Forms.Label();
			this.label3 = new System.Windows.Forms.Label();
			this.txtValor = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.label2 = new System.Windows.Forms.Label();
			this.txtvalorPendiente = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.btnNube = new System.Windows.Forms.Button();
			this.btnGuardar = new System.Windows.Forms.Button();
			this.btnSalir = new System.Windows.Forms.Button();
			this.cdsSeteoF = new C1.Data.C1DataSet();
			((System.ComponentModel.ISupportInitialize)(this.dtFecha)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.optTransporte)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtOrigen)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtValor)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtvalorPendiente)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).BeginInit();
			this.SuspendLayout();
			// 
			// label6
			// 
			this.label6.AutoSize = true;
			this.label6.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label6.Location = new System.Drawing.Point(16, 10);
			this.label6.Name = "label6";
			this.label6.Size = new System.Drawing.Size(36, 16);
			this.label6.TabIndex = 201;
			this.label6.Text = "Fecha";
			// 
			// dtFecha
			// 
			appearance1.ForeColorDisabled = System.Drawing.Color.Black;
			this.dtFecha.Appearance = appearance1;
			this.dtFecha.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			dateButton1.Caption = "Today";
			this.dtFecha.DateButtons.Add(dateButton1);
			this.dtFecha.Format = "dd/MM/yyyy";
			this.dtFecha.Location = new System.Drawing.Point(64, 8);
			this.dtFecha.Name = "dtFecha";
			this.dtFecha.NonAutoSizeHeight = 23;
			this.dtFecha.Size = new System.Drawing.Size(104, 21);
			this.dtFecha.SpinButtonsVisible = true;
			this.dtFecha.TabIndex = 200;
			this.dtFecha.Value = ((object)(resources.GetObject("dtFecha.Value")));
			// 
			// optTransporte
			// 
			appearance2.ForeColorDisabled = System.Drawing.Color.Black;
			this.optTransporte.Appearance = appearance2;
			this.optTransporte.BorderStyle = Infragistics.Win.UIElementBorderStyle.Solid;
			this.optTransporte.CheckedIndex = 0;
			this.optTransporte.ItemAppearance = appearance3;
			this.optTransporte.ItemOrigin = new System.Drawing.Point(10, 0);
			valueListItem1.DataValue = 1;
			valueListItem1.DisplayText = "Normal";
			valueListItem2.DataValue = 2;
			valueListItem2.DisplayText = "Nomina";
			this.optTransporte.Items.Add(valueListItem1);
			this.optTransporte.Items.Add(valueListItem2);
			this.optTransporte.ItemSpacingVertical = 10;
			this.optTransporte.Location = new System.Drawing.Point(176, 8);
			this.optTransporte.Name = "optTransporte";
			this.optTransporte.Size = new System.Drawing.Size(152, 24);
			this.optTransporte.TabIndex = 1020;
			this.optTransporte.Text = "Normal";
			// 
			// txtOrigen
			// 
			appearance4.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtOrigen.Appearance = appearance4;
			this.txtOrigen.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtOrigen.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtOrigen.Location = new System.Drawing.Point(64, 48);
			this.txtOrigen.Multiline = true;
			this.txtOrigen.Name = "txtOrigen";
			this.txtOrigen.Size = new System.Drawing.Size(304, 63);
			this.txtOrigen.TabIndex = 1021;
			// 
			// label1
			// 
			this.label1.AutoSize = true;
			this.label1.Location = new System.Drawing.Point(16, 56);
			this.label1.Name = "label1";
			this.label1.Size = new System.Drawing.Size(28, 16);
			this.label1.TabIndex = 1022;
			this.label1.Text = "Nota";
			this.label1.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// label3
			// 
			this.label3.AutoSize = true;
			this.label3.Location = new System.Drawing.Point(16, 130);
			this.label3.Name = "label3";
			this.label3.Size = new System.Drawing.Size(36, 16);
			this.label3.TabIndex = 1024;
			this.label3.Text = "Monto";
			this.label3.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// txtValor
			// 
			this.txtValor.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtValor.FormatString = "#,##0.00";
			this.txtValor.Location = new System.Drawing.Point(64, 128);
			this.txtValor.Name = "txtValor";
			this.txtValor.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtValor.PromptChar = ' ';
			this.txtValor.Size = new System.Drawing.Size(112, 21);
			this.txtValor.TabIndex = 1023;
			// 
			// label2
			// 
			this.label2.AutoSize = true;
			this.label2.Location = new System.Drawing.Point(192, 130);
			this.label2.Name = "label2";
			this.label2.Size = new System.Drawing.Size(55, 16);
			this.label2.TabIndex = 1026;
			this.label2.Text = "Pendiente";
			this.label2.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// txtvalorPendiente
			// 
			this.txtvalorPendiente.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtvalorPendiente.Enabled = false;
			this.txtvalorPendiente.FormatString = "#,##0.00";
			this.txtvalorPendiente.Location = new System.Drawing.Point(256, 128);
			this.txtvalorPendiente.Name = "txtvalorPendiente";
			this.txtvalorPendiente.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtvalorPendiente.PromptChar = ' ';
			this.txtvalorPendiente.Size = new System.Drawing.Size(112, 21);
			this.txtvalorPendiente.TabIndex = 1025;
			// 
			// btnNube
			// 
			this.btnNube.Anchor = System.Windows.Forms.AnchorStyles.Top;
			this.btnNube.Image = ((System.Drawing.Image)(resources.GetObject("btnNube.Image")));
			this.btnNube.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnNube.Location = new System.Drawing.Point(352, 8);
			this.btnNube.Name = "btnNube";
			this.btnNube.Size = new System.Drawing.Size(25, 24);
			this.btnNube.TabIndex = 1027;
			this.btnNube.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnNube.Click += new System.EventHandler(this.btnNube_Click);
			// 
			// btnGuardar
			// 
			this.btnGuardar.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btnGuardar.Image = ((System.Drawing.Image)(resources.GetObject("btnGuardar.Image")));
			this.btnGuardar.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnGuardar.Location = new System.Drawing.Point(16, 168);
			this.btnGuardar.Name = "btnGuardar";
			this.btnGuardar.Size = new System.Drawing.Size(76, 24);
			this.btnGuardar.TabIndex = 1028;
			this.btnGuardar.Text = "&Guardar";
			this.btnGuardar.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnGuardar.Click += new System.EventHandler(this.btnGuardar_Click);
			// 
			// btnSalir
			// 
			this.btnSalir.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btnSalir.CausesValidation = false;
			this.btnSalir.Image = ((System.Drawing.Image)(resources.GetObject("btnSalir.Image")));
			this.btnSalir.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnSalir.Location = new System.Drawing.Point(104, 168);
			this.btnSalir.Name = "btnSalir";
			this.btnSalir.TabIndex = 1029;
			this.btnSalir.Text = "&Salir";
			this.btnSalir.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnSalir.Click += new System.EventHandler(this.btnSalir_Click);
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
			// frmPagoFioCredito
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 13);
			this.ClientSize = new System.Drawing.Size(392, 198);
			this.ControlBox = false;
			this.Controls.Add(this.btnSalir);
			this.Controls.Add(this.btnGuardar);
			this.Controls.Add(this.btnNube);
			this.Controls.Add(this.label2);
			this.Controls.Add(this.txtvalorPendiente);
			this.Controls.Add(this.label3);
			this.Controls.Add(this.txtValor);
			this.Controls.Add(this.label1);
			this.Controls.Add(this.txtOrigen);
			this.Controls.Add(this.optTransporte);
			this.Controls.Add(this.label6);
			this.Controls.Add(this.dtFecha);
			this.KeyPreview = true;
			this.Name = "frmPagoFioCredito";
			this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
			this.Text = "Cobro AlFio";
			this.KeyDown += new System.Windows.Forms.KeyEventHandler(this.frmPagoFioCredito_KeyDown);
			this.Load += new System.EventHandler(this.frmPagoFioCredito_Load);
			((System.ComponentModel.ISupportInitialize)(this.dtFecha)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.optTransporte)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtOrigen)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtValor)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtvalorPendiente)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).EndInit();
			this.ResumeLayout(false);

		}
		#endregion

		private void btnGuardar_Click(object sender, System.EventArgs e)
		{
			if (!Validacion.vbFechaEnDocumentos(this.dtFecha, "Seleccione la Fecha De Pago", false, "", cdsSeteoF)) return;
			if (!Validacion.vbTexto(this.txtOrigen, 10, 1000, "Notas")) return;
			if (!Validacion.vbCampoDecimalVacio(this.txtValor, "Ingrese el valor", 0.01m, dSaldo)) return;
      if (!Validacion.vbVariableTexto(sCarpeta, 10, 1000, "JPG")) return;
			if (!Validacion.vbVariableTexto(sUrlNube, 10, 1000, "JPG")) return;

			this.DialogResult = DialogResult.OK;
		}

		private void btnSalir_Click(object sender, System.EventArgs e)
		{
						this.DialogResult = DialogResult.No;
		}

		private void frmPagoFioCredito_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Escape) this.DialogResult = DialogResult.No;
		}
   Decimal dSaldo =  0.00m ;
			int icont = 0;
		private void frmPagoFioCredito_Load(object sender, System.EventArgs e)
		{
			
			dSaldo = Funcion.decEscalarSQL(cdsSeteoF, string.Format("Exec [SaldoFio] {0}",
				idFioCredito));	
			icont = Funcion.iEscalarSQL(cdsSeteoF, string.Format("select convert (int, (count(1) + FLOOR(RAND()*(1500-1)+1)))  from dbo.PagosAlFioCredito "));	

			this.txtvalorPendiente.Value = dSaldo;
			FuncionesProcedimientos.RetornaFechaServidor(this.dtFecha, cdsSeteoF, true);
		}

		private void cdsSeteoF_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			
			Funcion miFuncion = new Funcion();
			cdsSeteoF.Schema.Connections[0].ConnectionString = miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);		
	
		}
 public string sCarpeta ="";
		public string sUrlNube = "";
		private void btnNube_Click(object sender, System.EventArgs e)
		{
			#region Deposito
			try
			{
				MessageBox.Show("Acontinuación seleccione el voucher de transporte en formato JPG.", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
				string txtDirectorioarchivo ="";
				string txtURL ="";
				string SArhivoNuevo = Convert.ToString( idFioCredito)+"_"+ Convert.ToString( icont)+ "_"+ Convert.ToString( icont);
				OpenFileDialog buscar = new OpenFileDialog ();
				buscar.Filter = "Text files (*.jpg)|*.jpg";
				if (buscar.ShowDialog () == DialogResult.OK)
				{

					txtDirectorioarchivo = buscar.FileName;//nombre del archivo
					FileInfo fi= new FileInfo(txtDirectorioarchivo);
					string nombre = fi.Name;//nombre.jpg
					string Directorio = txtDirectorioarchivo.Replace(nombre, "")+SArhivoNuevo+".jpg";//archivo nuevo
					File.Delete(Directorio); // First delete the new file if it exists, because it wont be overwritten by Move
					File.Move(txtDirectorioarchivo, Directorio); 
						
					string carpeta_1="";
					string carpeta_2="";
					string carpeta_3="";
					string carpeta ="";
					//string txtDirectorioarchivo ="";
					//	string txtURL ="";
					
					string ScARPETAUR = @"192.168.1.10";
					
					//
					carpeta_1 = @"\\"+ScARPETAUR+@"\Proyectos\VOUCHERFIO_Vou";
					//@"\\"+
					carpeta_2 = @"\\"+ScARPETAUR+@"\Proyectos\VOUCHERFIO_Vou\"+MenuLatinium.stDirFacturacion;
					// @"\\192.168.1.10\Proyectos\Transporte\"+MenuLatinium.stDirFacturacion+@"\" + this.txtNumero.Text.Replace(" ", "_") +"_"+this.txtRuc.Text+"_"+ Convert.ToString (this.txtidRegistroTransporte.Value);
					carpeta = @"\\"+ScARPETAUR+@"\Proyectos\VOUCHERFIO_Vou\"+MenuLatinium.stDirFacturacion+@"\" + Convert.ToString( idFioCredito)+"_"+ Convert.ToString( icont)+ "_"+ Convert.ToString( icont);
					carpeta_3 = @""+MenuLatinium.stDirFacturacion+@"\" + Convert.ToString( idFioCredito)+"_"+ Convert.ToString( icont)+ "_"+ Convert.ToString(icont);
					Process.Start("cmd.exe", "/C" + "net use \\192.168.1.10  /user:latinium 10dgpv15.p");
					Directory.GetDirectories(@"\\192.168.1.10\Proyectos");

					if (!Directory.Exists(carpeta_1))
					{
						Directory.CreateDirectory(carpeta);
					
					}
					if (!Directory.Exists(carpeta_2))
					{
						Directory.CreateDirectory(carpeta);
					
					}

					if (!Directory.Exists(carpeta))
					{
						Directory.CreateDirectory(carpeta);
					
					}
					txtURL =  Funcion.sEscalarSQL(cdsSeteoF, string.Format(" SELECT RIGHT('{0}', CHARINDEX('{1}', REVERSE('{0}')) - 1) AS Archivo",Directorio,@"\"));
					carpeta  = (carpeta + @"\" + txtURL.Replace(" ", "_"));

						System.IO.File.Copy (  Directorio, carpeta, false);
						sCarpeta = carpeta;
						sUrlNube = carpeta_3+ @"\" + txtURL;
//						string ssSQL = string.Format("Exec [GrabaUrlNubeALFIO] {0}, '{1}', '{2}'", Convert.ToInt32( e.Cell.Row.Cells["idFormularioSolicitud"].Value), sCarpeta, sUrlNube );
//						Funcion.EjecutaSQL(cdsSeteoF, ssSQL);
						MessageBox.Show(string.Format(" Archivo '{0}' subido . ", sUrlNube), "Validacion", MessageBoxButtons.OK, MessageBoxIcon.Information);
					

				}
			}
			catch (Exception ex)
			{
				MessageBox.Show ("Error: "+ ex.Message);
			}
			#endregion Deposito
		}
	}
}
