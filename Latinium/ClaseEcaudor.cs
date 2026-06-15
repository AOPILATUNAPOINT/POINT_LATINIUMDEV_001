using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using System.Data;
using System.Data.SqlClient;
using System.Data.OleDb;
using System.Globalization;

namespace Latinium
{
	/// <summary>
	/// Descripción breve de ClaseEcuador.
	/// </summary>
	public class ClaseEcaudor : DevExpress.XtraEditors.XtraForm
	{
		private Infragistics.Win.Misc.UltraButton btLee;
		private C1.Data.C1DataSet cdsCliente;
		/// <summary>
		/// Variable del diseñador requerida.
		/// </summary>
		private System.ComponentModel.Container components = null;
		private Infragistics.Win.Misc.UltraButton btAsiento;
		CultureInfo us = new CultureInfo("en-US");
		private Infragistics.Win.Misc.UltraButton btErrores;
		private Infragistics.Win.UltraWinEditors.UltraCheckEditor chkBorrarError;
		private System.Windows.Forms.OpenFileDialog openFileDialog;
		CultureInfo es = new CultureInfo("es");

		public ClaseEcaudor()
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
			this.btLee = new Infragistics.Win.Misc.UltraButton();
			this.cdsCliente = new C1.Data.C1DataSet();
			this.btAsiento = new Infragistics.Win.Misc.UltraButton();
			this.btErrores = new Infragistics.Win.Misc.UltraButton();
			this.chkBorrarError = new Infragistics.Win.UltraWinEditors.UltraCheckEditor();
			this.openFileDialog = new System.Windows.Forms.OpenFileDialog();
			((System.ComponentModel.ISupportInitialize)(this.cdsCliente)).BeginInit();
			this.SuspendLayout();
			// 
			// btLee
			// 
			this.btLee.Location = new System.Drawing.Point(160, 40);
			this.btLee.Name = "btLee";
			this.btLee.TabIndex = 3;
			this.btLee.Text = "Facturas";
			this.btLee.Click += new System.EventHandler(this.btLee_Click);
			// 
			// cdsCliente
			// 
			this.cdsCliente.BindingContextCtrl = this;
			this.cdsCliente.DataLibrary = "LibFacturacion";
			this.cdsCliente.DataLibraryUrl = "";
			this.cdsCliente.DataSetDef = "dsSeteoF";
			this.cdsCliente.Locale = new System.Globalization.CultureInfo("es-EC");
			this.cdsCliente.Name = "cdsCliente";
			this.cdsCliente.SchemaClassName = "LibFacturacion.DataClass";
			this.cdsCliente.SchemaDef = null;
			this.cdsCliente.BeforeFill += new C1.Data.FillEventHandler(this.cdsCliente_BeforeFill);
			// 
			// btAsiento
			// 
			this.btAsiento.Location = new System.Drawing.Point(160, 8);
			this.btAsiento.Name = "btAsiento";
			this.btAsiento.TabIndex = 4;
			this.btAsiento.Text = "Asientos";
			this.btAsiento.Click += new System.EventHandler(this.btAsiento_Click);
			// 
			// btErrores
			// 
			this.btErrores.Location = new System.Drawing.Point(32, 40);
			this.btErrores.Name = "btErrores";
			this.btErrores.TabIndex = 11;
			this.btErrores.Text = "Errores";
			this.btErrores.Click += new System.EventHandler(this.btErrores_Click);
			// 
			// chkBorrarError
			// 
			this.chkBorrarError.Location = new System.Drawing.Point(32, 8);
			this.chkBorrarError.Name = "chkBorrarError";
			this.chkBorrarError.Size = new System.Drawing.Size(88, 20);
			this.chkBorrarError.TabIndex = 12;
			this.chkBorrarError.Text = "Borrar Error";
			// 
			// ClaseEcaudor
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 13);
			this.ClientSize = new System.Drawing.Size(272, 94);
			this.Controls.Add(this.chkBorrarError);
			this.Controls.Add(this.btErrores);
			this.Controls.Add(this.btAsiento);
			this.Controls.Add(this.btLee);
			this.Name = "ClaseEcaudor";
			this.Text = "Clase Ecuador";
			this.Closing += new System.ComponentModel.CancelEventHandler(this.TelcoSan_Closing);
			this.Load += new System.EventHandler(this.TelcoSan_Load);
			((System.ComponentModel.ISupportInitialize)(this.cdsCliente)).EndInit();
			this.ResumeLayout(false);

		}
		#endregion

		DataSet miDataSet = new DataSet();
		private void ultraButton1_Click(object sender, System.EventArgs e)
		{
		}

		private bool EsRepetido(string stNumero, string stSerie)
		{
			string stSelect = "Select Count(*) From Compra Where Numero = '"
				+ stNumero.Trim() + "' And SerieFactura = '" + stSerie.Trim() + "'";
			int iCuenta = Funcion.iEscalarSQL(cdsCliente, stSelect, true);
			if (iCuenta > 0) return true;
			else return false;
		}

		private int CalculoId (string stNumero, string stSerie)
		{
			string stSelect = "Select IdCompra From Compra Where Numero = '"
				+ stNumero.Trim() + "' And SerieFactura = '" + stSerie.Trim() + "'";
			int IdCompra = Funcion.iEscalarSQL(cdsCliente, stSelect, true);
			return IdCompra;
		}

		private void btLee_Click(object sender, System.EventArgs e)
		{
			string stSelect =	"Select Count(*) From compra";
			int iCuenta = Funcion.iEscalarSQL(cdsCliente, stSelect, true);

			//string stNombre = cmbTipo.Text;
			//if (DialogResult.No == MessageBox.Show("Desea Subir las facturas ", "Confirmación", MessageBoxButtons.YesNo
			//	, MessageBoxIcon.Question)) return;
			string stArchivo = "";
			
			if (DialogResult.No == openFileDialog.ShowDialog()) return;
			stArchivo = openFileDialog.FileName;
			int largo=stArchivo.Length; 
			 stSelect =	"Select Count(*) From compra";
			iCuenta = Funcion.iEscalarSQL(cdsCliente, stSelect, true);
			string stTipo=stArchivo.Substring(14,largo-18); 	
			stSelect = "EXEC ClaseEcuFactura '" + stTipo +"'";
			Funcion.EjecutaSQL(cdsCliente, stSelect, true);

			stSelect =	"Select Count(*) From compra";
			iCuenta = Funcion.iEscalarSQL(cdsCliente, stSelect, true) - iCuenta;
			MessageBox.Show("Proceso realizado exitosamente, se han creado " + iCuenta.ToString() + " Notas de Venta ", "Subida de Notas de Venta");
		}

		private void TelcoSan_Load(object sender, System.EventArgs e)
		{
			string stAudita = "Exec AuditaCrear 53, 6, 'Telcosan'";
			Funcion.EjecutaSQL(cdsCliente, stAudita, true);
			//cmbTipo.Value = 0;		
		}

		private void cdsCliente_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			cdsCliente.Schema.Connections[0].ConnectionString = 
				miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);
		}

		private void btAsiento_Click(object sender, System.EventArgs e)
		{
			string stArchivo,stSelect;
			if (chkBorrarError.Checked)
			{
				string stBorra = "Delete ErrorSalida";
				Funcion.EjecutaSQL(cdsCliente, stBorra, true);
			}
			if (DialogResult.No == openFileDialog.ShowDialog()) return;
			stArchivo = openFileDialog.FileName;
			Cursor = Cursors.WaitCursor;
			
			int largo=stArchivo.Length; 
			string stTipo=stArchivo.Substring(14,largo-18); 	

			stSelect="select count(*) from asiento where numero='"+ stTipo +"'";
			int cuantos=Funcion.iEscalarSQL(cdsCliente,stSelect);
			if(cuantos==0)
			{
				stSelect = "EXEC ClaseEcuAsiFactura '" + stTipo +"'";
				Funcion.EjecutaSQL(cdsCliente, stSelect, true);
				Cursor = Cursors.Default;
				MessageBox.Show("Asiento creado exitosamente"); 
			}
			else
			{
				MessageBox.Show("Esta recaudacion ya tiene su asiento");  
			}
		}

		
		private void ultraButton2_Click(object sender, System.EventArgs e)
		{
		}

		private void btErrores_Click(object sender, System.EventArgs e)
		{
			TesterError miError = new TesterError();
			miError.MdiParent = this.MdiParent;
			miError.Show();
		}

		private void TelcoSan_Closing(object sender, System.ComponentModel.CancelEventArgs e)
		{
			string stAudita = "Exec AuditaCrear 53, 7, 'Telcosan'";
			Funcion.EjecutaSQL(cdsCliente, stAudita, true);
		}

		private void txtFile_ValueChanged(object sender, System.EventArgs e)
		{
		
		}
	}
}
