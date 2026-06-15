using System;
using System.IO;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;

namespace Latinium
{
	/// <summary>
	/// Descripción breve de GrabarBase.
	/// </summary>
	public class GrabarBase : DevExpress.XtraEditors.XtraForm
	{
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtDestino;
		private Infragistics.Win.Misc.UltraLabel Destino;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor ultraTextEditor1;
		private Infragistics.Win.Misc.UltraLabel ultraLabel1;
		private Infragistics.Win.Misc.UltraButton btAceptar;
		private Infragistics.Win.Misc.UltraButton btCancelar;
		private C1.Data.C1DataSet cdsCuenta;
		/// <summary>
		/// Variable del diseñador requerida.
		/// </summary>
		private System.ComponentModel.Container components = null;

		public GrabarBase()
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
			this.btAceptar = new Infragistics.Win.Misc.UltraButton();
			this.btCancelar = new Infragistics.Win.Misc.UltraButton();
			this.txtDestino = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.Destino = new Infragistics.Win.Misc.UltraLabel();
			this.ultraTextEditor1 = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.ultraLabel1 = new Infragistics.Win.Misc.UltraLabel();
			this.cdsCuenta = new C1.Data.C1DataSet();
			((System.ComponentModel.ISupportInitialize)(this.txtDestino)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraTextEditor1)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsCuenta)).BeginInit();
			this.SuspendLayout();
			// 
			// btAceptar
			// 
			this.btAceptar.Location = new System.Drawing.Point(40, 104);
			this.btAceptar.Name = "btAceptar";
			this.btAceptar.TabIndex = 0;
			this.btAceptar.Text = "&Aceptar";
			this.btAceptar.Click += new System.EventHandler(this.btAceptar_Click);
			// 
			// btCancelar
			// 
			this.btCancelar.DialogResult = System.Windows.Forms.DialogResult.Cancel;
			this.btCancelar.Location = new System.Drawing.Point(200, 104);
			this.btCancelar.Name = "btCancelar";
			this.btCancelar.TabIndex = 1;
			this.btCancelar.Text = "&Cancelar";
			this.btCancelar.Click += new System.EventHandler(this.btCancelar_Click);
			// 
			// txtDestino
			// 
			this.txtDestino.Location = new System.Drawing.Point(88, 40);
			this.txtDestino.Name = "txtDestino";
			this.txtDestino.Size = new System.Drawing.Size(224, 21);
			this.txtDestino.TabIndex = 2;
			this.txtDestino.Text = "c:\\Latinium\\Respaldo\\";
			// 
			// Destino
			// 
			this.Destino.Location = new System.Drawing.Point(176, 96);
			this.Destino.Name = "Destino";
			this.Destino.TabIndex = 3;
			this.Destino.Text = "ultraLabel1";
			// 
			// ultraTextEditor1
			// 
			this.ultraTextEditor1.Location = new System.Drawing.Point(40, 24);
			this.ultraTextEditor1.Name = "ultraTextEditor1";
			this.ultraTextEditor1.Size = new System.Drawing.Size(224, 21);
			this.ultraTextEditor1.TabIndex = 2;
			this.ultraTextEditor1.Text = "ultraTextEditor1";
			// 
			// ultraLabel1
			// 
			this.ultraLabel1.AutoSize = true;
			this.ultraLabel1.Location = new System.Drawing.Point(24, 40);
			this.ultraLabel1.Name = "ultraLabel1";
			this.ultraLabel1.Size = new System.Drawing.Size(43, 14);
			this.ultraLabel1.TabIndex = 3;
			this.ultraLabel1.Text = "Destino";
			// 
			// cdsCuenta
			// 
			this.cdsCuenta.BindingContextCtrl = this;
			this.cdsCuenta.DataLibrary = "LibContabilidad";
			this.cdsCuenta.DataLibraryUrl = "";
			this.cdsCuenta.DataSetDef = "dsCuentaCorriente";
			this.cdsCuenta.Locale = new System.Globalization.CultureInfo("es-EC");
			this.cdsCuenta.Name = "cdsCuenta";
			this.cdsCuenta.SchemaClassName = "LibContabilidad.DataClass";
			this.cdsCuenta.SchemaDef = null;
			// 
			// GrabarBase
			// 
			this.AcceptButton = this.btAceptar;
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 13);
			this.CancelButton = this.btCancelar;
			this.ClientSize = new System.Drawing.Size(336, 149);
			this.Controls.Add(this.ultraLabel1);
			this.Controls.Add(this.txtDestino);
			this.Controls.Add(this.btCancelar);
			this.Controls.Add(this.btAceptar);
			this.MaximizeBox = false;
			this.Name = "GrabarBase";
			this.Text = "Graba Base de Datos";
			this.Closing += new System.ComponentModel.CancelEventHandler(this.GrabarBase_Closing);
			this.Load += new System.EventHandler(this.GrabarBase_Load);
			((System.ComponentModel.ISupportInitialize)(this.txtDestino)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraTextEditor1)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsCuenta)).EndInit();
			this.ResumeLayout(false);

		}
		#endregion

		private void btCancelar_Click(object sender, System.EventArgs e)
		{
			Close();
		}

		private void btAceptar_Click(object sender, System.EventArgs e)
		{
			string stDir = txtDestino.Text.Trim();
			if (!Directory.Exists(stDir))
			{
				MessageBox.Show("Respaldo Cancelado porque NO existe Directorio: " + stDir,
					"Información", MessageBoxButtons.OK, MessageBoxIcon.Information);
				return;
			}
			DateTime dtIni = DateTime.Now;

			string stBase = MenuLatinium.stDirFacturacion; 
			DateTime dtHoy = DateTime.Now;
			stDir += stBase + "_" + dtHoy.ToString("dd-MMM-yyyy_HH_mm") + ".Bak";
			if (File.Exists(stDir))
			{
				if (DialogResult.No == MessageBox.Show("El Archivo: " + stDir
					+ " ya existe.\n¿Desea Sobrescribirlo?",
					"Confirmación", MessageBoxButtons.YesNo, MessageBoxIcon.Question)) return;
			}
			if (DialogResult.No == MessageBox.Show("¿Desea Generar Respaldo en "
				+ stDir + "?", "Confirmación", MessageBoxButtons.YesNo,
				MessageBoxIcon.Warning)) return;

			Cursor = Cursors.WaitCursor;

			// Respaldo de la base de datos
			string stSelect = string.Format("Exec RespaldoDataBase '{0}', '{1}'", 
				MenuLatinium.stDirFacturacion, stDir);
			string stMensaje = Funcion.sEscalarSQL(cdsCuenta, stSelect);
			if (stMensaje.Length > 0)
			{
				MessageBox.Show("No se puede obtener respaldo. \n"+stMensaje, "Informacion",
					MessageBoxButtons.OK, MessageBoxIcon.Warning);
				return;
			}

			Cursor = Cursors.Default;
			TimeSpan ts = DateTime.Now.Subtract(dtIni);
			MessageBox.Show("Respaldo " + stDir + "\nGenerado en : " + ts.TotalSeconds.ToString("n2") + " seg.", 
				"Información", MessageBoxButtons.OK, MessageBoxIcon.Information);
		}

		private void GrabarBase_Load(object sender, System.EventArgs e)
		{
			string stAudita = "Exec AuditaCrear 53, 6, 'Graba Base'";
			Funcion.EjecutaSQL(cdsCuenta, stAudita, true);
		}

		private void GrabarBase_Closing(object sender, System.ComponentModel.CancelEventArgs e)
		{
			string stAudita = "Exec AuditaCrear 53, 7, 'Graba Base'";
			Funcion.EjecutaSQL(cdsCuenta, stAudita, true);
		}
	}
}
