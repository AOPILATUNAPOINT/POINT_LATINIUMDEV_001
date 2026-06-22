using System;
using System.Data;
using System.Data.SqlClient;
using System.Drawing;
using System.Collections;
using System.Windows.Forms;

namespace Latinium
{
	public class frmResumenCierreCajaConfiguracion : System.Windows.Forms.Form
	{
		private System.Windows.Forms.Label lblTitulo;
		private System.Windows.Forms.Label lblInstruccion;
		private System.Windows.Forms.ListView listFormas;
		private System.Windows.Forms.ColumnHeader colFormaPago;
		private System.Windows.Forms.ColumnHeader colCod;
		private System.Windows.Forms.Button btnGuardar;
		private System.Windows.Forms.Button btnCerrar;

		private C1.Data.C1DataSet cdsSeteoF;

		public frmResumenCierreCajaConfiguracion(C1.Data.C1DataSet cdsSeteoF)
		{
			this.cdsSeteoF = cdsSeteoF;
			InitializeComponent();
			CargarDatos();
		}

		private void InitializeComponent()
		{
			this.lblTitulo      = new System.Windows.Forms.Label();
			this.lblInstruccion = new System.Windows.Forms.Label();
			this.listFormas     = new System.Windows.Forms.ListView();
			this.colFormaPago   = new System.Windows.Forms.ColumnHeader();
			this.colCod         = new System.Windows.Forms.ColumnHeader();
			this.btnGuardar     = new System.Windows.Forms.Button();
			this.btnCerrar      = new System.Windows.Forms.Button();

			this.SuspendLayout();

			// lblTitulo
			this.lblTitulo.Location = new System.Drawing.Point(8, 8);
			this.lblTitulo.Size     = new System.Drawing.Size(480, 20);
			this.lblTitulo.Font     = new System.Drawing.Font("Tahoma", 9F, System.Drawing.FontStyle.Bold);
			this.lblTitulo.Text     = "Configuración - Validación de Números Duplicados";

			// lblInstruccion
			this.lblInstruccion.Location = new System.Drawing.Point(8, 32);
			this.lblInstruccion.Size     = new System.Drawing.Size(480, 36);
			this.lblInstruccion.Font     = new System.Drawing.Font("Tahoma", 8F);
			this.lblInstruccion.Text     = "Marque las formas de pago que deben validar números duplicados en el cierre de caja.\r\n" +
			                              "Duplicado = misma forma de pago + misma cuenta bancaria + mismo número.";

			// colFormaPago
			this.colFormaPago.Text  = "Forma de Pago";
			this.colFormaPago.Width = 300;

			// colCod
			this.colCod.Text  = "Código";
			this.colCod.Width = 60;

			// listFormas
			this.listFormas.Location   = new System.Drawing.Point(8, 76);
			this.listFormas.Size       = new System.Drawing.Size(480, 296);
			this.listFormas.View       = System.Windows.Forms.View.Details;
			this.listFormas.CheckBoxes = true;
			this.listFormas.FullRowSelect = true;
			this.listFormas.GridLines  = true;
			this.listFormas.Font       = new System.Drawing.Font("Tahoma", 8.25F);
			this.listFormas.Columns.AddRange(new System.Windows.Forms.ColumnHeader[] {
				this.colFormaPago,
				this.colCod
			});

			// btnGuardar
			this.btnGuardar.Location = new System.Drawing.Point(328, 380);
			this.btnGuardar.Size     = new System.Drawing.Size(75, 24);
			this.btnGuardar.Text     = "Guardar";
			this.btnGuardar.Font     = new System.Drawing.Font("Tahoma", 8.25F);
			this.btnGuardar.Click   += new System.EventHandler(this.btnGuardar_Click);

			// btnCerrar
			this.btnCerrar.Location = new System.Drawing.Point(412, 380);
			this.btnCerrar.Size     = new System.Drawing.Size(75, 24);
			this.btnCerrar.Text     = "Cerrar";
			this.btnCerrar.Font     = new System.Drawing.Font("Tahoma", 8.25F);
			this.btnCerrar.Click   += new System.EventHandler(this.btnCerrar_Click);

			// Form
			this.ClientSize        = new System.Drawing.Size(496, 412);
			this.FormBorderStyle   = System.Windows.Forms.FormBorderStyle.FixedDialog;
			this.MaximizeBox       = false;
			this.MinimizeBox       = false;
			this.StartPosition     = System.Windows.Forms.FormStartPosition.CenterScreen;
			this.Text              = "Configuración - Validación Duplicados Cierre de Caja";
			this.Font              = new System.Drawing.Font("Tahoma", 8.25F);
			this.Controls.AddRange(new System.Windows.Forms.Control[] {
				this.lblTitulo,
				this.lblInstruccion,
				this.listFormas,
				this.btnGuardar,
				this.btnCerrar
			});

			this.ResumeLayout(false);
		}

		private void CargarDatos()
		{
			try
			{
				listFormas.Items.Clear();

				// Lee configuracion actual de la tabla
				Hashtable htActivos = new Hashtable();
				SqlDataReader drCfg = Funcion.rEscalarSQL(cdsSeteoF,
					"SELECT Valor, Activo FROM DetalleFormaPagoConfiguracion WHERE Clave='ValidaDepDuplicadoFP'", true);
				while (drCfg.Read())
					htActivos[drCfg.GetInt32(0)] = drCfg.GetBoolean(1);
				drCfg.Close();

				// Carga formas de pago disponibles en el grid de depositos
				SqlDataReader drFP = Funcion.rEscalarSQL(cdsSeteoF,
					"SELECT idFormaPago, FormaPago FROM CompraFormaPago " +
					"WHERE idFormaPago IN (1,3,7,13,31,32,33,34,35,36,38,41,42,43,44,45) " +
					"ORDER BY FormaPago", true);

				while (drFP.Read())
				{
					int    idFP   = drFP.GetInt32(0);
					string nombre = drFP.GetString(1);
					bool   activo = htActivos.ContainsKey(idFP) ? (bool)htActivos[idFP] : false;

					ListViewItem item = new ListViewItem(nombre);
					item.SubItems.Add(idFP.ToString());
					item.Checked = activo;
					item.Tag     = idFP;
					listFormas.Items.Add(item);
				}
				drFP.Close();
			}
			catch (Exception ex)
			{
				MessageBox.Show("Error al cargar la configuración: " + ex.Message,
					"Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
			}
		}

		private void btnGuardar_Click(object sender, System.EventArgs e)
		{
			try
			{
				foreach (ListViewItem item in listFormas.Items)
				{
					int idFP   = Convert.ToInt32(item.Tag);
					int activo = item.Checked ? 1 : 0;

					int existe = Convert.ToInt32(Funcion.iEscalarSQL(cdsSeteoF,
						string.Format("SELECT COUNT(*) FROM DetalleFormaPagoConfiguracion WHERE Clave='ValidaDepDuplicadoFP' AND Valor={0}", idFP)));

					if (existe > 0)
					{
						Funcion.EjecutaSQL(cdsSeteoF,
							string.Format("UPDATE DetalleFormaPagoConfiguracion SET Activo={0} WHERE Clave='ValidaDepDuplicadoFP' AND Valor={1}",
							activo, idFP));
					}
					else if (activo == 1)
					{
						Funcion.EjecutaSQL(cdsSeteoF,
							string.Format("INSERT INTO DetalleFormaPagoConfiguracion (Clave, Descripcion, Valor, Activo) " +
							"VALUES ('ValidaDepDuplicadoFP', '{0} - valida numero duplicado', {1}, 1)",
							item.Text.Replace("'", "''"), idFP));
					}
				}

				MessageBox.Show("Configuración guardada correctamente.",
					"Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Information);
			}
			catch (Exception ex)
			{
				MessageBox.Show("Error al guardar: " + ex.Message,
					"Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
			}
		}

		private void btnCerrar_Click(object sender, System.EventArgs e)
		{
			this.Close();
		}

		protected override void Dispose(bool disposing)
		{
			base.Dispose(disposing);
		}
	}
}
