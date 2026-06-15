using System;
using System.Collections;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Windows.Forms;
using Latinium.Services.Pinpad.Configuracion;
using Latinium.Models.Pinpad.Configuracion;

namespace Latinium
{
	public class frmPinPadConfiguracion : Form
	{
		private Container components = null;

		private DataGrid dgConfiguracion;
		private Button btnEditar;
		private Button btnEliminar;
		private Button btnRefrescar;
		private PinpadConfiguracionRepository repo;

		public frmPinPadConfiguracion()
		{
			InitializeComponent();
		}

		protected override void Dispose(bool disposing)
		{
			if (disposing)
			{
				if (components != null)
					components.Dispose();
			}
			base.Dispose(disposing);
		}

		private void InitializeComponent()
		{
			this.dgConfiguracion = new System.Windows.Forms.DataGrid();
			this.btnEditar = new System.Windows.Forms.Button();
			this.btnEliminar = new System.Windows.Forms.Button();
			this.btnRefrescar = new System.Windows.Forms.Button();
			((System.ComponentModel.ISupportInitialize)(this.dgConfiguracion)).BeginInit();
			this.SuspendLayout();
			// 
			// dgConfiguracion
			// 
			this.dgConfiguracion.DataMember = "";
			this.dgConfiguracion.HeaderForeColor = System.Drawing.SystemColors.ControlText;
			this.dgConfiguracion.Location = new System.Drawing.Point(10, 10);
			this.dgConfiguracion.Name = "dgConfiguracion";
			this.dgConfiguracion.ReadOnly = true;
			this.dgConfiguracion.Size = new System.Drawing.Size(560, 250);
			this.dgConfiguracion.TabIndex = 0;
			// 
			// btnEditar
			// 
			this.btnEditar.Location = new System.Drawing.Point(12, 270);
			this.btnEditar.Name = "btnEditar";
			this.btnEditar.TabIndex = 2;
			this.btnEditar.Text = "Editar";
			this.btnEditar.Click += new System.EventHandler(this.btnEditar_Click);
			// 
			// btnEliminar
			// 
			this.btnEliminar.Location = new System.Drawing.Point(104, 270);
			this.btnEliminar.Name = "btnEliminar";
			this.btnEliminar.TabIndex = 3;
			this.btnEliminar.Text = "Eliminar";
			this.btnEliminar.Click += new System.EventHandler(this.btnEliminar_Click);
			// 
			// btnRefrescar
			// 
			this.btnRefrescar.Location = new System.Drawing.Point(200, 270);
			this.btnRefrescar.Name = "btnRefrescar";
			this.btnRefrescar.TabIndex = 4;
			this.btnRefrescar.Text = "Refrescar";
			this.btnRefrescar.Click += new System.EventHandler(this.btnRefrescar_Click);
			// 
			// frmPinPadConfiguracion
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 13);
			this.ClientSize = new System.Drawing.Size(584, 302);
			this.Controls.Add(this.dgConfiguracion);
			this.Controls.Add(this.btnEditar);
			this.Controls.Add(this.btnEliminar);
			this.Controls.Add(this.btnRefrescar);
			this.Name = "frmPinPadConfiguracion";
			this.Text = "Configuración PinPad";
			this.Load += new System.EventHandler(this.frmPinPadConfiguracion_Load);
			((System.ComponentModel.ISupportInitialize)(this.dgConfiguracion)).EndInit();
			this.ResumeLayout(false);

		}

		// =========================
		// LOAD
		// =========================
		private void frmPinPadConfiguracion_Load(object sender, EventArgs e)
		{
			repo = new PinpadConfiguracionRepository(MenuLatinium.sconexionEmpresa);
			CargarDatos();
		}

		private void CargarDatos()
		{
			ArrayList lista = repo.Listar();

			DataTable dt = new DataTable();

			dt.Columns.Add("Id");
			dt.Columns.Add("DeviceId");
			dt.Columns.Add("PC");
			dt.Columns.Add("BodegaNombre");
			dt.Columns.Add("Caja");
			dt.Columns.Add("IP");
			dt.Columns.Add("MID");
			dt.Columns.Add("TID");
			dt.Columns.Add("CID");
			dt.Columns.Add("Puerto");
			dt.Columns.Add("Timeout");

			foreach (PinpadConfiguracion cfg in lista)
			{
				DataRow row = dt.NewRow();

				row["Id"] = cfg.IdPinpadConfiguracion.ToString();
				row["DeviceId"] = cfg.DeviceId;
				row["PC"] = cfg.NombrePC;
				row["BodegaNombre"] = cfg.NombreBodega;
				row["Caja"] = cfg.NumeroCaja;
				row["IP"] = cfg.IpPinPad == null ? "" : cfg.IpPinPad;

				row["MID"] = cfg.MID;
				row["TID"] = cfg.TID;
				row["CID"] = cfg.CID;

				row["Puerto"] = cfg.PuertoPinPad;
				row["Timeout"] = cfg.TimeoutPinPad;

				dt.Rows.Add(row);
			}

			dgConfiguracion.DataSource = dt;
		}
		// =========================
		// NUEVO
		// =========================
		private void btnNuevo_Click(object sender, EventArgs e)
		{
			    frmPinPadConfiguracionDetalle frm = new frmPinPadConfiguracionDetalle();
			frm.ShowDialog();
			CargarDatos();
		}

		// =========================
		// EDITAR
		// =========================
		private void btnEditar_Click(object sender, EventArgs e)
		{
			if (dgConfiguracion.DataSource == null) return;
			if (dgConfiguracion.CurrentRowIndex < 0) return;

			DataRowView drv =
				(DataRowView)dgConfiguracion.BindingContext[dgConfiguracion.DataSource].Current;

			Guid id = new Guid(drv["Id"].ToString());

			frmPinPadConfiguracionDetalle frm =
				new frmPinPadConfiguracionDetalle(id);

			frm.ShowDialog();

			CargarDatos();
		}

		// =========================
		// ELIMINAR
		// =========================
		private void btnEliminar_Click(object sender, EventArgs e)
		{
			if (dgConfiguracion.DataSource == null) return;
if (dgConfiguracion.CurrentRowIndex < 0) return;

			DataRowView drv =
				(DataRowView)dgConfiguracion.BindingContext[dgConfiguracion.DataSource].Current;

			Guid id = new Guid(drv["Id"].ToString());

			DialogResult dr = MessageBox.Show(
				"¿Eliminar registro?",
				"Confirmar",
				MessageBoxButtons.YesNo);

			if (dr == DialogResult.Yes)
			{
				repo.Eliminar(id);
				CargarDatos();
			}
		}
		// =========================
		// REFRESCAR
		// =========================
		private void btnRefrescar_Click(object sender, EventArgs e)
		{
			CargarDatos();
		}
	}
}