using System;
using System.Collections;
using System.Data;
using System.Data.SqlClient;
using System.Drawing;
using System.Windows.Forms;
using Infragistics.Win;
using Infragistics.Win.UltraWinGrid;
using Latinium.Services.Requerimiento.TurnoSoporte;
using Latinium.Models.Requerimiento.TurnoSoporte;
using System.ComponentModel;

namespace Latinium
{
	public class frmRequerimientoTurnoSoporteConsulta : Form
	{
		private Container components = null;

		// Controles
		private UltraGrid ugTurnos;
		private Button btnNuevo;
		private Button btnInhabilitar;

		// Servicio y conexión
		private SqlConnection _cn;
		private System.Windows.Forms.GroupBox groupBox1;
		private System.Windows.Forms.Label label3;
		private System.Windows.Forms.Label label2;
		private System.Windows.Forms.TextBox txtBuscar;
		private System.Windows.Forms.DateTimePicker dtpDesde;
		private System.Windows.Forms.DateTimePicker dtpHasta;
		private System.Windows.Forms.Button btnBuscar;
		private System.Windows.Forms.Button btnSalir;
		private RequerimientoTurnoSoporteService _srv;

		public frmRequerimientoTurnoSoporteConsulta()
		{
			InitializeComponent();
		}

		protected override void Dispose(bool disposing)
		{
			if (disposing && components != null)
				components.Dispose();
			base.Dispose(disposing);
		}

		private void InitializeComponent()
		{
			this.btnNuevo = new System.Windows.Forms.Button();
			this.btnInhabilitar = new System.Windows.Forms.Button();
			this.ugTurnos = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.groupBox1 = new System.Windows.Forms.GroupBox();
			this.label3 = new System.Windows.Forms.Label();
			this.label2 = new System.Windows.Forms.Label();
			this.txtBuscar = new System.Windows.Forms.TextBox();
			this.dtpDesde = new System.Windows.Forms.DateTimePicker();
			this.dtpHasta = new System.Windows.Forms.DateTimePicker();
			this.btnBuscar = new System.Windows.Forms.Button();
			this.btnSalir = new System.Windows.Forms.Button();
			((System.ComponentModel.ISupportInitialize)(this.ugTurnos)).BeginInit();
			this.groupBox1.SuspendLayout();
			this.SuspendLayout();
			// 
			// btnNuevo
			// 
			this.btnNuevo.Location = new System.Drawing.Point(824, 16);
			this.btnNuevo.Name = "btnNuevo";
			this.btnNuevo.Size = new System.Drawing.Size(80, 25);
			this.btnNuevo.TabIndex = 0;
			this.btnNuevo.Text = "Nuevo";
			this.btnNuevo.Click += new System.EventHandler(this.BtnNuevo_Click);
			// 
			// btnInhabilitar
			// 
			this.btnInhabilitar.Location = new System.Drawing.Point(824, 48);
			this.btnInhabilitar.Name = "btnInhabilitar";
			this.btnInhabilitar.Size = new System.Drawing.Size(80, 25);
			this.btnInhabilitar.TabIndex = 1;
			this.btnInhabilitar.Text = "Inhabilitar";
			this.btnInhabilitar.Click += new System.EventHandler(this.BtnInhabilitar_Click);
			// 
			// ugTurnos
			// 
			this.ugTurnos.Cursor = System.Windows.Forms.Cursors.Default;
			this.ugTurnos.DisplayLayout.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.No;
			this.ugTurnos.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.False;
			this.ugTurnos.DisplayLayout.Override.CellClickAction = Infragistics.Win.UltraWinGrid.CellClickAction.RowSelect;
			this.ugTurnos.Location = new System.Drawing.Point(16, 136);
			this.ugTurnos.Name = "ugTurnos";
			this.ugTurnos.Size = new System.Drawing.Size(880, 256);
			this.ugTurnos.TabIndex = 2;
			this.ugTurnos.DoubleClick += new System.EventHandler(this.UgTurnos_DoubleClick);
			// 
			// groupBox1
			// 
			this.groupBox1.Controls.Add(this.label3);
			this.groupBox1.Controls.Add(this.label2);
			this.groupBox1.Controls.Add(this.txtBuscar);
			this.groupBox1.Controls.Add(this.dtpDesde);
			this.groupBox1.Controls.Add(this.dtpHasta);
			this.groupBox1.Controls.Add(this.btnBuscar);
			this.groupBox1.Location = new System.Drawing.Point(16, 8);
			this.groupBox1.Name = "groupBox1";
			this.groupBox1.Size = new System.Drawing.Size(784, 112);
			this.groupBox1.TabIndex = 10;
			this.groupBox1.TabStop = false;
			this.groupBox1.Text = "Filtro";
			// 
			// label3
			// 
			this.label3.ForeColor = System.Drawing.Color.Gray;
			this.label3.Location = new System.Drawing.Point(24, 80);
			this.label3.Name = "label3";
			this.label3.Size = new System.Drawing.Size(416, 23);
			this.label3.TabIndex = 16;
			this.label3.Text = "Inhabilitar: Desactiva el turno seleccionado sin eliminarlo del sistema.";
			// 
			// label2
			// 
			this.label2.ForeColor = System.Drawing.Color.Gray;
			this.label2.Location = new System.Drawing.Point(24, 56);
			this.label2.Name = "label2";
			this.label2.Size = new System.Drawing.Size(416, 23);
			this.label2.TabIndex = 15;
			this.label2.Text = "Edición: Doble clic en el registro para editar el turno seleccionado.";
			// 
			// txtBuscar
			// 
			this.txtBuscar.Location = new System.Drawing.Point(440, 24);
			this.txtBuscar.Name = "txtBuscar";
			this.txtBuscar.Size = new System.Drawing.Size(320, 22);
			this.txtBuscar.TabIndex = 10;
			this.txtBuscar.Text = "";
			// 
			// dtpDesde
			// 
			this.dtpDesde.Format = System.Windows.Forms.DateTimePickerFormat.Short;
			this.dtpDesde.Location = new System.Drawing.Point(24, 24);
			this.dtpDesde.Name = "dtpDesde";
			this.dtpDesde.TabIndex = 11;
			// 
			// dtpHasta
			// 
			this.dtpHasta.Format = System.Windows.Forms.DateTimePickerFormat.Short;
			this.dtpHasta.Location = new System.Drawing.Point(232, 24);
			this.dtpHasta.Name = "dtpHasta";
			this.dtpHasta.TabIndex = 12;
			// 
			// btnBuscar
			// 
			this.btnBuscar.Location = new System.Drawing.Point(680, 56);
			this.btnBuscar.Name = "btnBuscar";
			this.btnBuscar.Size = new System.Drawing.Size(80, 25);
			this.btnBuscar.TabIndex = 13;
			this.btnBuscar.Text = "Buscar";
			this.btnBuscar.Click += new System.EventHandler(this.BtnBuscar_Click);
			// 
			// btnSalir
			// 
			this.btnSalir.Location = new System.Drawing.Point(824, 80);
			this.btnSalir.Name = "btnSalir";
			this.btnSalir.Size = new System.Drawing.Size(80, 25);
			this.btnSalir.TabIndex = 11;
			this.btnSalir.Text = "Salir";
			this.btnSalir.Click += new System.EventHandler(this.btnSalir_Click);
			// 
			// frmRequerimientoTurnoSoporteConsulta
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(6, 15);
			this.ClientSize = new System.Drawing.Size(912, 400);
			this.Controls.Add(this.btnSalir);
			this.Controls.Add(this.groupBox1);
			this.Controls.Add(this.btnNuevo);
			this.Controls.Add(this.btnInhabilitar);
			this.Controls.Add(this.ugTurnos);
			this.MaximizeBox = false;
			this.Name = "frmRequerimientoTurnoSoporteConsulta";
			this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
			this.Text = "Consultar Turnos de Soporte";
			this.Load += new System.EventHandler(this.FrmRequerimientoTurnoSoporteConsulta_Load);
			((System.ComponentModel.ISupportInitialize)(this.ugTurnos)).EndInit();
			this.groupBox1.ResumeLayout(false);
			this.ResumeLayout(false);

		}


		private void FrmRequerimientoTurnoSoporteConsulta_Load(object sender, EventArgs e)
		{
			try
			{
				// VISUAL: fechas por defecto
				dtpDesde.Value = new DateTime(DateTime.Now.Year, DateTime.Now.Month, 1);
				dtpHasta.Value = dtpDesde.Value.AddMonths(1).AddDays(-1);

				_cn = new SqlConnection(frmRequerimiento.sconexionPoint);
				_srv = new RequerimientoTurnoSoporteService(_cn);

				// Llamada normal
				CargarTurnos(dtpDesde.Value, dtpHasta.Value, "");
			}
			catch (Exception ex)
			{
				MessageBox.Show("Error al inicializar el formulario: " + ex.Message);
			}
		}

		private void CargarTurnos(DateTime desde, DateTime hasta, string filtro)
		{
			try
			{
				ArrayList turnos = _srv.ObtenerTodosTurnos(desde, hasta, filtro);

				DataTable dt = new DataTable();
				dt.Columns.Add("ID", typeof(int));
				dt.Columns.Add("Departamento");
				dt.Columns.Add("Técnico");
				dt.Columns.Add("Tipo");
				dt.Columns.Add("FechaDesde");
				dt.Columns.Add("FechaHasta");
				dt.Columns.Add("Prioridad");
				dt.Columns.Add("Activo");

				foreach (RequerimientoTurnoSoporteModel t in turnos)
				{
					DataRow row = dt.NewRow();
					row["ID"] = t.idRequerimientoTurnoSoporte;
					row["Departamento"] = t.NombreDepartamento;
					row["Técnico"] = t.NombreTecnico;
					row["Tipo"] = t.TipoTurnoDescripcion;
					row["FechaDesde"] = t.FechaDesde.ToString("dd/MM/yyyy");
					row["FechaHasta"] = t.FechaHasta.ToString("dd/MM/yyyy");
					row["Prioridad"] = t.Prioridad;
					row["Activo"] = t.Activo ? "Sí" : "No";
					dt.Rows.Add(row);
				}

				ugTurnos.DataSource = dt;

				// Configurar columnas
				Infragistics.Win.UltraWinGrid.UltraGridBand band = ugTurnos.DisplayLayout.Bands[0];
				band.Columns["ID"].Hidden = true;
				band.Columns["Activo"].CellActivation = Infragistics.Win.UltraWinGrid.Activation.NoEdit;
				band.Columns["Departamento"].Width = 150;
				band.Columns["Técnico"].Width = 150;
				band.Columns["Tipo"].Width = 100;
				band.Columns["FechaDesde"].Width = 90;
				band.Columns["FechaHasta"].Width = 90;
			}
			catch (Exception ex)
			{
				MessageBox.Show("Error al cargar turnos: " + ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
			}
		}

		private void UgTurnos_DoubleClick(object sender, EventArgs e)
		{
			if (ugTurnos.Selected.Rows.Count == 0) return;
			int idTurno = Convert.ToInt32(ugTurnos.Selected.Rows[0].Cells["ID"].Value);
			EditarTurno(idTurno);
		}

		private void EditarTurno(int idTurno)
		{
			frmRequerimientoTurnoSoporte frm = new frmRequerimientoTurnoSoporte(idTurno);
			frm.ShowDialog();
			CargarTurnos(dtpDesde.Value, dtpHasta.Value, txtBuscar.Text.Trim());
		}

		private void BtnInhabilitar_Click(object sender, EventArgs e)
		{
			if (ugTurnos.Selected.Rows.Count == 0) return;
			int idTurno = Convert.ToInt32(ugTurnos.Selected.Rows[0].Cells["ID"].Value);

			if (MessageBox.Show("¿Desea inhabilitar este turno?", "Confirmar", MessageBoxButtons.YesNo) == DialogResult.Yes)
			{
				_srv.InhabilitarTurno(idTurno);
				CargarTurnos(dtpDesde.Value, dtpHasta.Value, txtBuscar.Text.Trim());
			}
		}

		private void BtnBuscar_Click(object sender, EventArgs e)
		{
			if (dtpDesde.Value > dtpHasta.Value)
			{
				MessageBox.Show("La fecha 'Desde' no puede ser mayor a la fecha 'Hasta'.", "Aviso", MessageBoxButtons.OK, MessageBoxIcon.Warning);
				return;
			}
			CargarTurnos(dtpDesde.Value, dtpHasta.Value, txtBuscar.Text.Trim());
		}

		private void BtnNuevo_Click(object sender, EventArgs e)
		{
			frmRequerimientoTurnoSoporte frm = new frmRequerimientoTurnoSoporte();
			frm.ShowDialog();
			CargarTurnos(dtpDesde.Value, dtpHasta.Value, txtBuscar.Text.Trim());
		}

		private void btnSalir_Click(object sender, System.EventArgs e)
		{
			this.Close();
		}

	}
}
