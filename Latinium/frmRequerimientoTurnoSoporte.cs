using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using Latinium.Models.Requerimiento.TurnoSoporte;
using Latinium.Services.Requerimiento.TurnoSoporte;
using System.Data;
using System.Data.SqlClient;

namespace Latinium
{
	/// <summary>
	/// Descripción breve de frmRequerimientoTurnoSoporte.
	/// </summary>
	public class frmRequerimientoTurnoSoporte : System.Windows.Forms.Form
	{
		/// <summary>
		/// Variable del diseñador requerida.
		/// </summary>
		private System.ComponentModel.Container components = null;

		// ===== CONTROLES =====
		private Label lblTecnico;
		private ComboBox cmbTecnico;

		private Label lblDepartamento;
		private ComboBox cmbDepartamento;

		private Label lblTipo;
		private ComboBox cmbTipo;

		private Label lblDesde;
		private DateTimePicker dtpDesde;

		private Label lblHasta;
		private DateTimePicker dtpHasta;

		private Label lblPrioridad;
		private NumericUpDown nudPrioridad;

		private CheckBox chkActivo;

		private Button btnGuardar;
		private Button btnSalir;

		// ===== CONTROL INTERNO =====
		private int _idTurno = 0;
		private bool _cargando = false;

		private TextBox txtIdTurno;

		public frmRequerimientoTurnoSoporte()
		{
			//
			// Necesario para admitir el Diseñador de Windows Forms
			//
			InitializeComponent();

			//
			// TODO: agregar código de constructor después de llamar a InitializeComponent
			//
		}

		public frmRequerimientoTurnoSoporte(int idTurno)
		{
			InitializeComponent();
			_idTurno = idTurno;
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
			this.lblTecnico = new System.Windows.Forms.Label();
			this.cmbTecnico = new System.Windows.Forms.ComboBox();
			this.lblDepartamento = new System.Windows.Forms.Label();
			this.cmbDepartamento = new System.Windows.Forms.ComboBox();
			this.lblTipo = new System.Windows.Forms.Label();
			this.cmbTipo = new System.Windows.Forms.ComboBox();
			this.lblDesde = new System.Windows.Forms.Label();
			this.dtpDesde = new System.Windows.Forms.DateTimePicker();
			this.lblHasta = new System.Windows.Forms.Label();
			this.dtpHasta = new System.Windows.Forms.DateTimePicker();
			this.lblPrioridad = new System.Windows.Forms.Label();
			this.nudPrioridad = new System.Windows.Forms.NumericUpDown();
			this.chkActivo = new System.Windows.Forms.CheckBox();
			this.btnGuardar = new System.Windows.Forms.Button();
			this.btnSalir = new System.Windows.Forms.Button();
			((System.ComponentModel.ISupportInitialize)(this.nudPrioridad)).BeginInit();
			this.SuspendLayout();
			// 
			// lblTecnico
			// 
			this.lblTecnico.Location = new System.Drawing.Point(20, 48);
			this.lblTecnico.Name = "lblTecnico";
			this.lblTecnico.TabIndex = 0;
			this.lblTecnico.Text = "Técnico:";
			// 
			// cmbTecnico
			// 
			this.cmbTecnico.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
			this.cmbTecnico.Location = new System.Drawing.Point(150, 48);
			this.cmbTecnico.Name = "cmbTecnico";
			this.cmbTecnico.Size = new System.Drawing.Size(320, 24);
			this.cmbTecnico.TabIndex = 1;
			// 
			// lblDepartamento
			// 
			this.lblDepartamento.Location = new System.Drawing.Point(24, 16);
			this.lblDepartamento.Name = "lblDepartamento";
			this.lblDepartamento.TabIndex = 2;
			this.lblDepartamento.Text = "Departamento:";
			// 
			// cmbDepartamento
			// 
			this.cmbDepartamento.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
			this.cmbDepartamento.Location = new System.Drawing.Point(150, 16);
			this.cmbDepartamento.Name = "cmbDepartamento";
			this.cmbDepartamento.Size = new System.Drawing.Size(320, 24);
			this.cmbDepartamento.TabIndex = 3;
			this.cmbDepartamento.SelectedIndexChanged += new System.EventHandler(this.cmbDepartamento_SelectedIndexChanged);
			// 
			// lblTipo
			// 
			this.lblTipo.Location = new System.Drawing.Point(20, 80);
			this.lblTipo.Name = "lblTipo";
			this.lblTipo.Size = new System.Drawing.Size(100, 32);
			this.lblTipo.TabIndex = 4;
			this.lblTipo.Text = "Tipo Requerimiento:";
			// 
			// cmbTipo
			// 
			this.cmbTipo.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
			this.cmbTipo.Location = new System.Drawing.Point(150, 88);
			this.cmbTipo.Name = "cmbTipo";
			this.cmbTipo.Size = new System.Drawing.Size(320, 24);
			this.cmbTipo.TabIndex = 5;
			// 
			// lblDesde
			// 
			this.lblDesde.Location = new System.Drawing.Point(20, 125);
			this.lblDesde.Name = "lblDesde";
			this.lblDesde.TabIndex = 6;
			this.lblDesde.Text = "Desde:";
			// 
			// dtpDesde
			// 
			this.dtpDesde.Format = System.Windows.Forms.DateTimePickerFormat.Short;
			this.dtpDesde.Location = new System.Drawing.Point(150, 123);
			this.dtpDesde.Name = "dtpDesde";
			this.dtpDesde.TabIndex = 7;
			// 
			// lblHasta
			// 
			this.lblHasta.Location = new System.Drawing.Point(20, 160);
			this.lblHasta.Name = "lblHasta";
			this.lblHasta.TabIndex = 8;
			this.lblHasta.Text = "Hasta:";
			// 
			// dtpHasta
			// 
			this.dtpHasta.Format = System.Windows.Forms.DateTimePickerFormat.Short;
			this.dtpHasta.Location = new System.Drawing.Point(150, 158);
			this.dtpHasta.Name = "dtpHasta";
			this.dtpHasta.TabIndex = 9;
			// 
			// lblPrioridad
			// 
			this.lblPrioridad.Location = new System.Drawing.Point(20, 195);
			this.lblPrioridad.Name = "lblPrioridad";
			this.lblPrioridad.TabIndex = 10;
			this.lblPrioridad.Text = "Prioridad:";
			// 
			// nudPrioridad
			// 
			this.nudPrioridad.Location = new System.Drawing.Point(150, 193);
			this.nudPrioridad.Name = "nudPrioridad";
			this.nudPrioridad.TabIndex = 11;
			// 
			// chkActivo
			// 
			this.chkActivo.Checked = true;
			this.chkActivo.CheckState = System.Windows.Forms.CheckState.Checked;
			this.chkActivo.Location = new System.Drawing.Point(150, 225);
			this.chkActivo.Name = "chkActivo";
			this.chkActivo.TabIndex = 12;
			this.chkActivo.Text = "Activo";
			// 
			// btnGuardar
			// 
			this.btnGuardar.Location = new System.Drawing.Point(16, 256);
			this.btnGuardar.Name = "btnGuardar";
			this.btnGuardar.TabIndex = 13;
			this.btnGuardar.Text = "Guardar";
			this.btnGuardar.Click += new System.EventHandler(this.btnGuardar_Click);
			// 
			// btnSalir
			// 
			this.btnSalir.Location = new System.Drawing.Point(104, 256);
			this.btnSalir.Name = "btnSalir";
			this.btnSalir.TabIndex = 15;
			this.btnSalir.Text = "Salir";
			this.btnSalir.Click += new System.EventHandler(this.btnSalir_Click);
			// 
			// frmRequerimientoTurnoSoporte
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(6, 15);
			this.ClientSize = new System.Drawing.Size(496, 288);
			this.Controls.Add(this.lblTecnico);
			this.Controls.Add(this.cmbTecnico);
			this.Controls.Add(this.lblDepartamento);
			this.Controls.Add(this.cmbDepartamento);
			this.Controls.Add(this.lblTipo);
			this.Controls.Add(this.cmbTipo);
			this.Controls.Add(this.lblDesde);
			this.Controls.Add(this.dtpDesde);
			this.Controls.Add(this.lblHasta);
			this.Controls.Add(this.dtpHasta);
			this.Controls.Add(this.lblPrioridad);
			this.Controls.Add(this.nudPrioridad);
			this.Controls.Add(this.chkActivo);
			this.Controls.Add(this.btnGuardar);
			this.Controls.Add(this.btnSalir);
			this.MaximizeBox = false;
			this.Name = "frmRequerimientoTurnoSoporte";
			this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
			this.Text = "Turno de Soporte";
			this.Load += new System.EventHandler(this.frmRequerimientoTurnoSoporte_Load);
			((System.ComponentModel.ISupportInitialize)(this.nudPrioridad)).EndInit();
			this.ResumeLayout(false);

		}
		#endregion

		private RequerimientoTurnoSoporteService _srv;

		private void frmRequerimientoTurnoSoporte_Load(object sender, EventArgs e)
		{
			_cargando = true;

			SqlConnection cn =
				new SqlConnection(frmRequerimiento.sconexionPoint);

			_srv = new RequerimientoTurnoSoporteService(cn);

			cmbDepartamento.DataSource = _srv.ObtenerDepartamentos();
			cmbDepartamento.DisplayMember = "Descripcion";
			cmbDepartamento.ValueMember = "idRequerimientoDepartamento";

			cmbTecnico.Enabled = false;
			cmbTipo.Enabled = false;

			_cargando = false;

			if (_idTurno > 0)
				CargarParaEdicion(_idTurno);
		}




		private void Limpiar()
		{
			_idTurno = 0;

			cmbDepartamento.SelectedIndex = -1;
			cmbTecnico.SelectedIndex = -1;
			cmbTipo.SelectedIndex = -1;

			cmbTecnico.Enabled = false;
			cmbTipo.Enabled = false;

			nudPrioridad.Value = 0;
			chkActivo.Checked = true;

			dtpDesde.Value = DateTime.Today;
			dtpHasta.Value = DateTime.Today;
		}


		private void btnNuevo_Click(object sender, System.EventArgs e)
		{
			Limpiar();
		}

		private void btnSalir_Click(object sender, System.EventArgs e)
		{
			this.Close();
		}

		private void btnGuardar_Click(object sender, EventArgs e)
		{

			if (cmbDepartamento.SelectedIndex == -1 ||
				cmbTecnico.SelectedIndex == -1 ||
				cmbTipo.SelectedIndex == -1)
			{
				MessageBox.Show("Seleccione Departamento, Técnico y Tipo",
					"Validación", MessageBoxButtons.OK, MessageBoxIcon.Warning);
				return;
			}

			RequerimientoTurnoSoporteModel m =
				new RequerimientoTurnoSoporteModel();

			m.idRequerimientoTurnoSoporte = _idTurno;
			m.IdRequerimientoDepartamento = Convert.ToInt32(cmbDepartamento.SelectedValue);
			m.IdRequerimientoPerTecnico = Convert.ToInt32(cmbTecnico.SelectedValue);
			m.IdRequerimientoTipo = Convert.ToInt32(cmbTipo.SelectedValue);
			m.FechaDesde = dtpDesde.Value;
			m.FechaHasta = dtpHasta.Value;
			m.Prioridad = Convert.ToInt32(nudPrioridad.Value);
			m.Activo = chkActivo.Checked;

			_srv.Guardar(m);

			MessageBox.Show("Guardado correctamente");
		}

		private void CargarParaEdicion(int idTurno)
		{
			_cargando = true;

			RequerimientoTurnoSoporteModel m = _srv.ObtenerPorId(idTurno);

			if (m == null)
			{
				MessageBox.Show("Registro no encontrado",
					"Advertencia", MessageBoxButtons.OK,
					MessageBoxIcon.Warning);
				_cargando = false;
				return;
			}

			// Asegurarse de que el combo de departamento tenga los datos
			cmbDepartamento.DataSource = _srv.ObtenerDepartamentos();
			cmbDepartamento.DisplayMember = "NombreDepartamento";
			cmbDepartamento.ValueMember = "IdRequerimientoDepartamento";

			// Setear el departamento primero
			cmbDepartamento.SelectedValue = m.IdRequerimientoDepartamento;
			cmbDepartamento.Refresh(); // Fuerza que se actualice el combo

			// Cargar combos dependientes manualmente para evitar problemas con SelectedIndexChanged
			cmbTecnico.DataSource = _srv.ObtenerTecnicosPorDepartamento(m.IdRequerimientoDepartamento);
			cmbTecnico.DisplayMember = "Nombre";
			cmbTecnico.ValueMember = "idRequerimientoPerTecnico";
			cmbTecnico.SelectedValue = m.IdRequerimientoPerTecnico;

			cmbTipo.DataSource = _srv.ObtenerTiposPorDepartamento(m.IdRequerimientoDepartamento);
			cmbTipo.DisplayMember = "Descripcion";
			cmbTipo.ValueMember = "idRequerimientoTipo";
			cmbTipo.SelectedValue = m.IdRequerimientoTipo;

			dtpDesde.Value = m.FechaDesde;
			dtpHasta.Value = m.FechaHasta;
			nudPrioridad.Value = m.Prioridad;
			chkActivo.Checked = m.Activo;

			_cargando = false; // Solo al final
		}

		private void cmbDepartamento_SelectedIndexChanged(object sender, EventArgs e)
		{
			if (_cargando || cmbDepartamento.SelectedIndex == -1)
				return;

			cmbTecnico.DataSource = null;
			cmbTipo.DataSource = null;

			int id = Convert.ToInt32(cmbDepartamento.SelectedValue);

			cmbTecnico.DataSource = _srv.ObtenerTecnicosPorDepartamento(id);
			cmbTecnico.DisplayMember = "Nombre";
			cmbTecnico.ValueMember = "idRequerimientoPerTecnico";
			cmbTecnico.SelectedIndex = -1;

			cmbTipo.DataSource = _srv.ObtenerTiposPorDepartamento(id);
			cmbTipo.DisplayMember = "Descripcion";
			cmbTipo.ValueMember = "idRequerimientoTipo";
			cmbTipo.SelectedIndex = -1;

			cmbTecnico.Enabled = true;
			cmbTipo.Enabled = true;
		}
	}
}
