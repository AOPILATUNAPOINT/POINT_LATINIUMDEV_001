using System;
using System.Windows.Forms;
using System.Drawing;
using Latinium.Services.Pinpad.Configuracion;
using Latinium.Models.Pinpad.Configuracion;
using Latinium.Services.Transferencia;
using System.Data;             
using System.Data.SqlClient;   

namespace Latinium
{
	public class frmPinPadConfiguracionDetalle : Form
	{
		private TextBox txtDeviceId;
		private TextBox txtNombrePC;
		private ComboBox cmbBodega;
		private TextBox txtCaja;
		private TextBox txtIP;
		private TextBox txtPuerto;
		private TextBox txtTimeout;
		private TextBox txtMID;
		private TextBox txtTID;
		private TextBox txtCID;
		private CheckBox chkActivo;

		private Button btnGuardar;

		private Label lblDeviceId;
		private Label lblNombrePC;
		private Label lblBodega;
		private Label lblCaja;
		private Label lblIP;
		private Label lblPuerto;
		private Label lblTimeout;
		private Label lblMID;
		private Label lblTID;
		private Label lblCID;
		private Label lblActivo;

		private PinpadConfiguracionRepository repo;
		private Guid idEditar = Guid.Empty;

		public frmPinPadConfiguracionDetalle()
		{
			InitializeComponent();
		}

		public frmPinPadConfiguracionDetalle(Guid id)
		{
			idEditar = id;
			InitializeComponent();
		}

		// ======================================================
		// FORM SIMPLE (SIN LAYOUT DINÁMICO)
		// ======================================================
		private void InitializeComponent()
		{
			this.lblDeviceId = new System.Windows.Forms.Label();
			this.lblNombrePC = new System.Windows.Forms.Label();
			this.lblBodega = new System.Windows.Forms.Label();
			this.lblCaja = new System.Windows.Forms.Label();
			this.lblIP = new System.Windows.Forms.Label();
			this.lblPuerto = new System.Windows.Forms.Label();
			this.lblTimeout = new System.Windows.Forms.Label();
			this.lblMID = new System.Windows.Forms.Label();
			this.lblTID = new System.Windows.Forms.Label();
			this.lblCID = new System.Windows.Forms.Label();
			this.txtDeviceId = new System.Windows.Forms.TextBox();
			this.txtNombrePC = new System.Windows.Forms.TextBox();
			this.cmbBodega = new System.Windows.Forms.ComboBox();
			this.txtCaja = new System.Windows.Forms.TextBox();
			this.txtIP = new System.Windows.Forms.TextBox();
			this.txtPuerto = new System.Windows.Forms.TextBox();
			this.txtTimeout = new System.Windows.Forms.TextBox();
			this.txtMID = new System.Windows.Forms.TextBox();
			this.txtTID = new System.Windows.Forms.TextBox();
			this.txtCID = new System.Windows.Forms.TextBox();
			this.btnGuardar = new System.Windows.Forms.Button();
			this.lblActivo = new System.Windows.Forms.Label();
			this.chkActivo = new System.Windows.Forms.CheckBox();
			this.SuspendLayout();
			// 
			// lblDeviceId
			// 
			this.lblDeviceId.Location = new System.Drawing.Point(20, 20);
			this.lblDeviceId.Name = "lblDeviceId";
			this.lblDeviceId.TabIndex = 0;
			this.lblDeviceId.Text = "DeviceId";
			// 
			// lblNombrePC
			// 
			this.lblNombrePC.Location = new System.Drawing.Point(20, 50);
			this.lblNombrePC.Name = "lblNombrePC";
			this.lblNombrePC.TabIndex = 1;
			this.lblNombrePC.Text = "PC";
			// 
			// lblBodega
			// 
			this.lblBodega.Location = new System.Drawing.Point(20, 80);
			this.lblBodega.Name = "lblBodega";
			this.lblBodega.TabIndex = 2;
			this.lblBodega.Text = "Bodega";
			// 
			// lblCaja
			// 
			this.lblCaja.Location = new System.Drawing.Point(20, 110);
			this.lblCaja.Name = "lblCaja";
			this.lblCaja.TabIndex = 3;
			this.lblCaja.Text = "Caja";
			// 
			// lblIP
			// 
			this.lblIP.Location = new System.Drawing.Point(20, 140);
			this.lblIP.Name = "lblIP";
			this.lblIP.TabIndex = 4;
			this.lblIP.Text = "IP";
			// 
			// lblPuerto
			// 
			this.lblPuerto.Location = new System.Drawing.Point(20, 170);
			this.lblPuerto.Name = "lblPuerto";
			this.lblPuerto.TabIndex = 5;
			this.lblPuerto.Text = "Puerto";
			// 
			// lblTimeout
			// 
			this.lblTimeout.Location = new System.Drawing.Point(20, 200);
			this.lblTimeout.Name = "lblTimeout";
			this.lblTimeout.TabIndex = 6;
			this.lblTimeout.Text = "Timeout";
			// 
			// lblMID
			// 
			this.lblMID.Location = new System.Drawing.Point(20, 230);
			this.lblMID.Name = "lblMID";
			this.lblMID.TabIndex = 7;
			this.lblMID.Text = "MID";
			// 
			// lblTID
			// 
			this.lblTID.Location = new System.Drawing.Point(20, 260);
			this.lblTID.Name = "lblTID";
			this.lblTID.TabIndex = 8;
			this.lblTID.Text = "TID";
			// 
			// lblCID
			// 
			this.lblCID.Location = new System.Drawing.Point(20, 290);
			this.lblCID.Name = "lblCID";
			this.lblCID.TabIndex = 9;
			this.lblCID.Text = "CID";
			// 
			// txtDeviceId
			// 
			this.txtDeviceId.Location = new System.Drawing.Point(120, 20);
			this.txtDeviceId.Name = "txtDeviceId";
			this.txtDeviceId.Size = new System.Drawing.Size(200, 20);
			this.txtDeviceId.TabIndex = 10;
			this.txtDeviceId.Text = "";
			// 
			// txtNombrePC
			// 
			this.txtNombrePC.Location = new System.Drawing.Point(120, 48);
			this.txtNombrePC.Name = "txtNombrePC";
			this.txtNombrePC.Size = new System.Drawing.Size(200, 20);
			this.txtNombrePC.TabIndex = 11;
			this.txtNombrePC.Text = "";
			// 
			// cmbBodega
			// 
			this.cmbBodega.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
			this.cmbBodega.Location = new System.Drawing.Point(120, 80);
			this.cmbBodega.Name = "cmbBodega";
			this.cmbBodega.Size = new System.Drawing.Size(200, 21);
			this.cmbBodega.TabIndex = 12;
			// 
			// txtCaja
			// 
			this.txtCaja.Location = new System.Drawing.Point(120, 110);
			this.txtCaja.Name = "txtCaja";
			this.txtCaja.Size = new System.Drawing.Size(200, 20);
			this.txtCaja.TabIndex = 13;
			this.txtCaja.Text = "";
			// 
			// txtIP
			// 
			this.txtIP.Location = new System.Drawing.Point(120, 140);
			this.txtIP.Name = "txtIP";
			this.txtIP.Size = new System.Drawing.Size(200, 20);
			this.txtIP.TabIndex = 14;
			this.txtIP.Text = "";
			// 
			// txtPuerto
			// 
			this.txtPuerto.Location = new System.Drawing.Point(120, 170);
			this.txtPuerto.Name = "txtPuerto";
			this.txtPuerto.Size = new System.Drawing.Size(200, 20);
			this.txtPuerto.TabIndex = 15;
			this.txtPuerto.Text = "";
			// 
			// txtTimeout
			// 
			this.txtTimeout.Location = new System.Drawing.Point(120, 200);
			this.txtTimeout.Name = "txtTimeout";
			this.txtTimeout.Size = new System.Drawing.Size(200, 20);
			this.txtTimeout.TabIndex = 16;
			this.txtTimeout.Text = "";
			// 
			// txtMID
			// 
			this.txtMID.Location = new System.Drawing.Point(120, 230);
			this.txtMID.Name = "txtMID";
			this.txtMID.Size = new System.Drawing.Size(200, 20);
			this.txtMID.TabIndex = 17;
			this.txtMID.Text = "";
			// 
			// txtTID
			// 
			this.txtTID.Location = new System.Drawing.Point(120, 260);
			this.txtTID.Name = "txtTID";
			this.txtTID.Size = new System.Drawing.Size(200, 20);
			this.txtTID.TabIndex = 18;
			this.txtTID.Text = "";
			// 
			// txtCID
			// 
			this.txtCID.Location = new System.Drawing.Point(120, 290);
			this.txtCID.Name = "txtCID";
			this.txtCID.Size = new System.Drawing.Size(200, 20);
			this.txtCID.TabIndex = 19;
			this.txtCID.Text = "";
			// 
			// btnGuardar
			// 
			this.btnGuardar.Location = new System.Drawing.Point(120, 352);
			this.btnGuardar.Name = "btnGuardar";
			this.btnGuardar.TabIndex = 20;
			this.btnGuardar.Text = "Guardar";
			this.btnGuardar.Click += new System.EventHandler(this.btnGuardar_Click);
			// 
			// lblActivo
			// 
			this.lblActivo.Location = new System.Drawing.Point(20, 320);
			this.lblActivo.Name = "lblActivo";
			this.lblActivo.Size = new System.Drawing.Size(100, 20);
			this.lblActivo.TabIndex = 21;
			this.lblActivo.Text = "Activo";
			// 
			// chkActivo
			// 
			this.chkActivo.Checked = true;
			this.chkActivo.CheckState = System.Windows.Forms.CheckState.Checked;
			this.chkActivo.Location = new System.Drawing.Point(120, 320);
			this.chkActivo.Name = "chkActivo";
			this.chkActivo.TabIndex = 22;
			// 
			// frmPinPadConfiguracionDetalle
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 13);
			this.ClientSize = new System.Drawing.Size(368, 390);
			this.Controls.Add(this.lblDeviceId);
			this.Controls.Add(this.lblNombrePC);
			this.Controls.Add(this.lblBodega);
			this.Controls.Add(this.lblCaja);
			this.Controls.Add(this.lblIP);
			this.Controls.Add(this.lblPuerto);
			this.Controls.Add(this.lblTimeout);
			this.Controls.Add(this.lblMID);
			this.Controls.Add(this.lblTID);
			this.Controls.Add(this.lblCID);
			this.Controls.Add(this.txtDeviceId);
			this.Controls.Add(this.txtNombrePC);
			this.Controls.Add(this.txtCaja);
			this.Controls.Add(this.cmbBodega);
			this.Controls.Add(this.txtIP);
			this.Controls.Add(this.txtPuerto);
			this.Controls.Add(this.txtTimeout);
			this.Controls.Add(this.txtMID);
			this.Controls.Add(this.txtTID);
			this.Controls.Add(this.txtCID);
			this.Controls.Add(this.btnGuardar);
			this.Controls.Add(this.lblActivo);
			this.Controls.Add(this.chkActivo);
			this.Name = "frmPinPadConfiguracionDetalle";
			this.Load += new System.EventHandler(this.frmPinPadConfiguracionDetalle_Load);
			this.ResumeLayout(false);

		}

		// ======================================================
		// CARGAR DATOS
		// ======================================================
		private void CargarDatos()
		{
			if (idEditar == Guid.Empty)
				return;

			PinpadConfiguracion cfg = repo.ObtenerPorId(idEditar);

			if (cfg == null)
			{
				MessageBox.Show("No se encontró la configuración.");
				return;
			}

			txtDeviceId.Text = cfg.DeviceId;
			txtNombrePC.Text = cfg.NombrePC;
			txtCaja.Text = cfg.NumeroCaja;
			txtIP.Text = cfg.IpPinPad;
			txtPuerto.Text = cfg.PuertoPinPad.ToString();
			txtTimeout.Text = cfg.TimeoutPinPad.ToString();
			txtMID.Text = cfg.MID;
			txtTID.Text = cfg.TID;
			txtCID.Text = cfg.CID;
			chkActivo.Checked = cfg.Activo; // FALTA ESTO

			// SELECCIÓN SEGURA
			for (int i = 0; i < cmbBodega.Items.Count; i++)
			{
				DataRowView row = (DataRowView)cmbBodega.Items[i];

				if (Convert.ToInt32(row["Bodega"]) == cfg.Bodega)
				{
					cmbBodega.SelectedIndex = i;
					break;
				}
			}

			// BLOQUEAR CAMPOS EN EDICIÓN
			if (idEditar != Guid.Empty)
			{
				txtDeviceId.ReadOnly = true;
				txtNombrePC.ReadOnly = true;
				txtCaja.ReadOnly = true;
			}
		}

		// ======================================================
		// GUARDAR
		// ======================================================
		private void btnGuardar_Click(object sender, EventArgs e)
		{
			try
			{
				PinpadConfiguracion cfg = new PinpadConfiguracion();

				cfg.DeviceId = txtDeviceId.Text;
				cfg.NombrePC = txtNombrePC.Text;
				cfg.Bodega = (cmbBodega.SelectedValue == null)
					? 0
					: Convert.ToInt32(cmbBodega.SelectedValue);
				cfg.NumeroCaja = txtCaja.Text;
				cfg.IpPinPad = txtIP.Text;
				cfg.PuertoPinPad = SafeInt(txtPuerto.Text, 9999);
				cfg.TimeoutPinPad = SafeInt(txtTimeout.Text, 15000);
				cfg.MID = txtMID.Text;
				cfg.TID = txtTID.Text;
				cfg.CID = txtCID.Text;

				cfg.Activo = chkActivo.Checked;
				cfg.Usuario = MenuLatinium.stUsuario;

				if (idEditar != Guid.Empty)
				{
					cfg.IdPinpadConfiguracion = idEditar;
					repo.Actualizar(cfg);
				}
				else
				{
					repo.Insertar(cfg);
				}

				MessageBox.Show("Guardado correctamente");
				this.Close();
			}
			catch (Exception ex)
			{
				MessageBox.Show("Error: " + ex.Message);
			}
		}

		// ======================================================
		// SAFE INT (.NET 1.1)
		// ======================================================
		private int SafeInt(string value)
		{
			try { return Convert.ToInt32(value); }
			catch { return 0; }
		}


		private void CargarBodegas()
		{

			ReporteTransferenciaService service = new ReporteTransferenciaService(MenuLatinium.sconexionEmpresa);

			// ============================
			// BODEGAS DESDE SERVICE
			// ============================
			int bodegaAsignada = 0;

			DataTable dtBodega = service.ObtenerBodegasPorUsuario(
				MenuLatinium.IdUsuario,
				1,
				DateTime.Today,
				false,
				out bodegaAsignada
				);

			cmbBodega.DataSource = dtBodega;
			cmbBodega.DisplayMember = "Nombre";
			cmbBodega.ValueMember = "Bodega";

			// LÓGICA DE NEGOCIO
			if (bodegaAsignada > 0)
			{
				for (int i = 0; i < cmbBodega.Items.Count; i++)
				{
					DataRowView row = (DataRowView)cmbBodega.Items[i];

					if (Convert.ToInt32(row["Bodega"]) == bodegaAsignada)
					{
						cmbBodega.SelectedIndex = i;
						break;
					}
				}

				cmbBodega.Enabled = false;
			}
			else
			{
				cmbBodega.Enabled = true;
			}
		}

		private void frmPinPadConfiguracionDetalle_Load(object sender, System.EventArgs e)
		{
			repo = new PinpadConfiguracionRepository(MenuLatinium.sconexionEmpresa);
			CargarBodegas();  // 1. llena combo
			CargarDatos();    // 2. selecciona valor
		}

		private int SafeInt(string value, int defaultValue)
		{
			try { return Convert.ToInt32(value); }
			catch { return defaultValue; }
		}
	}
}