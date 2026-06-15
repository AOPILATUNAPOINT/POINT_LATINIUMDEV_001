using System;
using System.Windows.Forms;
using Latinium.Services.Pinpad;
using Latinium.Models.Pinpad;
using Latinium.Models.Pinpad.PP;
using Latinium.Models.Pinpad.Configuracion;
using Latinium.Services.Pinpad.Configuracion;
using Latinium.Services.Pinpad.PP;
using System.Collections; // para ArrayList

namespace Latinium
{
	public class frmPinPad : Form
	{
		private Button btnLeerTarjeta;
		private Button btnProcesarPago;
		private TextBox txtResultado;
		private TextBox txtMonto;
		private Label lblMonto;

		// INFO TARJETA
		private Label lblTarjeta;
		private Label lblBin;
		private Label lblRedC;
		private Label lblRedD;

		// OPCIONES
		private GroupBox grpOpciones;
		private RadioButton rbCorriente;
		private ComboBox comboBoxDiferidos;

		private PinpadService _service = new PinpadService();
		// Declaración de variable
		private PinpadDiferidoService _diferidoService;
		private PinpadConfiguracionRepository _pinpadConfiguracion;

		// VARIABLES INTERNAS
		private string redC = "0";
		private string redD = "0";
		private string tarjeta = "";
		private string bin = "";

		private TextBox txtPropina;
		private Label lblPropina;

		private TextBox txtBase0;
		private Label lblBase0;

		private TextBox txtBase15;
		private Label lblBase15;
		
		private RadioButton rbDiferido;

		private TextBox txtIVA15;
		private System.Windows.Forms.Label lblTipoTarjeta;
		private System.Windows.Forms.Label lblMensajeRespuesta;
		private Label lblIVA15;

		public frmPinPad()
		{
			_diferidoService = new PinpadDiferidoService(MenuLatinium.sconexionEmpresa);
			_pinpadConfiguracion = new PinpadConfiguracionRepository(MenuLatinium.sconexionEmpresa);
			InitializeComponent();
		}

		private void InitializeComponent()
		{
			System.Resources.ResourceManager resources = new System.Resources.ResourceManager(typeof(frmPinPad));
			this.rbDiferido = new System.Windows.Forms.RadioButton();
			this.lblBase15 = new System.Windows.Forms.Label();
			this.txtBase15 = new System.Windows.Forms.TextBox();
			this.lblBase0 = new System.Windows.Forms.Label();
			this.txtBase0 = new System.Windows.Forms.TextBox();
			this.lblPropina = new System.Windows.Forms.Label();
			this.txtPropina = new System.Windows.Forms.TextBox();
			this.comboBoxDiferidos = new System.Windows.Forms.ComboBox();
			this.lblMonto = new System.Windows.Forms.Label();
			this.txtMonto = new System.Windows.Forms.TextBox();
			this.btnLeerTarjeta = new System.Windows.Forms.Button();
			this.txtResultado = new System.Windows.Forms.TextBox();
			this.lblTarjeta = new System.Windows.Forms.Label();
			this.lblBin = new System.Windows.Forms.Label();
			this.lblRedC = new System.Windows.Forms.Label();
			this.lblRedD = new System.Windows.Forms.Label();
			this.grpOpciones = new System.Windows.Forms.GroupBox();
			this.rbCorriente = new System.Windows.Forms.RadioButton();
			this.btnProcesarPago = new System.Windows.Forms.Button();
			this.lblIVA15 = new System.Windows.Forms.Label();
			this.txtIVA15 = new System.Windows.Forms.TextBox();
			this.lblTipoTarjeta = new System.Windows.Forms.Label();
			this.lblMensajeRespuesta = new System.Windows.Forms.Label();
			this.grpOpciones.SuspendLayout();
			this.SuspendLayout();
			// 
			// rbDiferido
			// 
			this.rbDiferido.Location = new System.Drawing.Point(128, 24);
			this.rbDiferido.Name = "rbDiferido";
			this.rbDiferido.TabIndex = 1;
			this.rbDiferido.Text = "Diferido";
			this.rbDiferido.CheckedChanged += new System.EventHandler(this.rbDiferido_CheckedChanged);
			// 
			// lblBase15
			// 
			this.lblBase15.Location = new System.Drawing.Point(20, 45);
			this.lblBase15.Name = "lblBase15";
			this.lblBase15.TabIndex = 10;
			this.lblBase15.Text = "Base IVA 15%:";
			// 
			// txtBase15
			// 
			this.txtBase15.Location = new System.Drawing.Point(120, 45);
			this.txtBase15.Name = "txtBase15";
			this.txtBase15.ReadOnly = true;
			this.txtBase15.Size = new System.Drawing.Size(160, 20);
			this.txtBase15.TabIndex = 11;
			this.txtBase15.Text = "10.50";
			this.txtBase15.TextChanged += new System.EventHandler(this.txtBase15_TextChanged);
			// 
			// lblBase0
			// 
			this.lblBase0.Location = new System.Drawing.Point(20, 70);
			this.lblBase0.Name = "lblBase0";
			this.lblBase0.TabIndex = 12;
			this.lblBase0.Text = "Base 0%:";
			// 
			// txtBase0
			// 
			this.txtBase0.Location = new System.Drawing.Point(120, 70);
			this.txtBase0.Name = "txtBase0";
			this.txtBase0.ReadOnly = true;
			this.txtBase0.Size = new System.Drawing.Size(160, 20);
			this.txtBase0.TabIndex = 13;
			this.txtBase0.Text = "0.00";
			this.txtBase0.TextChanged += new System.EventHandler(this.txtBase0_TextChanged);
			// 
			// lblPropina
			// 
			this.lblPropina.Location = new System.Drawing.Point(20, 95);
			this.lblPropina.Name = "lblPropina";
			this.lblPropina.TabIndex = 14;
			this.lblPropina.Text = "Propina:";
			// 
			// txtPropina
			// 
			this.txtPropina.Location = new System.Drawing.Point(120, 95);
			this.txtPropina.Name = "txtPropina";
			this.txtPropina.ReadOnly = true;
			this.txtPropina.Size = new System.Drawing.Size(160, 20);
			this.txtPropina.TabIndex = 15;
			this.txtPropina.Text = "0.00";
			this.txtPropina.TextChanged += new System.EventHandler(this.txtPropina_TextChanged);
			// 
			// comboBoxDiferidos
			// 
			this.comboBoxDiferidos.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
			this.comboBoxDiferidos.Location = new System.Drawing.Point(288, 120);
			this.comboBoxDiferidos.Name = "comboBoxDiferidos";
			this.comboBoxDiferidos.Size = new System.Drawing.Size(352, 21);
			this.comboBoxDiferidos.TabIndex = 0;
			this.comboBoxDiferidos.SelectedIndexChanged += new System.EventHandler(this.comboBoxDiferidos_SelectedIndexChanged);
			// 
			// lblMonto
			// 
			this.lblMonto.Location = new System.Drawing.Point(20, 20);
			this.lblMonto.Name = "lblMonto";
			this.lblMonto.TabIndex = 0;
			this.lblMonto.Text = "Monto:";
			// 
			// txtMonto
			// 
			this.txtMonto.Location = new System.Drawing.Point(120, 16);
			this.txtMonto.Name = "txtMonto";
			this.txtMonto.Size = new System.Drawing.Size(160, 20);
			this.txtMonto.TabIndex = 1;
			this.txtMonto.Text = "10.50";
			this.txtMonto.TextChanged += new System.EventHandler(this.txtMonto_TextChanged);
			// 
			// btnLeerTarjeta
			// 
			this.btnLeerTarjeta.Font = new System.Drawing.Font("Microsoft Sans Serif", 15.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.btnLeerTarjeta.Image = ((System.Drawing.Image)(resources.GetObject("btnLeerTarjeta.Image")));
			this.btnLeerTarjeta.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnLeerTarjeta.Location = new System.Drawing.Point(648, 72);
			this.btnLeerTarjeta.Name = "btnLeerTarjeta";
			this.btnLeerTarjeta.Size = new System.Drawing.Size(272, 32);
			this.btnLeerTarjeta.TabIndex = 2;
			this.btnLeerTarjeta.Text = "Leer Tarjeta";
			this.btnLeerTarjeta.Click += new System.EventHandler(this.btnLeerTarjeta_Click);
			// 
			// txtResultado
			// 
			this.txtResultado.Location = new System.Drawing.Point(16, 168);
			this.txtResultado.Multiline = true;
			this.txtResultado.Name = "txtResultado";
			this.txtResultado.ScrollBars = System.Windows.Forms.ScrollBars.Vertical;
			this.txtResultado.Size = new System.Drawing.Size(904, 144);
			this.txtResultado.TabIndex = 3;
			this.txtResultado.Text = "";
			// 
			// lblTarjeta
			// 
			this.lblTarjeta.Location = new System.Drawing.Point(296, 16);
			this.lblTarjeta.Name = "lblTarjeta";
			this.lblTarjeta.Size = new System.Drawing.Size(344, 20);
			this.lblTarjeta.TabIndex = 4;
			this.lblTarjeta.Text = "Tarjeta:";
			// 
			// lblBin
			// 
			this.lblBin.Location = new System.Drawing.Point(296, 40);
			this.lblBin.Name = "lblBin";
			this.lblBin.Size = new System.Drawing.Size(344, 20);
			this.lblBin.TabIndex = 5;
			this.lblBin.Text = "BIN:";
			this.lblBin.Visible = false;
			// 
			// lblRedC
			// 
			this.lblRedC.Location = new System.Drawing.Point(296, 64);
			this.lblRedC.Name = "lblRedC";
			this.lblRedC.Size = new System.Drawing.Size(344, 20);
			this.lblRedC.TabIndex = 6;
			this.lblRedC.Text = "Red Corriente:";
			// 
			// lblRedD
			// 
			this.lblRedD.Location = new System.Drawing.Point(296, 88);
			this.lblRedD.Name = "lblRedD";
			this.lblRedD.Size = new System.Drawing.Size(344, 20);
			this.lblRedD.TabIndex = 7;
			this.lblRedD.Text = "Red Diferido:";
			// 
			// grpOpciones
			// 
			this.grpOpciones.Controls.Add(this.rbCorriente);
			this.grpOpciones.Controls.Add(this.rbDiferido);
			this.grpOpciones.Location = new System.Drawing.Point(648, 8);
			this.grpOpciones.Name = "grpOpciones";
			this.grpOpciones.Size = new System.Drawing.Size(272, 56);
			this.grpOpciones.TabIndex = 8;
			this.grpOpciones.TabStop = false;
			this.grpOpciones.Text = "Forma de Pago";
			// 
			// rbCorriente
			// 
			this.rbCorriente.Location = new System.Drawing.Point(10, 20);
			this.rbCorriente.Name = "rbCorriente";
			this.rbCorriente.TabIndex = 0;
			this.rbCorriente.Text = "Corriente";
			this.rbCorriente.CheckedChanged += new System.EventHandler(this.rbCorriente_CheckedChanged);
			// 
			// btnProcesarPago
			// 
			this.btnProcesarPago.Font = new System.Drawing.Font("Microsoft Sans Serif", 15.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.btnProcesarPago.Image = ((System.Drawing.Image)(resources.GetObject("btnProcesarPago.Image")));
			this.btnProcesarPago.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnProcesarPago.Location = new System.Drawing.Point(648, 112);
			this.btnProcesarPago.Name = "btnProcesarPago";
			this.btnProcesarPago.Size = new System.Drawing.Size(272, 32);
			this.btnProcesarPago.TabIndex = 9;
			this.btnProcesarPago.Text = "Procesar Pago";
			this.btnProcesarPago.Click += new System.EventHandler(this.btnProcesarPago_Click);
			// 
			// lblIVA15
			// 
			this.lblIVA15.Location = new System.Drawing.Point(20, 120);
			this.lblIVA15.Name = "lblIVA15";
			this.lblIVA15.Size = new System.Drawing.Size(100, 20);
			this.lblIVA15.TabIndex = 0;
			this.lblIVA15.Text = "IVA 15%:";
			// 
			// txtIVA15
			// 
			this.txtIVA15.Location = new System.Drawing.Point(120, 120);
			this.txtIVA15.Name = "txtIVA15";
			this.txtIVA15.ReadOnly = true;
			this.txtIVA15.Size = new System.Drawing.Size(160, 20);
			this.txtIVA15.TabIndex = 1;
			this.txtIVA15.Text = "0.00";
			this.txtIVA15.TextChanged += new System.EventHandler(this.txtIVA15_TextChanged);
			// 
			// lblTipoTarjeta
			// 
			this.lblTipoTarjeta.Location = new System.Drawing.Point(296, 40);
			this.lblTipoTarjeta.Name = "lblTipoTarjeta";
			this.lblTipoTarjeta.Size = new System.Drawing.Size(328, 20);
			this.lblTipoTarjeta.TabIndex = 16;
			this.lblTipoTarjeta.Text = "Tipo Tarjeta:";
			// 
			// lblMensajeRespuesta
			// 
			this.lblMensajeRespuesta.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.lblMensajeRespuesta.ForeColor = System.Drawing.Color.Red;
			this.lblMensajeRespuesta.Location = new System.Drawing.Point(512, 16);
			this.lblMensajeRespuesta.Name = "lblMensajeRespuesta";
			this.lblMensajeRespuesta.Size = new System.Drawing.Size(128, 20);
			this.lblMensajeRespuesta.TabIndex = 17;
			// 
			// frmPinPad
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 13);
			this.ClientSize = new System.Drawing.Size(928, 326);
			this.Controls.Add(this.lblMensajeRespuesta);
			this.Controls.Add(this.lblTipoTarjeta);
			this.Controls.Add(this.lblIVA15);
			this.Controls.Add(this.txtIVA15);
			this.Controls.Add(this.txtMonto);
			this.Controls.Add(this.txtResultado);
			this.Controls.Add(this.txtBase15);
			this.Controls.Add(this.txtBase0);
			this.Controls.Add(this.txtPropina);
			this.Controls.Add(this.comboBoxDiferidos);
			this.Controls.Add(this.lblMonto);
			this.Controls.Add(this.btnLeerTarjeta);
			this.Controls.Add(this.lblTarjeta);
			this.Controls.Add(this.lblBin);
			this.Controls.Add(this.lblRedC);
			this.Controls.Add(this.lblRedD);
			this.Controls.Add(this.grpOpciones);
			this.Controls.Add(this.btnProcesarPago);
			this.Controls.Add(this.lblBase15);
			this.Controls.Add(this.lblBase0);
			this.Controls.Add(this.lblPropina);
			this.Name = "frmPinPad";
			this.Text = "Simulador PinPad";
			this.Load += new System.EventHandler(this.frmPinPad_Load);
			this.grpOpciones.ResumeLayout(false);
			this.ResumeLayout(false);

		}
		// =========================
		// OPCIONES
		// =========================
		private void GenerarOpciones()
		{
			// Habilitar radio buttons solo si la red correspondiente existe
			rbCorriente.Enabled = (redC != "0");
			rbDiferido.Enabled = (redD != "0");

			// Selección automática
			if (redC != "0")
				rbCorriente.Checked = true;
			else if (redD != "0")
				rbDiferido.Checked = true;

			// Habilitar botón Procesar Pago solo después de leer tarjeta
			btnProcesarPago.Enabled = true;
		}

		// =========================
		// COMPRA
		// =========================

		private void btnProcesarPago_Click(object sender, EventArgs e)
		{
//			try
//			{
//				txtResultado.Clear();
//				txtResultado.AppendText("Procesando pago...\r\n");
//				decimal monto;
//				try { monto = decimal.Parse(txtMonto.Text); }
//				catch { throw new Exception("Monto inválido"); }
//
//				// ===== Crear request =====
//				ProcesoPagoRequest request = new ProcesoPagoRequest();
//				// ===== MONTO Y BASES =====
//				request.MontoTotal = monto;
//				request.BaseIVA = decimal.Parse(txtBase15.Text);
//				request.Base0 = decimal.Parse(txtBase0.Text);
//				request.IVA = decimal.Parse(txtIVA15.Text);
//				request.Propina = decimal.Parse(txtPropina.Text);
//				request.Usuario = MenuLatinium.stUsuario;
//				// ===== TIPO DE TRANSACCIÓN =====
//				request.TipoMensaje = "PP"; // Pago
//				if (rbCorriente.Checked)
//				{
//					request.TipoTransaccion = "01";
//				}
//				else if (rbDiferido.Checked)
//				{
//					request.TipoTransaccion = "02";
//				}
//				else
//				{
//					throw new Exception("Debe seleccionar tipo de transacción");
//				}
//
//				if (redC == null || redC.Trim() == "" || redD == null || redD.Trim() == "")
//				{
//					throw new Exception("Primero debe leer la tarjeta");
//				}
//
//				// ===== RED =====
//				request.CodigoRed = rbCorriente.Checked ? redC : redD;
//
//				// ===== DIFERIDO =====
//				if (rbDiferido.Checked && comboBoxDiferidos.SelectedItem != null)
//				{
//					PinpadDiferido d = (PinpadDiferido)comboBoxDiferidos.SelectedItem;
//					request.CodigoDiferido = d.Codigo;   // ajusta si tu modelo usa otro campo
//					request.PlazoDiferido = d.Plazo;
//					request.MesesGracia = d.Gracia;
//				}
//				else
//				{
//					request.CodigoDiferido = null;
//					request.PlazoDiferido = 0;
//					request.MesesGracia = 0;
//				}
//
//				// ===== CAJA / COMERCIO =====
//				// ===== OBTENER CONFIGURACIÓN =====
//				PinpadConfiguracion cfg = _pinpadConfiguracion.ObtenerConfiguracionPorUsuario(MenuLatinium.IdUsuario);
//
//				if (cfg == null)
//				{
//					throw new Exception("No existe configuración de PINPAD para el usuario");
//				}
//
//				// ===== ASIGNAR A REQUEST =====
//				request.NumeroCaja = cfg.NumeroCaja;
//
//				// IMPORTANTE (para tu trama PP)
//				request.MID = cfg.MID;
//				request.TID = cfg.TID;
//				request.CID = cfg.CID;
//
//				request.idPinpad_ProcesoPago = Guid.NewGuid();
//				// ===== COMPRA (si aplica) =====
//				request.idCompra = 0; // si luego lo manejas, aquí se asigna
//				// ===== PROCESAR =====
//				ProcesarPagoResponse resp = _service.ProcesarPago(request);
//				// ===== RESULTADO =====
//				if (resp.Success)
//				{
//					txtResultado.AppendText("APROBADO\r\n");
//					txtResultado.AppendText("Código: " + resp.CodigoRespuestaAutorizador + "\r\n");
//					txtResultado.AppendText("Mensaje: " + resp.CodigoRespuestaMensaje + "\r\n");
//				}
//				else
//				{
//					txtResultado.AppendText("ERROR\r\n");
//					txtResultado.AppendText(resp.ErrorMessage + "\r\n");
//				}
//				// ===== Recalcular partes =====
//				RecalcularPartes();
//			}
//			catch (Exception ex)
//			{
//				RecalcularPartes();
//				MessageBox.Show(ex.Message);
//			}

			ProcesarPago cliente = new ProcesarPago("http://192.168.2.21:5000/api/pp/procesar");

			string respuesta = cliente.Ejecutar(
				10.50m,
				9.00m,
				0.00m,
				1.50m,
				0.00m,
				"01",
				"admin",
				"C1",
				123
				);

			Console.WriteLine(respuesta);
		}

		private void frmPinPad_Load(object sender, System.EventArgs e)
		{
			// Al iniciar la pantalla
			btnProcesarPago.Enabled = false;
			rbCorriente.Enabled = false;
			rbDiferido.Enabled = false;
			comboBoxDiferidos.Enabled = false;
		}

		private void btnLeerTarjeta_Click(object sender, EventArgs e)
		{
			try
			{
				txtResultado.Clear();
				txtResultado.AppendText("Leyendo tarjeta...\r\n");
				decimal monto = ObtenerMonto();
				LecturaTarjetaResponse responseLT = EjecutarLecturaTarjeta(monto);
				ConsultaTarjetaResponse responseCT = EjecutarConsultaTarjeta(monto);
				MostrarDatos(responseLT, responseCT);
				GenerarOpciones();
				RecalcularPartes();
				txtResultado.AppendText("Tarjeta leída correctamente\r\n");
			}
			catch (Exception ex)
			{
				MessageBox.Show(ex.Message);
			}
		}

//		private void btnLeerTarjeta_Click(object sender, EventArgs e)
//		{
//			try
//			{
//				// =========================
//				// VALIDAR SESSION
//				// =========================
//				if (PinpadSession.ConfigActual == null)
//				{
//					MessageBox.Show("No existe configuración de PinPad en memoria.");
//					return;
//				}
//
//				// =========================
//				// VALIDAR DEVICE ID
//				// =========================
//				string deviceId = PinpadSession.ConfigActual.DeviceId;
//
//				if (deviceId == null || deviceId.Trim() == "")
//				{
//					MessageBox.Show("El DeviceId no está configurado. Verifique la configuración del PinPad.");
//					return;
//				}
//
//				// =========================
//				// VALIDAR MONTO
//				// =========================
//				decimal monto = ObtenerMonto();
//
//				if (monto <= 0)
//				{
//					MessageBox.Show("El monto debe ser mayor a 0.");
//					return;
//				}
//
//				// =========================
//				// LLAMADA A API LT
//				// =========================
//				ProcesarPago cliente = new ProcesarPago("http://192.168.2.21:5000/api/lt/leer");
//
//				string respuesta = cliente.EjecutarLT(monto, deviceId);
//
//				// =========================
//				// VALIDAR RESPUESTA
//				// =========================
//				if (respuesta == null || respuesta.Trim() == "")
//				{
//					MessageBox.Show("Respuesta vacía del PinPad.");
//					return;
//				}
//
//				if (respuesta.StartsWith("Error"))
//				{
//					MessageBox.Show(respuesta);
//					return;
//				}
//
//				// =========================
//				// OK
//				// =========================
//				MessageBox.Show("Lectura OK:\n" + respuesta);
//			}
//			catch (Exception ex)
//			{
//				MessageBox.Show("Error inesperado: " + ex.Message);
//			}
//		}
		private decimal ObtenerMonto()
		{
			decimal monto;

			try
			{
				monto = decimal.Parse(txtMonto.Text);
			}
			catch
			{
				throw new Exception("Monto inválido");
			}

			return monto;
		}

		private LecturaTarjetaResponse EjecutarLecturaTarjeta(decimal monto)
		{
			LecturaTarjetaRequest request = new LecturaTarjetaRequest();
			request.TipoMensaje = "LT";
			request.MontoTotal = monto;
			request.NumeroCaja = "0001";
			request.NumeroComercio = "0010";
			LecturaTarjetaResponse response = _service.LeerTarjeta(request);
			if (response.CodigoRespuesta != "00")
			{
				throw new Exception(response.MensajeRespuesta);
			}
			return response;
		}

		private ConsultaTarjetaResponse EjecutarConsultaTarjeta(decimal monto)
		{
			ConsultaTarjetaRequest request = new ConsultaTarjetaRequest();
			request.TipoMensaje = "CT";
			request.MontoTotal = monto;
			request.NumeroCaja = "0001";
			request.NumeroComercio = "0010";
			ConsultaTarjetaResponse response = _service.ConsultaTarjeta(request);
			if (response.CodigoRespuesta!= "00")
			{
				throw new Exception(response.MensajeRespuesta);
			}
			return response;
		}

		private string Safe(string value)
		{
			return value != null ? value : "";
		}

		private void MostrarDatos(LecturaTarjetaResponse lt, ConsultaTarjetaResponse ct)
		{
			tarjeta = lt.NumeroTarjetaTruncado;
			redC = lt.RedCorriente;
			redD = lt.RedDiferido;
			bin = ct.Bin;
			lblTarjeta.Text = "Tarjeta: " + tarjeta;
			lblBin.Text = "BIN: " + bin;
			lblRedC.Text = "Corriente: " + Safe(lt.RedCorrienteDescripcion);
			lblRedD.Text = "Diferido: " + Safe(lt.RedDiferidoDescripcion);
			lblTipoTarjeta.Text = "Tipo Tarjeta: " + Safe(ct.TipoTarjetaDescripcion);
			lblMensajeRespuesta.Text = Safe(ct.MensajeRespuesta);
		}

		private void MostrarDiferidos()
		{
			comboBoxDiferidos.Items.Clear();

			// Solo si hay red diferido válida
			if (redD == "0")
			{
				comboBoxDiferidos.Enabled = false;
				return;
			}

			comboBoxDiferidos.Enabled = true;

			// Obtener los diferidos reales
			ArrayList diferidos = _diferidoService.ObtenerDiferidosPorRed(redD);

			// VALIDAR DESPUÉS de obtenerlos
			if (diferidos == null || diferidos.Count == 0)
			{
				comboBoxDiferidos.Enabled = false;
				return;
			}

			foreach (PinpadDiferido d in diferidos)
			{
				comboBoxDiferidos.Items.Add(d);
			}

			if (comboBoxDiferidos.Items.Count > 0)
				comboBoxDiferidos.SelectedIndex = 0;
		}

		private void rbCorriente_CheckedChanged(object sender, EventArgs e)
		{
			if (rbCorriente.Checked)
			{
				comboBoxDiferidos.Enabled = false;
				comboBoxDiferidos.SelectedIndex = -1;
				comboBoxDiferidos.Text = "";

				
				// Recalcular según corriente
				RecalcularPartes();
			}
		}

		private void rbDiferido_CheckedChanged(object sender, EventArgs e)
		{
			if (rbDiferido.Checked)
			{
				if (redD == "0")
				{
					MessageBox.Show("La tarjeta no permite diferido");
					rbCorriente.Checked = true;
					return;
				}

				comboBoxDiferidos.Enabled = true;
				MostrarDiferidos();   // Llenar combo solo ahora
				RecalcularPartes();
			}
			else
			{
				comboBoxDiferidos.Enabled = false;
				comboBoxDiferidos.SelectedIndex = -1;
			}
		}

		private void comboBoxDiferidos_SelectedIndexChanged(object sender, EventArgs e)
		{
			if (comboBoxDiferidos.SelectedItem != null)
			{
				PinpadDiferido d = (PinpadDiferido)comboBoxDiferidos.SelectedItem;

				txtResultado.Text += "Diferido seleccionado: " 
					+ d.Plazo + " meses | Gracia: " + d.Gracia + "\r\n";

				RecalcularPartes();
			}
		}

		private void RecalcularPartes()
		{
			decimal total = 0;
			try
			{
				total = decimal.Parse(txtMonto.Text);
			}
			catch
			{
				txtResultado.AppendText("Monto inválido para recalcular\r\n");
				return;
			}

			// Base con IVA 15%
			decimal baseIVA = total / 1.15m;

			// IVA calculado
			decimal iva = baseIVA * 0.15m;

			// Propina
			decimal propina = 0; // si hay propina, se ajusta aquí

			// Base0 = total - baseIVA - iva - propina
			decimal base0 = total - baseIVA - iva - propina;
			if (base0 < 0) base0 = 0;

			// Asignar valores a TextBox con formato normal "0.00"
			txtBase15.Text = baseIVA.ToString("0.00");
			txtBase0.Text = base0.ToString("0.00");
			txtPropina.Text = propina.ToString("0.00");

			// Nuevo campo IVA, mismo formato que los demás
			txtIVA15.Text = iva.ToString("0.00");
		}

		private void txtMonto_TextChanged(object sender, EventArgs e)
		{
			RecalcularPartes();
		}

		private void txtBase15_TextChanged(object sender, System.EventArgs e)
		{
			RecalcularPartes();
		}

		private void txtBase0_TextChanged(object sender, System.EventArgs e)
		{
			RecalcularPartes();
		}

		private void txtPropina_TextChanged(object sender, System.EventArgs e)
		{
			RecalcularPartes();
		}

		private void txtIVA15_TextChanged(object sender, System.EventArgs e)
		{
			RecalcularPartes();
		}
	}
}