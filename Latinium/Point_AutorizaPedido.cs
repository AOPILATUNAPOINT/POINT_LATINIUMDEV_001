using System;
using System.Data;
using System.Data.SqlClient;
using C1.Data;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using LibContabilidad.DataObjects;
using CrystalDecisions.Shared;
using Infragistics.Win.UltraWinGrid;
using System.Text;


namespace Latinium
{
	/// <summary>
	/// Descripción breve de Point_AutorizaPedido.
	/// </summary>
	public class Point_AutorizaPedido : System.Windows.Forms.Form
	{
		private System.Windows.Forms.Button btnCancelar;
		private System.Windows.Forms.Button btnAceptar;
		private System.Windows.Forms.Label label1;
		private System.Windows.Forms.Label label4;
		private System.Windows.Forms.Label label7;
		private System.Windows.Forms.Label label9;
		private System.Windows.Forms.GroupBox groupBox1;
		private System.Windows.Forms.GroupBox groupBox2;
		private System.Windows.Forms.GroupBox groupBox3;
		private System.Windows.Forms.Label label10;
		private System.Windows.Forms.Label label12;
		private System.Windows.Forms.Label label14;
		private System.Windows.Forms.Label label16;
		private System.Windows.Forms.Label label18;
		private System.Windows.Forms.GroupBox groupBox4;
		private System.Windows.Forms.Label label22;
		private System.Windows.Forms.Label label24;
		private System.Windows.Forms.Label label28;
		private System.Windows.Forms.Label label30;
		private System.Windows.Forms.Label label35;
		private System.Windows.Forms.GroupBox groupBox5;
		private System.Windows.Forms.GroupBox groupBox6;
		public Infragistics.Win.UltraWinEditors.UltraTextEditor txtNotas;
		private System.Windows.Forms.Label label39;
		private System.Windows.Forms.Label label34;
		private System.Windows.Forms.Label label40;
		private System.Windows.Forms.Label lblNombreClte;
		private System.Windows.Forms.Label lblCodigo;
		private System.Windows.Forms.Label lblRuc;
		private System.Windows.Forms.Label lblFIngreso;
		private System.Windows.Forms.Label lblLimeteCredito;
		private System.Windows.Forms.Label lblSaldoActual;
		private System.Windows.Forms.Label lblDisponible;
		private System.Windows.Forms.Label lblCheqDevueltos;
		private System.Windows.Forms.Label lblFechaChDevuelto;
		private System.Windows.Forms.Label lblBanco;
		private System.Windows.Forms.Label lblValChqDevuelto;
		private System.Windows.Forms.Label lblFcxPagar;
		private System.Windows.Forms.Label lblNFcAntigua;
		private System.Windows.Forms.Label lblFchFcAntigua;
		private System.Windows.Forms.Label lblValFcAntigua;
		private System.Windows.Forms.Label lblUltimaFc;
		private System.Windows.Forms.Label lblFchUltimaFc;
		private System.Windows.Forms.Label lblValUlmimaFc;
		private System.Windows.Forms.Label lblUltimoPago;
		private System.Windows.Forms.Label lblFchUltimoPago;
		private System.Windows.Forms.Label lblValUltimoPago;
		private System.Windows.Forms.Label lblDiasCredito;
		private System.Windows.Forms.Label lblDiasPago;
		private System.Windows.Forms.Label lblDiasMora;
		private System.Windows.Forms.Label label2;
		private System.Windows.Forms.Label label5;
		private System.Windows.Forms.Label lblDiasCred2;
		private System.Windows.Forms.Label lblDiasCred3;
		private System.Windows.Forms.Label lblFormaPago;
		public Infragistics.Win.UltraWinEditors.UltraTextEditor txtDocumentos;
		private System.Windows.Forms.ErrorProvider errorProvider;
		/// <summary>
		/// Variable del diseñador requerida.
		/// </summary>
		private System.ComponentModel.Container components = null;
		
		public System.Windows.Forms.Label lblEstado;
		private System.Windows.Forms.Button btRechazar;		
		private C1.Data.C1DataSet cdsSeteoF;
		private System.Windows.Forms.Label label6;
		private System.Windows.Forms.Label label21;
		public Infragistics.Win.UltraWinSchedule.UltraCalendarCombo dtFecha;
		public Infragistics.Win.UltraWinEditors.UltraNumericEditor txtNumCuotas;

		int ICliente = 0;
		int IPedido = 0;
		int IEstado = 0;		
		public int iRetEstado = 0;
		public int iContSerial = 0;	
		int iCuotas = 0;
		DateTime dtFechaPrimerPago = DateTime.Today;
		string sNotas = "";
	
		public Point_AutorizaPedido(int iCliente, int iPedido, int iEstado, int ICuotas, DateTime DTFechaPrimerPago, string SNotas)
		{
			InitializeComponent();
			ICliente = iCliente;
			IPedido = iPedido;
			IEstado = iEstado;
			sNotas = SNotas;
			iCuotas = ICuotas;
			dtFechaPrimerPago = DTFechaPrimerPago;
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
			System.Resources.ResourceManager resources = new System.Resources.ResourceManager(typeof(Point_AutorizaPedido));
			Infragistics.Win.Appearance appearance1 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance2 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance3 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton dateButton1 = new Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton();
			Infragistics.Win.Appearance appearance4 = new Infragistics.Win.Appearance();
			this.btnCancelar = new System.Windows.Forms.Button();
			this.btnAceptar = new System.Windows.Forms.Button();
			this.btRechazar = new System.Windows.Forms.Button();
			this.lblNombreClte = new System.Windows.Forms.Label();
			this.label1 = new System.Windows.Forms.Label();
			this.lblCodigo = new System.Windows.Forms.Label();
			this.lblRuc = new System.Windows.Forms.Label();
			this.label4 = new System.Windows.Forms.Label();
			this.lblFIngreso = new System.Windows.Forms.Label();
			this.label7 = new System.Windows.Forms.Label();
			this.lblEstado = new System.Windows.Forms.Label();
			this.label9 = new System.Windows.Forms.Label();
			this.groupBox1 = new System.Windows.Forms.GroupBox();
			this.groupBox2 = new System.Windows.Forms.GroupBox();
			this.groupBox3 = new System.Windows.Forms.GroupBox();
			this.lblValChqDevuelto = new System.Windows.Forms.Label();
			this.lblBanco = new System.Windows.Forms.Label();
			this.label18 = new System.Windows.Forms.Label();
			this.lblFechaChDevuelto = new System.Windows.Forms.Label();
			this.label16 = new System.Windows.Forms.Label();
			this.lblCheqDevueltos = new System.Windows.Forms.Label();
			this.label14 = new System.Windows.Forms.Label();
			this.lblDisponible = new System.Windows.Forms.Label();
			this.label12 = new System.Windows.Forms.Label();
			this.lblSaldoActual = new System.Windows.Forms.Label();
			this.label10 = new System.Windows.Forms.Label();
			this.lblLimeteCredito = new System.Windows.Forms.Label();
			this.groupBox4 = new System.Windows.Forms.GroupBox();
			this.lblFormaPago = new System.Windows.Forms.Label();
			this.lblDiasCred3 = new System.Windows.Forms.Label();
			this.label5 = new System.Windows.Forms.Label();
			this.lblDiasCred2 = new System.Windows.Forms.Label();
			this.label2 = new System.Windows.Forms.Label();
			this.lblDiasMora = new System.Windows.Forms.Label();
			this.lblDiasPago = new System.Windows.Forms.Label();
			this.label40 = new System.Windows.Forms.Label();
			this.label34 = new System.Windows.Forms.Label();
			this.label39 = new System.Windows.Forms.Label();
			this.lblValUltimoPago = new System.Windows.Forms.Label();
			this.lblUltimoPago = new System.Windows.Forms.Label();
			this.lblFchUltimoPago = new System.Windows.Forms.Label();
			this.label35 = new System.Windows.Forms.Label();
			this.lblValUlmimaFc = new System.Windows.Forms.Label();
			this.lblFchUltimaFc = new System.Windows.Forms.Label();
			this.label30 = new System.Windows.Forms.Label();
			this.lblUltimaFc = new System.Windows.Forms.Label();
			this.label28 = new System.Windows.Forms.Label();
			this.lblDiasCredito = new System.Windows.Forms.Label();
			this.lblValFcAntigua = new System.Windows.Forms.Label();
			this.lblNFcAntigua = new System.Windows.Forms.Label();
			this.label24 = new System.Windows.Forms.Label();
			this.lblFchFcAntigua = new System.Windows.Forms.Label();
			this.label22 = new System.Windows.Forms.Label();
			this.lblFcxPagar = new System.Windows.Forms.Label();
			this.groupBox5 = new System.Windows.Forms.GroupBox();
			this.txtNotas = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.groupBox6 = new System.Windows.Forms.GroupBox();
			this.txtDocumentos = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.errorProvider = new System.Windows.Forms.ErrorProvider();
			this.cdsSeteoF = new C1.Data.C1DataSet();
			this.label6 = new System.Windows.Forms.Label();
			this.label21 = new System.Windows.Forms.Label();
			this.dtFecha = new Infragistics.Win.UltraWinSchedule.UltraCalendarCombo();
			this.txtNumCuotas = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.groupBox1.SuspendLayout();
			this.groupBox3.SuspendLayout();
			this.groupBox4.SuspendLayout();
			this.groupBox5.SuspendLayout();
			((System.ComponentModel.ISupportInitialize)(this.txtNotas)).BeginInit();
			this.groupBox6.SuspendLayout();
			((System.ComponentModel.ISupportInitialize)(this.txtDocumentos)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.dtFecha)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNumCuotas)).BeginInit();
			this.SuspendLayout();
			// 
			// btnCancelar
			// 
			this.btnCancelar.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.btnCancelar.Image = ((System.Drawing.Image)(resources.GetObject("btnCancelar.Image")));
			this.btnCancelar.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnCancelar.Location = new System.Drawing.Point(848, 400);
			this.btnCancelar.Name = "btnCancelar";
			this.btnCancelar.Size = new System.Drawing.Size(88, 23);
			this.btnCancelar.TabIndex = 6;
			this.btnCancelar.Text = "&Salir";
			this.btnCancelar.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnCancelar.Click += new System.EventHandler(this.btnCancelar_Click);
			// 
			// btnAceptar
			// 
			this.btnAceptar.Enabled = false;
			this.btnAceptar.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.btnAceptar.Image = ((System.Drawing.Image)(resources.GetObject("btnAceptar.Image")));
			this.btnAceptar.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnAceptar.Location = new System.Drawing.Point(656, 400);
			this.btnAceptar.Name = "btnAceptar";
			this.btnAceptar.Size = new System.Drawing.Size(88, 23);
			this.btnAceptar.TabIndex = 4;
			this.btnAceptar.Text = "&Autorizar";
			this.btnAceptar.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnAceptar.Click += new System.EventHandler(this.btnAceptar_Click);
			// 
			// btRechazar
			// 
			this.btRechazar.Enabled = false;
			this.btRechazar.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.btRechazar.Image = ((System.Drawing.Image)(resources.GetObject("btRechazar.Image")));
			this.btRechazar.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btRechazar.Location = new System.Drawing.Point(752, 400);
			this.btRechazar.Name = "btRechazar";
			this.btRechazar.Size = new System.Drawing.Size(88, 23);
			this.btRechazar.TabIndex = 5;
			this.btRechazar.Text = "&Rechazar";
			this.btRechazar.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btRechazar.Click += new System.EventHandler(this.btRechazar_Click);
			// 
			// lblNombreClte
			// 
			this.lblNombreClte.AutoSize = true;
			this.lblNombreClte.BackColor = System.Drawing.Color.Transparent;
			this.lblNombreClte.Font = new System.Drawing.Font("Microsoft Sans Serif", 20F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.lblNombreClte.Location = new System.Drawing.Point(8, 32);
			this.lblNombreClte.Name = "lblNombreClte";
			this.lblNombreClte.Size = new System.Drawing.Size(208, 34);
			this.lblNombreClte.TabIndex = 189;
			this.lblNombreClte.Text = "Nombre Cliente";
			// 
			// label1
			// 
			this.label1.AutoSize = true;
			this.label1.BackColor = System.Drawing.Color.Transparent;
			this.label1.Location = new System.Drawing.Point(16, 88);
			this.label1.Name = "label1";
			this.label1.Size = new System.Drawing.Size(42, 17);
			this.label1.TabIndex = 190;
			this.label1.Text = "Código:";
			// 
			// lblCodigo
			// 
			this.lblCodigo.AutoSize = true;
			this.lblCodigo.BackColor = System.Drawing.Color.Transparent;
			this.lblCodigo.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.lblCodigo.Location = new System.Drawing.Point(72, 88);
			this.lblCodigo.Name = "lblCodigo";
			this.lblCodigo.Size = new System.Drawing.Size(89, 16);
			this.lblCodigo.TabIndex = 191;
			this.lblCodigo.Text = "1709689805001";
			// 
			// lblRuc
			// 
			this.lblRuc.AutoSize = true;
			this.lblRuc.BackColor = System.Drawing.Color.Transparent;
			this.lblRuc.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.lblRuc.Location = new System.Drawing.Point(248, 88);
			this.lblRuc.Name = "lblRuc";
			this.lblRuc.Size = new System.Drawing.Size(89, 16);
			this.lblRuc.TabIndex = 193;
			this.lblRuc.Text = "1709689805001";
			// 
			// label4
			// 
			this.label4.AutoSize = true;
			this.label4.BackColor = System.Drawing.Color.Transparent;
			this.label4.Location = new System.Drawing.Point(192, 88);
			this.label4.Name = "label4";
			this.label4.Size = new System.Drawing.Size(45, 17);
			this.label4.TabIndex = 192;
			this.label4.Text = "RUC/CI:";
			// 
			// lblFIngreso
			// 
			this.lblFIngreso.AutoSize = true;
			this.lblFIngreso.BackColor = System.Drawing.Color.Transparent;
			this.lblFIngreso.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.lblFIngreso.Location = new System.Drawing.Point(456, 88);
			this.lblFIngreso.Name = "lblFIngreso";
			this.lblFIngreso.Size = new System.Drawing.Size(64, 16);
			this.lblFIngreso.TabIndex = 195;
			this.lblFIngreso.Text = "24-07-2013";
			// 
			// label7
			// 
			this.label7.AutoSize = true;
			this.label7.BackColor = System.Drawing.Color.Transparent;
			this.label7.Location = new System.Drawing.Point(368, 88);
			this.label7.Name = "label7";
			this.label7.Size = new System.Drawing.Size(75, 17);
			this.label7.TabIndex = 194;
			this.label7.Text = "Cliente desde:";
			// 
			// lblEstado
			// 
			this.lblEstado.AutoSize = true;
			this.lblEstado.BackColor = System.Drawing.Color.Transparent;
			this.lblEstado.Font = new System.Drawing.Font("Microsoft Sans Serif", 17F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.lblEstado.Location = new System.Drawing.Point(736, 40);
			this.lblEstado.Name = "lblEstado";
			this.lblEstado.Size = new System.Drawing.Size(146, 29);
			this.lblEstado.TabIndex = 196;
			this.lblEstado.Text = "PENDIENTE";
			this.lblEstado.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
			// 
			// label9
			// 
			this.label9.AutoSize = true;
			this.label9.BackColor = System.Drawing.Color.Transparent;
			this.label9.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label9.Location = new System.Drawing.Point(744, 80);
			this.label9.Name = "label9";
			this.label9.Size = new System.Drawing.Size(134, 16);
			this.label9.TabIndex = 197;
			this.label9.Text = "24-07-2013 11:00:00 AM";
			this.label9.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
			// 
			// groupBox1
			// 
			this.groupBox1.BackColor = System.Drawing.SystemColors.Control;
			this.groupBox1.Controls.Add(this.label7);
			this.groupBox1.Controls.Add(this.label1);
			this.groupBox1.Controls.Add(this.lblRuc);
			this.groupBox1.Controls.Add(this.label4);
			this.groupBox1.Controls.Add(this.lblFIngreso);
			this.groupBox1.Controls.Add(this.lblNombreClte);
			this.groupBox1.Controls.Add(this.lblCodigo);
			this.groupBox1.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.groupBox1.Location = new System.Drawing.Point(8, 8);
			this.groupBox1.Name = "groupBox1";
			this.groupBox1.Size = new System.Drawing.Size(640, 128);
			this.groupBox1.TabIndex = 198;
			this.groupBox1.TabStop = false;
			this.groupBox1.Text = "Datos Cliente";
			// 
			// groupBox2
			// 
			this.groupBox2.BackColor = System.Drawing.SystemColors.Control;
			this.groupBox2.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.groupBox2.Location = new System.Drawing.Point(656, 8);
			this.groupBox2.Name = "groupBox2";
			this.groupBox2.Size = new System.Drawing.Size(288, 96);
			this.groupBox2.TabIndex = 199;
			this.groupBox2.TabStop = false;
			this.groupBox2.Text = "Estado";
			// 
			// groupBox3
			// 
			this.groupBox3.Controls.Add(this.lblValChqDevuelto);
			this.groupBox3.Controls.Add(this.lblBanco);
			this.groupBox3.Controls.Add(this.label18);
			this.groupBox3.Controls.Add(this.lblFechaChDevuelto);
			this.groupBox3.Controls.Add(this.label16);
			this.groupBox3.Controls.Add(this.lblCheqDevueltos);
			this.groupBox3.Controls.Add(this.label14);
			this.groupBox3.Controls.Add(this.lblDisponible);
			this.groupBox3.Controls.Add(this.label12);
			this.groupBox3.Controls.Add(this.lblSaldoActual);
			this.groupBox3.Controls.Add(this.label10);
			this.groupBox3.Controls.Add(this.lblLimeteCredito);
			this.groupBox3.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.groupBox3.Location = new System.Drawing.Point(8, 160);
			this.groupBox3.Name = "groupBox3";
			this.groupBox3.Size = new System.Drawing.Size(640, 112);
			this.groupBox3.TabIndex = 200;
			this.groupBox3.TabStop = false;
			this.groupBox3.Text = "Detalle Financiero";
			// 
			// lblValChqDevuelto
			// 
			this.lblValChqDevuelto.AutoSize = true;
			this.lblValChqDevuelto.BackColor = System.Drawing.Color.Transparent;
			this.lblValChqDevuelto.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.lblValChqDevuelto.Location = new System.Drawing.Point(368, 80);
			this.lblValChqDevuelto.Name = "lblValChqDevuelto";
			this.lblValChqDevuelto.Size = new System.Drawing.Size(50, 16);
			this.lblValChqDevuelto.TabIndex = 206;
			this.lblValChqDevuelto.Text = "1,000.00";
			// 
			// lblBanco
			// 
			this.lblBanco.AutoSize = true;
			this.lblBanco.BackColor = System.Drawing.Color.Transparent;
			this.lblBanco.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.lblBanco.Location = new System.Drawing.Point(368, 64);
			this.lblBanco.Name = "lblBanco";
			this.lblBanco.Size = new System.Drawing.Size(92, 16);
			this.lblBanco.TabIndex = 205;
			this.lblBanco.Text = "Banco Pichincha";
			// 
			// label18
			// 
			this.label18.AutoSize = true;
			this.label18.BackColor = System.Drawing.Color.Transparent;
			this.label18.Location = new System.Drawing.Point(240, 48);
			this.label18.Name = "label18";
			this.label18.Size = new System.Drawing.Size(126, 17);
			this.label18.TabIndex = 204;
			this.label18.Text = "Ultimo cheque devuelto:";
			// 
			// lblFechaChDevuelto
			// 
			this.lblFechaChDevuelto.AutoSize = true;
			this.lblFechaChDevuelto.BackColor = System.Drawing.Color.Transparent;
			this.lblFechaChDevuelto.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.lblFechaChDevuelto.Location = new System.Drawing.Point(368, 48);
			this.lblFechaChDevuelto.Name = "lblFechaChDevuelto";
			this.lblFechaChDevuelto.Size = new System.Drawing.Size(64, 16);
			this.lblFechaChDevuelto.TabIndex = 203;
			this.lblFechaChDevuelto.Text = "25-03-2013";
			// 
			// label16
			// 
			this.label16.AutoSize = true;
			this.label16.BackColor = System.Drawing.Color.Transparent;
			this.label16.Location = new System.Drawing.Point(262, 32);
			this.label16.Name = "label16";
			this.label16.Size = new System.Drawing.Size(102, 17);
			this.label16.TabIndex = 202;
			this.label16.Text = "Cheques devueltos:";
			// 
			// lblCheqDevueltos
			// 
			this.lblCheqDevueltos.AutoSize = true;
			this.lblCheqDevueltos.BackColor = System.Drawing.Color.Transparent;
			this.lblCheqDevueltos.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.lblCheqDevueltos.Location = new System.Drawing.Point(368, 32);
			this.lblCheqDevueltos.Name = "lblCheqDevueltos";
			this.lblCheqDevueltos.Size = new System.Drawing.Size(24, 16);
			this.lblCheqDevueltos.TabIndex = 201;
			this.lblCheqDevueltos.Text = "100";
			// 
			// label14
			// 
			this.label14.AutoSize = true;
			this.label14.BackColor = System.Drawing.Color.Transparent;
			this.label14.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label14.ForeColor = System.Drawing.Color.LimeGreen;
			this.label14.Location = new System.Drawing.Point(57, 64);
			this.label14.Name = "label14";
			this.label14.Size = new System.Drawing.Size(63, 16);
			this.label14.TabIndex = 200;
			this.label14.Text = "Disponible:";
			// 
			// lblDisponible
			// 
			this.lblDisponible.AutoSize = true;
			this.lblDisponible.BackColor = System.Drawing.Color.Transparent;
			this.lblDisponible.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.lblDisponible.ForeColor = System.Drawing.Color.LimeGreen;
			this.lblDisponible.Location = new System.Drawing.Point(128, 64);
			this.lblDisponible.Name = "lblDisponible";
			this.lblDisponible.Size = new System.Drawing.Size(70, 16);
			this.lblDisponible.TabIndex = 199;
			this.lblDisponible.Text = "1000,000.00";
			// 
			// label12
			// 
			this.label12.AutoSize = true;
			this.label12.BackColor = System.Drawing.Color.Transparent;
			this.label12.Location = new System.Drawing.Point(46, 48);
			this.label12.Name = "label12";
			this.label12.Size = new System.Drawing.Size(70, 17);
			this.label12.TabIndex = 198;
			this.label12.Text = "Saldo Actual:";
			// 
			// lblSaldoActual
			// 
			this.lblSaldoActual.AutoSize = true;
			this.lblSaldoActual.BackColor = System.Drawing.Color.Transparent;
			this.lblSaldoActual.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.lblSaldoActual.Location = new System.Drawing.Point(128, 48);
			this.lblSaldoActual.Name = "lblSaldoActual";
			this.lblSaldoActual.Size = new System.Drawing.Size(70, 16);
			this.lblSaldoActual.TabIndex = 197;
			this.lblSaldoActual.Text = "1000,000.00";
			// 
			// label10
			// 
			this.label10.AutoSize = true;
			this.label10.BackColor = System.Drawing.Color.Transparent;
			this.label10.Location = new System.Drawing.Point(24, 32);
			this.label10.Name = "label10";
			this.label10.Size = new System.Drawing.Size(93, 17);
			this.label10.TabIndex = 196;
			this.label10.Text = "Limite de Crédito:";
			// 
			// lblLimeteCredito
			// 
			this.lblLimeteCredito.AutoSize = true;
			this.lblLimeteCredito.BackColor = System.Drawing.Color.Transparent;
			this.lblLimeteCredito.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.lblLimeteCredito.Location = new System.Drawing.Point(128, 32);
			this.lblLimeteCredito.Name = "lblLimeteCredito";
			this.lblLimeteCredito.Size = new System.Drawing.Size(70, 16);
			this.lblLimeteCredito.TabIndex = 196;
			this.lblLimeteCredito.Text = "1000,000.00";
			// 
			// groupBox4
			// 
			this.groupBox4.Controls.Add(this.lblFormaPago);
			this.groupBox4.Controls.Add(this.lblDiasCred3);
			this.groupBox4.Controls.Add(this.label5);
			this.groupBox4.Controls.Add(this.lblDiasCred2);
			this.groupBox4.Controls.Add(this.label2);
			this.groupBox4.Controls.Add(this.lblDiasMora);
			this.groupBox4.Controls.Add(this.lblDiasPago);
			this.groupBox4.Controls.Add(this.label40);
			this.groupBox4.Controls.Add(this.label34);
			this.groupBox4.Controls.Add(this.label39);
			this.groupBox4.Controls.Add(this.lblValUltimoPago);
			this.groupBox4.Controls.Add(this.lblUltimoPago);
			this.groupBox4.Controls.Add(this.lblFchUltimoPago);
			this.groupBox4.Controls.Add(this.label35);
			this.groupBox4.Controls.Add(this.lblValUlmimaFc);
			this.groupBox4.Controls.Add(this.lblFchUltimaFc);
			this.groupBox4.Controls.Add(this.label30);
			this.groupBox4.Controls.Add(this.lblUltimaFc);
			this.groupBox4.Controls.Add(this.label28);
			this.groupBox4.Controls.Add(this.lblDiasCredito);
			this.groupBox4.Controls.Add(this.lblValFcAntigua);
			this.groupBox4.Controls.Add(this.lblNFcAntigua);
			this.groupBox4.Controls.Add(this.label24);
			this.groupBox4.Controls.Add(this.lblFchFcAntigua);
			this.groupBox4.Controls.Add(this.label22);
			this.groupBox4.Controls.Add(this.lblFcxPagar);
			this.groupBox4.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.groupBox4.Location = new System.Drawing.Point(8, 296);
			this.groupBox4.Name = "groupBox4";
			this.groupBox4.Size = new System.Drawing.Size(640, 128);
			this.groupBox4.TabIndex = 201;
			this.groupBox4.TabStop = false;
			// 
			// lblFormaPago
			// 
			this.lblFormaPago.AutoSize = true;
			this.lblFormaPago.BackColor = System.Drawing.Color.Transparent;
			this.lblFormaPago.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.lblFormaPago.Location = new System.Drawing.Point(472, 64);
			this.lblFormaPago.Name = "lblFormaPago";
			this.lblFormaPago.Size = new System.Drawing.Size(81, 16);
			this.lblFormaPago.TabIndex = 233;
			this.lblFormaPago.Text = "30/60/90 DIAS";
			// 
			// lblDiasCred3
			// 
			this.lblDiasCred3.AutoSize = true;
			this.lblDiasCred3.BackColor = System.Drawing.Color.Transparent;
			this.lblDiasCred3.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.lblDiasCred3.Location = new System.Drawing.Point(608, 48);
			this.lblDiasCred3.Name = "lblDiasCred3";
			this.lblDiasCred3.Size = new System.Drawing.Size(17, 16);
			this.lblDiasCred3.TabIndex = 232;
			this.lblDiasCred3.Text = "52";
			// 
			// label5
			// 
			this.label5.AutoSize = true;
			this.label5.BackColor = System.Drawing.Color.Transparent;
			this.label5.Location = new System.Drawing.Point(592, 48);
			this.label5.Name = "label5";
			this.label5.Size = new System.Drawing.Size(8, 17);
			this.label5.TabIndex = 231;
			this.label5.Text = ",";
			// 
			// lblDiasCred2
			// 
			this.lblDiasCred2.AutoSize = true;
			this.lblDiasCred2.BackColor = System.Drawing.Color.Transparent;
			this.lblDiasCred2.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.lblDiasCred2.Location = new System.Drawing.Point(576, 48);
			this.lblDiasCred2.Name = "lblDiasCred2";
			this.lblDiasCred2.Size = new System.Drawing.Size(17, 16);
			this.lblDiasCred2.TabIndex = 230;
			this.lblDiasCred2.Text = "52";
			// 
			// label2
			// 
			this.label2.AutoSize = true;
			this.label2.BackColor = System.Drawing.Color.Transparent;
			this.label2.Location = new System.Drawing.Point(560, 48);
			this.label2.Name = "label2";
			this.label2.Size = new System.Drawing.Size(8, 17);
			this.label2.TabIndex = 229;
			this.label2.Text = ",";
			// 
			// lblDiasMora
			// 
			this.lblDiasMora.AutoSize = true;
			this.lblDiasMora.BackColor = System.Drawing.Color.OrangeRed;
			this.lblDiasMora.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.lblDiasMora.Location = new System.Drawing.Point(544, 96);
			this.lblDiasMora.Name = "lblDiasMora";
			this.lblDiasMora.Size = new System.Drawing.Size(17, 16);
			this.lblDiasMora.TabIndex = 228;
			this.lblDiasMora.Text = "52";
			// 
			// lblDiasPago
			// 
			this.lblDiasPago.AutoSize = true;
			this.lblDiasPago.BackColor = System.Drawing.Color.Transparent;
			this.lblDiasPago.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.lblDiasPago.Location = new System.Drawing.Point(544, 80);
			this.lblDiasPago.Name = "lblDiasPago";
			this.lblDiasPago.Size = new System.Drawing.Size(17, 16);
			this.lblDiasPago.TabIndex = 227;
			this.lblDiasPago.Text = "52";
			// 
			// label40
			// 
			this.label40.AutoSize = true;
			this.label40.BackColor = System.Drawing.Color.Transparent;
			this.label40.Location = new System.Drawing.Point(472, 96);
			this.label40.Name = "label40";
			this.label40.Size = new System.Drawing.Size(64, 17);
			this.label40.TabIndex = 226;
			this.label40.Text = "Días MORA:";
			// 
			// label34
			// 
			this.label34.AutoSize = true;
			this.label34.BackColor = System.Drawing.Color.Transparent;
			this.label34.Location = new System.Drawing.Point(472, 80);
			this.label34.Name = "label34";
			this.label34.Size = new System.Drawing.Size(62, 17);
			this.label34.TabIndex = 225;
			this.label34.Text = "Días PAGO:";
			// 
			// label39
			// 
			this.label39.AutoSize = true;
			this.label39.BackColor = System.Drawing.Color.Transparent;
			this.label39.Location = new System.Drawing.Point(472, 48);
			this.label39.Name = "label39";
			this.label39.Size = new System.Drawing.Size(67, 17);
			this.label39.TabIndex = 224;
			this.label39.Text = "Días crédito:";
			// 
			// lblValUltimoPago
			// 
			this.lblValUltimoPago.AutoSize = true;
			this.lblValUltimoPago.BackColor = System.Drawing.Color.Transparent;
			this.lblValUltimoPago.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.lblValUltimoPago.Location = new System.Drawing.Point(352, 80);
			this.lblValUltimoPago.Name = "lblValUltimoPago";
			this.lblValUltimoPago.Size = new System.Drawing.Size(79, 16);
			this.lblValUltimoPago.TabIndex = 223;
			this.lblValUltimoPago.Text = "1,000,000.00$";
			// 
			// lblUltimoPago
			// 
			this.lblUltimoPago.AutoSize = true;
			this.lblUltimoPago.BackColor = System.Drawing.Color.Transparent;
			this.lblUltimoPago.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.lblUltimoPago.Location = new System.Drawing.Point(352, 48);
			this.lblUltimoPago.Name = "lblUltimoPago";
			this.lblUltimoPago.Size = new System.Drawing.Size(81, 16);
			this.lblUltimoPago.TabIndex = 222;
			this.lblUltimoPago.Text = "BI0000220775";
			// 
			// lblFchUltimoPago
			// 
			this.lblFchUltimoPago.AutoSize = true;
			this.lblFchUltimoPago.BackColor = System.Drawing.Color.Transparent;
			this.lblFchUltimoPago.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.lblFchUltimoPago.Location = new System.Drawing.Point(352, 64);
			this.lblFchUltimoPago.Name = "lblFchUltimoPago";
			this.lblFchUltimoPago.Size = new System.Drawing.Size(64, 16);
			this.lblFchUltimoPago.TabIndex = 221;
			this.lblFchUltimoPago.Text = "25-03-2013";
			// 
			// label35
			// 
			this.label35.AutoSize = true;
			this.label35.BackColor = System.Drawing.Color.Transparent;
			this.label35.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label35.ForeColor = System.Drawing.Color.DarkGreen;
			this.label35.Location = new System.Drawing.Point(352, 32);
			this.label35.Name = "label35";
			this.label35.Size = new System.Drawing.Size(70, 16);
			this.label35.TabIndex = 220;
			this.label35.Text = "Ultimo pago:";
			// 
			// lblValUlmimaFc
			// 
			this.lblValUlmimaFc.AutoSize = true;
			this.lblValUlmimaFc.BackColor = System.Drawing.Color.Transparent;
			this.lblValUlmimaFc.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.lblValUlmimaFc.Location = new System.Drawing.Point(232, 80);
			this.lblValUlmimaFc.Name = "lblValUlmimaFc";
			this.lblValUlmimaFc.Size = new System.Drawing.Size(79, 16);
			this.lblValUlmimaFc.TabIndex = 218;
			this.lblValUlmimaFc.Text = "1,000,000.00$";
			// 
			// lblFchUltimaFc
			// 
			this.lblFchUltimaFc.AutoSize = true;
			this.lblFchUltimaFc.BackColor = System.Drawing.Color.Transparent;
			this.lblFchUltimaFc.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.lblFchUltimaFc.Location = new System.Drawing.Point(232, 64);
			this.lblFchUltimaFc.Name = "lblFchUltimaFc";
			this.lblFchUltimaFc.Size = new System.Drawing.Size(64, 16);
			this.lblFchUltimaFc.TabIndex = 217;
			this.lblFchUltimaFc.Text = "25-03-2013";
			// 
			// label30
			// 
			this.label30.AutoSize = true;
			this.label30.BackColor = System.Drawing.Color.Transparent;
			this.label30.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label30.ForeColor = System.Drawing.Color.Maroon;
			this.label30.Location = new System.Drawing.Point(232, 32);
			this.label30.Name = "label30";
			this.label30.Size = new System.Drawing.Size(81, 16);
			this.label30.TabIndex = 216;
			this.label30.Text = "Ultima factura:";
			// 
			// lblUltimaFc
			// 
			this.lblUltimaFc.AutoSize = true;
			this.lblUltimaFc.BackColor = System.Drawing.Color.Transparent;
			this.lblUltimaFc.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.lblUltimaFc.Location = new System.Drawing.Point(232, 48);
			this.lblUltimaFc.Name = "lblUltimaFc";
			this.lblUltimaFc.Size = new System.Drawing.Size(81, 16);
			this.lblUltimaFc.TabIndex = 215;
			this.lblUltimaFc.Text = "BI0000220775";
			// 
			// label28
			// 
			this.label28.AutoSize = true;
			this.label28.BackColor = System.Drawing.Color.LightSlateGray;
			this.label28.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label28.Location = new System.Drawing.Point(472, 32);
			this.label28.Name = "label28";
			this.label28.Size = new System.Drawing.Size(93, 16);
			this.label28.TabIndex = 214;
			this.label28.Text = "Formas de pago:";
			// 
			// lblDiasCredito
			// 
			this.lblDiasCredito.AutoSize = true;
			this.lblDiasCredito.BackColor = System.Drawing.Color.Transparent;
			this.lblDiasCredito.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.lblDiasCredito.Location = new System.Drawing.Point(544, 48);
			this.lblDiasCredito.Name = "lblDiasCredito";
			this.lblDiasCredito.Size = new System.Drawing.Size(17, 16);
			this.lblDiasCredito.TabIndex = 213;
			this.lblDiasCredito.Text = "52";
			// 
			// lblValFcAntigua
			// 
			this.lblValFcAntigua.AutoSize = true;
			this.lblValFcAntigua.BackColor = System.Drawing.Color.Transparent;
			this.lblValFcAntigua.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.lblValFcAntigua.Location = new System.Drawing.Point(120, 80);
			this.lblValFcAntigua.Name = "lblValFcAntigua";
			this.lblValFcAntigua.Size = new System.Drawing.Size(79, 16);
			this.lblValFcAntigua.TabIndex = 212;
			this.lblValFcAntigua.Text = "1,000,000.00$";
			// 
			// lblNFcAntigua
			// 
			this.lblNFcAntigua.AutoSize = true;
			this.lblNFcAntigua.BackColor = System.Drawing.Color.Transparent;
			this.lblNFcAntigua.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.lblNFcAntigua.Location = new System.Drawing.Point(120, 48);
			this.lblNFcAntigua.Name = "lblNFcAntigua";
			this.lblNFcAntigua.Size = new System.Drawing.Size(74, 16);
			this.lblNFcAntigua.TabIndex = 211;
			this.lblNFcAntigua.Text = "BI000021433";
			// 
			// label24
			// 
			this.label24.AutoSize = true;
			this.label24.BackColor = System.Drawing.Color.Transparent;
			this.label24.Location = new System.Drawing.Point(6, 48);
			this.label24.Name = "label24";
			this.label24.Size = new System.Drawing.Size(110, 17);
			this.label24.TabIndex = 210;
			this.label24.Text = "Factura mas antigua:";
			// 
			// lblFchFcAntigua
			// 
			this.lblFchFcAntigua.AutoSize = true;
			this.lblFchFcAntigua.BackColor = System.Drawing.Color.Transparent;
			this.lblFchFcAntigua.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.lblFchFcAntigua.Location = new System.Drawing.Point(120, 64);
			this.lblFchFcAntigua.Name = "lblFchFcAntigua";
			this.lblFchFcAntigua.Size = new System.Drawing.Size(64, 16);
			this.lblFchFcAntigua.TabIndex = 209;
			this.lblFchFcAntigua.Text = "26-03-2013";
			// 
			// label22
			// 
			this.label22.AutoSize = true;
			this.label22.BackColor = System.Drawing.Color.Transparent;
			this.label22.Location = new System.Drawing.Point(24, 32);
			this.label22.Name = "label22";
			this.label22.Size = new System.Drawing.Size(92, 17);
			this.label22.TabIndex = 208;
			this.label22.Text = "Facturas x pagar:";
			// 
			// lblFcxPagar
			// 
			this.lblFcxPagar.AutoSize = true;
			this.lblFcxPagar.BackColor = System.Drawing.Color.Transparent;
			this.lblFcxPagar.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.lblFcxPagar.Location = new System.Drawing.Point(120, 32);
			this.lblFcxPagar.Name = "lblFcxPagar";
			this.lblFcxPagar.Size = new System.Drawing.Size(24, 16);
			this.lblFcxPagar.TabIndex = 207;
			this.lblFcxPagar.Text = "100";
			// 
			// groupBox5
			// 
			this.groupBox5.Controls.Add(this.txtNotas);
			this.groupBox5.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.groupBox5.Location = new System.Drawing.Point(656, 112);
			this.groupBox5.Name = "groupBox5";
			this.groupBox5.Size = new System.Drawing.Size(288, 104);
			this.groupBox5.TabIndex = 202;
			this.groupBox5.TabStop = false;
			this.groupBox5.Text = "Observaciones";
			// 
			// txtNotas
			// 
			this.txtNotas.AcceptsReturn = true;
			this.txtNotas.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			appearance1.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtNotas.Appearance = appearance1;
			this.txtNotas.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtNotas.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtNotas.Location = new System.Drawing.Point(8, 17);
			this.txtNotas.Multiline = true;
			this.txtNotas.Name = "txtNotas";
			this.txtNotas.Size = new System.Drawing.Size(272, 79);
			this.txtNotas.TabIndex = 0;
			this.txtNotas.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.txtNotas_KeyPress);
			// 
			// groupBox6
			// 
			this.groupBox6.Controls.Add(this.txtDocumentos);
			this.groupBox6.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.groupBox6.Location = new System.Drawing.Point(656, 216);
			this.groupBox6.Name = "groupBox6";
			this.groupBox6.Size = new System.Drawing.Size(288, 120);
			this.groupBox6.TabIndex = 203;
			this.groupBox6.TabStop = false;
			this.groupBox6.Text = "Documentos";
			// 
			// txtDocumentos
			// 
			this.txtDocumentos.AcceptsReturn = true;
			this.txtDocumentos.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			appearance2.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtDocumentos.Appearance = appearance2;
			this.txtDocumentos.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtDocumentos.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtDocumentos.Location = new System.Drawing.Point(8, 16);
			this.txtDocumentos.Multiline = true;
			this.txtDocumentos.Name = "txtDocumentos";
			this.txtDocumentos.Size = new System.Drawing.Size(272, 96);
			this.txtDocumentos.TabIndex = 1;
			this.txtDocumentos.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.txtDocumentos_KeyPress);
			this.txtDocumentos.ValueChanged += new System.EventHandler(this.txtDocumentos_ValueChanged);
			// 
			// errorProvider
			// 
			this.errorProvider.ContainerControl = this;
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
			// label6
			// 
			this.label6.AutoSize = true;
			this.label6.Location = new System.Drawing.Point(768, 344);
			this.label6.Name = "label6";
			this.label6.Size = new System.Drawing.Size(36, 16);
			this.label6.TabIndex = 207;
			this.label6.Text = "Fecha";
			// 
			// label21
			// 
			this.label21.AutoSize = true;
			this.label21.BackColor = System.Drawing.Color.Transparent;
			this.label21.Location = new System.Drawing.Point(656, 344);
			this.label21.Name = "label21";
			this.label21.Size = new System.Drawing.Size(40, 16);
			this.label21.TabIndex = 206;
			this.label21.Text = "Cuotas";
			// 
			// dtFecha
			// 
			appearance3.ForeColorDisabled = System.Drawing.Color.Black;
			this.dtFecha.Appearance = appearance3;
			this.dtFecha.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			dateButton1.Caption = "Today";
			this.dtFecha.DateButtons.Add(dateButton1);
			this.dtFecha.Format = "dd/MM/yyyy";
			this.dtFecha.Location = new System.Drawing.Point(768, 368);
			this.dtFecha.Name = "dtFecha";
			this.dtFecha.NonAutoSizeHeight = 23;
			this.dtFecha.Size = new System.Drawing.Size(112, 21);
			this.dtFecha.SpinButtonsVisible = true;
			this.dtFecha.TabIndex = 3;
			this.dtFecha.Value = ((object)(resources.GetObject("dtFecha.Value")));
			this.dtFecha.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.dtFecha_KeyPress);
			// 
			// txtNumCuotas
			// 
			appearance4.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtNumCuotas.Appearance = appearance4;
			this.txtNumCuotas.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtNumCuotas.FormatString = "";
			this.txtNumCuotas.Location = new System.Drawing.Point(656, 368);
			this.txtNumCuotas.MaskInput = "nn";
			this.txtNumCuotas.MaxValue = 3;
			this.txtNumCuotas.MinValue = 1;
			this.txtNumCuotas.Name = "txtNumCuotas";
			this.txtNumCuotas.PromptChar = ' ';
			this.txtNumCuotas.Size = new System.Drawing.Size(88, 21);
			this.txtNumCuotas.SpinButtonDisplayStyle = Infragistics.Win.ButtonDisplayStyle.Always;
			this.txtNumCuotas.TabIndex = 2;
			this.txtNumCuotas.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.txtNumCuotas_KeyPress);
			// 
			// Point_AutorizaPedido
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 13);
			this.ClientSize = new System.Drawing.Size(946, 432);
			this.Controls.Add(this.label6);
			this.Controls.Add(this.label21);
			this.Controls.Add(this.label9);
			this.Controls.Add(this.lblEstado);
			this.Controls.Add(this.dtFecha);
			this.Controls.Add(this.txtNumCuotas);
			this.Controls.Add(this.groupBox6);
			this.Controls.Add(this.groupBox5);
			this.Controls.Add(this.groupBox4);
			this.Controls.Add(this.groupBox3);
			this.Controls.Add(this.btRechazar);
			this.Controls.Add(this.btnCancelar);
			this.Controls.Add(this.btnAceptar);
			this.Controls.Add(this.groupBox1);
			this.Controls.Add(this.groupBox2);
			this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog;
			this.KeyPreview = true;
			this.MaximizeBox = false;
			this.Name = "Point_AutorizaPedido";
			this.ShowInTaskbar = false;
			this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
			this.Text = "Autorización del pedido";
			this.KeyDown += new System.Windows.Forms.KeyEventHandler(this.Point_AutorizaPedido_KeyDown);
			this.Load += new System.EventHandler(this.Point_AutorizaPedido_Load);
			this.groupBox1.ResumeLayout(false);
			this.groupBox3.ResumeLayout(false);
			this.groupBox4.ResumeLayout(false);
			this.groupBox5.ResumeLayout(false);
			((System.ComponentModel.ISupportInitialize)(this.txtNotas)).EndInit();
			this.groupBox6.ResumeLayout(false);
			((System.ComponentModel.ISupportInitialize)(this.txtDocumentos)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.dtFecha)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNumCuotas)).EndInit();
			this.ResumeLayout(false);

		}
		#endregion

		private void Point_AutorizaPedido_Load(object sender, System.EventArgs e)
		{
			if (iCuotas == 0) this.txtNumCuotas.Value = 1; else this.txtNumCuotas.Value = iCuotas;
			this.dtFecha.CalendarInfo.MinDate = DateTime.Today;
			DateTime dtFechaMaxima = DateTime.Today.AddDays(45);
			dtFechaMaxima = DateTime.Parse(DateTime.DaysInMonth(dtFechaMaxima.Year, dtFechaMaxima.Month).ToString() + "/" + dtFechaMaxima.Month.ToString() + "/" + dtFechaMaxima.Year.ToString());
			this.dtFecha.CalendarInfo.MaxDate = dtFechaMaxima;
			this.dtFecha.Value = dtFechaPrimerPago;
			this.txtNotas.Text = sNotas;

			if (IEstado == 1) 
			{
				this.lblEstado.Text = "PENDIENTE";
				this.btnAceptar.Enabled = true;
				this.btRechazar.Enabled = true;
			}
			else if (IEstado == 4)
			{
				this.lblEstado.Text = "RECHAZADO";
				this.btnAceptar.Enabled = true;
			}
			else if (IEstado == 5 || IEstado == 3) 
			{
				this.lblEstado.Text = "APROBADO";
				this.btRechazar.Enabled = true;
			}

			string stExec = string.Format("Exec DatosAutorizacionDistribuidor {0}", ICliente);
			SqlDataReader dReader = Funcion.rEscalarSQL(cdsSeteoF, stExec, true);
			dReader.Read();
			if (dReader.HasRows)
			{
				lblCodigo.Text = dReader.GetValue(0).ToString();
				lblRuc.Text = dReader.GetValue(1).ToString();
				lblNombreClte.Text = dReader.GetValue(2).ToString();
				lblFIngreso.Text = dReader.GetValue(3).ToString();

				lblLimeteCredito.Text = dReader.GetValue(4).ToString();
				lblSaldoActual.Text = dReader.GetValue(5).ToString();
				lblDisponible.Text = dReader.GetValue(6).ToString();

				lblCheqDevueltos.Text = "0";
				lblFechaChDevuelto.Text = "";
				lblBanco.Text = "";
				lblValChqDevuelto.Text = "";

				lblFcxPagar.Text = dReader.GetValue(7).ToString();
				lblNFcAntigua.Text = dReader.GetValue(12).ToString();
				lblFchFcAntigua.Text = dReader.GetValue(11).ToString();
				lblValFcAntigua.Text = dReader.GetValue(13).ToString();

				lblUltimaFc.Text = dReader.GetValue(9).ToString();
				lblFchUltimaFc.Text = dReader.GetValue(8).ToString();
				lblValUlmimaFc.Text = dReader.GetValue(10).ToString();

				lblUltimoPago.Text = dReader.GetValue(14).ToString();
				lblFchUltimoPago.Text = dReader.GetValue(15).ToString();
				lblValUltimoPago.Text = dReader.GetValue(16).ToString();

				lblDiasCredito.Text = dReader.GetValue(17).ToString();
				lblDiasCred2.Text = dReader.GetValue(18).ToString();
				lblDiasCred3.Text = dReader.GetValue(19).ToString();

//				lblFormaPago.Text = dReader.GetValue(20).ToString();
//				lblDiasPago.Text = dReader.GetValue(21).ToString();
//				lblDiasMora.Text = dReader.GetValue(22).ToString();					
			}
			dReader.Close();

			this.txtDocumentos.Text = sNotas;
			this.txtNotas.Focus();
		}

		private void btnCancelar_Click(object sender, System.EventArgs e)
		{
			this.DialogResult = DialogResult.No;
		}

		private void btnAceptar_Click(object sender, System.EventArgs e)
		{
			if (!Validacion.vbTexto(this.txtNotas, 10, 300, "Observaciones")) return;

			if (!Validacion.vbCampoEnteroVacio(this.txtNumCuotas, "Ingrese el numero de pagos", 1, 3)) return;

			if (!Validacion.vbFechaEnDocumentos(this.dtFecha, "Ingrese la fecha del primer pago", false, "", cdsSeteoF)) return;
			
			iRetEstado = 5;

			this.DialogResult = DialogResult.OK;
		}

		private bool Verifica()
		{
			bool bOk = true;
		   
			LimpiaError();
			
			#region Verifica Campos Obligatorios
			if (txtNotas.Text.ToString().Trim().Length == 0)
			{
				errorProvider.SetError(txtNotas, "Ingrese Observaciones para pedido...!!!");
				bOk = false;
			}
			return bOk;
			#endregion Verifica Campos Obligatorios
		}

		private void LimpiaError()
		{
			errorProvider.SetError(txtNotas, "");  
		}

		private void btRechazar_Click(object sender, System.EventArgs e)
		{
			if (this.txtNotas.Text.ToString().Trim().Length == 0)
			{
				MessageBox.Show("Ingrese las Observaciones para el pedido.", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Information);
				this.txtNotas.Focus();
				return;
			}

			iContSerial = Funcion.iEscalarSQL(cdsSeteoF, string.Format("Select dbo.CuentaSerialesPedido({0})", IPedido));
			
			if (iContSerial > 0 &&  DialogResult.No == MessageBox.Show("Este pedido ya tiene registradas las seriales de sus Productos. \n ¿Esta seguro de rechazar?.", "Point Technology", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2)) return;
			
			iRetEstado = 4;

			this.DialogResult = DialogResult.OK;
		}

		private void Point_AutorizaPedido_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Escape) this.DialogResult = DialogResult.No;
		}

		private void txtDocumentos_ValueChanged(object sender, System.EventArgs e)
		{
		
		}

		private void txtNotas_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
			if (e.KeyChar == 13) this.txtDocumentos.Focus();
		}

		private void txtDocumentos_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
			if (e.KeyChar == 13) this.txtNumCuotas.Focus();
		}

		private void cdsSeteoF_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			cdsSeteoF.Schema.Connections[0].ConnectionString = miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);		
		}

		private void txtNumCuotas_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
			if (e.KeyChar == 13) this.dtFecha.Focus();
		}

		private void dtFecha_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
			if (e.KeyChar == 13)
			{
				if (this.btnAceptar.Enabled) this.btnAceptar.Focus();
				else this.btRechazar.Focus();
			}
		}

	}
}
