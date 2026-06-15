using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using DevExpress.XtraEditors;
using System.Security.Cryptography;

namespace Latinium
{
	/// <summary>
	/// Summary description for PiedraPapel.
	/// </summary>
	public class PiedraPapel : DevExpress.XtraEditors.XtraForm
	{
		private DevExpress.XtraEditors.SimpleButton btPiedra1;
		private DevExpress.XtraEditors.SimpleButton btPapel1;
		private DevExpress.XtraEditors.SimpleButton btTijera1;
		private DevExpress.XtraEditors.SimpleButton btResultado1;
		private DevExpress.XtraEditors.SimpleButton btResultado2;
		private DevExpress.XtraEditors.SimpleButton btTijera2;
		private DevExpress.XtraEditors.SimpleButton btPapel2;
		private DevExpress.XtraEditors.SimpleButton btPiedra2;
		private System.Windows.Forms.Timer timer1;
		private DevExpress.XtraEditors.SimpleButton btIniciar;
		private DevExpress.XtraEditors.LabelControl lblTiempo;
		private DevExpress.XtraEditors.LabelControl lblUsuario;
		private DevExpress.XtraEditors.LabelControl lblComputador;
		private System.Windows.Forms.Timer timer2;
		private Infragistics.Win.UltraWinTabControl.UltraTabControl ultraTabControl1;
		private Infragistics.Win.UltraWinTabControl.UltraTabSharedControlsPage ultraTabSharedControlsPage1;
		private Infragistics.Win.UltraWinTabControl.UltraTabPageControl tabBalance;
		private Infragistics.Win.UltraWinTabControl.UltraTabPageControl ultraTabPageControl1;
		private DevExpress.XtraEditors.SimpleButton btIniciarCM;
		private DevExpress.XtraEditors.LabelControl labelControl4;
		private DevExpress.XtraEditors.LabelControl lblIzquierda;
		private DevExpress.XtraEditors.LabelControl lblDerecha;
		private System.Windows.Forms.Timer timer3;
		private System.Windows.Forms.PictureBox imgCM;
		private System.Windows.Forms.CheckBox chkNumeros;
		private DevExpress.XtraEditors.SimpleButton btDerecha;
		private Infragistics.Win.UltraWinTabControl.UltraTabPageControl ultraTabPageControl2;
		private DevExpress.XtraEditors.LabelControl lblPalabra1;
		private DevExpress.XtraEditors.LabelControl lblPalabra2;
		private DevExpress.XtraEditors.LabelControl lblPalabra4;
		private DevExpress.XtraEditors.LabelControl lblPalabra3;
		private DevExpress.XtraEditors.LabelControl lblPalabra8;
		private DevExpress.XtraEditors.LabelControl lblPalabra7;
		private DevExpress.XtraEditors.LabelControl lblPalabra6;
		private DevExpress.XtraEditors.LabelControl lblPalabra5;
		private DevExpress.XtraEditors.LabelControl lblPalabra12;
		private DevExpress.XtraEditors.LabelControl lblPalabra11;
		private DevExpress.XtraEditors.LabelControl lblPalabra10;
		private DevExpress.XtraEditors.LabelControl lblPalabra9;
		private DevExpress.XtraEditors.LabelControl lblPalabra20;
		private DevExpress.XtraEditors.LabelControl lblPalabra19;
		private DevExpress.XtraEditors.LabelControl lblPalabra18;
		private DevExpress.XtraEditors.LabelControl lblPalabra17;
		private DevExpress.XtraEditors.LabelControl lblPalabra16;
		private DevExpress.XtraEditors.LabelControl lblPalabra15;
		private DevExpress.XtraEditors.LabelControl lblPalabra14;
		private DevExpress.XtraEditors.LabelControl lblPalabra13;
		private DevExpress.XtraEditors.LabelControl labelControl1;
		private DevExpress.XtraEditors.SimpleButton btResultadoColor;
		private DevExpress.XtraEditors.TextEdit txtColores;
		private Infragistics.Win.UltraWinTabControl.UltraTabPageControl ultraTabPageControl3;
		private DevExpress.XtraEditors.LabelControl labelControl2;
		private DevExpress.XtraEditors.SimpleButton btMentir;
		private DevExpress.XtraEditors.SimpleButton btVerdad;
		private DevExpress.XtraEditors.LabelControl lblVerdad;
		private DevExpress.XtraEditors.LabelControl lblInstruccionVerdad;
		private DevExpress.XtraEditors.SimpleButton btPar;
		private DevExpress.XtraEditors.SimpleButton btImpar;
		private DevExpress.XtraEditors.LabelControl lblIncorrecto;
		private DevExpress.XtraEditors.LabelControl lblCorrecto;
		private Infragistics.Win.UltraWinTabControl.UltraTabPageControl ultraTabPageControl4;
		private DevExpress.XtraEditors.LabelControl labelControl3;
		private System.Windows.Forms.PictureBox pictureBox1;
		private DevExpress.XtraEditors.SimpleButton btEmpezarMano;
		private System.Windows.Forms.Timer timer4;
		private System.ComponentModel.IContainer components;

		public PiedraPapel()
		{
			//
			// Required for Windows Form Designer support
			//
			InitializeComponent();

			//
			// TODO: Add any constructor code after InitializeComponent call
			//
		}

		/// <summary>
		/// Clean up any resources being used.
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

		#region Windows Form Designer generated code
		/// <summary>
		/// Required method for Designer support - do not modify
		/// the contents of this method with the code editor.
		/// </summary>
		private void InitializeComponent()
		{
			this.components = new System.ComponentModel.Container();
			System.Resources.ResourceManager resources = new System.Resources.ResourceManager(typeof(PiedraPapel));
			Infragistics.Win.UltraWinTabControl.UltraTab ultraTab1 = new Infragistics.Win.UltraWinTabControl.UltraTab();
			Infragistics.Win.UltraWinTabControl.UltraTab ultraTab2 = new Infragistics.Win.UltraWinTabControl.UltraTab();
			Infragistics.Win.UltraWinTabControl.UltraTab ultraTab3 = new Infragistics.Win.UltraWinTabControl.UltraTab();
			Infragistics.Win.UltraWinTabControl.UltraTab ultraTab4 = new Infragistics.Win.UltraWinTabControl.UltraTab();
			Infragistics.Win.UltraWinTabControl.UltraTab ultraTab5 = new Infragistics.Win.UltraWinTabControl.UltraTab();
			this.tabBalance = new Infragistics.Win.UltraWinTabControl.UltraTabPageControl();
			this.btTijera1 = new DevExpress.XtraEditors.SimpleButton();
			this.btIniciar = new DevExpress.XtraEditors.SimpleButton();
			this.btTijera2 = new DevExpress.XtraEditors.SimpleButton();
			this.lblTiempo = new DevExpress.XtraEditors.LabelControl();
			this.btPiedra1 = new DevExpress.XtraEditors.SimpleButton();
			this.btPiedra2 = new DevExpress.XtraEditors.SimpleButton();
			this.lblUsuario = new DevExpress.XtraEditors.LabelControl();
			this.lblComputador = new DevExpress.XtraEditors.LabelControl();
			this.btPapel1 = new DevExpress.XtraEditors.SimpleButton();
			this.btResultado1 = new DevExpress.XtraEditors.SimpleButton();
			this.btResultado2 = new DevExpress.XtraEditors.SimpleButton();
			this.btPapel2 = new DevExpress.XtraEditors.SimpleButton();
			this.ultraTabPageControl2 = new Infragistics.Win.UltraWinTabControl.UltraTabPageControl();
			this.labelControl1 = new DevExpress.XtraEditors.LabelControl();
			this.btResultadoColor = new DevExpress.XtraEditors.SimpleButton();
			this.txtColores = new DevExpress.XtraEditors.TextEdit();
			this.lblPalabra20 = new DevExpress.XtraEditors.LabelControl();
			this.lblPalabra19 = new DevExpress.XtraEditors.LabelControl();
			this.lblPalabra18 = new DevExpress.XtraEditors.LabelControl();
			this.lblPalabra17 = new DevExpress.XtraEditors.LabelControl();
			this.lblPalabra16 = new DevExpress.XtraEditors.LabelControl();
			this.lblPalabra15 = new DevExpress.XtraEditors.LabelControl();
			this.lblPalabra14 = new DevExpress.XtraEditors.LabelControl();
			this.lblPalabra13 = new DevExpress.XtraEditors.LabelControl();
			this.lblPalabra12 = new DevExpress.XtraEditors.LabelControl();
			this.lblPalabra11 = new DevExpress.XtraEditors.LabelControl();
			this.lblPalabra10 = new DevExpress.XtraEditors.LabelControl();
			this.lblPalabra9 = new DevExpress.XtraEditors.LabelControl();
			this.lblPalabra8 = new DevExpress.XtraEditors.LabelControl();
			this.lblPalabra7 = new DevExpress.XtraEditors.LabelControl();
			this.lblPalabra6 = new DevExpress.XtraEditors.LabelControl();
			this.lblPalabra5 = new DevExpress.XtraEditors.LabelControl();
			this.lblPalabra4 = new DevExpress.XtraEditors.LabelControl();
			this.lblPalabra3 = new DevExpress.XtraEditors.LabelControl();
			this.lblPalabra2 = new DevExpress.XtraEditors.LabelControl();
			this.lblPalabra1 = new DevExpress.XtraEditors.LabelControl();
			this.ultraTabPageControl3 = new Infragistics.Win.UltraWinTabControl.UltraTabPageControl();
			this.lblCorrecto = new DevExpress.XtraEditors.LabelControl();
			this.lblIncorrecto = new DevExpress.XtraEditors.LabelControl();
			this.btImpar = new DevExpress.XtraEditors.SimpleButton();
			this.btPar = new DevExpress.XtraEditors.SimpleButton();
			this.lblInstruccionVerdad = new DevExpress.XtraEditors.LabelControl();
			this.btVerdad = new DevExpress.XtraEditors.SimpleButton();
			this.btMentir = new DevExpress.XtraEditors.SimpleButton();
			this.lblVerdad = new DevExpress.XtraEditors.LabelControl();
			this.labelControl2 = new DevExpress.XtraEditors.LabelControl();
			this.ultraTabPageControl4 = new Infragistics.Win.UltraWinTabControl.UltraTabPageControl();
			this.btEmpezarMano = new DevExpress.XtraEditors.SimpleButton();
			this.pictureBox1 = new System.Windows.Forms.PictureBox();
			this.labelControl3 = new DevExpress.XtraEditors.LabelControl();
			this.ultraTabPageControl1 = new Infragistics.Win.UltraWinTabControl.UltraTabPageControl();
			this.btDerecha = new DevExpress.XtraEditors.SimpleButton();
			this.chkNumeros = new System.Windows.Forms.CheckBox();
			this.imgCM = new System.Windows.Forms.PictureBox();
			this.labelControl4 = new DevExpress.XtraEditors.LabelControl();
			this.btIniciarCM = new DevExpress.XtraEditors.SimpleButton();
			this.lblDerecha = new DevExpress.XtraEditors.LabelControl();
			this.lblIzquierda = new DevExpress.XtraEditors.LabelControl();
			this.timer1 = new System.Windows.Forms.Timer(this.components);
			this.timer2 = new System.Windows.Forms.Timer(this.components);
			this.ultraTabControl1 = new Infragistics.Win.UltraWinTabControl.UltraTabControl();
			this.ultraTabSharedControlsPage1 = new Infragistics.Win.UltraWinTabControl.UltraTabSharedControlsPage();
			this.timer3 = new System.Windows.Forms.Timer(this.components);
			this.timer4 = new System.Windows.Forms.Timer(this.components);
			this.tabBalance.SuspendLayout();
			this.ultraTabPageControl2.SuspendLayout();
			((System.ComponentModel.ISupportInitialize)(this.txtColores.Properties)).BeginInit();
			this.ultraTabPageControl3.SuspendLayout();
			this.ultraTabPageControl4.SuspendLayout();
			this.ultraTabPageControl1.SuspendLayout();
			((System.ComponentModel.ISupportInitialize)(this.ultraTabControl1)).BeginInit();
			this.ultraTabControl1.SuspendLayout();
			this.SuspendLayout();
			// 
			// tabBalance
			// 
			this.tabBalance.Controls.Add(this.btTijera1);
			this.tabBalance.Controls.Add(this.btIniciar);
			this.tabBalance.Controls.Add(this.btTijera2);
			this.tabBalance.Controls.Add(this.lblTiempo);
			this.tabBalance.Controls.Add(this.btPiedra1);
			this.tabBalance.Controls.Add(this.btPiedra2);
			this.tabBalance.Controls.Add(this.lblUsuario);
			this.tabBalance.Controls.Add(this.lblComputador);
			this.tabBalance.Controls.Add(this.btPapel1);
			this.tabBalance.Controls.Add(this.btResultado1);
			this.tabBalance.Controls.Add(this.btResultado2);
			this.tabBalance.Controls.Add(this.btPapel2);
			this.tabBalance.Location = new System.Drawing.Point(-10000, -10000);
			this.tabBalance.Name = "tabBalance";
			this.tabBalance.Size = new System.Drawing.Size(860, 390);
			// 
			// btTijera1
			// 
			this.btTijera1.Enabled = false;
			this.btTijera1.Location = new System.Drawing.Point(104, 256);
			this.btTijera1.Name = "btTijera1";
			this.btTijera1.Size = new System.Drawing.Size(88, 56);
			this.btTijera1.TabIndex = 2;
			this.btTijera1.Text = "&Tijera";
			this.btTijera1.Click += new System.EventHandler(this.btTijera1_Click);
			// 
			// btIniciar
			// 
			this.btIniciar.Location = new System.Drawing.Point(360, 256);
			this.btIniciar.Name = "btIniciar";
			this.btIniciar.Size = new System.Drawing.Size(88, 56);
			this.btIniciar.TabIndex = 11;
			this.btIniciar.Text = "&Iniciar";
			this.btIniciar.Click += new System.EventHandler(this.btIniciar_Click);
			// 
			// btTijera2
			// 
			this.btTijera2.Enabled = false;
			this.btTijera2.Location = new System.Drawing.Point(608, 256);
			this.btTijera2.Name = "btTijera2";
			this.btTijera2.Size = new System.Drawing.Size(88, 56);
			this.btTijera2.TabIndex = 7;
			this.btTijera2.Text = "Tijera";
			// 
			// lblTiempo
			// 
			this.lblTiempo.Appearance.Font = new System.Drawing.Font("Tahoma", 24F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.lblTiempo.Appearance.Options.UseFont = true;
			this.lblTiempo.Appearance.Options.UseTextOptions = true;
			this.lblTiempo.Appearance.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center;
			this.lblTiempo.Location = new System.Drawing.Point(216, 72);
			this.lblTiempo.Name = "lblTiempo";
			this.lblTiempo.Size = new System.Drawing.Size(360, 48);
			this.lblTiempo.TabIndex = 10;
			this.lblTiempo.Text = "Tiempo: 3 s";
			// 
			// btPiedra1
			// 
			this.btPiedra1.Enabled = false;
			this.btPiedra1.Location = new System.Drawing.Point(104, 64);
			this.btPiedra1.Name = "btPiedra1";
			this.btPiedra1.Size = new System.Drawing.Size(88, 56);
			this.btPiedra1.TabIndex = 0;
			this.btPiedra1.Text = "&Piedra";
			this.btPiedra1.Click += new System.EventHandler(this.btPiedra1_Click);
			// 
			// btPiedra2
			// 
			this.btPiedra2.Enabled = false;
			this.btPiedra2.Location = new System.Drawing.Point(608, 64);
			this.btPiedra2.Name = "btPiedra2";
			this.btPiedra2.Size = new System.Drawing.Size(88, 56);
			this.btPiedra2.TabIndex = 5;
			this.btPiedra2.Text = "Piedra";
			// 
			// lblUsuario
			// 
			this.lblUsuario.Appearance.Font = new System.Drawing.Font("Tahoma", 24F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.lblUsuario.Appearance.Options.UseFont = true;
			this.lblUsuario.Location = new System.Drawing.Point(48, 16);
			this.lblUsuario.Name = "lblUsuario";
			this.lblUsuario.Size = new System.Drawing.Size(240, 40);
			this.lblUsuario.TabIndex = 4;
			this.lblUsuario.Text = "Usuario: 0";
			// 
			// lblComputador
			// 
			this.lblComputador.Appearance.Font = new System.Drawing.Font("Tahoma", 24F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.lblComputador.Appearance.Options.UseFont = true;
			this.lblComputador.Location = new System.Drawing.Point(480, 16);
			this.lblComputador.Name = "lblComputador";
			this.lblComputador.Size = new System.Drawing.Size(312, 40);
			this.lblComputador.TabIndex = 9;
			this.lblComputador.Text = "Computador: 0";
			// 
			// btPapel1
			// 
			this.btPapel1.Enabled = false;
			this.btPapel1.Location = new System.Drawing.Point(104, 160);
			this.btPapel1.Name = "btPapel1";
			this.btPapel1.Size = new System.Drawing.Size(88, 56);
			this.btPapel1.TabIndex = 1;
			this.btPapel1.Text = "P&apel";
			this.btPapel1.Click += new System.EventHandler(this.btPapel1_Click);
			// 
			// btResultado1
			// 
			this.btResultado1.Location = new System.Drawing.Point(312, 160);
			this.btResultado1.Name = "btResultado1";
			this.btResultado1.Size = new System.Drawing.Size(88, 56);
			this.btResultado1.TabIndex = 3;
			this.btResultado1.Visible = false;
			// 
			// btResultado2
			// 
			this.btResultado2.Location = new System.Drawing.Point(400, 160);
			this.btResultado2.Name = "btResultado2";
			this.btResultado2.Size = new System.Drawing.Size(88, 56);
			this.btResultado2.TabIndex = 8;
			this.btResultado2.Visible = false;
			// 
			// btPapel2
			// 
			this.btPapel2.Enabled = false;
			this.btPapel2.Location = new System.Drawing.Point(608, 160);
			this.btPapel2.Name = "btPapel2";
			this.btPapel2.Size = new System.Drawing.Size(88, 56);
			this.btPapel2.TabIndex = 6;
			this.btPapel2.Text = "Papel";
			// 
			// ultraTabPageControl2
			// 
			this.ultraTabPageControl2.Controls.Add(this.labelControl1);
			this.ultraTabPageControl2.Controls.Add(this.btResultadoColor);
			this.ultraTabPageControl2.Controls.Add(this.txtColores);
			this.ultraTabPageControl2.Controls.Add(this.lblPalabra20);
			this.ultraTabPageControl2.Controls.Add(this.lblPalabra19);
			this.ultraTabPageControl2.Controls.Add(this.lblPalabra18);
			this.ultraTabPageControl2.Controls.Add(this.lblPalabra17);
			this.ultraTabPageControl2.Controls.Add(this.lblPalabra16);
			this.ultraTabPageControl2.Controls.Add(this.lblPalabra15);
			this.ultraTabPageControl2.Controls.Add(this.lblPalabra14);
			this.ultraTabPageControl2.Controls.Add(this.lblPalabra13);
			this.ultraTabPageControl2.Controls.Add(this.lblPalabra12);
			this.ultraTabPageControl2.Controls.Add(this.lblPalabra11);
			this.ultraTabPageControl2.Controls.Add(this.lblPalabra10);
			this.ultraTabPageControl2.Controls.Add(this.lblPalabra9);
			this.ultraTabPageControl2.Controls.Add(this.lblPalabra8);
			this.ultraTabPageControl2.Controls.Add(this.lblPalabra7);
			this.ultraTabPageControl2.Controls.Add(this.lblPalabra6);
			this.ultraTabPageControl2.Controls.Add(this.lblPalabra5);
			this.ultraTabPageControl2.Controls.Add(this.lblPalabra4);
			this.ultraTabPageControl2.Controls.Add(this.lblPalabra3);
			this.ultraTabPageControl2.Controls.Add(this.lblPalabra2);
			this.ultraTabPageControl2.Controls.Add(this.lblPalabra1);
			this.ultraTabPageControl2.Location = new System.Drawing.Point(-10000, -10000);
			this.ultraTabPageControl2.Name = "ultraTabPageControl2";
			this.ultraTabPageControl2.Size = new System.Drawing.Size(860, 390);
			// 
			// labelControl1
			// 
			this.labelControl1.Appearance.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.labelControl1.Appearance.Options.UseFont = true;
			this.labelControl1.Location = new System.Drawing.Point(40, 24);
			this.labelControl1.Name = "labelControl1";
			this.labelControl1.Size = new System.Drawing.Size(744, 14);
			this.labelControl1.TabIndex = 22;
			this.labelControl1.Text = "Escriba A (Amarillo) R (Rojo) V (Verde) Z (Azul) para cada color que vea";
			// 
			// btResultadoColor
			// 
			this.btResultadoColor.Location = new System.Drawing.Point(744, 328);
			this.btResultadoColor.Name = "btResultadoColor";
			this.btResultadoColor.Size = new System.Drawing.Size(80, 40);
			this.btResultadoColor.TabIndex = 21;
			this.btResultadoColor.Text = "Empezar";
			this.btResultadoColor.Click += new System.EventHandler(this.btResultadoColor_Click);
			// 
			// txtColores
			// 
			this.txtColores.Location = new System.Drawing.Point(48, 336);
			this.txtColores.Name = "txtColores";
			// 
			// txtColores.Properties
			// 
			this.txtColores.Properties.Appearance.Font = new System.Drawing.Font("Tahoma", 14.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.txtColores.Properties.Appearance.Options.UseFont = true;
			this.txtColores.Properties.ReadOnly = true;
			this.txtColores.Size = new System.Drawing.Size(680, 29);
			this.txtColores.TabIndex = 20;
			// 
			// lblPalabra20
			// 
			this.lblPalabra20.Appearance.Font = new System.Drawing.Font("Tahoma", 24F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.lblPalabra20.Appearance.Options.UseFont = true;
			this.lblPalabra20.Location = new System.Drawing.Point(640, 280);
			this.lblPalabra20.Name = "lblPalabra20";
			this.lblPalabra20.Size = new System.Drawing.Size(184, 40);
			this.lblPalabra20.TabIndex = 19;
			this.lblPalabra20.Text = "Palabra1";
			// 
			// lblPalabra19
			// 
			this.lblPalabra19.Appearance.Font = new System.Drawing.Font("Tahoma", 24F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.lblPalabra19.Appearance.Options.UseFont = true;
			this.lblPalabra19.Location = new System.Drawing.Point(440, 280);
			this.lblPalabra19.Name = "lblPalabra19";
			this.lblPalabra19.Size = new System.Drawing.Size(184, 40);
			this.lblPalabra19.TabIndex = 18;
			this.lblPalabra19.Text = "Palabra1";
			// 
			// lblPalabra18
			// 
			this.lblPalabra18.Appearance.Font = new System.Drawing.Font("Tahoma", 24F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.lblPalabra18.Appearance.Options.UseFont = true;
			this.lblPalabra18.Location = new System.Drawing.Point(240, 280);
			this.lblPalabra18.Name = "lblPalabra18";
			this.lblPalabra18.Size = new System.Drawing.Size(184, 40);
			this.lblPalabra18.TabIndex = 17;
			this.lblPalabra18.Text = "Palabra1";
			// 
			// lblPalabra17
			// 
			this.lblPalabra17.Appearance.Font = new System.Drawing.Font("Tahoma", 24F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.lblPalabra17.Appearance.Options.UseFont = true;
			this.lblPalabra17.Location = new System.Drawing.Point(40, 280);
			this.lblPalabra17.Name = "lblPalabra17";
			this.lblPalabra17.Size = new System.Drawing.Size(184, 40);
			this.lblPalabra17.TabIndex = 16;
			this.lblPalabra17.Text = "Palabra1";
			// 
			// lblPalabra16
			// 
			this.lblPalabra16.Appearance.Font = new System.Drawing.Font("Tahoma", 24F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.lblPalabra16.Appearance.Options.UseFont = true;
			this.lblPalabra16.Location = new System.Drawing.Point(640, 224);
			this.lblPalabra16.Name = "lblPalabra16";
			this.lblPalabra16.Size = new System.Drawing.Size(184, 40);
			this.lblPalabra16.TabIndex = 15;
			this.lblPalabra16.Text = "Palabra1";
			// 
			// lblPalabra15
			// 
			this.lblPalabra15.Appearance.Font = new System.Drawing.Font("Tahoma", 24F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.lblPalabra15.Appearance.Options.UseFont = true;
			this.lblPalabra15.Location = new System.Drawing.Point(440, 224);
			this.lblPalabra15.Name = "lblPalabra15";
			this.lblPalabra15.Size = new System.Drawing.Size(184, 40);
			this.lblPalabra15.TabIndex = 14;
			this.lblPalabra15.Text = "Palabra1";
			// 
			// lblPalabra14
			// 
			this.lblPalabra14.Appearance.Font = new System.Drawing.Font("Tahoma", 24F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.lblPalabra14.Appearance.Options.UseFont = true;
			this.lblPalabra14.Location = new System.Drawing.Point(240, 224);
			this.lblPalabra14.Name = "lblPalabra14";
			this.lblPalabra14.Size = new System.Drawing.Size(184, 40);
			this.lblPalabra14.TabIndex = 13;
			this.lblPalabra14.Text = "Palabra1";
			// 
			// lblPalabra13
			// 
			this.lblPalabra13.Appearance.Font = new System.Drawing.Font("Tahoma", 24F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.lblPalabra13.Appearance.Options.UseFont = true;
			this.lblPalabra13.Location = new System.Drawing.Point(40, 224);
			this.lblPalabra13.Name = "lblPalabra13";
			this.lblPalabra13.Size = new System.Drawing.Size(184, 40);
			this.lblPalabra13.TabIndex = 12;
			this.lblPalabra13.Text = "Palabra1";
			// 
			// lblPalabra12
			// 
			this.lblPalabra12.Appearance.Font = new System.Drawing.Font("Tahoma", 24F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.lblPalabra12.Appearance.Options.UseFont = true;
			this.lblPalabra12.Location = new System.Drawing.Point(640, 168);
			this.lblPalabra12.Name = "lblPalabra12";
			this.lblPalabra12.Size = new System.Drawing.Size(184, 40);
			this.lblPalabra12.TabIndex = 11;
			this.lblPalabra12.Text = "Palabra1";
			// 
			// lblPalabra11
			// 
			this.lblPalabra11.Appearance.Font = new System.Drawing.Font("Tahoma", 24F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.lblPalabra11.Appearance.Options.UseFont = true;
			this.lblPalabra11.Location = new System.Drawing.Point(440, 168);
			this.lblPalabra11.Name = "lblPalabra11";
			this.lblPalabra11.Size = new System.Drawing.Size(184, 40);
			this.lblPalabra11.TabIndex = 10;
			this.lblPalabra11.Text = "Palabra1";
			// 
			// lblPalabra10
			// 
			this.lblPalabra10.Appearance.Font = new System.Drawing.Font("Tahoma", 24F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.lblPalabra10.Appearance.Options.UseFont = true;
			this.lblPalabra10.Location = new System.Drawing.Point(240, 168);
			this.lblPalabra10.Name = "lblPalabra10";
			this.lblPalabra10.Size = new System.Drawing.Size(184, 40);
			this.lblPalabra10.TabIndex = 9;
			this.lblPalabra10.Text = "Palabra1";
			// 
			// lblPalabra9
			// 
			this.lblPalabra9.Appearance.Font = new System.Drawing.Font("Tahoma", 24F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.lblPalabra9.Appearance.Options.UseFont = true;
			this.lblPalabra9.Location = new System.Drawing.Point(40, 168);
			this.lblPalabra9.Name = "lblPalabra9";
			this.lblPalabra9.Size = new System.Drawing.Size(184, 40);
			this.lblPalabra9.TabIndex = 8;
			this.lblPalabra9.Text = "Palabra1";
			// 
			// lblPalabra8
			// 
			this.lblPalabra8.Appearance.Font = new System.Drawing.Font("Tahoma", 24F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.lblPalabra8.Appearance.Options.UseFont = true;
			this.lblPalabra8.Location = new System.Drawing.Point(640, 112);
			this.lblPalabra8.Name = "lblPalabra8";
			this.lblPalabra8.Size = new System.Drawing.Size(184, 40);
			this.lblPalabra8.TabIndex = 7;
			this.lblPalabra8.Text = "Palabra1";
			// 
			// lblPalabra7
			// 
			this.lblPalabra7.Appearance.Font = new System.Drawing.Font("Tahoma", 24F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.lblPalabra7.Appearance.Options.UseFont = true;
			this.lblPalabra7.Location = new System.Drawing.Point(440, 112);
			this.lblPalabra7.Name = "lblPalabra7";
			this.lblPalabra7.Size = new System.Drawing.Size(184, 40);
			this.lblPalabra7.TabIndex = 6;
			this.lblPalabra7.Text = "Palabra1";
			// 
			// lblPalabra6
			// 
			this.lblPalabra6.Appearance.Font = new System.Drawing.Font("Tahoma", 24F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.lblPalabra6.Appearance.Options.UseFont = true;
			this.lblPalabra6.Location = new System.Drawing.Point(240, 112);
			this.lblPalabra6.Name = "lblPalabra6";
			this.lblPalabra6.Size = new System.Drawing.Size(184, 40);
			this.lblPalabra6.TabIndex = 5;
			this.lblPalabra6.Text = "Palabra1";
			// 
			// lblPalabra5
			// 
			this.lblPalabra5.Appearance.Font = new System.Drawing.Font("Tahoma", 24F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.lblPalabra5.Appearance.Options.UseFont = true;
			this.lblPalabra5.Location = new System.Drawing.Point(40, 112);
			this.lblPalabra5.Name = "lblPalabra5";
			this.lblPalabra5.Size = new System.Drawing.Size(184, 40);
			this.lblPalabra5.TabIndex = 4;
			this.lblPalabra5.Text = "Palabra1";
			// 
			// lblPalabra4
			// 
			this.lblPalabra4.Appearance.Font = new System.Drawing.Font("Tahoma", 24F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.lblPalabra4.Appearance.Options.UseFont = true;
			this.lblPalabra4.Location = new System.Drawing.Point(640, 56);
			this.lblPalabra4.Name = "lblPalabra4";
			this.lblPalabra4.Size = new System.Drawing.Size(184, 40);
			this.lblPalabra4.TabIndex = 3;
			this.lblPalabra4.Text = "Palabra1";
			// 
			// lblPalabra3
			// 
			this.lblPalabra3.Appearance.Font = new System.Drawing.Font("Tahoma", 24F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.lblPalabra3.Appearance.Options.UseFont = true;
			this.lblPalabra3.Location = new System.Drawing.Point(440, 56);
			this.lblPalabra3.Name = "lblPalabra3";
			this.lblPalabra3.Size = new System.Drawing.Size(184, 40);
			this.lblPalabra3.TabIndex = 2;
			this.lblPalabra3.Text = "Palabra1";
			// 
			// lblPalabra2
			// 
			this.lblPalabra2.Appearance.Font = new System.Drawing.Font("Tahoma", 24F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.lblPalabra2.Appearance.Options.UseFont = true;
			this.lblPalabra2.Location = new System.Drawing.Point(240, 56);
			this.lblPalabra2.Name = "lblPalabra2";
			this.lblPalabra2.Size = new System.Drawing.Size(184, 40);
			this.lblPalabra2.TabIndex = 1;
			this.lblPalabra2.Text = "Palabra1";
			// 
			// lblPalabra1
			// 
			this.lblPalabra1.Appearance.Font = new System.Drawing.Font("Tahoma", 24F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.lblPalabra1.Appearance.ForeColor = System.Drawing.Color.Red;
			this.lblPalabra1.Appearance.Options.UseFont = true;
			this.lblPalabra1.Appearance.Options.UseForeColor = true;
			this.lblPalabra1.Location = new System.Drawing.Point(40, 56);
			this.lblPalabra1.Name = "lblPalabra1";
			this.lblPalabra1.Size = new System.Drawing.Size(184, 40);
			this.lblPalabra1.TabIndex = 0;
			this.lblPalabra1.Text = "Palabra1";
			// 
			// ultraTabPageControl3
			// 
			this.ultraTabPageControl3.Controls.Add(this.lblCorrecto);
			this.ultraTabPageControl3.Controls.Add(this.lblIncorrecto);
			this.ultraTabPageControl3.Controls.Add(this.btImpar);
			this.ultraTabPageControl3.Controls.Add(this.btPar);
			this.ultraTabPageControl3.Controls.Add(this.lblInstruccionVerdad);
			this.ultraTabPageControl3.Controls.Add(this.btVerdad);
			this.ultraTabPageControl3.Controls.Add(this.btMentir);
			this.ultraTabPageControl3.Controls.Add(this.lblVerdad);
			this.ultraTabPageControl3.Controls.Add(this.labelControl2);
			this.ultraTabPageControl3.Location = new System.Drawing.Point(1, 23);
			this.ultraTabPageControl3.Name = "ultraTabPageControl3";
			this.ultraTabPageControl3.Size = new System.Drawing.Size(860, 390);
			// 
			// lblCorrecto
			// 
			this.lblCorrecto.Appearance.Font = new System.Drawing.Font("Tahoma", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.lblCorrecto.Appearance.Options.UseFont = true;
			this.lblCorrecto.Appearance.Options.UseTextOptions = true;
			this.lblCorrecto.Appearance.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center;
			this.lblCorrecto.Location = new System.Drawing.Point(424, 344);
			this.lblCorrecto.Name = "lblCorrecto";
			this.lblCorrecto.Size = new System.Drawing.Size(232, 32);
			this.lblCorrecto.TabIndex = 8;
			this.lblCorrecto.Text = "Correctos: 0";
			// 
			// lblIncorrecto
			// 
			this.lblIncorrecto.Appearance.Font = new System.Drawing.Font("Tahoma", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.lblIncorrecto.Appearance.Options.UseFont = true;
			this.lblIncorrecto.Appearance.Options.UseTextOptions = true;
			this.lblIncorrecto.Appearance.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center;
			this.lblIncorrecto.Location = new System.Drawing.Point(96, 344);
			this.lblIncorrecto.Name = "lblIncorrecto";
			this.lblIncorrecto.Size = new System.Drawing.Size(232, 32);
			this.lblIncorrecto.TabIndex = 7;
			this.lblIncorrecto.Text = "Incorrectos: 0";
			// 
			// btImpar
			// 
			this.btImpar.Location = new System.Drawing.Point(488, 280);
			this.btImpar.Name = "btImpar";
			this.btImpar.Size = new System.Drawing.Size(128, 48);
			this.btImpar.TabIndex = 6;
			this.btImpar.Text = "El numero es Impar";
			this.btImpar.Visible = false;
			this.btImpar.Click += new System.EventHandler(this.btImpar_Click);
			// 
			// btPar
			// 
			this.btPar.Location = new System.Drawing.Point(168, 280);
			this.btPar.Name = "btPar";
			this.btPar.Size = new System.Drawing.Size(120, 48);
			this.btPar.TabIndex = 5;
			this.btPar.Text = "El numero es par";
			this.btPar.Visible = false;
			this.btPar.Click += new System.EventHandler(this.btPar_Click);
			// 
			// lblInstruccionVerdad
			// 
			this.lblInstruccionVerdad.Appearance.Font = new System.Drawing.Font("Tahoma", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.lblInstruccionVerdad.Appearance.Options.UseFont = true;
			this.lblInstruccionVerdad.Appearance.Options.UseTextOptions = true;
			this.lblInstruccionVerdad.Appearance.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center;
			this.lblInstruccionVerdad.Location = new System.Drawing.Point(128, 216);
			this.lblInstruccionVerdad.Name = "lblInstruccionVerdad";
			this.lblInstruccionVerdad.Size = new System.Drawing.Size(592, 32);
			this.lblInstruccionVerdad.TabIndex = 4;
			this.lblInstruccionVerdad.Text = "Diga la verdad";
			// 
			// btVerdad
			// 
			this.btVerdad.Location = new System.Drawing.Point(608, 72);
			this.btVerdad.Name = "btVerdad";
			this.btVerdad.Size = new System.Drawing.Size(96, 48);
			this.btVerdad.TabIndex = 3;
			this.btVerdad.Text = "La Verdad";
			this.btVerdad.Click += new System.EventHandler(this.btVerdad_Click);
			// 
			// btMentir
			// 
			this.btMentir.Location = new System.Drawing.Point(88, 72);
			this.btMentir.Name = "btMentir";
			this.btMentir.Size = new System.Drawing.Size(96, 48);
			this.btMentir.TabIndex = 2;
			this.btMentir.Text = "Mentir";
			this.btMentir.Click += new System.EventHandler(this.btMentir_Click);
			// 
			// lblVerdad
			// 
			this.lblVerdad.Appearance.Font = new System.Drawing.Font("Tahoma", 72F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.lblVerdad.Appearance.Options.UseFont = true;
			this.lblVerdad.Appearance.Options.UseTextOptions = true;
			this.lblVerdad.Appearance.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center;
			this.lblVerdad.Location = new System.Drawing.Point(328, 80);
			this.lblVerdad.Name = "lblVerdad";
			this.lblVerdad.Size = new System.Drawing.Size(144, 128);
			this.lblVerdad.TabIndex = 1;
			this.lblVerdad.Text = "1";
			// 
			// labelControl2
			// 
			this.labelControl2.Appearance.Font = new System.Drawing.Font("Tahoma", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.labelControl2.Appearance.Options.UseFont = true;
			this.labelControl2.Location = new System.Drawing.Point(24, 16);
			this.labelControl2.Name = "labelControl2";
			this.labelControl2.Size = new System.Drawing.Size(592, 32);
			this.labelControl2.TabIndex = 0;
			this.labelControl2.Text = "Siempre es más fácil y mejor decir la verdad. Pruebelo ...";
			// 
			// ultraTabPageControl4
			// 
			this.ultraTabPageControl4.Controls.Add(this.btEmpezarMano);
			this.ultraTabPageControl4.Controls.Add(this.pictureBox1);
			this.ultraTabPageControl4.Controls.Add(this.labelControl3);
			this.ultraTabPageControl4.Location = new System.Drawing.Point(-10000, -10000);
			this.ultraTabPageControl4.Name = "ultraTabPageControl4";
			this.ultraTabPageControl4.Size = new System.Drawing.Size(860, 390);
			// 
			// btEmpezarMano
			// 
			this.btEmpezarMano.Location = new System.Drawing.Point(720, 16);
			this.btEmpezarMano.Name = "btEmpezarMano";
			this.btEmpezarMano.Size = new System.Drawing.Size(104, 32);
			this.btEmpezarMano.TabIndex = 3;
			this.btEmpezarMano.Text = "Empezar";
			this.btEmpezarMano.Click += new System.EventHandler(this.btEmpezarMano_Click);
			// 
			// pictureBox1
			// 
			this.pictureBox1.Image = ((System.Drawing.Image)(resources.GetObject("pictureBox1.Image")));
			this.pictureBox1.Location = new System.Drawing.Point(152, 64);
			this.pictureBox1.Name = "pictureBox1";
			this.pictureBox1.Size = new System.Drawing.Size(528, 304);
			this.pictureBox1.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage;
			this.pictureBox1.TabIndex = 2;
			this.pictureBox1.TabStop = false;
			// 
			// labelControl3
			// 
			this.labelControl3.Appearance.Font = new System.Drawing.Font("Tahoma", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.labelControl3.Appearance.Options.UseFont = true;
			this.labelControl3.Location = new System.Drawing.Point(24, 16);
			this.labelControl3.Name = "labelControl3";
			this.labelControl3.Size = new System.Drawing.Size(592, 32);
			this.labelControl3.TabIndex = 1;
			this.labelControl3.Text = "Conecte el lado Izquierdo y Derecho de su cerebro";
			// 
			// ultraTabPageControl1
			// 
			this.ultraTabPageControl1.Controls.Add(this.btDerecha);
			this.ultraTabPageControl1.Controls.Add(this.chkNumeros);
			this.ultraTabPageControl1.Controls.Add(this.imgCM);
			this.ultraTabPageControl1.Controls.Add(this.labelControl4);
			this.ultraTabPageControl1.Controls.Add(this.btIniciarCM);
			this.ultraTabPageControl1.Controls.Add(this.lblDerecha);
			this.ultraTabPageControl1.Controls.Add(this.lblIzquierda);
			this.ultraTabPageControl1.Location = new System.Drawing.Point(-10000, -10000);
			this.ultraTabPageControl1.Name = "ultraTabPageControl1";
			this.ultraTabPageControl1.Size = new System.Drawing.Size(860, 390);
			// 
			// btDerecha
			// 
			this.btDerecha.Location = new System.Drawing.Point(608, 296);
			this.btDerecha.Name = "btDerecha";
			this.btDerecha.Size = new System.Drawing.Size(152, 48);
			this.btDerecha.TabIndex = 7;
			this.btDerecha.Text = "Derecha";
			this.btDerecha.Click += new System.EventHandler(this.btDerecha_Click);
			// 
			// chkNumeros
			// 
			this.chkNumeros.BackColor = System.Drawing.Color.Transparent;
			this.chkNumeros.Location = new System.Drawing.Point(712, 16);
			this.chkNumeros.Name = "chkNumeros";
			this.chkNumeros.Size = new System.Drawing.Size(144, 24);
			this.chkNumeros.TabIndex = 6;
			this.chkNumeros.Text = "Numeros";
			// 
			// imgCM
			// 
			this.imgCM.Image = ((System.Drawing.Image)(resources.GetObject("imgCM.Image")));
			this.imgCM.Location = new System.Drawing.Point(304, 72);
			this.imgCM.Name = "imgCM";
			this.imgCM.Size = new System.Drawing.Size(264, 192);
			this.imgCM.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage;
			this.imgCM.TabIndex = 5;
			this.imgCM.TabStop = false;
			this.imgCM.Visible = false;
			// 
			// labelControl4
			// 
			this.labelControl4.Appearance.Font = new System.Drawing.Font("Tahoma", 14.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.labelControl4.Appearance.Options.UseFont = true;
			this.labelControl4.Location = new System.Drawing.Point(16, 16);
			this.labelControl4.Name = "labelControl4";
			this.labelControl4.Size = new System.Drawing.Size(664, 24);
			this.labelControl4.TabIndex = 4;
			this.labelControl4.Text = "Escoge izquierda o derecha y piensa intensamente en ese lado";
			// 
			// btIniciarCM
			// 
			this.btIniciarCM.Location = new System.Drawing.Point(88, 296);
			this.btIniciarCM.Name = "btIniciarCM";
			this.btIniciarCM.Size = new System.Drawing.Size(152, 48);
			this.btIniciarCM.TabIndex = 2;
			this.btIniciarCM.Text = "Izquierda";
			this.btIniciarCM.Click += new System.EventHandler(this.btIniciarCM_Click);
			// 
			// lblDerecha
			// 
			this.lblDerecha.Appearance.Font = new System.Drawing.Font("Tahoma", 72F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.lblDerecha.Appearance.Options.UseFont = true;
			this.lblDerecha.Appearance.Options.UseTextOptions = true;
			this.lblDerecha.Appearance.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center;
			this.lblDerecha.Appearance.TextOptions.VAlignment = DevExpress.Utils.VertAlignment.Center;
			this.lblDerecha.BorderStyle = DevExpress.XtraEditors.Controls.BorderStyles.Office2003;
			this.lblDerecha.Location = new System.Drawing.Point(528, 72);
			this.lblDerecha.Name = "lblDerecha";
			this.lblDerecha.Size = new System.Drawing.Size(296, 160);
			this.lblDerecha.TabIndex = 1;
			this.lblDerecha.Text = "50";
			this.lblDerecha.Visible = false;
			// 
			// lblIzquierda
			// 
			this.lblIzquierda.Appearance.Font = new System.Drawing.Font("Tahoma", 72F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.lblIzquierda.Appearance.Options.UseFont = true;
			this.lblIzquierda.Appearance.Options.UseTextOptions = true;
			this.lblIzquierda.Appearance.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Center;
			this.lblIzquierda.Appearance.TextOptions.VAlignment = DevExpress.Utils.VertAlignment.Center;
			this.lblIzquierda.BorderStyle = DevExpress.XtraEditors.Controls.BorderStyles.Office2003;
			this.lblIzquierda.Location = new System.Drawing.Point(40, 72);
			this.lblIzquierda.Name = "lblIzquierda";
			this.lblIzquierda.Size = new System.Drawing.Size(288, 160);
			this.lblIzquierda.TabIndex = 0;
			this.lblIzquierda.Text = "50";
			this.lblIzquierda.Visible = false;
			// 
			// timer1
			// 
			this.timer1.Interval = 1000;
			this.timer1.Tick += new System.EventHandler(this.timer1_Tick);
			// 
			// timer2
			// 
			this.timer2.Tick += new System.EventHandler(this.timer2_Tick);
			// 
			// ultraTabControl1
			// 
			this.ultraTabControl1.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.ultraTabControl1.Controls.Add(this.ultraTabSharedControlsPage1);
			this.ultraTabControl1.Controls.Add(this.tabBalance);
			this.ultraTabControl1.Controls.Add(this.ultraTabPageControl1);
			this.ultraTabControl1.Controls.Add(this.ultraTabPageControl2);
			this.ultraTabControl1.Controls.Add(this.ultraTabPageControl3);
			this.ultraTabControl1.Controls.Add(this.ultraTabPageControl4);
			this.ultraTabControl1.Location = new System.Drawing.Point(8, 8);
			this.ultraTabControl1.Name = "ultraTabControl1";
			this.ultraTabControl1.SharedControlsPage = this.ultraTabSharedControlsPage1;
			this.ultraTabControl1.Size = new System.Drawing.Size(864, 416);
			this.ultraTabControl1.TabIndex = 13;
			ultraTab1.TabPage = this.tabBalance;
			ultraTab1.Text = "Piedra Papel o Tijera";
			ultraTab2.TabPage = this.ultraTabPageControl2;
			ultraTab2.Text = "Colores";
			ultraTab3.TabPage = this.ultraTabPageControl3;
			ultraTab3.Text = "La Verdad";
			ultraTab4.TabPage = this.ultraTabPageControl4;
			ultraTab4.Text = "Conexiones";
			ultraTab5.TabPage = this.ultraTabPageControl1;
			ultraTab5.Text = "Control Mental";
			this.ultraTabControl1.Tabs.AddRange(new Infragistics.Win.UltraWinTabControl.UltraTab[] {
																																															 ultraTab1,
																																															 ultraTab2,
																																															 ultraTab3,
																																															 ultraTab4,
																																															 ultraTab5});
			// 
			// ultraTabSharedControlsPage1
			// 
			this.ultraTabSharedControlsPage1.Location = new System.Drawing.Point(-10000, -10000);
			this.ultraTabSharedControlsPage1.Name = "ultraTabSharedControlsPage1";
			this.ultraTabSharedControlsPage1.Size = new System.Drawing.Size(860, 390);
			// 
			// timer3
			// 
			this.timer3.Interval = 500;
			this.timer3.Tick += new System.EventHandler(this.timer3_Tick);
			// 
			// timer4
			// 
			this.timer4.Interval = 2000;
			this.timer4.Tick += new System.EventHandler(this.timer4_Tick);
			// 
			// PiedraPapel
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 14);
			this.ClientSize = new System.Drawing.Size(880, 430);
			this.Controls.Add(this.ultraTabControl1);
			this.Name = "PiedraPapel";
			this.Text = "Juegos";
			this.Load += new System.EventHandler(this.PiedraPapel_Load);
			this.tabBalance.ResumeLayout(false);
			this.ultraTabPageControl2.ResumeLayout(false);
			((System.ComponentModel.ISupportInitialize)(this.txtColores.Properties)).EndInit();
			this.ultraTabPageControl3.ResumeLayout(false);
			this.ultraTabPageControl4.ResumeLayout(false);
			this.ultraTabPageControl1.ResumeLayout(false);
			((System.ComponentModel.ISupportInitialize)(this.ultraTabControl1)).EndInit();
			this.ultraTabControl1.ResumeLayout(false);
			this.ResumeLayout(false);

		}
		#endregion

		private void btIniciar_Click(object sender, System.EventArgs e)
		{
			lblTiempo.Text = "Tiempo: 3 s";
			timer1.Enabled = true;
			btPapel1.Enabled = true;
			btPiedra1.Enabled = true;
			btTijera1.Enabled = true;
			btResultado1.Location = Res1;
			btResultado2.Location = Res2;
		}

		string stUltimo = "";
		string[] stSerie = {"", "", "", "", "", "", "", "", ""};
		int iRepetido = 0;
		bool bTrampa = false;
		private string Jugar()
		{
			string [] stObjeto = {"Piedra", "Papel", "Tijera"};
			Random r = new Random(DateTime.Now.Millisecond);
			int iActual = r.Next(0, 3);
			if (iRepetido >= 2 && !bTrampa) // Siempre envia lo mismo
			{
				if (stUltimo == btResultado1.Text) 
					lblTiempo.Text = "Predecible";
				else
				{
					lblTiempo.Text = "Trampa";
					bTrampa = true;
				}

				if (stUltimo == "Piedra") return "Papel";
				if (stUltimo == "Papel") return "Tijera";
				if (stUltimo == "Tijera") return "Piedra";
			}
			if (stUltimo == btResultado1.Text) 
				iRepetido ++;
			else 
			{
				iRepetido = 0;
				stUltimo = btResultado1.Text;
			}

			for (int j=8; j > 0; j--)
				stSerie[j] = stSerie[j-1];
			stSerie[0] = btResultado1.Text;
			if (stSerie[0] == stSerie[3] && stSerie[0] == stSerie[6]) // Serie identificada
			{
				lblTiempo.Text = "Predecible Serie";
				if (btResultado1.Text == "Piedra") return "Papel";
				if (btResultado1.Text == "Papel") return "Tijera";
				if (btResultado1.Text == "Tijera") return "Piedra";
			}

			return stObjeto[iActual];
		}

		private void Resultado()
		{
			string Objeto1 = btResultado1.Text;
			string Objeto2 = btResultado2.Text;
			int iUsuario1 = iUsuario;
			int iComputador1 = iComputador;
			if (Objeto1 == "") iComputador++;
			if (Objeto1 == "Piedra" && Objeto2 == "Tijera") iUsuario++;
			if (Objeto1 == "Papel" && Objeto2 == "Piedra") iUsuario++;
			if (Objeto1 == "Tijera" && Objeto2 == "Papel") iUsuario++;

			if (Objeto2 == "Piedra" && Objeto1 == "Tijera") iComputador++;
			if (Objeto2 == "Papel" && Objeto1 == "Piedra") iComputador++;
			if (Objeto2 == "Tijera" && Objeto1 == "Papel") iComputador++;

			lblComputador.Text = "Computador: " + iComputador;
			lblUsuario.Text = "Usuario: " + iUsuario;
			if (iComputador > iComputador1 && Objeto1 != "") lblTiempo.Text = "Gana el PC";
			if (Objeto1 == "") lblTiempo.Text = "Gana PC por Lento";
			if (iUsuario > iUsuario1) lblTiempo.Text = "Tu Ganas";
			if (iUsuario1 == iUsuario && iComputador1 == iComputador) lblTiempo.Text = "Empate";
		}

		private int iTiempo = 3;
		int iComputador = 0;
		int iUsuario = 0;
		private void timer1_Tick(object sender, System.EventArgs e)
		{
			iTiempo--;
			btIniciar.Enabled = false;
			lblTiempo.Text = "Tiempo: " + iTiempo + " s";
			if (iTiempo == 0) 
			{
				iTiempo = 3;
				lblTiempo.Text = "Tiempo";
				timer1.Enabled = false;
				btPapel1.Enabled = false;
				btPiedra1.Enabled = false;
				btTijera1.Enabled = false;
				btResultado2.Text = Jugar();

				if (btResultado1.Text == "Piedra") btPiedra1.Visible = false;
				if (btResultado1.Text == "Papel") btPapel1.Visible = false;
				if (btResultado1.Text == "Tijera") btTijera1.Visible = false;

				if (btResultado2.Text == "Piedra") btPiedra2.Visible = false;
				if (btResultado2.Text == "Papel") btPapel2.Visible = false;
				if (btResultado2.Text == "Tijera") btTijera2.Visible = false;

				btResultado1.Visible = true;
				iAncho = (btResultado1.Location.X - btPiedra1.Location.X) / iPasos;
				btResultado2.Visible = true;
				iAncho2 = (btResultado2.Location.X - btPiedra2.Location.X) / iPasos;

				if (btResultado1.Text == "Piedra")
				{
					iAlto = (btResultado1.Location.Y - btPiedra1.Location.Y) / iPasos;
					btResultado1.Location = new Point(btPiedra1.Location.X, btPiedra1.Location.Y);
				}
				if (btResultado2.Text == "Piedra")
				{
					iAlto2 = (btResultado2.Location.Y - btPiedra2.Location.Y) / iPasos;
					btResultado2.Location = new Point(btPiedra2.Location.X, btPiedra2.Location.Y);
				}

				if (btResultado1.Text == "Papel")
				{
					iAlto = (btResultado1.Location.Y - btPapel1.Location.Y) / iPasos;
					btResultado1.Location = new Point(btPapel1.Location.X, btPapel1.Location.Y);
				}
				if (btResultado2.Text == "Papel")
				{
					iAlto2 = (btResultado2.Location.Y - btPapel2.Location.Y) / iPasos;
					btResultado2.Location = new Point(btPapel2.Location.X, btPapel2.Location.Y);
				}

				if (btResultado1.Text == "Tijera")
				{
					iAlto = (btResultado1.Location.Y - btTijera1.Location.Y) / iPasos;
					btResultado1.Location = new Point(btTijera1.Location.X, btTijera1.Location.Y);
				}
				if (btResultado2.Text == "Tijera")
				{
					iAlto2 = (btResultado2.Location.Y - btTijera2.Location.Y) / iPasos;
					btResultado2.Location = new Point(btTijera2.Location.X, btTijera2.Location.Y);
				}

//				txtSalidas.Text += btResultado2.Text + "\r\n";
				timer2.Enabled = true;
			}
		}

		private void btPiedra1_Click(object sender, System.EventArgs e)
		{
			btResultado1.Text = "Piedra";
		}

		private void btPapel1_Click(object sender, System.EventArgs e)
		{
			btResultado1.Text = "Papel";
		}

		private void btTijera1_Click(object sender, System.EventArgs e)
		{
			btResultado1.Text = "Tijera";
		}

		int iTiempoRes = 0;
		int iAncho = 0;
		int iAlto = 0;
		int iAncho2 = 0;
		int iAlto2 = 0;
		private int iPasos = 4;
		private void timer2_Tick(object sender, System.EventArgs e)
		{
			iTiempoRes ++;
			
			if (iTiempoRes == iPasos + 2)
			{
				Resultado();
			}
			else if (iTiempoRes <= iPasos)
			{
				if (btResultado1.Text != "")
					btResultado1.Location = new Point(btResultado1.Location.X + iAncho, btResultado1.Location.Y +iAlto);
				btResultado2.Location = new Point(btResultado2.Location.X + iAncho2, btResultado2.Location.Y +iAlto2);
			}
			if (iTiempoRes >= iPasos * 6)
			{
				iTiempoRes = 0;
				timer2.Enabled = false;
				btIniciar.Enabled = true;
				if (btResultado1.Text == "Piedra") btPiedra1.Visible = true;
				if (btResultado1.Text == "Papel") btPapel1.Visible = true;
				if (btResultado1.Text == "Tijera") btTijera1.Visible = true;
				if (btResultado2.Text == "Piedra") btPiedra2.Visible = true;
				if (btResultado2.Text == "Papel") btPapel2.Visible = true;
				if (btResultado2.Text == "Tijera") btTijera2.Visible = true;
				btResultado1.Visible = false;
				btResultado2.Visible = false;
				btResultado1.Text = "";
			}
		}

		Point Res1;
		Point Res2;
		private void PiedraPapel_Load(object sender, System.EventArgs e)
		{
			Res1 = btResultado1.Location;
			Res2 = btResultado2.Location;
			LlenaPalabras();
		}

		private string GeneraTextoColor()
		{
			string [] Colores = {"Rojo", "Verde", "Azul", "Amarillo"};
			int iRand = GeneraRandomico()/64;
			return Colores[iRand];
		}
		private Color GeneraColor()
		{
			Color [] Colores = new Color[4];
			Colores[0] = Color.Red;
			Colores[1] = Color.Yellow;
			Colores[2] = Color.Green;
			Colores[3] = Color.Blue;
			int iRand = GeneraRandomico()/64;
			return Colores[iRand];
		}

		private void LlenaPalabras()
		{
			lblPalabra1.Text = GeneraTextoColor();
			lblPalabra2.Text = GeneraTextoColor();
			lblPalabra3.Text = GeneraTextoColor();
			lblPalabra4.Text = GeneraTextoColor();
			lblPalabra5.Text = GeneraTextoColor();
			lblPalabra6.Text = GeneraTextoColor();
			lblPalabra7.Text = GeneraTextoColor();
			lblPalabra8.Text = GeneraTextoColor();
			lblPalabra9.Text = GeneraTextoColor();
			lblPalabra10.Text = GeneraTextoColor();
			lblPalabra11.Text = GeneraTextoColor();
			lblPalabra12.Text = GeneraTextoColor();
			lblPalabra13.Text = GeneraTextoColor();
			lblPalabra14.Text = GeneraTextoColor();
			lblPalabra15.Text = GeneraTextoColor();
			lblPalabra16.Text = GeneraTextoColor();
			lblPalabra17.Text = GeneraTextoColor();
			lblPalabra18.Text = GeneraTextoColor();
			lblPalabra19.Text = GeneraTextoColor();
			lblPalabra20.Text = GeneraTextoColor();

			lblPalabra1.ForeColor = GeneraColor();
			lblPalabra2.ForeColor = GeneraColor();
			lblPalabra3.ForeColor = GeneraColor();
			lblPalabra4.ForeColor = GeneraColor();
			lblPalabra5.ForeColor = GeneraColor();
			lblPalabra6.ForeColor = GeneraColor();
			lblPalabra7.ForeColor = GeneraColor();
			lblPalabra8.ForeColor = GeneraColor();
			lblPalabra9.ForeColor = GeneraColor();
			lblPalabra10.ForeColor = GeneraColor();
			lblPalabra11.ForeColor = GeneraColor();
			lblPalabra12.ForeColor = GeneraColor();
			lblPalabra13.ForeColor = GeneraColor();
			lblPalabra14.ForeColor = GeneraColor();
			lblPalabra15.ForeColor = GeneraColor();
			lblPalabra16.ForeColor = GeneraColor();
			lblPalabra17.ForeColor = GeneraColor();
			lblPalabra18.ForeColor = GeneraColor();
			lblPalabra19.ForeColor = GeneraColor();
			lblPalabra20.ForeColor = GeneraColor();
		}

		private int GeneraRandomico()
		{
			byte[] randomNumber = new byte[1];
			RNGCryptoServiceProvider Gen = new RNGCryptoServiceProvider();
			Gen.GetBytes(randomNumber);
			return Convert.ToInt32(randomNumber[0]);
		}

		private void EscogerLado(bool bizquierda)
		{
			bIzquierda = bizquierda;
			if (btIniciarCM.Text != "Ejecucion")
			{
				btIniciarCM.Text = "Ejecucion";
				iNumIzquierdos = 0;
				iNumDerechos = 0;
				if (chkNumeros.Checked)
				{
					lblIzquierda.Visible = true;
					lblDerecha.Visible = true;
					imgCM.Visible = false;
				}
				else
				{
					lblIzquierda.Visible = false;
					lblDerecha.Visible = false;
					imgCM.Visible = true;
				}
			}
			timer3.Enabled = true;
			btIniciarCM.Enabled = false;
		}

		private void btIniciarCM_Click(object sender, System.EventArgs e)
		{
			EscogerLado(true);
		}

		int iNumIzquierdos = 0;
		int iNumDerechos = 0;
		int iNumCM= 0;
		bool bIzquierda;
		private void timer3_Tick(object sender, System.EventArgs e)
		{
			iNumCM ++;
			if (iNumCM == 51)
			{
				iNumCM = 0;
				btIniciarCM.Text = "Izquierda";
				timer3.Enabled = false;
				if (iNumIzquierdos > iNumDerechos) 
				{
					imgCM.Location = lblIzquierda.Location;
					if (bIzquierda) MessageBox.Show("Felicitaciones Moviste la Foto a la Izquierda");
					else MessageBox.Show("Intenta imaginando que se mueve donde deseas");
				}
				else 
				{
					imgCM.Location = lblDerecha.Location;
					if (!bIzquierda) MessageBox.Show("Felicitaciones Moviste la Foto a la Derecha");
					else MessageBox.Show("Intenta imaginando que se mueve donde deseas");
				}
				lblIzquierda.Text = iNumIzquierdos.ToString();
				lblDerecha.Text = iNumDerechos.ToString();
				btIniciarCM.Enabled = true;
				btDerecha.Enabled = true;
				return;
			}
			int iMin = 1000;
			int iMax = 0;
			int iCuentaI = 0;
			int iCuentaD = 0;
			int iNumRep = 100000;
			for (int i=0; i < iNumRep; i++)
			{
        int iRand = GeneraRandomico();
				if (iRand < iMin) iMin = iRand;
				if (iRand > iMax) iMax = iRand;
				if (iRand > 127) iCuentaD ++;
				else iCuentaI ++;
			}
			double dCuentaD = iCuentaD * 100.0 / iNumRep;
			double dCuentaI = iCuentaI * 100.0 / iNumRep;
			if (dCuentaD > dCuentaI) iNumIzquierdos++;
			if (dCuentaD < dCuentaI) iNumDerechos++;
			if (chkNumeros.Checked)
			{
				lblIzquierda.Text = dCuentaI.ToString("0.00");
				lblDerecha.Text = dCuentaD.ToString("0.00");
			}
			else
			{
				if (dCuentaD > dCuentaI) imgCM.Location = lblIzquierda.Location;
				if (dCuentaD < dCuentaI) imgCM.Location = lblDerecha.Location;
			}
		}

		private void btDerecha_Click(object sender, System.EventArgs e)
		{
			EscogerLado(false);
			btDerecha.Enabled = false;
		}

		private bool ComparaColorPalabra(int i, string stLetra)
		{
			string stColor = "";
			if (i== 0) stColor = lblPalabra1.ForeColor.Name;
			if (i== 1) stColor = lblPalabra2.ForeColor.Name;
			if (i== 2) stColor = lblPalabra3.ForeColor.Name;
			if (i== 3) stColor = lblPalabra4.ForeColor.Name;
			if (i== 4) stColor = lblPalabra5.ForeColor.Name;
			if (i== 5) stColor = lblPalabra6.ForeColor.Name;
			if (i== 6) stColor = lblPalabra7.ForeColor.Name;
			if (i== 7) stColor = lblPalabra8.ForeColor.Name;
			if (i== 8) stColor = lblPalabra9.ForeColor.Name;
			if (i== 9) stColor = lblPalabra10.ForeColor.Name;
			if (i== 10) stColor = lblPalabra11.ForeColor.Name;
			if (i== 11) stColor = lblPalabra12.ForeColor.Name;
			if (i== 12) stColor = lblPalabra13.ForeColor.Name;
			if (i== 13) stColor = lblPalabra14.ForeColor.Name;
			if (i== 14) stColor = lblPalabra15.ForeColor.Name;
			if (i== 15) stColor = lblPalabra16.ForeColor.Name;
			if (i== 16) stColor = lblPalabra17.ForeColor.Name;
			if (i== 17) stColor = lblPalabra18.ForeColor.Name;
			if (i== 18) stColor = lblPalabra19.ForeColor.Name;
			if (i== 19) stColor = lblPalabra20.ForeColor.Name;
			string stInicial = "A";
			if (stColor == "Red") stInicial = "R";
			if (stColor == "Green") stInicial = "V";
			if (stColor == "Blue") stInicial = "Z";
			if (stLetra == stInicial) return true;
			else return false;
		}

		private void btResultadoColor_Click(object sender, System.EventArgs e)
		{
			if (btResultadoColor.Text == "Empezar")
			{
				LlenaPalabras();
				btResultadoColor.Text = "Resultado";
				txtColores.Properties.ReadOnly = false;
				txtColores.Text = "";
			}
			else
			{
				txtColores.Properties.ReadOnly = true;
				string stLetras = txtColores.Text;
				int iCorrecto = 0;
				for(int i=0;i<stLetras.Length;i++)
				{
					string stLetra = stLetras.Substring(i, 1).ToUpper();
					if (ComparaColorPalabra(i, stLetra)) iCorrecto++;
				}
				MessageBox.Show("Numero de palabrar correctas " + iCorrecto, "Informacion");
				btResultadoColor.Text = "Empezar";
			}
		}

		DateTime dtInicio;
		private bool bVerdad = false;
		private void IniciaParImpar(bool bverdad)
		{
			bVerdad = bverdad;
			dtInicio= DateTime.Now;
			lblInstruccionVerdad.Visible = true;
			if (!bVerdad)
				lblInstruccionVerdad.Text = "Mienta sobre el numero que ve. Es par o Impar?";
			else
				lblInstruccionVerdad.Text = "Diga la Verdad sobre el numero que ve.";

			btPar.Visible = true;
			btImpar.Visible = true;
			btVerdad.Visible = false;
			btMentir.Visible = false;
			NumeroParImpar();
		}
		private void btMentir_Click(object sender, System.EventArgs e)
		{
			IniciaParImpar(false);
		}

		private int iNumeroParImpar = 0;
		private void NumeroParImpar()
		{
			int iRand = GeneraRandomico();
			double dRand = iRand * 100.0 / 255.0;
			lblVerdad.Text =  dRand.ToString("00");
			lblCorrecto.Text = "Correctos: " + iCorrecto;
			lblIncorrecto.Text = "Incorrectos: " + iInCorrecto;
			if (++iNumeroParImpar == 11)
			{
				iNumeroParImpar = 0;
				DateTime dtFin = DateTime.Now;
				TimeSpan ts = dtFin.Subtract(dtInicio);
				MessageBox.Show("Tiempo " + ts.TotalSeconds.ToString("0.00"));
				btPar.Visible = false;
				btImpar.Visible = false;
				btVerdad.Visible = true;
				btMentir.Visible = true;
				iCorrecto = 0;
				iInCorrecto = 0;
			}
		}

		private int iCorrecto = 0;
		private int iInCorrecto = 0;
		private void btPar_Click(object sender, System.EventArgs e)
		{
			int iPar = int.Parse(lblVerdad.Text);
			if((iPar % 2) == 0) iCorrecto++;
			else iInCorrecto++;
			NumeroParImpar();
		}

		private void btImpar_Click(object sender, System.EventArgs e)
		{
			int iImPar = int.Parse(lblVerdad.Text);
			if((iImPar % 2) != 0) iCorrecto++;
			else iInCorrecto++;
			NumeroParImpar();		
		}

		private void btVerdad_Click(object sender, System.EventArgs e)
		{
			IniciaParImpar(true);
		}

		private void btEmpezarMano_Click(object sender, System.EventArgs e)
		{
			timer4.Enabled = true;
		}

		string stArchivo = @"c:\Latinium\Mano2.jpg";
		private void timer4_Tick(object sender, System.EventArgs e)
		{
			try
			{
				if (stArchivo == @"c:\Latinium\Mano1.jpg")
				{
					pictureBox1.Image =new Bitmap(@"c:\Latinium\Mano1.jpg");
					stArchivo = @"c:\Latinium\Mano2.jpg";
				}
				else
				{
					pictureBox1.Image =new Bitmap(@"c:\Latinium\Mano2.jpg");
					stArchivo = @"c:\Latinium\Mano1.jpg";
				}
			}
			catch(System.Exception ex)
			{
				timer4.Enabled = false;
				MessageBox.Show(ex.Message, "Informacion");

			}
		}
	}
}

