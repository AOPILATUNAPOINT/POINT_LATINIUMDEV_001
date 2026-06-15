using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;

namespace Latinium
{
	/// <summary>
	/// Descripción breve de frmRequerimientoNotificacionAlerta.
	/// </summary>
	public class frmRequerimientoNotificacionAlerta : System.Windows.Forms.Form
	{
		private AxWMPLib.AxWindowsMediaPlayer axWindowsMediaPlayer1;
		private System.Windows.Forms.Timer tRequerimientoNotAudio;
		private System.ComponentModel.IContainer components;

		public frmRequerimientoNotificacionAlerta()
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
			this.components = new System.ComponentModel.Container();
			System.Resources.ResourceManager resources = new System.Resources.ResourceManager(typeof(frmRequerimientoNotificacionAlerta));
			this.axWindowsMediaPlayer1 = new AxWMPLib.AxWindowsMediaPlayer();
			this.tRequerimientoNotAudio = new System.Windows.Forms.Timer(this.components);
			((System.ComponentModel.ISupportInitialize)(this.axWindowsMediaPlayer1)).BeginInit();
			this.SuspendLayout();
			// 
			// axWindowsMediaPlayer1
			// 
//			this.axWindowsMediaPlayer1.AccessibleDescription = resources.GetString("axWindowsMediaPlayer1.AccessibleDescription");
//			this.axWindowsMediaPlayer1.AccessibleName = resources.GetString("axWindowsMediaPlayer1.AccessibleName");
//			//this.axWindowsMediaPlayer1.Anchor = ((System.Windows.Forms.AnchorStyles)(resources.GetObject("axWindowsMediaPlayer1.Anchor")));
//			this.axWindowsMediaPlayer1.BackgroundImage = ((System.Drawing.Image)(resources.GetObject("axWindowsMediaPlayer1.BackgroundImage")));
//			//this.axWindowsMediaPlayer1.Dock = ((System.Windows.Forms.DockStyle)(resources.GetObject("axWindowsMediaPlayer1.Dock")));
//			this.axWindowsMediaPlayer1.Enabled = ((bool)(resources.GetObject("axWindowsMediaPlayer1.Enabled")));
//			this.axWindowsMediaPlayer1.Font = ((System.Drawing.Font)(resources.GetObject("axWindowsMediaPlayer1.Font")));
//			this.axWindowsMediaPlayer1.ImeMode = ((System.Windows.Forms.ImeMode)(resources.GetObject("axWindowsMediaPlayer1.ImeMode")));
//			this.axWindowsMediaPlayer1.Location = ((System.Drawing.Point)(resources.GetObject("axWindowsMediaPlayer1.Location")));
//			this.axWindowsMediaPlayer1.Name = "axWindowsMediaPlayer1";
//			this.axWindowsMediaPlayer1.OcxState = ((System.Windows.Forms.AxHost.State)(resources.GetObject("axWindowsMediaPlayer1.OcxState")));
//			this.axWindowsMediaPlayer1.RightToLeft = ((bool)(resources.GetObject("axWindowsMediaPlayer1.RightToLeft")));
//			this.axWindowsMediaPlayer1.Size = ((System.Drawing.Size)(resources.GetObject("axWindowsMediaPlayer1.Size")));
//			this.axWindowsMediaPlayer1.TabIndex = ((int)(resources.GetObject("axWindowsMediaPlayer1.TabIndex")));
//			this.axWindowsMediaPlayer1.Text = resources.GetString("axWindowsMediaPlayer1.Text");
//			this.axWindowsMediaPlayer1.Visible = ((bool)(resources.GetObject("axWindowsMediaPlayer1.Visible")));
			// 
			// tRequerimientoNotAudio
			// 
			this.tRequerimientoNotAudio.Enabled = true;
			this.tRequerimientoNotAudio.Tick += new System.EventHandler(this.tRequerimientoNotAudio_Tick);
			// 
			// frmRequerimientoNotificacionAlerta
			// 
//			this.AccessibleDescription = resources.GetString("$this.AccessibleDescription");
//			this.AccessibleName = resources.GetString("$this.AccessibleName");
//			this.AccessibleRole = System.Windows.Forms.AccessibleRole.None;
//			this.AutoScale = false;
//			this.AutoScaleBaseSize = ((System.Drawing.Size)(resources.GetObject("$this.AutoScaleBaseSize")));
//			this.AutoScroll = ((bool)(resources.GetObject("$this.AutoScroll")));
//			this.AutoScrollMargin = ((System.Drawing.Size)(resources.GetObject("$this.AutoScrollMargin")));
//			this.AutoScrollMinSize = ((System.Drawing.Size)(resources.GetObject("$this.AutoScrollMinSize")));
//			this.BackColor = System.Drawing.SystemColors.GrayText;
//			this.BackgroundImage = ((System.Drawing.Image)(resources.GetObject("$this.BackgroundImage")));
//			this.CausesValidation = false;
//			this.ClientSize = ((System.Drawing.Size)(resources.GetObject("$this.ClientSize")));
//			this.Controls.Add(this.axWindowsMediaPlayer1);
//			this.Enabled = ((bool)(resources.GetObject("$this.Enabled")));
//			this.Font = ((System.Drawing.Font)(resources.GetObject("$this.Font")));
//			this.ForeColor = System.Drawing.Color.FromArgb(((System.Byte)(25)), ((System.Byte)(25)), ((System.Byte)(25)));
//			this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.None;
//			this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
//			this.ImeMode = ((System.Windows.Forms.ImeMode)(resources.GetObject("$this.ImeMode")));
//			this.Location = ((System.Drawing.Point)(resources.GetObject("$this.Location")));
//			this.MaximizeBox = false;
//			this.MaximumSize = ((System.Drawing.Size)(resources.GetObject("$this.MaximumSize")));
//			this.MinimumSize = ((System.Drawing.Size)(resources.GetObject("$this.MinimumSize")));
//			this.Name = "frmRequerimientoNotificacionAlerta";
//			this.Opacity = 0;
//			this.RightToLeft = ((System.Windows.Forms.RightToLeft)(resources.GetObject("$this.RightToLeft")));
//			this.ShowInTaskbar = false;
//			this.StartPosition = ((System.Windows.Forms.FormStartPosition)(resources.GetObject("$this.StartPosition")));
//			this.Text = resources.GetString("$this.Text");
//			this.Load += new System.EventHandler(this.frmRequerimientoNotificacionAlerta_Load);
//			((System.ComponentModel.ISupportInitialize)(this.axWindowsMediaPlayer1)).EndInit();
//			this.ResumeLayout(false);

		}
		#endregion

		public static int iContadorNotificaciones = 0;
		int iContadorUrl = 0;

		private void frmRequerimientoNotificacionAlerta_Load(object sender, System.EventArgs e)
		{
			int Color1 = Funcion.iEjecucion(frmRequerimiento.sconexionPoint, String.Format("EXEC RequerimientoConfiguracionModel '{0}', {1}", "Color1", frmRequerimiento.iConfiguracionAudioNotificacionAlerta),false);
			int Color2 = Funcion.iEjecucion(frmRequerimiento.sconexionPoint, String.Format("EXEC RequerimientoConfiguracionModel '{0}', {1}", "Color2", frmRequerimiento.iConfiguracionAudioNotificacionAlerta),false);
			int Color3 = Funcion.iEjecucion(frmRequerimiento.sconexionPoint, String.Format("EXEC RequerimientoConfiguracionModel '{0}', {1}", "Color3", frmRequerimiento.iConfiguracionAudioNotificacionAlerta),false);
			this.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(Color1)), ((System.Byte)(Color2)), ((System.Byte)(Color3)));;
			//Intervalo de refresh al timer parametrizado
			this.tRequerimientoNotAudio.Interval = Funcion.iEjecucion(frmRequerimiento.sconexionPoint, String.Format("EXEC RequerimientoConfiguracionModel '{0}', {1}", "Valor", frmRequerimiento.iConfiguracionTimerbtnRequerimiento),false);
			iContadorNotificaciones = frmRequerimiento.RequerimientoNotificacionNumero(MenuLatinium.idUsuarioRequerimiento);
			//Notificación al inicio si tiene notificaciones en cola
			NotificacionesAudioInicio();

			//Mostrar formulario requerimiento modal.
//			using (frmRequerimientoNotificacionModal FRMNOTIMODAL = new frmRequerimientoNotificacionModal())
//			{		
//				if (DialogResult.OK == FRMNOTIMODAL.ShowDialog())
//				{
//				}
//				NotificacionesAudioInicio();
//			}

		}

		public void NotificacionesAudioInicio()
		{
			if(iContadorNotificaciones > 0)
			{
				if(iContadorUrl == 0)
				{
					axWindowsMediaPlayer1.URL = Funcion.sEjecucion(frmRequerimiento.sconexionPoint, String.Format("EXEC RequerimientoConfiguracionModel '{0}', {1}", "Excepcion", frmRequerimiento.iConfiguracionAudioNotificacion));
					iContadorUrl = 1;
				}
				axWindowsMediaPlayer1.Ctlcontrols.play();
			}

		}

		public void NotificacionesAudio()
		{
			int getiContadorNotificaciones = frmRequerimiento.RequerimientoNotificacionNumero(MenuLatinium.idUsuarioRequerimiento);
			if(getiContadorNotificaciones > 0)
			{
				if(getiContadorNotificaciones > iContadorNotificaciones)
				{
//					axWindowsMediaPlayer1.Ctlcontrols.play();
					iContadorNotificaciones = getiContadorNotificaciones; //frmRequerimiento.RequerimientoNotificacionNumero(MenuLatinium.idUsuarioRequerimiento);
				}
			}
		}
		private void tRequerimientoNotAudio_Tick(object sender, System.EventArgs e)
		{
			NotificacionesAudio();
		}
	}
}
