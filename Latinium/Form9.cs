using System;
//using System.Collections.Generic;
using System.Text;
using System.Net;
using System.Net.Sockets;
using System.Text.RegularExpressions;
using System.Configuration;
using System.Threading;

namespace Latinium
{
	/// <summary>
	/// Descripción breve de Form9.
	/// </summary>
	public class Form9 : System.Windows.Forms.Form
	{
		private System.Windows.Forms.Button btnExportar;
		/// <summary>
		/// Variable del diseñador requerida.
		/// </summary>
		private System.ComponentModel.Container components = null;

		public Form9()
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
			System.Resources.ResourceManager resources = new System.Resources.ResourceManager(typeof(Form9));
			this.btnExportar = new System.Windows.Forms.Button();
			this.SuspendLayout();
			// 
			// btnExportar
			// 
			this.btnExportar.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			this.btnExportar.CausesValidation = false;
			this.btnExportar.Image = ((System.Drawing.Image)(resources.GetObject("btnExportar.Image")));
			this.btnExportar.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnExportar.ImeMode = System.Windows.Forms.ImeMode.NoControl;
			this.btnExportar.Location = new System.Drawing.Point(106, 122);
			this.btnExportar.Name = "btnExportar";
			this.btnExportar.Size = new System.Drawing.Size(80, 23);
			this.btnExportar.TabIndex = 382;
			this.btnExportar.Text = "&Exportar";
			this.btnExportar.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnExportar.Click += new System.EventHandler(this.btnExportar_Click);
			// 
			// Form9
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 13);
			this.ClientSize = new System.Drawing.Size(292, 266);
			this.Controls.Add(this.btnExportar);
			this.Name = "Form9";
			this.Text = "Form9";
			this.Load += new System.EventHandler(this.Form9_Load);
			this.ResumeLayout(false);

		}
		#endregion

		private void Form9_Load(object sender, System.EventArgs e)
		{
		
		}

		private void btnExportar_Click(object sender, System.EventArgs e)
		{
//			string m_numerotmp = "0982119820";
//			if (m_numerotmp.Length == 0)
//			{
//				Console.WriteLine("Number incorrect");
//				return;
//			}
//
//			string m_numero = string.Empty;
//			for (int i = 0; i < m_numerotmp.Length; i++)
//			{
//				int test = 0;
//				if (int.Parse(m_numerotmp.Substring(i,1),out test))
//				{
//					m_numero = m_numero + m_numerotmp.Substring(i, 1);
//				}
//			}

			

		/*	System.Configuration.Configuration config = ConfigurationManager.OpenExeConfiguration(ConfigurationUserLevel.None);
			ConfigurationSectionGroup grpApp = config.GetSectionGroup("applicationSettings");
			ClientSettingsSection secPropSet = (ClientSettingsSection)grpApp.Sections["EdiClick2Call.Properties.Settings"];
			string m_Extension = secPropSet.Settings.Get("extension").Value.ValueXml.InnerText;
			string m_ippbx = secPropSet.Settings.Get("ippbx").Value.ValueXml.InnerText;
			string m_user = secPropSet.Settings.Get("login").Value.ValueXml.InnerText;
			string m_pass = secPropSet.Settings.Get("password").Value.ValueXml.InnerText;*/

			/*Socket clientSocket = new Socket(AddressFamily.InterNetwork, SocketType.Stream, ProtocolType.Tcp);
			IPEndPoint serverEndPoint = new IPEndPoint(IPAddress.Parse(m_ippbx), 5038);
			clientSocket.Connect(serverEndPoint);
			clientSocket.Send(Encoding.ASCII.GetBytes("Action: Login\r\nUsername: " + m_user + "\r\nSecret: "+m_pass+"\r\nActionID: 1\r\n\r\n"));
			int bytesRead = 0;*/

			//do
			//{
				//byte[] buffer = new byte[1024];
				//bytesRead = clientSocket.Receive(buffer);

				//string response = Encoding.ASCII.GetString(buffer, 0, bytesRead);
				
			//	if(Regex.Match(response, "Message: Authentication accepted", RegexOptions.IgnoreCase).Success)
			//	{

					/*clientSocket.Send(Encoding.ASCII.GetBytes("Action: Originate\r\nChannel: SIP/" + m_Extension + "\r\nExten: " + m_numero + "\r\nContext: from-internal\r\nPriority: 1\r\nCallerID: " + m_Extension + "\r\n\r\n"));
					Console.WriteLine("Calling......");
					Thread.Sleep(1000);*/
			//		return;
			//	}
			//}
	//while(bytesRead != 0);		
		}
	}
}
