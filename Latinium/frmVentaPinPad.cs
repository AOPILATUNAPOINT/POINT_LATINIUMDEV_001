using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using Latinium.Models;
using Latinium.Services;

namespace Latinium
{
	/// <summary>
	/// Descripción breve de frmVentaPinPad.
	/// </summary>
	public class frmVentaPinPad : System.Windows.Forms.Form
	{
		/// <summary>
		/// Variable del diseñador requerida.
		/// </summary>
		private System.ComponentModel.Container components = null;

		public frmVentaPinPad()
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
			// 
			// frmVentaPinPad
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(6, 15);
			this.ClientSize = new System.Drawing.Size(292, 260);
			this.Name = "frmVentaPinPad";
			this.Text = "frmVentaPinPad";
			this.Load += new System.EventHandler(this.frmVentaPinPad_Load);

		}
		#endregion

		private void frmVentaPinPad_Load(object sender, System.EventArgs e)
		{
			PinpadService pinpadService = new PinpadService();
			pinpadService.ModoSimulado = true; // Activar modo simulación

			LecturaTarjetaRequest request = new LecturaTarjetaRequest();
			request.MontoTotal = 100.00m;

			LecturaTarjetaResponse response = pinpadService.LeerTarjeta(request);

			Console.WriteLine("Código Respuesta: " + response.CodigoRespuesta);
			Console.WriteLine("Mensaje: " + response.MensajeRespuesta);
			Console.WriteLine("Número Tarjeta: " + response.NumeroTarjetaTruncado);
			Console.WriteLine("Fecha Vencimiento: " + response.FechaVencimiento);

			MessageBox.Show("Código Respuesta: " + response.CodigoRespuesta + "\n" +
				"Mensaje: " + response.MensajeRespuesta + "\n" +
				"Número Tarjeta: " + response.NumeroTarjetaTruncado + "\n" +
				"Fecha Vencimiento: " + response.FechaVencimiento);
		}
	}
}
