using System;
using System.Collections;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.ServiceProcess;

namespace Latinium
{
	public class Service1 : System.ServiceProcess.ServiceBase
	{
		/// <summary> 
		/// Variable del diseñador requerida.
		/// </summary>
		private System.ComponentModel.Container components = null;

		public Service1()
		{
			// Llamada necesaria para el Diseñador de componentes Windows.Forms.
			InitializeComponent();

			// TODO: agregar cualquier inicialización después de llamar a InitComponent
		}

		#region Código generado por el Diseñador de componentes
		/// <summary> 
		/// Método necesario para admitir el Diseñador. No se puede modificar 
		/// el contenido del método con el editor de código.
		/// </summary>
		private void InitializeComponent()
		{
			components = new System.ComponentModel.Container();
			this.ServiceName = "Service1";
		}
		#endregion

		/// <summary>
		/// Limpiar los recursos que se estén utilizando.
		/// </summary>
		protected override void Dispose( bool disposing ) 
		{
			if( disposing )
			{
				if (components != null) 
				{
					components.Dispose();
				}
			}
			base.Dispose( disposing );
		}

		/// <summary>
		/// Poner en movimiento los elementos para que el servicio pueda funcionar.
		/// </summary>
		protected override void OnStart(string[] args)
		{
			// TODO: agregar aquí código para iniciar el servicio.
		}
 
		/// <summary>
		/// Detener el servicio.
		/// </summary>
		protected override void OnStop()
		{
			// TODO: agregar código aquí para realizar cualquier anulación necesaria para detener el servicio.
		}
	}
}
