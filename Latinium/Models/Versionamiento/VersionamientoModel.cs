using System;

namespace Latinium.Models.Versionamiento
{
	/// <summary>
	/// Descripción breve de VersionamientoModel.
	/// </summary>
	public class VersionamientoModel
	{
		private int _intervalo;

		public int Intervalo
		{
			get { return _intervalo; }
			set { _intervalo = value; }
		}

		public VersionamientoModel()
		{
			_intervalo = 5; // valor por defecto
		}
	}
}
