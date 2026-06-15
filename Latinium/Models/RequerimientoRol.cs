using System;

namespace Latinium.Models
{
	/// <summary>
	/// Representa los roles definidos para los requerimientos.
	/// </summary>
	public class RequerimientoRol
	{
		private int idRequerimientoRol;
		private string descripcion;
		private bool estado;
		private string subtitulo;

		public RequerimientoRol()
		{
			// Constructor por defecto
		}

		public int IdRequerimientoRol
		{
			get { return idRequerimientoRol; }
			set { idRequerimientoRol = value; }
		}

		public string Descripcion
		{
			get { return descripcion; }
			set { descripcion = value; }
		}

		public bool Estado
		{
			get { return estado; }
			set { estado = value; }
		}

		public string Subtitulo
		{
			get { return Subtitulo; }
			set { subtitulo = value; }
		}
	}
}
