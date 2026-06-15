using System;

namespace Latinium.Models
{
	/// <summary>
	/// Representa la relación entre departamento y grupo para acceso a la ticketera.
	/// </summary>
	public class RequerimientoDepartamentoRol
	{
		private int idRequerimientoDepartamentoRol;
		private int idRequerimientoDepartamento;
		private int idGrupo;
		private bool estado;
		private string nombreGrupo;

		public RequerimientoDepartamentoRol()
		{
			// Constructor por defecto (puedes inicializar valores si lo deseas)
		}

		public int IdRequerimientoDepartamentoRol
		{
			get { return idRequerimientoDepartamentoRol; }
			set { idRequerimientoDepartamentoRol = value; }
		}

		public int IdRequerimientoDepartamento
		{
			get { return idRequerimientoDepartamento; }
			set { idRequerimientoDepartamento = value; }
		}

		public int IdGrupo
		{
			get { return idGrupo; }
			set { idGrupo = value; }
		}

		public bool Estado
		{
			get { return estado; }
			set { estado = value; }
		}

		public string NombreGrupo
		{
			get { return nombreGrupo; }
			set { nombreGrupo = value; }
		}
	}
}
