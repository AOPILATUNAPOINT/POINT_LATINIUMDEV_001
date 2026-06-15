using System;

namespace Latinium.Models
{
	/// <summary>
	/// Descripción breve de RequerimientoDepartamento.
	/// </summary>
	public class RequerimientoDepartamento
	{
		// Campos privados con nombres en minúscula
		private int idRequerimientoDepartamento;
		private string descripcion;
		private bool estado;
		private string usuario;
		private DateTime fechaSistema;
		private string estacion;
		private int grupo;
		private int idTipoFactura;
		private bool checklist;
		private bool pantalla;
		private bool bDescripcion;
		private bool asignacion;
		private bool todoUsuario;
		private bool bDepartamento;
		private bool asignacionAutomatica;
		private string usuarioAutomatico;
		private bool asignacionSoporte;
		private int documentosMinimo;
		private int cargoSolicitado;
		private bool asignacionPorDisponibilidad; // Corregí el nombre aquí, faltaba una 'a'

		// Propiedades públicas que exponen los campos privados.
		// El nombre de la propiedad comienza con mayúscula.
		public int IdRequerimientoDepartamento
		{
			get { return idRequerimientoDepartamento; }
			set { idRequerimientoDepartamento = value; }
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

		public string Usuario
		{
			get { return usuario; }
			set { usuario = value; }
		}

		public DateTime FechaSistema
		{
			get { return fechaSistema; }
			set { fechaSistema = value; }
		}

		public string Estacion
		{
			get { return estacion; }
			set { estacion = value; }
		}

		public int Grupo
		{
			get { return grupo; }
			set { grupo = value; }
		}

		public int IdTipoFactura
		{
			get { return idTipoFactura; }
			set { idTipoFactura = value; }
		}

		public bool Checklist
		{
			get { return checklist; }
			set { checklist = value; }
		}

		public bool Pantalla
		{
			get { return pantalla; }
			set { pantalla = value; }
		}

		public bool BDescripcion
		{
			get { return bDescripcion; }
			set { bDescripcion = value; }
		}

		public bool Asignacion
		{
			get { return asignacion; }
			set { asignacion = value; }
		}

		public bool TodoUsuario
		{
			get { return todoUsuario; }
			set { todoUsuario = value; }
		}

		public bool BDepartamento
		{
			get { return bDepartamento; }
			set { bDepartamento = value; }
		}

		public bool AsignacionAutomatica
		{
			get { return asignacionAutomatica; }
			set { asignacionAutomatica = value; }
		}

		public string UsuarioAutomatico
		{
			get { return usuarioAutomatico; }
			set { usuarioAutomatico = value; }
		}

		public bool AsignacionSoporte
		{
			get { return asignacionSoporte; }
			set { asignacionSoporte = value; }
		}

		public int DocumentosMinimo
		{
			get { return documentosMinimo; }
			set { documentosMinimo = value; }
		}

		public int CargoSolicitado
		{
			get { return cargoSolicitado; }
			set { cargoSolicitado = value; }
		}

		public bool AsignacionPorDisponibilidad
		{
			get { return asignacionPorDisponibilidad; }
			set { asignacionPorDisponibilidad = value; }
		}

		public RequerimientoDepartamento()
		{
			// Constructor vacío para Visual Studio 2003
		}

		// Aquí agregas el método ToString
		public override string ToString()
		{
			return this.Descripcion;
		}
	}
}