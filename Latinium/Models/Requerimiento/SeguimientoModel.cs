using System;

namespace Latinium.Models.Requerimiento
{
	public class SeguimientoModel
	{
		// Campos privados
		private string _Notas;
		private int _idPersonal;
		private int _idRequerimiento;
		private int _idEstado;
		private int _idAsignado;
		private string _NotaInterna;
		private string _UsuarioRespondeNota;

		private int _idRequerimientoUrgencia;
		private int _idRequerimientoComplejidad;
		private DateTime _FechaEntregaAdm;
		private DateTime _FechaEntregaTec;

		private int _idRequerimientoDepartamento;
		private int _idRequerimientoTipo;

		private string _NombreCompletoCol;
		private DateTime _FechaIngresoCol;
		private string _CargoAsignadoCol;

		// Constructor
		public SeguimientoModel()
		{
		}

		// Propiedades

		public string Notas
		{
			get { return _Notas; }
			set { _Notas = value; }
		}

		public int idPersonal
		{
			get { return _idPersonal; }
			set { _idPersonal = value; }
		}

		public int idRequerimiento
		{
			get { return _idRequerimiento; }
			set { _idRequerimiento = value; }
		}

		public int idEstado
		{
			get { return _idEstado; }
			set { _idEstado = value; }
		}

		public int idAsignado
		{
			get { return _idAsignado; }
			set { _idAsignado = value; }
		}

		public string NotaInterna
		{
			get { return _NotaInterna; }
			set { _NotaInterna = value; }
		}

		public string UsuarioRespondeNota
		{
			get { return _UsuarioRespondeNota; }
			set { _UsuarioRespondeNota = value; }
		}

		public int idRequerimientoUrgencia
		{
			get { return _idRequerimientoUrgencia; }
			set { _idRequerimientoUrgencia = value; }
		}

		public int idRequerimientoComplejidad
		{
			get { return _idRequerimientoComplejidad; }
			set { _idRequerimientoComplejidad = value; }
		}

		public DateTime FechaEntregaAdm
		{
			get { return _FechaEntregaAdm; }
			set { _FechaEntregaAdm = value; }
		}

		public DateTime FechaEntregaTec
		{
			get { return _FechaEntregaTec; }
			set { _FechaEntregaTec = value; }
		}

		public int idRequerimientoDepartamento
		{
			get { return _idRequerimientoDepartamento; }
			set { _idRequerimientoDepartamento = value; }
		}

		public int idRequerimientoTipo
		{
			get { return _idRequerimientoTipo; }
			set { _idRequerimientoTipo = value; }
		}

		public string NombreCompletoCol
		{
			get { return _NombreCompletoCol; }
			set { _NombreCompletoCol = value; }
		}

		public DateTime FechaIngresoCol
		{
			get { return _FechaIngresoCol; }
			set { _FechaIngresoCol = value; }
		}

		public string CargoAsignadoCol
		{
			get { return _CargoAsignadoCol; }
			set { _CargoAsignadoCol = value; }
		}
	}
}