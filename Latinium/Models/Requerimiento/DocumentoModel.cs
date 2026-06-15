using System;

namespace Latinium.Models.Requerimiento
{
	/// <summary>
	/// Modelo alineado exactamente al SP RequerimientoDocumentoCon
	/// Compatible con .NET 1.1
	/// </summary>
	public class DocumentoModel
	{
		// Campos privados
		private int _idRequerimientoDocumento;
		private DateTime _fecha;
		private int _idSeguimiento;
		private string _txtDirectorioarchivo;
		private int _idResponsable;
		private int _idAsignado;
		private string _sCarpeta;
		private string _sUrlNube;
		private string _nombre;

		// Propiedades (MISMO NOMBRE QUE EL SP)
		public int idRequerimientoDocumento
		{
			get { return _idRequerimientoDocumento; }
			set { _idRequerimientoDocumento = value; }
		}

		public DateTime Fecha
		{
			get { return _fecha; }
			set { _fecha = value; }
		}

		public int idSeguimiento
		{
			get { return _idSeguimiento; }
			set { _idSeguimiento = value; }
		}

		public string txtDirectorioarchivo
		{
			get { return _txtDirectorioarchivo; }
			set { _txtDirectorioarchivo = value; }
		}

		public int idResponsable
		{
			get { return _idResponsable; }
			set { _idResponsable = value; }
		}

		public int idAsignado
		{
			get { return _idAsignado; }
			set { _idAsignado = value; }
		}

		public string sCarpeta
		{
			get { return _sCarpeta; }
			set { _sCarpeta = value; }
		}

		public string sUrlNube
		{
			get { return _sUrlNube; }
			set { _sUrlNube = value; }
		}

		public string Nombre
		{
			get { return _nombre; }
			set { _nombre = value; }
		}

		// Constructor vacío (OBLIGATORIO para grids y DataSet)
		public DocumentoModel()
		{
		}
	}
}
