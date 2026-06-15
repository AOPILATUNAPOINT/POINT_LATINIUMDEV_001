using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;

namespace Latinium
{
	/// <summary>
	/// Descripción breve de frmReporteNominaGeneral.
	/// </summary>
	public class frmReporteNominaGeneral : System.Windows.Forms.Form
	{
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource1;
		public Infragistics.Win.UltraWinGrid.UltraGrid uGridNomina;
		private System.Windows.Forms.Label label7;
		private System.Windows.Forms.Label label6;
		private System.Windows.Forms.Button btnVer;
		private Infragistics.Win.UltraWinSchedule.UltraCalendarCombo dtpFecha;
		private System.Windows.Forms.Button btnExportar;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtBuscar;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbNEEstadoCivil;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource4;
		private System.Windows.Forms.Label label2;
		private Infragistics.Win.UltraWinEditors.UltraOptionSet optEstados;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource9;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbNECargo;
		private System.Windows.Forms.Label label13;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbNEDepartamento;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource8;
		private System.Windows.Forms.Label label12;
		private Infragistics.Win.UltraWinEditors.UltraOptionSet optDiscapacidad;
		private Infragistics.Win.UltraWinEditors.UltraOptionSet optAplicaARol;
		private Infragistics.Win.UltraWinEditors.UltraOptionSet optFormaDePago;
		private Infragistics.Win.UltraWinEditors.UltraOptionSet optCargasFamiliares;
		private Infragistics.Win.UltraWinEditors.UltraOptionSet optAFDR;
		private Infragistics.Win.UltraWinEditors.UltraOptionSet optDFR;
		private Infragistics.Win.UltraWinEditors.UltraOptionSet optAportaIESS;
		private Infragistics.Win.UltraWinEditors.UltraOptionSet optComisiona;
		private Infragistics.Win.UltraWinEditors.UltraOptionSet optVentas;
		private System.Windows.Forms.Label lblContador;
		private C1.Data.C1DataSet cdsSeteoF;
		private System.Windows.Forms.GroupBox groupBox1;
		/// <summary>
		/// Variable del diseñador requerida.
		/// </summary>
		private System.ComponentModel.Container components = null;

		public frmReporteNominaGeneral()
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
			Infragistics.Win.Appearance appearance1 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand1 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn1 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idNomina");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn2 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Estado");
			Infragistics.Win.Appearance appearance2 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn3 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Codigo");
			Infragistics.Win.Appearance appearance3 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn4 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("NIdentificacion");
			Infragistics.Win.Appearance appearance4 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn5 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Empleado");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn6 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Prov_Nacimiento");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn7 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Ciud_Nacimiento");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn8 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("FechaNacimiento");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn9 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Cumpleaños");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn10 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Sexo");
			Infragistics.Win.Appearance appearance5 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn11 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("TipoDeSangre");
			Infragistics.Win.Appearance appearance6 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn12 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("EstadoCivil");
			Infragistics.Win.Appearance appearance7 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn13 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Cargas");
			Infragistics.Win.Appearance appearance8 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn14 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Discapacidad");
			Infragistics.Win.Appearance appearance9 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn15 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("PorcentajeDiscapacidad");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn16 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("NumeroConadis");
			Infragistics.Win.Appearance appearance10 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn17 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("TipoDiscapacidad");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn18 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Prov_Domicilio");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn19 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Ciud_Domicilio");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn20 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Parr_Domicilio");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn21 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Barrio_Domicilio");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn22 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Direccion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn23 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Telefono");
			Infragistics.Win.Appearance appearance11 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn24 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Celular");
			Infragistics.Win.Appearance appearance12 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn25 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("EMailPersonal");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn26 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("EMailEmpresarial");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn27 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Fecha_Entrada");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn28 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("FechaTerminaPrueba");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn29 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Fecha_Salida");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn30 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Centro_Costo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn31 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Departamento");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn32 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Cargo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn33 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Codigo_Sectorial");
			Infragistics.Win.Appearance appearance13 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn34 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Tipo_Pago");
			Infragistics.Win.Appearance appearance14 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn35 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Banco");
			Infragistics.Win.Appearance appearance15 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn36 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("NumeroDeCuenta");
			Infragistics.Win.Appearance appearance16 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn37 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Sueldo");
			Infragistics.Win.Appearance appearance17 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn38 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Bonos");
			Infragistics.Win.Appearance appearance18 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn39 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Movilizacion");
			Infragistics.Win.Appearance appearance19 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn40 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("AplicaRol");
			Infragistics.Win.Appearance appearance20 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn41 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("AporteIESS");
			Infragistics.Win.Appearance appearance21 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn42 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("DerechoFondosReserva");
			Infragistics.Win.Appearance appearance22 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn43 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("FondosDeReservaMensuales");
			Infragistics.Win.Appearance appearance23 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn44 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Vendedor");
			Infragistics.Win.Appearance appearance24 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn45 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Comisiona");
			Infragistics.Win.Appearance appearance25 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance26 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance27 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance28 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance29 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance30 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance31 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn1 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idNomina");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn2 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Estado");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn3 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Codigo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn4 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("NIdentificacion");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn5 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Empleado");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn6 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Prov_Nacimiento");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn7 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Ciud_Nacimiento");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn8 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("FechaNacimiento");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn9 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Cumpleaños");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn10 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Sexo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn11 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("TipoDeSangre");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn12 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("EstadoCivil");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn13 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Cargas");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn14 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Discapacidad");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn15 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("PorcentajeDiscapacidad");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn16 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("NumeroConadis");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn17 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("TipoDiscapacidad");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn18 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Prov_Domicilio");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn19 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Ciud_Domicilio");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn20 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Parr_Domicilio");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn21 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Barrio_Domicilio");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn22 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Direccion");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn23 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Telefono");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn24 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Celular");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn25 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("EMailPersonal");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn26 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("EMailEmpresarial");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn27 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Fecha_Entrada");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn28 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("FechaTerminaPrueba");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn29 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Fecha_Salida");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn30 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Centro_Costo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn31 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Departamento");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn32 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Cargo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn33 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Codigo_Sectorial");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn34 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Tipo_Pago");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn35 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Banco");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn36 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("NumeroDeCuenta");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn37 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Sueldo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn38 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Bonos");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn39 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Movilizacion");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn40 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("AplicaRol");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn41 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("AporteIESS");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn42 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("DerechoFondosReserva");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn43 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("FondosDeReservaMensuales");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn44 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Vendedor");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn45 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Comisiona");
			Infragistics.Win.Appearance appearance32 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance33 = new Infragistics.Win.Appearance();
			Infragistics.Win.ValueListItem valueListItem1 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.ValueListItem valueListItem2 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.ValueListItem valueListItem3 = new Infragistics.Win.ValueListItem();
			System.Resources.ResourceManager resources = new System.Resources.ResourceManager(typeof(frmReporteNominaGeneral));
			Infragistics.Win.Appearance appearance34 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton dateButton1 = new Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand2 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn46 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idEdoCivil");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn47 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn46 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idEdoCivil");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn47 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Nombre");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn48 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idCargo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn49 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Descripcion");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn50 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Com_IdCargo");
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand3 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn48 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCargo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn49 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Descripcion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn50 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Com_IdCargo");
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand4 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn51 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idDepartamento");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn52 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Descripcion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn53 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idModulo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn51 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idDepartamento");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn52 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Descripcion");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn53 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idModulo");
			Infragistics.Win.Appearance appearance35 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance36 = new Infragistics.Win.Appearance();
			Infragistics.Win.ValueListItem valueListItem4 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.ValueListItem valueListItem5 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.ValueListItem valueListItem6 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.Appearance appearance37 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance38 = new Infragistics.Win.Appearance();
			Infragistics.Win.ValueListItem valueListItem7 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.ValueListItem valueListItem8 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.ValueListItem valueListItem9 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.Appearance appearance39 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance40 = new Infragistics.Win.Appearance();
			Infragistics.Win.ValueListItem valueListItem10 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.ValueListItem valueListItem11 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.ValueListItem valueListItem12 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.ValueListItem valueListItem13 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.Appearance appearance41 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance42 = new Infragistics.Win.Appearance();
			Infragistics.Win.ValueListItem valueListItem14 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.ValueListItem valueListItem15 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.ValueListItem valueListItem16 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.Appearance appearance43 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance44 = new Infragistics.Win.Appearance();
			Infragistics.Win.ValueListItem valueListItem17 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.ValueListItem valueListItem18 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.ValueListItem valueListItem19 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.Appearance appearance45 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance46 = new Infragistics.Win.Appearance();
			Infragistics.Win.ValueListItem valueListItem20 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.ValueListItem valueListItem21 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.ValueListItem valueListItem22 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.Appearance appearance47 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance48 = new Infragistics.Win.Appearance();
			Infragistics.Win.ValueListItem valueListItem23 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.ValueListItem valueListItem24 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.ValueListItem valueListItem25 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.Appearance appearance49 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance50 = new Infragistics.Win.Appearance();
			Infragistics.Win.ValueListItem valueListItem26 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.ValueListItem valueListItem27 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.ValueListItem valueListItem28 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.Appearance appearance51 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance52 = new Infragistics.Win.Appearance();
			Infragistics.Win.ValueListItem valueListItem29 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.ValueListItem valueListItem30 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.ValueListItem valueListItem31 = new Infragistics.Win.ValueListItem();
			this.uGridNomina = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.ultraDataSource1 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.optEstados = new Infragistics.Win.UltraWinEditors.UltraOptionSet();
			this.label7 = new System.Windows.Forms.Label();
			this.label6 = new System.Windows.Forms.Label();
			this.btnVer = new System.Windows.Forms.Button();
			this.dtpFecha = new Infragistics.Win.UltraWinSchedule.UltraCalendarCombo();
			this.btnExportar = new System.Windows.Forms.Button();
			this.txtBuscar = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.cmbNEEstadoCivil = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.ultraDataSource4 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.label2 = new System.Windows.Forms.Label();
			this.ultraDataSource9 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.cmbNECargo = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.label13 = new System.Windows.Forms.Label();
			this.cmbNEDepartamento = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.ultraDataSource8 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.label12 = new System.Windows.Forms.Label();
			this.optDiscapacidad = new Infragistics.Win.UltraWinEditors.UltraOptionSet();
			this.optAplicaARol = new Infragistics.Win.UltraWinEditors.UltraOptionSet();
			this.optFormaDePago = new Infragistics.Win.UltraWinEditors.UltraOptionSet();
			this.optCargasFamiliares = new Infragistics.Win.UltraWinEditors.UltraOptionSet();
			this.optAFDR = new Infragistics.Win.UltraWinEditors.UltraOptionSet();
			this.optDFR = new Infragistics.Win.UltraWinEditors.UltraOptionSet();
			this.optAportaIESS = new Infragistics.Win.UltraWinEditors.UltraOptionSet();
			this.optComisiona = new Infragistics.Win.UltraWinEditors.UltraOptionSet();
			this.optVentas = new Infragistics.Win.UltraWinEditors.UltraOptionSet();
			this.lblContador = new System.Windows.Forms.Label();
			this.cdsSeteoF = new C1.Data.C1DataSet();
			this.groupBox1 = new System.Windows.Forms.GroupBox();
			((System.ComponentModel.ISupportInitialize)(this.uGridNomina)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource1)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.optEstados)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.dtpFecha)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtBuscar)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbNEEstadoCivil)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource4)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource9)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbNECargo)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbNEDepartamento)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource8)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.optDiscapacidad)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.optAplicaARol)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.optFormaDePago)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.optCargasFamiliares)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.optAFDR)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.optDFR)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.optAportaIESS)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.optComisiona)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.optVentas)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).BeginInit();
			this.SuspendLayout();
			// 
			// uGridNomina
			// 
			this.uGridNomina.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.uGridNomina.Cursor = System.Windows.Forms.Cursors.Default;
			this.uGridNomina.DataSource = this.ultraDataSource1;
			appearance1.BackColor = System.Drawing.Color.White;
			this.uGridNomina.DisplayLayout.Appearance = appearance1;
			ultraGridBand1.AddButtonCaption = "DetalleFormaPago";
			ultraGridColumn1.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn1.Header.VisiblePosition = 0;
			ultraGridColumn1.Hidden = true;
			ultraGridColumn2.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance2.TextHAlign = Infragistics.Win.HAlign.Center;
			ultraGridColumn2.CellAppearance = appearance2;
			ultraGridColumn2.Header.VisiblePosition = 1;
			ultraGridColumn2.Width = 80;
			ultraGridColumn3.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance3.TextHAlign = Infragistics.Win.HAlign.Center;
			ultraGridColumn3.CellAppearance = appearance3;
			ultraGridColumn3.Header.VisiblePosition = 2;
			ultraGridColumn3.Width = 80;
			ultraGridColumn4.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance4.TextHAlign = Infragistics.Win.HAlign.Center;
			ultraGridColumn4.CellAppearance = appearance4;
			ultraGridColumn4.Header.Caption = "Cedula";
			ultraGridColumn4.Header.VisiblePosition = 3;
			ultraGridColumn4.Width = 80;
			ultraGridColumn5.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn5.Header.VisiblePosition = 4;
			ultraGridColumn5.Width = 220;
			ultraGridColumn6.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn6.Header.Caption = "Prov. Nacimiento";
			ultraGridColumn6.Header.VisiblePosition = 5;
			ultraGridColumn6.Width = 140;
			ultraGridColumn7.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn7.Header.Caption = "Ciud. Nacimiento";
			ultraGridColumn7.Header.VisiblePosition = 6;
			ultraGridColumn7.Width = 140;
			ultraGridColumn8.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn8.Header.Caption = "Fec. Nacim.";
			ultraGridColumn8.Header.VisiblePosition = 7;
			ultraGridColumn8.Width = 80;
			ultraGridColumn9.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn9.Header.VisiblePosition = 8;
			ultraGridColumn9.Width = 95;
			ultraGridColumn10.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance5.TextHAlign = Infragistics.Win.HAlign.Center;
			ultraGridColumn10.CellAppearance = appearance5;
			ultraGridColumn10.Header.VisiblePosition = 9;
			ultraGridColumn10.Width = 80;
			ultraGridColumn11.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance6.TextHAlign = Infragistics.Win.HAlign.Center;
			ultraGridColumn11.CellAppearance = appearance6;
			ultraGridColumn11.Header.Caption = "Tipo De Sangre";
			ultraGridColumn11.Header.VisiblePosition = 10;
			ultraGridColumn11.Width = 90;
			ultraGridColumn12.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance7.TextHAlign = Infragistics.Win.HAlign.Center;
			ultraGridColumn12.CellAppearance = appearance7;
			ultraGridColumn12.Header.Caption = "Estado Civil";
			ultraGridColumn12.Header.VisiblePosition = 11;
			ultraGridColumn12.Width = 100;
			ultraGridColumn13.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance8.TextHAlign = Infragistics.Win.HAlign.Center;
			ultraGridColumn13.CellAppearance = appearance8;
			ultraGridColumn13.Format = "#,##0";
			ultraGridColumn13.Header.VisiblePosition = 12;
			ultraGridColumn13.Width = 50;
			ultraGridColumn14.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance9.TextHAlign = Infragistics.Win.HAlign.Center;
			ultraGridColumn14.CellAppearance = appearance9;
			ultraGridColumn14.Header.VisiblePosition = 13;
			ultraGridColumn14.Width = 80;
			ultraGridColumn15.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn15.Header.Caption = "% Discapacidad";
			ultraGridColumn15.Header.VisiblePosition = 14;
			ultraGridColumn16.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance10.TextHAlign = Infragistics.Win.HAlign.Center;
			ultraGridColumn16.CellAppearance = appearance10;
			ultraGridColumn16.Header.Caption = "Numero Conadis";
			ultraGridColumn16.Header.VisiblePosition = 15;
			ultraGridColumn16.Width = 100;
			ultraGridColumn17.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn17.Header.Caption = "Tipo Discapacidad";
			ultraGridColumn17.Header.VisiblePosition = 16;
			ultraGridColumn17.Width = 120;
			ultraGridColumn18.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn18.Header.Caption = "Prov. Domicilio";
			ultraGridColumn18.Header.VisiblePosition = 17;
			ultraGridColumn18.Width = 140;
			ultraGridColumn19.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn19.Header.Caption = "Ciud. Domicilio";
			ultraGridColumn19.Header.VisiblePosition = 18;
			ultraGridColumn19.Width = 140;
			ultraGridColumn20.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn20.Header.Caption = "Parr. Domicilio";
			ultraGridColumn20.Header.VisiblePosition = 19;
			ultraGridColumn20.Width = 140;
			ultraGridColumn21.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn21.Header.Caption = "Barrio Domicilio";
			ultraGridColumn21.Header.VisiblePosition = 20;
			ultraGridColumn21.Width = 140;
			ultraGridColumn22.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn22.Header.VisiblePosition = 21;
			ultraGridColumn22.Width = 200;
			ultraGridColumn23.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance11.TextHAlign = Infragistics.Win.HAlign.Center;
			ultraGridColumn23.CellAppearance = appearance11;
			ultraGridColumn23.Header.VisiblePosition = 22;
			ultraGridColumn23.Width = 80;
			ultraGridColumn24.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance12.TextHAlign = Infragistics.Win.HAlign.Center;
			ultraGridColumn24.CellAppearance = appearance12;
			ultraGridColumn24.Header.VisiblePosition = 23;
			ultraGridColumn24.Width = 80;
			ultraGridColumn25.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn25.Header.Caption = "E-Mail Personal";
			ultraGridColumn25.Header.VisiblePosition = 24;
			ultraGridColumn25.Width = 120;
			ultraGridColumn26.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn26.Header.Caption = "E-Mail Empresarial";
			ultraGridColumn26.Header.VisiblePosition = 25;
			ultraGridColumn26.Width = 120;
			ultraGridColumn27.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn27.Header.Caption = "Fecha De Entrada";
			ultraGridColumn27.Header.VisiblePosition = 26;
			ultraGridColumn27.Width = 100;
			ultraGridColumn28.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn28.Header.Caption = "Termino de Prueba";
			ultraGridColumn28.Header.VisiblePosition = 27;
			ultraGridColumn28.Width = 105;
			ultraGridColumn29.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn29.Header.Caption = "Fecha De Salida";
			ultraGridColumn29.Header.VisiblePosition = 28;
			ultraGridColumn29.Width = 100;
			ultraGridColumn30.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn30.Header.Caption = "Centro de Costo";
			ultraGridColumn30.Header.VisiblePosition = 29;
			ultraGridColumn30.Width = 140;
			ultraGridColumn31.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn31.Header.VisiblePosition = 30;
			ultraGridColumn31.Width = 140;
			ultraGridColumn32.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn32.Header.VisiblePosition = 31;
			ultraGridColumn32.Width = 140;
			ultraGridColumn33.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance13.TextHAlign = Infragistics.Win.HAlign.Center;
			ultraGridColumn33.CellAppearance = appearance13;
			ultraGridColumn33.Header.Caption = "Codigo Sectorial";
			ultraGridColumn33.Header.VisiblePosition = 32;
			ultraGridColumn33.Width = 140;
			ultraGridColumn34.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance14.TextHAlign = Infragistics.Win.HAlign.Center;
			ultraGridColumn34.CellAppearance = appearance14;
			ultraGridColumn34.Header.Caption = "Forma de Pago";
			ultraGridColumn34.Header.VisiblePosition = 33;
			ultraGridColumn34.Width = 100;
			ultraGridColumn35.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance15.TextHAlign = Infragistics.Win.HAlign.Center;
			ultraGridColumn35.CellAppearance = appearance15;
			ultraGridColumn35.Header.VisiblePosition = 34;
			ultraGridColumn35.Width = 100;
			ultraGridColumn36.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance16.TextHAlign = Infragistics.Win.HAlign.Center;
			ultraGridColumn36.CellAppearance = appearance16;
			ultraGridColumn36.Header.Caption = "Cuenta Bancaria";
			ultraGridColumn36.Header.VisiblePosition = 35;
			ultraGridColumn36.Width = 120;
			ultraGridColumn37.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance17.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn37.CellAppearance = appearance17;
			ultraGridColumn37.Format = "#,##0.00";
			ultraGridColumn37.Header.VisiblePosition = 36;
			ultraGridColumn37.Width = 80;
			ultraGridColumn38.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance18.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn38.CellAppearance = appearance18;
			ultraGridColumn38.Format = "#,##0.00";
			ultraGridColumn38.Header.VisiblePosition = 37;
			ultraGridColumn38.Width = 80;
			ultraGridColumn39.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance19.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn39.CellAppearance = appearance19;
			ultraGridColumn39.Format = "#,##0.00";
			ultraGridColumn39.Header.VisiblePosition = 38;
			ultraGridColumn39.Width = 80;
			ultraGridColumn40.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance20.TextHAlign = Infragistics.Win.HAlign.Center;
			ultraGridColumn40.CellAppearance = appearance20;
			ultraGridColumn40.Header.Caption = "Aplica a Rol";
			ultraGridColumn40.Header.VisiblePosition = 39;
			ultraGridColumn40.Width = 75;
			ultraGridColumn41.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance21.TextHAlign = Infragistics.Win.HAlign.Center;
			ultraGridColumn41.CellAppearance = appearance21;
			ultraGridColumn41.Header.Caption = "Aporte I.E.S.S.";
			ultraGridColumn41.Header.VisiblePosition = 40;
			ultraGridColumn41.Width = 90;
			ultraGridColumn42.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance22.TextHAlign = Infragistics.Win.HAlign.Center;
			ultraGridColumn42.CellAppearance = appearance22;
			ultraGridColumn42.Header.Caption = "Der Fondos de Reserva";
			ultraGridColumn42.Header.VisiblePosition = 41;
			ultraGridColumn42.Width = 130;
			ultraGridColumn43.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance23.TextHAlign = Infragistics.Win.HAlign.Center;
			ultraGridColumn43.CellAppearance = appearance23;
			ultraGridColumn43.Header.Caption = "Acum. Fondos De Reserva";
			ultraGridColumn43.Header.VisiblePosition = 42;
			ultraGridColumn43.Width = 145;
			ultraGridColumn44.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance24.TextHAlign = Infragistics.Win.HAlign.Center;
			ultraGridColumn44.CellAppearance = appearance24;
			ultraGridColumn44.Header.VisiblePosition = 43;
			ultraGridColumn44.Width = 70;
			ultraGridColumn45.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance25.TextHAlign = Infragistics.Win.HAlign.Center;
			ultraGridColumn45.CellAppearance = appearance25;
			ultraGridColumn45.Header.VisiblePosition = 44;
			ultraGridColumn45.Width = 70;
			ultraGridBand1.Columns.AddRange(new object[] {
																										 ultraGridColumn1,
																										 ultraGridColumn2,
																										 ultraGridColumn3,
																										 ultraGridColumn4,
																										 ultraGridColumn5,
																										 ultraGridColumn6,
																										 ultraGridColumn7,
																										 ultraGridColumn8,
																										 ultraGridColumn9,
																										 ultraGridColumn10,
																										 ultraGridColumn11,
																										 ultraGridColumn12,
																										 ultraGridColumn13,
																										 ultraGridColumn14,
																										 ultraGridColumn15,
																										 ultraGridColumn16,
																										 ultraGridColumn17,
																										 ultraGridColumn18,
																										 ultraGridColumn19,
																										 ultraGridColumn20,
																										 ultraGridColumn21,
																										 ultraGridColumn22,
																										 ultraGridColumn23,
																										 ultraGridColumn24,
																										 ultraGridColumn25,
																										 ultraGridColumn26,
																										 ultraGridColumn27,
																										 ultraGridColumn28,
																										 ultraGridColumn29,
																										 ultraGridColumn30,
																										 ultraGridColumn31,
																										 ultraGridColumn32,
																										 ultraGridColumn33,
																										 ultraGridColumn34,
																										 ultraGridColumn35,
																										 ultraGridColumn36,
																										 ultraGridColumn37,
																										 ultraGridColumn38,
																										 ultraGridColumn39,
																										 ultraGridColumn40,
																										 ultraGridColumn41,
																										 ultraGridColumn42,
																										 ultraGridColumn43,
																										 ultraGridColumn44,
																										 ultraGridColumn45});
			this.uGridNomina.DisplayLayout.BandsSerializer.Add(ultraGridBand1);
			this.uGridNomina.DisplayLayout.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.No;
			this.uGridNomina.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.False;
			appearance26.BackColor = System.Drawing.Color.Transparent;
			this.uGridNomina.DisplayLayout.Override.CardAreaAppearance = appearance26;
			appearance27.ForeColor = System.Drawing.Color.Black;
			appearance27.ForeColorDisabled = System.Drawing.Color.Black;
			this.uGridNomina.DisplayLayout.Override.CellAppearance = appearance27;
			appearance28.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance28.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance28.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance28.FontData.BoldAsString = "True";
			appearance28.FontData.Name = "Arial";
			appearance28.FontData.SizeInPoints = 8F;
			appearance28.ForeColor = System.Drawing.Color.White;
			appearance28.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.uGridNomina.DisplayLayout.Override.HeaderAppearance = appearance28;
			this.uGridNomina.DisplayLayout.Override.HeaderClickAction = Infragistics.Win.UltraWinGrid.HeaderClickAction.SortSingle;
			appearance29.BackColor = System.Drawing.SystemColors.ControlLight;
			appearance29.BackColor2 = System.Drawing.Color.Gainsboro;
			appearance29.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridNomina.DisplayLayout.Override.RowAlternateAppearance = appearance29;
			appearance30.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance30.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance30.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridNomina.DisplayLayout.Override.RowSelectorAppearance = appearance30;
			appearance31.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance31.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance31.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridNomina.DisplayLayout.Override.SelectedRowAppearance = appearance31;
			this.uGridNomina.Font = new System.Drawing.Font("Microsoft Sans Serif", 7.5F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.uGridNomina.Location = new System.Drawing.Point(8, 184);
			this.uGridNomina.Name = "uGridNomina";
			this.uGridNomina.Size = new System.Drawing.Size(1072, 192);
			this.uGridNomina.TabIndex = 2;
			// 
			// ultraDataSource1
			// 
			ultraDataColumn1.DataType = typeof(int);
			ultraDataColumn8.DataType = typeof(System.DateTime);
			ultraDataColumn13.DataType = typeof(int);
			ultraDataColumn15.DataType = typeof(int);
			ultraDataColumn15.DefaultValue = 0;
			ultraDataColumn27.DataType = typeof(System.DateTime);
			ultraDataColumn28.DataType = typeof(System.DateTime);
			ultraDataColumn28.DefaultValue = new System.DateTime(((long)(0)));
			ultraDataColumn29.DataType = typeof(System.DateTime);
			ultraDataColumn37.DataType = typeof(System.Decimal);
			ultraDataColumn37.DefaultValue = new System.Decimal(new int[] {
																																			0,
																																			0,
																																			0,
																																			0});
			ultraDataColumn38.DataType = typeof(System.Decimal);
			ultraDataColumn38.DefaultValue = new System.Decimal(new int[] {
																																			0,
																																			0,
																																			0,
																																			0});
			ultraDataColumn39.DataType = typeof(System.Decimal);
			ultraDataColumn39.DefaultValue = new System.Decimal(new int[] {
																																			0,
																																			0,
																																			0,
																																			0});
			this.ultraDataSource1.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn1,
																																 ultraDataColumn2,
																																 ultraDataColumn3,
																																 ultraDataColumn4,
																																 ultraDataColumn5,
																																 ultraDataColumn6,
																																 ultraDataColumn7,
																																 ultraDataColumn8,
																																 ultraDataColumn9,
																																 ultraDataColumn10,
																																 ultraDataColumn11,
																																 ultraDataColumn12,
																																 ultraDataColumn13,
																																 ultraDataColumn14,
																																 ultraDataColumn15,
																																 ultraDataColumn16,
																																 ultraDataColumn17,
																																 ultraDataColumn18,
																																 ultraDataColumn19,
																																 ultraDataColumn20,
																																 ultraDataColumn21,
																																 ultraDataColumn22,
																																 ultraDataColumn23,
																																 ultraDataColumn24,
																																 ultraDataColumn25,
																																 ultraDataColumn26,
																																 ultraDataColumn27,
																																 ultraDataColumn28,
																																 ultraDataColumn29,
																																 ultraDataColumn30,
																																 ultraDataColumn31,
																																 ultraDataColumn32,
																																 ultraDataColumn33,
																																 ultraDataColumn34,
																																 ultraDataColumn35,
																																 ultraDataColumn36,
																																 ultraDataColumn37,
																																 ultraDataColumn38,
																																 ultraDataColumn39,
																																 ultraDataColumn40,
																																 ultraDataColumn41,
																																 ultraDataColumn42,
																																 ultraDataColumn43,
																																 ultraDataColumn44,
																																 ultraDataColumn45});
			// 
			// optEstados
			// 
			appearance32.ForeColorDisabled = System.Drawing.Color.Black;
			this.optEstados.Appearance = appearance32;
			this.optEstados.BorderStyle = Infragistics.Win.UIElementBorderStyle.Solid;
			this.optEstados.CheckedIndex = 1;
			this.optEstados.ItemAppearance = appearance33;
			this.optEstados.ItemOrigin = new System.Drawing.Point(10, 0);
			valueListItem1.DataValue = 0;
			valueListItem1.DisplayText = "Todos";
			valueListItem2.DataValue = 1;
			valueListItem2.DisplayText = "Activos";
			valueListItem3.DataValue = 2;
			valueListItem3.DisplayText = "Inactivos";
			this.optEstados.Items.Add(valueListItem1);
			this.optEstados.Items.Add(valueListItem2);
			this.optEstados.Items.Add(valueListItem3);
			this.optEstados.ItemSpacingVertical = 10;
			this.optEstados.Location = new System.Drawing.Point(416, 8);
			this.optEstados.Name = "optEstados";
			this.optEstados.Size = new System.Drawing.Size(312, 24);
			this.optEstados.TabIndex = 7;
			this.optEstados.Text = "Activos";
			// 
			// label7
			// 
			this.label7.AutoSize = true;
			this.label7.Location = new System.Drawing.Point(8, 137);
			this.label7.Name = "label7";
			this.label7.Size = new System.Drawing.Size(39, 16);
			this.label7.TabIndex = 755;
			this.label7.Text = "Buscar";
			this.label7.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// label6
			// 
			this.label6.AutoSize = true;
			this.label6.Location = new System.Drawing.Point(8, 10);
			this.label6.Name = "label6";
			this.label6.Size = new System.Drawing.Size(14, 16);
			this.label6.TabIndex = 752;
			this.label6.Text = "Al";
			this.label6.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// btnVer
			// 
			this.btnVer.CausesValidation = false;
			this.btnVer.Image = ((System.Drawing.Image)(resources.GetObject("btnVer.Image")));
			this.btnVer.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnVer.Location = new System.Drawing.Point(320, 134);
			this.btnVer.Name = "btnVer";
			this.btnVer.Size = new System.Drawing.Size(72, 23);
			this.btnVer.TabIndex = 1;
			this.btnVer.Text = "Ver";
			this.btnVer.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnVer.Click += new System.EventHandler(this.btnVer_Click);
			// 
			// dtpFecha
			// 
			appearance34.ForeColorDisabled = System.Drawing.Color.Black;
			this.dtpFecha.Appearance = appearance34;
			this.dtpFecha.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			dateButton1.Caption = "Today";
			this.dtpFecha.DateButtons.Add(dateButton1);
			this.dtpFecha.Format = "dd/MM/yyyy";
			this.dtpFecha.Location = new System.Drawing.Point(32, 8);
			this.dtpFecha.MonthOrientation = Infragistics.Win.UltraWinSchedule.MonthOrientation.DownThenAcross;
			this.dtpFecha.Name = "dtpFecha";
			this.dtpFecha.NonAutoSizeHeight = 23;
			this.dtpFecha.Size = new System.Drawing.Size(112, 21);
			this.dtpFecha.SpinButtonsVisible = true;
			this.dtpFecha.TabIndex = 2;
			this.dtpFecha.Value = ((object)(resources.GetObject("dtpFecha.Value")));
			this.dtpFecha.KeyDown += new System.Windows.Forms.KeyEventHandler(this.dtpFecha_KeyDown);
			// 
			// btnExportar
			// 
			this.btnExportar.CausesValidation = false;
			this.btnExportar.Enabled = false;
			this.btnExportar.Image = ((System.Drawing.Image)(resources.GetObject("btnExportar.Image")));
			this.btnExportar.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnExportar.Location = new System.Drawing.Point(320, 103);
			this.btnExportar.Name = "btnExportar";
			this.btnExportar.Size = new System.Drawing.Size(72, 23);
			this.btnExportar.TabIndex = 6;
			this.btnExportar.Text = "Exportar";
			this.btnExportar.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnExportar.Click += new System.EventHandler(this.btnExportar_Click);
			// 
			// txtBuscar
			// 
			this.txtBuscar.BorderStyle = Infragistics.Win.UIElementBorderStyle.Solid;
			this.txtBuscar.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtBuscar.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtBuscar.Location = new System.Drawing.Point(96, 136);
			this.txtBuscar.Name = "txtBuscar";
			this.txtBuscar.Size = new System.Drawing.Size(208, 19);
			this.txtBuscar.TabIndex = 0;
			this.txtBuscar.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtBuscar_KeyDown);
			// 
			// cmbNEEstadoCivil
			// 
			this.cmbNEEstadoCivil.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbNEEstadoCivil.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbNEEstadoCivil.DataSource = this.ultraDataSource4;
			ultraGridBand2.AddButtonCaption = "RolSucursal";
			ultraGridColumn46.Header.VisiblePosition = 0;
			ultraGridColumn46.Hidden = true;
			ultraGridColumn47.Header.VisiblePosition = 1;
			ultraGridColumn47.Width = 208;
			ultraGridBand2.Columns.AddRange(new object[] {
																										 ultraGridColumn46,
																										 ultraGridColumn47});
			this.cmbNEEstadoCivil.DisplayLayout.BandsSerializer.Add(ultraGridBand2);
			this.cmbNEEstadoCivil.DisplayMember = "Nombre";
			this.cmbNEEstadoCivil.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbNEEstadoCivil.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbNEEstadoCivil.Location = new System.Drawing.Point(96, 40);
			this.cmbNEEstadoCivil.Name = "cmbNEEstadoCivil";
			this.cmbNEEstadoCivil.Size = new System.Drawing.Size(208, 21);
			this.cmbNEEstadoCivil.TabIndex = 3;
			this.cmbNEEstadoCivil.ValueMember = "idEdoCivil";
			this.cmbNEEstadoCivil.KeyDown += new System.Windows.Forms.KeyEventHandler(this.cmbNEEstadoCivil_KeyDown);
			// 
			// ultraDataSource4
			// 
			ultraDataColumn46.DataType = typeof(int);
			this.ultraDataSource4.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn46,
																																 ultraDataColumn47});
			// 
			// label2
			// 
			this.label2.AutoSize = true;
			this.label2.BackColor = System.Drawing.Color.Transparent;
			this.label2.Location = new System.Drawing.Point(8, 42);
			this.label2.Name = "label2";
			this.label2.Size = new System.Drawing.Size(64, 16);
			this.label2.TabIndex = 758;
			this.label2.Text = "Estado Civil";
			this.label2.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// ultraDataSource9
			// 
			ultraDataColumn48.DataType = typeof(int);
			ultraDataColumn50.DataType = typeof(int);
			this.ultraDataSource9.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn48,
																																 ultraDataColumn49,
																																 ultraDataColumn50});
			// 
			// cmbNECargo
			// 
			this.cmbNECargo.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbNECargo.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbNECargo.DataSource = this.ultraDataSource9;
			ultraGridBand3.AddButtonCaption = "RolTipo";
			ultraGridColumn48.Header.VisiblePosition = 0;
			ultraGridColumn48.Hidden = true;
			ultraGridColumn49.Header.VisiblePosition = 1;
			ultraGridColumn49.Width = 208;
			ultraGridColumn50.Header.VisiblePosition = 2;
			ultraGridColumn50.Hidden = true;
			ultraGridBand3.Columns.AddRange(new object[] {
																										 ultraGridColumn48,
																										 ultraGridColumn49,
																										 ultraGridColumn50});
			this.cmbNECargo.DisplayLayout.BandsSerializer.Add(ultraGridBand3);
			this.cmbNECargo.DisplayMember = "Descripcion";
			this.cmbNECargo.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbNECargo.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbNECargo.Location = new System.Drawing.Point(96, 104);
			this.cmbNECargo.Name = "cmbNECargo";
			this.cmbNECargo.Size = new System.Drawing.Size(208, 21);
			this.cmbNECargo.TabIndex = 5;
			this.cmbNECargo.ValueMember = "idCargo";
			this.cmbNECargo.KeyDown += new System.Windows.Forms.KeyEventHandler(this.cmbNECargo_KeyDown);
			// 
			// label13
			// 
			this.label13.AutoSize = true;
			this.label13.BackColor = System.Drawing.Color.Transparent;
			this.label13.Location = new System.Drawing.Point(8, 106);
			this.label13.Name = "label13";
			this.label13.Size = new System.Drawing.Size(35, 16);
			this.label13.TabIndex = 762;
			this.label13.Text = "Cargo";
			// 
			// cmbNEDepartamento
			// 
			this.cmbNEDepartamento.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbNEDepartamento.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbNEDepartamento.DataSource = this.ultraDataSource8;
			ultraGridBand4.AddButtonCaption = "RolTipo";
			ultraGridColumn51.Header.VisiblePosition = 0;
			ultraGridColumn51.Hidden = true;
			ultraGridColumn52.Header.VisiblePosition = 1;
			ultraGridColumn52.Width = 208;
			ultraGridColumn53.Header.VisiblePosition = 2;
			ultraGridColumn53.Hidden = true;
			ultraGridBand4.Columns.AddRange(new object[] {
																										 ultraGridColumn51,
																										 ultraGridColumn52,
																										 ultraGridColumn53});
			this.cmbNEDepartamento.DisplayLayout.BandsSerializer.Add(ultraGridBand4);
			this.cmbNEDepartamento.DisplayMember = "Descripcion";
			this.cmbNEDepartamento.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbNEDepartamento.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbNEDepartamento.Location = new System.Drawing.Point(96, 72);
			this.cmbNEDepartamento.Name = "cmbNEDepartamento";
			this.cmbNEDepartamento.Size = new System.Drawing.Size(208, 21);
			this.cmbNEDepartamento.TabIndex = 4;
			this.cmbNEDepartamento.ValueMember = "idDepartamento";
			this.cmbNEDepartamento.KeyDown += new System.Windows.Forms.KeyEventHandler(this.cmbNEDepartamento_KeyDown);
			// 
			// ultraDataSource8
			// 
			ultraDataColumn51.DataType = typeof(int);
			ultraDataColumn53.DataType = typeof(int);
			this.ultraDataSource8.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn51,
																																 ultraDataColumn52,
																																 ultraDataColumn53});
			// 
			// label12
			// 
			this.label12.AutoSize = true;
			this.label12.BackColor = System.Drawing.Color.Transparent;
			this.label12.Location = new System.Drawing.Point(8, 74);
			this.label12.Name = "label12";
			this.label12.Size = new System.Drawing.Size(76, 16);
			this.label12.TabIndex = 760;
			this.label12.Text = "Departamento";
			// 
			// optDiscapacidad
			// 
			appearance35.ForeColorDisabled = System.Drawing.Color.Black;
			this.optDiscapacidad.Appearance = appearance35;
			this.optDiscapacidad.BorderStyle = Infragistics.Win.UIElementBorderStyle.Solid;
			this.optDiscapacidad.CheckedIndex = 0;
			this.optDiscapacidad.ItemAppearance = appearance36;
			this.optDiscapacidad.ItemOrigin = new System.Drawing.Point(10, 0);
			valueListItem4.DataValue = -1;
			valueListItem4.DisplayText = "Todos";
			valueListItem5.DataValue = 1;
			valueListItem5.DisplayText = "Con Discapacidad";
			valueListItem6.DataValue = 0;
			valueListItem6.DisplayText = "Sin Discapacidad";
			this.optDiscapacidad.Items.Add(valueListItem4);
			this.optDiscapacidad.Items.Add(valueListItem5);
			this.optDiscapacidad.Items.Add(valueListItem6);
			this.optDiscapacidad.ItemSpacingVertical = 10;
			this.optDiscapacidad.Location = new System.Drawing.Point(416, 40);
			this.optDiscapacidad.Name = "optDiscapacidad";
			this.optDiscapacidad.Size = new System.Drawing.Size(312, 24);
			this.optDiscapacidad.TabIndex = 9;
			this.optDiscapacidad.Text = "Todos";
			// 
			// optAplicaARol
			// 
			appearance37.ForeColorDisabled = System.Drawing.Color.Black;
			this.optAplicaARol.Appearance = appearance37;
			this.optAplicaARol.BorderStyle = Infragistics.Win.UIElementBorderStyle.Solid;
			this.optAplicaARol.CheckedIndex = 0;
			this.optAplicaARol.ItemAppearance = appearance38;
			this.optAplicaARol.ItemOrigin = new System.Drawing.Point(10, 0);
			valueListItem7.DataValue = -1;
			valueListItem7.DisplayText = "Todos";
			valueListItem8.DataValue = 1;
			valueListItem8.DisplayText = "Aplica Rol";
			valueListItem9.DataValue = 0;
			valueListItem9.DisplayText = "No Aplica a Rol";
			this.optAplicaARol.Items.Add(valueListItem7);
			this.optAplicaARol.Items.Add(valueListItem8);
			this.optAplicaARol.Items.Add(valueListItem9);
			this.optAplicaARol.ItemSpacingVertical = 10;
			this.optAplicaARol.Location = new System.Drawing.Point(416, 136);
			this.optAplicaARol.Name = "optAplicaARol";
			this.optAplicaARol.Size = new System.Drawing.Size(312, 24);
			this.optAplicaARol.TabIndex = 15;
			this.optAplicaARol.Text = "Todos";
			// 
			// optFormaDePago
			// 
			appearance39.ForeColorDisabled = System.Drawing.Color.Black;
			this.optFormaDePago.Appearance = appearance39;
			this.optFormaDePago.BorderStyle = Infragistics.Win.UIElementBorderStyle.Solid;
			this.optFormaDePago.CheckedIndex = 0;
			this.optFormaDePago.ItemAppearance = appearance40;
			this.optFormaDePago.ItemOrigin = new System.Drawing.Point(10, 0);
			valueListItem10.DataValue = 0;
			valueListItem10.DisplayText = "Todos";
			valueListItem11.DataValue = 1;
			valueListItem11.DisplayText = "Ahorros";
			valueListItem12.DataValue = 2;
			valueListItem12.DisplayText = "Corriente";
			valueListItem13.DataValue = 4;
			valueListItem13.DisplayText = "Cheque";
			this.optFormaDePago.Items.Add(valueListItem10);
			this.optFormaDePago.Items.Add(valueListItem11);
			this.optFormaDePago.Items.Add(valueListItem12);
			this.optFormaDePago.Items.Add(valueListItem13);
			this.optFormaDePago.ItemSpacingVertical = 10;
			this.optFormaDePago.Location = new System.Drawing.Point(416, 104);
			this.optFormaDePago.Name = "optFormaDePago";
			this.optFormaDePago.Size = new System.Drawing.Size(312, 24);
			this.optFormaDePago.TabIndex = 13;
			this.optFormaDePago.Text = "Todos";
			// 
			// optCargasFamiliares
			// 
			appearance41.ForeColorDisabled = System.Drawing.Color.Black;
			this.optCargasFamiliares.Appearance = appearance41;
			this.optCargasFamiliares.BorderStyle = Infragistics.Win.UIElementBorderStyle.Solid;
			this.optCargasFamiliares.CheckedIndex = 0;
			this.optCargasFamiliares.ItemAppearance = appearance42;
			this.optCargasFamiliares.ItemOrigin = new System.Drawing.Point(10, 0);
			valueListItem14.DataValue = -1;
			valueListItem14.DisplayText = "Todos";
			valueListItem15.DataValue = 1;
			valueListItem15.DisplayText = "Con Cargas Famil.";
			valueListItem16.DataValue = 0;
			valueListItem16.DisplayText = "Sin Cargas Famil.";
			this.optCargasFamiliares.Items.Add(valueListItem14);
			this.optCargasFamiliares.Items.Add(valueListItem15);
			this.optCargasFamiliares.Items.Add(valueListItem16);
			this.optCargasFamiliares.ItemSpacingVertical = 10;
			this.optCargasFamiliares.Location = new System.Drawing.Point(416, 72);
			this.optCargasFamiliares.Name = "optCargasFamiliares";
			this.optCargasFamiliares.Size = new System.Drawing.Size(312, 24);
			this.optCargasFamiliares.TabIndex = 11;
			this.optCargasFamiliares.Text = "Todos";
			// 
			// optAFDR
			// 
			appearance43.ForeColorDisabled = System.Drawing.Color.Black;
			this.optAFDR.Appearance = appearance43;
			this.optAFDR.BorderStyle = Infragistics.Win.UIElementBorderStyle.Solid;
			this.optAFDR.CheckedIndex = 0;
			this.optAFDR.ItemAppearance = appearance44;
			this.optAFDR.ItemOrigin = new System.Drawing.Point(10, 0);
			valueListItem17.DataValue = -1;
			valueListItem17.DisplayText = "Todos";
			valueListItem18.DataValue = 1;
			valueListItem18.DisplayText = "Acumula F. R.";
			valueListItem19.DataValue = 0;
			valueListItem19.DisplayText = "No Acumula F. R.";
			this.optAFDR.Items.Add(valueListItem17);
			this.optAFDR.Items.Add(valueListItem18);
			this.optAFDR.Items.Add(valueListItem19);
			this.optAFDR.ItemSpacingVertical = 10;
			this.optAFDR.Location = new System.Drawing.Point(736, 72);
			this.optAFDR.Name = "optAFDR";
			this.optAFDR.Size = new System.Drawing.Size(344, 24);
			this.optAFDR.TabIndex = 12;
			this.optAFDR.Text = "Todos";
			// 
			// optDFR
			// 
			appearance45.ForeColorDisabled = System.Drawing.Color.Black;
			this.optDFR.Appearance = appearance45;
			this.optDFR.BorderStyle = Infragistics.Win.UIElementBorderStyle.Solid;
			this.optDFR.CheckedIndex = 0;
			this.optDFR.ItemAppearance = appearance46;
			this.optDFR.ItemOrigin = new System.Drawing.Point(10, 0);
			valueListItem20.DataValue = -1;
			valueListItem20.DisplayText = "Todos";
			valueListItem21.DataValue = 1;
			valueListItem21.DisplayText = "Tiene Derecho F. R.";
			valueListItem22.DataValue = 0;
			valueListItem22.DisplayText = "No Tiene Derecho F. R.";
			this.optDFR.Items.Add(valueListItem20);
			this.optDFR.Items.Add(valueListItem21);
			this.optDFR.Items.Add(valueListItem22);
			this.optDFR.ItemSpacingVertical = 10;
			this.optDFR.Location = new System.Drawing.Point(736, 40);
			this.optDFR.Name = "optDFR";
			this.optDFR.Size = new System.Drawing.Size(344, 24);
			this.optDFR.TabIndex = 10;
			this.optDFR.Text = "Todos";
			// 
			// optAportaIESS
			// 
			appearance47.ForeColorDisabled = System.Drawing.Color.Black;
			this.optAportaIESS.Appearance = appearance47;
			this.optAportaIESS.BorderStyle = Infragistics.Win.UIElementBorderStyle.Solid;
			this.optAportaIESS.CheckedIndex = 0;
			this.optAportaIESS.ItemAppearance = appearance48;
			this.optAportaIESS.ItemOrigin = new System.Drawing.Point(10, 0);
			valueListItem23.DataValue = -1;
			valueListItem23.DisplayText = "Todos";
			valueListItem24.DataValue = 1;
			valueListItem24.DisplayText = "Aporta I.E.S.S.";
			valueListItem25.DataValue = 0;
			valueListItem25.DisplayText = "No Aporta I.E.S.S.";
			this.optAportaIESS.Items.Add(valueListItem23);
			this.optAportaIESS.Items.Add(valueListItem24);
			this.optAportaIESS.Items.Add(valueListItem25);
			this.optAportaIESS.ItemSpacingVertical = 10;
			this.optAportaIESS.Location = new System.Drawing.Point(736, 8);
			this.optAportaIESS.Name = "optAportaIESS";
			this.optAportaIESS.Size = new System.Drawing.Size(344, 24);
			this.optAportaIESS.TabIndex = 8;
			this.optAportaIESS.Text = "Todos";
			// 
			// optComisiona
			// 
			appearance49.ForeColorDisabled = System.Drawing.Color.Black;
			this.optComisiona.Appearance = appearance49;
			this.optComisiona.BorderStyle = Infragistics.Win.UIElementBorderStyle.Solid;
			this.optComisiona.CheckedIndex = 0;
			this.optComisiona.ItemAppearance = appearance50;
			this.optComisiona.ItemOrigin = new System.Drawing.Point(10, 0);
			valueListItem26.DataValue = -1;
			valueListItem26.DisplayText = "Todos";
			valueListItem27.DataValue = 1;
			valueListItem27.DisplayText = "Comisiona";
			valueListItem28.DataValue = 0;
			valueListItem28.DisplayText = "No Comisiona";
			this.optComisiona.Items.Add(valueListItem26);
			this.optComisiona.Items.Add(valueListItem27);
			this.optComisiona.Items.Add(valueListItem28);
			this.optComisiona.ItemSpacingVertical = 10;
			this.optComisiona.Location = new System.Drawing.Point(736, 136);
			this.optComisiona.Name = "optComisiona";
			this.optComisiona.Size = new System.Drawing.Size(344, 24);
			this.optComisiona.TabIndex = 16;
			this.optComisiona.Text = "Todos";
			// 
			// optVentas
			// 
			appearance51.ForeColorDisabled = System.Drawing.Color.Black;
			this.optVentas.Appearance = appearance51;
			this.optVentas.BorderStyle = Infragistics.Win.UIElementBorderStyle.Solid;
			this.optVentas.CheckedIndex = 0;
			this.optVentas.ItemAppearance = appearance52;
			this.optVentas.ItemOrigin = new System.Drawing.Point(10, 0);
			valueListItem29.DataValue = -1;
			valueListItem29.DisplayText = "Todos";
			valueListItem30.DataValue = 1;
			valueListItem30.DisplayText = "Vendedor";
			valueListItem31.DataValue = 0;
			valueListItem31.DisplayText = "No Vendedor";
			this.optVentas.Items.Add(valueListItem29);
			this.optVentas.Items.Add(valueListItem30);
			this.optVentas.Items.Add(valueListItem31);
			this.optVentas.ItemSpacingVertical = 10;
			this.optVentas.Location = new System.Drawing.Point(736, 104);
			this.optVentas.Name = "optVentas";
			this.optVentas.Size = new System.Drawing.Size(344, 24);
			this.optVentas.TabIndex = 14;
			this.optVentas.Text = "Todos";
			// 
			// lblContador
			// 
			this.lblContador.AutoSize = true;
			this.lblContador.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.lblContador.ForeColor = System.Drawing.Color.Firebrick;
			this.lblContador.Location = new System.Drawing.Point(160, 10);
			this.lblContador.Name = "lblContador";
			this.lblContador.Size = new System.Drawing.Size(0, 17);
			this.lblContador.TabIndex = 773;
			this.lblContador.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
			// 
			// cdsSeteoF
			// 
			this.cdsSeteoF.BindingContextCtrl = this;
			this.cdsSeteoF.DataLibrary = "LibFacturacion";
			this.cdsSeteoF.DataLibraryUrl = "";
			this.cdsSeteoF.DataSetDef = "dsSeteoF";
			this.cdsSeteoF.Locale = new System.Globalization.CultureInfo("es-EC");
			this.cdsSeteoF.Name = "cdsSeteoF";
			this.cdsSeteoF.SchemaClassName = "LibFacturacion.DataClass";
			this.cdsSeteoF.SchemaDef = null;
			this.cdsSeteoF.BeforeFill += new C1.Data.FillEventHandler(this.cdsSeteoF_BeforeFill);
			// 
			// groupBox1
			// 
			this.groupBox1.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.groupBox1.Location = new System.Drawing.Point(0, 170);
			this.groupBox1.Name = "groupBox1";
			this.groupBox1.Size = new System.Drawing.Size(1088, 8);
			this.groupBox1.TabIndex = 774;
			this.groupBox1.TabStop = false;
			// 
			// frmReporteNominaGeneral
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 13);
			this.ClientSize = new System.Drawing.Size(1088, 382);
			this.Controls.Add(this.groupBox1);
			this.Controls.Add(this.lblContador);
			this.Controls.Add(this.optComisiona);
			this.Controls.Add(this.optVentas);
			this.Controls.Add(this.optCargasFamiliares);
			this.Controls.Add(this.optAFDR);
			this.Controls.Add(this.optDFR);
			this.Controls.Add(this.optAportaIESS);
			this.Controls.Add(this.optAplicaARol);
			this.Controls.Add(this.optFormaDePago);
			this.Controls.Add(this.optDiscapacidad);
			this.Controls.Add(this.cmbNEDepartamento);
			this.Controls.Add(this.label12);
			this.Controls.Add(this.label2);
			this.Controls.Add(this.txtBuscar);
			this.Controls.Add(this.optEstados);
			this.Controls.Add(this.label7);
			this.Controls.Add(this.label6);
			this.Controls.Add(this.btnVer);
			this.Controls.Add(this.dtpFecha);
			this.Controls.Add(this.btnExportar);
			this.Controls.Add(this.uGridNomina);
			this.Controls.Add(this.cmbNEEstadoCivil);
			this.Controls.Add(this.cmbNECargo);
			this.Controls.Add(this.label13);
			this.KeyPreview = true;
			this.Name = "frmReporteNominaGeneral";
			this.ShowInTaskbar = false;
			this.Text = "Nomina General";
			this.KeyDown += new System.Windows.Forms.KeyEventHandler(this.frmReporteNominaGeneral_KeyDown);
			this.Load += new System.EventHandler(this.frmReporteNominaGeneral_Load);
			((System.ComponentModel.ISupportInitialize)(this.uGridNomina)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource1)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.optEstados)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.dtpFecha)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtBuscar)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbNEEstadoCivil)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource4)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource9)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbNECargo)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbNEDepartamento)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource8)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.optDiscapacidad)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.optAplicaARol)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.optFormaDePago)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.optCargasFamiliares)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.optAFDR)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.optDFR)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.optAportaIESS)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.optComisiona)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.optVentas)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).EndInit();
			this.ResumeLayout(false);

		}
		#endregion
		private Acceso miAcceso;
		private void frmReporteNominaGeneral_Load(object sender, System.EventArgs e)
		{
	
		
			miAcceso = new Acceso (cdsSeteoF,"200304");

			if (miAcceso.BExportar) this.btnExportar.Enabled = true;
			this.dtpFecha.Value = DateTime.Today;
			this.dtpFecha.CalendarInfo.MinDate = DateTime.Parse("01/01/2015");

			FuncionesProcedimientos.RetornaFechaServidor(this.dtpFecha, cdsSeteoF, false);

			this.cmbNEEstadoCivil.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "Exec Listas 12, 0");
			this.cmbNEDepartamento.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "Exec Listas 48, 0");
			this.cmbNECargo.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "Exec Listas 49, 0");
		}

		private void btnExportar_Click(object sender, System.EventArgs e)
		{
			FuncionesProcedimientos.ExportaExcel(this.uGridNomina);
		}

		private void btnVer_Click(object sender, System.EventArgs e)
		{
			if (!Validacion.vbFechaEnDocumentos(this.dtpFecha, "Seleccione una Fecha", false, "", cdsSeteoF)) return;

			int idEstadoCivil = 0;
			int idDepartamento = 0;
			int idCargo = 0;

			if (this.cmbNEEstadoCivil.ActiveRow != null) idEstadoCivil = (int)this.cmbNEEstadoCivil.Value;
			if (this.cmbNEDepartamento.ActiveRow != null) idDepartamento = (int)this.cmbNEDepartamento.Value;
			if (this.cmbNECargo.ActiveRow != null) idCargo = (int)this.cmbNECargo.Value;

			string sSQL = string.Format("Exec ReporteNominaGeneral '{0}', '{1}', {2}, {3}, {4}, {5}, {6}, {7}, {8}, {9}, {10}, {11}, {12}, {13}, {14}",
				Convert.ToDateTime(this.dtpFecha.Value).ToString("yyyyMMdd"), this.txtBuscar.Text.ToString(),
				idEstadoCivil, idDepartamento, idCargo,
				(int)this.optEstados.Value, (int)this.optCargasFamiliares.Value, (int)this.optDiscapacidad.Value,
				(int)this.optFormaDePago.Value, (int)this.optAplicaARol.Value, (int)this.optAportaIESS.Value,
				(int)this.optDFR.Value, (int)this.optAFDR.Value,
				(int)this.optVentas.Value, (int)this.optComisiona.Value);
			FuncionesProcedimientos.dsGeneral(sSQL, this.uGridNomina);

			this.lblContador.Text = this.uGridNomina.Rows.Count + " REGISTROS ENCONTRADOS";
		}

		private void cdsSeteoF_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			cdsSeteoF.Schema.Connections[0].ConnectionString = miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);	
		}

		private void dtpFecha_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Enter) this.cmbNEEstadoCivil.Focus();
		}

		private void cmbNEEstadoCivil_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Enter) this.cmbNEDepartamento.Focus();
		}

		private void cmbNEDepartamento_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Enter) this.cmbNECargo.Focus();
		}

		private void cmbNECargo_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Enter) this.txtBuscar.Focus();
		}

		private void txtBuscar_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Enter) this.btnVer_Click(sender, e);
		}

		private void frmReporteNominaGeneral_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Escape) this.Close();
		}
	}
}
