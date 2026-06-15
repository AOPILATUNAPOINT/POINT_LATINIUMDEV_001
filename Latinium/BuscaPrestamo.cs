using System;
using System.Drawing;
using System.Globalization;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;


namespace Latinium
{
	/// <summary>
	/// Descripción breve de BuscaAsiento.
	/// </summary>
	public class BuscaPrestamo : DevExpress.XtraEditors.XtraForm
	{
		public System.Windows.Forms.TextBox idBusca;
		private Infragistics.Win.UltraWinGrid.UltraGrid ultraGrid1;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtCodigo;
		private Infragistics.Win.UltraWinEditors.UltraOptionSet rdgOpcion;
		private Infragistics.Win.UltraWinEditors.UltraOptionSet rdgTipoBusqueda;
		private Infragistics.Win.Misc.UltraButton btBuscar;
		private Infragistics.Win.Misc.UltraButton btAceptar;
		private Infragistics.Win.Misc.UltraButton btCancelar;
		private C1.Data.C1DataView cdvPersonal;
		private C1.Data.C1DataSet cdsPersonal;
		private System.Windows.Forms.ErrorProvider errorProvider;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbNombre;
		private Infragistics.Win.UltraWinSchedule.UltraCalendarCombo cmbFechaEntrega;
		private Infragistics.Win.UltraWinSchedule.UltraCalendarCombo cmbFechaCobro;
		private Infragistics.Win.UltraWinGrid.UltraDropDown ultraDropDown1;
		private C1.Data.C1DataView cdvRolPrestamo;
		private C1.Data.C1DataSet cdsRolPrestamo;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtNumero;
		private System.ComponentModel.IContainer components=null;

		public BuscaPrestamo()
		{
			//
			// Necesario para admitir el Diseñador de Windows Forms
			//
			InitializeComponent();

			//
			// TODO: Agregar código de constructor después de llamar a InitializeComponent
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

		#region Windows Form Designer generated code
		/// <summary>
		/// Método necesario para admitir el Diseñador, no se puede modificar
		/// el contenido del método con el editor de código.
		/// </summary>
		private void InitializeComponent()
		{
			Infragistics.Win.Appearance appearance1 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand1 = new Infragistics.Win.UltraWinGrid.UltraGridBand("RolPrestamo", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn1 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idRolPrestamo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn2 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idSucursal");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn3 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Capital");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn4 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("NumCuotas");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn5 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Interes");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn6 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idPersonal", -1, "ultraDropDown1");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn7 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("FechaEntrega");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn8 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("FechaCobro");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn9 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("FechaIngreso");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn10 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Cobrado");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn11 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Concepto");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn12 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("UltimoPago");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn13 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idAsiento");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn14 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idRolTipoPrestamo");
			Infragistics.Win.Appearance appearance2 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance3 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance4 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance5 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance6 = new Infragistics.Win.Appearance();
			Infragistics.Win.ValueListItem valueListItem1 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.ValueListItem valueListItem2 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.ValueListItem valueListItem3 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.ValueListItem valueListItem4 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.ValueListItem valueListItem5 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.Appearance appearance7 = new Infragistics.Win.Appearance();
			Infragistics.Win.ValueListItem valueListItem6 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.ValueListItem valueListItem7 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.ValueListItem valueListItem8 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand2 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Personal", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn15 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idPersonal");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn16 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idAsiDistGrupo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn17 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCentroCosto");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn18 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idDepartamento");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn19 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCargo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn20 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idRolSucursal");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn21 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idRolGrupo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn22 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Jornal");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn23 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn24 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Apellido");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn25 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Codigo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn26 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Jornada");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn27 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Direccion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn28 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Telefono");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn29 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Celular");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn30 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Cedula");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn31 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Afiliacion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn32 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Reserva");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn33 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Renta");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn34 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Iess");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn35 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Sueldo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn36 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Comision");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn37 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Tipo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn38 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("FechaIngreso");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn39 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("FechaNacimiento");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn40 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("FechaSalida");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn41 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Notas");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn42 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("CargasFamiliares");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn43 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idProvincia");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn44 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCiudad");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn45 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Numero");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn46 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("SistemaSalario");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn47 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Cuenta");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn48 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("CodigoSectorial");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn49 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Condicion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn50 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("FechaNovedad");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn51 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Acceso");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn52 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Vendedor");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn53 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Borrar");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn54 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Cobrador");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn55 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Sexo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn56 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Estado_Civil");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn57 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Banco");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn58 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("CuentaCorr");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn59 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idTipoCuenta");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn60 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("CedulaMilitar");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn61 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idFormaPago");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn62 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Quincena");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn63 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("CodIngreso");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn64 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("CodTelefono");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn65 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("TarjEdificio");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn66 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("TarjetaOficina");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn67 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("email");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn68 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idProyecto");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn69 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCuenta");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn70 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Bono");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn71 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Bono1");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn72 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Desc1");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn73 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Desc2");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn74 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("TieneQuincena");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn75 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Fotos");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn76 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Convenio");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn77 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idSubProyecto");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn78 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Sindicato");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn79 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Asociacion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn80 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idRolHorario");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn81 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("FechaContrato");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn82 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Educacion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn83 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("NumeroDias");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn84 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idRelacionT");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn85 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("gSalud");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn86 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("gEducacion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn87 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("gVivienda");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn88 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("gVestimenta");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn89 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("gAlimentacion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn90 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("gTotal");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn91 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("BaseComision");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn92 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("PorcentajeComision");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn93 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("gDesahucio");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn94 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("gDiscapacidad");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn95 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("IngGravOE");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn96 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("DeduccionOE");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn97 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("OtrasRebOE");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn98 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("ImpRentaAsumido");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn99 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("gRebajaIIIEdad");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn100 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("RentaRetenidaOE");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn101 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("RentaRetOE");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn102 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("TipoSangre");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn103 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("RegEntrada");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn104 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("RegSalida");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn105 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idTipoContrato");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn106 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("CodVendedor");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn107 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Clave");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn108 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idSupervisor");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn109 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("bSupervisor");
			Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton dateButton1 = new Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton();
			Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton dateButton2 = new Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand3 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Personal", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn110 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idPersonal");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn111 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idAsiDistGrupo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn112 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCentroCosto");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn113 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idDepartamento");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn114 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCargo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn115 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idRolSucursal");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn116 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idRolGrupo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn117 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Jornal");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn118 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn119 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Apellido");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn120 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Codigo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn121 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Jornada");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn122 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Direccion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn123 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Telefono");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn124 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Celular");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn125 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Cedula");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn126 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Afiliacion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn127 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Reserva");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn128 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Renta");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn129 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Iess");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn130 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Sueldo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn131 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Comision");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn132 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Tipo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn133 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("FechaIngreso");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn134 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("FechaNacimiento");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn135 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("FechaSalida");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn136 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Notas");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn137 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("CargasFamiliares");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn138 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idProvincia");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn139 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCiudad");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn140 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Numero");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn141 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("SistemaSalario");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn142 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Cuenta");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn143 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("CodigoSectorial");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn144 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Condicion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn145 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("FechaNovedad");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn146 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Acceso");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn147 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Vendedor");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn148 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Borrar");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn149 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Cobrador");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn150 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Sexo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn151 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Estado_Civil");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn152 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Banco");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn153 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("CuentaCorr");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn154 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idTipoCuenta");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn155 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("CedulaMilitar");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn156 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idFormaPago");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn157 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Quincena");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn158 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("CodIngreso");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn159 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("CodTelefono");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn160 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("TarjEdificio");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn161 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("TarjetaOficina");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn162 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("email");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn163 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idProyecto");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn164 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCuenta");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn165 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Bono");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn166 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Bono1");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn167 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Desc1");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn168 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Desc2");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn169 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("TieneQuincena");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn170 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Fotos");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn171 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Convenio");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn172 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idSubProyecto");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn173 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Sindicato");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn174 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Asociacion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn175 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idRolHorario");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn176 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("FechaContrato");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn177 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Educacion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn178 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("NumeroDias");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn179 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idRelacionT");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn180 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("gSalud");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn181 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("gEducacion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn182 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("gVivienda");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn183 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("gVestimenta");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn184 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("gAlimentacion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn185 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("gTotal");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn186 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("BaseComision");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn187 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("PorcentajeComision");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn188 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("gDesahucio");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn189 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("gDiscapacidad");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn190 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("IngGravOE");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn191 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("DeduccionOE");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn192 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("OtrasRebOE");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn193 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("ImpRentaAsumido");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn194 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("gRebajaIIIEdad");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn195 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("RentaRetenidaOE");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn196 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("RentaRetOE");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn197 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("TipoSangre");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn198 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("RegEntrada");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn199 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("RegSalida");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn200 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idTipoContrato");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn201 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("CodVendedor");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn202 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Clave");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn203 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idSupervisor");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn204 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("bSupervisor");
			this.idBusca = new System.Windows.Forms.TextBox();
			this.cdvRolPrestamo = new C1.Data.C1DataView();
			this.cdsRolPrestamo = new C1.Data.C1DataSet();
			this.ultraGrid1 = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.txtCodigo = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.rdgOpcion = new Infragistics.Win.UltraWinEditors.UltraOptionSet();
			this.rdgTipoBusqueda = new Infragistics.Win.UltraWinEditors.UltraOptionSet();
			this.btBuscar = new Infragistics.Win.Misc.UltraButton();
			this.btAceptar = new Infragistics.Win.Misc.UltraButton();
			this.btCancelar = new Infragistics.Win.Misc.UltraButton();
			this.cdvPersonal = new C1.Data.C1DataView();
			this.cdsPersonal = new C1.Data.C1DataSet();
			this.errorProvider = new System.Windows.Forms.ErrorProvider();
			this.cmbNombre = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.cmbFechaEntrega = new Infragistics.Win.UltraWinSchedule.UltraCalendarCombo();
			this.cmbFechaCobro = new Infragistics.Win.UltraWinSchedule.UltraCalendarCombo();
			this.ultraDropDown1 = new Infragistics.Win.UltraWinGrid.UltraDropDown();
			this.txtNumero = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			((System.ComponentModel.ISupportInitialize)(this.cdvRolPrestamo)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsRolPrestamo)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraGrid1)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtCodigo)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.rdgOpcion)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.rdgTipoBusqueda)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdvPersonal)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsPersonal)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbNombre)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbFechaEntrega)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbFechaCobro)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDropDown1)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNumero)).BeginInit();
			this.SuspendLayout();
			// 
			// idBusca
			// 
			this.idBusca.DataBindings.Add(new System.Windows.Forms.Binding("Text", this.cdvRolPrestamo, "idRolPrestamo"));
			this.idBusca.Location = new System.Drawing.Point(480, 17);
			this.idBusca.Name = "idBusca";
			this.idBusca.Size = new System.Drawing.Size(48, 20);
			this.idBusca.TabIndex = 23;
			this.idBusca.Text = "";
			// 
			// cdvRolPrestamo
			// 
			this.cdvRolPrestamo.BindingContextCtrl = this;
			this.cdvRolPrestamo.DataSet = this.cdsRolPrestamo;
			this.cdvRolPrestamo.TableName = "";
			this.cdvRolPrestamo.TableViewName = "RolPrestamo";
			// 
			// cdsRolPrestamo
			// 
			this.cdsRolPrestamo.BindingContextCtrl = this;
			this.cdsRolPrestamo.DataLibrary = "LibPersonal";
			this.cdsRolPrestamo.DataLibraryUrl = "";
			this.cdsRolPrestamo.DataSetDef = "dsRolGrupo";
			this.cdsRolPrestamo.FillOnRequest = false;
			this.cdsRolPrestamo.Locale = new System.Globalization.CultureInfo("es-EC");
			this.cdsRolPrestamo.Name = "cdsRolPrestamo";
			this.cdsRolPrestamo.SchemaClassName = "LibPersonal.DataClass";
			this.cdsRolPrestamo.SchemaDef = null;
			this.cdsRolPrestamo.BeforeFill += new C1.Data.FillEventHandler(this.cdsRolAnticipo_BeforeFill);
			// 
			// ultraGrid1
			// 
			this.ultraGrid1.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.ultraGrid1.Cursor = System.Windows.Forms.Cursors.Default;
			this.ultraGrid1.DataSource = this.cdvRolPrestamo;
			appearance1.BackColor = System.Drawing.Color.White;
			this.ultraGrid1.DisplayLayout.Appearance = appearance1;
			ultraGridBand1.AddButtonCaption = "RolAnticipo";
			ultraGridColumn1.Header.VisiblePosition = 0;
			ultraGridColumn1.Hidden = true;
			ultraGridColumn2.Header.VisiblePosition = 12;
			ultraGridColumn2.Hidden = true;
			ultraGridColumn3.Header.VisiblePosition = 2;
			ultraGridColumn4.Header.VisiblePosition = 3;
			ultraGridColumn5.Header.VisiblePosition = 4;
			ultraGridColumn6.Header.VisiblePosition = 1;
			ultraGridColumn6.Hidden = true;
			ultraGridColumn6.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.DropDownList;
			ultraGridColumn7.Header.VisiblePosition = 5;
			ultraGridColumn8.Header.VisiblePosition = 6;
			ultraGridColumn9.Header.VisiblePosition = 10;
			ultraGridColumn9.Hidden = true;
			ultraGridColumn10.Header.VisiblePosition = 9;
			ultraGridColumn10.Hidden = true;
			ultraGridColumn11.Header.VisiblePosition = 7;
			ultraGridColumn12.Header.VisiblePosition = 11;
			ultraGridColumn12.Hidden = true;
			ultraGridColumn13.Header.VisiblePosition = 8;
			ultraGridColumn13.Hidden = true;
			ultraGridColumn14.Header.VisiblePosition = 13;
			ultraGridColumn14.Hidden = true;
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
																										 ultraGridColumn14});
			this.ultraGrid1.DisplayLayout.BandsSerializer.Add(ultraGridBand1);
			this.ultraGrid1.DisplayLayout.GroupByBox.Prompt = "Arrastre columna a Agrupar";
			this.ultraGrid1.DisplayLayout.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.No;
			this.ultraGrid1.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.False;
			this.ultraGrid1.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.False;
			appearance2.BackColor = System.Drawing.Color.Transparent;
			this.ultraGrid1.DisplayLayout.Override.CardAreaAppearance = appearance2;
			this.ultraGrid1.DisplayLayout.Override.CellClickAction = Infragistics.Win.UltraWinGrid.CellClickAction.RowSelect;
			this.ultraGrid1.DisplayLayout.Override.ColumnAutoSizeMode = Infragistics.Win.UltraWinGrid.ColumnAutoSizeMode.AllRowsInBand;
			appearance3.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance3.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance3.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance3.FontData.BoldAsString = "True";
			appearance3.FontData.Name = "Arial";
			appearance3.FontData.SizeInPoints = 10F;
			appearance3.ForeColor = System.Drawing.Color.White;
			appearance3.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.ultraGrid1.DisplayLayout.Override.HeaderAppearance = appearance3;
			this.ultraGrid1.DisplayLayout.Override.HeaderClickAction = Infragistics.Win.UltraWinGrid.HeaderClickAction.SortMulti;
			appearance4.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance4.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance4.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ultraGrid1.DisplayLayout.Override.RowSelectorAppearance = appearance4;
			appearance5.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance5.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance5.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ultraGrid1.DisplayLayout.Override.SelectedRowAppearance = appearance5;
			this.ultraGrid1.DisplayLayout.Override.SelectTypeCell = Infragistics.Win.UltraWinGrid.SelectType.None;
			this.ultraGrid1.DisplayLayout.Override.SelectTypeCol = Infragistics.Win.UltraWinGrid.SelectType.None;
			this.ultraGrid1.DisplayLayout.Override.SelectTypeRow = Infragistics.Win.UltraWinGrid.SelectType.Single;
			this.ultraGrid1.DisplayLayout.ScrollBounds = Infragistics.Win.UltraWinGrid.ScrollBounds.ScrollToFill;
			this.ultraGrid1.DisplayLayout.ScrollStyle = Infragistics.Win.UltraWinGrid.ScrollStyle.Immediate;
			this.ultraGrid1.DisplayLayout.TabNavigation = Infragistics.Win.UltraWinGrid.TabNavigation.NextControl;
			this.ultraGrid1.DisplayLayout.ViewStyle = Infragistics.Win.UltraWinGrid.ViewStyle.SingleBand;
			this.ultraGrid1.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.ultraGrid1.Location = new System.Drawing.Point(16, 160);
			this.ultraGrid1.Name = "ultraGrid1";
			this.ultraGrid1.Size = new System.Drawing.Size(512, 264);
			this.ultraGrid1.TabIndex = 27;
			// 
			// txtCodigo
			// 
			this.txtCodigo.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtCodigo.Location = new System.Drawing.Point(136, 43);
			this.txtCodigo.Name = "txtCodigo";
			this.txtCodigo.Size = new System.Drawing.Size(128, 22);
			this.txtCodigo.TabIndex = 38;
			// 
			// rdgOpcion
			// 
			this.rdgOpcion.ItemAppearance = appearance6;
			this.rdgOpcion.ItemOrigin = new System.Drawing.Point(6, 0);
			valueListItem1.DataValue = 0;
			valueListItem1.DisplayText = "Nombre";
			valueListItem2.DataValue = 1;
			valueListItem2.DisplayText = "Concepto";
			valueListItem3.DataValue = 2;
			valueListItem3.DisplayText = "Valor";
			valueListItem4.DataValue = 3;
			valueListItem4.DisplayText = "Fecha Entrega";
			valueListItem5.DataValue = 4;
			valueListItem5.DisplayText = "Fecha Cobro";
			this.rdgOpcion.Items.Add(valueListItem1);
			this.rdgOpcion.Items.Add(valueListItem2);
			this.rdgOpcion.Items.Add(valueListItem3);
			this.rdgOpcion.Items.Add(valueListItem4);
			this.rdgOpcion.Items.Add(valueListItem5);
			this.rdgOpcion.ItemSpacingVertical = 9;
			this.rdgOpcion.Location = new System.Drawing.Point(16, 17);
			this.rdgOpcion.Name = "rdgOpcion";
			this.rdgOpcion.Size = new System.Drawing.Size(112, 129);
			this.rdgOpcion.TabIndex = 40;
			this.rdgOpcion.ValueChanged += new System.EventHandler(this.rdgOpcion_SelectedIndexChanged);
			// 
			// rdgTipoBusqueda
			// 
			this.rdgTipoBusqueda.CheckedIndex = 2;
			this.rdgTipoBusqueda.ItemAppearance = appearance7;
			this.rdgTipoBusqueda.ItemOrigin = new System.Drawing.Point(10, 0);
			valueListItem6.DataValue = 0;
			valueListItem6.DisplayText = "Que Empiece";
			valueListItem7.DataValue = 1;
			valueListItem7.DisplayText = "Exacta";
			valueListItem8.DataValue = 2;
			valueListItem8.DisplayText = "En cualquier parte";
			this.rdgTipoBusqueda.Items.Add(valueListItem6);
			this.rdgTipoBusqueda.Items.Add(valueListItem7);
			this.rdgTipoBusqueda.Items.Add(valueListItem8);
			this.rdgTipoBusqueda.ItemSpacingVertical = 10;
			this.rdgTipoBusqueda.Location = new System.Drawing.Point(296, 52);
			this.rdgTipoBusqueda.Name = "rdgTipoBusqueda";
			this.rdgTipoBusqueda.Size = new System.Drawing.Size(136, 86);
			this.rdgTipoBusqueda.TabIndex = 41;
			this.rdgTipoBusqueda.Text = "En cualquier parte";
			// 
			// btBuscar
			// 
			this.btBuscar.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			this.btBuscar.Location = new System.Drawing.Point(448, 52);
			this.btBuscar.Name = "btBuscar";
			this.btBuscar.Size = new System.Drawing.Size(75, 24);
			this.btBuscar.TabIndex = 42;
			this.btBuscar.Text = "&Buscar";
			this.btBuscar.Click += new System.EventHandler(this.btBuscar_Click);
			// 
			// btAceptar
			// 
			this.btAceptar.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			this.btAceptar.Location = new System.Drawing.Point(448, 86);
			this.btAceptar.Name = "btAceptar";
			this.btAceptar.Size = new System.Drawing.Size(75, 25);
			this.btAceptar.TabIndex = 43;
			this.btAceptar.Text = "&Aceptar";
			this.btAceptar.Click += new System.EventHandler(this.btAceptar_Click);
			// 
			// btCancelar
			// 
			this.btCancelar.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			this.btCancelar.DialogResult = System.Windows.Forms.DialogResult.Cancel;
			this.btCancelar.Location = new System.Drawing.Point(448, 121);
			this.btCancelar.Name = "btCancelar";
			this.btCancelar.Size = new System.Drawing.Size(75, 24);
			this.btCancelar.TabIndex = 44;
			this.btCancelar.Text = "&Cancelar";
			// 
			// cdvPersonal
			// 
			this.cdvPersonal.BindingContextCtrl = this;
			this.cdvPersonal.DataSet = this.cdsPersonal;
			this.cdvPersonal.Sort = "Nombre";
			this.cdvPersonal.TableName = "";
			this.cdvPersonal.TableViewName = "Personal";
			// 
			// cdsPersonal
			// 
			this.cdsPersonal.BindingContextCtrl = this;
			this.cdsPersonal.DataLibrary = "LibPersonal";
			this.cdsPersonal.DataLibraryUrl = "";
			this.cdsPersonal.DataSetDef = "dsPersonal";
			this.cdsPersonal.Locale = new System.Globalization.CultureInfo("es-EC");
			this.cdsPersonal.Name = "cdsPersonal";
			this.cdsPersonal.SchemaClassName = "LibPersonal.DataClass";
			this.cdsPersonal.SchemaDef = null;
			this.cdsPersonal.BeforeFill += new C1.Data.FillEventHandler(this.cdsPersonal_BeforeFill);
			// 
			// errorProvider
			// 
			this.errorProvider.ContainerControl = this;
			// 
			// cmbNombre
			// 
			this.cmbNombre.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbNombre.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbNombre.DataSource = this.cdvPersonal;
			this.cmbNombre.DisplayLayout.AutoFitColumns = true;
			ultraGridColumn15.Header.VisiblePosition = 21;
			ultraGridColumn15.Hidden = true;
			ultraGridColumn16.Header.VisiblePosition = 59;
			ultraGridColumn16.Hidden = true;
			ultraGridColumn17.Header.VisiblePosition = 19;
			ultraGridColumn17.Hidden = true;
			ultraGridColumn18.Header.VisiblePosition = 20;
			ultraGridColumn18.Hidden = true;
			ultraGridColumn19.Header.VisiblePosition = 18;
			ultraGridColumn19.Hidden = true;
			ultraGridColumn20.Header.VisiblePosition = 22;
			ultraGridColumn20.Hidden = true;
			ultraGridColumn21.Header.VisiblePosition = 37;
			ultraGridColumn21.Hidden = true;
			ultraGridColumn22.Header.VisiblePosition = 25;
			ultraGridColumn22.Hidden = true;
			ultraGridColumn23.Header.VisiblePosition = 0;
			ultraGridColumn23.Width = 253;
			ultraGridColumn24.Header.VisiblePosition = 1;
			ultraGridColumn24.Hidden = true;
			ultraGridColumn25.Header.VisiblePosition = 35;
			ultraGridColumn25.Hidden = true;
			ultraGridColumn26.Header.VisiblePosition = 24;
			ultraGridColumn26.Hidden = true;
			ultraGridColumn27.Header.VisiblePosition = 13;
			ultraGridColumn27.Hidden = true;
			ultraGridColumn28.Header.VisiblePosition = 32;
			ultraGridColumn28.Hidden = true;
			ultraGridColumn29.Header.VisiblePosition = 7;
			ultraGridColumn29.Hidden = true;
			ultraGridColumn30.Header.VisiblePosition = 6;
			ultraGridColumn30.Hidden = true;
			ultraGridColumn31.Header.VisiblePosition = 3;
			ultraGridColumn31.Hidden = true;
			ultraGridColumn32.Header.VisiblePosition = 29;
			ultraGridColumn32.Hidden = true;
			ultraGridColumn33.Header.VisiblePosition = 28;
			ultraGridColumn33.Hidden = true;
			ultraGridColumn34.Header.VisiblePosition = 23;
			ultraGridColumn34.Hidden = true;
			ultraGridColumn35.Header.VisiblePosition = 31;
			ultraGridColumn35.Hidden = true;
			ultraGridColumn36.Header.VisiblePosition = 10;
			ultraGridColumn36.Hidden = true;
			ultraGridColumn37.Header.VisiblePosition = 33;
			ultraGridColumn37.Hidden = true;
			ultraGridColumn38.Header.VisiblePosition = 14;
			ultraGridColumn38.Hidden = true;
			ultraGridColumn39.Header.VisiblePosition = 15;
			ultraGridColumn39.Hidden = true;
			ultraGridColumn40.Header.VisiblePosition = 17;
			ultraGridColumn40.Hidden = true;
			ultraGridColumn41.Header.VisiblePosition = 26;
			ultraGridColumn41.Hidden = true;
			ultraGridColumn42.Header.VisiblePosition = 5;
			ultraGridColumn42.Hidden = true;
			ultraGridColumn43.Header.VisiblePosition = 36;
			ultraGridColumn43.Hidden = true;
			ultraGridColumn44.Header.VisiblePosition = 38;
			ultraGridColumn44.Hidden = true;
			ultraGridColumn45.Header.VisiblePosition = 27;
			ultraGridColumn45.Hidden = true;
			ultraGridColumn46.Header.VisiblePosition = 30;
			ultraGridColumn46.Hidden = true;
			ultraGridColumn47.Header.VisiblePosition = 12;
			ultraGridColumn47.Hidden = true;
			ultraGridColumn48.Header.VisiblePosition = 9;
			ultraGridColumn48.Hidden = true;
			ultraGridColumn49.Header.VisiblePosition = 11;
			ultraGridColumn49.Hidden = true;
			ultraGridColumn50.Header.VisiblePosition = 16;
			ultraGridColumn50.Hidden = true;
			ultraGridColumn51.Header.VisiblePosition = 2;
			ultraGridColumn51.Hidden = true;
			ultraGridColumn52.Header.VisiblePosition = 34;
			ultraGridColumn52.Hidden = true;
			ultraGridColumn53.Header.VisiblePosition = 4;
			ultraGridColumn53.Hidden = true;
			ultraGridColumn54.Header.VisiblePosition = 8;
			ultraGridColumn54.Hidden = true;
			ultraGridColumn55.Header.VisiblePosition = 39;
			ultraGridColumn55.Hidden = true;
			ultraGridColumn56.Header.VisiblePosition = 40;
			ultraGridColumn56.Hidden = true;
			ultraGridColumn57.Header.VisiblePosition = 41;
			ultraGridColumn57.Hidden = true;
			ultraGridColumn58.Header.VisiblePosition = 42;
			ultraGridColumn58.Hidden = true;
			ultraGridColumn59.Header.VisiblePosition = 43;
			ultraGridColumn59.Hidden = true;
			ultraGridColumn60.Header.VisiblePosition = 44;
			ultraGridColumn60.Hidden = true;
			ultraGridColumn61.Header.VisiblePosition = 45;
			ultraGridColumn61.Hidden = true;
			ultraGridColumn62.Header.VisiblePosition = 46;
			ultraGridColumn62.Hidden = true;
			ultraGridColumn63.Header.VisiblePosition = 47;
			ultraGridColumn63.Hidden = true;
			ultraGridColumn64.Header.VisiblePosition = 48;
			ultraGridColumn64.Hidden = true;
			ultraGridColumn65.Header.VisiblePosition = 49;
			ultraGridColumn65.Hidden = true;
			ultraGridColumn66.Header.VisiblePosition = 50;
			ultraGridColumn66.Hidden = true;
			ultraGridColumn67.Header.VisiblePosition = 51;
			ultraGridColumn67.Hidden = true;
			ultraGridColumn68.Header.VisiblePosition = 52;
			ultraGridColumn68.Hidden = true;
			ultraGridColumn69.Header.VisiblePosition = 53;
			ultraGridColumn69.Hidden = true;
			ultraGridColumn70.Header.VisiblePosition = 54;
			ultraGridColumn70.Hidden = true;
			ultraGridColumn71.Header.VisiblePosition = 55;
			ultraGridColumn71.Hidden = true;
			ultraGridColumn72.Header.VisiblePosition = 56;
			ultraGridColumn72.Hidden = true;
			ultraGridColumn73.Header.VisiblePosition = 57;
			ultraGridColumn73.Hidden = true;
			ultraGridColumn74.Header.VisiblePosition = 58;
			ultraGridColumn74.Hidden = true;
			ultraGridColumn75.Header.VisiblePosition = 60;
			ultraGridColumn75.Hidden = true;
			ultraGridColumn76.Header.VisiblePosition = 61;
			ultraGridColumn76.Hidden = true;
			ultraGridColumn77.Header.VisiblePosition = 62;
			ultraGridColumn77.Hidden = true;
			ultraGridColumn78.Header.VisiblePosition = 63;
			ultraGridColumn78.Hidden = true;
			ultraGridColumn79.Header.VisiblePosition = 64;
			ultraGridColumn79.Hidden = true;
			ultraGridColumn80.Header.VisiblePosition = 65;
			ultraGridColumn80.Hidden = true;
			ultraGridColumn80.Width = 8;
			ultraGridColumn81.Header.VisiblePosition = 66;
			ultraGridColumn81.Hidden = true;
			ultraGridColumn81.Width = 17;
			ultraGridColumn82.Header.VisiblePosition = 67;
			ultraGridColumn82.Hidden = true;
			ultraGridColumn82.Width = 8;
			ultraGridColumn83.Header.VisiblePosition = 68;
			ultraGridColumn83.Hidden = true;
			ultraGridColumn83.Width = 8;
			ultraGridColumn84.Header.VisiblePosition = 69;
			ultraGridColumn84.Hidden = true;
			ultraGridColumn84.Width = 8;
			ultraGridColumn85.Header.VisiblePosition = 70;
			ultraGridColumn85.Hidden = true;
			ultraGridColumn85.Width = 8;
			ultraGridColumn86.Header.VisiblePosition = 71;
			ultraGridColumn86.Hidden = true;
			ultraGridColumn86.Width = 8;
			ultraGridColumn87.Header.VisiblePosition = 72;
			ultraGridColumn87.Hidden = true;
			ultraGridColumn87.Width = 8;
			ultraGridColumn88.Header.VisiblePosition = 73;
			ultraGridColumn88.Hidden = true;
			ultraGridColumn88.Width = 8;
			ultraGridColumn89.Header.VisiblePosition = 74;
			ultraGridColumn89.Hidden = true;
			ultraGridColumn89.Width = 8;
			ultraGridColumn90.Header.VisiblePosition = 75;
			ultraGridColumn90.Hidden = true;
			ultraGridColumn90.Width = 8;
			ultraGridColumn91.Header.VisiblePosition = 76;
			ultraGridColumn91.Hidden = true;
			ultraGridColumn91.Width = 8;
			ultraGridColumn92.Header.VisiblePosition = 77;
			ultraGridColumn92.Hidden = true;
			ultraGridColumn92.Width = 8;
			ultraGridColumn93.Header.VisiblePosition = 78;
			ultraGridColumn93.Hidden = true;
			ultraGridColumn93.Width = 8;
			ultraGridColumn94.Header.VisiblePosition = 79;
			ultraGridColumn94.Hidden = true;
			ultraGridColumn94.Width = 8;
			ultraGridColumn95.Header.VisiblePosition = 80;
			ultraGridColumn95.Hidden = true;
			ultraGridColumn95.Width = 8;
			ultraGridColumn96.Header.VisiblePosition = 81;
			ultraGridColumn96.Hidden = true;
			ultraGridColumn96.Width = 8;
			ultraGridColumn97.Header.VisiblePosition = 82;
			ultraGridColumn97.Hidden = true;
			ultraGridColumn97.Width = 8;
			ultraGridColumn98.Header.VisiblePosition = 83;
			ultraGridColumn98.Hidden = true;
			ultraGridColumn98.Width = 8;
			ultraGridColumn99.Header.VisiblePosition = 84;
			ultraGridColumn99.Hidden = true;
			ultraGridColumn99.Width = 8;
			ultraGridColumn100.Header.VisiblePosition = 85;
			ultraGridColumn100.Hidden = true;
			ultraGridColumn100.Width = 8;
			ultraGridColumn101.Header.VisiblePosition = 86;
			ultraGridColumn101.Hidden = true;
			ultraGridColumn101.Width = 8;
			ultraGridColumn102.Header.VisiblePosition = 87;
			ultraGridColumn102.Hidden = true;
			ultraGridColumn102.Width = 8;
			ultraGridColumn103.Header.VisiblePosition = 88;
			ultraGridColumn103.Hidden = true;
			ultraGridColumn103.Width = 8;
			ultraGridColumn104.Header.VisiblePosition = 89;
			ultraGridColumn104.Hidden = true;
			ultraGridColumn104.Width = 8;
			ultraGridColumn105.Header.VisiblePosition = 90;
			ultraGridColumn105.Hidden = true;
			ultraGridColumn105.Width = 8;
			ultraGridColumn106.Header.VisiblePosition = 91;
			ultraGridColumn106.Hidden = true;
			ultraGridColumn106.Width = 8;
			ultraGridColumn107.Header.VisiblePosition = 92;
			ultraGridColumn107.Hidden = true;
			ultraGridColumn107.Width = 8;
			ultraGridColumn108.Header.VisiblePosition = 93;
			ultraGridColumn108.Hidden = true;
			ultraGridColumn108.Width = 8;
			ultraGridColumn109.Header.VisiblePosition = 94;
			ultraGridColumn109.Hidden = true;
			ultraGridColumn109.Width = 17;
			ultraGridBand2.Columns.AddRange(new object[] {
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
																										 ultraGridColumn45,
																										 ultraGridColumn46,
																										 ultraGridColumn47,
																										 ultraGridColumn48,
																										 ultraGridColumn49,
																										 ultraGridColumn50,
																										 ultraGridColumn51,
																										 ultraGridColumn52,
																										 ultraGridColumn53,
																										 ultraGridColumn54,
																										 ultraGridColumn55,
																										 ultraGridColumn56,
																										 ultraGridColumn57,
																										 ultraGridColumn58,
																										 ultraGridColumn59,
																										 ultraGridColumn60,
																										 ultraGridColumn61,
																										 ultraGridColumn62,
																										 ultraGridColumn63,
																										 ultraGridColumn64,
																										 ultraGridColumn65,
																										 ultraGridColumn66,
																										 ultraGridColumn67,
																										 ultraGridColumn68,
																										 ultraGridColumn69,
																										 ultraGridColumn70,
																										 ultraGridColumn71,
																										 ultraGridColumn72,
																										 ultraGridColumn73,
																										 ultraGridColumn74,
																										 ultraGridColumn75,
																										 ultraGridColumn76,
																										 ultraGridColumn77,
																										 ultraGridColumn78,
																										 ultraGridColumn79,
																										 ultraGridColumn80,
																										 ultraGridColumn81,
																										 ultraGridColumn82,
																										 ultraGridColumn83,
																										 ultraGridColumn84,
																										 ultraGridColumn85,
																										 ultraGridColumn86,
																										 ultraGridColumn87,
																										 ultraGridColumn88,
																										 ultraGridColumn89,
																										 ultraGridColumn90,
																										 ultraGridColumn91,
																										 ultraGridColumn92,
																										 ultraGridColumn93,
																										 ultraGridColumn94,
																										 ultraGridColumn95,
																										 ultraGridColumn96,
																										 ultraGridColumn97,
																										 ultraGridColumn98,
																										 ultraGridColumn99,
																										 ultraGridColumn100,
																										 ultraGridColumn101,
																										 ultraGridColumn102,
																										 ultraGridColumn103,
																										 ultraGridColumn104,
																										 ultraGridColumn105,
																										 ultraGridColumn106,
																										 ultraGridColumn107,
																										 ultraGridColumn108,
																										 ultraGridColumn109});
			this.cmbNombre.DisplayLayout.BandsSerializer.Add(ultraGridBand2);
			this.cmbNombre.DisplayMember = "Nombre";
			this.cmbNombre.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbNombre.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbNombre.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbNombre.Location = new System.Drawing.Point(136, 17);
			this.cmbNombre.Name = "cmbNombre";
			this.cmbNombre.Size = new System.Drawing.Size(272, 21);
			this.cmbNombre.TabIndex = 45;
			this.cmbNombre.ValueMember = "idPersonal";
			// 
			// cmbFechaEntrega
			// 
			this.cmbFechaEntrega.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			dateButton1.Caption = "Today";
			this.cmbFechaEntrega.DateButtons.Add(dateButton1);
			this.cmbFechaEntrega.Format = "dd/MMM/yyyy";
			this.cmbFechaEntrega.Location = new System.Drawing.Point(136, 95);
			this.cmbFechaEntrega.Name = "cmbFechaEntrega";
			this.cmbFechaEntrega.NonAutoSizeHeight = 8;
			this.cmbFechaEntrega.Size = new System.Drawing.Size(128, 21);
			this.cmbFechaEntrega.SpinButtonIncrement = Infragistics.Win.UltraWinSchedule.SpinIncrementUnit.Months;
			this.cmbFechaEntrega.SpinButtonsVisible = true;
			this.cmbFechaEntrega.TabIndex = 46;
			this.cmbFechaEntrega.Value = new System.DateTime(2012, 4, 11, 0, 0, 0, 0);
			// 
			// cmbFechaCobro
			// 
			this.cmbFechaCobro.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			dateButton2.Caption = "Today";
			this.cmbFechaCobro.DateButtons.Add(dateButton2);
			this.cmbFechaCobro.Format = "dd/MMM/yyyy";
			this.cmbFechaCobro.Location = new System.Drawing.Point(136, 121);
			this.cmbFechaCobro.Name = "cmbFechaCobro";
			this.cmbFechaCobro.NonAutoSizeHeight = 8;
			this.cmbFechaCobro.Size = new System.Drawing.Size(128, 21);
			this.cmbFechaCobro.SpinButtonIncrement = Infragistics.Win.UltraWinSchedule.SpinIncrementUnit.Months;
			this.cmbFechaCobro.SpinButtonsVisible = true;
			this.cmbFechaCobro.TabIndex = 47;
			this.cmbFechaCobro.Value = new System.DateTime(2012, 4, 11, 0, 0, 0, 0);
			// 
			// ultraDropDown1
			// 
			this.ultraDropDown1.Cursor = System.Windows.Forms.Cursors.Default;
			this.ultraDropDown1.DataSource = this.cdvPersonal;
			ultraGridColumn110.Header.VisiblePosition = 21;
			ultraGridColumn110.Hidden = true;
			ultraGridColumn111.Header.VisiblePosition = 59;
			ultraGridColumn111.Hidden = true;
			ultraGridColumn112.Header.VisiblePosition = 19;
			ultraGridColumn112.Hidden = true;
			ultraGridColumn113.Header.VisiblePosition = 20;
			ultraGridColumn113.Hidden = true;
			ultraGridColumn114.Header.VisiblePosition = 18;
			ultraGridColumn114.Hidden = true;
			ultraGridColumn115.Header.VisiblePosition = 22;
			ultraGridColumn115.Hidden = true;
			ultraGridColumn116.Header.VisiblePosition = 37;
			ultraGridColumn116.Hidden = true;
			ultraGridColumn117.Header.VisiblePosition = 25;
			ultraGridColumn117.Hidden = true;
			ultraGridColumn118.Header.VisiblePosition = 1;
			ultraGridColumn119.Header.VisiblePosition = 0;
			ultraGridColumn119.Width = 143;
			ultraGridColumn120.Header.VisiblePosition = 35;
			ultraGridColumn120.Hidden = true;
			ultraGridColumn121.Header.VisiblePosition = 24;
			ultraGridColumn121.Hidden = true;
			ultraGridColumn122.Header.VisiblePosition = 13;
			ultraGridColumn122.Hidden = true;
			ultraGridColumn123.Header.VisiblePosition = 32;
			ultraGridColumn123.Hidden = true;
			ultraGridColumn124.Header.VisiblePosition = 7;
			ultraGridColumn124.Hidden = true;
			ultraGridColumn125.Header.VisiblePosition = 6;
			ultraGridColumn125.Hidden = true;
			ultraGridColumn126.Header.VisiblePosition = 3;
			ultraGridColumn126.Hidden = true;
			ultraGridColumn127.Header.VisiblePosition = 29;
			ultraGridColumn127.Hidden = true;
			ultraGridColumn128.Header.VisiblePosition = 28;
			ultraGridColumn128.Hidden = true;
			ultraGridColumn129.Header.VisiblePosition = 23;
			ultraGridColumn129.Hidden = true;
			ultraGridColumn130.Header.VisiblePosition = 31;
			ultraGridColumn130.Hidden = true;
			ultraGridColumn131.Header.VisiblePosition = 10;
			ultraGridColumn131.Hidden = true;
			ultraGridColumn132.Header.VisiblePosition = 33;
			ultraGridColumn132.Hidden = true;
			ultraGridColumn133.Header.VisiblePosition = 14;
			ultraGridColumn133.Hidden = true;
			ultraGridColumn134.Header.VisiblePosition = 15;
			ultraGridColumn134.Hidden = true;
			ultraGridColumn135.Header.VisiblePosition = 17;
			ultraGridColumn135.Hidden = true;
			ultraGridColumn136.Header.VisiblePosition = 26;
			ultraGridColumn136.Hidden = true;
			ultraGridColumn137.Header.VisiblePosition = 5;
			ultraGridColumn137.Hidden = true;
			ultraGridColumn138.Header.VisiblePosition = 36;
			ultraGridColumn138.Hidden = true;
			ultraGridColumn139.Header.VisiblePosition = 38;
			ultraGridColumn139.Hidden = true;
			ultraGridColumn140.Header.VisiblePosition = 27;
			ultraGridColumn140.Hidden = true;
			ultraGridColumn141.Header.VisiblePosition = 30;
			ultraGridColumn141.Hidden = true;
			ultraGridColumn142.Header.VisiblePosition = 12;
			ultraGridColumn142.Hidden = true;
			ultraGridColumn143.Header.VisiblePosition = 9;
			ultraGridColumn143.Hidden = true;
			ultraGridColumn144.Header.VisiblePosition = 11;
			ultraGridColumn144.Hidden = true;
			ultraGridColumn145.Header.VisiblePosition = 16;
			ultraGridColumn145.Hidden = true;
			ultraGridColumn146.Header.VisiblePosition = 2;
			ultraGridColumn146.Hidden = true;
			ultraGridColumn147.Header.VisiblePosition = 34;
			ultraGridColumn147.Hidden = true;
			ultraGridColumn148.Header.VisiblePosition = 4;
			ultraGridColumn148.Hidden = true;
			ultraGridColumn149.Header.VisiblePosition = 8;
			ultraGridColumn149.Hidden = true;
			ultraGridColumn150.Header.VisiblePosition = 39;
			ultraGridColumn150.Hidden = true;
			ultraGridColumn151.Header.VisiblePosition = 40;
			ultraGridColumn151.Hidden = true;
			ultraGridColumn152.Header.VisiblePosition = 41;
			ultraGridColumn152.Hidden = true;
			ultraGridColumn153.Header.VisiblePosition = 42;
			ultraGridColumn153.Hidden = true;
			ultraGridColumn154.Header.VisiblePosition = 43;
			ultraGridColumn154.Hidden = true;
			ultraGridColumn155.Header.VisiblePosition = 44;
			ultraGridColumn155.Hidden = true;
			ultraGridColumn156.Header.VisiblePosition = 45;
			ultraGridColumn156.Hidden = true;
			ultraGridColumn157.Header.VisiblePosition = 46;
			ultraGridColumn157.Hidden = true;
			ultraGridColumn158.Header.VisiblePosition = 47;
			ultraGridColumn158.Hidden = true;
			ultraGridColumn159.Header.VisiblePosition = 48;
			ultraGridColumn159.Hidden = true;
			ultraGridColumn160.Header.VisiblePosition = 49;
			ultraGridColumn160.Hidden = true;
			ultraGridColumn161.Header.VisiblePosition = 50;
			ultraGridColumn161.Hidden = true;
			ultraGridColumn162.Header.VisiblePosition = 51;
			ultraGridColumn162.Hidden = true;
			ultraGridColumn163.Header.VisiblePosition = 52;
			ultraGridColumn163.Hidden = true;
			ultraGridColumn164.Header.VisiblePosition = 53;
			ultraGridColumn164.Hidden = true;
			ultraGridColumn165.Header.VisiblePosition = 54;
			ultraGridColumn165.Hidden = true;
			ultraGridColumn166.Header.VisiblePosition = 55;
			ultraGridColumn166.Hidden = true;
			ultraGridColumn167.Header.VisiblePosition = 56;
			ultraGridColumn167.Hidden = true;
			ultraGridColumn168.Header.VisiblePosition = 57;
			ultraGridColumn168.Hidden = true;
			ultraGridColumn169.Header.VisiblePosition = 58;
			ultraGridColumn169.Hidden = true;
			ultraGridColumn170.Header.VisiblePosition = 60;
			ultraGridColumn171.Header.VisiblePosition = 61;
			ultraGridColumn172.Header.VisiblePosition = 62;
			ultraGridColumn173.Header.VisiblePosition = 63;
			ultraGridColumn174.Header.VisiblePosition = 64;
			ultraGridColumn175.Header.VisiblePosition = 65;
			ultraGridColumn176.Header.VisiblePosition = 66;
			ultraGridColumn177.Header.VisiblePosition = 67;
			ultraGridColumn178.Header.VisiblePosition = 68;
			ultraGridColumn179.Header.VisiblePosition = 69;
			ultraGridColumn180.Header.VisiblePosition = 70;
			ultraGridColumn181.Header.VisiblePosition = 71;
			ultraGridColumn182.Header.VisiblePosition = 72;
			ultraGridColumn183.Header.VisiblePosition = 73;
			ultraGridColumn184.Header.VisiblePosition = 74;
			ultraGridColumn185.Header.VisiblePosition = 75;
			ultraGridColumn186.Header.VisiblePosition = 76;
			ultraGridColumn187.Header.VisiblePosition = 77;
			ultraGridColumn188.Header.VisiblePosition = 78;
			ultraGridColumn189.Header.VisiblePosition = 79;
			ultraGridColumn190.Header.VisiblePosition = 80;
			ultraGridColumn191.Header.VisiblePosition = 81;
			ultraGridColumn192.Header.VisiblePosition = 82;
			ultraGridColumn193.Header.VisiblePosition = 83;
			ultraGridColumn194.Header.VisiblePosition = 84;
			ultraGridColumn195.Header.VisiblePosition = 85;
			ultraGridColumn196.Header.VisiblePosition = 86;
			ultraGridColumn197.Header.VisiblePosition = 87;
			ultraGridColumn198.Header.VisiblePosition = 88;
			ultraGridColumn199.Header.VisiblePosition = 89;
			ultraGridColumn200.Header.VisiblePosition = 90;
			ultraGridColumn201.Header.VisiblePosition = 91;
			ultraGridColumn202.Header.VisiblePosition = 92;
			ultraGridColumn203.Header.VisiblePosition = 93;
			ultraGridColumn204.Header.VisiblePosition = 94;
			ultraGridBand3.Columns.AddRange(new object[] {
																										 ultraGridColumn110,
																										 ultraGridColumn111,
																										 ultraGridColumn112,
																										 ultraGridColumn113,
																										 ultraGridColumn114,
																										 ultraGridColumn115,
																										 ultraGridColumn116,
																										 ultraGridColumn117,
																										 ultraGridColumn118,
																										 ultraGridColumn119,
																										 ultraGridColumn120,
																										 ultraGridColumn121,
																										 ultraGridColumn122,
																										 ultraGridColumn123,
																										 ultraGridColumn124,
																										 ultraGridColumn125,
																										 ultraGridColumn126,
																										 ultraGridColumn127,
																										 ultraGridColumn128,
																										 ultraGridColumn129,
																										 ultraGridColumn130,
																										 ultraGridColumn131,
																										 ultraGridColumn132,
																										 ultraGridColumn133,
																										 ultraGridColumn134,
																										 ultraGridColumn135,
																										 ultraGridColumn136,
																										 ultraGridColumn137,
																										 ultraGridColumn138,
																										 ultraGridColumn139,
																										 ultraGridColumn140,
																										 ultraGridColumn141,
																										 ultraGridColumn142,
																										 ultraGridColumn143,
																										 ultraGridColumn144,
																										 ultraGridColumn145,
																										 ultraGridColumn146,
																										 ultraGridColumn147,
																										 ultraGridColumn148,
																										 ultraGridColumn149,
																										 ultraGridColumn150,
																										 ultraGridColumn151,
																										 ultraGridColumn152,
																										 ultraGridColumn153,
																										 ultraGridColumn154,
																										 ultraGridColumn155,
																										 ultraGridColumn156,
																										 ultraGridColumn157,
																										 ultraGridColumn158,
																										 ultraGridColumn159,
																										 ultraGridColumn160,
																										 ultraGridColumn161,
																										 ultraGridColumn162,
																										 ultraGridColumn163,
																										 ultraGridColumn164,
																										 ultraGridColumn165,
																										 ultraGridColumn166,
																										 ultraGridColumn167,
																										 ultraGridColumn168,
																										 ultraGridColumn169,
																										 ultraGridColumn170,
																										 ultraGridColumn171,
																										 ultraGridColumn172,
																										 ultraGridColumn173,
																										 ultraGridColumn174,
																										 ultraGridColumn175,
																										 ultraGridColumn176,
																										 ultraGridColumn177,
																										 ultraGridColumn178,
																										 ultraGridColumn179,
																										 ultraGridColumn180,
																										 ultraGridColumn181,
																										 ultraGridColumn182,
																										 ultraGridColumn183,
																										 ultraGridColumn184,
																										 ultraGridColumn185,
																										 ultraGridColumn186,
																										 ultraGridColumn187,
																										 ultraGridColumn188,
																										 ultraGridColumn189,
																										 ultraGridColumn190,
																										 ultraGridColumn191,
																										 ultraGridColumn192,
																										 ultraGridColumn193,
																										 ultraGridColumn194,
																										 ultraGridColumn195,
																										 ultraGridColumn196,
																										 ultraGridColumn197,
																										 ultraGridColumn198,
																										 ultraGridColumn199,
																										 ultraGridColumn200,
																										 ultraGridColumn201,
																										 ultraGridColumn202,
																										 ultraGridColumn203,
																										 ultraGridColumn204});
			this.ultraDropDown1.DisplayLayout.BandsSerializer.Add(ultraGridBand3);
			this.ultraDropDown1.DisplayMember = "Apellido";
			this.ultraDropDown1.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.ultraDropDown1.Location = new System.Drawing.Point(272, 276);
			this.ultraDropDown1.Name = "ultraDropDown1";
			this.ultraDropDown1.Size = new System.Drawing.Size(240, 103);
			this.ultraDropDown1.TabIndex = 48;
			this.ultraDropDown1.ValueMember = "idPersonal";
			this.ultraDropDown1.Visible = false;
			// 
			// txtNumero
			// 
			this.txtNumero.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtNumero.FormatString = "#,##0.00";
			this.txtNumero.Location = new System.Drawing.Point(136, 69);
			this.txtNumero.Name = "txtNumero";
			this.txtNumero.NullText = "0";
			this.txtNumero.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtNumero.PromptChar = ' ';
			this.txtNumero.Size = new System.Drawing.Size(128, 22);
			this.txtNumero.TabIndex = 49;
			// 
			// BuscaPrestamo
			// 
			this.AcceptButton = this.btBuscar;
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 14);
			this.CancelButton = this.btCancelar;
			this.ClientSize = new System.Drawing.Size(544, 445);
			this.Controls.Add(this.txtNumero);
			this.Controls.Add(this.ultraDropDown1);
			this.Controls.Add(this.cmbFechaCobro);
			this.Controls.Add(this.cmbFechaEntrega);
			this.Controls.Add(this.cmbNombre);
			this.Controls.Add(this.btCancelar);
			this.Controls.Add(this.btAceptar);
			this.Controls.Add(this.btBuscar);
			this.Controls.Add(this.rdgTipoBusqueda);
			this.Controls.Add(this.rdgOpcion);
			this.Controls.Add(this.txtCodigo);
			this.Controls.Add(this.ultraGrid1);
			this.Controls.Add(this.idBusca);
			this.MinimizeBox = false;
			this.Name = "BuscaPrestamo";
			this.Text = "Búsqueda de Prestamos";
			this.Resize += new System.EventHandler(this.BuscaArticulo_Resize);
			this.Load += new System.EventHandler(this.BuscaAsiento_Load);
			((System.ComponentModel.ISupportInitialize)(this.cdvRolPrestamo)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsRolPrestamo)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraGrid1)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtCodigo)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.rdgOpcion)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.rdgTipoBusqueda)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdvPersonal)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsPersonal)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbNombre)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbFechaEntrega)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbFechaCobro)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDropDown1)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNumero)).EndInit();
			this.ResumeLayout(false);

		}
		#endregion

		private void BuscaAsiento_Load(object sender, System.EventArgs e)
		{
			Cursor = Cursors.Default;
			idBusca.Width =0;
			this.rdgOpcion.Value = 0;
		}

		CultureInfo us = new CultureInfo("en-US");

		private bool Verifica()
		{
			this.errorProvider.SetError(this.cmbNombre, "");
			this.errorProvider.SetError(this.txtCodigo, "");
			this.errorProvider.SetError(this.txtNumero, "");
			this.errorProvider.SetError(this.cmbFechaEntrega, "");
			this.errorProvider.SetError(this.cmbFechaCobro, "");

			bool bOk = true;
			int iTipo = (int) this.rdgOpcion.Value;

			if (iTipo == 0)
				if (this.cmbNombre.Text.Trim().Length == 0)
				{
					this.errorProvider.SetError(this.cmbNombre, "Ingrese Nombre");
					bOk = false;
				}

			if (iTipo == 1)
				if (this.txtCodigo.Text.Trim().Length == 0)
				{
					this.errorProvider.SetError(this.txtCodigo, "Ingrese Codigo");
					bOk = false;
				}
			if (iTipo == 2)
				if (this.txtNumero.Text.Trim().Length == 0)
				{
					this.errorProvider.SetError(this.txtNumero, "Ingrese Valor");
					bOk = false;
				}
			if (iTipo == 3)
				if (this.cmbFechaEntrega.Value == null)
				{
					this.errorProvider.SetError(this.cmbFechaEntrega, "Ingrese Fecha de Entregar");
					bOk = false;
				}
			if (iTipo == 4)
				if (this.cmbFechaCobro.Value == null)
				{
					this.errorProvider.SetError(this.cmbFechaCobro, "Ingrese SubGrupo");
					bOk = false;
				}
			return bOk;
		}

		private void btBuscar_Click(object sender, System.EventArgs e)
		{
			if (!Verifica()) return;
			string strFiltro;
			string strAntes = "";
			string strDespues="'";

//		Opciones de comodines
			if ((int) this.rdgTipoBusqueda.Value == 0) strDespues = "%'";
			if ((int) this.rdgTipoBusqueda.Value == 1) strAntes = "";
			if ((int) this.rdgTipoBusqueda.Value == 2) strDespues = "%'";
			if ((int) this.rdgTipoBusqueda.Value == 2) strAntes = "%";

			strFiltro = "";
			string stBusca = "";
			switch((int) this.rdgOpcion.Value)
			{
				case 0:
					strFiltro = "idPersonal  = " +
						this.cmbNombre.Value.ToString().Trim();
					break;
				case 1:
					if (txtCodigo.Value == null)
					{
						errorProvider.SetError(txtCodigo, "Ingrese Codigo");
						return;
					}
					errorProvider.SetError(txtCodigo, "");

					stBusca = this.txtCodigo.Value.ToString().Trim();
					stBusca = stBusca.Replace("'", "''");

					strFiltro = "Concepto LIKE '" + strAntes + stBusca + strDespues;
					break;
				case 2:
					double dNumero = (double) this.txtNumero.Value;
					strFiltro = "Capital = " + dNumero.ToString("0.00", us).Trim();
					break;
				case 3:
					DateTime dtFechaEntrega = (DateTime) this.cmbFechaEntrega.Value;
					strFiltro = "FechaEntrega = '" + dtFechaEntrega.ToString("yyyyMMdd").Trim() + "'";
					break;
				case 4:
					DateTime dtFechaCobro = (DateTime) this.cmbFechaCobro.Value;
					strFiltro = "FechaCobro = '" + dtFechaCobro.ToString("yyyyMMdd").Trim() + "'";
					break;
			}

			C1.Data.FilterCondition fcPadre = new 
				C1.Data.FilterCondition(this.cdsRolPrestamo, "RolPrestamo", strFiltro);
			C1.Data.FilterConditions fcTotal = new C1.Data.FilterConditions();
			fcTotal.Add(fcPadre);

			this.cdsRolPrestamo.Fill(fcTotal, false);

			if (this.cdvRolPrestamo.Count == 0)
			{
				MessageBox.Show("No existen registros que cumplan la condicion de busqueda", "Información");
				return;
			}
			if (this.cdvRolPrestamo.Count == 1)
			{
				this.ultraGrid1.Select();
				this.DialogResult = DialogResult.OK;
			}
			else
			{
				this.AcceptButton=this.btAceptar;
				this.ultraGrid1.Select();
			}
		}

//		public int IdBuscado()
//		{
//			int iFila = this.gridView1.FocusedRowHandle;
//			return int.Parse(this.cdvSocio[iFila]["idSocio"].ToString());
//		}

		private void btAceptar_Click(object sender, System.EventArgs e)
		{
			this.DialogResult = DialogResult.OK;
		}

		private void BuscaArticulo_Resize(object sender, System.EventArgs e)
		{
			this.ultraGrid1.DisplayLayout.AutoFitColumns = true;
		}

		private void rdgOpcion_SelectedIndexChanged(object sender, System.EventArgs e)
		{
			this.cmbNombre.Enabled = false;
			this.txtCodigo.Enabled = false;
			this.txtNumero.Enabled = false;
			this.cmbFechaCobro.Enabled = false;
			this.cmbFechaEntrega.Enabled = false;

			switch((int) this.rdgOpcion.Value)
			{
				case 0:
					this.cmbNombre.Enabled = true;
					this.cmbNombre.Select();
					break;
				case 1:
					this.txtCodigo.Enabled = true;
					this.txtCodigo.Select();
					break;
				case 2:
					this.txtNumero.Enabled = true;
					this.txtNumero.Select();
					break;
				case 3:
					this.cmbFechaEntrega.Enabled = true;
					this.cmbFechaEntrega.Select();
					break;
				case 4:
					this.cmbFechaCobro.Enabled = true;
					this.cmbFechaCobro.Select();
					break;
			}

		}

		private void cdsRolAnticipo_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			this.cdsRolPrestamo.Schema.Connections[0].ConnectionString = 
				miFuncion.AccesoBase(MenuLatinium.stDirRoles);
		}

		private void cdsPersonal_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			this.cdsPersonal.Schema.Connections[0].ConnectionString = 
				miFuncion.AccesoBase(MenuLatinium.stDirRoles);
		}


	
	}
}
