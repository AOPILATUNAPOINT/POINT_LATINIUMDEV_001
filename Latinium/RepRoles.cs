using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;

namespace Latinium
{
	/// <summary>
	/// Descripción breve de RepRoles.
	/// </summary>
	public class RepRoles : DevExpress.XtraEditors.XtraForm
	{
		private C1.Data.C1DataSet cdsPersonal;
		private Infragistics.Win.UltraWinEditors.UltraCheckEditor chkPersonal;
		private Infragistics.Win.UltraWinEditors.UltraOptionSet optTipo;
		private C1.Data.C1DataSet cdsDepartamento;
		private Infragistics.Win.UltraWinEditors.UltraCheckEditor chkDepartamento;
		private Infragistics.Win.UltraWinSchedule.UltraCalendarCombo cmbFecha;
		private System.Windows.Forms.Label label1;
		private Infragistics.Win.Misc.UltraButton btAceptar;
		private Infragistics.Win.Misc.UltraButton btCancelar;
		private C1.Data.C1DataSet cdsReporte;
		private C1.Data.C1DataView cdvReporte;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbReporte;
		private System.Windows.Forms.Label label2;
		private System.Windows.Forms.ErrorProvider errorProvider;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbPersonal;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbDepartamento;
		private Infragistics.Win.UltraWinEditors.UltraCheckEditor chkSucursal;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbSucursal;
		private Infragistics.Win.UltraWinEditors.UltraCheckEditor chkCargo;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbCargo;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbGrupo;
		private C1.Data.C1DataSet cdsRolGrupo;
		private Infragistics.Win.UltraWinEditors.UltraCheckEditor chkGrupo;
		private System.Windows.Forms.Label lblHasta;
		private Infragistics.Win.UltraWinSchedule.UltraCalendarCombo cmbHasta;
		private Infragistics.Win.UltraWinEditors.UltraCheckEditor chkResumen;
		private C1.Data.C1DataView cdvPersonal;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbBanco;
		private C1.Data.C1DataSet cdsBanco;
		private Infragistics.Win.UltraWinEditors.UltraCheckEditor chkBanco;
		private C1.Data.C1DataView cdvSucursal;
		private Infragistics.Win.UltraWinEditors.UltraCheckEditor chkDespedidos;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbCentroCosto;
		private Infragistics.Win.UltraWinEditors.UltraCheckEditor chkCentroCosto;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbDistrib;
		private Infragistics.Win.UltraWinEditors.UltraCheckEditor chkDistribucion;
		/// <summary>
		/// Variable del diseñador requerida.
		/// </summary>
		private System.ComponentModel.Container components = null;

		public RepRoles()
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
			Infragistics.Win.ValueListItem valueListItem1 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.ValueListItem valueListItem2 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.ValueListItem valueListItem3 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.ValueListItem valueListItem4 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.ValueListItem valueListItem5 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand1 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Personal", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn1 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idPersonal");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn2 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idAsiDistGrupo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn3 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCentroCosto");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn4 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idDepartamento");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn5 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCargo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn6 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idRolSucursal");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn7 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idRolGrupo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn8 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Jornal");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn9 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn10 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Apellido");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn11 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Codigo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn12 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Jornada");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn13 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Direccion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn14 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Telefono");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn15 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Celular");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn16 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Cedula");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn17 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Afiliacion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn18 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Reserva");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn19 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Renta");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn20 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Iess");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn21 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Sueldo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn22 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Comision");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn23 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Tipo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn24 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("FechaIngreso");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn25 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("FechaNacimiento");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn26 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("FechaSalida");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn27 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Notas");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn28 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("CargasFamiliares");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn29 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idProvincia");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn30 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCiudad");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn31 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Numero");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn32 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("SistemaSalario");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn33 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Cuenta");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn34 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("CodigoSectorial");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn35 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Condicion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn36 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("FechaNovedad");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn37 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Acceso");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn38 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Vendedor");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn39 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Borrar");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn40 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Cobrador");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn41 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Sexo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn42 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Estado_Civil");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn43 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Banco");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn44 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("CuentaCorr");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn45 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idTipoCuenta");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn46 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("CedulaMilitar");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn47 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idFormaPago");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn48 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Quincena");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn49 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("CodIngreso");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn50 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("CodTelefono");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn51 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("TarjEdificio");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn52 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("TarjetaOficina");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn53 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("email");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn54 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idProyecto");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn55 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCuenta");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn56 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Bono");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn57 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Bono1");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn58 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Desc1");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn59 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Desc2");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn60 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("TieneQuincena");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn61 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Fotos");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn62 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Convenio");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn63 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idSubProyecto");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn64 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Sindicato");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn65 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Asociacion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn66 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idRolHorario");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn67 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("FechaContrato");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn68 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Educacion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn69 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("NumeroDias");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn70 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idRelacionT");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn71 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("gSalud");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn72 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("gEducacion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn73 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("gVivienda");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn74 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("gVestimenta");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn75 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("gAlimentacion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn76 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("gTotal");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn77 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("BaseComision");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn78 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("PorcentajeComision");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn79 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("gDesahucio");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn80 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("gDiscapacidad");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn81 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("IngGravOE");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn82 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("DeduccionOE");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn83 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("OtrasRebOE");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn84 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("ImpRentaAsumido");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn85 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("gRebajaIIIEdad");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn86 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("RentaRetenidaOE");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn87 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("RentaRetOE");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn88 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("TipoSangre");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn89 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("RegEntrada");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn90 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("RegSalida");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn91 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idTipoContrato");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn92 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("CodVendedor");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn93 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Clave");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn94 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idSupervisor");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn95 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("bSupervisor");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn96 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("E_Profesor");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn97 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Dias360");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn98 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("DiasXIII");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn99 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idJefe");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn100 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("PoseeDiscapacidad");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn101 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("PorcDiscapacidad");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn102 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("TipoDiscapacidad");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn103 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("NumConadis");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn104 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("TipoVivienda");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn105 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Vehiculo");
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand2 = new Infragistics.Win.UltraWinGrid.UltraGridBand("RolDepartamento", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn106 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idDepartamento");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn107 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Descripcion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn108 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Borrar");
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand3 = new Infragistics.Win.UltraWinGrid.UltraGridBand("RolSucursal", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn109 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idRolSucursal");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn110 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Sucursal");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn111 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCodigo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn112 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("OrigenFondos");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn113 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("CodigoSeguro");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn114 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idRolTipoEmpleador");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn115 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idRelacionTrabajo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn116 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Numero");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn117 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Codigo");
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand4 = new Infragistics.Win.UltraWinGrid.UltraGridBand("RolCargo", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn118 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCargo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn119 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Descripcion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn120 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Codigo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn121 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Borrar");
			Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton dateButton1 = new Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand5 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Reporte", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn122 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Detalle");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn123 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idGrupo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn124 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idReporte");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn125 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Reporte");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn126 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Tipo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn127 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Filtro");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn128 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("NoFiltro");
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand6 = new Infragistics.Win.UltraWinGrid.UltraGridBand("RolGrupo", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn129 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idRolGrupo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn130 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Grupo");
			Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton dateButton2 = new Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand7 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Banco", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn131 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idBanco");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn132 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn133 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Pais");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn134 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Codigo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn135 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Borrar");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn136 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Hoja");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn137 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("ProcesoOk");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn138 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idPais");
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand8 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Banco", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn139 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idProyecto");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn140 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand9 = new Infragistics.Win.UltraWinGrid.UltraGridBand("AsientoDistGrupo", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn141 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idAsiDistGrupo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn142 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Distribucion");
			System.Resources.ResourceManager resources = new System.Resources.ResourceManager(typeof(RepRoles));
			this.optTipo = new Infragistics.Win.UltraWinEditors.UltraOptionSet();
			this.cdsPersonal = new C1.Data.C1DataSet();
			this.cdsDepartamento = new C1.Data.C1DataSet();
			this.cmbPersonal = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.cdvPersonal = new C1.Data.C1DataView();
			this.chkPersonal = new Infragistics.Win.UltraWinEditors.UltraCheckEditor();
			this.chkDepartamento = new Infragistics.Win.UltraWinEditors.UltraCheckEditor();
			this.cmbDepartamento = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.chkSucursal = new Infragistics.Win.UltraWinEditors.UltraCheckEditor();
			this.cmbSucursal = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.cdvSucursal = new C1.Data.C1DataView();
			this.chkCargo = new Infragistics.Win.UltraWinEditors.UltraCheckEditor();
			this.cmbCargo = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.cmbFecha = new Infragistics.Win.UltraWinSchedule.UltraCalendarCombo();
			this.label1 = new System.Windows.Forms.Label();
			this.btAceptar = new Infragistics.Win.Misc.UltraButton();
			this.btCancelar = new Infragistics.Win.Misc.UltraButton();
			this.cdsReporte = new C1.Data.C1DataSet();
			this.cdvReporte = new C1.Data.C1DataView();
			this.cmbReporte = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.label2 = new System.Windows.Forms.Label();
			this.errorProvider = new System.Windows.Forms.ErrorProvider();
			this.cmbGrupo = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.cdsRolGrupo = new C1.Data.C1DataSet();
			this.chkGrupo = new Infragistics.Win.UltraWinEditors.UltraCheckEditor();
			this.lblHasta = new System.Windows.Forms.Label();
			this.cmbHasta = new Infragistics.Win.UltraWinSchedule.UltraCalendarCombo();
			this.chkResumen = new Infragistics.Win.UltraWinEditors.UltraCheckEditor();
			this.chkBanco = new Infragistics.Win.UltraWinEditors.UltraCheckEditor();
			this.cmbBanco = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.cdsBanco = new C1.Data.C1DataSet();
			this.chkDespedidos = new Infragistics.Win.UltraWinEditors.UltraCheckEditor();
			this.chkCentroCosto = new Infragistics.Win.UltraWinEditors.UltraCheckEditor();
			this.cmbCentroCosto = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.cmbDistrib = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.chkDistribucion = new Infragistics.Win.UltraWinEditors.UltraCheckEditor();
			((System.ComponentModel.ISupportInitialize)(this.optTipo)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsPersonal)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsDepartamento)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbPersonal)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdvPersonal)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbDepartamento)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbSucursal)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdvSucursal)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbCargo)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbFecha)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsReporte)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdvReporte)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbReporte)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbGrupo)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsRolGrupo)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbHasta)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbBanco)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsBanco)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbCentroCosto)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbDistrib)).BeginInit();
			this.SuspendLayout();
			// 
			// optTipo
			// 
			this.optTipo.ItemAppearance = appearance1;
			this.optTipo.ItemOrigin = new System.Drawing.Point(10, 5);
			valueListItem1.DataValue = 0;
			valueListItem1.DisplayText = "General";
			valueListItem2.DataValue = 1;
			valueListItem2.DisplayText = "Individual";
			valueListItem3.DataValue = 2;
			valueListItem3.DisplayText = "Planilla IESS";
			valueListItem4.DataValue = 3;
			valueListItem4.DisplayText = "Provisiones";
			valueListItem5.DataValue = 7;
			valueListItem5.DisplayText = "Quincena";
			this.optTipo.Items.Add(valueListItem1);
			this.optTipo.Items.Add(valueListItem2);
			this.optTipo.Items.Add(valueListItem3);
			this.optTipo.Items.Add(valueListItem4);
			this.optTipo.Items.Add(valueListItem5);
			this.optTipo.ItemSpacingVertical = 8;
			this.optTipo.Location = new System.Drawing.Point(24, 52);
			this.optTipo.Name = "optTipo";
			this.optTipo.Size = new System.Drawing.Size(160, 146);
			this.optTipo.TabIndex = 0;
			this.optTipo.ValueChanged += new System.EventHandler(this.optTipo_ValueChanged);
			// 
			// cdsPersonal
			// 
			this.cdsPersonal.BindingContextCtrl = this;
			this.cdsPersonal.CaseSensitive = false;
			this.cdsPersonal.DataLibrary = "LibPersonal";
			this.cdsPersonal.DataLibraryUrl = "";
			this.cdsPersonal.DataSetDef = "dsPersonal";
			this.cdsPersonal.Locale = new System.Globalization.CultureInfo("es-EC");
			this.cdsPersonal.Name = "cdsPersonal";
			this.cdsPersonal.SchemaClassName = "LibPersonal.DataClass";
			this.cdsPersonal.SchemaDef = null;
			this.cdsPersonal.BeforeFill += new C1.Data.FillEventHandler(this.cdsPersonal_BeforeFill);
			// 
			// cdsDepartamento
			// 
			this.cdsDepartamento.BindingContextCtrl = this;
			this.cdsDepartamento.DataLibrary = "LibPersonal";
			this.cdsDepartamento.DataLibraryUrl = "";
			this.cdsDepartamento.DataSetDef = "dsRolTablas";
			this.cdsDepartamento.Locale = new System.Globalization.CultureInfo("es-EC");
			this.cdsDepartamento.Name = "cdsDepartamento";
			this.cdsDepartamento.SchemaClassName = "LibPersonal.DataClass";
			this.cdsDepartamento.SchemaDef = null;
			this.cdsDepartamento.BeforeFill += new C1.Data.FillEventHandler(this.cdsDepartamento_BeforeFill);
			// 
			// cmbPersonal
			// 
			this.cmbPersonal.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbPersonal.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbPersonal.DataSource = this.cdvPersonal;
			ultraGridColumn1.Header.VisiblePosition = 22;
			ultraGridColumn1.Hidden = true;
			ultraGridColumn2.Header.VisiblePosition = 59;
			ultraGridColumn2.Hidden = true;
			ultraGridColumn2.Width = 30;
			ultraGridColumn3.Header.VisiblePosition = 20;
			ultraGridColumn3.Hidden = true;
			ultraGridColumn4.Header.VisiblePosition = 21;
			ultraGridColumn4.Hidden = true;
			ultraGridColumn5.Header.VisiblePosition = 19;
			ultraGridColumn5.Hidden = true;
			ultraGridColumn6.Header.VisiblePosition = 23;
			ultraGridColumn6.Hidden = true;
			ultraGridColumn7.Header.VisiblePosition = 36;
			ultraGridColumn7.Hidden = true;
			ultraGridColumn8.Header.VisiblePosition = 26;
			ultraGridColumn8.Hidden = true;
			ultraGridColumn9.Header.VisiblePosition = 0;
			ultraGridColumn9.Width = 296;
			ultraGridColumn10.Header.VisiblePosition = 5;
			ultraGridColumn10.Hidden = true;
			ultraGridColumn11.Header.VisiblePosition = 2;
			ultraGridColumn11.Width = 89;
			ultraGridColumn12.Header.VisiblePosition = 25;
			ultraGridColumn12.Hidden = true;
			ultraGridColumn13.Header.VisiblePosition = 14;
			ultraGridColumn13.Hidden = true;
			ultraGridColumn14.Header.VisiblePosition = 33;
			ultraGridColumn14.Hidden = true;
			ultraGridColumn15.Header.VisiblePosition = 1;
			ultraGridColumn15.Width = 101;
			ultraGridColumn16.Header.VisiblePosition = 8;
			ultraGridColumn16.Hidden = true;
			ultraGridColumn17.Header.VisiblePosition = 4;
			ultraGridColumn17.Hidden = true;
			ultraGridColumn18.Header.VisiblePosition = 30;
			ultraGridColumn18.Hidden = true;
			ultraGridColumn19.Header.VisiblePosition = 29;
			ultraGridColumn19.Hidden = true;
			ultraGridColumn20.Header.VisiblePosition = 24;
			ultraGridColumn20.Hidden = true;
			ultraGridColumn21.Header.VisiblePosition = 32;
			ultraGridColumn21.Hidden = true;
			ultraGridColumn22.Header.VisiblePosition = 11;
			ultraGridColumn22.Hidden = true;
			ultraGridColumn23.Header.VisiblePosition = 34;
			ultraGridColumn23.Hidden = true;
			ultraGridColumn24.Header.VisiblePosition = 15;
			ultraGridColumn24.Hidden = true;
			ultraGridColumn25.Header.VisiblePosition = 16;
			ultraGridColumn25.Hidden = true;
			ultraGridColumn26.Header.VisiblePosition = 18;
			ultraGridColumn26.Hidden = true;
			ultraGridColumn27.Header.VisiblePosition = 27;
			ultraGridColumn27.Hidden = true;
			ultraGridColumn28.Header.VisiblePosition = 7;
			ultraGridColumn28.Hidden = true;
			ultraGridColumn29.Header.VisiblePosition = 37;
			ultraGridColumn29.Hidden = true;
			ultraGridColumn30.Header.VisiblePosition = 38;
			ultraGridColumn30.Hidden = true;
			ultraGridColumn31.Header.VisiblePosition = 28;
			ultraGridColumn31.Hidden = true;
			ultraGridColumn32.Header.VisiblePosition = 31;
			ultraGridColumn32.Hidden = true;
			ultraGridColumn33.Header.VisiblePosition = 13;
			ultraGridColumn33.Hidden = true;
			ultraGridColumn34.Header.VisiblePosition = 10;
			ultraGridColumn34.Hidden = true;
			ultraGridColumn35.Header.VisiblePosition = 12;
			ultraGridColumn35.Hidden = true;
			ultraGridColumn36.Header.VisiblePosition = 17;
			ultraGridColumn36.Hidden = true;
			ultraGridColumn37.Header.VisiblePosition = 3;
			ultraGridColumn37.Hidden = true;
			ultraGridColumn38.Header.VisiblePosition = 35;
			ultraGridColumn38.Hidden = true;
			ultraGridColumn39.Header.VisiblePosition = 6;
			ultraGridColumn39.Hidden = true;
			ultraGridColumn40.Header.VisiblePosition = 9;
			ultraGridColumn40.Hidden = true;
			ultraGridColumn41.Header.VisiblePosition = 39;
			ultraGridColumn41.Hidden = true;
			ultraGridColumn41.Width = 8;
			ultraGridColumn42.Header.VisiblePosition = 40;
			ultraGridColumn42.Hidden = true;
			ultraGridColumn42.Width = 14;
			ultraGridColumn43.Header.VisiblePosition = 41;
			ultraGridColumn43.Hidden = true;
			ultraGridColumn43.Width = 14;
			ultraGridColumn44.Header.VisiblePosition = 42;
			ultraGridColumn44.Hidden = true;
			ultraGridColumn44.Width = 14;
			ultraGridColumn45.Header.VisiblePosition = 43;
			ultraGridColumn45.Hidden = true;
			ultraGridColumn45.Width = 11;
			ultraGridColumn46.Header.VisiblePosition = 44;
			ultraGridColumn46.Hidden = true;
			ultraGridColumn46.Width = 14;
			ultraGridColumn47.Header.VisiblePosition = 45;
			ultraGridColumn47.Hidden = true;
			ultraGridColumn47.Width = 11;
			ultraGridColumn48.Header.VisiblePosition = 46;
			ultraGridColumn48.Hidden = true;
			ultraGridColumn48.Width = 11;
			ultraGridColumn49.Header.VisiblePosition = 47;
			ultraGridColumn49.Hidden = true;
			ultraGridColumn49.Width = 14;
			ultraGridColumn50.Header.VisiblePosition = 48;
			ultraGridColumn50.Hidden = true;
			ultraGridColumn50.Width = 14;
			ultraGridColumn51.Header.VisiblePosition = 49;
			ultraGridColumn51.Hidden = true;
			ultraGridColumn51.Width = 14;
			ultraGridColumn52.Header.VisiblePosition = 50;
			ultraGridColumn52.Hidden = true;
			ultraGridColumn52.Width = 14;
			ultraGridColumn53.Header.VisiblePosition = 51;
			ultraGridColumn53.Hidden = true;
			ultraGridColumn53.Width = 17;
			ultraGridColumn54.Header.VisiblePosition = 52;
			ultraGridColumn54.Hidden = true;
			ultraGridColumn54.Width = 41;
			ultraGridColumn55.Header.VisiblePosition = 53;
			ultraGridColumn55.Hidden = true;
			ultraGridColumn55.Width = 38;
			ultraGridColumn56.Header.VisiblePosition = 54;
			ultraGridColumn56.Hidden = true;
			ultraGridColumn56.Width = 22;
			ultraGridColumn57.Header.VisiblePosition = 55;
			ultraGridColumn57.Hidden = true;
			ultraGridColumn57.Width = 22;
			ultraGridColumn58.Header.VisiblePosition = 56;
			ultraGridColumn58.Hidden = true;
			ultraGridColumn58.Width = 22;
			ultraGridColumn59.Header.VisiblePosition = 57;
			ultraGridColumn59.Hidden = true;
			ultraGridColumn59.Width = 22;
			ultraGridColumn60.Header.VisiblePosition = 58;
			ultraGridColumn60.Hidden = true;
			ultraGridColumn60.Width = 25;
			ultraGridColumn61.Header.VisiblePosition = 60;
			ultraGridColumn61.Hidden = true;
			ultraGridColumn61.Width = 27;
			ultraGridColumn62.Header.VisiblePosition = 61;
			ultraGridColumn62.Hidden = true;
			ultraGridColumn62.Width = 59;
			ultraGridColumn63.Header.VisiblePosition = 62;
			ultraGridColumn63.Hidden = true;
			ultraGridColumn63.Width = 53;
			ultraGridColumn64.Header.VisiblePosition = 63;
			ultraGridColumn64.Hidden = true;
			ultraGridColumn64.Width = 37;
			ultraGridColumn65.Header.VisiblePosition = 64;
			ultraGridColumn65.Hidden = true;
			ultraGridColumn65.Width = 42;
			ultraGridColumn66.Header.VisiblePosition = 65;
			ultraGridColumn66.Hidden = true;
			ultraGridColumn66.Width = 191;
			ultraGridColumn67.Header.VisiblePosition = 66;
			ultraGridColumn67.Hidden = true;
			ultraGridColumn67.Width = 224;
			ultraGridColumn68.Header.VisiblePosition = 67;
			ultraGridColumn68.Hidden = true;
			ultraGridColumn68.Width = 32;
			ultraGridColumn69.Header.VisiblePosition = 68;
			ultraGridColumn69.Hidden = true;
			ultraGridColumn69.Width = 36;
			ultraGridColumn70.Header.VisiblePosition = 69;
			ultraGridColumn70.Hidden = true;
			ultraGridColumn70.Width = 37;
			ultraGridColumn71.Header.VisiblePosition = 70;
			ultraGridColumn71.Hidden = true;
			ultraGridColumn72.Header.VisiblePosition = 71;
			ultraGridColumn72.Hidden = true;
			ultraGridColumn73.Header.VisiblePosition = 72;
			ultraGridColumn73.Hidden = true;
			ultraGridColumn74.Header.VisiblePosition = 73;
			ultraGridColumn74.Hidden = true;
			ultraGridColumn75.Header.VisiblePosition = 74;
			ultraGridColumn75.Hidden = true;
			ultraGridColumn76.Header.VisiblePosition = 75;
			ultraGridColumn76.Hidden = true;
			ultraGridColumn77.Header.VisiblePosition = 76;
			ultraGridColumn77.Hidden = true;
			ultraGridColumn78.Header.VisiblePosition = 77;
			ultraGridColumn78.Hidden = true;
			ultraGridColumn79.Header.VisiblePosition = 78;
			ultraGridColumn79.Hidden = true;
			ultraGridColumn80.Header.VisiblePosition = 79;
			ultraGridColumn80.Hidden = true;
			ultraGridColumn81.Header.VisiblePosition = 80;
			ultraGridColumn81.Hidden = true;
			ultraGridColumn82.Header.VisiblePosition = 81;
			ultraGridColumn82.Hidden = true;
			ultraGridColumn83.Header.VisiblePosition = 82;
			ultraGridColumn83.Hidden = true;
			ultraGridColumn84.Header.VisiblePosition = 83;
			ultraGridColumn84.Hidden = true;
			ultraGridColumn85.Header.VisiblePosition = 84;
			ultraGridColumn85.Hidden = true;
			ultraGridColumn86.Header.VisiblePosition = 85;
			ultraGridColumn86.Hidden = true;
			ultraGridColumn87.Header.VisiblePosition = 86;
			ultraGridColumn87.Hidden = true;
			ultraGridColumn88.Header.VisiblePosition = 87;
			ultraGridColumn88.Hidden = true;
			ultraGridColumn89.Header.VisiblePosition = 88;
			ultraGridColumn89.Hidden = true;
			ultraGridColumn90.Header.VisiblePosition = 89;
			ultraGridColumn90.Hidden = true;
			ultraGridColumn91.Header.VisiblePosition = 90;
			ultraGridColumn91.Hidden = true;
			ultraGridColumn92.Header.VisiblePosition = 91;
			ultraGridColumn93.Header.VisiblePosition = 92;
			ultraGridColumn94.Header.VisiblePosition = 93;
			ultraGridColumn95.Header.VisiblePosition = 94;
			ultraGridColumn96.Header.VisiblePosition = 95;
			ultraGridColumn97.Header.VisiblePosition = 96;
			ultraGridColumn98.Header.VisiblePosition = 97;
			ultraGridColumn99.Header.VisiblePosition = 98;
			ultraGridColumn100.Header.VisiblePosition = 99;
			ultraGridColumn101.Header.VisiblePosition = 100;
			ultraGridColumn102.Header.VisiblePosition = 101;
			ultraGridColumn103.Header.VisiblePosition = 102;
			ultraGridColumn104.Header.VisiblePosition = 103;
			ultraGridColumn105.Header.VisiblePosition = 104;
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
																										 ultraGridColumn105});
			this.cmbPersonal.DisplayLayout.BandsSerializer.Add(ultraGridBand1);
			this.cmbPersonal.DisplayMember = "Nombre";
			this.cmbPersonal.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbPersonal.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbPersonal.Enabled = false;
			this.cmbPersonal.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbPersonal.Location = new System.Drawing.Point(344, 24);
			this.cmbPersonal.Name = "cmbPersonal";
			this.cmbPersonal.Size = new System.Drawing.Size(216, 21);
			this.cmbPersonal.TabIndex = 1;
			this.cmbPersonal.ValueMember = "idPersonal";
			this.cmbPersonal.InitializeLayout += new Infragistics.Win.UltraWinGrid.InitializeLayoutEventHandler(this.cmbPersonal_InitializeLayout);
			// 
			// cdvPersonal
			// 
			this.cdvPersonal.BindingContextCtrl = this;
			this.cdvPersonal.DataSet = this.cdsPersonal;
			this.cdvPersonal.Sort = "nombre";
			this.cdvPersonal.TableName = "";
			this.cdvPersonal.TableViewName = "Personal";
			// 
			// chkPersonal
			// 
			this.chkPersonal.Checked = true;
			this.chkPersonal.CheckState = System.Windows.Forms.CheckState.Checked;
			this.chkPersonal.Location = new System.Drawing.Point(208, 24);
			this.chkPersonal.Name = "chkPersonal";
			this.chkPersonal.Size = new System.Drawing.Size(104, 21);
			this.chkPersonal.TabIndex = 2;
			this.chkPersonal.Text = "Todo Personal";
			this.chkPersonal.CheckedChanged += new System.EventHandler(this.chkPersonal_CheckedChanged);
			// 
			// chkDepartamento
			// 
			this.chkDepartamento.Checked = true;
			this.chkDepartamento.CheckState = System.Windows.Forms.CheckState.Checked;
			this.chkDepartamento.Location = new System.Drawing.Point(208, 56);
			this.chkDepartamento.Name = "chkDepartamento";
			this.chkDepartamento.Size = new System.Drawing.Size(128, 22);
			this.chkDepartamento.TabIndex = 4;
			this.chkDepartamento.Text = "Todo Departamento";
			this.chkDepartamento.CheckedChanged += new System.EventHandler(this.chkDepartamento_CheckedChanged);
			// 
			// cmbDepartamento
			// 
			this.cmbDepartamento.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbDepartamento.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbDepartamento.DataMember = "RolDepartamento";
			this.cmbDepartamento.DataSource = this.cdsDepartamento;
			this.cmbDepartamento.DisplayLayout.AutoFitColumns = true;
			ultraGridBand2.AddButtonCaption = "Personal";
			ultraGridColumn106.Header.VisiblePosition = 2;
			ultraGridColumn106.Hidden = true;
			ultraGridColumn107.Header.VisiblePosition = 0;
			ultraGridColumn107.Width = 197;
			ultraGridColumn108.Header.VisiblePosition = 1;
			ultraGridColumn108.Hidden = true;
			ultraGridBand2.Columns.AddRange(new object[] {
																										 ultraGridColumn106,
																										 ultraGridColumn107,
																										 ultraGridColumn108});
			this.cmbDepartamento.DisplayLayout.BandsSerializer.Add(ultraGridBand2);
			this.cmbDepartamento.DisplayMember = "Descripcion";
			this.cmbDepartamento.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbDepartamento.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbDepartamento.Enabled = false;
			this.cmbDepartamento.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbDepartamento.Location = new System.Drawing.Point(344, 56);
			this.cmbDepartamento.Name = "cmbDepartamento";
			this.cmbDepartamento.Size = new System.Drawing.Size(216, 21);
			this.cmbDepartamento.TabIndex = 3;
			this.cmbDepartamento.ValueMember = "idDepartamento";
			// 
			// chkSucursal
			// 
			this.chkSucursal.Checked = true;
			this.chkSucursal.CheckState = System.Windows.Forms.CheckState.Checked;
			this.chkSucursal.Location = new System.Drawing.Point(208, 120);
			this.chkSucursal.Name = "chkSucursal";
			this.chkSucursal.Size = new System.Drawing.Size(104, 22);
			this.chkSucursal.TabIndex = 6;
			this.chkSucursal.Text = "Toda Sucursal";
			this.chkSucursal.CheckedChanged += new System.EventHandler(this.chkSucursal_CheckedChanged);
			// 
			// cmbSucursal
			// 
			this.cmbSucursal.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbSucursal.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbSucursal.DataSource = this.cdvSucursal;
			this.cmbSucursal.DisplayLayout.AutoFitColumns = true;
			ultraGridBand3.AddButtonCaption = "Personal";
			ultraGridColumn109.Header.VisiblePosition = 6;
			ultraGridColumn109.Hidden = true;
			ultraGridColumn110.Header.VisiblePosition = 0;
			ultraGridColumn110.Width = 130;
			ultraGridColumn111.Header.VisiblePosition = 2;
			ultraGridColumn111.Hidden = true;
			ultraGridColumn112.Header.VisiblePosition = 5;
			ultraGridColumn112.Hidden = true;
			ultraGridColumn113.Header.VisiblePosition = 1;
			ultraGridColumn113.Hidden = true;
			ultraGridColumn114.Header.VisiblePosition = 4;
			ultraGridColumn114.Hidden = true;
			ultraGridColumn115.Header.VisiblePosition = 3;
			ultraGridColumn115.Hidden = true;
			ultraGridColumn116.Header.VisiblePosition = 7;
			ultraGridColumn116.Hidden = true;
			ultraGridColumn117.Header.VisiblePosition = 8;
			ultraGridColumn117.Width = 67;
			ultraGridBand3.Columns.AddRange(new object[] {
																										 ultraGridColumn109,
																										 ultraGridColumn110,
																										 ultraGridColumn111,
																										 ultraGridColumn112,
																										 ultraGridColumn113,
																										 ultraGridColumn114,
																										 ultraGridColumn115,
																										 ultraGridColumn116,
																										 ultraGridColumn117});
			this.cmbSucursal.DisplayLayout.BandsSerializer.Add(ultraGridBand3);
			this.cmbSucursal.DisplayMember = "Sucursal";
			this.cmbSucursal.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbSucursal.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbSucursal.Enabled = false;
			this.cmbSucursal.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbSucursal.Location = new System.Drawing.Point(344, 120);
			this.cmbSucursal.Name = "cmbSucursal";
			this.cmbSucursal.Size = new System.Drawing.Size(216, 21);
			this.cmbSucursal.TabIndex = 5;
			this.cmbSucursal.ValueMember = "idRolSucursal";
			// 
			// cdvSucursal
			// 
			this.cdvSucursal.BindingContextCtrl = this;
			this.cdvSucursal.DataSet = this.cdsDepartamento;
			this.cdvSucursal.Sort = "Sucursal";
			this.cdvSucursal.TableName = "";
			this.cdvSucursal.TableViewName = "RolSucursal";
			// 
			// chkCargo
			// 
			this.chkCargo.Checked = true;
			this.chkCargo.CheckState = System.Windows.Forms.CheckState.Checked;
			this.chkCargo.Location = new System.Drawing.Point(208, 88);
			this.chkCargo.Name = "chkCargo";
			this.chkCargo.Size = new System.Drawing.Size(104, 21);
			this.chkCargo.TabIndex = 8;
			this.chkCargo.Text = "Todo Cargo";
			this.chkCargo.CheckedChanged += new System.EventHandler(this.chkCargo_CheckedChanged);
			// 
			// cmbCargo
			// 
			this.cmbCargo.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbCargo.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbCargo.DataMember = "RolCargo";
			this.cmbCargo.DataSource = this.cdsDepartamento;
			this.cmbCargo.DisplayLayout.AutoFitColumns = true;
			ultraGridBand4.AddButtonCaption = "Personal";
			ultraGridColumn118.Header.VisiblePosition = 2;
			ultraGridColumn118.Hidden = true;
			ultraGridColumn119.Header.VisiblePosition = 0;
			ultraGridColumn119.Width = 197;
			ultraGridColumn120.Header.VisiblePosition = 3;
			ultraGridColumn120.Hidden = true;
			ultraGridColumn121.Header.VisiblePosition = 1;
			ultraGridColumn121.Hidden = true;
			ultraGridBand4.Columns.AddRange(new object[] {
																										 ultraGridColumn118,
																										 ultraGridColumn119,
																										 ultraGridColumn120,
																										 ultraGridColumn121});
			this.cmbCargo.DisplayLayout.BandsSerializer.Add(ultraGridBand4);
			this.cmbCargo.DisplayMember = "Descripcion";
			this.cmbCargo.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbCargo.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbCargo.Enabled = false;
			this.cmbCargo.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbCargo.Location = new System.Drawing.Point(344, 88);
			this.cmbCargo.Name = "cmbCargo";
			this.cmbCargo.Size = new System.Drawing.Size(216, 21);
			this.cmbCargo.TabIndex = 7;
			this.cmbCargo.ValueMember = "idCargo";
			// 
			// cmbFecha
			// 
			dateButton1.Caption = "Today";
			this.cmbFecha.DateButtons.Add(dateButton1);
			this.cmbFecha.Format = "MMM/yyyy";
			this.cmbFecha.Location = new System.Drawing.Point(72, 233);
			this.cmbFecha.Name = "cmbFecha";
			this.cmbFecha.NonAutoSizeHeight = 23;
			this.cmbFecha.Size = new System.Drawing.Size(120, 21);
			this.cmbFecha.SpinButtonIncrement = Infragistics.Win.UltraWinSchedule.SpinIncrementUnit.Months;
			this.cmbFecha.SpinButtonsVisible = true;
			this.cmbFecha.TabIndex = 9;
			this.cmbFecha.Value = new System.DateTime(2009, 1, 17, 0, 0, 0, 0);
			this.cmbFecha.BeforeDropDown += new System.ComponentModel.CancelEventHandler(this.cmbFecha_BeforeDropDown);
			// 
			// label1
			// 
			this.label1.AutoSize = true;
			this.label1.Location = new System.Drawing.Point(24, 241);
			this.label1.Name = "label1";
			this.label1.Size = new System.Drawing.Size(24, 16);
			this.label1.TabIndex = 10;
			this.label1.Text = "Mes";
			// 
			// btAceptar
			// 
			this.btAceptar.Location = new System.Drawing.Point(240, 336);
			this.btAceptar.Name = "btAceptar";
			this.btAceptar.Size = new System.Drawing.Size(75, 24);
			this.btAceptar.TabIndex = 11;
			this.btAceptar.Text = "&Aceptar";
			this.btAceptar.Click += new System.EventHandler(this.btAceptar_Click);
			// 
			// btCancelar
			// 
			this.btCancelar.DialogResult = System.Windows.Forms.DialogResult.Cancel;
			this.btCancelar.Location = new System.Drawing.Point(456, 336);
			this.btCancelar.Name = "btCancelar";
			this.btCancelar.Size = new System.Drawing.Size(75, 24);
			this.btCancelar.TabIndex = 12;
			this.btCancelar.Text = "&Cancelar";
			this.btCancelar.Click += new System.EventHandler(this.btCancelar_Click);
			// 
			// cdsReporte
			// 
			this.cdsReporte.BindingContextCtrl = this;
			this.cdsReporte.DataLibrary = "LibEmpresa";
			this.cdsReporte.DataLibraryUrl = "";
			this.cdsReporte.DataSetDef = "dsReporte";
			this.cdsReporte.Locale = new System.Globalization.CultureInfo("es-EC");
			this.cdsReporte.Name = "cdsReporte";
			this.cdsReporte.SchemaClassName = "LibEmpresa.DataClass";
			this.cdsReporte.SchemaDef = null;
			this.cdsReporte.BeforeFill += new C1.Data.FillEventHandler(this.cdsReporte_BeforeFill);
			// 
			// cdvReporte
			// 
			this.cdvReporte.BindingContextCtrl = this;
			this.cdvReporte.DataSet = this.cdsReporte;
			this.cdvReporte.RowFilter = "Tipo = 6";
			this.cdvReporte.TableName = "";
			this.cdvReporte.TableViewName = "Reporte";
			// 
			// cmbReporte
			// 
			this.cmbReporte.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbReporte.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbReporte.DataSource = this.cdvReporte;
			this.cmbReporte.DisplayLayout.AutoFitColumns = true;
			ultraGridColumn122.Header.VisiblePosition = 0;
			ultraGridColumn122.Width = 197;
			ultraGridColumn123.Header.VisiblePosition = 1;
			ultraGridColumn123.Hidden = true;
			ultraGridColumn124.Header.VisiblePosition = 2;
			ultraGridColumn124.Hidden = true;
			ultraGridColumn125.Header.VisiblePosition = 3;
			ultraGridColumn125.Hidden = true;
			ultraGridColumn126.Header.VisiblePosition = 4;
			ultraGridColumn126.Hidden = true;
			ultraGridColumn127.Header.VisiblePosition = 5;
			ultraGridColumn127.Hidden = true;
			ultraGridColumn127.Width = 25;
			ultraGridColumn128.Header.VisiblePosition = 6;
			ultraGridColumn128.Hidden = true;
			ultraGridColumn128.Width = 26;
			ultraGridBand5.Columns.AddRange(new object[] {
																										 ultraGridColumn122,
																										 ultraGridColumn123,
																										 ultraGridColumn124,
																										 ultraGridColumn125,
																										 ultraGridColumn126,
																										 ultraGridColumn127,
																										 ultraGridColumn128});
			this.cmbReporte.DisplayLayout.BandsSerializer.Add(ultraGridBand5);
			this.cmbReporte.DisplayMember = "Reporte";
			this.cmbReporte.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbReporte.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbReporte.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbReporte.Location = new System.Drawing.Point(344, 296);
			this.cmbReporte.Name = "cmbReporte";
			this.cmbReporte.Size = new System.Drawing.Size(216, 21);
			this.cmbReporte.TabIndex = 44;
			this.cmbReporte.ValueMember = "idReporte";
			// 
			// label2
			// 
			this.label2.AutoSize = true;
			this.label2.Location = new System.Drawing.Point(208, 296);
			this.label2.Name = "label2";
			this.label2.Size = new System.Drawing.Size(43, 16);
			this.label2.TabIndex = 45;
			this.label2.Text = "Reporte";
			// 
			// errorProvider
			// 
			this.errorProvider.ContainerControl = this;
			// 
			// cmbGrupo
			// 
			this.cmbGrupo.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbGrupo.Cursor = System.Windows.Forms.Cursors.WaitCursor;
			this.cmbGrupo.DataMember = "RolGrupo";
			this.cmbGrupo.DataSource = this.cdsRolGrupo;
			this.cmbGrupo.DisplayLayout.AutoFitColumns = true;
			ultraGridBand6.AddButtonCaption = "Reporte";
			ultraGridColumn129.Header.VisiblePosition = 1;
			ultraGridColumn129.Hidden = true;
			ultraGridColumn129.Width = 84;
			ultraGridColumn130.Header.VisiblePosition = 0;
			ultraGridColumn130.Width = 197;
			ultraGridBand6.Columns.AddRange(new object[] {
																										 ultraGridColumn129,
																										 ultraGridColumn130});
			this.cmbGrupo.DisplayLayout.BandsSerializer.Add(ultraGridBand6);
			this.cmbGrupo.DisplayMember = "Grupo";
			this.cmbGrupo.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbGrupo.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbGrupo.Enabled = false;
			this.cmbGrupo.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbGrupo.Location = new System.Drawing.Point(344, 152);
			this.cmbGrupo.Name = "cmbGrupo";
			this.cmbGrupo.Size = new System.Drawing.Size(216, 21);
			this.cmbGrupo.TabIndex = 46;
			this.cmbGrupo.ValueMember = "idRolGrupo";
			// 
			// cdsRolGrupo
			// 
			this.cdsRolGrupo.BindingContextCtrl = this;
			this.cdsRolGrupo.DataLibrary = "LibPersonal";
			this.cdsRolGrupo.DataLibraryUrl = "";
			this.cdsRolGrupo.DataSetDef = "dsRolGrupo";
			this.cdsRolGrupo.Locale = new System.Globalization.CultureInfo("es-EC");
			this.cdsRolGrupo.Name = "cdsRolGrupo";
			this.cdsRolGrupo.SchemaClassName = "LibPersonal.DataClass";
			this.cdsRolGrupo.SchemaDef = null;
			this.cdsRolGrupo.BeforeFill += new C1.Data.FillEventHandler(this.cdsRolGrupo_BeforeFill);
			// 
			// chkGrupo
			// 
			this.chkGrupo.Checked = true;
			this.chkGrupo.CheckState = System.Windows.Forms.CheckState.Checked;
			this.chkGrupo.Location = new System.Drawing.Point(208, 152);
			this.chkGrupo.Name = "chkGrupo";
			this.chkGrupo.Size = new System.Drawing.Size(104, 21);
			this.chkGrupo.TabIndex = 47;
			this.chkGrupo.Text = "Todo Grupo";
			this.chkGrupo.CheckedChanged += new System.EventHandler(this.chkGrupo_CheckedChanged);
			// 
			// lblHasta
			// 
			this.lblHasta.AutoSize = true;
			this.lblHasta.Location = new System.Drawing.Point(24, 267);
			this.lblHasta.Name = "lblHasta";
			this.lblHasta.Size = new System.Drawing.Size(33, 16);
			this.lblHasta.TabIndex = 49;
			this.lblHasta.Text = "Hasta";
			this.lblHasta.Visible = false;
			// 
			// cmbHasta
			// 
			dateButton2.Caption = "Today";
			this.cmbHasta.DateButtons.Add(dateButton2);
			this.cmbHasta.Format = "MMM/yyyy";
			this.cmbHasta.Location = new System.Drawing.Point(72, 267);
			this.cmbHasta.Name = "cmbHasta";
			this.cmbHasta.NonAutoSizeHeight = 23;
			this.cmbHasta.Size = new System.Drawing.Size(120, 21);
			this.cmbHasta.SpinButtonIncrement = Infragistics.Win.UltraWinSchedule.SpinIncrementUnit.Months;
			this.cmbHasta.SpinButtonsVisible = true;
			this.cmbHasta.TabIndex = 48;
			this.cmbHasta.Value = new System.DateTime(2009, 1, 17, 0, 0, 0, 0);
			this.cmbHasta.Visible = false;
			// 
			// chkResumen
			// 
			this.chkResumen.Location = new System.Drawing.Point(32, 302);
			this.chkResumen.Name = "chkResumen";
			this.chkResumen.Size = new System.Drawing.Size(96, 21);
			this.chkResumen.TabIndex = 50;
			this.chkResumen.Text = "Resumen";
			this.chkResumen.Visible = false;
			this.chkResumen.CheckedChanged += new System.EventHandler(this.chkResumen_CheckedChanged);
			// 
			// chkBanco
			// 
			this.chkBanco.Checked = true;
			this.chkBanco.CheckState = System.Windows.Forms.CheckState.Checked;
			this.chkBanco.Location = new System.Drawing.Point(208, 184);
			this.chkBanco.Name = "chkBanco";
			this.chkBanco.Size = new System.Drawing.Size(104, 22);
			this.chkBanco.TabIndex = 52;
			this.chkBanco.Text = "Todo Banco";
			this.chkBanco.CheckedChanged += new System.EventHandler(this.chkBanco_CheckedChanged);
			// 
			// cmbBanco
			// 
			this.cmbBanco.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbBanco.Cursor = System.Windows.Forms.Cursors.WaitCursor;
			this.cmbBanco.DataMember = "Banco";
			this.cmbBanco.DataSource = this.cdsBanco;
			this.cmbBanco.DisplayLayout.AutoFitColumns = true;
			ultraGridBand7.AddButtonCaption = "Reporte";
			ultraGridColumn131.Header.VisiblePosition = 1;
			ultraGridColumn131.Hidden = true;
			ultraGridColumn131.Width = 19;
			ultraGridColumn132.Header.VisiblePosition = 0;
			ultraGridColumn132.Width = 151;
			ultraGridColumn133.Header.VisiblePosition = 2;
			ultraGridColumn133.Hidden = true;
			ultraGridColumn133.Width = 32;
			ultraGridColumn134.Header.VisiblePosition = 3;
			ultraGridColumn134.Hidden = true;
			ultraGridColumn134.Width = 32;
			ultraGridColumn135.Header.VisiblePosition = 4;
			ultraGridColumn135.Hidden = true;
			ultraGridColumn135.Width = 17;
			ultraGridColumn136.Header.VisiblePosition = 5;
			ultraGridColumn136.Hidden = true;
			ultraGridColumn136.Width = 32;
			ultraGridColumn137.Header.VisiblePosition = 6;
			ultraGridColumn137.Hidden = true;
			ultraGridColumn137.Width = 33;
			ultraGridColumn138.Header.VisiblePosition = 7;
			ultraGridColumn138.Width = 46;
			ultraGridBand7.Columns.AddRange(new object[] {
																										 ultraGridColumn131,
																										 ultraGridColumn132,
																										 ultraGridColumn133,
																										 ultraGridColumn134,
																										 ultraGridColumn135,
																										 ultraGridColumn136,
																										 ultraGridColumn137,
																										 ultraGridColumn138});
			this.cmbBanco.DisplayLayout.BandsSerializer.Add(ultraGridBand7);
			this.cmbBanco.DisplayMember = "Nombre";
			this.cmbBanco.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbBanco.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbBanco.Enabled = false;
			this.cmbBanco.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbBanco.Location = new System.Drawing.Point(344, 184);
			this.cmbBanco.Name = "cmbBanco";
			this.cmbBanco.Size = new System.Drawing.Size(216, 21);
			this.cmbBanco.TabIndex = 51;
			this.cmbBanco.ValueMember = "Nombre";
			// 
			// cdsBanco
			// 
			this.cdsBanco.BindingContextCtrl = this;
			this.cdsBanco.DataLibrary = "LibContabilidad";
			this.cdsBanco.DataLibraryUrl = "";
			this.cdsBanco.DataSetDef = "dsCuentaCorriente";
			this.cdsBanco.Locale = new System.Globalization.CultureInfo("es-EC");
			this.cdsBanco.Name = "cdsBanco";
			this.cdsBanco.SchemaClassName = "LibContabilidad.DataClass";
			this.cdsBanco.SchemaDef = null;
			// 
			// chkDespedidos
			// 
			this.chkDespedidos.Location = new System.Drawing.Point(32, 327);
			this.chkDespedidos.Name = "chkDespedidos";
			this.chkDespedidos.Size = new System.Drawing.Size(184, 22);
			this.chkDespedidos.TabIndex = 53;
			this.chkDespedidos.Text = "Imprimir los que salieron";
			// 
			// chkCentroCosto
			// 
			this.chkCentroCosto.Checked = true;
			this.chkCentroCosto.CheckState = System.Windows.Forms.CheckState.Checked;
			this.chkCentroCosto.Location = new System.Drawing.Point(208, 216);
			this.chkCentroCosto.Name = "chkCentroCosto";
			this.chkCentroCosto.Size = new System.Drawing.Size(120, 22);
			this.chkCentroCosto.TabIndex = 55;
			this.chkCentroCosto.Text = "Todo Centro Costo";
			this.chkCentroCosto.CheckedChanged += new System.EventHandler(this.chkCentroCosto_CheckedChanged);
			// 
			// cmbCentroCosto
			// 
			this.cmbCentroCosto.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbCentroCosto.Cursor = System.Windows.Forms.Cursors.Default;
			ultraGridColumn139.Header.VisiblePosition = 0;
			ultraGridColumn139.Hidden = true;
			ultraGridColumn140.Header.VisiblePosition = 1;
			ultraGridBand8.Columns.AddRange(new object[] {
																										 ultraGridColumn139,
																										 ultraGridColumn140});
			this.cmbCentroCosto.DisplayLayout.BandsSerializer.Add(ultraGridBand8);
			this.cmbCentroCosto.DisplayMember = "Nombre";
			this.cmbCentroCosto.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbCentroCosto.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbCentroCosto.Enabled = false;
			this.cmbCentroCosto.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbCentroCosto.Location = new System.Drawing.Point(344, 216);
			this.cmbCentroCosto.Name = "cmbCentroCosto";
			this.cmbCentroCosto.Size = new System.Drawing.Size(216, 21);
			this.cmbCentroCosto.TabIndex = 54;
			this.cmbCentroCosto.ValueMember = "idProyecto";
			// 
			// cmbDistrib
			// 
			this.cmbDistrib.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbDistrib.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbDistrib.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsPersonal, "Personal.idAsiDistGrupo"));
			ultraGridColumn141.Header.VisiblePosition = 0;
			ultraGridColumn141.Hidden = true;
			ultraGridColumn142.Header.VisiblePosition = 1;
			ultraGridBand9.Columns.AddRange(new object[] {
																										 ultraGridColumn141,
																										 ultraGridColumn142});
			this.cmbDistrib.DisplayLayout.BandsSerializer.Add(ultraGridBand9);
			this.cmbDistrib.DisplayMember = "Distribucion";
			this.cmbDistrib.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbDistrib.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbDistrib.Enabled = false;
			this.cmbDistrib.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbDistrib.Location = new System.Drawing.Point(344, 248);
			this.cmbDistrib.Name = "cmbDistrib";
			this.cmbDistrib.Size = new System.Drawing.Size(216, 21);
			this.cmbDistrib.TabIndex = 56;
			this.cmbDistrib.ValueMember = "idAsiDistGrupo";
			// 
			// chkDistribucion
			// 
			this.chkDistribucion.Checked = true;
			this.chkDistribucion.CheckState = System.Windows.Forms.CheckState.Checked;
			this.chkDistribucion.Location = new System.Drawing.Point(208, 248);
			this.chkDistribucion.Name = "chkDistribucion";
			this.chkDistribucion.Size = new System.Drawing.Size(120, 22);
			this.chkDistribucion.TabIndex = 57;
			this.chkDistribucion.Text = "Toda Distribución";
			this.chkDistribucion.CheckedChanged += new System.EventHandler(this.chkDistribucion_CheckedChanged);
			// 
			// RepRoles
			// 
			this.AcceptButton = this.btAceptar;
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 14);
			this.CancelButton = this.btCancelar;
			this.ClientSize = new System.Drawing.Size(598, 374);
			this.Controls.Add(this.chkDistribucion);
			this.Controls.Add(this.cmbDistrib);
			this.Controls.Add(this.chkCentroCosto);
			this.Controls.Add(this.cmbCentroCosto);
			this.Controls.Add(this.chkDespedidos);
			this.Controls.Add(this.chkBanco);
			this.Controls.Add(this.cmbBanco);
			this.Controls.Add(this.chkResumen);
			this.Controls.Add(this.lblHasta);
			this.Controls.Add(this.label2);
			this.Controls.Add(this.label1);
			this.Controls.Add(this.cmbHasta);
			this.Controls.Add(this.chkGrupo);
			this.Controls.Add(this.cmbGrupo);
			this.Controls.Add(this.cmbReporte);
			this.Controls.Add(this.btCancelar);
			this.Controls.Add(this.btAceptar);
			this.Controls.Add(this.cmbFecha);
			this.Controls.Add(this.chkCargo);
			this.Controls.Add(this.cmbCargo);
			this.Controls.Add(this.chkSucursal);
			this.Controls.Add(this.cmbSucursal);
			this.Controls.Add(this.chkDepartamento);
			this.Controls.Add(this.cmbDepartamento);
			this.Controls.Add(this.chkPersonal);
			this.Controls.Add(this.cmbPersonal);
			this.Controls.Add(this.optTipo);
			this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
			this.MaximizeBox = false;
			this.Name = "RepRoles";
			this.Text = "Reportes de Roles";
			this.Closing += new System.ComponentModel.CancelEventHandler(this.RepRoles_Closing);
			this.Load += new System.EventHandler(this.RepRoles_Load);
			((System.ComponentModel.ISupportInitialize)(this.optTipo)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsPersonal)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsDepartamento)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbPersonal)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdvPersonal)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbDepartamento)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbSucursal)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdvSucursal)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbCargo)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbFecha)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsReporte)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdvReporte)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbReporte)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbGrupo)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsRolGrupo)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbHasta)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbBanco)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsBanco)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbCentroCosto)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbDistrib)).EndInit();
			this.ResumeLayout(false);

		}
		#endregion

		private void cdsPersonal_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			cdsPersonal.Schema.Connections[0].ConnectionString = 
				miFuncion.AccesoBase(MenuLatinium.stDirRoles);		
		}

		private void cdsDepartamento_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			cdsDepartamento.Schema.Connections[0].ConnectionString = 
				miFuncion.AccesoBase(MenuLatinium.stDirRoles);		
		}

		private void cdsReporte_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
		}

		private void optTipo_ValueChanged(object sender, System.EventArgs e)
		{
			int IdGrupo = (int) optTipo.Value;
			string stFiltro = "Tipo = 6 And idGrupo = " + IdGrupo.ToString();
			this.cdvReporte.RowFilter = stFiltro;
			System.Data.DataRow[] drRep = this.cdsReporte.StorageDataSet.Tables["Reporte"].Select(stFiltro);
			if (drRep.Length > 0)
				cmbReporte.Value = drRep[0]["idReporte"];
			bool bHasta = false;
			if (IdGrupo == 3) bHasta = true;
			
			lblHasta.Visible = bHasta;
			cmbHasta.Visible = bHasta;
			chkResumen.Visible = bHasta;
		}

		private void RepRoles_Load(object sender, System.EventArgs e)
		{
			string stAudita = "Exec AuditaCrear 53, 6, 'RepRoles'";
			Funcion.EjecutaSQL(cdsDepartamento, stAudita, true);
			Cursor = Cursors.Default;
			optTipo.Value = 0;
			cmbFecha.Value = DateTime.Today;
			cmbHasta.Value = DateTime.Today;
			cmbCentroCosto.DataSource = Funcion.dvProcedimiento(cdsPersonal, "Exec ProyectoCarga 0");
			cmbDistrib.DataSource = Funcion.dvProcedimiento(cdsPersonal, "Select idAsiDistGrupo, Distribucion From AsientoDistGrupo");

			bool bFiltroSucursal = Funcion.bEjecutaSQL(cdsPersonal, 
				"Select IsNull(FiltroSucursal, 0) From SeteoRol");
			if (bFiltroSucursal)
			{			
				string iSuc = Funcion.sEscalarSQL(cdsPersonal, "Exec VerificaUsuarioRol");
				if (iSuc != "")
				{
					cdvSucursal.RowFilter = "idRolSucursal in(" + iSuc.ToString()+ " )";
					chkSucursal.Checked = false;
					chkSucursal.Enabled = false;
					cmbSucursal.Value=iSuc;
				}	
			}
		}

		private void btCancelar_Click(object sender, System.EventArgs e)
		{
			this.Close();
		}

		private void BorrarErrores()
		{
			errorProvider.SetError(cmbReporte, "");
			errorProvider.SetError(cmbPersonal, "");
			errorProvider.SetError(cmbDepartamento, "");
			errorProvider.SetError(cmbCargo, "");
			errorProvider.SetError(cmbSucursal, "");
			errorProvider.SetError(cmbBanco, "");
			errorProvider.SetError(optTipo, "");
		}

		private string[] Filtro()
		{
			bool bOk = true;
			DateTime dtFecha = (DateTime) cmbFecha.Value;
			string stFiltro = "";
			string stTitulo = "Mes " + dtFecha.ToString("MMMM") + " ";
			BorrarErrores();
      if (!chkBanco.Checked)
				if (cmbBanco.Text.Trim().Length == 0)
				{
					errorProvider.SetError(cmbBanco, "Ingrese Banco");
					bOk = false;
				}
				else
				{
					stTitulo += "Banco " + cmbBanco.Text.Trim();
					stFiltro += " And {Personal.Banco} = '" + cmbBanco.Value.ToString().Trim() + "' ";
				}

			if (!chkDespedidos.Checked)
				stFiltro += " And IsNull({Personal.FechaSalida}) ";
			else
				stTitulo += "Incluye Retirados ";

			if (!chkDistribucion.Checked)
				if (cmbDistrib.Text.Trim().Length == 0)
				{
					errorProvider.SetError(cmbDistrib, "Ingrese Distribucion");
					bOk = false;
				}
				else
				{
					stTitulo += "Distribucion " + cmbDistrib.Text.Trim();
					stFiltro += " And {RolPago.idAsiDistGrupo} = " + cmbDistrib.Value;
				}

			if (!chkCentroCosto.Checked)
				if (cmbCentroCosto.Text.Trim().Length == 0)
				{
					errorProvider.SetError(cmbCentroCosto, "Ingrese Centro de Costo");
					bOk = false;
				}
				else
				{
					stTitulo += "Centro de Costo " + cmbCentroCosto.Text.Trim();
					stFiltro += " And {RolPago.idCentroCosto} = " + cmbCentroCosto.Value;
				}

			if (!chkDepartamento.Checked)
				if (cmbDepartamento.Text.Trim().Length == 0)
				{
					errorProvider.SetError(cmbDepartamento, "Ingrese Departamento");
					bOk = false;
				}
				else
				{
					stTitulo += "Departamento " + cmbDepartamento.Text.Trim();
					stFiltro += " And {RolPago.idDepartamento} = " + cmbDepartamento.Value;
				}

			if (!chkCargo.Checked)
				if (cmbCargo.Text.Trim().Length == 0)
				{
					errorProvider.SetError(cmbCargo, "Ingrese Cargo");
					bOk = false;
				}
				else 
				{
					stTitulo += "Cargo " + cmbCargo.Text.Trim();
					stFiltro += " And {Personal.idCargo} = " + cmbCargo.Value;
				}

			if (!chkPersonal.Checked)
				if (cmbPersonal.Text.Trim().Length == 0)
				{
					errorProvider.SetError(cmbPersonal, "Ingrese Personal");
					bOk = false;
				}
				else 
				{
					stFiltro += " And {Personal.idPersonal} = " + cmbPersonal.Value;
				}

			if (!chkSucursal.Checked)
				if (cmbSucursal.Text.Trim().Length == 0)
				{
					errorProvider.SetError(cmbSucursal, "Ingrese Sucursal");
					bOk = false;
				}
				else 
				{
					stTitulo += "Sucursal " + cmbSucursal.Text.Trim();
					stFiltro += " And {Personal.idRolSucursal} = " + cmbSucursal.Value;
				}

			if (!chkGrupo.Checked)
				if (cmbGrupo.Text.Trim().Length == 0)
				{
					errorProvider.SetError(cmbGrupo, "Ingrese Grupo");
					bOk = false;
				}
				else 
				{
					stTitulo += "Grupo " + cmbGrupo.Text.Trim();
					stFiltro += " And {RolPago.idRolGrupo} = " + cmbGrupo.Value;
				}
			string [] stSalida = new string[2];
			stSalida[0] = stTitulo;
			stSalida[1] = stFiltro;
			if (!bOk) stSalida[0] = "";
			return stSalida;
		}

		private bool Verifica()
		{
			bool bOk = true;
			BorrarErrores();
			if (!chkDepartamento.Checked)
				if (cmbDepartamento.Text.Trim().Length == 0)
				{
					errorProvider.SetError(cmbDepartamento, "Ingrese Departamento");
					bOk = false;
				}

			if (!chkCargo.Checked)
				if (cmbCargo.Text.Trim().Length == 0)
				{
					errorProvider.SetError(cmbCargo, "Ingrese Cargo");
					bOk = false;
				}

			if (!chkPersonal.Checked)
				if (cmbPersonal.Text.Trim().Length == 0)
				{
					errorProvider.SetError(cmbPersonal, "Ingrese Personal");
					bOk = false;
				}

			if (!chkSucursal.Checked)
				if (cmbSucursal.Text.Trim().Length == 0)
				{
					errorProvider.SetError(cmbSucursal, "Ingrese Sucursal");
					bOk = false;
				}

			if (!chkGrupo.Checked)
				if (cmbGrupo.Text.Trim().Length == 0)
				{
					errorProvider.SetError(cmbGrupo, "Ingrese Grupo");
					bOk = false;
				}
			DateTime dtFecha = (DateTime) cmbFecha.Value;
			//Decimo III en Diciembre
			if ((int) optTipo.Value == 4 && dtFecha.Month != 12)
			{
				errorProvider.SetError(optTipo, "Decimo Tercero solo se puede imprimir en Diciembre");
				bOk = false;
			}

//			// Decimo IV en Septiembre
//			if ((int) optTipo.Value == 5 && dtFecha.Month != 9)
//			{
//				errorProvider.SetError(optTipo, "Decimo Cuarto solo se puede imprimir en Septiembre");
//				bOk = false;
//			}
//
//			// Fondo de Reserva
//			if ((int) optTipo.Value == 6 && dtFecha.Month != 8)
//			{
//				errorProvider.SetError(optTipo, "Fondo de Reserva Solo se imprime en Agosto");
//				bOk = false;
//			}
			return bOk;
		}
		private void btAceptar_Click(object sender, System.EventArgs e)
		{
			if (!Verifica()) return;
			string [] stSalida = new string[2];
			stSalida = Filtro();
			if (stSalida[0].Length == 0) return;
			Cursor = Cursors.WaitCursor;

			string stReporte = this.cmbReporte.Text.Trim() + ".Rpt";
			DateTime dtFecha = (DateTime) cmbFecha.Value;
			string stTitulo = stSalida[0];
			string stFiltro = "";
		
      DateTime dtHasta = (DateTime) cmbHasta.Value;
			if (lblHasta.Visible)
			{
				if (dtHasta.Year == dtFecha.Year)
				{
					stFiltro = "({RolPago.Anio} = " + dtFecha.Year.ToString().Trim()
						+ " And {RolPago.Mes} >= " + dtFecha.Month.ToString().Trim() 
						+ " And {RolPago.Mes} <= " + dtHasta.Month.ToString().Trim() + ") " + stSalida[1];
				}
				else
				{
					stFiltro = "( ({RolPago.Anio} = " + dtFecha.Year.ToString().Trim()
						+ " And {RolPago.Mes} >= " + dtFecha.Month.ToString().Trim() 
						+ ") Or ({RolPago.Anio} = " + dtHasta.Year.ToString().Trim()
						+ " And {RolPago.Mes} <= " + dtHasta.Month.ToString().Trim() + ") )" + stSalida[1];
				}
					stTitulo = "Desde " + dtFecha.ToString("MMMM Del yyyy") + " Hasta " + dtHasta.ToString("MMMM Del yyyy");
			}
			else
			{
				stFiltro = "{RolPago.Anio} = " + dtFecha.Year.ToString().Trim()
					+ " And {RolPago.Mes} = " + dtFecha.Month.ToString().Trim() + stSalida[1];
			}
			string stFiltroRep = cmbReporte.ActiveRow.Cells["Filtro"].Value.ToString();
			if (stFiltroRep != null && stFiltroRep.Length > 0)
			{
				if (stFiltroRep.Trim().StartsWith("And")) 
					stFiltro += " " + stFiltroRep;
				else
				{
					string stExec = string.Format("Exec RolXIV '{0}01'", dtHasta.ToString("yyyyMM"));
					Funcion.EjecutaSQL(cdsBanco, stExec);
					stFiltro = stFiltroRep;
				}
			}

			Reporte miReporte = new Reporte(stReporte, stFiltro);
			miReporte.Titulo(stTitulo);
			if (chkResumen.Visible && chkResumen.Checked) miReporte.Resumen();
			miReporte.MdiParent = this.MdiParent;
			miReporte.Show();
			Cursor = Cursors.Default;
		}

		private void cdsRolGrupo_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			cdsRolGrupo.Schema.Connections[0].ConnectionString = 
				miFuncion.AccesoBase(MenuLatinium.stDirRoles);		
		}

		private void chkPersonal_CheckedChanged(object sender, System.EventArgs e)
		{
			cmbPersonal.Enabled = !chkPersonal.Checked;
			if (!chkPersonal.Checked) cmbPersonal.Select();
		}

		private void chkDepartamento_CheckedChanged(object sender, System.EventArgs e)
		{
			cmbDepartamento.Enabled = !chkDepartamento.Checked;
			if (!chkDepartamento.Checked) cmbDepartamento.Select();
		}

		private void chkCargo_CheckedChanged(object sender, System.EventArgs e)
		{
			cmbCargo.Enabled = !chkCargo.Checked;
			if (!chkCargo.Checked) cmbCargo.Select();
		}

		private void chkSucursal_CheckedChanged(object sender, System.EventArgs e)
		{
			cmbSucursal.Enabled = !chkSucursal.Checked;
			if (!chkSucursal.Checked) cmbSucursal.Select();
		}

		private void chkGrupo_CheckedChanged(object sender, System.EventArgs e)
		{
			cmbGrupo.Enabled = !chkGrupo.Checked;
			if (!chkGrupo.Checked) cmbGrupo.Select();
		}

		private void RepRoles_Closing(object sender, System.ComponentModel.CancelEventArgs e)
		{
			string stAudita = "Exec AuditaCrear 53, 7, 'RepRoles'";
			Funcion.EjecutaSQL(cdsDepartamento, stAudita, true);
		}

		private void chkBanco_CheckedChanged(object sender, System.EventArgs e)
		{
			cmbBanco.Enabled = !chkBanco.Checked;
			if (!chkBanco.Checked) cmbBanco.Select();
		}

		private void chkCentroCosto_CheckedChanged(object sender, System.EventArgs e)
		{
			cmbCentroCosto.Enabled = !chkCentroCosto.Checked;
			if (!chkCentroCosto.Checked) cmbCentroCosto.Select();
		}

		private void chkDistribucion_CheckedChanged(object sender, System.EventArgs e)
		{
			cmbDistrib.Enabled = !chkDistribucion.Checked;
			if (!chkDistribucion.Checked) cmbDistrib.Select();
		}

		private void cmbPersonal_InitializeLayout(object sender, Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs e)
		{
		
		}

		private void chkResumen_CheckedChanged(object sender, System.EventArgs e)
		{
		
		}

		private void cmbFecha_BeforeDropDown(object sender, System.ComponentModel.CancelEventArgs e)
		{
		
		}
	}
}
