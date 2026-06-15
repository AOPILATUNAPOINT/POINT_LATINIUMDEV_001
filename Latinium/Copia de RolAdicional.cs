using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using System.Globalization;

namespace Latinium
{
	/// <summary>
	/// Descripción breve de RolAdicional.
	/// </summary>
	public class RolAdicional : DevExpress.XtraEditors.XtraForm
	{
		private Infragistics.Win.UltraWinGrid.UltraGrid grdRetRenta;
		private BarraDatoSQL.BarraDatoSQL barraDato1;
		private C1.Data.C1DataSet cdsRolAdicional;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbPersonal;
		private Infragistics.Win.Misc.UltraButton btCarga;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtIdRolAdicional;
		private Infragistics.Win.UltraWinGrid.UltraDropDown cmbPersonales;
		private Infragistics.Win.Misc.UltraButton btSubida;
		private System.Windows.Forms.Label label1;
		private System.Windows.Forms.Label label2;
		private Infragistics.Win.UltraWinEditors.UltraCheckEditor chkResumen;
		private Infragistics.Win.UltraWinSchedule.UltraCalendarCombo cmbFecha;
		private System.Windows.Forms.Button button1;
		private System.Windows.Forms.Button button2;
		private System.Windows.Forms.Button button3;
		/// <summary>
		/// Variable del diseñador requerida.
		/// </summary>
		private System.ComponentModel.Container components = null;
		private C1.Data.C1DataSet cdsPersonal;
		private C1.Data.C1DataSet cdsSeguridad;
		private Infragistics.Win.UltraWinEditors.UltraCheckEditor chkRevisado;
		int IdRolAdicional = 0;
		public RolAdicional(int idRolA)
		{
			InitializeComponent();
			IdRolAdicional = idRolA;
		}
		public RolAdicional()
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
			Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton dateButton1 = new Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton();
			Infragistics.Win.Appearance appearance1 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand1 = new Infragistics.Win.UltraWinGrid.UltraGridBand("_RolAdicionalesDetalle", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn1 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idRolAdicionalDetalle");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn2 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idRolAdicionales");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn3 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idPersonal", -1, "cmbPersonales");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn4 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Valor1");
			Infragistics.Win.Appearance appearance2 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn5 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Valor2");
			Infragistics.Win.Appearance appearance3 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn6 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Valor3");
			Infragistics.Win.Appearance appearance4 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn7 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("HoraEntrada");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn8 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("HoraSalida");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn9 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idHorario");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn10 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Libre");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn11 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("CL");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn12 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Diferencia");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn13 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Observaciones");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn14 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("HorasTotales");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn15 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Atrasos");
			Infragistics.Win.Appearance appearance5 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance6 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance7 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance8 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand2 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn16 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idPersonal");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn17 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand3 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Personal", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn18 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idPersonal");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn19 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idAsiDistGrupo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn20 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCentroCosto");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn21 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idDepartamento");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn22 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCargo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn23 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idRolSucursal");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn24 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idRolGrupo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn25 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Jornal");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn26 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn27 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Apellido");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn28 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Codigo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn29 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Jornada");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn30 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Direccion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn31 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Telefono");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn32 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Celular");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn33 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Cedula");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn34 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Afiliacion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn35 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Reserva");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn36 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Renta");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn37 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Iess");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn38 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Sueldo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn39 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Comision");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn40 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Tipo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn41 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("FechaIngreso");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn42 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("FechaNacimiento");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn43 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("FechaSalida");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn44 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Notas");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn45 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("CargasFamiliares");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn46 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idProvincia");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn47 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCiudad");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn48 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Numero");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn49 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("SistemaSalario");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn50 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Cuenta");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn51 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("CodigoSectorial");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn52 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Condicion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn53 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("FechaNovedad");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn54 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Acceso");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn55 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Vendedor");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn56 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Borrar");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn57 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Cobrador");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn58 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Sexo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn59 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Estado_Civil");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn60 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Banco");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn61 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("CuentaCorr");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn62 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idTipoCuenta");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn63 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("CedulaMilitar");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn64 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idFormaPago");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn65 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Quincena");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn66 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("CodIngreso");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn67 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("CodTelefono");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn68 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("TarjEdificio");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn69 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("TarjetaOficina");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn70 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("email");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn71 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idProyecto");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn72 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCuenta");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn73 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Bono");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn74 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Bono1");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn75 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Desc1");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn76 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Desc2");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn77 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("TieneQuincena");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn78 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Fotos");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn79 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Convenio");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn80 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idSubProyecto");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn81 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Sindicato");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn82 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Asociacion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn83 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idRolHorario");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn84 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("FechaContrato");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn85 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Educacion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn86 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("NumeroDias");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn87 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idRelacionT");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn88 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("gSalud");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn89 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("gEducacion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn90 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("gVivienda");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn91 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("gVestimenta");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn92 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("gAlimentacion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn93 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("gTotal");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn94 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("BaseComision");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn95 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("PorcentajeComision");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn96 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("gDesahucio");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn97 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("gDiscapacidad");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn98 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("IngGravOE");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn99 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("DeduccionOE");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn100 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("OtrasRebOE");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn101 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("ImpRentaAsumido");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn102 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("gRebajaIIIEdad");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn103 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("RentaRetenidaOE");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn104 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("RentaRetOE");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn105 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("TipoSangre");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn106 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("RegEntrada");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn107 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("RegSalida");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn108 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idTipoContrato");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn109 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("CodVendedor");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn110 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Clave");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn111 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idSupervisor");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn112 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("bSupervisor");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn113 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("E_Profesor");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn114 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Dias360");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn115 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("DiasXIII");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn116 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idJefe");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn117 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("PoseeDiscapacidad");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn118 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("PorcDiscapacidad");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn119 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("TipoDiscapacidad");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn120 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("NumConadis");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn121 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("TipoVivienda");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn122 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Vehiculo");
			this.cmbFecha = new Infragistics.Win.UltraWinSchedule.UltraCalendarCombo();
			this.cdsRolAdicional = new C1.Data.C1DataSet();
			this.grdRetRenta = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.barraDato1 = new BarraDatoSQL.BarraDatoSQL();
			this.cmbPersonal = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.btCarga = new Infragistics.Win.Misc.UltraButton();
			this.txtIdRolAdicional = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.cmbPersonales = new Infragistics.Win.UltraWinGrid.UltraDropDown();
			this.cdsPersonal = new C1.Data.C1DataSet();
			this.btSubida = new Infragistics.Win.Misc.UltraButton();
			this.label1 = new System.Windows.Forms.Label();
			this.label2 = new System.Windows.Forms.Label();
			this.chkResumen = new Infragistics.Win.UltraWinEditors.UltraCheckEditor();
			this.button1 = new System.Windows.Forms.Button();
			this.button2 = new System.Windows.Forms.Button();
			this.button3 = new System.Windows.Forms.Button();
			this.chkRevisado = new Infragistics.Win.UltraWinEditors.UltraCheckEditor();
			this.cdsSeguridad = new C1.Data.C1DataSet();
			((System.ComponentModel.ISupportInitialize)(this.cmbFecha)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsRolAdicional)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.grdRetRenta)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbPersonal)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIdRolAdicional)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbPersonales)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsPersonal)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeguridad)).BeginInit();
			this.SuspendLayout();
			// 
			// cmbFecha
			// 
			this.cmbFecha.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsRolAdicional, "_RolAdicionales.Fecha"));
			dateButton1.Caption = "Today";
			this.cmbFecha.DateButtons.Add(dateButton1);
			this.cmbFecha.Location = new System.Drawing.Point(360, 26);
			this.cmbFecha.Name = "cmbFecha";
			this.cmbFecha.NonAutoSizeHeight = 23;
			this.cmbFecha.Size = new System.Drawing.Size(96, 21);
			this.cmbFecha.TabIndex = 30;
			this.cmbFecha.Value = new System.DateTime(2011, 9, 7, 0, 0, 0, 0);
			// 
			// cdsRolAdicional
			// 
			this.cdsRolAdicional.BindingContextCtrl = this;
			this.cdsRolAdicional.DataLibrary = "LibPersonal";
			this.cdsRolAdicional.DataLibraryUrl = "";
			this.cdsRolAdicional.DataSetDef = "dsRolAdicional";
			this.cdsRolAdicional.EnforceConstraints = false;
			this.cdsRolAdicional.Locale = new System.Globalization.CultureInfo("es-EC");
			this.cdsRolAdicional.Name = "cdsRolAdicional";
			this.cdsRolAdicional.SchemaClassName = "LibPersonal.DataClass";
			this.cdsRolAdicional.SchemaDef = null;
			this.cdsRolAdicional.BeforeFill += new C1.Data.FillEventHandler(this.cdsRolAdicional_BeforeFill);
			// 
			// grdRetRenta
			// 
			this.grdRetRenta.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.grdRetRenta.Cursor = System.Windows.Forms.Cursors.Default;
			this.grdRetRenta.DataMember = "_RolAdicionales.RolAdicionales - RolAdicionalesDetalle";
			this.grdRetRenta.DataSource = this.cdsRolAdicional;
			appearance1.BackColor = System.Drawing.Color.White;
			this.grdRetRenta.DisplayLayout.Appearance = appearance1;
			this.grdRetRenta.DisplayLayout.AutoFitColumns = true;
			ultraGridBand1.AddButtonCaption = "RetDetalle";
			ultraGridColumn1.Header.VisiblePosition = 0;
			ultraGridColumn1.Hidden = true;
			ultraGridColumn1.Width = 94;
			ultraGridColumn2.Header.VisiblePosition = 1;
			ultraGridColumn2.Hidden = true;
			ultraGridColumn2.Width = 75;
			ultraGridColumn3.Header.Caption = "Personal";
			ultraGridColumn3.Header.VisiblePosition = 2;
			ultraGridColumn3.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.DropDownList;
			ultraGridColumn3.Width = 182;
			appearance2.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn4.CellAppearance = appearance2;
			ultraGridColumn4.Format = "n2";
			ultraGridColumn4.Header.VisiblePosition = 5;
			ultraGridColumn4.Width = 45;
			appearance3.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn5.CellAppearance = appearance3;
			ultraGridColumn5.Format = "n2";
			ultraGridColumn5.Header.VisiblePosition = 6;
			ultraGridColumn5.Width = 41;
			appearance4.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn6.CellAppearance = appearance4;
			ultraGridColumn6.Format = "n2";
			ultraGridColumn6.Header.VisiblePosition = 7;
			ultraGridColumn6.Width = 47;
			ultraGridColumn7.EditorControl = this.cmbFecha;
			ultraGridColumn7.Format = "dd/MMM/yyyy HH:mm";
			ultraGridColumn7.Header.VisiblePosition = 3;
			ultraGridColumn7.VertScrollBar = true;
			ultraGridColumn7.Width = 110;
			ultraGridColumn8.EditorControl = this.cmbFecha;
			ultraGridColumn8.Format = "dd/MMM/yyyy HH:mm";
			ultraGridColumn8.Header.VisiblePosition = 4;
			ultraGridColumn8.VertScrollBar = true;
			ultraGridColumn8.Width = 89;
			ultraGridColumn9.Header.VisiblePosition = 8;
			ultraGridColumn9.Hidden = true;
			ultraGridColumn9.Width = 42;
			ultraGridColumn10.Header.Caption = "Doblar";
			ultraGridColumn10.Header.VisiblePosition = 9;
			ultraGridColumn10.Hidden = true;
			ultraGridColumn10.Width = 23;
			ultraGridColumn11.Header.Caption = "CLibre";
			ultraGridColumn11.Header.VisiblePosition = 10;
			ultraGridColumn11.Width = 23;
			ultraGridColumn12.Header.VisiblePosition = 11;
			ultraGridColumn12.Hidden = true;
			ultraGridColumn12.Width = 40;
			ultraGridColumn13.Header.VisiblePosition = 14;
			ultraGridColumn13.Width = 85;
			ultraGridColumn14.Header.VisiblePosition = 12;
			ultraGridColumn14.Width = 112;
			ultraGridColumn15.Header.VisiblePosition = 13;
			ultraGridColumn15.Width = 85;
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
																										 ultraGridColumn15});
			this.grdRetRenta.DisplayLayout.BandsSerializer.Add(ultraGridBand1);
			this.grdRetRenta.DisplayLayout.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.TemplateOnBottom;
			this.grdRetRenta.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.True;
			this.grdRetRenta.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.True;
			appearance5.BackColor = System.Drawing.Color.Transparent;
			this.grdRetRenta.DisplayLayout.Override.CardAreaAppearance = appearance5;
			appearance6.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance6.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance6.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance6.FontData.BoldAsString = "True";
			appearance6.FontData.Name = "Arial";
			appearance6.FontData.SizeInPoints = 10F;
			appearance6.ForeColor = System.Drawing.Color.White;
			appearance6.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.grdRetRenta.DisplayLayout.Override.HeaderAppearance = appearance6;
			appearance7.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance7.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance7.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.grdRetRenta.DisplayLayout.Override.RowSelectorAppearance = appearance7;
			appearance8.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance8.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance8.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.grdRetRenta.DisplayLayout.Override.SelectedRowAppearance = appearance8;
			this.grdRetRenta.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.grdRetRenta.Location = new System.Drawing.Point(8, 78);
			this.grdRetRenta.Name = "grdRetRenta";
			this.grdRetRenta.Size = new System.Drawing.Size(840, 292);
			this.grdRetRenta.TabIndex = 27;
			this.grdRetRenta.InitializeLayout += new Infragistics.Win.UltraWinGrid.InitializeLayoutEventHandler(this.grdRetRenta_InitializeLayout);
			// 
			// barraDato1
			// 
			this.barraDato1.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.barraDato1.BarraMovimiento = true;
			this.barraDato1.DataMember = "RolAdicionales";
			this.barraDato1.DataNombreId = "idRolAdicionales";
			this.barraDato1.DataOrden = "";
			this.barraDato1.DataSource = this.cdsRolAdicional;
			this.barraDato1.DataTabla = "RolAdicionales";
			this.barraDato1.DataTablaHija = "";
			this.barraDato1.Location = new System.Drawing.Point(24, 405);
			this.barraDato1.Name = "barraDato1";
			this.barraDato1.Size = new System.Drawing.Size(304, 26);
			this.barraDato1.TabIndex = 28;
			this.barraDato1.VisibleBorrar = true;
			this.barraDato1.VisibleBuscar = false;
			this.barraDato1.VisibleEditar = true;
			this.barraDato1.VisibleImprimir = true;
			this.barraDato1.VisibleNuevo = true;
			this.barraDato1.Load += new System.EventHandler(this.barraDato1_Load);
			this.barraDato1.Borrar += new System.EventHandler(this.barraDato1_Borrar_1);
			this.barraDato1.Imprimir += new System.EventHandler(this.barraDato1_Imprimir);
			this.barraDato1.Restaurar += new System.EventHandler(this.barraDato1_Restaurar);
			this.barraDato1.Refresca += new System.EventHandler(this.barraDato1_Refresca);
			this.barraDato1.Crear += new System.EventHandler(this.barraDato1_Crear);
			this.barraDato1.Deshacer += new System.EventHandler(this.barraDato1_Deshacer);
			this.barraDato1.Impresora += new System.EventHandler(this.barraDato1_Impresora);
			this.barraDato1.Editar += new System.EventHandler(this.barraDato1_Editar);
			this.barraDato1.Grabar += new System.EventHandler(this.barraDato1_Grabar);
			// 
			// cmbPersonal
			// 
			this.cmbPersonal.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbPersonal.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbPersonal.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsRolAdicional, "_RolAdicionales.idPersonal"));
			ultraGridColumn16.Header.VisiblePosition = 0;
			ultraGridColumn16.Hidden = true;
			ultraGridColumn17.Header.VisiblePosition = 1;
			ultraGridColumn17.Width = 245;
			ultraGridBand2.Columns.AddRange(new object[] {
																										 ultraGridColumn16,
																										 ultraGridColumn17});
			this.cmbPersonal.DisplayLayout.BandsSerializer.Add(ultraGridBand2);
			this.cmbPersonal.DisplayMember = "Nombre";
			this.cmbPersonal.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbPersonal.Location = new System.Drawing.Point(88, 26);
			this.cmbPersonal.Name = "cmbPersonal";
			this.cmbPersonal.Size = new System.Drawing.Size(208, 21);
			this.cmbPersonal.TabIndex = 31;
			this.cmbPersonal.Text = "ultraCombo1";
			this.cmbPersonal.ValueMember = "idPersonal";
			// 
			// btCarga
			// 
			this.btCarga.Location = new System.Drawing.Point(472, 26);
			this.btCarga.Name = "btCarga";
			this.btCarga.Size = new System.Drawing.Size(88, 25);
			this.btCarga.TabIndex = 32;
			this.btCarga.Text = "Carga";
			this.btCarga.Click += new System.EventHandler(this.btCarga_Click);
			// 
			// txtIdRolAdicional
			// 
			this.txtIdRolAdicional.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.txtIdRolAdicional.DataBindings.Add(new System.Windows.Forms.Binding("Text", this.cdsRolAdicional, "_RolAdicionales.idRolAdicionales"));
			this.txtIdRolAdicional.Location = new System.Drawing.Point(584, 405);
			this.txtIdRolAdicional.Name = "txtIdRolAdicional";
			this.txtIdRolAdicional.Size = new System.Drawing.Size(48, 22);
			this.txtIdRolAdicional.TabIndex = 33;
			this.txtIdRolAdicional.Text = "id";
			// 
			// cmbPersonales
			// 
			this.cmbPersonales.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbPersonales.DataSource = this.cdsPersonal;
			this.cmbPersonales.DisplayLayout.AutoFitColumns = true;
			ultraGridBand3.AddButtonCaption = "Band 0";
			ultraGridColumn18.Header.VisiblePosition = 0;
			ultraGridColumn18.Hidden = true;
			ultraGridColumn19.Header.VisiblePosition = 1;
			ultraGridColumn19.Hidden = true;
			ultraGridColumn19.Width = 8;
			ultraGridColumn20.Header.VisiblePosition = 3;
			ultraGridColumn20.Hidden = true;
			ultraGridColumn20.Width = 8;
			ultraGridColumn21.Header.VisiblePosition = 4;
			ultraGridColumn21.Hidden = true;
			ultraGridColumn21.Width = 8;
			ultraGridColumn22.Header.VisiblePosition = 5;
			ultraGridColumn22.Hidden = true;
			ultraGridColumn22.Width = 8;
			ultraGridColumn23.Header.VisiblePosition = 6;
			ultraGridColumn23.Hidden = true;
			ultraGridColumn23.Width = 8;
			ultraGridColumn24.Header.VisiblePosition = 7;
			ultraGridColumn24.Hidden = true;
			ultraGridColumn24.Width = 8;
			ultraGridColumn25.Header.VisiblePosition = 8;
			ultraGridColumn25.Hidden = true;
			ultraGridColumn25.Width = 8;
			ultraGridColumn26.Header.VisiblePosition = 2;
			ultraGridColumn26.Width = 63;
			ultraGridColumn27.Header.VisiblePosition = 9;
			ultraGridColumn27.Hidden = true;
			ultraGridColumn27.Width = 8;
			ultraGridColumn28.Header.VisiblePosition = 10;
			ultraGridColumn28.Hidden = true;
			ultraGridColumn28.Width = 8;
			ultraGridColumn29.Header.VisiblePosition = 11;
			ultraGridColumn29.Hidden = true;
			ultraGridColumn29.Width = 8;
			ultraGridColumn30.Header.VisiblePosition = 12;
			ultraGridColumn30.Hidden = true;
			ultraGridColumn30.Width = 8;
			ultraGridColumn31.Header.VisiblePosition = 13;
			ultraGridColumn31.Hidden = true;
			ultraGridColumn31.Width = 8;
			ultraGridColumn32.Header.VisiblePosition = 14;
			ultraGridColumn32.Hidden = true;
			ultraGridColumn32.Width = 8;
			ultraGridColumn33.Header.VisiblePosition = 15;
			ultraGridColumn33.Hidden = true;
			ultraGridColumn33.Width = 8;
			ultraGridColumn34.Header.VisiblePosition = 16;
			ultraGridColumn34.Hidden = true;
			ultraGridColumn34.Width = 8;
			ultraGridColumn35.Header.VisiblePosition = 17;
			ultraGridColumn35.Hidden = true;
			ultraGridColumn35.Width = 16;
			ultraGridColumn36.Header.VisiblePosition = 18;
			ultraGridColumn36.Hidden = true;
			ultraGridColumn36.Width = 16;
			ultraGridColumn37.Header.VisiblePosition = 19;
			ultraGridColumn37.Hidden = true;
			ultraGridColumn37.Width = 16;
			ultraGridColumn38.Header.VisiblePosition = 20;
			ultraGridColumn38.Hidden = true;
			ultraGridColumn38.Width = 8;
			ultraGridColumn39.Header.VisiblePosition = 21;
			ultraGridColumn39.Hidden = true;
			ultraGridColumn39.Width = 8;
			ultraGridColumn40.Header.VisiblePosition = 22;
			ultraGridColumn40.Hidden = true;
			ultraGridColumn40.Width = 8;
			ultraGridColumn41.Header.VisiblePosition = 23;
			ultraGridColumn41.Hidden = true;
			ultraGridColumn41.Width = 16;
			ultraGridColumn42.Header.VisiblePosition = 24;
			ultraGridColumn42.Hidden = true;
			ultraGridColumn42.Width = 16;
			ultraGridColumn43.Header.VisiblePosition = 25;
			ultraGridColumn43.Hidden = true;
			ultraGridColumn43.Width = 16;
			ultraGridColumn44.Header.VisiblePosition = 26;
			ultraGridColumn44.Hidden = true;
			ultraGridColumn44.Width = 8;
			ultraGridColumn45.Header.VisiblePosition = 27;
			ultraGridColumn45.Hidden = true;
			ultraGridColumn45.Width = 8;
			ultraGridColumn46.Header.VisiblePosition = 28;
			ultraGridColumn46.Hidden = true;
			ultraGridColumn46.Width = 8;
			ultraGridColumn47.Header.VisiblePosition = 29;
			ultraGridColumn47.Hidden = true;
			ultraGridColumn47.Width = 8;
			ultraGridColumn48.Header.VisiblePosition = 30;
			ultraGridColumn48.Hidden = true;
			ultraGridColumn48.Width = 8;
			ultraGridColumn49.Header.VisiblePosition = 31;
			ultraGridColumn49.Hidden = true;
			ultraGridColumn49.Width = 8;
			ultraGridColumn50.Header.VisiblePosition = 32;
			ultraGridColumn50.Hidden = true;
			ultraGridColumn50.Width = 8;
			ultraGridColumn51.Header.VisiblePosition = 33;
			ultraGridColumn51.Hidden = true;
			ultraGridColumn51.Width = 8;
			ultraGridColumn52.Header.VisiblePosition = 34;
			ultraGridColumn52.Hidden = true;
			ultraGridColumn52.Width = 8;
			ultraGridColumn53.Header.VisiblePosition = 35;
			ultraGridColumn53.Hidden = true;
			ultraGridColumn53.Width = 16;
			ultraGridColumn54.Header.VisiblePosition = 36;
			ultraGridColumn54.Hidden = true;
			ultraGridColumn54.Width = 16;
			ultraGridColumn55.Header.VisiblePosition = 37;
			ultraGridColumn55.Hidden = true;
			ultraGridColumn55.Width = 16;
			ultraGridColumn56.Header.VisiblePosition = 38;
			ultraGridColumn56.Hidden = true;
			ultraGridColumn56.Width = 16;
			ultraGridColumn57.Header.VisiblePosition = 39;
			ultraGridColumn57.Hidden = true;
			ultraGridColumn57.Width = 16;
			ultraGridColumn58.Header.VisiblePosition = 40;
			ultraGridColumn58.Hidden = true;
			ultraGridColumn58.Width = 8;
			ultraGridColumn59.Header.VisiblePosition = 41;
			ultraGridColumn59.Hidden = true;
			ultraGridColumn59.Width = 8;
			ultraGridColumn60.Header.VisiblePosition = 42;
			ultraGridColumn60.Hidden = true;
			ultraGridColumn60.Width = 8;
			ultraGridColumn61.Header.VisiblePosition = 43;
			ultraGridColumn61.Hidden = true;
			ultraGridColumn61.Width = 8;
			ultraGridColumn62.Header.VisiblePosition = 44;
			ultraGridColumn62.Hidden = true;
			ultraGridColumn62.Width = 8;
			ultraGridColumn63.Header.VisiblePosition = 45;
			ultraGridColumn63.Hidden = true;
			ultraGridColumn63.Width = 8;
			ultraGridColumn64.Header.VisiblePosition = 46;
			ultraGridColumn64.Hidden = true;
			ultraGridColumn64.Width = 8;
			ultraGridColumn65.Header.VisiblePosition = 47;
			ultraGridColumn65.Hidden = true;
			ultraGridColumn65.Width = 8;
			ultraGridColumn66.Header.VisiblePosition = 48;
			ultraGridColumn66.Hidden = true;
			ultraGridColumn66.Width = 8;
			ultraGridColumn67.Header.VisiblePosition = 49;
			ultraGridColumn67.Hidden = true;
			ultraGridColumn67.Width = 8;
			ultraGridColumn68.Header.VisiblePosition = 50;
			ultraGridColumn68.Hidden = true;
			ultraGridColumn68.Width = 8;
			ultraGridColumn69.Header.VisiblePosition = 51;
			ultraGridColumn69.Hidden = true;
			ultraGridColumn69.Width = 8;
			ultraGridColumn70.Header.VisiblePosition = 52;
			ultraGridColumn70.Hidden = true;
			ultraGridColumn70.Width = 8;
			ultraGridColumn71.Header.VisiblePosition = 53;
			ultraGridColumn71.Hidden = true;
			ultraGridColumn71.Width = 8;
			ultraGridColumn72.Header.VisiblePosition = 54;
			ultraGridColumn72.Hidden = true;
			ultraGridColumn72.Width = 8;
			ultraGridColumn73.Header.VisiblePosition = 55;
			ultraGridColumn73.Hidden = true;
			ultraGridColumn73.Width = 8;
			ultraGridColumn74.Header.VisiblePosition = 56;
			ultraGridColumn74.Hidden = true;
			ultraGridColumn74.Width = 8;
			ultraGridColumn75.Header.VisiblePosition = 57;
			ultraGridColumn75.Hidden = true;
			ultraGridColumn75.Width = 8;
			ultraGridColumn76.Header.VisiblePosition = 58;
			ultraGridColumn76.Hidden = true;
			ultraGridColumn76.Width = 8;
			ultraGridColumn77.Header.VisiblePosition = 59;
			ultraGridColumn77.Hidden = true;
			ultraGridColumn77.Width = 16;
			ultraGridColumn78.Header.VisiblePosition = 60;
			ultraGridColumn78.Hidden = true;
			ultraGridColumn78.Width = 8;
			ultraGridColumn79.Header.VisiblePosition = 61;
			ultraGridColumn79.Hidden = true;
			ultraGridColumn79.Width = 8;
			ultraGridColumn80.Header.VisiblePosition = 62;
			ultraGridColumn80.Hidden = true;
			ultraGridColumn80.Width = 8;
			ultraGridColumn81.Header.VisiblePosition = 63;
			ultraGridColumn81.Hidden = true;
			ultraGridColumn81.Width = 16;
			ultraGridColumn82.Header.VisiblePosition = 64;
			ultraGridColumn82.Hidden = true;
			ultraGridColumn82.Width = 16;
			ultraGridColumn83.Header.VisiblePosition = 65;
			ultraGridColumn83.Hidden = true;
			ultraGridColumn83.Width = 8;
			ultraGridColumn84.Header.VisiblePosition = 66;
			ultraGridColumn84.Hidden = true;
			ultraGridColumn84.Width = 16;
			ultraGridColumn85.Header.VisiblePosition = 67;
			ultraGridColumn85.Hidden = true;
			ultraGridColumn85.Width = 8;
			ultraGridColumn86.Header.VisiblePosition = 68;
			ultraGridColumn86.Hidden = true;
			ultraGridColumn86.Width = 8;
			ultraGridColumn87.Header.VisiblePosition = 69;
			ultraGridColumn87.Hidden = true;
			ultraGridColumn87.Width = 8;
			ultraGridColumn88.Header.VisiblePosition = 70;
			ultraGridColumn88.Hidden = true;
			ultraGridColumn88.Width = 8;
			ultraGridColumn89.Header.VisiblePosition = 71;
			ultraGridColumn89.Hidden = true;
			ultraGridColumn89.Width = 8;
			ultraGridColumn90.Header.VisiblePosition = 72;
			ultraGridColumn90.Hidden = true;
			ultraGridColumn90.Width = 8;
			ultraGridColumn91.Header.VisiblePosition = 73;
			ultraGridColumn91.Hidden = true;
			ultraGridColumn91.Width = 8;
			ultraGridColumn92.Header.VisiblePosition = 74;
			ultraGridColumn92.Hidden = true;
			ultraGridColumn92.Width = 8;
			ultraGridColumn93.Header.VisiblePosition = 75;
			ultraGridColumn93.Hidden = true;
			ultraGridColumn93.Width = 8;
			ultraGridColumn94.Header.VisiblePosition = 76;
			ultraGridColumn94.Hidden = true;
			ultraGridColumn94.Width = 8;
			ultraGridColumn95.Header.VisiblePosition = 77;
			ultraGridColumn95.Hidden = true;
			ultraGridColumn95.Width = 8;
			ultraGridColumn96.Header.VisiblePosition = 78;
			ultraGridColumn96.Hidden = true;
			ultraGridColumn96.Width = 8;
			ultraGridColumn97.Header.VisiblePosition = 79;
			ultraGridColumn97.Hidden = true;
			ultraGridColumn97.Width = 8;
			ultraGridColumn98.Header.VisiblePosition = 80;
			ultraGridColumn98.Hidden = true;
			ultraGridColumn98.Width = 8;
			ultraGridColumn99.Header.VisiblePosition = 81;
			ultraGridColumn99.Hidden = true;
			ultraGridColumn99.Width = 8;
			ultraGridColumn100.Header.VisiblePosition = 82;
			ultraGridColumn100.Hidden = true;
			ultraGridColumn100.Width = 8;
			ultraGridColumn101.Header.VisiblePosition = 83;
			ultraGridColumn101.Hidden = true;
			ultraGridColumn101.Width = 8;
			ultraGridColumn102.Header.VisiblePosition = 84;
			ultraGridColumn102.Hidden = true;
			ultraGridColumn102.Width = 8;
			ultraGridColumn103.Header.VisiblePosition = 85;
			ultraGridColumn103.Hidden = true;
			ultraGridColumn103.Width = 8;
			ultraGridColumn104.Header.VisiblePosition = 86;
			ultraGridColumn104.Hidden = true;
			ultraGridColumn104.Width = 8;
			ultraGridColumn105.Header.VisiblePosition = 87;
			ultraGridColumn105.Hidden = true;
			ultraGridColumn105.Width = 8;
			ultraGridColumn106.Header.VisiblePosition = 88;
			ultraGridColumn106.Hidden = true;
			ultraGridColumn106.Width = 8;
			ultraGridColumn107.Header.VisiblePosition = 89;
			ultraGridColumn107.Hidden = true;
			ultraGridColumn107.Width = 8;
			ultraGridColumn108.Header.VisiblePosition = 90;
			ultraGridColumn108.Hidden = true;
			ultraGridColumn108.Width = 8;
			ultraGridColumn109.Header.VisiblePosition = 91;
			ultraGridColumn109.Hidden = true;
			ultraGridColumn109.Width = 8;
			ultraGridColumn110.Header.VisiblePosition = 92;
			ultraGridColumn110.Hidden = true;
			ultraGridColumn110.Width = 8;
			ultraGridColumn111.Header.VisiblePosition = 93;
			ultraGridColumn111.Hidden = true;
			ultraGridColumn111.Width = 8;
			ultraGridColumn112.Header.VisiblePosition = 94;
			ultraGridColumn112.Hidden = true;
			ultraGridColumn112.Width = 16;
			ultraGridColumn113.Header.VisiblePosition = 95;
			ultraGridColumn113.Width = 17;
			ultraGridColumn114.Header.VisiblePosition = 96;
			ultraGridColumn114.Width = 14;
			ultraGridColumn115.Header.VisiblePosition = 97;
			ultraGridColumn115.Width = 14;
			ultraGridColumn116.Header.VisiblePosition = 98;
			ultraGridColumn116.Width = 14;
			ultraGridColumn117.Header.VisiblePosition = 99;
			ultraGridColumn117.Width = 27;
			ultraGridColumn118.Header.VisiblePosition = 100;
			ultraGridColumn118.Width = 26;
			ultraGridColumn119.Header.VisiblePosition = 101;
			ultraGridColumn119.Width = 25;
			ultraGridColumn120.Header.VisiblePosition = 102;
			ultraGridColumn120.Width = 24;
			ultraGridColumn121.Header.VisiblePosition = 103;
			ultraGridColumn121.Width = 19;
			ultraGridColumn122.Header.VisiblePosition = 104;
			ultraGridColumn122.Width = 18;
			ultraGridBand3.Columns.AddRange(new object[] {
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
																										 ultraGridColumn109,
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
																										 ultraGridColumn122});
			this.cmbPersonales.DisplayLayout.BandsSerializer.Add(ultraGridBand3);
			this.cmbPersonales.DisplayMember = "Nombre";
			this.cmbPersonales.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbPersonales.Location = new System.Drawing.Point(40, 181);
			this.cmbPersonales.Name = "cmbPersonales";
			this.cmbPersonales.Size = new System.Drawing.Size(280, 86);
			this.cmbPersonales.TabIndex = 34;
			this.cmbPersonales.Text = "ultraDropDown1";
			this.cmbPersonales.ValueMember = "idPersonal";
			this.cmbPersonales.Visible = false;
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
			// 
			// btSubida
			// 
			this.btSubida.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btSubida.Location = new System.Drawing.Point(504, 405);
			this.btSubida.Name = "btSubida";
			this.btSubida.Size = new System.Drawing.Size(75, 25);
			this.btSubida.TabIndex = 35;
			this.btSubida.Text = "Genera Rol";
			this.btSubida.Click += new System.EventHandler(this.btSubida_Click);
			// 
			// label1
			// 
			this.label1.AutoSize = true;
			this.label1.Location = new System.Drawing.Point(312, 26);
			this.label1.Name = "label1";
			this.label1.Size = new System.Drawing.Size(34, 16);
			this.label1.TabIndex = 36;
			this.label1.Text = "Fecha";
			// 
			// label2
			// 
			this.label2.AutoSize = true;
			this.label2.Location = new System.Drawing.Point(16, 26);
			this.label2.Name = "label2";
			this.label2.Size = new System.Drawing.Size(57, 16);
			this.label2.TabIndex = 37;
			this.label2.Text = "Supervisor";
			// 
			// chkResumen
			// 
			this.chkResumen.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.chkResumen.Location = new System.Drawing.Point(344, 405);
			this.chkResumen.Name = "chkResumen";
			this.chkResumen.Size = new System.Drawing.Size(72, 21);
			this.chkResumen.TabIndex = 38;
			this.chkResumen.Text = "Resumen";
			// 
			// button1
			// 
			this.button1.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			this.button1.Location = new System.Drawing.Point(696, 26);
			this.button1.Name = "button1";
			this.button1.Size = new System.Drawing.Size(72, 26);
			this.button1.TabIndex = 41;
			this.button1.Text = "Calcular";
			this.button1.Click += new System.EventHandler(this.button1_Click);
			// 
			// button2
			// 
			this.button2.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			this.button2.Location = new System.Drawing.Point(776, 26);
			this.button2.Name = "button2";
			this.button2.Size = new System.Drawing.Size(72, 26);
			this.button2.TabIndex = 45;
			this.button2.Text = "Turnos";
			this.button2.Click += new System.EventHandler(this.button2_Click);
			// 
			// button3
			// 
			this.button3.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.button3.Location = new System.Drawing.Point(456, 405);
			this.button3.Name = "button3";
			this.button3.Size = new System.Drawing.Size(24, 26);
			this.button3.TabIndex = 46;
			this.button3.Text = "RI";
			this.button3.Click += new System.EventHandler(this.button4_Click);
			this.button3.MouseUp += new System.Windows.Forms.MouseEventHandler(this.button3_MouseUp);
			// 
			// chkRevisado
			// 
			this.chkRevisado.DataBindings.Add(new System.Windows.Forms.Binding("CheckedValue", this.cdsRolAdicional, "RolAdicionales.Revisado"));
			this.chkRevisado.Location = new System.Drawing.Point(592, 26);
			this.chkRevisado.Name = "chkRevisado";
			this.chkRevisado.Size = new System.Drawing.Size(72, 21);
			this.chkRevisado.TabIndex = 47;
			this.chkRevisado.Text = "Revisado";
			// 
			// cdsSeguridad
			// 
			this.cdsSeguridad.BindingContextCtrl = this;
			this.cdsSeguridad.DataLibrary = "LibContabilidad";
			this.cdsSeguridad.DataLibraryUrl = "";
			this.cdsSeguridad.DataSetDef = "dsSeguridad";
			this.cdsSeguridad.Locale = new System.Globalization.CultureInfo("es-EC");
			this.cdsSeguridad.Name = "cdsSeguridad";
			this.cdsSeguridad.SchemaClassName = "LibContabilidad.DataClass";
			this.cdsSeguridad.SchemaDef = null;
			// 
			// RolAdicional
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 14);
			this.ClientSize = new System.Drawing.Size(872, 453);
			this.Controls.Add(this.chkRevisado);
			this.Controls.Add(this.button3);
			this.Controls.Add(this.button2);
			this.Controls.Add(this.button1);
			this.Controls.Add(this.chkResumen);
			this.Controls.Add(this.label2);
			this.Controls.Add(this.label1);
			this.Controls.Add(this.btSubida);
			this.Controls.Add(this.cmbPersonales);
			this.Controls.Add(this.txtIdRolAdicional);
			this.Controls.Add(this.btCarga);
			this.Controls.Add(this.cmbPersonal);
			this.Controls.Add(this.cmbFecha);
			this.Controls.Add(this.barraDato1);
			this.Controls.Add(this.grdRetRenta);
			this.Name = "RolAdicional";
			this.Text = "Valores Adiciones Adicional";
			this.Load += new System.EventHandler(this.RolAdicional_Load);
			((System.ComponentModel.ISupportInitialize)(this.cmbFecha)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsRolAdicional)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.grdRetRenta)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbPersonal)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIdRolAdicional)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbPersonales)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsPersonal)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeguridad)).EndInit();
			this.ResumeLayout(false);

		}
		#endregion

		private void barraDato1_Refresca(object sender, System.EventArgs e)
		{
			drRolAdicional = cdsRolAdicional.TableViews["RolAdicionales"].Rows[0];

		}
//		private bool bUsuario = false;
		private C1.Data.C1DataRow drRolAdicional;
		private bool bFiltroSucursal = false;
		private Acceso miAcceso;

		private void RolAdicional_Load(object sender, System.EventArgs e)
		{
			txtIdRolAdicional.Width = 0;
			#region Seguridad
			miAcceso =new Acceso("0516", cdsSeguridad.StorageDataSet.Tables["Seguridad"]);
	
	
		#endregion
			#region Filtro de Sucursales
		bFiltroSucursal = Funcion.bEjecutaSQL(cdsRolAdicional, 	"Select IsNull(FiltroSucursal, 0) From SeteoRol");
			if (bFiltroSucursal)
			{
				string iSuc = Funcion.sEscalarSQL(cdsRolAdicional, " Exec RolAdicionalUsuario ");
				if (iSuc != "")
				{
					barraDato1.stFiltroExt = " And  Roladicionales.idPersonal in ( " + iSuc.ToString()+ " )";
					string procedimiento = "Select idPersonal, Nombre From Personal  Where bSupervisor = 1 and IsNull(idpersonal,0) in ("+iSuc.ToString()+")";
					cmbPersonal.DataSource = Funcion.dvProcedimiento(cdsRolAdicional, procedimiento);
					cmbPersonales.DataSource = Funcion.dvProcedimiento(cdsRolAdicional, 
						"Select idPersonal, Nombre From Personal Where IsNull(idpersonal,'') in ("+iSuc.ToString()+")");
				}
			}
			#endregion Filtro de Sucursales

			if (IdRolAdicional > 0)
			{
				BuscarAdicional(IdRolAdicional);
			}
			else
			{
				this.barraDato1.PosUltima();
			}
			barraDato1.HabilitaControles(false);
			//barraDato1.ProximoId(4);
		}

		private void cdsRolAdicional_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			cdsRolAdicional.Schema.Connections[0].ConnectionString = miFuncion.AccesoBase(MenuLatinium.stDirRoles);
		}

		private void barraDato1_Borrar(object sender, System.EventArgs e)
		{
		
			if (drRolAdicional == null)
			{
				MessageBox.Show("No existe un registro para borrar", "Información",
					MessageBoxButtons.OK, MessageBoxIcon.Information);
				return;
			}
			if (DialogResult.No == MessageBox.Show("¿Está seguro que desea borrar el Registro?"
				+ ".\nEste proceso no puede ser revertido",	"Confirmación", MessageBoxButtons.YesNo, MessageBoxIcon.Stop, 
				MessageBoxDefaultButton.Button2)) return;

			barraDato1.BorraRegistro();
		}

		private void barraDato1_Crear(object sender, System.EventArgs e)
		{
			if (!miAcceso.Nuevo)
			{
				MessageBox.Show("No tiene Acceso a Crear " + Text.Trim(),
					"Seguridad", MessageBoxButtons.OK, MessageBoxIcon.Information);
				return;
			}
			barraDato1.CrearRegistro();		
		}

		private void barraDato1_Deshacer(object sender, System.EventArgs e)
		{
			barraDato1.DeshacerRegistro();
		}

		private void barraDato1_Editar(object sender, System.EventArgs e)
		{		
			if (!miAcceso.Editar)
			{
				MessageBox.Show("No tiene Acceso a Editar " + Text.Trim(),
					"Seguridad", MessageBoxButtons.OK, MessageBoxIcon.Information);
				return;
			}
			
			if (!miAcceso.EditarRevisadas && chkRevisado.Checked)
			{
				MessageBox.Show("No puede Editar después de Revisado", "Seguridad", 
					MessageBoxButtons.OK, MessageBoxIcon.Information);
				return;
			}
			barraDato1.EditarRegistro();
		}

		private void barraDato1_Grabar(object sender, System.EventArgs e)
		{
			barraDato1.GrabaRegistro();
		}

		CultureInfo us = new CultureInfo("en-US");
		private void barraDato1_Impresora(object sender, System.EventArgs e)
		{
		}

		private void barraDato1_Imprimir(object sender, System.EventArgs e)
		{
			Cursor = Cursors.WaitCursor;
			DateTime dtFecha = (DateTime) cmbFecha.Value;
			DateTime dtFechaIni = DateTime.Parse(dtFecha.ToString("MM/01/yyyy"), us);
			DateTime dtFechaFin = DateTime.Parse(dtFecha.AddMonths(1).ToString("MM/01/yyyy", us));

			string stFiltro = "{RolAdicionales.Fecha} >= #" + dtFechaIni.ToString("MM/dd/yyyy", us) 
				+ "# And {RolAdicionales.Fecha} >= #" + dtFechaFin.ToString("MM/dd/yyyy", us) + "#";
			Reporte miRep = new Reporte("RolAdicional.rpt", stFiltro);
			miRep.Titulo("Reporte horas extras de " + dtFecha.ToString("MMMM del yyyy"));
			if (chkResumen.Checked) miRep.Resumen();
			miRep.MdiParent = MdiParent;
			miRep.Show();
			Cursor = Cursors.Default;
		}

		private void barraDato1_Restaurar(object sender, System.EventArgs e)
		{
		
		}

		private void btCarga_Click(object sender, System.EventArgs e)
		{
			string stExec = "Exec RolAdicionalCarga " + txtIdRolAdicional.Text;
			Funcion.EjecutaSQL(cdsRolAdicional, stExec);
			barraDato1.ProximoId(5);
		}

		private void btSubida_Click(object sender, System.EventArgs e)
		{
			string stExec = "Exec RolAdicionalVerificaSubida " + txtIdRolAdicional.Text;
			string stMensaje = Funcion.sEscalarSQL(cdsRolAdicional, stExec);
			if (stMensaje.Length > 0)
			{
				if (stMensaje.StartsWith("Precauci"))
				{
					if (DialogResult.No == MessageBox.Show(stMensaje, "Confirmación",
						MessageBoxButtons.YesNo, MessageBoxIcon.Question))
						return;
				}
				else
				{
					MessageBox.Show(stMensaje, "Información");
					return;
				}
			}
			stExec = "Exec RolAdicionalSubida " + txtIdRolAdicional.Text;
			Funcion.EjecutaSQL(cdsRolAdicional, stExec);
		}

		private void button1_Click(object sender, System.EventArgs e)
		{
			DateTime dtFecha = DateTime.Parse(cmbFecha.Value.ToString());
			Funcion.EjecutaSQL(cdsRolAdicional, "Exec RolAdicionalCalcula "+ txtIdRolAdicional.Text
				+",'"+ dtFecha.ToString("yyyyMMdd")+"'");
			barraDato1.ProximoId(5);
		}

		private void button4_Click(object sender, System.EventArgs e)
		{
			string stFiltro="";
			Cursor = Cursors.WaitCursor;
			DateTime dtFecha = (DateTime) cmbFecha.Value;
			DateTime dtFechaIni = DateTime.Parse(dtFecha.AddMonths(-1).ToString("MM/27/yyyy"), us);
			DateTime dtFechaFin = DateTime.Parse(dtFecha.ToString("MM/27/yyyy"), us);

			stFiltro = "{RolAdicionales.Fecha} >= #" + dtFechaIni.ToString("MM/dd/yyyy", us) 
				+ "# And {RolAdicionales.Fecha} >= #" + dtFechaFin.ToString("MM/dd/yyyy", us) + "#";
		
			if (grdRetRenta.ActiveRow == null)
			{
				MessageBox.Show("Seleccione Fila de la cual desea ver la Reporte", "Información");
				return;
			}
			int IdPersonal = (int) grdRetRenta.ActiveRow.Cells["idPersonal"].Value;
			if (IdPersonal > 0)
			{
				stFiltro = stFiltro + " and {Personal.idPersonal} = " + IdPersonal.ToString();			
			}
			string stReporte = "RolAdicionalInd.rpt";
			Reporte miReporte = new Reporte(stReporte, stFiltro);
			miReporte.Show();
			Cursor = Cursors.Default;
		}
		private void BuscarAdicional(int idPersona)
		{
			if (idPersona > 0) 
			{
				this.barraDato1.IdRegistro = idPersona;
				this.barraDato1.ProximoId(5);
			}
		}
		private void button2_Click(object sender, System.EventArgs e)
		{
		Cursor = Cursors.WaitCursor;
			RolAdicionalTurno miTurno = new RolAdicionalTurno();
			miTurno.MdiParent = MdiParent;
			miTurno.Show();

		Cursor = Cursors.Default;
		}

		private void button3_MouseUp(object sender, System.Windows.Forms.MouseEventArgs e)
		{
			if (e.Button != MouseButtons.Right) return;
			string stFiltro="";
			Cursor = Cursors.WaitCursor;
			DateTime dtFecha = (DateTime) cmbFecha.Value;
			DateTime dtFechaIni = DateTime.Parse(dtFecha.AddMonths(-1).ToString("MM/27/yyyy"), us);
			DateTime dtFechaFin = DateTime.Parse(dtFecha.ToString("MM/27/yyyy"), us);

			stFiltro = "{RolAdicionales.Fecha} >= #" + dtFechaIni.ToString("MM/dd/yyyy", us) 
				+ "# And {RolAdicionales.Fecha} <= #" + dtFechaFin.ToString("MM/dd/yyyy", us) + "#";
		
			string stReporte = "RolAdicionalInd.rpt";
			Reporte miReporte = new Reporte(stReporte, stFiltro);
			miReporte.Show();
			Cursor = Cursors.Default;


		}

		private void grdRetRenta_InitializeLayout(object sender, Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs e)
		{
		
		}

		private void barraDato1_Borrar_1(object sender, System.EventArgs e)
		{
			if (!miAcceso.Eliminar)
			{
				MessageBox.Show("No tiene Acceso a Eliminar " + Text.Trim(),
					"Seguridad", MessageBoxButtons.OK, MessageBoxIcon.Information);
				return;
			}
			barraDato1.BorraRegistro();		
		}

		private void barraDato1_Load(object sender, System.EventArgs e)
		{
		
		}


	}
}
