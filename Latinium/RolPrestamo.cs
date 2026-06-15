using System;
using System.Globalization;
using System.Data.SqlClient;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using Infragistics.Win.UltraWinGrid;

namespace Latinium
{
	/// <summary>
	/// Descripción breve de RolPrestamo.
	/// </summary>
	public class RolPrestamo : DevExpress.XtraEditors.XtraForm
	{
		private C1.Data.C1DataSet cdsRolPrestamo;
		//private C1.Data.FillEventArgs ef;
		private C1.Data.C1DataView cdvRolPrestamo;
		private C1.Data.C1DataSet cdsPersonal;
		private C1.Data.C1DataView cdvPersonal;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtIdRolPrestamo;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtIdAsiento;
		private Infragistics.Win.Misc.UltraButton btEgreso;
		private System.Windows.Forms.Label label5;
		private Infragistics.Win.UltraWinSchedule.UltraCalendarCombo cmbFechaCobro;
		private System.Windows.Forms.Label label4;
		private Infragistics.Win.UltraWinSchedule.UltraCalendarCombo cmbFechaEntrega;
		private System.Windows.Forms.Label label3;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtValor;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbNombre;
		private System.Windows.Forms.Label label2;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtConcepto;
		private System.Windows.Forms.Label label1;
		private System.Windows.Forms.Label label6;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtNumCuotas;
		private System.Windows.Forms.ErrorProvider errorProvider;
		private Infragistics.Win.Misc.UltraButton DetallePrestamo;
		private System.Windows.Forms.ToolTip toolTip;
		private Infragistics.Win.UltraWinGrid.UltraGrid ultraGrid1;
		private C1.Data.C1DataSet cdsAsiento;
		private C1.Data.C1DataSet cdsRolTipoPrest;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbPrestamo;
		private System.Windows.Forms.Label label7;
		private Infragistics.Win.Misc.UltraButton btTipoPrestamo;
		private BarraDatoSQL.BarraDatoSQL barraDato1;
		private Infragistics.Win.Misc.UltraButton btSolicitud;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtNumero;
		private System.Windows.Forms.Label label8;
		private System.Windows.Forms.GroupBox groupBox1;
		private Infragistics.Win.Misc.UltraButton btAceptar;
		private Infragistics.Win.Misc.UltraButton btCancelar;
		private System.ComponentModel.IContainer components;

		bool BCruceAnticipo = false;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtEstado;
		double DValor = 0;

		public RolPrestamo(bool bCruceAnticipo, double dValor)
		{
			//
			// Necesario para admitir el Diseñador de Windows Forms
			//
			InitializeComponent();
			BCruceAnticipo = bCruceAnticipo;
			DValor = dValor;
			//
			// TODO: agregar código de constructor después de llamar a InitializeComponent
			//
		}

		public RolPrestamo()
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
			this.components = new System.ComponentModel.Container();
			Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton dateButton1 = new Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton();
			Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton dateButton2 = new Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton();
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
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn11 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Jornada");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn12 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Direccion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn13 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Telefono");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn14 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Celular");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn15 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Cedula");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn16 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Afiliacion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn17 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Reserva");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn18 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Renta");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn19 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Iess");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn20 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Sueldo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn21 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Comision");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn22 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Tipo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn23 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("FechaIngreso");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn24 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("FechaNacimiento");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn25 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("FechaSalida");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn26 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Notas");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn27 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("CargasFamiliares");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn28 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idProvincia");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn29 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCiudad");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn30 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Numero");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn31 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("SistemaSalario");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn32 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Cuenta");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn33 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("CodigoSectorial");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn34 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Condicion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn35 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("FechaNovedad");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn36 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Acceso");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn37 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Vendedor");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn38 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Borrar");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn39 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Cobrador");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn40 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Sexo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn41 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Estado_Civil");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn42 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Banco");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn43 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("CuentaCorr");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn44 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idTipoCuenta");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn45 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("CedulaMilitar");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn46 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idFormaPago");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn47 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Quincena");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn48 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("CodIngreso");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn49 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("CodTelefono");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn50 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("TarjEdificio");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn51 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("TarjetaOficina");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn52 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("email");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn53 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idProyecto");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn54 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCuenta");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn55 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Bono");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn56 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Bono1");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn57 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Desc1");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn58 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Desc2");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn59 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("TieneQuincena");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn60 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Convenio");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn61 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idSubProyecto");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn62 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Sindicato");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn63 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Asociacion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn64 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idRolHorario");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn65 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("FechaContrato");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn66 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Educacion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn67 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("NumeroDias");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn68 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idRelacionT");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn69 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("gSalud");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn70 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("gEducacion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn71 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("gVivienda");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn72 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("gVestimenta");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn73 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("gAlimentacion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn74 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("gTotal");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn75 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("BaseComision");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn76 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("PorcentajeComision");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn77 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("gDesahucio");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn78 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("gDiscapacidad");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn79 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("IngGravOE");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn80 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("DeduccionOE");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn81 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("OtrasRebOE");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn82 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("ImpRentaAsumido");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn83 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("gRebajaIIIEdad");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn84 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("RentaRetenidaOE");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn85 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("RentaRetOE");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn86 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("TipoSangre");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn87 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("RegEntrada");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn88 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("RegSalida");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn89 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idTipoContrato");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn90 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("CodVendedor");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn91 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Clave");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn92 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idSupervisor");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn93 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("bSupervisor");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn94 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("E_Profesor");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn95 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Dias360");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn96 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("DiasXIII");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn97 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idJefe");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn98 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("PoseeDiscapacidad");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn99 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("PorcDiscapacidad");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn100 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("TipoDiscapacidad");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn101 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("NumConadis");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn102 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("TipoVivienda");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn103 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Vehiculo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn104 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("CodigoUsuario");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn105 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Estado");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn106 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("FechaReIngreso");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn107 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("GTC");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn108 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Codigo");
			Infragistics.Win.Appearance appearance1 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand2 = new Infragistics.Win.UltraWinGrid.UltraGridBand("_RolPrestamoDetalle", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn109 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idRolPrestDet");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn110 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idSucursal");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn111 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("FechaPropuesta");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn112 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("FechaCobro");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn113 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Capital");
			Infragistics.Win.Appearance appearance2 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn114 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Interes");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn115 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("CapCobrado");
			Infragistics.Win.Appearance appearance3 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn116 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("IntCobrado");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn117 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idRolPrestamo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn118 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idTipoPrestamo");
			Infragistics.Win.Appearance appearance4 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance5 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance6 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance7 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand3 = new Infragistics.Win.UltraWinGrid.UltraGridBand("RolTipoPrestamo", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn119 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idRolTipoPrest");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn120 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("TipoPrestamo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn121 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Campo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn122 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCuenta");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn123 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idRolGrupo");
			this.cmbFechaCobro = new Infragistics.Win.UltraWinSchedule.UltraCalendarCombo();
			this.cdsRolPrestamo = new C1.Data.C1DataSet();
			this.cdvRolPrestamo = new C1.Data.C1DataView();
			this.cdsPersonal = new C1.Data.C1DataSet();
			this.cdvPersonal = new C1.Data.C1DataView();
			this.txtIdRolPrestamo = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.txtIdAsiento = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.btEgreso = new Infragistics.Win.Misc.UltraButton();
			this.label5 = new System.Windows.Forms.Label();
			this.label4 = new System.Windows.Forms.Label();
			this.cmbFechaEntrega = new Infragistics.Win.UltraWinSchedule.UltraCalendarCombo();
			this.label3 = new System.Windows.Forms.Label();
			this.txtValor = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.cmbNombre = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.label2 = new System.Windows.Forms.Label();
			this.txtConcepto = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.label1 = new System.Windows.Forms.Label();
			this.txtNumCuotas = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.label6 = new System.Windows.Forms.Label();
			this.errorProvider = new System.Windows.Forms.ErrorProvider();
			this.DetallePrestamo = new Infragistics.Win.Misc.UltraButton();
			this.toolTip = new System.Windows.Forms.ToolTip(this.components);
			this.btSolicitud = new Infragistics.Win.Misc.UltraButton();
			this.ultraGrid1 = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.cdsAsiento = new C1.Data.C1DataSet();
			this.cmbPrestamo = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.cdsRolTipoPrest = new C1.Data.C1DataSet();
			this.label7 = new System.Windows.Forms.Label();
			this.btTipoPrestamo = new Infragistics.Win.Misc.UltraButton();
			this.barraDato1 = new BarraDatoSQL.BarraDatoSQL();
			this.txtNumero = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.label8 = new System.Windows.Forms.Label();
			this.groupBox1 = new System.Windows.Forms.GroupBox();
			this.btAceptar = new Infragistics.Win.Misc.UltraButton();
			this.btCancelar = new Infragistics.Win.Misc.UltraButton();
			this.txtEstado = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			((System.ComponentModel.ISupportInitialize)(this.cmbFechaCobro)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsRolPrestamo)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdvRolPrestamo)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsPersonal)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdvPersonal)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIdRolPrestamo)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIdAsiento)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbFechaEntrega)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtValor)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbNombre)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtConcepto)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNumCuotas)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraGrid1)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsAsiento)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbPrestamo)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsRolTipoPrest)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNumero)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtEstado)).BeginInit();
			this.SuspendLayout();
			// 
			// cmbFechaCobro
			// 
			this.cmbFechaCobro.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			this.cmbFechaCobro.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsRolPrestamo, "_RolPrestamo.FechaCobro"));
			dateButton1.Caption = "Today";
			this.cmbFechaCobro.DateButtons.Add(dateButton1);
			this.cmbFechaCobro.Format = "dd/MM/yyyy";
			this.cmbFechaCobro.Location = new System.Drawing.Point(320, 72);
			this.cmbFechaCobro.Name = "cmbFechaCobro";
			this.cmbFechaCobro.NonAutoSizeHeight = 23;
			this.cmbFechaCobro.Size = new System.Drawing.Size(112, 21);
			this.cmbFechaCobro.SpinButtonsVisible = true;
			this.cmbFechaCobro.TabIndex = 24;
			this.cmbFechaCobro.Value = new System.DateTime(2006, 1, 14, 0, 0, 0, 0);
			// 
			// cdsRolPrestamo
			// 
			this.cdsRolPrestamo.BindingContextCtrl = this;
			this.cdsRolPrestamo.DataLibrary = "LibPersonal";
			this.cdsRolPrestamo.DataLibraryUrl = "";
			this.cdsRolPrestamo.DataSetDef = "dsRolGrupo";
			this.cdsRolPrestamo.Locale = new System.Globalization.CultureInfo("es-EC");
			this.cdsRolPrestamo.Name = "cdsRolPrestamo";
			this.cdsRolPrestamo.SchemaClassName = "LibPersonal.DataClass";
			this.cdsRolPrestamo.SchemaDef = null;
			this.cdsRolPrestamo.BeforeFill += new C1.Data.FillEventHandler(this.cdsRolPrestamo_BeforeFill);
			// 
			// cdvRolPrestamo
			// 
			this.cdvRolPrestamo.BindingContextCtrl = this;
			this.cdvRolPrestamo.DataSet = this.cdsRolPrestamo;
			this.cdvRolPrestamo.Sort = "idRolPrestamo";
			this.cdvRolPrestamo.TableName = "";
			this.cdvRolPrestamo.TableViewName = "RolPrestamo";
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
			// cdvPersonal
			// 
			this.cdvPersonal.BindingContextCtrl = this;
			this.cdvPersonal.DataSet = this.cdsPersonal;
			this.cdvPersonal.Sort = "Nombre";
			this.cdvPersonal.TableName = "";
			this.cdvPersonal.TableViewName = "Personal";
			// 
			// txtIdRolPrestamo
			// 
			this.txtIdRolPrestamo.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsRolPrestamo, "_RolPrestamo.idRolPrestamo"));
			this.txtIdRolPrestamo.Location = new System.Drawing.Point(344, 400);
			this.txtIdRolPrestamo.Name = "txtIdRolPrestamo";
			this.txtIdRolPrestamo.Size = new System.Drawing.Size(40, 22);
			this.txtIdRolPrestamo.TabIndex = 28;
			// 
			// txtIdAsiento
			// 
			this.txtIdAsiento.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsRolPrestamo, "_RolPrestamo.idAsiento"));
			this.txtIdAsiento.Location = new System.Drawing.Point(312, 400);
			this.txtIdAsiento.Name = "txtIdAsiento";
			this.txtIdAsiento.Size = new System.Drawing.Size(40, 22);
			this.txtIdAsiento.TabIndex = 27;
			// 
			// btEgreso
			// 
			this.btEgreso.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btEgreso.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			this.btEgreso.Location = new System.Drawing.Point(424, 400);
			this.btEgreso.Name = "btEgreso";
			this.btEgreso.Size = new System.Drawing.Size(24, 25);
			this.btEgreso.TabIndex = 26;
			this.btEgreso.Text = "A";
			this.toolTip.SetToolTip(this.btEgreso, "Generacion de Egreso");
			this.btEgreso.Click += new System.EventHandler(this.btEgreso_Click);
			this.btEgreso.MouseUp += new System.Windows.Forms.MouseEventHandler(this.btEgreso_MouseDown);
			// 
			// label5
			// 
			this.label5.AutoSize = true;
			this.label5.Location = new System.Drawing.Point(232, 72);
			this.label5.Name = "label5";
			this.label5.Size = new System.Drawing.Size(34, 16);
			this.label5.TabIndex = 25;
			this.label5.Text = "Cobro";
			this.label5.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// label4
			// 
			this.label4.AutoSize = true;
			this.label4.Location = new System.Drawing.Point(8, 72);
			this.label4.Name = "label4";
			this.label4.Size = new System.Drawing.Size(46, 16);
			this.label4.TabIndex = 23;
			this.label4.Text = "Solicitud";
			this.label4.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// cmbFechaEntrega
			// 
			this.cmbFechaEntrega.AllowMonthSelection = true;
			this.cmbFechaEntrega.AutoSelectionUpdate = true;
			this.cmbFechaEntrega.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			this.cmbFechaEntrega.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsRolPrestamo, "_RolPrestamo.FechaEntrega"));
			dateButton2.Caption = "Today";
			this.cmbFechaEntrega.DateButtons.Add(dateButton2);
			this.cmbFechaEntrega.DayDisplayStyle = Infragistics.Win.UltraWinSchedule.DayDisplayStyle.DayNumberAndImage;
			this.cmbFechaEntrega.DayOfWeekDisplayStyle = Infragistics.Win.UltraWinSchedule.DayOfWeekDisplayStyle.FirstRow;
			this.cmbFechaEntrega.Format = "dd/MM/yyyy";
			this.cmbFechaEntrega.Location = new System.Drawing.Point(104, 72);
			this.cmbFechaEntrega.MonthOrientation = Infragistics.Win.UltraWinSchedule.MonthOrientation.DownThenAcross;
			this.cmbFechaEntrega.Name = "cmbFechaEntrega";
			this.cmbFechaEntrega.NonAutoSizeHeight = 23;
			this.cmbFechaEntrega.Size = new System.Drawing.Size(112, 21);
			this.cmbFechaEntrega.SpinButtonsVisible = true;
			this.cmbFechaEntrega.TabIndex = 22;
			this.cmbFechaEntrega.Value = new System.DateTime(2006, 1, 14, 0, 0, 0, 0);
			// 
			// label3
			// 
			this.label3.AutoSize = true;
			this.label3.Location = new System.Drawing.Point(8, 104);
			this.label3.Name = "label3";
			this.label3.Size = new System.Drawing.Size(30, 16);
			this.label3.TabIndex = 21;
			this.label3.Text = "Valor";
			this.label3.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// txtValor
			// 
			this.txtValor.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsRolPrestamo, "_RolPrestamo.Capital"));
			this.txtValor.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtValor.FormatString = "#,##0.00";
			this.txtValor.Location = new System.Drawing.Point(104, 104);
			this.txtValor.Name = "txtValor";
			this.txtValor.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtValor.PromptChar = ' ';
			this.txtValor.Size = new System.Drawing.Size(112, 22);
			this.txtValor.SpinButtonDisplayStyle = Infragistics.Win.ButtonDisplayStyle.Always;
			this.txtValor.TabIndex = 20;
			this.txtValor.Click += new System.EventHandler(this.txtValor_Click);
			this.txtValor.Enter += new System.EventHandler(this.txtValor_Enter);
			// 
			// cmbNombre
			// 
			this.cmbNombre.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbNombre.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbNombre.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsRolPrestamo, "_RolPrestamo.idPersonal"));
			this.cmbNombre.DataSource = this.cdvPersonal;
			this.cmbNombre.DisplayLayout.AutoFitColumns = true;
			ultraGridColumn1.Header.VisiblePosition = 21;
			ultraGridColumn1.Hidden = true;
			ultraGridColumn2.Header.VisiblePosition = 59;
			ultraGridColumn2.Hidden = true;
			ultraGridColumn2.Width = 49;
			ultraGridColumn3.Header.VisiblePosition = 19;
			ultraGridColumn3.Hidden = true;
			ultraGridColumn4.Header.VisiblePosition = 20;
			ultraGridColumn4.Hidden = true;
			ultraGridColumn5.Header.VisiblePosition = 18;
			ultraGridColumn5.Hidden = true;
			ultraGridColumn6.Header.VisiblePosition = 22;
			ultraGridColumn6.Hidden = true;
			ultraGridColumn7.Header.VisiblePosition = 37;
			ultraGridColumn7.Hidden = true;
			ultraGridColumn8.Header.VisiblePosition = 25;
			ultraGridColumn8.Hidden = true;
			ultraGridColumn9.Header.VisiblePosition = 0;
			ultraGridColumn9.Width = 148;
			ultraGridColumn10.Header.VisiblePosition = 1;
			ultraGridColumn10.Hidden = true;
			ultraGridColumn11.Header.VisiblePosition = 24;
			ultraGridColumn11.Hidden = true;
			ultraGridColumn12.Header.VisiblePosition = 13;
			ultraGridColumn12.Hidden = true;
			ultraGridColumn13.Header.VisiblePosition = 32;
			ultraGridColumn13.Hidden = true;
			ultraGridColumn14.Header.VisiblePosition = 7;
			ultraGridColumn14.Hidden = true;
			ultraGridColumn15.Header.VisiblePosition = 6;
			ultraGridColumn15.Hidden = true;
			ultraGridColumn16.Header.VisiblePosition = 3;
			ultraGridColumn16.Hidden = true;
			ultraGridColumn17.Header.VisiblePosition = 29;
			ultraGridColumn17.Hidden = true;
			ultraGridColumn18.Header.VisiblePosition = 28;
			ultraGridColumn18.Hidden = true;
			ultraGridColumn19.Header.VisiblePosition = 23;
			ultraGridColumn19.Hidden = true;
			ultraGridColumn20.Header.VisiblePosition = 31;
			ultraGridColumn20.Hidden = true;
			ultraGridColumn21.Header.VisiblePosition = 10;
			ultraGridColumn21.Hidden = true;
			ultraGridColumn22.Header.VisiblePosition = 33;
			ultraGridColumn22.Hidden = true;
			ultraGridColumn23.Header.VisiblePosition = 14;
			ultraGridColumn23.Hidden = true;
			ultraGridColumn24.Header.VisiblePosition = 15;
			ultraGridColumn24.Hidden = true;
			ultraGridColumn25.Header.VisiblePosition = 17;
			ultraGridColumn25.Hidden = true;
			ultraGridColumn26.Header.VisiblePosition = 26;
			ultraGridColumn26.Hidden = true;
			ultraGridColumn27.Header.VisiblePosition = 5;
			ultraGridColumn27.Hidden = true;
			ultraGridColumn28.Header.VisiblePosition = 36;
			ultraGridColumn28.Hidden = true;
			ultraGridColumn29.Header.VisiblePosition = 38;
			ultraGridColumn29.Hidden = true;
			ultraGridColumn30.Header.VisiblePosition = 27;
			ultraGridColumn30.Hidden = true;
			ultraGridColumn31.Header.VisiblePosition = 30;
			ultraGridColumn31.Hidden = true;
			ultraGridColumn32.Header.VisiblePosition = 12;
			ultraGridColumn32.Hidden = true;
			ultraGridColumn33.Header.VisiblePosition = 9;
			ultraGridColumn33.Hidden = true;
			ultraGridColumn34.Header.VisiblePosition = 11;
			ultraGridColumn34.Hidden = true;
			ultraGridColumn35.Header.VisiblePosition = 16;
			ultraGridColumn35.Hidden = true;
			ultraGridColumn36.Header.VisiblePosition = 2;
			ultraGridColumn36.Hidden = true;
			ultraGridColumn37.Header.VisiblePosition = 34;
			ultraGridColumn37.Hidden = true;
			ultraGridColumn38.Header.VisiblePosition = 4;
			ultraGridColumn38.Hidden = true;
			ultraGridColumn39.Header.VisiblePosition = 8;
			ultraGridColumn39.Hidden = true;
			ultraGridColumn40.Header.VisiblePosition = 39;
			ultraGridColumn40.Hidden = true;
			ultraGridColumn40.Width = 11;
			ultraGridColumn41.Header.VisiblePosition = 40;
			ultraGridColumn41.Hidden = true;
			ultraGridColumn41.Width = 19;
			ultraGridColumn42.Header.VisiblePosition = 41;
			ultraGridColumn42.Hidden = true;
			ultraGridColumn42.Width = 19;
			ultraGridColumn43.Header.VisiblePosition = 42;
			ultraGridColumn43.Hidden = true;
			ultraGridColumn43.Width = 19;
			ultraGridColumn44.Header.VisiblePosition = 43;
			ultraGridColumn44.Hidden = true;
			ultraGridColumn44.Width = 16;
			ultraGridColumn45.Header.VisiblePosition = 44;
			ultraGridColumn45.Hidden = true;
			ultraGridColumn45.Width = 19;
			ultraGridColumn46.Header.VisiblePosition = 45;
			ultraGridColumn46.Hidden = true;
			ultraGridColumn46.Width = 16;
			ultraGridColumn47.Header.VisiblePosition = 46;
			ultraGridColumn47.Hidden = true;
			ultraGridColumn47.Width = 15;
			ultraGridColumn48.Header.VisiblePosition = 47;
			ultraGridColumn48.Hidden = true;
			ultraGridColumn48.Width = 19;
			ultraGridColumn49.Header.VisiblePosition = 48;
			ultraGridColumn49.Hidden = true;
			ultraGridColumn49.Width = 19;
			ultraGridColumn50.Header.VisiblePosition = 49;
			ultraGridColumn50.Hidden = true;
			ultraGridColumn50.Width = 19;
			ultraGridColumn51.Header.VisiblePosition = 50;
			ultraGridColumn51.Hidden = true;
			ultraGridColumn51.Width = 19;
			ultraGridColumn52.Header.VisiblePosition = 51;
			ultraGridColumn52.Hidden = true;
			ultraGridColumn52.Width = 26;
			ultraGridColumn53.Header.VisiblePosition = 52;
			ultraGridColumn53.Hidden = true;
			ultraGridColumn53.Width = 108;
			ultraGridColumn54.Header.VisiblePosition = 53;
			ultraGridColumn54.Hidden = true;
			ultraGridColumn54.Width = 96;
			ultraGridColumn55.Header.VisiblePosition = 54;
			ultraGridColumn55.Hidden = true;
			ultraGridColumn55.Width = 39;
			ultraGridColumn56.Header.VisiblePosition = 55;
			ultraGridColumn56.Hidden = true;
			ultraGridColumn56.Width = 39;
			ultraGridColumn57.Header.VisiblePosition = 56;
			ultraGridColumn57.Hidden = true;
			ultraGridColumn57.Width = 39;
			ultraGridColumn58.Header.VisiblePosition = 57;
			ultraGridColumn58.Hidden = true;
			ultraGridColumn58.Width = 39;
			ultraGridColumn59.Header.VisiblePosition = 58;
			ultraGridColumn59.Hidden = true;
			ultraGridColumn59.Width = 45;
			ultraGridColumn60.Header.VisiblePosition = 60;
			ultraGridColumn60.Hidden = true;
			ultraGridColumn60.Width = 58;
			ultraGridColumn61.Header.VisiblePosition = 61;
			ultraGridColumn61.Hidden = true;
			ultraGridColumn61.Width = 61;
			ultraGridColumn62.Header.VisiblePosition = 62;
			ultraGridColumn62.Hidden = true;
			ultraGridColumn62.Width = 43;
			ultraGridColumn63.Header.VisiblePosition = 63;
			ultraGridColumn63.Hidden = true;
			ultraGridColumn63.Width = 50;
			ultraGridColumn64.Header.VisiblePosition = 64;
			ultraGridColumn64.Hidden = true;
			ultraGridColumn64.Width = 49;
			ultraGridColumn65.Header.VisiblePosition = 65;
			ultraGridColumn65.Hidden = true;
			ultraGridColumn65.Width = 59;
			ultraGridColumn66.Header.VisiblePosition = 66;
			ultraGridColumn66.Hidden = true;
			ultraGridColumn66.Width = 18;
			ultraGridColumn67.Header.VisiblePosition = 67;
			ultraGridColumn67.Hidden = true;
			ultraGridColumn67.Width = 20;
			ultraGridColumn68.Header.VisiblePosition = 68;
			ultraGridColumn68.Hidden = true;
			ultraGridColumn68.Width = 20;
			ultraGridColumn69.Header.VisiblePosition = 69;
			ultraGridColumn69.Hidden = true;
			ultraGridColumn69.Width = 21;
			ultraGridColumn70.Header.VisiblePosition = 70;
			ultraGridColumn70.Hidden = true;
			ultraGridColumn70.Width = 21;
			ultraGridColumn71.Header.VisiblePosition = 71;
			ultraGridColumn71.Hidden = true;
			ultraGridColumn71.Width = 21;
			ultraGridColumn72.Header.VisiblePosition = 72;
			ultraGridColumn72.Hidden = true;
			ultraGridColumn72.Width = 21;
			ultraGridColumn73.Header.VisiblePosition = 73;
			ultraGridColumn73.Hidden = true;
			ultraGridColumn73.Width = 23;
			ultraGridColumn74.Header.VisiblePosition = 74;
			ultraGridColumn74.Hidden = true;
			ultraGridColumn74.Width = 21;
			ultraGridColumn75.Header.VisiblePosition = 75;
			ultraGridColumn75.Hidden = true;
			ultraGridColumn75.Width = 23;
			ultraGridColumn76.Header.VisiblePosition = 76;
			ultraGridColumn76.Hidden = true;
			ultraGridColumn76.Width = 36;
			ultraGridColumn77.Header.VisiblePosition = 77;
			ultraGridColumn77.Hidden = true;
			ultraGridColumn77.Width = 19;
			ultraGridColumn78.Header.VisiblePosition = 78;
			ultraGridColumn78.Hidden = true;
			ultraGridColumn78.Width = 21;
			ultraGridColumn79.Header.VisiblePosition = 79;
			ultraGridColumn79.Hidden = true;
			ultraGridColumn79.Width = 19;
			ultraGridColumn80.Header.VisiblePosition = 80;
			ultraGridColumn80.Hidden = true;
			ultraGridColumn80.Width = 20;
			ultraGridColumn81.Header.VisiblePosition = 81;
			ultraGridColumn81.Hidden = true;
			ultraGridColumn81.Width = 19;
			ultraGridColumn82.Header.VisiblePosition = 82;
			ultraGridColumn82.Hidden = true;
			ultraGridColumn82.Width = 25;
			ultraGridColumn83.Header.VisiblePosition = 83;
			ultraGridColumn83.Hidden = true;
			ultraGridColumn83.Width = 22;
			ultraGridColumn84.Header.VisiblePosition = 84;
			ultraGridColumn84.Hidden = true;
			ultraGridColumn84.Width = 25;
			ultraGridColumn85.Header.VisiblePosition = 85;
			ultraGridColumn85.Hidden = true;
			ultraGridColumn85.Width = 19;
			ultraGridColumn86.Header.VisiblePosition = 86;
			ultraGridColumn86.Hidden = true;
			ultraGridColumn86.Width = 23;
			ultraGridColumn87.Header.VisiblePosition = 87;
			ultraGridColumn87.Hidden = true;
			ultraGridColumn87.Width = 23;
			ultraGridColumn88.Header.VisiblePosition = 88;
			ultraGridColumn88.Hidden = true;
			ultraGridColumn88.Width = 23;
			ultraGridColumn89.Header.VisiblePosition = 89;
			ultraGridColumn89.Hidden = true;
			ultraGridColumn89.Width = 25;
			ultraGridColumn90.Header.VisiblePosition = 90;
			ultraGridColumn90.Hidden = true;
			ultraGridColumn90.Width = 39;
			ultraGridColumn91.Header.VisiblePosition = 91;
			ultraGridColumn91.Hidden = true;
			ultraGridColumn91.Width = 46;
			ultraGridColumn92.Header.VisiblePosition = 92;
			ultraGridColumn92.Hidden = true;
			ultraGridColumn92.Width = 36;
			ultraGridColumn93.Header.VisiblePosition = 93;
			ultraGridColumn93.Hidden = true;
			ultraGridColumn93.Width = 37;
			ultraGridColumn94.Header.VisiblePosition = 94;
			ultraGridColumn94.Hidden = true;
			ultraGridColumn94.Width = 19;
			ultraGridColumn95.Header.VisiblePosition = 95;
			ultraGridColumn95.Hidden = true;
			ultraGridColumn95.Width = 15;
			ultraGridColumn96.Header.VisiblePosition = 96;
			ultraGridColumn96.Hidden = true;
			ultraGridColumn96.Width = 15;
			ultraGridColumn97.Header.VisiblePosition = 97;
			ultraGridColumn97.Hidden = true;
			ultraGridColumn97.Width = 15;
			ultraGridColumn98.Header.VisiblePosition = 98;
			ultraGridColumn98.Hidden = true;
			ultraGridColumn98.Width = 30;
			ultraGridColumn99.Header.VisiblePosition = 99;
			ultraGridColumn99.Hidden = true;
			ultraGridColumn99.Width = 28;
			ultraGridColumn100.Header.VisiblePosition = 100;
			ultraGridColumn100.Hidden = true;
			ultraGridColumn100.Width = 27;
			ultraGridColumn101.Header.VisiblePosition = 101;
			ultraGridColumn101.Hidden = true;
			ultraGridColumn101.Width = 26;
			ultraGridColumn102.Header.VisiblePosition = 102;
			ultraGridColumn102.Hidden = true;
			ultraGridColumn102.Width = 21;
			ultraGridColumn103.Header.VisiblePosition = 103;
			ultraGridColumn103.Hidden = true;
			ultraGridColumn103.Width = 21;
			ultraGridColumn104.Header.VisiblePosition = 104;
			ultraGridColumn104.Width = 47;
			ultraGridColumn105.Header.VisiblePosition = 105;
			ultraGridColumn105.Width = 28;
			ultraGridColumn106.Header.VisiblePosition = 107;
			ultraGridColumn106.Width = 51;
			ultraGridColumn107.Header.VisiblePosition = 106;
			ultraGridColumn107.Width = 35;
			ultraGridColumn108.Header.VisiblePosition = 35;
			ultraGridColumn108.Hidden = true;
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
																										 ultraGridColumn105,
																										 ultraGridColumn106,
																										 ultraGridColumn107,
																										 ultraGridColumn108});
			this.cmbNombre.DisplayLayout.BandsSerializer.Add(ultraGridBand1);
			this.cmbNombre.DisplayMember = "Nombre";
			this.cmbNombre.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbNombre.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbNombre.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbNombre.Location = new System.Drawing.Point(104, 40);
			this.cmbNombre.Name = "cmbNombre";
			this.cmbNombre.Size = new System.Drawing.Size(328, 21);
			this.cmbNombre.TabIndex = 19;
			this.cmbNombre.ValueMember = "idPersonal";
			this.cmbNombre.InitializeLayout += new Infragistics.Win.UltraWinGrid.InitializeLayoutEventHandler(this.cmbNombre_InitializeLayout);
			// 
			// label2
			// 
			this.label2.AutoSize = true;
			this.label2.Location = new System.Drawing.Point(8, 40);
			this.label2.Name = "label2";
			this.label2.Size = new System.Drawing.Size(44, 16);
			this.label2.TabIndex = 18;
			this.label2.Text = "Nombre";
			this.label2.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// txtConcepto
			// 
			this.txtConcepto.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtConcepto.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsRolPrestamo, "_RolPrestamo.Concepto"));
			this.txtConcepto.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtConcepto.Location = new System.Drawing.Point(104, 136);
			this.txtConcepto.Name = "txtConcepto";
			this.txtConcepto.Size = new System.Drawing.Size(464, 22);
			this.txtConcepto.TabIndex = 17;
			// 
			// label1
			// 
			this.label1.AutoSize = true;
			this.label1.Location = new System.Drawing.Point(8, 136);
			this.label1.Name = "label1";
			this.label1.Size = new System.Drawing.Size(51, 16);
			this.label1.TabIndex = 16;
			this.label1.Text = "Concepto";
			this.label1.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// txtNumCuotas
			// 
			this.txtNumCuotas.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsRolPrestamo, "_RolPrestamo.NumCuotas"));
			this.txtNumCuotas.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtNumCuotas.FormatString = "";
			this.txtNumCuotas.Location = new System.Drawing.Point(320, 104);
			this.txtNumCuotas.MaskInput = "nnn";
			this.txtNumCuotas.Name = "txtNumCuotas";
			this.txtNumCuotas.PromptChar = ' ';
			this.txtNumCuotas.RightToLeft = System.Windows.Forms.RightToLeft.Yes;
			this.txtNumCuotas.Size = new System.Drawing.Size(88, 22);
			this.txtNumCuotas.SpinButtonDisplayStyle = Infragistics.Win.ButtonDisplayStyle.Always;
			this.txtNumCuotas.TabIndex = 29;
			this.txtNumCuotas.Click += new System.EventHandler(this.txtNumCuotas_Click);
			this.txtNumCuotas.Enter += new System.EventHandler(this.txtNumCuotas_Enter);
			// 
			// label6
			// 
			this.label6.AutoSize = true;
			this.label6.Location = new System.Drawing.Point(232, 104);
			this.label6.Name = "label6";
			this.label6.Size = new System.Drawing.Size(69, 16);
			this.label6.TabIndex = 30;
			this.label6.Text = "N. de Cuotas";
			this.label6.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// errorProvider
			// 
			this.errorProvider.ContainerControl = this;
			// 
			// DetallePrestamo
			// 
			this.DetallePrestamo.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.DetallePrestamo.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			this.DetallePrestamo.Location = new System.Drawing.Point(456, 400);
			this.DetallePrestamo.Name = "DetallePrestamo";
			this.DetallePrestamo.Size = new System.Drawing.Size(24, 25);
			this.DetallePrestamo.TabIndex = 31;
			this.DetallePrestamo.Text = "P";
			this.toolTip.SetToolTip(this.DetallePrestamo, "Detalle de Pagos");
			this.DetallePrestamo.Click += new System.EventHandler(this.DetallePrestamo_Click);
			// 
			// btSolicitud
			// 
			this.btSolicitud.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btSolicitud.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			this.btSolicitud.Location = new System.Drawing.Point(488, 400);
			this.btSolicitud.Name = "btSolicitud";
			this.btSolicitud.Size = new System.Drawing.Size(24, 25);
			this.btSolicitud.TabIndex = 38;
			this.btSolicitud.Text = "S";
			this.toolTip.SetToolTip(this.btSolicitud, "Solicitud de Prestamo");
			this.btSolicitud.Click += new System.EventHandler(this.btSolicitud_Click);
			// 
			// ultraGrid1
			// 
			this.ultraGrid1.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.ultraGrid1.Cursor = System.Windows.Forms.Cursors.Default;
			this.ultraGrid1.DataMember = "_RolPrestamo.RolPrestamo - RolPrestamoDetalle";
			this.ultraGrid1.DataSource = this.cdsRolPrestamo;
			appearance1.BackColor = System.Drawing.Color.White;
			this.ultraGrid1.DisplayLayout.Appearance = appearance1;
			this.ultraGrid1.DisplayLayout.AutoFitColumns = true;
			ultraGridBand2.AddButtonCaption = "RolPrestamoDetalle";
			ultraGridColumn109.Header.VisiblePosition = 6;
			ultraGridColumn109.Hidden = true;
			ultraGridColumn110.Header.VisiblePosition = 8;
			ultraGridColumn110.Hidden = true;
			ultraGridColumn110.Width = 60;
			ultraGridColumn111.CellActivation = Infragistics.Win.UltraWinGrid.Activation.NoEdit;
			ultraGridColumn111.Header.Caption = "F. Propuesta";
			ultraGridColumn111.Header.VisiblePosition = 0;
			ultraGridColumn111.Width = 152;
			ultraGridColumn112.Format = "dd/MMM/yyyy";
			ultraGridColumn112.Header.Caption = "F. Pago";
			ultraGridColumn112.Header.VisiblePosition = 1;
			ultraGridColumn112.MaskInput = "";
			ultraGridColumn112.Width = 136;
			ultraGridColumn113.CellActivation = Infragistics.Win.UltraWinGrid.Activation.NoEdit;
			appearance2.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn113.CellAppearance = appearance2;
			ultraGridColumn113.Format = "#,##0.00";
			ultraGridColumn113.Header.VisiblePosition = 2;
			ultraGridColumn113.Width = 112;
			ultraGridColumn114.Header.VisiblePosition = 4;
			ultraGridColumn114.Hidden = true;
			appearance3.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn115.CellAppearance = appearance3;
			ultraGridColumn115.Format = "#,##0.00";
			ultraGridColumn115.Header.Caption = "Cobrado";
			ultraGridColumn115.Header.VisiblePosition = 3;
			ultraGridColumn115.Width = 139;
			ultraGridColumn116.Header.VisiblePosition = 5;
			ultraGridColumn116.Hidden = true;
			ultraGridColumn117.Header.VisiblePosition = 7;
			ultraGridColumn117.Hidden = true;
			ultraGridColumn117.Width = 75;
			ultraGridColumn118.Header.VisiblePosition = 9;
			ultraGridColumn118.Hidden = true;
			ultraGridBand2.Columns.AddRange(new object[] {
																										 ultraGridColumn109,
																										 ultraGridColumn110,
																										 ultraGridColumn111,
																										 ultraGridColumn112,
																										 ultraGridColumn113,
																										 ultraGridColumn114,
																										 ultraGridColumn115,
																										 ultraGridColumn116,
																										 ultraGridColumn117,
																										 ultraGridColumn118});
			this.ultraGrid1.DisplayLayout.BandsSerializer.Add(ultraGridBand2);
			this.ultraGrid1.DisplayLayout.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.TemplateOnBottom;
			this.ultraGrid1.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.True;
			this.ultraGrid1.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.True;
			appearance4.BackColor = System.Drawing.Color.Transparent;
			this.ultraGrid1.DisplayLayout.Override.CardAreaAppearance = appearance4;
			appearance5.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance5.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance5.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance5.FontData.BoldAsString = "True";
			appearance5.FontData.Name = "Arial";
			appearance5.FontData.SizeInPoints = 10F;
			appearance5.ForeColor = System.Drawing.Color.White;
			appearance5.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.ultraGrid1.DisplayLayout.Override.HeaderAppearance = appearance5;
			appearance6.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance6.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance6.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ultraGrid1.DisplayLayout.Override.RowSelectorAppearance = appearance6;
			appearance7.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance7.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance7.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ultraGrid1.DisplayLayout.Override.SelectedRowAppearance = appearance7;
			this.ultraGrid1.Location = new System.Drawing.Point(8, 176);
			this.ultraGrid1.Name = "ultraGrid1";
			this.ultraGrid1.Size = new System.Drawing.Size(560, 216);
			this.ultraGrid1.TabIndex = 32;
			this.ultraGrid1.KeyDown += new System.Windows.Forms.KeyEventHandler(this.ultraGrid1_KeyDown);
			// 
			// cdsAsiento
			// 
			this.cdsAsiento.BindingContextCtrl = this;
			this.cdsAsiento.DataLibrary = "LibContabilidad";
			this.cdsAsiento.DataLibraryUrl = "";
			this.cdsAsiento.DataSetDef = "dsAsiento";
			this.cdsAsiento.FillOnRequest = false;
			this.cdsAsiento.Locale = new System.Globalization.CultureInfo("es-EC");
			this.cdsAsiento.Name = "cdsAsiento";
			this.cdsAsiento.SchemaClassName = "LibContabilidad.DataClass";
			this.cdsAsiento.SchemaDef = null;
			// 
			// cmbPrestamo
			// 
			this.cmbPrestamo.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbPrestamo.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbPrestamo.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsRolPrestamo, "_RolPrestamo.idRolTipoPrestamo"));
			this.cmbPrestamo.DataMember = "RolTipoPrestamo";
			this.cmbPrestamo.DataSource = this.cdsRolTipoPrest;
			ultraGridBand3.ColHeadersVisible = false;
			ultraGridColumn119.Header.VisiblePosition = 1;
			ultraGridColumn119.Hidden = true;
			ultraGridColumn120.Header.VisiblePosition = 0;
			ultraGridColumn120.Width = 250;
			ultraGridColumn121.Header.VisiblePosition = 2;
			ultraGridColumn121.Hidden = true;
			ultraGridColumn122.Header.VisiblePosition = 3;
			ultraGridColumn122.Hidden = true;
			ultraGridColumn123.Header.VisiblePosition = 4;
			ultraGridColumn123.Hidden = true;
			ultraGridBand3.Columns.AddRange(new object[] {
																										 ultraGridColumn119,
																										 ultraGridColumn120,
																										 ultraGridColumn121,
																										 ultraGridColumn122,
																										 ultraGridColumn123});
			this.cmbPrestamo.DisplayLayout.BandsSerializer.Add(ultraGridBand3);
			this.cmbPrestamo.DisplayMember = "TipoPrestamo";
			this.cmbPrestamo.Font = new System.Drawing.Font("Tahoma", 8.25F);
			this.cmbPrestamo.Location = new System.Drawing.Point(104, 8);
			this.cmbPrestamo.Name = "cmbPrestamo";
			this.cmbPrestamo.Size = new System.Drawing.Size(240, 22);
			this.cmbPrestamo.TabIndex = 34;
			this.cmbPrestamo.ValueMember = "idRolTipoPrest";
			// 
			// cdsRolTipoPrest
			// 
			this.cdsRolTipoPrest.BindingContextCtrl = this;
			this.cdsRolTipoPrest.DataLibrary = "LibPersonal";
			this.cdsRolTipoPrest.DataLibraryUrl = "";
			this.cdsRolTipoPrest.DataSetDef = "dsRolTipoPrestamo";
			this.cdsRolTipoPrest.Locale = new System.Globalization.CultureInfo("es-EC");
			this.cdsRolTipoPrest.Name = "cdsRolTipoPrest";
			this.cdsRolTipoPrest.SchemaClassName = "LibPersonal.DataClass";
			this.cdsRolTipoPrest.SchemaDef = null;
			// 
			// label7
			// 
			this.label7.AutoSize = true;
			this.label7.Location = new System.Drawing.Point(8, 8);
			this.label7.Name = "label7";
			this.label7.Size = new System.Drawing.Size(26, 16);
			this.label7.TabIndex = 35;
			this.label7.Text = "Tipo";
			this.label7.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// btTipoPrestamo
			// 
			this.btTipoPrestamo.Location = new System.Drawing.Point(360, 8);
			this.btTipoPrestamo.Name = "btTipoPrestamo";
			this.btTipoPrestamo.Size = new System.Drawing.Size(24, 22);
			this.btTipoPrestamo.TabIndex = 36;
			this.btTipoPrestamo.Text = "...";
			this.btTipoPrestamo.Visible = false;
			this.btTipoPrestamo.Click += new System.EventHandler(this.btTipoPrestamo_Click);
			// 
			// barraDato1
			// 
			this.barraDato1.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.barraDato1.BarraMovimiento = true;
			this.barraDato1.DataMember = "_RolPrestamo";
			this.barraDato1.DataNombreId = "idRolPrestamo";
			this.barraDato1.DataOrden = "";
			this.barraDato1.DataSource = this.cdsRolPrestamo;
			this.barraDato1.DataTabla = "RolPrestamo";
			this.barraDato1.DataTablaHija = "";
			this.barraDato1.Location = new System.Drawing.Point(8, 399);
			this.barraDato1.Name = "barraDato1";
			this.barraDato1.Size = new System.Drawing.Size(304, 26);
			this.barraDato1.TabIndex = 37;
			this.barraDato1.VisibleBorrar = true;
			this.barraDato1.VisibleBuscar = true;
			this.barraDato1.VisibleEditar = true;
			this.barraDato1.VisibleImprimir = true;
			this.barraDato1.VisibleNuevo = true;
			this.barraDato1.Buscar += new System.EventHandler(this.barraDato1_Buscar);
			this.barraDato1.Load += new System.EventHandler(this.barraDato1_Load);
			this.barraDato1.Borrar += new System.EventHandler(this.barraDato1_Borrar);
			this.barraDato1.Imprimir += new System.EventHandler(this.barraDato1_Imprimir);
			this.barraDato1.Restaurar += new System.EventHandler(this.barraDato1_Restaurar);
			this.barraDato1.Refresca += new System.EventHandler(this.barraDato1_Refresca);
			this.barraDato1.Crear += new System.EventHandler(this.barraDato1_Crear);
			this.barraDato1.Deshacer += new System.EventHandler(this.barraDato1_Deshacer);
			this.barraDato1.Copiar += new System.EventHandler(this.barraDato1_Copiar);
			this.barraDato1.Impresora += new System.EventHandler(this.barraDato1_Impresora);
			this.barraDato1.Editar += new System.EventHandler(this.barraDato1_Editar);
			this.barraDato1.Grabar += new System.EventHandler(this.barraDato1_Grabar);
			// 
			// txtNumero
			// 
			this.txtNumero.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtNumero.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsRolPrestamo, "_RolPrestamo.Numero"));
			this.txtNumero.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtNumero.Location = new System.Drawing.Point(472, 8);
			this.txtNumero.Name = "txtNumero";
			this.txtNumero.Size = new System.Drawing.Size(96, 22);
			this.txtNumero.TabIndex = 39;
			// 
			// label8
			// 
			this.label8.AutoSize = true;
			this.label8.Location = new System.Drawing.Point(400, 8);
			this.label8.Name = "label8";
			this.label8.Size = new System.Drawing.Size(44, 16);
			this.label8.TabIndex = 40;
			this.label8.Text = "Número";
			this.label8.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// groupBox1
			// 
			this.groupBox1.Location = new System.Drawing.Point(0, 160);
			this.groupBox1.Name = "groupBox1";
			this.groupBox1.Size = new System.Drawing.Size(576, 8);
			this.groupBox1.TabIndex = 41;
			this.groupBox1.TabStop = false;
			// 
			// btAceptar
			// 
			this.btAceptar.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btAceptar.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			this.btAceptar.Location = new System.Drawing.Point(8, 400);
			this.btAceptar.Name = "btAceptar";
			this.btAceptar.Size = new System.Drawing.Size(88, 24);
			this.btAceptar.TabIndex = 159;
			this.btAceptar.Text = "&Aceptar";
			this.btAceptar.Visible = false;
			this.btAceptar.Click += new System.EventHandler(this.btAceptar_Click);
			// 
			// btCancelar
			// 
			this.btCancelar.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btCancelar.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			this.btCancelar.DialogResult = System.Windows.Forms.DialogResult.Cancel;
			this.btCancelar.Location = new System.Drawing.Point(104, 400);
			this.btCancelar.Name = "btCancelar";
			this.btCancelar.Size = new System.Drawing.Size(88, 24);
			this.btCancelar.TabIndex = 158;
			this.btCancelar.Text = "&Cancelar";
			this.btCancelar.Visible = false;
			this.btCancelar.Click += new System.EventHandler(this.btCancelar_Click);
			// 
			// txtEstado
			// 
			this.txtEstado.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsRolPrestamo, "_RolPrestamo.idAsiento"));
			this.txtEstado.Location = new System.Drawing.Point(384, 400);
			this.txtEstado.Name = "txtEstado";
			this.txtEstado.Size = new System.Drawing.Size(40, 22);
			this.txtEstado.TabIndex = 160;
			// 
			// RolPrestamo
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 14);
			this.ClientSize = new System.Drawing.Size(576, 430);
			this.Controls.Add(this.txtEstado);
			this.Controls.Add(this.btAceptar);
			this.Controls.Add(this.btCancelar);
			this.Controls.Add(this.groupBox1);
			this.Controls.Add(this.label8);
			this.Controls.Add(this.label7);
			this.Controls.Add(this.label6);
			this.Controls.Add(this.label5);
			this.Controls.Add(this.label4);
			this.Controls.Add(this.label3);
			this.Controls.Add(this.label2);
			this.Controls.Add(this.label1);
			this.Controls.Add(this.txtNumero);
			this.Controls.Add(this.btSolicitud);
			this.Controls.Add(this.barraDato1);
			this.Controls.Add(this.btTipoPrestamo);
			this.Controls.Add(this.cmbPrestamo);
			this.Controls.Add(this.ultraGrid1);
			this.Controls.Add(this.DetallePrestamo);
			this.Controls.Add(this.txtNumCuotas);
			this.Controls.Add(this.txtIdRolPrestamo);
			this.Controls.Add(this.txtIdAsiento);
			this.Controls.Add(this.btEgreso);
			this.Controls.Add(this.cmbFechaCobro);
			this.Controls.Add(this.cmbFechaEntrega);
			this.Controls.Add(this.txtValor);
			this.Controls.Add(this.cmbNombre);
			this.Controls.Add(this.txtConcepto);
			this.MaximizeBox = false;
			this.Name = "RolPrestamo";
			this.Text = "Prestámos y Anticipos";
			this.Load += new System.EventHandler(this.RolPrestamo_Load);
			((System.ComponentModel.ISupportInitialize)(this.cmbFechaCobro)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsRolPrestamo)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdvRolPrestamo)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsPersonal)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdvPersonal)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIdRolPrestamo)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIdAsiento)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbFechaEntrega)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtValor)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbNombre)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtConcepto)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNumCuotas)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraGrid1)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsAsiento)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbPrestamo)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsRolTipoPrest)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNumero)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtEstado)).EndInit();
			this.ResumeLayout(false);

		}
		#endregion

		private void barraDato1_Borrar(object sender, System.EventArgs e)
		{
			if (ValorPagos() > 0)
			{
				MessageBox.Show("No puede borrar Prestamos que tengan valores ya pagados");
			}

			C1.Data.C1DataRow drPrestamo = this.cdsRolPrestamo.TableViews["RolPrestamo"].Rows.Find(int.Parse(this.txtIdRolPrestamo.Value.ToString()));

			if ((int)drPrestamo["Estado"] == 2)
			{
				MessageBox.Show("No puede Anular ya esta Pagado", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				return;
			}

			if ((int)drPrestamo["Estado"] == 3)
			{
				MessageBox.Show("El registro ya esta Anulado", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				return;
			}

			if (DialogResult.No == MessageBox.Show("Desea Borrar este Prestamo",
				"Información", MessageBoxButtons.YesNo, MessageBoxIcon.Question,
				MessageBoxDefaultButton.Button2)) return;

			this.barraDato1.BorraRegistro();
		}

		private void barraDato1_Buscar(object sender, System.EventArgs e)
		{
			using (BuscaPrestamo miBusqueda = new BuscaPrestamo())
			{
				if (miBusqueda.ShowDialog() == DialogResult.OK)
				{
					int id = int.Parse(miBusqueda.idBusca.Text);
					barraDato1.IdRegistro = id;
					this.barraDato1.ProximoId(5);
				}
			}
		}

		private void PrestamoBuscado(int id)
		{
			int idx = this.cdvRolPrestamo.Find(id);
			if (idx == -1) 
				MessageBox.Show("IdRolPrestamo no encontrado " + id.ToString(), "Error");
			else BindingContext[this.cdsRolPrestamo,"_RolPrestamo"].Position = idx;
			this.barraDato1.ProximoId(5);
		}

		private void barraDato1_Copiar(object sender, System.EventArgs e)
		{
		}

		private void barraDato1_Crear(object sender, System.EventArgs e)
		{
			this.barraDato1.CrearRegistro();
			C1.Data.C1DataRow drPrestamo = this.cdsRolPrestamo.TableViews["RolPrestamo"].Rows.Find(int.Parse(this.txtIdRolPrestamo.Value.ToString()));
			drPrestamo["FechaEntrega"] = DateTime.Today;
			drPrestamo["FechaCobro"] = DateTime.Today;		
			this.txtNumero.Enabled = false;

			if (BCruceAnticipo)
			{
				drPrestamo["idRolTipoPrestamo"] = 8;
				drPrestamo["Capital"] = DValor;
				this.cmbPrestamo.Enabled = false;
				this.label5.Visible = false;
				this.cmbFechaCobro.Visible = false;
				this.label6.Visible = false;
				this.txtNumCuotas.Visible = false;
				this.txtValor.Enabled = false;
				this.cmbNombre.Select();
			}
			else 
			{
				drPrestamo["idRolTipoPrestamo"] = 1;
				this.cmbPrestamo.Select();
			}
		}

		private void barraDato1_Deshacer(object sender, System.EventArgs e)
		{
			this.BorraErrores();
			this.barraDato1.DeshacerRegistro();
		}

		private void barraDato1_Editar(object sender, System.EventArgs e)
		{
			C1.Data.C1DataRow drPrestamo = this.cdsRolPrestamo.TableViews["RolPrestamo"].Rows.Find(int.Parse(this.txtIdRolPrestamo.Value.ToString()));

			if ((int)drPrestamo["Estado"] == 2)
			{
				MessageBox.Show("No puede editar ya esta Pagado", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				return;
			}

			if ((int)drPrestamo["Estado"] == 3)
			{
				MessageBox.Show("No puede editar Registro Anulado", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				return;
			}

			this.barraDato1.EditarRegistro();		
			this.txtNumero.Enabled = false;
		}

		private void BorraErrores()
		{
			this.errorProvider.SetError(this.cmbNombre, "");
			this.errorProvider.SetError(this.txtValor, "");
			this.errorProvider.SetError(this.cmbFechaCobro, "");
			this.errorProvider.SetError(this.cmbFechaEntrega, "");
			this.errorProvider.SetError(this.txtNumCuotas, "");
		}

		private bool Verifica()
		{
			bool bOk = true;
			BorraErrores();
			if (this.cmbNombre.Value == System.DBNull.Value)
			{
				this.errorProvider.SetError(this.cmbNombre, "Ingreso Nombre");
				bOk = false;
			}
			if (this.txtValor.Value == System.DBNull.Value || (double) this.txtValor.Value == 0)
			{
				this.errorProvider.SetError(this.txtValor, "Ingreso Valor del Prestamo");
				bOk = false;
			}
			if (!BCruceAnticipo)
			{
				if (this.txtNumCuotas.Value == System.DBNull.Value || (int) this.txtNumCuotas.Value == 0)
				{
					this.errorProvider.SetError(this.txtNumCuotas, "Ingreso Numero de Cuotas");
					bOk = false;
				}
						
				if (this.cmbFechaCobro.Value == System.DBNull.Value)
				{
					this.errorProvider.SetError(this.cmbFechaCobro, "Ingreso Fecha de Cobro");
					bOk = false;
				}
			}
			if (this.cmbFechaEntrega.Value == System.DBNull.Value)
			{
				this.errorProvider.SetError(this.cmbFechaEntrega, "Ingreso Fecha de Entrega");
				bOk = false;
			}
			
			return bOk;
		}

		private float ValorPagos()
		{
			string stSelect = "SELECT SUM(CapCobrado) FROM RolPrestamoDetalle"
				+ " WHERE idRolPrestamo = " + this.txtIdRolPrestamo.Value.ToString();
			SqlCommand cmdActualiza;
			this.cdsRolPrestamo.Schema.Connections[0].Open();
			C1.Data.SchemaObjects.Connection miConeccion = 
				this.cdsRolPrestamo.Schema.Connections[0];

			cmdActualiza = new SqlCommand (stSelect
				, (SqlConnection) miConeccion.DbConnection);
			cmdActualiza.Transaction = (SqlTransaction)
				miConeccion.DbTransaction;
			float fPagado = 0;
			try
			{
				fPagado = (float) cmdActualiza.ExecuteScalar();
			}
			catch
			{
			}
			this.cdsRolPrestamo.Schema.Connections[0].Close();
			return fPagado;
		}

		CultureInfo us = new CultureInfo("en-US");

		private void barraDato1_Grabar(object sender, System.EventArgs e)
		{
			if (!Verifica()) return;
			if (ValorPagos() > 0) 
			{
				MessageBox.Show("No puede modificar prestamos que tienen valores ya cancelados");
			}

			if (this.barraDato1.bNuevo) this.txtNumero.Value = Funcion.sEscalarSQL(cdsAsiento, string.Format("Select dbo.SecuencialPrestamoEmpleados({0})", (int)this.cmbPrestamo.Value));
			this.barraDato1.GrabaRegistro();
			
			DateTime FechaSol = (DateTime)this.cmbFechaEntrega.Value;
			int idPrest = int.Parse(this.txtIdRolPrestamo.Value.ToString());
			DateTime dtFecha = (DateTime) this.cmbFechaCobro.Value;
			double dValor = (double) this.txtValor.Value;
			
//			if ((int)this.cmbPrestamo.Value == 1 && (int)this.cmbPrestamo.Value == 2 && (int)this.cmbPrestamo.Value == 3 && (int)this.cmbPrestamo.Value == 6)
//			{
			string sSQL = string.Format("Exec CreaAutorizacionesDePago {0}, {1}, '{2}', '{3}', {4}, 2, '1 DE 1', '{5}', 0", 
				(int)this.cmbPrestamo.Value, idPrest, FechaSol.ToString("yyyyMMdd"), FechaSol.ToString("yyyyMMdd"), (double)this.txtValor.Value, this.txtConcepto.Text.ToString());			
			Funcion.EjecutaSQL(cdsAsiento, sSQL, true);
//			}

			if (!BCruceAnticipo)
			{
				if ((int)this.cmbPrestamo.Value == 2 || (int)this.cmbPrestamo.Value == 3 || (int)this.cmbPrestamo.Value == 6)
				{
					string stSelect = "EXEC TablaAmortizacion "
						+ dValor.ToString("0.00", us)
						+ ", " + this.txtNumCuotas.Value.ToString()
						+ ", " + this.txtIdRolPrestamo.Value.ToString()
						+ ", '" + dtFecha.ToString("yyyyMMdd")
						+ "'";

					SqlCommand cmdActualiza;
					this.cdsRolPrestamo.Schema.Connections[0].Open();
					C1.Data.SchemaObjects.Connection miConeccion = this.cdsRolPrestamo.Schema.Connections[0];

					cmdActualiza = new SqlCommand (stSelect, (SqlConnection) miConeccion.DbConnection);
					cmdActualiza.Transaction = (SqlTransaction)miConeccion.DbTransaction;

					try
					{
						cmdActualiza.ExecuteNonQuery();
					}
					catch(System.Exception ex)
					{
						MessageBox.Show(ex.Message, "TablaAmortizacion");
					}
			
					this.cdsRolPrestamo.Schema.Connections[0].Close();
					this.cdsRolPrestamo.Clear();
					this.cdsRolPrestamo.Fill();
				}

				PrestamoBuscado(idPrest);
			}
		}

		private void barraDato1_Impresora(object sender, System.EventArgs e)
		{
		
		}

		private void barraDato1_Imprimir(object sender, System.EventArgs e)
		{
			RolPrestamoImp miImpresion = new RolPrestamoImp();
			miImpresion.MdiParent = this.MdiParent;
			miImpresion.Show();
		}

		private void barraDato1_Restaurar(object sender, System.EventArgs e)
		{
		
		}

		private void cdsRolPrestamo_BeforeFill(object sender, C1.Data.FillEventArgs e)
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

		private void ModificaGrilla(bool bModifica)
		{
			ultraGrid1.Enabled = true;
			if (bModifica)
			{
				ultraGrid1.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.True;
				ultraGrid1.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.True;
			}
			else
			{
				ultraGrid1.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.False;
				ultraGrid1.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.False;
			}
		}

		private void RolPrestamo_Load(object sender, System.EventArgs e)
		{
			Cursor = Cursors.Default;
			#region Filtro			
			this.barraDato1.stFiltroExt = " And idRolTipoPrestamo In (1, 2, 3, 6)";

			string sFiltro = "idRolTipoPrest In (1, 2, 3, 6)";						
			C1.Data.FilterCondition fcP = new C1.Data.FilterCondition(cdsRolTipoPrest, "RolTipoPrestamo", sFiltro);
			C1.Data.FilterConditions fcT = new C1.Data.FilterConditions();
			fcT.Add(fcP);
			cdsRolTipoPrest.Fill(fcT, false);			
			#endregion Filtro

			this.txtIdRolPrestamo.Width = 0;
			this.txtIdAsiento.Width = 0;
			this.barraDato1.PosUltima();
			this.barraDato1.HabilitaControles(false);

			ModificaGrilla(false);

			if (BCruceAnticipo)
			{
				this.btAceptar.Visible = true;
				this.btCancelar.Visible = true;
				this.barraDato1.Visible = false;
				this.barraDato1_Crear(sender, e);
			}
		}

		private void btEgreso_Click(object sender, System.EventArgs e)
		{
			C1.Data.C1DataRow drAsi = null;
			bool bRepetido = false;
			if (int.Parse(this.txtIdAsiento.Value.ToString()) > 0) 
			{
				if (DialogResult.No == MessageBox.Show("Asiento ya Existe. ¿Desea Reemplazarlo?", "Confirmación", MessageBoxButtons.YesNo, MessageBoxIcon.Question)) return;
				string stBorrar = "Delete AsientoDetalle Where idAsiento = " + this.txtIdAsiento.Value.ToString();
				Funcion.EjecutaSQL(cdsAsiento, stBorrar, true);
				bRepetido = true;
			}
			else
			{
				if (DialogResult.No == MessageBox.Show("¿Desea Generar un Egreso de este anticipo?", "Confirmación", MessageBoxButtons.YesNo, MessageBoxIcon.Question)) return;
				drAsi = cdsAsiento.TableViews["Asiento"].AddNew();
			}

			if (bRepetido)
			{
				string stFiltro = "idAsiento = " + txtIdAsiento.Text.ToString();
				C1.Data.FilterCondition fcPadre = new C1.Data.FilterCondition(cdsAsiento, "Asiento", stFiltro);
				C1.Data.FilterConditions fcTotal = new C1.Data.FilterConditions();
				fcTotal.Add(fcPadre);

				C1.Data.FilterCondition fcHija = new C1.Data.FilterCondition(cdsAsiento, "AsientoDetalle", "idAsiento = 0");
				fcTotal.Add(fcHija);
			
				this.cdsAsiento.Fill(fcTotal, false);

				drAsi = cdsAsiento.TableViews["Asiento"].Rows[0];
			}
			else
			{
				drAsi["Fecha"] = DateTime.Today;
				drAsi["CodAsiento"] = Contabilidad.NumeroAsiento(cdsAsiento, 3);
				drAsi["Descripcion"] = "Prestamo de " + cmbNombre.Text;
				drAsi["Total"] = txtValor.Value;
				drAsi["TotalH"] = txtValor.Value;
				drAsi["idTipoAsiento"] = 3;
				cdsAsiento.Update();
			}
			int IdAsi = (int) drAsi["idAsiento"];
			C1.Data.C1DataRow drDetAsi = cdsAsiento.TableViews["AsientoDetalle"].AddNew();
			drDetAsi["Debe"] = txtValor.Value;
			drDetAsi["idAsiento"] = IdAsi;
			string stSelect = "Select IsNull(idCuenta, 0) From Personal Where idPersonal = " + cmbNombre.Value.ToString();

			int iCuentaDebe = Funcion.iEscalarSQL(cdsAsiento, stSelect, false);
			if (iCuentaDebe < 1)
			{
				stSelect = "Select IsNull(idCuenta, 0) From RolTipoPrestamo Where idRolTipoPrest = " + cmbPrestamo.Value.ToString();

				iCuentaDebe = Funcion.iEscalarSQL(cdsAsiento, stSelect, false);
				if (iCuentaDebe < 1)
				{
					stSelect = "Select IsNull(idCuenta, 0) From AsignaCuenta Where idAsignacion = 19";
					iCuentaDebe = Funcion.iEscalarSQL(cdsAsiento, stSelect, false);
				}
			}
			drDetAsi["idCuenta"] = Funcion.iEscalarSQL(cdsAsiento, stSelect);
			
			drDetAsi = cdsAsiento.TableViews["AsientoDetalle"].AddNew();
			drDetAsi["idAsiento"] = IdAsi;
			drDetAsi["Haber"] = txtValor.Value;
			stSelect = "Select idCuenta From AsignaCuenta Where idAsignacion = 10";
			int iCuentaHaber = Funcion.iEscalarSQL(cdsAsiento, stSelect, false);
			drDetAsi["idCuenta"] = iCuentaHaber;
			cdsAsiento.Update();

			stSelect = "Update RolPrestamo SET idAsiento = " + IdAsi.ToString() + " Where idRolPrestamo = " + txtIdRolPrestamo.Value.ToString();
			txtIdAsiento.Value = IdAsi;
			Funcion.EjecutaSQL(cdsRolPrestamo, stSelect);
			MessageBox.Show("Asiento Generado", "Información", MessageBoxButtons.OK, MessageBoxIcon.Information);
		}

		private void btEgreso_MouseDown(object sender, System.Windows.Forms.MouseEventArgs e)
		{
			if (e.Button != MouseButtons.Right) return;
			if (int.Parse(this.txtIdAsiento.Value.ToString()) > 0) 
			{
				Cursor = Cursors.WaitCursor;
				frmAsientos Asiento = new frmAsientos(int.Parse(this.txtIdAsiento.Value.ToString()));
				Asiento.MdiParent = this.MdiParent;
				Asiento.Show();
				Cursor = Cursors.Default;
			}
			else
			{
				MessageBox.Show("Asiento NO se ha creado", "Información", MessageBoxButtons.OK, MessageBoxIcon.Information);
				return;
			}		
		}

		private void btTipoPrestamo_Click(object sender, System.EventArgs e)
		{
			RolPrestamoTipo miTipoPrest = new RolPrestamoTipo();
			miTipoPrest.MdiParent = this.MdiParent;
			miTipoPrest.Show();
		}

		private void DetallePrestamo_Click(object sender, System.EventArgs e)
		{
			if (int.Parse(this.txtIdAsiento.Value.ToString()) > 0) 
			{
				if (DialogResult.No == MessageBox.Show("¿Asiento ya Existe desea Reemplazarlo?", "Confirmación", MessageBoxButtons.YesNo, MessageBoxIcon.Question)) return;
			}
			else
			{
				if (DialogResult.No == MessageBox.Show("¿Desea Generar un Egreso de éste prestamo?", "Confirmación", MessageBoxButtons.YesNo, MessageBoxIcon.Question)) return;
			}

      Funcion.EjecutaSQL(cdsAsiento, "Exec PrestamoAsiento " + txtIdRolPrestamo.Value.ToString());
			barraDato1.ProximoId(5);
      MessageBox.Show("Asiento Generado", "Información");
		}

		private void barraDato1_Refresca(object sender, System.EventArgs e)
		{
		
		}

		private void ultraGrid1_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (ultraGrid1.ActiveCell == null) return;
			try
			{
				if (ultraGrid1.ActiveCell.Column.Key.ToString().ToUpper() == "NOTAS") return;
				switch(e.KeyValue)
				{
					case (int) Keys.PageDown:
						if(ultraGrid1.ActiveCell.DroppedDown == false)
							ultraGrid1.ActiveCell.DroppedDown = true;
						break;

					case (int) Keys.Up:
						if (ultraGrid1.ActiveCell.DroppedDown) return;

						ultraGrid1.PerformAction(UltraGridAction.ExitEditMode, false, false);
						ultraGrid1.PerformAction(UltraGridAction.AboveCell, false, false);
						e.Handled = true;
						ultraGrid1.PerformAction(UltraGridAction.EnterEditMode, false, false);
						break;
					case (int) Keys.Left:
						ultraGrid1.PerformAction(UltraGridAction.ExitEditMode, false, false);
						ultraGrid1.PerformAction(UltraGridAction.PrevCellByTab, false, false);
						e.Handled = true;
						ultraGrid1.PerformAction(UltraGridAction.EnterEditMode, false, false);
						break;
					case (int) Keys.Down:
						if (ultraGrid1.ActiveCell.DroppedDown) return;
			
						ultraGrid1.PerformAction(UltraGridAction.ExitEditMode, false, false);
						ultraGrid1.PerformAction(UltraGridAction.BelowCell, false, false);
						e.Handled = true;
						ultraGrid1.PerformAction(UltraGridAction.EnterEditMode, false, false);
						break;
					case (int) Keys.Enter:
						ultraGrid1.PerformAction(UltraGridAction.ExitEditMode, false, false);
						ultraGrid1.PerformAction(UltraGridAction.NextCellByTab, false, false);
						e.Handled = true;
						ultraGrid1.PerformAction(UltraGridAction.EnterEditMode, false, false);
						if(ultraGrid1.ActiveCell.DroppedDown == false)
							ultraGrid1.ActiveCell.DroppedDown = true;
						//						ultraGrid1.PerformAction(UltraGridAction.EnterEditModeAndDropdown, false, false);
						break;
				}
			}
			catch(System.Exception ex)
			{
				Console.WriteLine(ex.Message);
				//				MessageBox.Show(ex.Message, "Tecla");
			}

		}

		private void btSolicitud_Click(object sender, System.EventArgs e)
		{
			string stFiltro = "{RolPrestamo.idRolPrestamo} = " + txtIdRolPrestamo.Value.ToString();
			Cursor = Cursors.WaitCursor;
			Reporte miRep = new Reporte("PrestamoSolicitud.rpt", stFiltro);
			miRep.MdiParent = this.MdiParent;
			miRep.Show();
			Cursor = Cursors.Default;
		}

		private void barraDato1_Load(object sender, System.EventArgs e)
		{
		
		}

		private void txtValor_Click(object sender, System.EventArgs e)
		{
			this.txtValor.SelectAll();
		}

		private void txtValor_Enter(object sender, System.EventArgs e)
		{
			this.txtValor.SelectAll();
		}

		private void txtNumCuotas_Click(object sender, System.EventArgs e)
		{
			this.txtNumCuotas.SelectAll();
		}

		private void txtNumCuotas_Enter(object sender, System.EventArgs e)
		{
			this.txtNumCuotas.SelectAll();
		}

		private void btCancelar_Click(object sender, System.EventArgs e)
		{
			this.DialogResult = DialogResult.No;
		}

		private void btAceptar_Click(object sender, System.EventArgs e)
		{
			this.barraDato1_Grabar(sender, e);
			this.DialogResult = DialogResult.Yes;
		}

		private void cmbNombre_InitializeLayout(object sender, Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs e)
		{
		
		}
	}
}
