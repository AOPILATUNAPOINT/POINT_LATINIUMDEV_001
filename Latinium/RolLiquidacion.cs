using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using System.Globalization;
using CrystalDecisions.Shared;

namespace Latinium
{
	/// <summary>
	/// Descripción breve de RolLiquidacion.
	/// </summary>
	public class RolLiquidacion : DevExpress.XtraEditors.XtraForm
	{
		private C1.Data.C1DataSet cdsPersonal;
		private C1.Data.C1DataView cdvPersonal;
		private Infragistics.Win.Misc.UltraLabel ultraLabel1;
		private C1.Data.C1DataSet cdsRolPago;
		private System.Windows.Forms.Label label1;
		private System.Windows.Forms.Label label2;
		private System.Windows.Forms.Label label3;
		private System.Windows.Forms.Label label4;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtDecimoXIV;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtFondoReserva;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtVacaciones;
		private Infragistics.Win.UltraWinEditors.UltraDateTimeEditor cmbDecXIV;
		private Infragistics.Win.UltraWinEditors.UltraDateTimeEditor cmbDecXIII;
		private Infragistics.Win.UltraWinEditors.UltraDateTimeEditor cmbFondoReserva;
		private Infragistics.Win.UltraWinEditors.UltraDateTimeEditor cmbVacaciones;
		private Infragistics.Win.Misc.UltraButton btAceptar;
		private Infragistics.Win.Misc.UltraButton btCancelar;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbPersonal;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtDecimoXIII;
		private System.Windows.Forms.ErrorProvider errorProvider;
		private Infragistics.Win.Misc.UltraButton btImprimir;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtProvision1;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtProvision2;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtProvision3;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtProvision4;
		private C1.Data.C1DataSet cdsSeteoRol;
		private System.Windows.Forms.Label label5;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtAportable;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtProvAporte;
		private BarraDatoSQL.BarraDatoSQL barraDato1;
		private C1.Data.C1DataSet cdsLiquidacion;
		/// <summary>
		/// Variable del diseñador requerida.
		/// </summary>
		private System.ComponentModel.Container components = null;

		public RolLiquidacion()
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
			this.cdsPersonal = new C1.Data.C1DataSet();
			this.cdvPersonal = new C1.Data.C1DataView();
			this.cmbPersonal = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.cdsLiquidacion = new C1.Data.C1DataSet();
			this.ultraLabel1 = new Infragistics.Win.Misc.UltraLabel();
			this.cdsRolPago = new C1.Data.C1DataSet();
			this.label1 = new System.Windows.Forms.Label();
			this.label2 = new System.Windows.Forms.Label();
			this.label3 = new System.Windows.Forms.Label();
			this.label4 = new System.Windows.Forms.Label();
			this.txtDecimoXIV = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtFondoReserva = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtVacaciones = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.cmbDecXIV = new Infragistics.Win.UltraWinEditors.UltraDateTimeEditor();
			this.cmbDecXIII = new Infragistics.Win.UltraWinEditors.UltraDateTimeEditor();
			this.cmbFondoReserva = new Infragistics.Win.UltraWinEditors.UltraDateTimeEditor();
			this.cmbVacaciones = new Infragistics.Win.UltraWinEditors.UltraDateTimeEditor();
			this.btAceptar = new Infragistics.Win.Misc.UltraButton();
			this.btCancelar = new Infragistics.Win.Misc.UltraButton();
			this.txtDecimoXIII = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.errorProvider = new System.Windows.Forms.ErrorProvider();
			this.btImprimir = new Infragistics.Win.Misc.UltraButton();
			this.txtProvision1 = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.cdsSeteoRol = new C1.Data.C1DataSet();
			this.txtProvision2 = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.txtProvision3 = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.txtProvision4 = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.label5 = new System.Windows.Forms.Label();
			this.txtProvAporte = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.txtAportable = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.barraDato1 = new BarraDatoSQL.BarraDatoSQL();
			((System.ComponentModel.ISupportInitialize)(this.cdsPersonal)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdvPersonal)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbPersonal)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsLiquidacion)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsRolPago)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtDecimoXIV)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtFondoReserva)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtVacaciones)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbDecXIV)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbDecXIII)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbFondoReserva)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbVacaciones)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtDecimoXIII)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtProvision1)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoRol)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtProvision2)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtProvision3)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtProvision4)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtProvAporte)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtAportable)).BeginInit();
			this.SuspendLayout();
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
			// cmbPersonal
			// 
			this.cmbPersonal.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbPersonal.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbPersonal.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsLiquidacion, "RolLiquidacion.idPersonal"));
			this.cmbPersonal.DataSource = this.cdvPersonal;
			ultraGridColumn1.Header.VisiblePosition = 21;
			ultraGridColumn1.Hidden = true;
			ultraGridColumn2.Header.VisiblePosition = 59;
			ultraGridColumn2.Hidden = true;
			ultraGridColumn3.Header.VisiblePosition = 19;
			ultraGridColumn3.Hidden = true;
			ultraGridColumn4.Header.VisiblePosition = 20;
			ultraGridColumn4.Hidden = true;
			ultraGridColumn5.Header.VisiblePosition = 18;
			ultraGridColumn5.Hidden = true;
			ultraGridColumn6.Header.VisiblePosition = 22;
			ultraGridColumn6.Hidden = true;
			ultraGridColumn7.Header.VisiblePosition = 36;
			ultraGridColumn7.Hidden = true;
			ultraGridColumn8.Header.VisiblePosition = 25;
			ultraGridColumn8.Hidden = true;
			ultraGridColumn9.Header.VisiblePosition = 0;
			ultraGridColumn9.Width = 364;
			ultraGridColumn10.Header.VisiblePosition = 3;
			ultraGridColumn10.Hidden = true;
			ultraGridColumn11.Header.VisiblePosition = 35;
			ultraGridColumn11.Hidden = true;
			ultraGridColumn12.Header.VisiblePosition = 24;
			ultraGridColumn12.Hidden = true;
			ultraGridColumn13.Header.VisiblePosition = 13;
			ultraGridColumn13.Hidden = true;
			ultraGridColumn14.Header.VisiblePosition = 32;
			ultraGridColumn14.Hidden = true;
			ultraGridColumn15.Header.VisiblePosition = 7;
			ultraGridColumn15.Hidden = true;
			ultraGridColumn16.Header.VisiblePosition = 6;
			ultraGridColumn16.Hidden = true;
			ultraGridColumn17.Header.VisiblePosition = 2;
			ultraGridColumn17.Hidden = true;
			ultraGridColumn18.Header.VisiblePosition = 29;
			ultraGridColumn18.Hidden = true;
			ultraGridColumn19.Header.VisiblePosition = 28;
			ultraGridColumn19.Hidden = true;
			ultraGridColumn20.Header.VisiblePosition = 23;
			ultraGridColumn20.Hidden = true;
			ultraGridColumn21.Header.VisiblePosition = 31;
			ultraGridColumn21.Hidden = true;
			ultraGridColumn22.Header.VisiblePosition = 10;
			ultraGridColumn22.Hidden = true;
			ultraGridColumn23.Header.VisiblePosition = 33;
			ultraGridColumn23.Hidden = true;
			ultraGridColumn24.Header.VisiblePosition = 14;
			ultraGridColumn24.Hidden = true;
			ultraGridColumn25.Header.VisiblePosition = 15;
			ultraGridColumn25.Hidden = true;
			ultraGridColumn26.Header.VisiblePosition = 17;
			ultraGridColumn26.Hidden = true;
			ultraGridColumn27.Header.VisiblePosition = 26;
			ultraGridColumn27.Hidden = true;
			ultraGridColumn28.Header.VisiblePosition = 5;
			ultraGridColumn28.Hidden = true;
			ultraGridColumn29.Header.VisiblePosition = 37;
			ultraGridColumn29.Hidden = true;
			ultraGridColumn30.Header.VisiblePosition = 38;
			ultraGridColumn30.Hidden = true;
			ultraGridColumn31.Header.VisiblePosition = 27;
			ultraGridColumn31.Hidden = true;
			ultraGridColumn32.Header.VisiblePosition = 30;
			ultraGridColumn32.Hidden = true;
			ultraGridColumn33.Header.VisiblePosition = 12;
			ultraGridColumn33.Hidden = true;
			ultraGridColumn34.Header.VisiblePosition = 9;
			ultraGridColumn34.Hidden = true;
			ultraGridColumn35.Header.VisiblePosition = 11;
			ultraGridColumn35.Hidden = true;
			ultraGridColumn36.Header.VisiblePosition = 16;
			ultraGridColumn36.Hidden = true;
			ultraGridColumn37.Header.VisiblePosition = 1;
			ultraGridColumn37.Hidden = true;
			ultraGridColumn38.Header.VisiblePosition = 34;
			ultraGridColumn38.Hidden = true;
			ultraGridColumn39.Header.VisiblePosition = 4;
			ultraGridColumn39.Hidden = true;
			ultraGridColumn40.Header.VisiblePosition = 8;
			ultraGridColumn40.Hidden = true;
			ultraGridColumn41.Header.VisiblePosition = 39;
			ultraGridColumn41.Hidden = true;
			ultraGridColumn42.Header.VisiblePosition = 40;
			ultraGridColumn42.Hidden = true;
			ultraGridColumn43.Header.VisiblePosition = 41;
			ultraGridColumn43.Hidden = true;
			ultraGridColumn44.Header.VisiblePosition = 42;
			ultraGridColumn44.Hidden = true;
			ultraGridColumn45.Header.VisiblePosition = 43;
			ultraGridColumn45.Hidden = true;
			ultraGridColumn46.Header.VisiblePosition = 44;
			ultraGridColumn46.Hidden = true;
			ultraGridColumn47.Header.VisiblePosition = 45;
			ultraGridColumn47.Hidden = true;
			ultraGridColumn48.Header.VisiblePosition = 46;
			ultraGridColumn48.Hidden = true;
			ultraGridColumn49.Header.VisiblePosition = 47;
			ultraGridColumn49.Hidden = true;
			ultraGridColumn50.Header.VisiblePosition = 48;
			ultraGridColumn50.Hidden = true;
			ultraGridColumn51.Header.VisiblePosition = 49;
			ultraGridColumn51.Hidden = true;
			ultraGridColumn52.Header.VisiblePosition = 50;
			ultraGridColumn52.Hidden = true;
			ultraGridColumn53.Header.VisiblePosition = 51;
			ultraGridColumn53.Hidden = true;
			ultraGridColumn54.Header.VisiblePosition = 52;
			ultraGridColumn54.Hidden = true;
			ultraGridColumn55.Header.VisiblePosition = 53;
			ultraGridColumn55.Hidden = true;
			ultraGridColumn56.Header.VisiblePosition = 54;
			ultraGridColumn56.Hidden = true;
			ultraGridColumn57.Header.VisiblePosition = 55;
			ultraGridColumn57.Hidden = true;
			ultraGridColumn58.Header.VisiblePosition = 56;
			ultraGridColumn58.Hidden = true;
			ultraGridColumn59.Header.VisiblePosition = 57;
			ultraGridColumn59.Hidden = true;
			ultraGridColumn60.Header.VisiblePosition = 58;
			ultraGridColumn60.Hidden = true;
			ultraGridColumn61.Header.VisiblePosition = 60;
			ultraGridColumn61.Hidden = true;
			ultraGridColumn62.Header.VisiblePosition = 61;
			ultraGridColumn62.Hidden = true;
			ultraGridColumn63.Header.VisiblePosition = 62;
			ultraGridColumn63.Hidden = true;
			ultraGridColumn64.Header.VisiblePosition = 63;
			ultraGridColumn64.Hidden = true;
			ultraGridColumn65.Header.VisiblePosition = 64;
			ultraGridColumn65.Hidden = true;
			ultraGridColumn66.Header.VisiblePosition = 65;
			ultraGridColumn66.Hidden = true;
			ultraGridColumn67.Header.VisiblePosition = 66;
			ultraGridColumn67.Hidden = true;
			ultraGridColumn68.Header.VisiblePosition = 67;
			ultraGridColumn68.Hidden = true;
			ultraGridColumn69.Header.VisiblePosition = 68;
			ultraGridColumn69.Hidden = true;
			ultraGridColumn70.Header.VisiblePosition = 69;
			ultraGridColumn70.Hidden = true;
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
			ultraGridColumn92.Hidden = true;
			ultraGridColumn93.Header.VisiblePosition = 92;
			ultraGridColumn93.Hidden = true;
			ultraGridColumn94.Header.VisiblePosition = 93;
			ultraGridColumn94.Hidden = true;
			ultraGridColumn95.Header.VisiblePosition = 94;
			ultraGridColumn95.Hidden = true;
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
			this.cmbPersonal.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbPersonal.Enabled = false;
			this.cmbPersonal.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbPersonal.Location = new System.Drawing.Point(152, 26);
			this.cmbPersonal.Name = "cmbPersonal";
			this.cmbPersonal.Size = new System.Drawing.Size(272, 21);
			this.cmbPersonal.TabIndex = 0;
			this.cmbPersonal.ValueMember = "idPersonal";
			// 
			// cdsLiquidacion
			// 
			this.cdsLiquidacion.BindingContextCtrl = this;
			this.cdsLiquidacion.DataLibrary = "LibPersonal";
			this.cdsLiquidacion.DataLibraryUrl = "";
			this.cdsLiquidacion.DataSetDef = "dsRolTablas";
			this.cdsLiquidacion.Locale = new System.Globalization.CultureInfo("es-EC");
			this.cdsLiquidacion.Name = "cdsLiquidacion";
			this.cdsLiquidacion.SchemaClassName = "LibPersonal.DataClass";
			this.cdsLiquidacion.SchemaDef = null;
			this.cdsLiquidacion.BeforeFill += new C1.Data.FillEventHandler(this.cdsLiquidacion_BeforeFill);
			// 
			// ultraLabel1
			// 
			this.ultraLabel1.AutoSize = true;
			this.ultraLabel1.Location = new System.Drawing.Point(32, 26);
			this.ultraLabel1.Name = "ultraLabel1";
			this.ultraLabel1.Size = new System.Drawing.Size(112, 15);
			this.ultraLabel1.TabIndex = 1;
			this.ultraLabel1.Text = "Nombre de Empleado";
			// 
			// cdsRolPago
			// 
			this.cdsRolPago.BindingContextCtrl = this;
			this.cdsRolPago.DataLibrary = "LibPersonal";
			this.cdsRolPago.DataLibraryUrl = "";
			this.cdsRolPago.DataSetDef = "dsRolGrupo";
			this.cdsRolPago.Locale = new System.Globalization.CultureInfo("es-EC");
			this.cdsRolPago.Name = "cdsRolPago";
			this.cdsRolPago.SchemaClassName = "LibPersonal.DataClass";
			this.cdsRolPago.SchemaDef = null;
			this.cdsRolPago.BeforeFill += new C1.Data.FillEventHandler(this.cdsRolPago_BeforeFill);
			// 
			// label1
			// 
			this.label1.AutoSize = true;
			this.label1.Location = new System.Drawing.Point(32, 95);
			this.label1.Name = "label1";
			this.label1.Size = new System.Drawing.Size(83, 16);
			this.label1.TabIndex = 2;
			this.label1.Text = "Decimo Tercero";
			// 
			// label2
			// 
			this.label2.AutoSize = true;
			this.label2.Location = new System.Drawing.Point(32, 129);
			this.label2.Name = "label2";
			this.label2.Size = new System.Drawing.Size(78, 16);
			this.label2.TabIndex = 3;
			this.label2.Text = "Decimo Cuarto";
			// 
			// label3
			// 
			this.label3.AutoSize = true;
			this.label3.Location = new System.Drawing.Point(32, 164);
			this.label3.Name = "label3";
			this.label3.Size = new System.Drawing.Size(78, 16);
			this.label3.TabIndex = 4;
			this.label3.Text = "Fondo Reserva";
			// 
			// label4
			// 
			this.label4.AutoSize = true;
			this.label4.Location = new System.Drawing.Point(32, 198);
			this.label4.Name = "label4";
			this.label4.Size = new System.Drawing.Size(59, 16);
			this.label4.TabIndex = 5;
			this.label4.Text = "Vacaciones";
			// 
			// txtDecimoXIV
			// 
			this.txtDecimoXIV.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsLiquidacion, "RolLiquidacion.DCuarto"));
			this.txtDecimoXIV.Enabled = false;
			this.txtDecimoXIV.FormatString = "#,##0.00";
			this.txtDecimoXIV.Location = new System.Drawing.Point(152, 129);
			this.txtDecimoXIV.Name = "txtDecimoXIV";
			this.txtDecimoXIV.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtDecimoXIV.PromptChar = ' ';
			this.txtDecimoXIV.Size = new System.Drawing.Size(136, 22);
			this.txtDecimoXIV.TabIndex = 7;
			// 
			// txtFondoReserva
			// 
			this.txtFondoReserva.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsLiquidacion, "RolLiquidacion.FondoR"));
			this.txtFondoReserva.Enabled = false;
			this.txtFondoReserva.FormatString = "#,##0.00";
			this.txtFondoReserva.Location = new System.Drawing.Point(152, 164);
			this.txtFondoReserva.Name = "txtFondoReserva";
			this.txtFondoReserva.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtFondoReserva.PromptChar = ' ';
			this.txtFondoReserva.Size = new System.Drawing.Size(136, 22);
			this.txtFondoReserva.TabIndex = 8;
			// 
			// txtVacaciones
			// 
			this.txtVacaciones.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsLiquidacion, "RolLiquidacion.Vacaciones"));
			this.txtVacaciones.Enabled = false;
			this.txtVacaciones.FormatString = "#,##0.00";
			this.txtVacaciones.Location = new System.Drawing.Point(152, 198);
			this.txtVacaciones.Name = "txtVacaciones";
			this.txtVacaciones.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtVacaciones.PromptChar = ' ';
			this.txtVacaciones.Size = new System.Drawing.Size(136, 22);
			this.txtVacaciones.TabIndex = 9;
			// 
			// cmbDecXIV
			// 
			this.cmbDecXIV.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsLiquidacion, "RolLiquidacion.DTerceroFecha"));
			this.cmbDecXIV.DateTime = new System.DateTime(2010, 2, 9, 0, 0, 0, 0);
			this.cmbDecXIV.Enabled = false;
			this.cmbDecXIV.FormatString = "dd/MMM/yyyy";
			this.cmbDecXIV.Location = new System.Drawing.Point(304, 129);
			this.cmbDecXIV.Name = "cmbDecXIV";
			this.cmbDecXIV.Size = new System.Drawing.Size(120, 22);
			this.cmbDecXIV.TabIndex = 11;
			this.cmbDecXIV.Value = new System.DateTime(2010, 2, 9, 0, 0, 0, 0);
			// 
			// cmbDecXIII
			// 
			this.cmbDecXIII.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsLiquidacion, "RolLiquidacion.AportableFecha"));
			this.cmbDecXIII.DateTime = new System.DateTime(2010, 2, 9, 0, 0, 0, 0);
			this.cmbDecXIII.Enabled = false;
			this.cmbDecXIII.FormatString = "dd/MMM/yyyy";
			this.cmbDecXIII.Location = new System.Drawing.Point(304, 95);
			this.cmbDecXIII.Name = "cmbDecXIII";
			this.cmbDecXIII.PromptChar = ' ';
			this.cmbDecXIII.Size = new System.Drawing.Size(120, 22);
			this.cmbDecXIII.TabIndex = 12;
			this.cmbDecXIII.Value = new System.DateTime(2010, 2, 9, 0, 0, 0, 0);
			// 
			// cmbFondoReserva
			// 
			this.cmbFondoReserva.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsLiquidacion, "RolLiquidacion.FondorRFecha"));
			this.cmbFondoReserva.DateTime = new System.DateTime(2010, 2, 9, 0, 0, 0, 0);
			this.cmbFondoReserva.Enabled = false;
			this.cmbFondoReserva.FormatString = "dd/MMM/yyyy";
			this.cmbFondoReserva.Location = new System.Drawing.Point(304, 164);
			this.cmbFondoReserva.Name = "cmbFondoReserva";
			this.cmbFondoReserva.Size = new System.Drawing.Size(120, 22);
			this.cmbFondoReserva.TabIndex = 13;
			this.cmbFondoReserva.Value = new System.DateTime(2010, 2, 9, 0, 0, 0, 0);
			// 
			// cmbVacaciones
			// 
			this.cmbVacaciones.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsLiquidacion, "RolLiquidacion.VacacionesFecha"));
			this.cmbVacaciones.DateTime = new System.DateTime(2010, 2, 9, 0, 0, 0, 0);
			this.cmbVacaciones.Enabled = false;
			this.cmbVacaciones.FormatString = "dd/MMM/yyyy";
			this.cmbVacaciones.Location = new System.Drawing.Point(304, 198);
			this.cmbVacaciones.Name = "cmbVacaciones";
			this.cmbVacaciones.Size = new System.Drawing.Size(120, 22);
			this.cmbVacaciones.TabIndex = 14;
			this.cmbVacaciones.Value = new System.DateTime(2010, 2, 9, 0, 0, 0, 0);
			// 
			// btAceptar
			// 
			this.btAceptar.Enabled = false;
			this.btAceptar.Location = new System.Drawing.Point(32, 241);
			this.btAceptar.Name = "btAceptar";
			this.btAceptar.Size = new System.Drawing.Size(75, 25);
			this.btAceptar.TabIndex = 15;
			this.btAceptar.Text = "&Calcular";
			this.btAceptar.Click += new System.EventHandler(this.btAceptar_Click);
			// 
			// btCancelar
			// 
			this.btCancelar.DialogResult = System.Windows.Forms.DialogResult.Cancel;
			this.btCancelar.Location = new System.Drawing.Point(344, 241);
			this.btCancelar.Name = "btCancelar";
			this.btCancelar.Size = new System.Drawing.Size(75, 25);
			this.btCancelar.TabIndex = 16;
			this.btCancelar.Text = "&Salir";
			this.btCancelar.Visible = false;
			this.btCancelar.Click += new System.EventHandler(this.btCancelar_Click);
			// 
			// txtDecimoXIII
			// 
			this.txtDecimoXIII.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsLiquidacion, "RolLiquidacion.DTercero"));
			this.txtDecimoXIII.Enabled = false;
			this.txtDecimoXIII.FormatString = "#,##0.00";
			this.txtDecimoXIII.Location = new System.Drawing.Point(152, 95);
			this.txtDecimoXIII.Name = "txtDecimoXIII";
			this.txtDecimoXIII.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtDecimoXIII.PromptChar = ' ';
			this.txtDecimoXIII.Size = new System.Drawing.Size(136, 22);
			this.txtDecimoXIII.TabIndex = 19;
			// 
			// errorProvider
			// 
			this.errorProvider.ContainerControl = this;
			// 
			// btImprimir
			// 
			this.btImprimir.Location = new System.Drawing.Point(128, 241);
			this.btImprimir.Name = "btImprimir";
			this.btImprimir.Size = new System.Drawing.Size(75, 25);
			this.btImprimir.TabIndex = 20;
			this.btImprimir.Text = "&Imprimir";
			this.btImprimir.Click += new System.EventHandler(this.btImprimir_Click);
			// 
			// txtProvision1
			// 
			this.txtProvision1.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsLiquidacion, "RolLiquidacion.Provision1"));
			this.txtProvision1.Enabled = false;
			this.txtProvision1.Location = new System.Drawing.Point(440, 95);
			this.txtProvision1.Name = "txtProvision1";
			this.txtProvision1.Size = new System.Drawing.Size(80, 22);
			this.txtProvision1.TabIndex = 21;
			this.txtProvision1.Text = "Provision1";
			// 
			// cdsSeteoRol
			// 
			this.cdsSeteoRol.BindingContextCtrl = this;
			this.cdsSeteoRol.DataLibrary = "LibPersonal";
			this.cdsSeteoRol.DataLibraryUrl = "";
			this.cdsSeteoRol.DataSetDef = "dsSeteoRol";
			this.cdsSeteoRol.Locale = new System.Globalization.CultureInfo("es-EC");
			this.cdsSeteoRol.Name = "cdsSeteoRol";
			this.cdsSeteoRol.SchemaClassName = "LibPersonal.DataClass";
			this.cdsSeteoRol.SchemaDef = null;
			this.cdsSeteoRol.BeforeFill += new C1.Data.FillEventHandler(this.cdsSeteoRol_BeforeFill);
			// 
			// txtProvision2
			// 
			this.txtProvision2.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsLiquidacion, "RolLiquidacion.Provision2"));
			this.txtProvision2.Enabled = false;
			this.txtProvision2.Location = new System.Drawing.Point(440, 129);
			this.txtProvision2.Name = "txtProvision2";
			this.txtProvision2.Size = new System.Drawing.Size(80, 22);
			this.txtProvision2.TabIndex = 22;
			this.txtProvision2.Text = "Provision2";
			this.txtProvision2.ValueChanged += new System.EventHandler(this.txtProvision2_ValueChanged);
			// 
			// txtProvision3
			// 
			this.txtProvision3.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsLiquidacion, "RolLiquidacion.Provision3"));
			this.txtProvision3.Enabled = false;
			this.txtProvision3.Location = new System.Drawing.Point(440, 164);
			this.txtProvision3.Name = "txtProvision3";
			this.txtProvision3.Size = new System.Drawing.Size(80, 22);
			this.txtProvision3.TabIndex = 23;
			this.txtProvision3.Text = "Provision3";
			// 
			// txtProvision4
			// 
			this.txtProvision4.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsLiquidacion, "RolLiquidacion.Provision4"));
			this.txtProvision4.Enabled = false;
			this.txtProvision4.Location = new System.Drawing.Point(440, 198);
			this.txtProvision4.Name = "txtProvision4";
			this.txtProvision4.Size = new System.Drawing.Size(80, 22);
			this.txtProvision4.TabIndex = 24;
			this.txtProvision4.Text = "Provision4";
			// 
			// label5
			// 
			this.label5.AutoSize = true;
			this.label5.Location = new System.Drawing.Point(32, 60);
			this.label5.Name = "label5";
			this.label5.Size = new System.Drawing.Size(52, 16);
			this.label5.TabIndex = 25;
			this.label5.Text = "Aportable";
			// 
			// txtProvAporte
			// 
			this.txtProvAporte.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsLiquidacion, "RolLiquidacion.ProvAporte"));
			this.txtProvAporte.Enabled = false;
			this.txtProvAporte.Location = new System.Drawing.Point(440, 60);
			this.txtProvAporte.Name = "txtProvAporte";
			this.txtProvAporte.Size = new System.Drawing.Size(80, 22);
			this.txtProvAporte.TabIndex = 26;
			this.txtProvAporte.Text = "Valor6";
			// 
			// txtAportable
			// 
			this.txtAportable.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsLiquidacion, "RolLiquidacion.Aportable"));
			this.txtAportable.Enabled = false;
			this.txtAportable.FormatString = "#,##0.00";
			this.txtAportable.Location = new System.Drawing.Point(152, 60);
			this.txtAportable.Name = "txtAportable";
			this.txtAportable.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtAportable.PromptChar = ' ';
			this.txtAportable.Size = new System.Drawing.Size(136, 22);
			this.txtAportable.TabIndex = 27;
			// 
			// barraDato1
			// 
			this.barraDato1.BarraMovimiento = true;
			this.barraDato1.DataMember = "RolLiquidacion";
			this.barraDato1.DataNombreId = "idLiquidacion";
			this.barraDato1.DataOrden = "";
			this.barraDato1.DataSource = this.cdsLiquidacion;
			this.barraDato1.DataTabla = "RolLiquidacion";
			this.barraDato1.DataTablaHija = "";
			this.barraDato1.Location = new System.Drawing.Point(24, 276);
			this.barraDato1.Name = "barraDato1";
			this.barraDato1.Size = new System.Drawing.Size(304, 26);
			this.barraDato1.TabIndex = 28;
			this.barraDato1.VisibleBorrar = true;
			this.barraDato1.VisibleBuscar = false;
			this.barraDato1.VisibleEditar = true;
			this.barraDato1.VisibleImprimir = false;
			this.barraDato1.VisibleNuevo = true;
			this.barraDato1.Buscar += new System.EventHandler(this.barraDato1_Buscar);
			this.barraDato1.Load += new System.EventHandler(this.barraDato1_Load);
			this.barraDato1.Borrar += new System.EventHandler(this.barraDato1_Borrar);
			this.barraDato1.Imprimir += new System.EventHandler(this.barraDato1_Imprimir);
			this.barraDato1.Refresca += new System.EventHandler(this.barraDato1_Refresca);
			this.barraDato1.Crear += new System.EventHandler(this.barraDato1_Crear);
			this.barraDato1.Deshacer += new System.EventHandler(this.barraDato1_Deshacer);
			this.barraDato1.Copiar += new System.EventHandler(this.barraDato1_Copiar);
			this.barraDato1.Impresora += new System.EventHandler(this.barraDato1_Impresora);
			this.barraDato1.Editar += new System.EventHandler(this.barraDato1_Editar);
			this.barraDato1.Grabar += new System.EventHandler(this.barraDato1_Grabar);
			// 
			// RolLiquidacion
			// 
			this.AcceptButton = this.btAceptar;
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 14);
			this.CancelButton = this.btCancelar;
			this.ClientSize = new System.Drawing.Size(528, 315);
			this.Controls.Add(this.barraDato1);
			this.Controls.Add(this.txtAportable);
			this.Controls.Add(this.txtProvAporte);
			this.Controls.Add(this.label5);
			this.Controls.Add(this.txtProvision4);
			this.Controls.Add(this.txtProvision3);
			this.Controls.Add(this.txtProvision2);
			this.Controls.Add(this.txtProvision1);
			this.Controls.Add(this.btImprimir);
			this.Controls.Add(this.txtDecimoXIII);
			this.Controls.Add(this.btCancelar);
			this.Controls.Add(this.btAceptar);
			this.Controls.Add(this.cmbVacaciones);
			this.Controls.Add(this.cmbFondoReserva);
			this.Controls.Add(this.cmbDecXIII);
			this.Controls.Add(this.cmbDecXIV);
			this.Controls.Add(this.txtVacaciones);
			this.Controls.Add(this.txtFondoReserva);
			this.Controls.Add(this.txtDecimoXIV);
			this.Controls.Add(this.label4);
			this.Controls.Add(this.label3);
			this.Controls.Add(this.label2);
			this.Controls.Add(this.label1);
			this.Controls.Add(this.ultraLabel1);
			this.Controls.Add(this.cmbPersonal);
			this.MaximizeBox = false;
			this.Name = "RolLiquidacion";
			this.Text = "Liquidacion";
			this.Closing += new System.ComponentModel.CancelEventHandler(this.RolLiquidacion_Closing);
			this.Load += new System.EventHandler(this.RolLiquidacion_Load);
			((System.ComponentModel.ISupportInitialize)(this.cdsPersonal)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdvPersonal)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbPersonal)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsLiquidacion)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsRolPago)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtDecimoXIV)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtFondoReserva)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtVacaciones)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbDecXIV)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbDecXIII)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbFondoReserva)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbVacaciones)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtDecimoXIII)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtProvision1)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoRol)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtProvision2)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtProvision3)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtProvision4)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtProvAporte)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtAportable)).EndInit();
			this.ResumeLayout(false);

		}
		#endregion


		private void RolLiquidacion_Load(object sender, System.EventArgs e)
		{
			string stAudita = "Exec AuditaCrear 53, 6, 'Rol Liquidacion'";
			Funcion.EjecutaSQL(cdsPersonal, stAudita, true);

			IFormatProvider us = new CultureInfo("en-US", true);
			DateTime dtFecha = DateTime.Today;
			if (dtFecha.Month != 12)
				cmbDecXIII.Value = DateTime.Parse("12/01/" + dtFecha.AddYears(-1).Year.ToString(), us);
			else
				cmbDecXIII.Value = DateTime.Parse("12/01/" + dtFecha.AddYears(-1).Year.ToString(), us);

			if (dtFecha.Month <= 9)
				cmbDecXIV.Value = DateTime.Parse("09/01/" + dtFecha.AddYears(-1).Year.ToString(), us);
			else
				cmbDecXIV.Value = DateTime.Parse("09/01/" + dtFecha.Year.ToString(), us);

			if (dtFecha.Month <= 7)
				cmbFondoReserva.Value = DateTime.Parse("07/01/" + dtFecha.AddYears(-1).Year.ToString(), us);
			else
				cmbFondoReserva.Value = DateTime.Parse("07/01/" + dtFecha.Year.ToString(), us);

			this.barraDato1.PosUltima();
		}

		private void BorraErrores()
		{
			errorProvider.SetError(cmbDecXIII, "");
			errorProvider.SetError(cmbDecXIV, "");
			errorProvider.SetError(cmbFondoReserva, "");
			errorProvider.SetError(cmbVacaciones, "");
			errorProvider.SetError(cmbPersonal, "");
		}

		private bool Verifica()
		{
			bool bOk = true;
			BorraErrores();
			if (cmbDecXIII.Value == null)
			{
				errorProvider.SetError(cmbDecXIII, "Ingrese Fecha de Inicio de Calculo de XIII");
				bOk = false;
			}
			if (cmbDecXIV.Value == null)
			{
				errorProvider.SetError(cmbDecXIV, "Ingrese Fecha de Inicio de Calculo de XIV");
				bOk = false;
			}
			if (cmbFondoReserva.Value == null)
			{
				errorProvider.SetError(cmbFondoReserva, "Ingrese Fecha de Inicio de Calculo de Fondo de Reserva");
				bOk = false;
			}
			if (cmbVacaciones.Value == null)
			{
				errorProvider.SetError(cmbVacaciones, "Ingrese Fecha de Inicio de Calculo de Vacaciones");
				bOk = false;
			}
			string stSelect = "select ISNULL(p.FechaSalida, '19000101') from Personal p where p.idPersonal = " 
				+ cmbPersonal.Value.ToString();
			dtHoy = Funcion.dtEscalarSQL(cdsPersonal, stSelect);
			if (dtHoy < DateTime.Today.AddYears(-50))
			{
				errorProvider.SetError(cmbPersonal, "Ingrese Fecha de Salida de la persona");
				bOk = false;
			}
			if (cmbPersonal.Text.Trim().Length < 1)
			{
				errorProvider.SetError(cmbPersonal, "Ingrese Empleado");
				bOk = false;
			}
			return bOk;
		}

		DateTime dtHoy;
		private double CalculoProvision(string stProvision, DateTime dtFecha)
		{
			string stSelect = string.Format("Exec RolLiquidaPersona '{0}', {1}, '{2}'", 
				dtFecha.ToString("yyyyMMdd"), cmbPersonal.Value, stProvision);
			double dValor = Funcion.dEscalarSQL(cdsRolPago, stSelect, false);
			return dValor;
		}

		private void btAceptar_Click(object sender, System.EventArgs e)
		{
			if (!Verifica()) return;
			DateTime dtFecha = (DateTime) cmbDecXIII.Value;
			txtDecimoXIII.Value = CalculoProvision(txtProvision1.Text, dtFecha);
			txtAportable.Value = CalculoProvision(txtProvAporte.Text, dtFecha);
			dtFecha = (DateTime) cmbDecXIV.Value;
			txtDecimoXIV.Value = CalculoProvision(txtProvision2.Text, dtFecha);
			dtFecha = (DateTime) cmbFondoReserva.Value;
			txtFondoReserva.Value = CalculoProvision(txtProvision3.Text, dtFecha);
			dtFecha = (DateTime) cmbVacaciones.Value;
			txtVacaciones.Value = CalculoProvision(txtProvision4.Text, dtFecha);
		}

		private void cdsPersonal_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			this.cdsPersonal.Schema.Connections[0].ConnectionString = 
				miFuncion.AccesoBase(MenuLatinium.stDirRoles);
		}

		private void cdsRolPago_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			this.cdsRolPago.Schema.Connections[0].ConnectionString = 
				miFuncion.AccesoBase(MenuLatinium.stDirRoles);
		}

		private void btCancelar_Click(object sender, System.EventArgs e)
		{
			Close();
		}

		private void btImprimir_Click(object sender, System.EventArgs e)
		{
			if (!Verifica()) return;
			Cursor = Cursors.WaitCursor;
			int idPersonal = (int) cmbPersonal.Value;
			Reporte miReporte = new Reporte("RolLiquidacion.rpt", 
				"{Personal.idPersonal} = " + idPersonal.ToString());

			#region Parametros
			ParameterFields paramFields = new ParameterFields ();
			ParameterField paramField = new ParameterField ();
			ParameterDiscreteValue discreteVal = new ParameterDiscreteValue ();
			paramField.ParameterFieldName = "dDecimoXIII";
			discreteVal.Value = (double) txtDecimoXIII.Value;
			paramField.CurrentValues.Add (discreteVal);
			paramFields.Add (paramField);

			ParameterDiscreteValue discreteVal2 = new ParameterDiscreteValue ();
			ParameterField paramField2 = new ParameterField ();
			paramField2.ParameterFieldName = "dDecimoXIV";
			discreteVal2.Value = (double) txtDecimoXIV.Value;
			paramField2.CurrentValues.Add (discreteVal2);
			paramFields.Add (paramField2);

			ParameterDiscreteValue discreteVal3 = new ParameterDiscreteValue ();
			ParameterField paramField3 = new ParameterField ();
			paramField3.ParameterFieldName = "dFondoReserva";
			discreteVal3.Value = (double) txtFondoReserva.Value;
			paramField3.CurrentValues.Add (discreteVal3);
			paramFields.Add (paramField3);

			ParameterDiscreteValue discreteVal4 = new ParameterDiscreteValue ();
			ParameterField paramField4 = new ParameterField ();
			paramField4.ParameterFieldName = "dVacaciones";
			discreteVal4.Value = (double) txtVacaciones.Value;
			paramField4.CurrentValues.Add (discreteVal4);
			paramFields.Add (paramField4);

			// Fechas
			ParameterDiscreteValue discreteVal5 = new ParameterDiscreteValue ();
			ParameterField paramField5 = new ParameterField ();
			paramField5.ParameterFieldName = "tDecimoXIII";
			discreteVal5.Value = (DateTime) cmbDecXIII.Value;
			paramField5.CurrentValues.Add (discreteVal5);
			paramFields.Add (paramField5);

			ParameterDiscreteValue discreteVal6 = new ParameterDiscreteValue ();
			ParameterField paramField6 = new ParameterField ();
			paramField6.ParameterFieldName = "tDecimoXIV";
			discreteVal6.Value = (DateTime) cmbDecXIV.Value;
			paramField6.CurrentValues.Add (discreteVal6);
			paramFields.Add (paramField6);

			ParameterDiscreteValue discreteVal7 = new ParameterDiscreteValue ();
			ParameterField paramField7 = new ParameterField ();
			paramField7.ParameterFieldName = "tFondoReserva";
			discreteVal7.Value = (DateTime) cmbFondoReserva.Value;
			paramField7.CurrentValues.Add (discreteVal7);
			paramFields.Add (paramField7);

			ParameterDiscreteValue discreteVal8 = new ParameterDiscreteValue ();
			ParameterField paramField8 = new ParameterField ();
			paramField8.ParameterFieldName = "tVacaciones";
			discreteVal8.Value = (DateTime) cmbVacaciones.Value;
			paramField8.CurrentValues.Add (discreteVal8);
			paramFields.Add (paramField8);

			ParameterDiscreteValue discreteVal9 = new ParameterDiscreteValue ();
			ParameterField paramField9 = new ParameterField ();
			paramField9.ParameterFieldName = "dAportable";
			discreteVal9.Value = Double.Parse(txtAportable.Value.ToString());
			paramField9.CurrentValues.Add (discreteVal9);
			paramFields.Add (paramField9);

			#endregion Parametros

			DateTime dtHasta = dtHoy;
			miReporte.MdiParent = this.MdiParent;
			miReporte.crVista.ParameterFieldInfo = paramFields;
			miReporte.Titulo("Liquidacion de " + cmbPersonal.Text + 
				" al " + dtHasta.ToString("dd/MMM/yyyy"));
			miReporte.Show();
			Cursor = Cursors.Default;
		}

		private void RolLiquidacion_Closing(object sender, System.ComponentModel.CancelEventArgs e)
		{
			string stAudita = "Exec AuditaCrear 53, 7, 'Rol Liquidacion'";
			Funcion.EjecutaSQL(cdsPersonal, stAudita, true);
		}

		private void cdsSeteoRol_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			this.cdsSeteoRol.Schema.Connections[0].ConnectionString = 
				miFuncion.AccesoBase(MenuLatinium.stDirRoles);
		}

		private void cdsLiquidacion_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			this.cdsLiquidacion.Schema.Connections[0].ConnectionString = 
				miFuncion.AccesoBase(MenuLatinium.stDirRoles);
		}

		private void barraDato1_Copiar(object sender, System.EventArgs e)
		{
		
		}

		private void barraDato1_Crear(object sender, System.EventArgs e)
		{
		barraDato1.CrearRegistro();
			btAceptar.Enabled = true;
			
		}

		private void barraDato1_Deshacer(object sender, System.EventArgs e)
		{
		barraDato1.DeshacerRegistro();
			btAceptar.Enabled = false;
		}

		private void barraDato1_Borrar(object sender, System.EventArgs e)
		{
			if (MessageBox.Show("¿Desea Borrar esta Liquidación?", "Confirmación",
				MessageBoxButtons.YesNo, MessageBoxIcon.Warning,
				MessageBoxDefaultButton.Button2) == DialogResult.No) 
				return;

		barraDato1.BorraRegistro();
			string stAudita = "Exec AuditaCrear 61, 4, 'Liquidacion'";
			Funcion.EjecutaSQL(cdsLiquidacion, stAudita, true);
		}

		private void barraDato1_Buscar(object sender, System.EventArgs e)
		{
		
		}

		private void barraDato1_Editar(object sender, System.EventArgs e)
		{
			barraDato1.EditarRegistro();
			btAceptar.Enabled = true;
		}

		private void barraDato1_Grabar(object sender, System.EventArgs e)
		{
			barraDato1.GrabaRegistro();
			btAceptar.Enabled = false;
			this.barraDato1.ProximoId(5);
			string stAudita = "Exec AuditaCrear 61, 1, 'Liquidacion'";
			Funcion.EjecutaSQL(cdsLiquidacion, stAudita, true);
		}

		private void barraDato1_Impresora(object sender, System.EventArgs e)
		{
		
		}

		private void barraDato1_Imprimir(object sender, System.EventArgs e)
		{
		
		}

		private void txtProvision2_ValueChanged(object sender, System.EventArgs e)
		{
		
		}

		private void barraDato1_Refresca(object sender, System.EventArgs e)
		{
		
		}

		private void barraDato1_Load(object sender, System.EventArgs e)
		{
		
		}
	}
}
