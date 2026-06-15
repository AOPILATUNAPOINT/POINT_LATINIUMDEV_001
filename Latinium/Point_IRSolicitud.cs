using System;
using System.Data;
using System.Data.SqlClient;
using C1.Data;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using LibContabilidad.DataObjects;
using CrystalDecisions.Shared;
using Infragistics.Win.UltraWinGrid;

namespace Latinium
{
	/// <summary>
	/// Descripción breve de Point_IRSolicitud.
	/// </summary>
	public class Point_IRSolicitud : System.Windows.Forms.Form
	{
		private System.Windows.Forms.GroupBox groupBox2;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbTipoEstado;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbEstado;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbCalificacion;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbCallCenter;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbInstitucion;
		private Infragistics.Win.Misc.UltraButton btGarante;
		private System.Windows.Forms.Label label28;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor ultraTextEditor3;
		private System.Windows.Forms.Label label27;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtNumTarjeta;
		private System.Windows.Forms.Label label26;
		private System.Windows.Forms.Label label24;
		private System.Windows.Forms.Label label23;
		private System.Windows.Forms.Label label22;
		private System.Windows.Forms.Label label21;
		private System.Windows.Forms.Label label20;
		private System.Windows.Forms.Label label19;
		private System.Windows.Forms.Label label18;
		private System.Windows.Forms.Label label17;
		private System.Windows.Forms.Label label25;
		private System.Windows.Forms.GroupBox groupBox1;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtRefUbicacion;
		private System.Windows.Forms.Label label16;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtNumTelefono;
		private System.Windows.Forms.Label label15;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtCalleTransversal;
		private System.Windows.Forms.Label label14;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtNumCasa;
		private System.Windows.Forms.Label label13;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbBarrio;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbCanton;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbProvincia;
		private System.Windows.Forms.Label label11;
		private System.Windows.Forms.Label label10;
		private System.Windows.Forms.Label label9;
		private System.Windows.Forms.Label label8;
		private System.Windows.Forms.Label label12;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtAvPrincipal;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtSegundoNombre;
		private System.Windows.Forms.Label label7;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtPrimerNombre;
		private System.Windows.Forms.Label label6;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtApeMaterno;
		private System.Windows.Forms.Label label5;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtApePaterno;
		private System.Windows.Forms.Label label4;
		private System.Windows.Forms.Label label3;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbTipoDocumento;
		private System.Windows.Forms.Label label2;
		private System.Windows.Forms.Label label1;
		private C1.Data.C1DataSet cdsSolcitud;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbEntFinanciera;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbTipTarjeta;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbLocal;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbVendedor;
		private DevExpress.XtraEditors.RadioGroup rgGarante;
		private System.Windows.Forms.ErrorProvider errorProvider;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtNumDocumento;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtidSolicitud;
		private BarraDatoSQL.BarraDatoSQL barraDato1;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbParroquia;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtCuota;
		/// <summary>
		/// Variable del diseñador requerida.
		/// </summary>
		private System.ComponentModel.Container components = null;

		public Point_IRSolicitud()
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
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn79 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("gRebajaIIIEdad");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn80 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("IngGravOE");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn81 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("DeduccionOE");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn82 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("OtrasRebOE");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn83 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("ImpRentaAsumido");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn84 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("RentaRetenidaOE");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn85 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("RentaRetOE");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn86 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("TipoSangre");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn87 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("RegEntrada");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn88 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("RegSalida");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn89 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idTipoContrato");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn90 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("E_Profesor");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn91 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("CodVendedor");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn92 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Clave");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn93 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idSupervisor");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn94 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("bSupervisor");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn95 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Fotos");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn96 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Dias360");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn97 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("DiasXIII");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn98 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idJefe");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn99 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("NumConadis");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn100 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("PorcDiscapacidad");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn101 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("PoseeDiscapacidad");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn102 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("TipoDiscapacidad");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn103 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("TipoVivienda");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn104 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Vehiculo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn105 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("CodigoUsuario");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn106 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Codigo");
			Infragistics.Win.Appearance appearance2 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand2 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Bodega", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn107 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("IdBodega");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn108 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn109 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Bodega");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn110 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Usuario");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn111 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCiudad");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn112 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idTipoFactura");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn113 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Codigo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn114 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idProyecto");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn115 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Activo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn116 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Sel");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn117 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Transferencia");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn118 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Direccion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn119 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Establecimiento");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn120 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("PuntoEmision");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn121 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Factura");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn122 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Telefono");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn123 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Email");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn124 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("CentroCosto");
			Infragistics.Win.Appearance appearance3 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand3 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Cre_TipoTarjeta", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn125 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idTipoTarjeta");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn126 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Codigo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn127 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.Appearance appearance4 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance5 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand4 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Cre_EtdoSolicitud", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn128 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idEstado");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn129 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Codigo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn130 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.Appearance appearance6 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand5 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Cre_TipoCalificacion", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn131 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idTipoCalificacion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn132 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Codigo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn133 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.Appearance appearance7 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand6 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Cre_CallCenter", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn134 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCallCenter");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn135 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Codigo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn136 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.Appearance appearance8 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand7 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Cre_Instituciones", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn137 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idInstitucion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn138 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Codigo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn139 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Descripcion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn140 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("PctDcto");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn141 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("PctRecargoTC");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn142 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Tarifa");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn143 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("EnviarBanco");
			Infragistics.Win.Appearance appearance9 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance10 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance11 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance12 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance13 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance14 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance15 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance16 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance17 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand8 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Cre_Barrio", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn144 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idBarrio");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn145 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Codigo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn146 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.Appearance appearance18 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance19 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand9 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Cre_Provincia", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn147 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idProvincia");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn148 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Codigo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn149 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.Appearance appearance20 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance21 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance22 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance23 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance24 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance25 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand10 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Cre_TipoDocumento", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn150 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idTipoDoc");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn151 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Codigo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn152 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.Appearance appearance26 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand11 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Cre_EntidadFinanciera", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn153 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idEntidadFinanciera");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn154 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Codigo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn155 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.Appearance appearance27 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance28 = new Infragistics.Win.Appearance();
			this.groupBox2 = new System.Windows.Forms.GroupBox();
			this.rgGarante = new DevExpress.XtraEditors.RadioGroup();
			this.cdsSolcitud = new C1.Data.C1DataSet();
			this.cmbVendedor = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.cmbLocal = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.cmbTipTarjeta = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.cmbTipoEstado = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.cmbEstado = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.cmbCalificacion = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.cmbCallCenter = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.cmbInstitucion = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.btGarante = new Infragistics.Win.Misc.UltraButton();
			this.label28 = new System.Windows.Forms.Label();
			this.ultraTextEditor3 = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.label27 = new System.Windows.Forms.Label();
			this.txtNumTarjeta = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.label26 = new System.Windows.Forms.Label();
			this.label24 = new System.Windows.Forms.Label();
			this.label23 = new System.Windows.Forms.Label();
			this.label22 = new System.Windows.Forms.Label();
			this.label21 = new System.Windows.Forms.Label();
			this.label20 = new System.Windows.Forms.Label();
			this.label19 = new System.Windows.Forms.Label();
			this.label18 = new System.Windows.Forms.Label();
			this.label17 = new System.Windows.Forms.Label();
			this.label25 = new System.Windows.Forms.Label();
			this.txtCuota = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.groupBox1 = new System.Windows.Forms.GroupBox();
			this.cmbParroquia = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.txtRefUbicacion = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.label16 = new System.Windows.Forms.Label();
			this.txtNumTelefono = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.label15 = new System.Windows.Forms.Label();
			this.txtCalleTransversal = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.label14 = new System.Windows.Forms.Label();
			this.txtNumCasa = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.label13 = new System.Windows.Forms.Label();
			this.cmbBarrio = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.cmbCanton = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.cmbProvincia = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.label11 = new System.Windows.Forms.Label();
			this.label10 = new System.Windows.Forms.Label();
			this.label9 = new System.Windows.Forms.Label();
			this.label8 = new System.Windows.Forms.Label();
			this.label12 = new System.Windows.Forms.Label();
			this.txtAvPrincipal = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.txtSegundoNombre = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.label7 = new System.Windows.Forms.Label();
			this.txtPrimerNombre = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.label6 = new System.Windows.Forms.Label();
			this.txtApeMaterno = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.label5 = new System.Windows.Forms.Label();
			this.txtApePaterno = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.label4 = new System.Windows.Forms.Label();
			this.label3 = new System.Windows.Forms.Label();
			this.cmbTipoDocumento = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.label2 = new System.Windows.Forms.Label();
			this.label1 = new System.Windows.Forms.Label();
			this.cmbEntFinanciera = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.errorProvider = new System.Windows.Forms.ErrorProvider();
			this.txtNumDocumento = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.txtidSolicitud = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.barraDato1 = new BarraDatoSQL.BarraDatoSQL();
			this.groupBox2.SuspendLayout();
			((System.ComponentModel.ISupportInitialize)(this.rgGarante.Properties)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSolcitud)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbVendedor)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbLocal)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbTipTarjeta)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbTipoEstado)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbEstado)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbCalificacion)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbCallCenter)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbInstitucion)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraTextEditor3)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNumTarjeta)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtCuota)).BeginInit();
			this.groupBox1.SuspendLayout();
			((System.ComponentModel.ISupportInitialize)(this.cmbParroquia)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtRefUbicacion)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNumTelefono)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtCalleTransversal)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNumCasa)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbBarrio)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbCanton)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbProvincia)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtAvPrincipal)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtSegundoNombre)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtPrimerNombre)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtApeMaterno)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtApePaterno)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbTipoDocumento)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbEntFinanciera)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNumDocumento)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtidSolicitud)).BeginInit();
			this.SuspendLayout();
			// 
			// groupBox2
			// 
			this.groupBox2.Controls.Add(this.rgGarante);
			this.groupBox2.Controls.Add(this.cmbVendedor);
			this.groupBox2.Controls.Add(this.cmbLocal);
			this.groupBox2.Controls.Add(this.cmbTipTarjeta);
			this.groupBox2.Controls.Add(this.cmbTipoEstado);
			this.groupBox2.Controls.Add(this.cmbEstado);
			this.groupBox2.Controls.Add(this.cmbCalificacion);
			this.groupBox2.Controls.Add(this.cmbCallCenter);
			this.groupBox2.Controls.Add(this.cmbInstitucion);
			this.groupBox2.Controls.Add(this.btGarante);
			this.groupBox2.Controls.Add(this.label28);
			this.groupBox2.Controls.Add(this.ultraTextEditor3);
			this.groupBox2.Controls.Add(this.label27);
			this.groupBox2.Controls.Add(this.txtNumTarjeta);
			this.groupBox2.Controls.Add(this.label26);
			this.groupBox2.Controls.Add(this.label24);
			this.groupBox2.Controls.Add(this.label23);
			this.groupBox2.Controls.Add(this.label22);
			this.groupBox2.Controls.Add(this.label21);
			this.groupBox2.Controls.Add(this.label20);
			this.groupBox2.Controls.Add(this.label19);
			this.groupBox2.Controls.Add(this.label18);
			this.groupBox2.Controls.Add(this.label17);
			this.groupBox2.Controls.Add(this.label25);
			this.groupBox2.Controls.Add(this.txtCuota);
			this.groupBox2.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.groupBox2.Location = new System.Drawing.Point(8, 256);
			this.groupBox2.Name = "groupBox2";
			this.groupBox2.Size = new System.Drawing.Size(712, 152);
			this.groupBox2.TabIndex = 55;
			this.groupBox2.TabStop = false;
			this.groupBox2.Text = "Datos Solicitud";
			this.groupBox2.Enter += new System.EventHandler(this.groupBox2_Enter);
			// 
			// rgGarante
			// 
			this.rgGarante.DataBindings.Add(new System.Windows.Forms.Binding("EditValue", this.cdsSolcitud, "Cre_Solicitud.Garante"));
			this.rgGarante.Location = new System.Drawing.Point(521, 80);
			this.rgGarante.Name = "rgGarante";
			// 
			// rgGarante.Properties
			// 
			this.rgGarante.Properties.Appearance.BackColor = System.Drawing.Color.Gainsboro;
			this.rgGarante.Properties.Appearance.BackColor2 = System.Drawing.Color.White;
			this.rgGarante.Properties.Appearance.GradientMode = System.Drawing.Drawing2D.LinearGradientMode.Vertical;
			this.rgGarante.Properties.Appearance.Options.UseBackColor = true;
			this.rgGarante.Properties.BorderStyle = DevExpress.XtraEditors.Controls.BorderStyles.Simple;
			this.rgGarante.Properties.Items.AddRange(new DevExpress.XtraEditors.Controls.RadioGroupItem[] {
																																																			new DevExpress.XtraEditors.Controls.RadioGroupItem(0, "No"),
																																																			new DevExpress.XtraEditors.Controls.RadioGroupItem(1, "Si")});
			this.rgGarante.Size = new System.Drawing.Size(103, 24);
			this.rgGarante.TabIndex = 27;
			// 
			// cdsSolcitud
			// 
			this.cdsSolcitud.BindingContextCtrl = this;
			this.cdsSolcitud.DataLibrary = "LibFacturacion";
			this.cdsSolcitud.DataLibraryUrl = "";
			this.cdsSolcitud.DataSetDef = "dsPointCredito";
			this.cdsSolcitud.Locale = new System.Globalization.CultureInfo("es-EC");
			this.cdsSolcitud.Name = "cdsSolcitud";
			this.cdsSolcitud.SchemaClassName = "LibFacturacion.DataClass";
			this.cdsSolcitud.SchemaDef = null;
			this.cdsSolcitud.BeforeFill += new C1.Data.FillEventHandler(this.cdsSolcitud_BeforeFill);
			// 
			// cmbVendedor
			// 
			appearance1.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(224)), ((System.Byte)(224)), ((System.Byte)(224)));
			appearance1.BackColor2 = System.Drawing.Color.White;
			appearance1.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.cmbVendedor.Appearance = appearance1;
			this.cmbVendedor.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbVendedor.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbVendedor.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsSolcitud, "Cre_Solicitud.Vendedor"));
			this.cmbVendedor.DataMember = "Personal";
			this.cmbVendedor.DataSource = this.cdsSolcitud;
			ultraGridBand1.AddButtonCaption = "RolSucursal";
			ultraGridColumn1.Header.VisiblePosition = 0;
			ultraGridColumn1.Hidden = true;
			ultraGridColumn2.Header.VisiblePosition = 1;
			ultraGridColumn2.Hidden = true;
			ultraGridColumn3.Header.VisiblePosition = 4;
			ultraGridColumn3.Hidden = true;
			ultraGridColumn4.Header.VisiblePosition = 5;
			ultraGridColumn4.Hidden = true;
			ultraGridColumn5.Header.VisiblePosition = 6;
			ultraGridColumn5.Hidden = true;
			ultraGridColumn6.Header.VisiblePosition = 7;
			ultraGridColumn6.Hidden = true;
			ultraGridColumn7.Header.VisiblePosition = 8;
			ultraGridColumn7.Hidden = true;
			ultraGridColumn8.Header.VisiblePosition = 9;
			ultraGridColumn8.Hidden = true;
			ultraGridColumn9.Header.VisiblePosition = 2;
			ultraGridColumn9.Width = 300;
			ultraGridColumn10.Header.VisiblePosition = 10;
			ultraGridColumn10.Hidden = true;
			ultraGridColumn11.Header.VisiblePosition = 11;
			ultraGridColumn11.Hidden = true;
			ultraGridColumn12.Header.VisiblePosition = 12;
			ultraGridColumn12.Hidden = true;
			ultraGridColumn13.Header.VisiblePosition = 13;
			ultraGridColumn13.Hidden = true;
			ultraGridColumn14.Header.VisiblePosition = 14;
			ultraGridColumn14.Hidden = true;
			ultraGridColumn15.Header.VisiblePosition = 15;
			ultraGridColumn15.Hidden = true;
			ultraGridColumn16.Header.VisiblePosition = 16;
			ultraGridColumn16.Hidden = true;
			ultraGridColumn17.Header.VisiblePosition = 17;
			ultraGridColumn17.Hidden = true;
			ultraGridColumn18.Header.VisiblePosition = 18;
			ultraGridColumn18.Hidden = true;
			ultraGridColumn19.Header.VisiblePosition = 19;
			ultraGridColumn19.Hidden = true;
			ultraGridColumn20.Header.VisiblePosition = 20;
			ultraGridColumn20.Hidden = true;
			ultraGridColumn21.Header.VisiblePosition = 21;
			ultraGridColumn21.Hidden = true;
			ultraGridColumn22.Header.VisiblePosition = 22;
			ultraGridColumn22.Hidden = true;
			ultraGridColumn23.Header.VisiblePosition = 23;
			ultraGridColumn23.Hidden = true;
			ultraGridColumn24.Header.VisiblePosition = 24;
			ultraGridColumn24.Hidden = true;
			ultraGridColumn25.Header.VisiblePosition = 25;
			ultraGridColumn25.Hidden = true;
			ultraGridColumn26.Header.VisiblePosition = 26;
			ultraGridColumn26.Hidden = true;
			ultraGridColumn27.Header.VisiblePosition = 27;
			ultraGridColumn27.Hidden = true;
			ultraGridColumn28.Header.VisiblePosition = 28;
			ultraGridColumn28.Hidden = true;
			ultraGridColumn29.Header.VisiblePosition = 29;
			ultraGridColumn29.Hidden = true;
			ultraGridColumn30.Header.VisiblePosition = 30;
			ultraGridColumn30.Hidden = true;
			ultraGridColumn31.Header.VisiblePosition = 31;
			ultraGridColumn31.Hidden = true;
			ultraGridColumn32.Header.VisiblePosition = 32;
			ultraGridColumn32.Hidden = true;
			ultraGridColumn33.Header.VisiblePosition = 33;
			ultraGridColumn33.Hidden = true;
			ultraGridColumn34.Header.VisiblePosition = 34;
			ultraGridColumn34.Hidden = true;
			ultraGridColumn35.Header.VisiblePosition = 35;
			ultraGridColumn35.Hidden = true;
			ultraGridColumn36.Header.VisiblePosition = 36;
			ultraGridColumn36.Hidden = true;
			ultraGridColumn37.Header.VisiblePosition = 37;
			ultraGridColumn37.Hidden = true;
			ultraGridColumn38.Header.VisiblePosition = 38;
			ultraGridColumn38.Hidden = true;
			ultraGridColumn39.Header.VisiblePosition = 39;
			ultraGridColumn39.Hidden = true;
			ultraGridColumn40.Header.VisiblePosition = 40;
			ultraGridColumn40.Hidden = true;
			ultraGridColumn41.Header.VisiblePosition = 41;
			ultraGridColumn41.Hidden = true;
			ultraGridColumn42.Header.VisiblePosition = 42;
			ultraGridColumn42.Hidden = true;
			ultraGridColumn43.Header.VisiblePosition = 43;
			ultraGridColumn43.Hidden = true;
			ultraGridColumn44.Header.VisiblePosition = 44;
			ultraGridColumn44.Hidden = true;
			ultraGridColumn45.Header.VisiblePosition = 45;
			ultraGridColumn45.Hidden = true;
			ultraGridColumn46.Header.VisiblePosition = 46;
			ultraGridColumn46.Hidden = true;
			ultraGridColumn47.Header.VisiblePosition = 47;
			ultraGridColumn47.Hidden = true;
			ultraGridColumn48.Header.VisiblePosition = 48;
			ultraGridColumn48.Hidden = true;
			ultraGridColumn49.Header.VisiblePosition = 49;
			ultraGridColumn49.Hidden = true;
			ultraGridColumn50.Header.VisiblePosition = 50;
			ultraGridColumn50.Hidden = true;
			ultraGridColumn51.Header.VisiblePosition = 51;
			ultraGridColumn51.Hidden = true;
			ultraGridColumn52.Header.VisiblePosition = 52;
			ultraGridColumn52.Hidden = true;
			ultraGridColumn53.Header.VisiblePosition = 53;
			ultraGridColumn53.Hidden = true;
			ultraGridColumn54.Header.VisiblePosition = 54;
			ultraGridColumn54.Hidden = true;
			ultraGridColumn55.Header.VisiblePosition = 55;
			ultraGridColumn55.Hidden = true;
			ultraGridColumn56.Header.VisiblePosition = 56;
			ultraGridColumn56.Hidden = true;
			ultraGridColumn57.Header.VisiblePosition = 57;
			ultraGridColumn57.Hidden = true;
			ultraGridColumn58.Header.VisiblePosition = 58;
			ultraGridColumn58.Hidden = true;
			ultraGridColumn59.Header.VisiblePosition = 59;
			ultraGridColumn59.Hidden = true;
			ultraGridColumn60.Header.VisiblePosition = 60;
			ultraGridColumn60.Hidden = true;
			ultraGridColumn61.Header.VisiblePosition = 61;
			ultraGridColumn61.Hidden = true;
			ultraGridColumn62.Header.VisiblePosition = 62;
			ultraGridColumn62.Hidden = true;
			ultraGridColumn63.Header.VisiblePosition = 63;
			ultraGridColumn63.Hidden = true;
			ultraGridColumn64.Header.VisiblePosition = 64;
			ultraGridColumn64.Hidden = true;
			ultraGridColumn65.Header.VisiblePosition = 65;
			ultraGridColumn65.Hidden = true;
			ultraGridColumn66.Header.VisiblePosition = 66;
			ultraGridColumn66.Hidden = true;
			ultraGridColumn67.Header.VisiblePosition = 67;
			ultraGridColumn67.Hidden = true;
			ultraGridColumn68.Header.VisiblePosition = 68;
			ultraGridColumn68.Hidden = true;
			ultraGridColumn69.Header.VisiblePosition = 69;
			ultraGridColumn69.Hidden = true;
			ultraGridColumn70.Header.VisiblePosition = 70;
			ultraGridColumn70.Hidden = true;
			ultraGridColumn71.Header.VisiblePosition = 71;
			ultraGridColumn71.Hidden = true;
			ultraGridColumn72.Header.VisiblePosition = 72;
			ultraGridColumn72.Hidden = true;
			ultraGridColumn73.Header.VisiblePosition = 73;
			ultraGridColumn73.Hidden = true;
			ultraGridColumn74.Header.VisiblePosition = 74;
			ultraGridColumn74.Hidden = true;
			ultraGridColumn75.Header.VisiblePosition = 75;
			ultraGridColumn75.Hidden = true;
			ultraGridColumn76.Header.VisiblePosition = 76;
			ultraGridColumn76.Hidden = true;
			ultraGridColumn77.Header.VisiblePosition = 77;
			ultraGridColumn77.Hidden = true;
			ultraGridColumn78.Header.VisiblePosition = 78;
			ultraGridColumn78.Hidden = true;
			ultraGridColumn79.Header.VisiblePosition = 79;
			ultraGridColumn79.Hidden = true;
			ultraGridColumn80.Header.VisiblePosition = 80;
			ultraGridColumn80.Hidden = true;
			ultraGridColumn81.Header.VisiblePosition = 81;
			ultraGridColumn81.Hidden = true;
			ultraGridColumn82.Header.VisiblePosition = 82;
			ultraGridColumn82.Hidden = true;
			ultraGridColumn83.Header.VisiblePosition = 83;
			ultraGridColumn83.Hidden = true;
			ultraGridColumn84.Header.VisiblePosition = 84;
			ultraGridColumn84.Hidden = true;
			ultraGridColumn85.Header.VisiblePosition = 85;
			ultraGridColumn85.Hidden = true;
			ultraGridColumn86.Header.VisiblePosition = 86;
			ultraGridColumn86.Hidden = true;
			ultraGridColumn87.Header.VisiblePosition = 87;
			ultraGridColumn87.Hidden = true;
			ultraGridColumn88.Header.VisiblePosition = 88;
			ultraGridColumn88.Hidden = true;
			ultraGridColumn89.Header.VisiblePosition = 89;
			ultraGridColumn89.Hidden = true;
			ultraGridColumn90.Header.VisiblePosition = 90;
			ultraGridColumn90.Hidden = true;
			ultraGridColumn91.Header.VisiblePosition = 91;
			ultraGridColumn91.Hidden = true;
			ultraGridColumn92.Header.VisiblePosition = 92;
			ultraGridColumn92.Hidden = true;
			ultraGridColumn93.Header.VisiblePosition = 93;
			ultraGridColumn93.Hidden = true;
			ultraGridColumn94.Header.VisiblePosition = 94;
			ultraGridColumn94.Hidden = true;
			ultraGridColumn95.Header.VisiblePosition = 95;
			ultraGridColumn95.Hidden = true;
			ultraGridColumn96.Header.VisiblePosition = 96;
			ultraGridColumn96.Hidden = true;
			ultraGridColumn97.Header.VisiblePosition = 97;
			ultraGridColumn97.Hidden = true;
			ultraGridColumn98.Header.VisiblePosition = 98;
			ultraGridColumn98.Hidden = true;
			ultraGridColumn99.Header.VisiblePosition = 99;
			ultraGridColumn99.Hidden = true;
			ultraGridColumn100.Header.VisiblePosition = 100;
			ultraGridColumn100.Hidden = true;
			ultraGridColumn101.Header.VisiblePosition = 101;
			ultraGridColumn101.Hidden = true;
			ultraGridColumn102.Header.VisiblePosition = 102;
			ultraGridColumn102.Hidden = true;
			ultraGridColumn103.Header.VisiblePosition = 103;
			ultraGridColumn103.Hidden = true;
			ultraGridColumn104.Header.VisiblePosition = 104;
			ultraGridColumn104.Hidden = true;
			ultraGridColumn105.Header.VisiblePosition = 105;
			ultraGridColumn106.Header.VisiblePosition = 3;
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
																										 ultraGridColumn106});
			this.cmbVendedor.DisplayLayout.BandsSerializer.Add(ultraGridBand1);
			this.cmbVendedor.DisplayMember = "Nombre";
			this.cmbVendedor.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbVendedor.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbVendedor.Enabled = false;
			this.cmbVendedor.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbVendedor.Location = new System.Drawing.Point(432, 40);
			this.cmbVendedor.Name = "cmbVendedor";
			this.cmbVendedor.Size = new System.Drawing.Size(112, 21);
			this.cmbVendedor.TabIndex = 21;
			this.cmbVendedor.ValueMember = "idPersonal";
			// 
			// cmbLocal
			// 
			appearance2.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(224)), ((System.Byte)(224)), ((System.Byte)(224)));
			appearance2.BackColor2 = System.Drawing.Color.White;
			appearance2.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.cmbLocal.Appearance = appearance2;
			this.cmbLocal.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbLocal.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbLocal.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsSolcitud, "Cre_Solicitud.Local"));
			this.cmbLocal.DataMember = "Bodega";
			this.cmbLocal.DataSource = this.cdsSolcitud;
			ultraGridBand2.AddButtonCaption = "RolSucursal";
			ultraGridColumn107.Header.VisiblePosition = 0;
			ultraGridColumn107.Hidden = true;
			ultraGridColumn108.Header.VisiblePosition = 1;
			ultraGridColumn108.Width = 150;
			ultraGridColumn109.Header.VisiblePosition = 3;
			ultraGridColumn109.Hidden = true;
			ultraGridColumn110.Header.VisiblePosition = 4;
			ultraGridColumn110.Hidden = true;
			ultraGridColumn111.Header.VisiblePosition = 8;
			ultraGridColumn112.Header.VisiblePosition = 5;
			ultraGridColumn112.Hidden = true;
			ultraGridColumn113.Header.VisiblePosition = 2;
			ultraGridColumn113.Width = 50;
			ultraGridColumn114.Header.VisiblePosition = 6;
			ultraGridColumn115.Header.VisiblePosition = 7;
			ultraGridColumn116.Header.VisiblePosition = 9;
			ultraGridColumn117.Header.VisiblePosition = 10;
			ultraGridColumn118.Header.VisiblePosition = 11;
			ultraGridColumn119.Header.VisiblePosition = 12;
			ultraGridColumn120.Header.VisiblePosition = 13;
			ultraGridColumn121.Header.VisiblePosition = 14;
			ultraGridColumn122.Header.VisiblePosition = 15;
			ultraGridColumn123.Header.VisiblePosition = 16;
			ultraGridColumn124.Header.VisiblePosition = 17;
			ultraGridBand2.Columns.AddRange(new object[] {
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
																										 ultraGridColumn122,
																										 ultraGridColumn123,
																										 ultraGridColumn124});
			this.cmbLocal.DisplayLayout.BandsSerializer.Add(ultraGridBand2);
			this.cmbLocal.DisplayMember = "Nombre";
			this.cmbLocal.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbLocal.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbLocal.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbLocal.Location = new System.Drawing.Point(128, 40);
			this.cmbLocal.Name = "cmbLocal";
			this.cmbLocal.Size = new System.Drawing.Size(96, 21);
			this.cmbLocal.TabIndex = 19;
			this.cmbLocal.ValueMember = "idBodega";
			// 
			// cmbTipTarjeta
			// 
			appearance3.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(224)), ((System.Byte)(224)), ((System.Byte)(224)));
			appearance3.BackColor2 = System.Drawing.Color.White;
			appearance3.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.cmbTipTarjeta.Appearance = appearance3;
			this.cmbTipTarjeta.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbTipTarjeta.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbTipTarjeta.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsSolcitud, "Cre_Solicitud.TipoTarjeta"));
			this.cmbTipTarjeta.DataMember = "Cre_TipoTarjeta";
			this.cmbTipTarjeta.DataSource = this.cdsSolcitud;
			ultraGridBand3.AddButtonCaption = "RolSucursal";
			ultraGridColumn125.Header.VisiblePosition = 0;
			ultraGridColumn125.Hidden = true;
			ultraGridColumn126.Header.VisiblePosition = 2;
			ultraGridColumn126.Width = 50;
			ultraGridColumn127.Header.VisiblePosition = 1;
			ultraGridColumn127.Width = 150;
			ultraGridBand3.Columns.AddRange(new object[] {
																										 ultraGridColumn125,
																										 ultraGridColumn126,
																										 ultraGridColumn127});
			this.cmbTipTarjeta.DisplayLayout.BandsSerializer.Add(ultraGridBand3);
			this.cmbTipTarjeta.DisplayMember = "Nombre";
			this.cmbTipTarjeta.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbTipTarjeta.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbTipTarjeta.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbTipTarjeta.Location = new System.Drawing.Point(8, 40);
			this.cmbTipTarjeta.Name = "cmbTipTarjeta";
			this.cmbTipTarjeta.Size = new System.Drawing.Size(112, 21);
			this.cmbTipTarjeta.TabIndex = 18;
			this.cmbTipTarjeta.ValueMember = "idTipoTarjeta";
			this.cmbTipTarjeta.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.cmbTipTarjeta_KeyPress);
			this.cmbTipTarjeta.Validating += new System.ComponentModel.CancelEventHandler(this.cmbTipTarjeta_Validating);
			this.cmbTipTarjeta.Validated += new System.EventHandler(this.cmbTipTarjeta_Validated);
			this.cmbTipTarjeta.InitializeLayout += new Infragistics.Win.UltraWinGrid.InitializeLayoutEventHandler(this.cmbTipTarjeta_InitializeLayout);
			// 
			// cmbTipoEstado
			// 
			appearance4.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(224)), ((System.Byte)(224)), ((System.Byte)(224)));
			appearance4.BackColor2 = System.Drawing.Color.White;
			appearance4.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.cmbTipoEstado.Appearance = appearance4;
			this.cmbTipoEstado.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbTipoEstado.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbTipoEstado.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsSolcitud, "Cre_Solicitud.TipoCalificacion"));
			this.cmbTipoEstado.DisplayMember = "Nombre";
			this.cmbTipoEstado.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbTipoEstado.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbTipoEstado.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbTipoEstado.Location = new System.Drawing.Point(128, 80);
			this.cmbTipoEstado.Name = "cmbTipoEstado";
			this.cmbTipoEstado.Size = new System.Drawing.Size(96, 21);
			this.cmbTipoEstado.TabIndex = 25;
			this.cmbTipoEstado.ValueMember = "idNegadoPendiente";
			// 
			// cmbEstado
			// 
			appearance5.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(224)), ((System.Byte)(224)), ((System.Byte)(224)));
			appearance5.BackColor2 = System.Drawing.Color.White;
			appearance5.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.cmbEstado.Appearance = appearance5;
			this.cmbEstado.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbEstado.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbEstado.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsSolcitud, "Cre_Solicitud.Estado"));
			this.cmbEstado.DataMember = "Cre_EtdoSolicitud";
			this.cmbEstado.DataSource = this.cdsSolcitud;
			ultraGridBand4.AddButtonCaption = "RolSucursal";
			ultraGridColumn128.Header.VisiblePosition = 0;
			ultraGridColumn128.Hidden = true;
			ultraGridColumn129.Header.VisiblePosition = 1;
			ultraGridColumn129.Width = 50;
			ultraGridColumn130.Header.VisiblePosition = 2;
			ultraGridColumn130.Width = 150;
			ultraGridBand4.Columns.AddRange(new object[] {
																										 ultraGridColumn128,
																										 ultraGridColumn129,
																										 ultraGridColumn130});
			this.cmbEstado.DisplayLayout.BandsSerializer.Add(ultraGridBand4);
			this.cmbEstado.DisplayMember = "Nombre";
			this.cmbEstado.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbEstado.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbEstado.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbEstado.Location = new System.Drawing.Point(8, 80);
			this.cmbEstado.Name = "cmbEstado";
			this.cmbEstado.Size = new System.Drawing.Size(112, 21);
			this.cmbEstado.TabIndex = 24;
			this.cmbEstado.ValueMember = "idEstado";
			this.cmbEstado.Validating += new System.ComponentModel.CancelEventHandler(this.cmbEstado_Validating);
			this.cmbEstado.Validated += new System.EventHandler(this.cmbEstado_Validated);
			this.cmbEstado.InitializeLayout += new Infragistics.Win.UltraWinGrid.InitializeLayoutEventHandler(this.cmbEstado_InitializeLayout);
			// 
			// cmbCalificacion
			// 
			appearance6.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(224)), ((System.Byte)(224)), ((System.Byte)(224)));
			appearance6.BackColor2 = System.Drawing.Color.White;
			appearance6.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.cmbCalificacion.Appearance = appearance6;
			this.cmbCalificacion.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbCalificacion.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbCalificacion.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsSolcitud, "Cre_Solicitud.Calificacion"));
			this.cmbCalificacion.DataMember = "Cre_TipoCalificacion";
			this.cmbCalificacion.DataSource = this.cdsSolcitud;
			ultraGridBand5.AddButtonCaption = "RolSucursal";
			ultraGridColumn131.Header.VisiblePosition = 0;
			ultraGridColumn131.Hidden = true;
			ultraGridColumn132.Header.VisiblePosition = 1;
			ultraGridColumn132.Width = 50;
			ultraGridColumn133.Header.VisiblePosition = 2;
			ultraGridColumn133.Width = 150;
			ultraGridBand5.Columns.AddRange(new object[] {
																										 ultraGridColumn131,
																										 ultraGridColumn132,
																										 ultraGridColumn133});
			this.cmbCalificacion.DisplayLayout.BandsSerializer.Add(ultraGridBand5);
			this.cmbCalificacion.DisplayMember = "Codigo";
			this.cmbCalificacion.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbCalificacion.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbCalificacion.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbCalificacion.Location = new System.Drawing.Point(648, 40);
			this.cmbCalificacion.Name = "cmbCalificacion";
			this.cmbCalificacion.Size = new System.Drawing.Size(56, 21);
			this.cmbCalificacion.TabIndex = 23;
			this.cmbCalificacion.ValueMember = "idTipoCalificacion";
			// 
			// cmbCallCenter
			// 
			appearance7.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(224)), ((System.Byte)(224)), ((System.Byte)(224)));
			appearance7.BackColor2 = System.Drawing.Color.White;
			appearance7.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.cmbCallCenter.Appearance = appearance7;
			this.cmbCallCenter.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbCallCenter.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbCallCenter.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsSolcitud, "Cre_Solicitud.CallCenter"));
			this.cmbCallCenter.DataMember = "Cre_CallCenter";
			this.cmbCallCenter.DataSource = this.cdsSolcitud;
			ultraGridBand6.AddButtonCaption = "RolSucursal";
			ultraGridColumn134.Header.VisiblePosition = 0;
			ultraGridColumn134.Hidden = true;
			ultraGridColumn135.Header.VisiblePosition = 2;
			ultraGridColumn135.Width = 50;
			ultraGridColumn136.Header.VisiblePosition = 1;
			ultraGridColumn136.Width = 150;
			ultraGridBand6.Columns.AddRange(new object[] {
																										 ultraGridColumn134,
																										 ultraGridColumn135,
																										 ultraGridColumn136});
			this.cmbCallCenter.DisplayLayout.BandsSerializer.Add(ultraGridBand6);
			this.cmbCallCenter.DisplayMember = "Nombre";
			this.cmbCallCenter.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbCallCenter.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbCallCenter.Enabled = false;
			this.cmbCallCenter.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbCallCenter.Location = new System.Drawing.Point(552, 40);
			this.cmbCallCenter.Name = "cmbCallCenter";
			this.cmbCallCenter.Size = new System.Drawing.Size(88, 21);
			this.cmbCallCenter.TabIndex = 22;
			this.cmbCallCenter.ValueMember = "idCallCenter";
			// 
			// cmbInstitucion
			// 
			appearance8.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(224)), ((System.Byte)(224)), ((System.Byte)(224)));
			appearance8.BackColor2 = System.Drawing.Color.White;
			appearance8.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.cmbInstitucion.Appearance = appearance8;
			this.cmbInstitucion.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbInstitucion.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbInstitucion.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsSolcitud, "Cre_Solicitud.Institucion"));
			this.cmbInstitucion.DataMember = "Cre_Instituciones";
			this.cmbInstitucion.DataSource = this.cdsSolcitud;
			ultraGridBand7.AddButtonCaption = "RolSucursal";
			ultraGridColumn137.Header.VisiblePosition = 0;
			ultraGridColumn137.Hidden = true;
			ultraGridColumn138.Header.VisiblePosition = 2;
			ultraGridColumn138.Width = 50;
			ultraGridColumn139.Header.VisiblePosition = 1;
			ultraGridColumn139.Width = 300;
			ultraGridColumn140.Header.VisiblePosition = 3;
			ultraGridColumn140.Hidden = true;
			ultraGridColumn141.Header.VisiblePosition = 4;
			ultraGridColumn141.Hidden = true;
			ultraGridColumn142.Header.VisiblePosition = 5;
			ultraGridColumn142.Hidden = true;
			ultraGridColumn143.Header.VisiblePosition = 6;
			ultraGridColumn143.Hidden = true;
			ultraGridBand7.Columns.AddRange(new object[] {
																										 ultraGridColumn137,
																										 ultraGridColumn138,
																										 ultraGridColumn139,
																										 ultraGridColumn140,
																										 ultraGridColumn141,
																										 ultraGridColumn142,
																										 ultraGridColumn143});
			this.cmbInstitucion.DisplayLayout.BandsSerializer.Add(ultraGridBand7);
			this.cmbInstitucion.DisplayMember = "Descripcion";
			this.cmbInstitucion.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbInstitucion.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbInstitucion.Enabled = false;
			this.cmbInstitucion.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbInstitucion.Location = new System.Drawing.Point(232, 40);
			this.cmbInstitucion.Name = "cmbInstitucion";
			this.cmbInstitucion.Size = new System.Drawing.Size(192, 21);
			this.cmbInstitucion.TabIndex = 20;
			this.cmbInstitucion.ValueMember = "idInstitucion";
			// 
			// btGarante
			// 
			this.btGarante.Enabled = false;
			this.btGarante.Font = new System.Drawing.Font("Microsoft Sans Serif", 18F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.btGarante.Location = new System.Drawing.Point(680, 80);
			this.btGarante.Name = "btGarante";
			this.btGarante.Size = new System.Drawing.Size(24, 25);
			this.btGarante.TabIndex = 28;
			this.btGarante.Text = "+";
			// 
			// label28
			// 
			this.label28.AutoSize = true;
			this.label28.BackColor = System.Drawing.Color.Transparent;
			this.label28.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label28.Location = new System.Drawing.Point(521, 64);
			this.label28.Name = "label28";
			this.label28.Size = new System.Drawing.Size(45, 16);
			this.label28.TabIndex = 64;
			this.label28.Text = "Garante";
			// 
			// ultraTextEditor3
			// 
			appearance9.BackColor = System.Drawing.Color.Gainsboro;
			appearance9.BackColor2 = System.Drawing.Color.White;
			appearance9.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ultraTextEditor3.Appearance = appearance9;
			this.ultraTextEditor3.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.ultraTextEditor3.DataBindings.Add(new System.Windows.Forms.Binding("Text", this.cdsSolcitud, "Cre_Solicitud.Observacion"));
			this.ultraTextEditor3.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.ultraTextEditor3.Location = new System.Drawing.Point(8, 120);
			this.ultraTextEditor3.Name = "ultraTextEditor3";
			this.ultraTextEditor3.Size = new System.Drawing.Size(696, 21);
			this.ultraTextEditor3.TabIndex = 29;
			// 
			// label27
			// 
			this.label27.AutoSize = true;
			this.label27.BackColor = System.Drawing.Color.Transparent;
			this.label27.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label27.Location = new System.Drawing.Point(8, 104);
			this.label27.Name = "label27";
			this.label27.Size = new System.Drawing.Size(80, 16);
			this.label27.TabIndex = 62;
			this.label27.Text = "Observaciones";
			// 
			// txtNumTarjeta
			// 
			appearance10.BackColor = System.Drawing.Color.Gainsboro;
			appearance10.BackColor2 = System.Drawing.Color.White;
			appearance10.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.txtNumTarjeta.Appearance = appearance10;
			this.txtNumTarjeta.DataBindings.Add(new System.Windows.Forms.Binding("Text", this.cdsSolcitud, "Cre_Solicitud.NumeroTarjeta"));
			this.txtNumTarjeta.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtNumTarjeta.Enabled = false;
			this.txtNumTarjeta.Location = new System.Drawing.Point(320, 80);
			this.txtNumTarjeta.Name = "txtNumTarjeta";
			this.txtNumTarjeta.Size = new System.Drawing.Size(184, 21);
			this.txtNumTarjeta.TabIndex = 100;
			// 
			// label26
			// 
			this.label26.AutoSize = true;
			this.label26.BackColor = System.Drawing.Color.Transparent;
			this.label26.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label26.Location = new System.Drawing.Point(320, 64);
			this.label26.Name = "label26";
			this.label26.Size = new System.Drawing.Size(49, 16);
			this.label26.TabIndex = 60;
			this.label26.Text = "# Tarjeta";
			// 
			// label24
			// 
			this.label24.AutoSize = true;
			this.label24.BackColor = System.Drawing.Color.Transparent;
			this.label24.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label24.Location = new System.Drawing.Point(128, 64);
			this.label24.Name = "label24";
			this.label24.Size = new System.Drawing.Size(65, 16);
			this.label24.TabIndex = 53;
			this.label24.Text = "Tipo Estado";
			// 
			// label23
			// 
			this.label23.AutoSize = true;
			this.label23.BackColor = System.Drawing.Color.Transparent;
			this.label23.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label23.Location = new System.Drawing.Point(8, 64);
			this.label23.Name = "label23";
			this.label23.Size = new System.Drawing.Size(39, 16);
			this.label23.TabIndex = 51;
			this.label23.Text = "Estado";
			// 
			// label22
			// 
			this.label22.AutoSize = true;
			this.label22.BackColor = System.Drawing.Color.Transparent;
			this.label22.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label22.Location = new System.Drawing.Point(648, 24);
			this.label22.Name = "label22";
			this.label22.Size = new System.Drawing.Size(62, 16);
			this.label22.TabIndex = 49;
			this.label22.Text = "Calificación";
			// 
			// label21
			// 
			this.label21.AutoSize = true;
			this.label21.BackColor = System.Drawing.Color.Transparent;
			this.label21.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label21.Location = new System.Drawing.Point(552, 24);
			this.label21.Name = "label21";
			this.label21.Size = new System.Drawing.Size(73, 16);
			this.label21.TabIndex = 47;
			this.label21.Text = "Call CENTER";
			// 
			// label20
			// 
			this.label20.AutoSize = true;
			this.label20.BackColor = System.Drawing.Color.Transparent;
			this.label20.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label20.Location = new System.Drawing.Point(432, 24);
			this.label20.Name = "label20";
			this.label20.Size = new System.Drawing.Size(53, 16);
			this.label20.TabIndex = 45;
			this.label20.Text = "Vendedor";
			// 
			// label19
			// 
			this.label19.AutoSize = true;
			this.label19.BackColor = System.Drawing.Color.Transparent;
			this.label19.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label19.Location = new System.Drawing.Point(232, 24);
			this.label19.Name = "label19";
			this.label19.Size = new System.Drawing.Size(55, 16);
			this.label19.TabIndex = 43;
			this.label19.Text = "Institución";
			// 
			// label18
			// 
			this.label18.AutoSize = true;
			this.label18.BackColor = System.Drawing.Color.Transparent;
			this.label18.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label18.Location = new System.Drawing.Point(128, 24);
			this.label18.Name = "label18";
			this.label18.Size = new System.Drawing.Size(31, 16);
			this.label18.TabIndex = 41;
			this.label18.Text = "Local";
			// 
			// label17
			// 
			this.label17.AutoSize = true;
			this.label17.BackColor = System.Drawing.Color.Transparent;
			this.label17.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label17.Location = new System.Drawing.Point(8, 24);
			this.label17.Name = "label17";
			this.label17.Size = new System.Drawing.Size(65, 16);
			this.label17.TabIndex = 39;
			this.label17.Text = "Tipo Tarjeta";
			// 
			// label25
			// 
			this.label25.AutoSize = true;
			this.label25.BackColor = System.Drawing.Color.Transparent;
			this.label25.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label25.Location = new System.Drawing.Point(232, 64);
			this.label25.Name = "label25";
			this.label25.Size = new System.Drawing.Size(34, 16);
			this.label25.TabIndex = 58;
			this.label25.Text = "Cuota";
			// 
			// txtCuota
			// 
			appearance11.BackColor = System.Drawing.Color.Gainsboro;
			appearance11.BackColor2 = System.Drawing.Color.White;
			appearance11.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance11.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtCuota.Appearance = appearance11;
			this.txtCuota.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsSolcitud, "Cre_Solicitud.CuotaAsignada"));
			this.txtCuota.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtCuota.Enabled = false;
			this.txtCuota.FormatString = "#,##0.00";
			this.txtCuota.Location = new System.Drawing.Point(232, 80);
			this.txtCuota.Name = "txtCuota";
			this.txtCuota.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtCuota.PromptChar = ' ';
			this.txtCuota.Size = new System.Drawing.Size(80, 21);
			this.txtCuota.TabIndex = 26;
			this.txtCuota.ValueChanged += new System.EventHandler(this.txtCuota_ValueChanged);
			// 
			// groupBox1
			// 
			this.groupBox1.Controls.Add(this.cmbParroquia);
			this.groupBox1.Controls.Add(this.txtRefUbicacion);
			this.groupBox1.Controls.Add(this.label16);
			this.groupBox1.Controls.Add(this.txtNumTelefono);
			this.groupBox1.Controls.Add(this.label15);
			this.groupBox1.Controls.Add(this.txtCalleTransversal);
			this.groupBox1.Controls.Add(this.label14);
			this.groupBox1.Controls.Add(this.txtNumCasa);
			this.groupBox1.Controls.Add(this.label13);
			this.groupBox1.Controls.Add(this.cmbBarrio);
			this.groupBox1.Controls.Add(this.cmbCanton);
			this.groupBox1.Controls.Add(this.cmbProvincia);
			this.groupBox1.Controls.Add(this.label11);
			this.groupBox1.Controls.Add(this.label10);
			this.groupBox1.Controls.Add(this.label9);
			this.groupBox1.Controls.Add(this.label8);
			this.groupBox1.Controls.Add(this.label12);
			this.groupBox1.Controls.Add(this.txtAvPrincipal);
			this.groupBox1.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.groupBox1.Location = new System.Drawing.Point(8, 96);
			this.groupBox1.Name = "groupBox1";
			this.groupBox1.Size = new System.Drawing.Size(712, 152);
			this.groupBox1.TabIndex = 54;
			this.groupBox1.TabStop = false;
			this.groupBox1.Text = "Dirección Cliente";
			// 
			// cmbParroquia
			// 
			appearance12.BackColor = System.Drawing.Color.GreenYellow;
			appearance12.BackColor2 = System.Drawing.Color.White;
			appearance12.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.cmbParroquia.Appearance = appearance12;
			this.cmbParroquia.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbParroquia.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbParroquia.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsSolcitud, "Cre_Solicitud.ParroquiaDomicilio"));
			this.cmbParroquia.DisplayMember = "Nombre";
			this.cmbParroquia.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbParroquia.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbParroquia.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbParroquia.Location = new System.Drawing.Point(360, 32);
			this.cmbParroquia.Name = "cmbParroquia";
			this.cmbParroquia.Size = new System.Drawing.Size(168, 21);
			this.cmbParroquia.TabIndex = 11;
			this.cmbParroquia.ValueMember = "idParroquia";
			this.cmbParroquia.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.cmbParroquia_KeyPress);
			this.cmbParroquia.InitializeLayout += new Infragistics.Win.UltraWinGrid.InitializeLayoutEventHandler(this.cmbParroquia_InitializeLayout);
			// 
			// txtRefUbicacion
			// 
			appearance13.BackColor = System.Drawing.Color.GreenYellow;
			appearance13.BackColor2 = System.Drawing.Color.White;
			appearance13.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.txtRefUbicacion.Appearance = appearance13;
			this.txtRefUbicacion.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtRefUbicacion.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsSolcitud, "Cre_Solicitud.ReferenciaUbicacion"));
			this.txtRefUbicacion.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtRefUbicacion.Location = new System.Drawing.Point(8, 120);
			this.txtRefUbicacion.Name = "txtRefUbicacion";
			this.txtRefUbicacion.Size = new System.Drawing.Size(696, 21);
			this.txtRefUbicacion.TabIndex = 17;
			this.txtRefUbicacion.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.txtRefUbicacion_KeyPress);
			// 
			// label16
			// 
			this.label16.AutoSize = true;
			this.label16.BackColor = System.Drawing.Color.Transparent;
			this.label16.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label16.Location = new System.Drawing.Point(8, 104);
			this.label16.Name = "label16";
			this.label16.Size = new System.Drawing.Size(161, 16);
			this.label16.TabIndex = 56;
			this.label16.Text = "Referencia Ubicación Domicilio";
			// 
			// txtNumTelefono
			// 
			appearance14.BackColor = System.Drawing.Color.GreenYellow;
			appearance14.BackColor2 = System.Drawing.Color.White;
			appearance14.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.txtNumTelefono.Appearance = appearance14;
			this.txtNumTelefono.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtNumTelefono.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsSolcitud, "Cre_Solicitud.TelefonoDomicilio1"));
			this.txtNumTelefono.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtNumTelefono.Location = new System.Drawing.Point(552, 80);
			this.txtNumTelefono.Name = "txtNumTelefono";
			this.txtNumTelefono.Size = new System.Drawing.Size(104, 21);
			this.txtNumTelefono.TabIndex = 16;
			this.txtNumTelefono.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.txtNumTelefono_KeyPress);
			this.txtNumTelefono.ValueChanged += new System.EventHandler(this.txtNumTelefono_ValueChanged);
			// 
			// label15
			// 
			this.label15.AutoSize = true;
			this.label15.BackColor = System.Drawing.Color.Transparent;
			this.label15.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label15.Location = new System.Drawing.Point(552, 64);
			this.label15.Name = "label15";
			this.label15.Size = new System.Drawing.Size(58, 16);
			this.label15.TabIndex = 54;
			this.label15.Text = "# Teléfono";
			// 
			// txtCalleTransversal
			// 
			appearance15.BackColor = System.Drawing.Color.GreenYellow;
			appearance15.BackColor2 = System.Drawing.Color.White;
			appearance15.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.txtCalleTransversal.Appearance = appearance15;
			this.txtCalleTransversal.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtCalleTransversal.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsSolcitud, "Cre_Solicitud.CalleTransversal"));
			this.txtCalleTransversal.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtCalleTransversal.Location = new System.Drawing.Point(320, 80);
			this.txtCalleTransversal.Name = "txtCalleTransversal";
			this.txtCalleTransversal.Size = new System.Drawing.Size(224, 21);
			this.txtCalleTransversal.TabIndex = 15;
			this.txtCalleTransversal.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.txtCalleTransversal_KeyPress);
			// 
			// label14
			// 
			this.label14.AutoSize = true;
			this.label14.BackColor = System.Drawing.Color.Transparent;
			this.label14.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label14.Location = new System.Drawing.Point(320, 64);
			this.label14.Name = "label14";
			this.label14.Size = new System.Drawing.Size(92, 16);
			this.label14.TabIndex = 52;
			this.label14.Text = "Calle Transversal";
			// 
			// txtNumCasa
			// 
			appearance16.BackColor = System.Drawing.Color.GreenYellow;
			appearance16.BackColor2 = System.Drawing.Color.White;
			appearance16.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.txtNumCasa.Appearance = appearance16;
			this.txtNumCasa.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtNumCasa.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsSolcitud, "Cre_Solicitud.NumeroCasa"));
			this.txtNumCasa.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtNumCasa.Location = new System.Drawing.Point(232, 80);
			this.txtNumCasa.Name = "txtNumCasa";
			this.txtNumCasa.Size = new System.Drawing.Size(80, 21);
			this.txtNumCasa.TabIndex = 14;
			this.txtNumCasa.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.txtNumCasa_KeyPress);
			// 
			// label13
			// 
			this.label13.AutoSize = true;
			this.label13.BackColor = System.Drawing.Color.Transparent;
			this.label13.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label13.Location = new System.Drawing.Point(232, 64);
			this.label13.Name = "label13";
			this.label13.Size = new System.Drawing.Size(40, 16);
			this.label13.TabIndex = 50;
			this.label13.Text = "# Casa";
			// 
			// cmbBarrio
			// 
			appearance17.BackColor = System.Drawing.Color.GreenYellow;
			appearance17.BackColor2 = System.Drawing.Color.White;
			appearance17.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.cmbBarrio.Appearance = appearance17;
			this.cmbBarrio.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbBarrio.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbBarrio.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsSolcitud, "Cre_Solicitud.BarrioDomicilio"));
			this.cmbBarrio.DataMember = "Cre_Barrio";
			this.cmbBarrio.DataSource = this.cdsSolcitud;
			ultraGridBand8.AddButtonCaption = "RolSucursal";
			ultraGridColumn144.Header.VisiblePosition = 0;
			ultraGridColumn144.Hidden = true;
			ultraGridColumn145.Header.VisiblePosition = 2;
			ultraGridColumn145.Width = 70;
			ultraGridColumn146.AutoSizeMode = Infragistics.Win.UltraWinGrid.ColumnAutoSizeMode.AllRowsInBand;
			ultraGridColumn146.Header.VisiblePosition = 1;
			ultraGridColumn146.Width = 300;
			ultraGridBand8.Columns.AddRange(new object[] {
																										 ultraGridColumn144,
																										 ultraGridColumn145,
																										 ultraGridColumn146});
			this.cmbBarrio.DisplayLayout.BandsSerializer.Add(ultraGridBand8);
			this.cmbBarrio.DisplayMember = "Nombre";
			this.cmbBarrio.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbBarrio.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbBarrio.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbBarrio.Location = new System.Drawing.Point(536, 32);
			this.cmbBarrio.Name = "cmbBarrio";
			this.cmbBarrio.Size = new System.Drawing.Size(168, 21);
			this.cmbBarrio.TabIndex = 12;
			this.cmbBarrio.ValueMember = "idBarrio";
			this.cmbBarrio.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.cmbBarrio_KeyPress);
			this.cmbBarrio.InitializeLayout += new Infragistics.Win.UltraWinGrid.InitializeLayoutEventHandler(this.cmbBarrio_InitializeLayout);
			// 
			// cmbCanton
			// 
			appearance18.BackColor = System.Drawing.Color.GreenYellow;
			appearance18.BackColor2 = System.Drawing.Color.White;
			appearance18.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.cmbCanton.Appearance = appearance18;
			this.cmbCanton.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbCanton.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbCanton.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsSolcitud, "Cre_Solicitud.CantonDomicilio"));
			this.cmbCanton.DisplayMember = "Nombre";
			this.cmbCanton.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbCanton.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbCanton.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbCanton.Location = new System.Drawing.Point(184, 32);
			this.cmbCanton.Name = "cmbCanton";
			this.cmbCanton.Size = new System.Drawing.Size(168, 21);
			this.cmbCanton.TabIndex = 10;
			this.cmbCanton.ValueMember = "idCanton";
			this.cmbCanton.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.cmbCanton_KeyPress);
			this.cmbCanton.Validating += new System.ComponentModel.CancelEventHandler(this.cmbCanton_Validating);
			this.cmbCanton.InitializeLayout += new Infragistics.Win.UltraWinGrid.InitializeLayoutEventHandler(this.cmbCanton_InitializeLayout);
			// 
			// cmbProvincia
			// 
			appearance19.BackColor = System.Drawing.Color.GreenYellow;
			appearance19.BackColor2 = System.Drawing.Color.White;
			appearance19.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.cmbProvincia.Appearance = appearance19;
			this.cmbProvincia.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbProvincia.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbProvincia.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsSolcitud, "Cre_Solicitud.ProvinciaDomicilio"));
			this.cmbProvincia.DataMember = "Cre_Provincia";
			this.cmbProvincia.DataSource = this.cdsSolcitud;
			ultraGridBand9.AddButtonCaption = "RolSucursal";
			ultraGridColumn147.Header.VisiblePosition = 0;
			ultraGridColumn147.Hidden = true;
			ultraGridColumn148.Header.VisiblePosition = 2;
			ultraGridColumn148.Width = 50;
			ultraGridColumn149.Header.VisiblePosition = 1;
			ultraGridColumn149.Width = 200;
			ultraGridBand9.Columns.AddRange(new object[] {
																										 ultraGridColumn147,
																										 ultraGridColumn148,
																										 ultraGridColumn149});
			this.cmbProvincia.DisplayLayout.BandsSerializer.Add(ultraGridBand9);
			this.cmbProvincia.DisplayMember = "Nombre";
			this.cmbProvincia.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbProvincia.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbProvincia.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbProvincia.Location = new System.Drawing.Point(8, 32);
			this.cmbProvincia.Name = "cmbProvincia";
			this.cmbProvincia.Size = new System.Drawing.Size(168, 21);
			this.cmbProvincia.TabIndex = 9;
			this.cmbProvincia.ValueMember = "idProvincia";
			this.cmbProvincia.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.cmbProvincia_KeyPress);
			this.cmbProvincia.Validating += new System.ComponentModel.CancelEventHandler(this.cmbProvincia_Validating);
			this.cmbProvincia.ValueChanged += new System.EventHandler(this.cmbProvincia_ValueChanged);
			this.cmbProvincia.InitializeLayout += new Infragistics.Win.UltraWinGrid.InitializeLayoutEventHandler(this.cmbProvincia_InitializeLayout);
			// 
			// label11
			// 
			this.label11.AutoSize = true;
			this.label11.BackColor = System.Drawing.Color.Transparent;
			this.label11.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label11.Location = new System.Drawing.Point(536, 16);
			this.label11.Name = "label11";
			this.label11.Size = new System.Drawing.Size(34, 16);
			this.label11.TabIndex = 44;
			this.label11.Text = "Barrio";
			// 
			// label10
			// 
			this.label10.AutoSize = true;
			this.label10.BackColor = System.Drawing.Color.Transparent;
			this.label10.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label10.Location = new System.Drawing.Point(360, 16);
			this.label10.Name = "label10";
			this.label10.Size = new System.Drawing.Size(53, 16);
			this.label10.TabIndex = 42;
			this.label10.Text = "Parroquia";
			// 
			// label9
			// 
			this.label9.AutoSize = true;
			this.label9.BackColor = System.Drawing.Color.Transparent;
			this.label9.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label9.Location = new System.Drawing.Point(184, 16);
			this.label9.Name = "label9";
			this.label9.Size = new System.Drawing.Size(41, 16);
			this.label9.TabIndex = 40;
			this.label9.Text = "Cantón";
			// 
			// label8
			// 
			this.label8.AutoSize = true;
			this.label8.BackColor = System.Drawing.Color.Transparent;
			this.label8.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label8.Location = new System.Drawing.Point(8, 16);
			this.label8.Name = "label8";
			this.label8.Size = new System.Drawing.Size(51, 16);
			this.label8.TabIndex = 38;
			this.label8.Text = "Provincia";
			// 
			// label12
			// 
			this.label12.AutoSize = true;
			this.label12.BackColor = System.Drawing.Color.Transparent;
			this.label12.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label12.Location = new System.Drawing.Point(8, 64);
			this.label12.Name = "label12";
			this.label12.Size = new System.Drawing.Size(67, 16);
			this.label12.TabIndex = 38;
			this.label12.Text = "Av. Principal";
			// 
			// txtAvPrincipal
			// 
			appearance20.BackColor = System.Drawing.Color.GreenYellow;
			appearance20.BackColor2 = System.Drawing.Color.White;
			appearance20.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.txtAvPrincipal.Appearance = appearance20;
			this.txtAvPrincipal.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtAvPrincipal.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsSolcitud, "Cre_Solicitud.AvenidaPrincipal"));
			this.txtAvPrincipal.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtAvPrincipal.Location = new System.Drawing.Point(8, 80);
			this.txtAvPrincipal.Name = "txtAvPrincipal";
			this.txtAvPrincipal.Size = new System.Drawing.Size(216, 21);
			this.txtAvPrincipal.TabIndex = 13;
			this.txtAvPrincipal.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.txtAvPrincipal_KeyPress);
			this.txtAvPrincipal.ValueChanged += new System.EventHandler(this.txtAvPrincipal_ValueChanged);
			// 
			// txtSegundoNombre
			// 
			appearance21.BackColor = System.Drawing.Color.Gainsboro;
			appearance21.BackColor2 = System.Drawing.Color.White;
			appearance21.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.txtSegundoNombre.Appearance = appearance21;
			this.txtSegundoNombre.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtSegundoNombre.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsSolcitud, "Cre_Solicitud.SegundoNombre"));
			this.txtSegundoNombre.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtSegundoNombre.Location = new System.Drawing.Point(488, 64);
			this.txtSegundoNombre.Name = "txtSegundoNombre";
			this.txtSegundoNombre.Size = new System.Drawing.Size(152, 21);
			this.txtSegundoNombre.TabIndex = 8;
			this.txtSegundoNombre.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.txtSegundoNombre_KeyPress);
			this.txtSegundoNombre.ValueChanged += new System.EventHandler(this.txtSegundoNombre_ValueChanged);
			// 
			// label7
			// 
			this.label7.AutoSize = true;
			this.label7.BackColor = System.Drawing.Color.Transparent;
			this.label7.Location = new System.Drawing.Point(488, 48);
			this.label7.Name = "label7";
			this.label7.Size = new System.Drawing.Size(93, 16);
			this.label7.TabIndex = 52;
			this.label7.Text = "Segundo Nombre";
			// 
			// txtPrimerNombre
			// 
			appearance22.BackColor = System.Drawing.Color.GreenYellow;
			appearance22.BackColor2 = System.Drawing.Color.White;
			appearance22.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.txtPrimerNombre.Appearance = appearance22;
			this.txtPrimerNombre.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtPrimerNombre.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsSolcitud, "Cre_Solicitud.PrimerNombre"));
			this.txtPrimerNombre.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtPrimerNombre.Location = new System.Drawing.Point(328, 64);
			this.txtPrimerNombre.Name = "txtPrimerNombre";
			this.txtPrimerNombre.Size = new System.Drawing.Size(152, 21);
			this.txtPrimerNombre.TabIndex = 7;
			this.txtPrimerNombre.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.txtPrimerNombre_KeyPress);
			this.txtPrimerNombre.ValueChanged += new System.EventHandler(this.txtPrimerNombre_ValueChanged);
			// 
			// label6
			// 
			this.label6.AutoSize = true;
			this.label6.BackColor = System.Drawing.Color.Transparent;
			this.label6.Location = new System.Drawing.Point(328, 48);
			this.label6.Name = "label6";
			this.label6.Size = new System.Drawing.Size(81, 16);
			this.label6.TabIndex = 50;
			this.label6.Text = "Primer Nombre";
			// 
			// txtApeMaterno
			// 
			appearance23.BackColor = System.Drawing.Color.Gainsboro;
			appearance23.BackColor2 = System.Drawing.Color.White;
			appearance23.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.txtApeMaterno.Appearance = appearance23;
			this.txtApeMaterno.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtApeMaterno.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsSolcitud, "Cre_Solicitud.ApellidoMaterno"));
			this.txtApeMaterno.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtApeMaterno.Location = new System.Drawing.Point(168, 64);
			this.txtApeMaterno.Name = "txtApeMaterno";
			this.txtApeMaterno.Size = new System.Drawing.Size(152, 21);
			this.txtApeMaterno.TabIndex = 6;
			this.txtApeMaterno.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.txtApeMaterno_KeyPress);
			// 
			// label5
			// 
			this.label5.AutoSize = true;
			this.label5.BackColor = System.Drawing.Color.Transparent;
			this.label5.Location = new System.Drawing.Point(168, 48);
			this.label5.Name = "label5";
			this.label5.Size = new System.Drawing.Size(89, 16);
			this.label5.TabIndex = 48;
			this.label5.Text = "Apellido Materno";
			// 
			// txtApePaterno
			// 
			appearance24.BackColor = System.Drawing.Color.GreenYellow;
			appearance24.BackColor2 = System.Drawing.Color.White;
			appearance24.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.txtApePaterno.Appearance = appearance24;
			this.txtApePaterno.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtApePaterno.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsSolcitud, "Cre_Solicitud.ApellidoPaterno"));
			this.txtApePaterno.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtApePaterno.Location = new System.Drawing.Point(8, 64);
			this.txtApePaterno.Name = "txtApePaterno";
			this.txtApePaterno.Size = new System.Drawing.Size(152, 21);
			this.txtApePaterno.TabIndex = 5;
			this.txtApePaterno.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.txtApePaterno_KeyPress);
			this.txtApePaterno.ValueChanged += new System.EventHandler(this.txtApePaterno_ValueChanged);
			// 
			// label4
			// 
			this.label4.AutoSize = true;
			this.label4.BackColor = System.Drawing.Color.Transparent;
			this.label4.Location = new System.Drawing.Point(8, 48);
			this.label4.Name = "label4";
			this.label4.Size = new System.Drawing.Size(87, 16);
			this.label4.TabIndex = 46;
			this.label4.Text = "Apellido Paterno";
			// 
			// label3
			// 
			this.label3.AutoSize = true;
			this.label3.BackColor = System.Drawing.Color.Transparent;
			this.label3.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label3.Location = new System.Drawing.Point(328, 8);
			this.label3.Name = "label3";
			this.label3.Size = new System.Drawing.Size(90, 16);
			this.label3.TabIndex = 44;
			this.label3.Text = "# DOCUMENTO";
			// 
			// cmbTipoDocumento
			// 
			appearance25.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(224)), ((System.Byte)(224)), ((System.Byte)(224)));
			appearance25.BackColor2 = System.Drawing.Color.White;
			appearance25.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.cmbTipoDocumento.Appearance = appearance25;
			this.cmbTipoDocumento.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbTipoDocumento.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbTipoDocumento.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsSolcitud, "Cre_Solicitud.TipoDocumento"));
			this.cmbTipoDocumento.DataMember = "Cre_TipoDocumento";
			this.cmbTipoDocumento.DataSource = this.cdsSolcitud;
			ultraGridBand10.AddButtonCaption = "RolSucursal";
			ultraGridColumn150.Header.VisiblePosition = 0;
			ultraGridColumn150.Hidden = true;
			ultraGridColumn151.Header.VisiblePosition = 1;
			ultraGridColumn151.Width = 50;
			ultraGridColumn152.Header.VisiblePosition = 2;
			ultraGridBand10.Columns.AddRange(new object[] {
																											ultraGridColumn150,
																											ultraGridColumn151,
																											ultraGridColumn152});
			this.cmbTipoDocumento.DisplayLayout.BandsSerializer.Add(ultraGridBand10);
			this.cmbTipoDocumento.DisplayMember = "Nombre";
			this.cmbTipoDocumento.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbTipoDocumento.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbTipoDocumento.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbTipoDocumento.Location = new System.Drawing.Point(168, 24);
			this.cmbTipoDocumento.Name = "cmbTipoDocumento";
			this.cmbTipoDocumento.Size = new System.Drawing.Size(120, 21);
			this.cmbTipoDocumento.TabIndex = 3;
			this.cmbTipoDocumento.ValueMember = "idTipoDoc";
			// 
			// label2
			// 
			this.label2.AutoSize = true;
			this.label2.BackColor = System.Drawing.Color.Transparent;
			this.label2.Location = new System.Drawing.Point(168, 8);
			this.label2.Name = "label2";
			this.label2.Size = new System.Drawing.Size(104, 16);
			this.label2.TabIndex = 42;
			this.label2.Text = "Tipo DOCUMENTO";
			// 
			// label1
			// 
			this.label1.AutoSize = true;
			this.label1.BackColor = System.Drawing.Color.Transparent;
			this.label1.Location = new System.Drawing.Point(8, 8);
			this.label1.Name = "label1";
			this.label1.Size = new System.Drawing.Size(99, 16);
			this.label1.TabIndex = 39;
			this.label1.Text = "Entidad Financiera";
			// 
			// cmbEntFinanciera
			// 
			appearance26.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(224)), ((System.Byte)(224)), ((System.Byte)(224)));
			appearance26.BackColor2 = System.Drawing.Color.White;
			appearance26.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.cmbEntFinanciera.Appearance = appearance26;
			this.cmbEntFinanciera.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbEntFinanciera.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbEntFinanciera.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsSolcitud, "Cre_Solicitud.EntidadFinanciera"));
			this.cmbEntFinanciera.DataMember = "Cre_EntidadFinanciera";
			this.cmbEntFinanciera.DataSource = this.cdsSolcitud;
			ultraGridBand11.AddButtonCaption = "_Cre_EntidadFinanciera";
			ultraGridColumn153.Header.VisiblePosition = 0;
			ultraGridColumn153.Hidden = true;
			ultraGridColumn154.Header.VisiblePosition = 1;
			ultraGridColumn154.Width = 50;
			ultraGridColumn155.Header.VisiblePosition = 2;
			ultraGridBand11.Columns.AddRange(new object[] {
																											ultraGridColumn153,
																											ultraGridColumn154,
																											ultraGridColumn155});
			this.cmbEntFinanciera.DisplayLayout.BandsSerializer.Add(ultraGridBand11);
			this.cmbEntFinanciera.DisplayMember = "Nombre";
			this.cmbEntFinanciera.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbEntFinanciera.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbEntFinanciera.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbEntFinanciera.Location = new System.Drawing.Point(8, 24);
			this.cmbEntFinanciera.Name = "cmbEntFinanciera";
			this.cmbEntFinanciera.Size = new System.Drawing.Size(152, 21);
			this.cmbEntFinanciera.TabIndex = 2;
			this.cmbEntFinanciera.ValueMember = "idEntidadFinanciera";
			this.cmbEntFinanciera.InitializeLayout += new Infragistics.Win.UltraWinGrid.InitializeLayoutEventHandler(this.cmbEntFinanciera_InitializeLayout);
			// 
			// errorProvider
			// 
			this.errorProvider.ContainerControl = this;
			// 
			// txtNumDocumento
			// 
			appearance27.BackColor = System.Drawing.Color.GreenYellow;
			appearance27.BackColor2 = System.Drawing.Color.White;
			appearance27.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.txtNumDocumento.Appearance = appearance27;
			this.txtNumDocumento.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsSolcitud, "Cre_Solicitud.NumeroDocumento"));
			this.txtNumDocumento.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtNumDocumento.Location = new System.Drawing.Point(328, 24);
			this.txtNumDocumento.Name = "txtNumDocumento";
			this.txtNumDocumento.Size = new System.Drawing.Size(120, 21);
			this.txtNumDocumento.TabIndex = 4;
			this.txtNumDocumento.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.txtNumDocumento_KeyPress);
			this.txtNumDocumento.Validating += new System.ComponentModel.CancelEventHandler(this.txtNumDocumento_Validating);
			this.txtNumDocumento.ValueChanged += new System.EventHandler(this.txtNumDocumento_ValueChanged);
			// 
			// txtidSolicitud
			// 
			this.txtidSolicitud.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			appearance28.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtidSolicitud.Appearance = appearance28;
			this.txtidSolicitud.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsSolcitud, "Cre_Solicitud.idSolicitud"));
			this.txtidSolicitud.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtidSolicitud.Enabled = false;
			this.txtidSolicitud.FormatString = "";
			this.txtidSolicitud.Location = new System.Drawing.Point(696, 8);
			this.txtidSolicitud.Name = "txtidSolicitud";
			this.txtidSolicitud.PromptChar = ' ';
			this.txtidSolicitud.Size = new System.Drawing.Size(24, 21);
			this.txtidSolicitud.TabIndex = 152;
			// 
			// barraDato1
			// 
			this.barraDato1.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.barraDato1.BarraMovimiento = true;
			this.barraDato1.DataMember = "Cre_Solicitud";
			this.barraDato1.DataNombreId = "idSolicitud";
			this.barraDato1.DataOrden = "";
			this.barraDato1.DataSource = this.cdsSolcitud;
			this.barraDato1.DataTabla = "Cre_Solicitud";
			this.barraDato1.DataTablaHija = "";
			this.barraDato1.Location = new System.Drawing.Point(8, 424);
			this.barraDato1.Name = "barraDato1";
			this.barraDato1.Size = new System.Drawing.Size(304, 32);
			this.barraDato1.TabIndex = 1;
			this.barraDato1.VisibleBorrar = true;
			this.barraDato1.VisibleBuscar = true;
			this.barraDato1.VisibleEditar = true;
			this.barraDato1.VisibleImprimir = false;
			this.barraDato1.VisibleNuevo = true;
			this.barraDato1.Buscar += new System.EventHandler(this.barraDato1_Buscar);
			this.barraDato1.Load += new System.EventHandler(this.barraDato1_Load);
			this.barraDato1.Borrar += new System.EventHandler(this.barraDato1_Borrar);
			this.barraDato1.Refresca += new System.EventHandler(this.barraDato1_Refresca);
			this.barraDato1.Crear += new System.EventHandler(this.barraDato1_Crear);
			this.barraDato1.Deshacer += new System.EventHandler(this.barraDato1_Deshacer);
			this.barraDato1.Editar += new System.EventHandler(this.barraDato1_Editar);
			this.barraDato1.Grabar += new System.EventHandler(this.barraDato1_Grabar);
			// 
			// Point_IRSolicitud
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 13);
			this.CausesValidation = false;
			this.ClientSize = new System.Drawing.Size(728, 462);
			this.Controls.Add(this.txtidSolicitud);
			this.Controls.Add(this.cmbEntFinanciera);
			this.Controls.Add(this.barraDato1);
			this.Controls.Add(this.groupBox1);
			this.Controls.Add(this.txtSegundoNombre);
			this.Controls.Add(this.label7);
			this.Controls.Add(this.label6);
			this.Controls.Add(this.label5);
			this.Controls.Add(this.label4);
			this.Controls.Add(this.label3);
			this.Controls.Add(this.label2);
			this.Controls.Add(this.label1);
			this.Controls.Add(this.txtPrimerNombre);
			this.Controls.Add(this.txtApeMaterno);
			this.Controls.Add(this.txtApePaterno);
			this.Controls.Add(this.txtNumDocumento);
			this.Controls.Add(this.cmbTipoDocumento);
			this.Controls.Add(this.groupBox2);
			this.Name = "Point_IRSolicitud";
			this.Text = "Solicitud ";
			this.Load += new System.EventHandler(this.Point_IRSolicitud_Load);
			this.groupBox2.ResumeLayout(false);
			((System.ComponentModel.ISupportInitialize)(this.rgGarante.Properties)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSolcitud)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbVendedor)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbLocal)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbTipTarjeta)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbTipoEstado)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbEstado)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbCalificacion)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbCallCenter)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbInstitucion)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraTextEditor3)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNumTarjeta)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtCuota)).EndInit();
			this.groupBox1.ResumeLayout(false);
			((System.ComponentModel.ISupportInitialize)(this.cmbParroquia)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtRefUbicacion)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNumTelefono)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtCalleTransversal)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNumCasa)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbBarrio)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbCanton)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbProvincia)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtAvPrincipal)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtSegundoNombre)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtPrimerNombre)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtApeMaterno)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtApePaterno)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbTipoDocumento)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbEntFinanciera)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNumDocumento)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtidSolicitud)).EndInit();
			this.ResumeLayout(false);

		}
		#endregion

		private C1.Data.C1DataRow drIRSolicitud;

		private void Point_IRSolicitud_Load(object sender, System.EventArgs e)
		{
			HabilitaControles(this.groupBox1, false);
			HabilitaControles(this.groupBox2, false);

			this.barraDato1.HabilitaControles(false);
			this.barraDato1.PosUltima();
		}

		private void cdsSolcitud_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			cdsSolcitud.Schema.Connections[0].ConnectionString = miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);		
		}

		private void barraDato1_Refresca(object sender, System.EventArgs e)
		{
			try
			{
				string sCliente = "Cliente: ";
				this.Text = sCliente + this.txtApePaterno.Text + " " + this.txtApeMaterno.Text + " " + this.txtPrimerNombre.Text + " " + this.txtSegundoNombre.Text;
				drIRSolicitud = cdsSolcitud.TableViews["Cre_Solicitud"].Rows[0];

				int iCodProv = (int) this.cmbProvincia.Value;
				string sSelect = string.Format("Exec SP_Filtro_CP '{0}', 1", iCodProv);
				this.cmbCanton.DataSource = Funcion.dvProcedimiento(cdsSolcitud, sSelect);
				this.cmbCanton.DisplayLayout.Bands[0].Columns["idCanton"].Hidden = true;
				this.cmbCanton.DisplayLayout.Bands[0].Columns["Nombre"].Width = 300;

				int iCodCant = (int) this.cmbCanton.Value;
				string sSel = string.Format("Exec SP_Filtro_CP '{0}', 2", iCodCant);
				this.cmbParroquia.DataSource = Funcion.dvProcedimiento(cdsSolcitud, sSel);
				this.cmbParroquia.DisplayLayout.Bands[0].Columns["idParroquia"].Hidden = true;
				this.cmbParroquia.DisplayLayout.Bands[0].Columns["Nombre"].Width = 300;

				if ((string) this.cmbEstado.Text != "")
				{
					string stSel = string.Format("Exec FiltroEstadoSolicitud {0}", this.cmbEstado.Value);
			
					if (Funcion.dvProcedimiento(cdsSolcitud, stSel).Count > 0)
					{		
						this.cmbTipoEstado.DataSource = Funcion.dvProcedimiento(cdsSolcitud, stSel);
					}		
				}
			}
			catch{};
		}

		private void barraDato1_Borrar(object sender, System.EventArgs e)
		{
			this.barraDato1.BorraRegistro();
		}

		private void barraDato1_Crear(object sender, System.EventArgs e)
		{
			this.barraDato1.CrearRegistro();
			this.txtidSolicitud.Width = 0;

			HabilitaControles(this.groupBox1, true);
			HabilitaControles(this.groupBox2, true);

			this.txtNumTarjeta.Enabled = false;
			this.cmbInstitucion.Enabled = false;
			this.cmbCallCenter.Enabled = false;
			this.cmbVendedor.Enabled = false;
			this.cmbTipoEstado.Enabled = false;
			this.txtCuota.Enabled = false;

			//drIRSolicitud["EntidadFinanciera"] = 1;
			drIRSolicitud["TipoDocumento"] = 1;

			//this.txtNumDocumento.Focus();
			this.cmbEntFinanciera.Select();
		}

		private void barraDato1_Deshacer(object sender, System.EventArgs e)
		{
			this.barraDato1.DeshacerRegistro();

			LimpiaError();

			HabilitaControles(this.groupBox1, false);
			HabilitaControles(this.groupBox2, false);
		}

		private void barraDato1_Editar(object sender, System.EventArgs e)
		{
			if (drIRSolicitud == null)
			{
				MessageBox.Show("No existe un registro para editar", "Información",
					MessageBoxButtons.OK, MessageBoxIcon.Information);
				return;
			}

			HabilitaControles(this.groupBox1, true);
			HabilitaControles(this.groupBox2, true);

			this.barraDato1.EditarRegistro();

			this.txtNumTarjeta.Enabled = false;
			
			if ((int) this.cmbTipTarjeta.Value == 1)
			{
				this.cmbInstitucion.Enabled = false;
				this.cmbCallCenter.Enabled = false;
				this.cmbVendedor.Enabled = false;
			}

			if ((int) this.cmbEstado.Value == 1)
			{
				this.cmbTipoEstado.Enabled = false;
				this.txtCuota.Enabled = true;
			}
			else
			{
				this.cmbTipoEstado.Enabled = false;
				this.txtCuota.Enabled = false;
			}
		}

		private void barraDato1_Grabar(object sender, System.EventArgs e)
		{
			if (!Verifica()) return;
			this.barraDato1.GrabaRegistro();
		
			HabilitaControles(this.groupBox1, false);
			HabilitaControles(this.groupBox2, false);

			int idSolicitud = (int) drIRSolicitud["idSolicitud"]; //(string)txtidSolicitud.Text.ToString();
			string sRUC = (string) drIRSolicitud["NumeroDocumento"];
			string stExec = string.Format("Exec Point_CreaCliente '{0}', {1}", sRUC, idSolicitud);
			string stCodigo = Funcion.sEscalarSQL(cdsSolcitud, stExec, true);
			
			if (stCodigo.Trim().Length > 0)
			{
				MessageBox.Show("No logro insertar datos en tabla cliente","Informacion"); 
				return;
			}
		}

		public void HabilitaControles(GroupBox Contenedor, bool Estado)
		{
			foreach (Control c in Contenedor.Controls) 
			{
				if (c is Infragistics.Win.UltraWinEditors.UltraTextEditor) 
				{
					c.Enabled = Estado;
				}
				if (c is Infragistics.Win.UltraWinGrid.UltraCombo) 
				{
					c.Enabled = Estado;
				}
				this.txtCuota.Enabled = Estado;
				this.rgGarante.Enabled = Estado;
			}
		}

		private void txtNumDocumento_Validating(object sender, System.ComponentModel.CancelEventArgs e)
		{
			if (txtNumDocumento.Text.ToString().Trim() != "")
			{
				string stExec = string.Format("Select dbo.ValidaCedula ('{0}')", txtNumDocumento.Text);
				string VlCed = Funcion.sEscalarSQL(cdsSolcitud, stExec);
				if (VlCed!="OK")
				{
					MessageBox.Show("# de Cédula Incorrecta");
					e.Cancel=true;
				}
			}
		
			string stSelect = string.Format("Exec buscaclientecredito '{0}'", this.txtNumDocumento.Text);
			int id = (int) Funcion.iEscalarSQL(cdsSolcitud, stSelect);
			
			if (id > 0)
			{
				this.barraDato1.IdRegistro = id;
				this.barraDato1.ProximoId(5);
				this.barraDato1.EditarRegistro();
			}
		}

		private void LimpiaError()
		{
			errorProvider.SetError(txtNumDocumento, "");  
			errorProvider.SetError(txtApePaterno, "");
			errorProvider.SetError(txtPrimerNombre, "");  
			errorProvider.SetError(cmbProvincia, "");
			errorProvider.SetError(cmbCanton, "");  
			errorProvider.SetError(cmbParroquia, "");
			errorProvider.SetError(cmbBarrio, "");  
			errorProvider.SetError(txtAvPrincipal, "");
			errorProvider.SetError(txtNumCasa, "");  
			errorProvider.SetError(txtCalleTransversal, "");
			errorProvider.SetError(txtNumTelefono, "");
			errorProvider.SetError(txtRefUbicacion, "");
		}

		private bool Verifica()
		{
			bool bOk = true;
		   
			LimpiaError();
			
			#region Verifica Campos Obligatorios
			if (txtNumDocumento.Text.ToString().Trim().Length == 0)
			{
				errorProvider.SetError(txtNumDocumento, "Ingrese Número de Documento");
				bOk = false;
			}
			if (txtApePaterno.Text.ToString().Trim().Length == 0)
			{
				errorProvider.SetError(txtApePaterno, "Ingrese Apellido Paterno");
				bOk = false;
			}
			if (txtPrimerNombre.Text.ToString().Trim().Length == 0)
			{
				errorProvider.SetError(txtPrimerNombre, "Ingrese El Primer Nombre");
				bOk = false;
			}			
			if (this.cmbProvincia.Value == System.DBNull.Value || (int) this.cmbProvincia.Value == 0) //.ActiveRow != null && cmbProvincia.ActiveRow.Cells["idProvincia"].Value != null && cmbProvincia.ActiveRow.Cells["idProvincia"].Value != DBNull.Value)
			{
				errorProvider.SetError(cmbProvincia, "Ingrese La Provinvia");
				bOk = false;
			}
			if (this.cmbCanton.Value == System.DBNull.Value || (int) this.cmbCanton.Value == 0) //.ActiveRow != null && cmbCanton.ActiveRow.Cells["idCanton"].Value != null && cmbCanton.ActiveRow.Cells["idCanton"].Value != DBNull.Value)
			{
				errorProvider.SetError(cmbCanton, "Ingrese Cantón");
				bOk = false;
			}
			if (this.cmbParroquia.Value == System.DBNull.Value || (int) this.cmbParroquia.Value == 0) //ActiveRow != null && cmbParroquia.ActiveRow.Cells["idParroquia"].Value != null && cmbParroquia.ActiveRow.Cells["idParroquia"].Value != DBNull.Value)
			{
				errorProvider.SetError(cmbParroquia, "Ingrese La Parroquia");
				bOk = false;
			}
			if (this.cmbBarrio.Value == System.DBNull.Value || (int) this.cmbBarrio.Value == 0) //.ActiveRow != null && cmbBarrio.ActiveRow.Cells["idBarrio"].Value != null && cmbBarrio.ActiveRow.Cells["idBarrio"].Value != DBNull.Value)
			{
				errorProvider.SetError(cmbBarrio, "Ingrese El Barrio");
				bOk = false;
			}
			if (txtAvPrincipal.Text.ToString().Trim().Length == 0)
			{
				errorProvider.SetError(txtAvPrincipal, "Ingrese Avenida Principal");
				bOk = false;
			}
			if (txtNumCasa.Text.ToString().Trim().Length == 0)
			{
				errorProvider.SetError(txtNumCasa, "Ingrese Número de Casa");
				bOk = false;
			}
			if (txtCalleTransversal.Text.ToString().Trim().Length == 0)
			{
				errorProvider.SetError(txtCalleTransversal, "Ingrese Calle Transversal");
				bOk = false;
			}
			if (txtNumTelefono.Text.ToString().Trim().Length == 0)
			{
				errorProvider.SetError(txtNumTelefono, "Ingrese Número de Teléfono");
				bOk = false;
			}
			if (txtRefUbicacion.Text.ToString().Trim().Length == 0)
			{
				errorProvider.SetError(txtRefUbicacion, "Ingrese Referencia de Ubicación");
				bOk = false;
			}
			return bOk;
			#endregion Verifica Campos Obligatorios
		}

		private void cmbEstado_Validated(object sender, System.EventArgs e)
		{
			/* Valida Estado Solicitud*/
			if ((string)cmbEstado.Text == "APROBADO")
			{
				string stExec = string.Format("Exec SP_Genera_Num_Tarjeta ");
				string NTarjt = Funcion.sEscalarSQL(cdsSolcitud, stExec);
				txtNumTarjeta.Value = (string)NTarjt;
			}
		}

		private void cmbTipTarjeta_Validated(object sender, System.EventArgs e)
		{
			/* Valida Tipo de Tarjeta*/
			if ((string)cmbTipTarjeta.Text == "NUEVA")
			{
				//				drIRSolicitud["Institucion"] = 0;
				//				drIRSolicitud["Vendedor"] = 0;
				//				drIRSolicitud["CallCenter"] = 0;
				this.cmbInstitucion.Value = 0;
				this.cmbVendedor.Value = 0;
				this.cmbCallCenter.Value = 0;

				cmbInstitucion.Enabled = false;
				cmbCallCenter.Enabled = false;
				cmbVendedor.Enabled = false;
			}
			else
			{
				cmbInstitucion.Enabled = true;
				cmbCallCenter.Enabled = true;
				cmbVendedor.Enabled = true;
			}	
		}

		private void barraDato1_Load(object sender, System.EventArgs e)
		{
		
		}

		private void cmbParroquia_InitializeLayout(object sender, Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs e)
		{
		
		}

		private void cmbProvincia_InitializeLayout(object sender, Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs e)
		{
		
		}

		private void cmbProvincia_ValueChanged(object sender, System.EventArgs e)
		{
			
		}

		private void cmbCanton_InitializeLayout(object sender, Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs e)
		{
		
		}

		private void cmbCanton_Validating(object sender, System.ComponentModel.CancelEventArgs e)
		{
			int iCodCant = (int) this.cmbCanton.Value;
			string sSelect = string.Format("Exec SP_Filtro_CP {0}, 2", iCodCant);
			this.cmbParroquia.DataSource = Funcion.dvProcedimiento(cdsSolcitud, sSelect);
			this.cmbParroquia.DisplayLayout.Bands[0].Columns["idParroquia"].Hidden = true;
			this.cmbParroquia.DisplayLayout.Bands[0].Columns["Nombre"].Width = 300;
		}

		private void cmbProvincia_Validating(object sender, System.ComponentModel.CancelEventArgs e)
		{
			int iCodProv = (int) this.cmbProvincia.Value;
			string sSelect = string.Format("Exec SP_Filtro_CP {0}, 1", iCodProv);
			this.cmbCanton.DataSource = Funcion.dvProcedimiento(cdsSolcitud, sSelect);
			this.cmbCanton.DisplayLayout.Bands[0].Columns["idCanton"].Hidden = true;
			this.cmbCanton.DisplayLayout.Bands[0].Columns["Nombre"].Width = 300;
		}

		private void cmbEstado_Validating(object sender, System.ComponentModel.CancelEventArgs e)
		{
			this.cmbTipoEstado.DataSource = null;

			if ((int) this.cmbEstado.Value > 0)
			{
				string stSel = string.Format("Exec FiltroEstadoSolicitud {0}", this.cmbEstado.Value);
			
				if (Funcion.dvProcedimiento(cdsSolcitud, stSel).Count > 0)
				{		
					this.cmbTipoEstado.DataSource = Funcion.dvProcedimiento(cdsSolcitud, stSel);
					this.cmbTipoEstado.DisplayLayout.Bands[0].Columns["idNegadoPendiente"].Hidden = true;
					this.cmbTipoEstado.DisplayLayout.Bands[0].Columns["Nombre"].Width = 150;
					this.cmbTipoEstado.Enabled = true;
					this.txtCuota.Enabled = false;
				}		
				else
				{
					this.cmbTipoEstado.Enabled = false;
					this.txtCuota.Enabled = true;
				}
			}
		}

		private void cmbEstado_InitializeLayout(object sender, Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs e)
		{
		
		}

		private void cmbTipTarjeta_Validating(object sender, System.ComponentModel.CancelEventArgs e)
		{
			
		}

		private void groupBox2_Enter(object sender, System.EventArgs e)
		{
		
		}

		private void cmbEntFinanciera_InitializeLayout(object sender, Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs e)
		{
		
		}

		private void cmbTipTarjeta_InitializeLayout(object sender, Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs e)
		{
		
		}

		private void barraDato1_Buscar(object sender, System.EventArgs e)
		{
			using (BuscaClienteCredito miBusqueda = new BuscaClienteCredito())
			{
				if (miBusqueda.ShowDialog() == DialogResult.OK)
				{
					if (miBusqueda.idBusca.Text.Length > 0)
					{
						this.barraDato1.IdRegistro = int.Parse(miBusqueda.idBusca.Text);
						this.barraDato1.ProximoId(5);
					}
				}
			}
		}

		private void ValidaCamposNumericos(object sender, KeyPressEventArgs e)
		{
			if (Char.IsDigit(e.KeyChar))
			{
				e.Handled = false;
			}
			else if (Char.IsControl(e.KeyChar))
			{
				e.Handled = false;
			}
			else if (Char.IsSeparator(e.KeyChar))
			{
				e.Handled = false;
			}
			else
			{
				e.Handled = true;
			}
		}

		private void txtNumTelefono_ValueChanged(object sender, System.EventArgs e)
		{
			
		}

		private void txtNumTelefono_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
			ValidaCamposNumericos(sender, e);

			if (e.KeyChar == 13)
				this.txtRefUbicacion.Focus();
		}

		private void txtNumDocumento_ValueChanged(object sender, System.EventArgs e)
		{
		
		}

		private void txtCuota_ValueChanged(object sender, System.EventArgs e)
		{
		
		}

		private void txtNumDocumento_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
			if (e.KeyChar == 13)
				this.txtApePaterno.Focus();
		}

		private void txtApePaterno_ValueChanged(object sender, System.EventArgs e)
		{
		
		}

		private void txtApePaterno_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
			if (e.KeyChar == 13)
				this.txtApeMaterno.Focus();
		}

		private void txtApeMaterno_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
			if (e.KeyChar == 13)
				this.txtPrimerNombre.Focus();
		}

		private void txtPrimerNombre_ValueChanged(object sender, System.EventArgs e)
		{
		
		}

		private void txtPrimerNombre_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
			if (e.KeyChar == 13)
				this.txtSegundoNombre.Focus();
		}

		private void txtSegundoNombre_ValueChanged(object sender, System.EventArgs e)
		{
		
		}

		private void txtSegundoNombre_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
			if (e.KeyChar == 13)
				
				this.cmbProvincia.Focus();
		}

		private void cmbProvincia_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
			if (e.KeyChar == 13)
				this.cmbCanton.Focus();
		}

		private void cmbCanton_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
			if (e.KeyChar == 13)
				this.cmbParroquia.Focus();
		}

		private void cmbParroquia_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
			if (e.KeyChar == 13)
				this.cmbBarrio.Focus();
		}

		private void cmbBarrio_InitializeLayout(object sender, Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs e)
		{
			
		}

		private void txtAvPrincipal_ValueChanged(object sender, System.EventArgs e)
		{
		
		}

		private void txtAvPrincipal_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
			if (e.KeyChar == 13)
				this.txtNumCasa.Focus();
		}

		private void cmbBarrio_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
			if (e.KeyChar == 13)
				this.txtAvPrincipal.Focus();
		}

		private void txtNumCasa_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
			if (e.KeyChar == 13)
				this.txtCalleTransversal.Focus();
		}

		private void txtCalleTransversal_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
			if (e.KeyChar == 13)
				this.txtNumTelefono.Focus();
		}

		private void txtRefUbicacion_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
			if (e.KeyChar == 13)
				this.cmbTipTarjeta.Focus();
		}

		private void cmbTipTarjeta_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
			if (e.KeyChar == 13)
				this.cmbLocal.Focus();
		}

	}
}
