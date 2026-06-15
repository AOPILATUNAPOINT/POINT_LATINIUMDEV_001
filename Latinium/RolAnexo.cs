using System;
using System.IO;
using System.Data.SqlClient;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using System.Globalization;
using Infragistics.Win.UltraWinGrid;
using CrystalDecisions.Shared;

namespace Latinium
{
	/// <summary>
	/// Descripción breve de RolAnexo.
	/// </summary>
	public class RolAnexo : DevExpress.XtraEditors.XtraForm
	{
		private Infragistics.Win.UltraWinGrid.UltraGrid ultraGrid1;
		private C1.Data.C1DataSet cdsPersonal;
		private C1.Data.C1DataSet cdsCiudad;
		private C1.Data.C1DataView cdvCiudad;
		private C1.Data.C1DataSet cdsProvincia;
		private C1.Data.C1DataView cdvProvincia;
		private Infragistics.Win.UltraWinGrid.UltraDropDown cmbCiudad;
		private Infragistics.Win.UltraWinGrid.UltraDropDown cmbProvincia;
		private Infragistics.Win.UltraWinSchedule.UltraCalendarCombo dtFecha;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource udsTipo;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource udsSistema;
		private Infragistics.Win.UltraWinGrid.UltraDropDown cmbTipo;
		private Infragistics.Win.UltraWinGrid.UltraDropDown cmbSalario;
		private Infragistics.Win.UltraWinEditors.UltraCheckEditor chkResumen;
		private C1.Data.C1DataView cdvRolPago;
		private C1.Data.C1DataSet cdsRolGrupo;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbRolGrupo;
		private DevExpress.XtraTab.XtraTabControl xtraTabControl1;
		private DevExpress.XtraTab.XtraTabPage xtraTabPage1;
		private DevExpress.XtraTab.XtraTabPage xtraTabPage2;
		private DevExpress.XtraEditors.PanelControl panelControl1;
		private DevExpress.XtraGrid.Views.Grid.GridView gridView2;
		private DevExpress.XtraEditors.SimpleButton btGeneraRDEP;
		private DevExpress.XtraEditors.SimpleButton btImprimir;
		private DevExpress.XtraGrid.GridControl grdRdep;
		private DevExpress.XtraGrid.Views.Grid.GridView grvRdep;
		private Infragistics.Win.Misc.UltraButton btArchivoEmpleado;
		private Infragistics.Win.Misc.UltraButton btGrabar;
		private DevExpress.XtraEditors.LookUpEdit cmbGrupo;
		private DevExpress.XtraTab.XtraTabPage xtraTabPage3;
		private DevExpress.XtraGrid.Views.Grid.GridView gridView3;
		private DevExpress.XtraGrid.GridControl grdDetalle;
		private DevExpress.XtraGrid.Views.Grid.GridView grvDetalle;
		private DevExpress.XtraEditors.SimpleButton btDetalle;
		private DevExpress.XtraEditors.SimpleButton btPesonal;
		private DevExpress.XtraEditors.SimpleButton btExcelDetalle;
		private DevExpress.XtraEditors.SimpleButton btExcelRDep;
		/// <summary>
		/// Variable del diseñador requerida.
		/// </summary>
		private System.ComponentModel.Container components = null;

		public RolAnexo()
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
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn23 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Tipo", -1, "cmbTipo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn24 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("FechaIngreso");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn25 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("FechaNacimiento");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn26 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("FechaSalida");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn27 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Notas");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn28 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("CargasFamiliares");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn29 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idProvincia", -1, "cmbProvincia");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn30 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCiudad", -1, "cmbCiudad");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn31 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Numero");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn32 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("SistemaSalario", -1, "cmbSalario");
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
			Infragistics.Win.Appearance appearance2 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn72 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("gEducacion");
			Infragistics.Win.Appearance appearance3 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn73 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("gVivienda");
			Infragistics.Win.Appearance appearance4 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn74 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("gVestimenta");
			Infragistics.Win.Appearance appearance5 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn75 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("gAlimentacion");
			Infragistics.Win.Appearance appearance6 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn76 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("gTotal");
			Infragistics.Win.Appearance appearance7 = new Infragistics.Win.Appearance();
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
			Infragistics.Win.Appearance appearance8 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance9 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance10 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance11 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand2 = new Infragistics.Win.UltraWinGrid.UltraGridBand("ClienteCiudad", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn106 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCiudad");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn107 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Ciudad");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn108 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Codigo");
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand3 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Provincia", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn109 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idProvincia");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn110 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("codigo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn111 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("provincia");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn112 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("borrar");
			Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton dateButton1 = new Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton();
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn1 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idTipo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn2 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("CedulaP");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn3 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idSistema");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn4 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Salario");
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand4 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn113 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idTipo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn114 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("CedulaP");
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand5 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn115 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idSistema");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn116 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Salario");
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand6 = new Infragistics.Win.UltraWinGrid.UltraGridBand("RolGrupo", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn117 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idRolGrupo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn118 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Grupo");
			this.ultraGrid1 = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.cdsPersonal = new C1.Data.C1DataSet();
			this.cdsCiudad = new C1.Data.C1DataSet();
			this.cdsProvincia = new C1.Data.C1DataSet();
			this.cdvCiudad = new C1.Data.C1DataView();
			this.cdvProvincia = new C1.Data.C1DataView();
			this.cmbCiudad = new Infragistics.Win.UltraWinGrid.UltraDropDown();
			this.cmbProvincia = new Infragistics.Win.UltraWinGrid.UltraDropDown();
			this.dtFecha = new Infragistics.Win.UltraWinSchedule.UltraCalendarCombo();
			this.udsTipo = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.udsSistema = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.cmbTipo = new Infragistics.Win.UltraWinGrid.UltraDropDown();
			this.cmbSalario = new Infragistics.Win.UltraWinGrid.UltraDropDown();
			this.btGrabar = new Infragistics.Win.Misc.UltraButton();
			this.chkResumen = new Infragistics.Win.UltraWinEditors.UltraCheckEditor();
			this.cdvRolPago = new C1.Data.C1DataView();
			this.cdsRolGrupo = new C1.Data.C1DataSet();
			this.cmbRolGrupo = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.xtraTabControl1 = new DevExpress.XtraTab.XtraTabControl();
			this.xtraTabPage1 = new DevExpress.XtraTab.XtraTabPage();
			this.cmbGrupo = new DevExpress.XtraEditors.LookUpEdit();
			this.btArchivoEmpleado = new Infragistics.Win.Misc.UltraButton();
			this.xtraTabPage2 = new DevExpress.XtraTab.XtraTabPage();
			this.grdRdep = new DevExpress.XtraGrid.GridControl();
			this.grvRdep = new DevExpress.XtraGrid.Views.Grid.GridView();
			this.gridView2 = new DevExpress.XtraGrid.Views.Grid.GridView();
			this.panelControl1 = new DevExpress.XtraEditors.PanelControl();
			this.btExcelRDep = new DevExpress.XtraEditors.SimpleButton();
			this.btPesonal = new DevExpress.XtraEditors.SimpleButton();
			this.btDetalle = new DevExpress.XtraEditors.SimpleButton();
			this.btGeneraRDEP = new DevExpress.XtraEditors.SimpleButton();
			this.btImprimir = new DevExpress.XtraEditors.SimpleButton();
			this.xtraTabPage3 = new DevExpress.XtraTab.XtraTabPage();
			this.btExcelDetalle = new DevExpress.XtraEditors.SimpleButton();
			this.grdDetalle = new DevExpress.XtraGrid.GridControl();
			this.grvDetalle = new DevExpress.XtraGrid.Views.Grid.GridView();
			this.gridView3 = new DevExpress.XtraGrid.Views.Grid.GridView();
			((System.ComponentModel.ISupportInitialize)(this.ultraGrid1)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsPersonal)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsCiudad)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsProvincia)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdvCiudad)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdvProvincia)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbCiudad)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbProvincia)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.dtFecha)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.udsTipo)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.udsSistema)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbTipo)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbSalario)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdvRolPago)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsRolGrupo)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbRolGrupo)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.xtraTabControl1)).BeginInit();
			this.xtraTabControl1.SuspendLayout();
			this.xtraTabPage1.SuspendLayout();
			((System.ComponentModel.ISupportInitialize)(this.cmbGrupo.Properties)).BeginInit();
			this.xtraTabPage2.SuspendLayout();
			((System.ComponentModel.ISupportInitialize)(this.grdRdep)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.grvRdep)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.gridView2)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.panelControl1)).BeginInit();
			this.panelControl1.SuspendLayout();
			this.xtraTabPage3.SuspendLayout();
			((System.ComponentModel.ISupportInitialize)(this.grdDetalle)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.grvDetalle)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.gridView3)).BeginInit();
			this.SuspendLayout();
			// 
			// ultraGrid1
			// 
			this.ultraGrid1.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.ultraGrid1.Cursor = System.Windows.Forms.Cursors.Hand;
			this.ultraGrid1.DataMember = "Personal";
			this.ultraGrid1.DataSource = this.cdsPersonal;
			appearance1.BackColor = System.Drawing.Color.White;
			this.ultraGrid1.DisplayLayout.Appearance = appearance1;
			ultraGridColumn1.Header.VisiblePosition = 15;
			ultraGridColumn1.Hidden = true;
			ultraGridColumn2.Header.VisiblePosition = 16;
			ultraGridColumn2.Hidden = true;
			ultraGridColumn3.Header.VisiblePosition = 17;
			ultraGridColumn3.Hidden = true;
			ultraGridColumn4.Header.VisiblePosition = 18;
			ultraGridColumn4.Hidden = true;
			ultraGridColumn5.Header.VisiblePosition = 19;
			ultraGridColumn5.Hidden = true;
			ultraGridColumn6.Header.VisiblePosition = 20;
			ultraGridColumn6.Hidden = true;
			ultraGridColumn7.Header.VisiblePosition = 21;
			ultraGridColumn7.Hidden = true;
			ultraGridColumn8.Header.VisiblePosition = 22;
			ultraGridColumn8.Hidden = true;
			ultraGridColumn9.Header.VisiblePosition = 2;
			ultraGridColumn10.Header.VisiblePosition = 23;
			ultraGridColumn10.Hidden = true;
			ultraGridColumn11.Header.VisiblePosition = 24;
			ultraGridColumn11.Hidden = true;
			ultraGridColumn12.Header.VisiblePosition = 25;
			ultraGridColumn12.Hidden = true;
			ultraGridColumn13.Header.VisiblePosition = 6;
			ultraGridColumn14.Header.VisiblePosition = 3;
			ultraGridColumn15.Header.VisiblePosition = 26;
			ultraGridColumn15.Hidden = true;
			ultraGridColumn16.Header.VisiblePosition = 1;
			ultraGridColumn17.Header.VisiblePosition = 27;
			ultraGridColumn17.Hidden = true;
			ultraGridColumn18.Header.VisiblePosition = 28;
			ultraGridColumn18.Hidden = true;
			ultraGridColumn19.Header.VisiblePosition = 29;
			ultraGridColumn19.Hidden = true;
			ultraGridColumn20.Header.VisiblePosition = 30;
			ultraGridColumn20.Hidden = true;
			ultraGridColumn21.Format = "n2";
			ultraGridColumn21.Header.VisiblePosition = 31;
			ultraGridColumn21.Hidden = true;
			ultraGridColumn22.Header.VisiblePosition = 32;
			ultraGridColumn22.Hidden = true;
			ultraGridColumn23.Header.VisiblePosition = 0;
			ultraGridColumn23.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.DropDownValidate;
			ultraGridColumn24.Header.VisiblePosition = 33;
			ultraGridColumn24.Hidden = true;
			ultraGridColumn25.Header.VisiblePosition = 34;
			ultraGridColumn25.Hidden = true;
			ultraGridColumn26.Header.VisiblePosition = 35;
			ultraGridColumn26.Hidden = true;
			ultraGridColumn27.Header.VisiblePosition = 36;
			ultraGridColumn27.Hidden = true;
			ultraGridColumn28.Header.VisiblePosition = 37;
			ultraGridColumn28.Hidden = true;
			ultraGridColumn29.Header.Caption = "Provincia";
			ultraGridColumn29.Header.VisiblePosition = 4;
			ultraGridColumn29.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.DropDownValidate;
			ultraGridColumn30.Header.Caption = "Ciudad";
			ultraGridColumn30.Header.VisiblePosition = 5;
			ultraGridColumn30.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.DropDownValidate;
			ultraGridColumn31.Header.VisiblePosition = 7;
			ultraGridColumn32.Header.VisiblePosition = 8;
			ultraGridColumn32.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.DropDownValidate;
			ultraGridColumn33.Header.VisiblePosition = 38;
			ultraGridColumn33.Hidden = true;
			ultraGridColumn34.Header.VisiblePosition = 39;
			ultraGridColumn34.Hidden = true;
			ultraGridColumn35.Header.VisiblePosition = 40;
			ultraGridColumn35.Hidden = true;
			ultraGridColumn36.Header.VisiblePosition = 41;
			ultraGridColumn36.Hidden = true;
			ultraGridColumn37.Header.VisiblePosition = 42;
			ultraGridColumn37.Hidden = true;
			ultraGridColumn38.Header.VisiblePosition = 43;
			ultraGridColumn38.Hidden = true;
			ultraGridColumn39.Header.VisiblePosition = 44;
			ultraGridColumn39.Hidden = true;
			ultraGridColumn40.Header.VisiblePosition = 45;
			ultraGridColumn40.Hidden = true;
			ultraGridColumn41.Header.VisiblePosition = 46;
			ultraGridColumn41.Hidden = true;
			ultraGridColumn42.Header.VisiblePosition = 47;
			ultraGridColumn42.Hidden = true;
			ultraGridColumn43.Header.VisiblePosition = 48;
			ultraGridColumn43.Hidden = true;
			ultraGridColumn44.Header.VisiblePosition = 49;
			ultraGridColumn44.Hidden = true;
			ultraGridColumn45.Header.VisiblePosition = 50;
			ultraGridColumn45.Hidden = true;
			ultraGridColumn46.Header.VisiblePosition = 51;
			ultraGridColumn46.Hidden = true;
			ultraGridColumn47.Header.VisiblePosition = 52;
			ultraGridColumn47.Hidden = true;
			ultraGridColumn48.Header.VisiblePosition = 53;
			ultraGridColumn48.Hidden = true;
			ultraGridColumn49.Header.VisiblePosition = 54;
			ultraGridColumn49.Hidden = true;
			ultraGridColumn50.Header.VisiblePosition = 55;
			ultraGridColumn50.Hidden = true;
			ultraGridColumn51.Header.VisiblePosition = 56;
			ultraGridColumn51.Hidden = true;
			ultraGridColumn52.Header.VisiblePosition = 57;
			ultraGridColumn52.Hidden = true;
			ultraGridColumn53.Header.VisiblePosition = 58;
			ultraGridColumn53.Hidden = true;
			ultraGridColumn54.Header.VisiblePosition = 59;
			ultraGridColumn54.Hidden = true;
			ultraGridColumn55.Header.VisiblePosition = 60;
			ultraGridColumn55.Hidden = true;
			ultraGridColumn56.Header.VisiblePosition = 61;
			ultraGridColumn56.Hidden = true;
			ultraGridColumn57.Header.VisiblePosition = 62;
			ultraGridColumn57.Hidden = true;
			ultraGridColumn58.Header.VisiblePosition = 63;
			ultraGridColumn58.Hidden = true;
			ultraGridColumn59.Header.VisiblePosition = 64;
			ultraGridColumn59.Hidden = true;
			ultraGridColumn60.Header.VisiblePosition = 65;
			ultraGridColumn60.Hidden = true;
			ultraGridColumn61.Header.VisiblePosition = 66;
			ultraGridColumn61.Hidden = true;
			ultraGridColumn62.Header.VisiblePosition = 67;
			ultraGridColumn62.Hidden = true;
			ultraGridColumn63.Header.VisiblePosition = 68;
			ultraGridColumn63.Hidden = true;
			ultraGridColumn64.Header.VisiblePosition = 69;
			ultraGridColumn64.Hidden = true;
			ultraGridColumn65.Header.VisiblePosition = 70;
			ultraGridColumn65.Hidden = true;
			ultraGridColumn66.Header.VisiblePosition = 71;
			ultraGridColumn66.Hidden = true;
			ultraGridColumn67.Header.VisiblePosition = 72;
			ultraGridColumn67.Hidden = true;
			ultraGridColumn68.Header.VisiblePosition = 73;
			ultraGridColumn68.Hidden = true;
			ultraGridColumn69.Header.VisiblePosition = 74;
			ultraGridColumn69.Hidden = true;
			ultraGridColumn70.Header.VisiblePosition = 75;
			ultraGridColumn70.Hidden = true;
			appearance2.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn71.CellAppearance = appearance2;
			ultraGridColumn71.Format = "n2";
			ultraGridColumn71.Header.VisiblePosition = 9;
			appearance3.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn72.CellAppearance = appearance3;
			ultraGridColumn72.Format = "n2";
			ultraGridColumn72.Header.VisiblePosition = 10;
			appearance4.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn73.CellAppearance = appearance4;
			ultraGridColumn73.Format = "n2";
			ultraGridColumn73.Header.VisiblePosition = 11;
			appearance5.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn74.CellAppearance = appearance5;
			ultraGridColumn74.Format = "n2";
			ultraGridColumn74.Header.VisiblePosition = 12;
			appearance6.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn75.CellAppearance = appearance6;
			ultraGridColumn75.Format = "n2";
			ultraGridColumn75.Header.VisiblePosition = 13;
			appearance7.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn76.CellAppearance = appearance7;
			ultraGridColumn76.Format = "n2";
			ultraGridColumn76.Header.VisiblePosition = 14;
			ultraGridColumn77.Header.VisiblePosition = 76;
			ultraGridColumn77.Hidden = true;
			ultraGridColumn78.Header.VisiblePosition = 77;
			ultraGridColumn78.Hidden = true;
			ultraGridColumn79.Formula = "n2";
			ultraGridColumn79.Header.VisiblePosition = 78;
			ultraGridColumn79.Hidden = true;
			ultraGridColumn80.Formula = "n2";
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
			ultraGridColumn96.Hidden = true;
			ultraGridColumn97.Header.VisiblePosition = 96;
			ultraGridColumn97.Hidden = true;
			ultraGridColumn98.Header.VisiblePosition = 97;
			ultraGridColumn98.Hidden = true;
			ultraGridColumn99.Header.VisiblePosition = 98;
			ultraGridColumn99.Hidden = true;
			ultraGridColumn100.Header.VisiblePosition = 99;
			ultraGridColumn100.Hidden = true;
			ultraGridColumn101.Header.VisiblePosition = 100;
			ultraGridColumn101.Hidden = true;
			ultraGridColumn102.Header.VisiblePosition = 101;
			ultraGridColumn102.Hidden = true;
			ultraGridColumn103.Header.VisiblePosition = 102;
			ultraGridColumn103.Hidden = true;
			ultraGridColumn104.Header.VisiblePosition = 103;
			ultraGridColumn104.Hidden = true;
			ultraGridColumn105.Header.VisiblePosition = 104;
			ultraGridColumn105.Hidden = true;
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
			this.ultraGrid1.DisplayLayout.BandsSerializer.Add(ultraGridBand1);
			this.ultraGrid1.DisplayLayout.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.No;
			this.ultraGrid1.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.True;
			this.ultraGrid1.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.True;
			appearance8.BackColor = System.Drawing.Color.Transparent;
			this.ultraGrid1.DisplayLayout.Override.CardAreaAppearance = appearance8;
			appearance9.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance9.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance9.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance9.FontData.BoldAsString = "True";
			appearance9.FontData.Name = "Arial";
			appearance9.FontData.SizeInPoints = 10F;
			appearance9.ForeColor = System.Drawing.Color.White;
			appearance9.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.ultraGrid1.DisplayLayout.Override.HeaderAppearance = appearance9;
			this.ultraGrid1.DisplayLayout.Override.HeaderClickAction = Infragistics.Win.UltraWinGrid.HeaderClickAction.SortMulti;
			appearance10.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance10.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance10.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ultraGrid1.DisplayLayout.Override.RowSelectorAppearance = appearance10;
			appearance11.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance11.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance11.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ultraGrid1.DisplayLayout.Override.SelectedRowAppearance = appearance11;
			this.ultraGrid1.DisplayLayout.ViewStyleBand = Infragistics.Win.UltraWinGrid.ViewStyleBand.OutlookGroupBy;
			this.ultraGrid1.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.ultraGrid1.Location = new System.Drawing.Point(8, 8);
			this.ultraGrid1.Name = "ultraGrid1";
			this.ultraGrid1.Size = new System.Drawing.Size(800, 281);
			this.ultraGrid1.TabIndex = 0;
			this.ultraGrid1.KeyDown += new System.Windows.Forms.KeyEventHandler(this.ultraGrid1_KeyDown);
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
			// cdsCiudad
			// 
			this.cdsCiudad.BindingContextCtrl = this;
			this.cdsCiudad.DataLibrary = "LibFacturacion";
			this.cdsCiudad.DataLibraryUrl = "";
			this.cdsCiudad.DataSetDef = "dsClienteTabla";
			this.cdsCiudad.Locale = new System.Globalization.CultureInfo("es-EC");
			this.cdsCiudad.Name = "cdsCiudad";
			this.cdsCiudad.SchemaClassName = "LibFacturacion.DataClass";
			this.cdsCiudad.SchemaDef = null;
			this.cdsCiudad.BeforeFill += new C1.Data.FillEventHandler(this.cdsCiudad_BeforeFill);
			// 
			// cdsProvincia
			// 
			this.cdsProvincia.BindingContextCtrl = this;
			this.cdsProvincia.DataLibrary = "LibPersonal";
			this.cdsProvincia.DataLibraryUrl = "";
			this.cdsProvincia.DataSetDef = "dsRolTablas";
			this.cdsProvincia.Locale = new System.Globalization.CultureInfo("es-EC");
			this.cdsProvincia.Name = "cdsProvincia";
			this.cdsProvincia.SchemaClassName = "LibPersonal.DataClass";
			this.cdsProvincia.SchemaDef = null;
			this.cdsProvincia.BeforeFill += new C1.Data.FillEventHandler(this.cdsProvincia_BeforeFill);
			// 
			// cdvCiudad
			// 
			this.cdvCiudad.BindingContextCtrl = this;
			this.cdvCiudad.DataSet = this.cdsCiudad;
			this.cdvCiudad.Sort = "Ciudad";
			this.cdvCiudad.TableName = "";
			this.cdvCiudad.TableViewName = "ClienteCiudad";
			// 
			// cdvProvincia
			// 
			this.cdvProvincia.BindingContextCtrl = this;
			this.cdvProvincia.DataSet = this.cdsProvincia;
			this.cdvProvincia.Sort = "Provincia";
			this.cdvProvincia.TableName = "";
			this.cdvProvincia.TableViewName = "Provincia";
			// 
			// cmbCiudad
			// 
			this.cmbCiudad.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbCiudad.DataSource = this.cdvCiudad;
			ultraGridColumn106.Header.VisiblePosition = 1;
			ultraGridColumn106.Hidden = true;
			ultraGridColumn107.Header.VisiblePosition = 0;
			ultraGridColumn108.Header.VisiblePosition = 2;
			ultraGridBand2.Columns.AddRange(new object[] {
																										 ultraGridColumn106,
																										 ultraGridColumn107,
																										 ultraGridColumn108});
			this.cmbCiudad.DisplayLayout.BandsSerializer.Add(ultraGridBand2);
			this.cmbCiudad.DisplayMember = "Ciudad";
			this.cmbCiudad.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbCiudad.Location = new System.Drawing.Point(528, 160);
			this.cmbCiudad.Name = "cmbCiudad";
			this.cmbCiudad.Size = new System.Drawing.Size(144, 78);
			this.cmbCiudad.TabIndex = 1;
			this.cmbCiudad.ValueMember = "idCiudad";
			this.cmbCiudad.Visible = false;
			// 
			// cmbProvincia
			// 
			this.cmbProvincia.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbProvincia.DataSource = this.cdvProvincia;
			ultraGridColumn109.Header.VisiblePosition = 1;
			ultraGridColumn109.Hidden = true;
			ultraGridColumn110.Header.VisiblePosition = 2;
			ultraGridColumn110.Hidden = true;
			ultraGridColumn111.Header.VisiblePosition = 0;
			ultraGridColumn112.Header.VisiblePosition = 3;
			ultraGridColumn112.Hidden = true;
			ultraGridBand3.Columns.AddRange(new object[] {
																										 ultraGridColumn109,
																										 ultraGridColumn110,
																										 ultraGridColumn111,
																										 ultraGridColumn112});
			this.cmbProvincia.DisplayLayout.BandsSerializer.Add(ultraGridBand3);
			this.cmbProvincia.DisplayMember = "provincia";
			this.cmbProvincia.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbProvincia.Location = new System.Drawing.Point(296, 152);
			this.cmbProvincia.Name = "cmbProvincia";
			this.cmbProvincia.Size = new System.Drawing.Size(128, 78);
			this.cmbProvincia.TabIndex = 2;
			this.cmbProvincia.ValueMember = "idProvincia";
			this.cmbProvincia.Visible = false;
			// 
			// dtFecha
			// 
			this.dtFecha.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			dateButton1.Caption = "Today";
			this.dtFecha.DateButtons.Add(dateButton1);
			this.dtFecha.Format = "yyyy";
			this.dtFecha.Location = new System.Drawing.Point(16, 8);
			this.dtFecha.Name = "dtFecha";
			this.dtFecha.NonAutoSizeHeight = 23;
			this.dtFecha.Size = new System.Drawing.Size(80, 21);
			this.dtFecha.SpinButtonIncrement = Infragistics.Win.UltraWinSchedule.SpinIncrementUnit.Years;
			this.dtFecha.SpinButtonsVisible = true;
			this.dtFecha.TabIndex = 4;
			this.dtFecha.Value = new System.DateTime(2009, 1, 1, 0, 0, 0, 0);
			this.dtFecha.BeforeDropDown += new System.ComponentModel.CancelEventHandler(this.dtFecha_BeforeDropDown);
			// 
			// udsTipo
			// 
			ultraDataColumn1.DataType = typeof(int);
			this.udsTipo.Band.Columns.AddRange(new object[] {
																												ultraDataColumn1,
																												ultraDataColumn2});
			this.udsTipo.Rows.AddRange(new object[] {
																								new Infragistics.Win.UltraWinDataSource.UltraDataRow(new object[] {
																																																										"idTipo",
																																																										2,
																																																										"CedulaP",
																																																										"Cédula"}, new object[0]),
																								new Infragistics.Win.UltraWinDataSource.UltraDataRow(new object[] {
																																																										"idTipo",
																																																										3,
																																																										"CedulaP",
																																																										"Pasaporte"}, new object[0])});
			this.udsTipo.CellDataRequested += new Infragistics.Win.UltraWinDataSource.CellDataRequestedEventHandler(this.udsTipo_CellDataRequested);
			// 
			// udsSistema
			// 
			ultraDataColumn3.DataType = typeof(int);
			this.udsSistema.Band.Columns.AddRange(new object[] {
																													 ultraDataColumn3,
																													 ultraDataColumn4});
			this.udsSistema.Rows.AddRange(new object[] {
																									 new Infragistics.Win.UltraWinDataSource.UltraDataRow(new object[] {
																																																											 "idSistema",
																																																											 1,
																																																											 "Salario",
																																																											 "Sin Salario Neto"}, new object[0]),
																									 new Infragistics.Win.UltraWinDataSource.UltraDataRow(new object[] {
																																																											 "idSistema",
																																																											 2,
																																																											 "Salario",
																																																											 "Con Salario Neto"}, new object[0])});
			this.udsSistema.CellDataRequested += new Infragistics.Win.UltraWinDataSource.CellDataRequestedEventHandler(this.udsSistema_CellDataRequested);
			// 
			// cmbTipo
			// 
			this.cmbTipo.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbTipo.DataMember = "Band 0";
			this.cmbTipo.DataSource = this.udsTipo;
			ultraGridColumn113.Header.VisiblePosition = 1;
			ultraGridColumn113.Hidden = true;
			ultraGridColumn114.Header.Caption = "Cedula Pasap";
			ultraGridColumn114.Header.VisiblePosition = 0;
			ultraGridBand4.Columns.AddRange(new object[] {
																										 ultraGridColumn113,
																										 ultraGridColumn114});
			this.cmbTipo.DisplayLayout.BandsSerializer.Add(ultraGridBand4);
			this.cmbTipo.DisplayMember = "CedulaP";
			this.cmbTipo.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbTipo.Location = new System.Drawing.Point(48, 112);
			this.cmbTipo.Name = "cmbTipo";
			this.cmbTipo.Size = new System.Drawing.Size(128, 69);
			this.cmbTipo.TabIndex = 5;
			this.cmbTipo.ValueMember = "idTipo";
			this.cmbTipo.Visible = false;
			// 
			// cmbSalario
			// 
			this.cmbSalario.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbSalario.DataMember = "Band 0";
			this.cmbSalario.DataSource = this.udsSistema;
			ultraGridColumn115.Header.VisiblePosition = 1;
			ultraGridColumn115.Hidden = true;
			ultraGridColumn116.Header.VisiblePosition = 0;
			ultraGridColumn116.Width = 132;
			ultraGridBand5.Columns.AddRange(new object[] {
																										 ultraGridColumn115,
																										 ultraGridColumn116});
			this.cmbSalario.DisplayLayout.BandsSerializer.Add(ultraGridBand5);
			this.cmbSalario.DisplayMember = "Salario";
			this.cmbSalario.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbSalario.Location = new System.Drawing.Point(368, 64);
			this.cmbSalario.Name = "cmbSalario";
			this.cmbSalario.Size = new System.Drawing.Size(208, 74);
			this.cmbSalario.TabIndex = 6;
			this.cmbSalario.ValueMember = "idSistema";
			this.cmbSalario.Visible = false;
			// 
			// btGrabar
			// 
			this.btGrabar.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btGrabar.Location = new System.Drawing.Point(16, 296);
			this.btGrabar.Name = "btGrabar";
			this.btGrabar.Size = new System.Drawing.Size(75, 24);
			this.btGrabar.TabIndex = 7;
			this.btGrabar.Text = "G&rabar";
			this.btGrabar.Click += new System.EventHandler(this.ultraButton2_Click);
			// 
			// chkResumen
			// 
			this.chkResumen.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.chkResumen.Location = new System.Drawing.Point(208, 8);
			this.chkResumen.Name = "chkResumen";
			this.chkResumen.Size = new System.Drawing.Size(72, 21);
			this.chkResumen.TabIndex = 9;
			this.chkResumen.Text = "Resumen";
			this.chkResumen.CheckedChanged += new System.EventHandler(this.chkResumen_CheckedChanged);
			// 
			// cdvRolPago
			// 
			this.cdvRolPago.BindingContextCtrl = this;
			this.cdvRolPago.DataSet = this.cdsRolGrupo;
			this.cdvRolPago.TableName = "";
			this.cdvRolPago.TableViewName = "RolPago";
			// 
			// cdsRolGrupo
			// 
			this.cdsRolGrupo.BindingContextCtrl = this;
			this.cdsRolGrupo.DataLibrary = "LibPersonal";
			this.cdsRolGrupo.DataLibraryUrl = "";
			this.cdsRolGrupo.DataSetDef = "dsRolGrupo";
			this.cdsRolGrupo.FillOnRequest = false;
			this.cdsRolGrupo.Locale = new System.Globalization.CultureInfo("es-EC");
			this.cdsRolGrupo.Name = "cdsRolGrupo";
			this.cdsRolGrupo.SchemaClassName = "LibPersonal.DataClass";
			this.cdsRolGrupo.SchemaDef = null;
			// 
			// cmbRolGrupo
			// 
			this.cmbRolGrupo.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.cmbRolGrupo.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbRolGrupo.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbRolGrupo.DataMember = "RolGrupo";
			this.cmbRolGrupo.DataSource = this.cdsRolGrupo;
			this.cmbRolGrupo.DisplayLayout.AutoFitColumns = true;
			ultraGridBand6.AddButtonCaption = "RolTipo";
			ultraGridColumn117.Header.VisiblePosition = 1;
			ultraGridColumn117.Hidden = true;
			ultraGridColumn118.Header.VisiblePosition = 0;
			ultraGridColumn118.Width = 133;
			ultraGridBand6.Columns.AddRange(new object[] {
																										 ultraGridColumn117,
																										 ultraGridColumn118});
			this.cmbRolGrupo.DisplayLayout.BandsSerializer.Add(ultraGridBand6);
			this.cmbRolGrupo.DisplayMember = "Grupo";
			this.cmbRolGrupo.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbRolGrupo.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbRolGrupo.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbRolGrupo.Location = new System.Drawing.Point(296, 8);
			this.cmbRolGrupo.Name = "cmbRolGrupo";
			this.cmbRolGrupo.Size = new System.Drawing.Size(152, 21);
			this.cmbRolGrupo.TabIndex = 16;
			this.cmbRolGrupo.ValueMember = "idRolGrupo";
			this.cmbRolGrupo.Visible = false;
			// 
			// xtraTabControl1
			// 
			this.xtraTabControl1.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.xtraTabControl1.Location = new System.Drawing.Point(16, 16);
			this.xtraTabControl1.Name = "xtraTabControl1";
			this.xtraTabControl1.SelectedTabPage = this.xtraTabPage1;
			this.xtraTabControl1.Size = new System.Drawing.Size(824, 360);
			this.xtraTabControl1.TabIndex = 20;
			this.xtraTabControl1.TabPages.AddRange(new DevExpress.XtraTab.XtraTabPage[] {
																																										this.xtraTabPage1,
																																										this.xtraTabPage2,
																																										this.xtraTabPage3});
			this.xtraTabControl1.Text = "xtraTabControl1";
			this.xtraTabControl1.Click += new System.EventHandler(this.xtraTabControl1_Click);
			// 
			// xtraTabPage1
			// 
			this.xtraTabPage1.Controls.Add(this.cmbGrupo);
			this.xtraTabPage1.Controls.Add(this.btArchivoEmpleado);
			this.xtraTabPage1.Controls.Add(this.cmbProvincia);
			this.xtraTabPage1.Controls.Add(this.cmbCiudad);
			this.xtraTabPage1.Controls.Add(this.cmbSalario);
			this.xtraTabPage1.Controls.Add(this.cmbTipo);
			this.xtraTabPage1.Controls.Add(this.ultraGrid1);
			this.xtraTabPage1.Controls.Add(this.btGrabar);
			this.xtraTabPage1.Name = "xtraTabPage1";
			this.xtraTabPage1.Size = new System.Drawing.Size(815, 330);
			this.xtraTabPage1.Text = "Personal";
			// 
			// cmbGrupo
			// 
			this.cmbGrupo.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.cmbGrupo.Location = new System.Drawing.Point(264, 297);
			this.cmbGrupo.Name = "cmbGrupo";
			// 
			// cmbGrupo.Properties
			// 
			this.cmbGrupo.Properties.Buttons.AddRange(new DevExpress.XtraEditors.Controls.EditorButton[] {
																																																		 new DevExpress.XtraEditors.Controls.EditorButton(DevExpress.XtraEditors.Controls.ButtonPredefines.Combo)});
			this.cmbGrupo.Properties.Columns.AddRange(new DevExpress.XtraEditors.Controls.LookUpColumnInfo[] {
																																																				 new DevExpress.XtraEditors.Controls.LookUpColumnInfo("Grupo", "Grupo", 49, DevExpress.Utils.FormatType.None, "", true, DevExpress.Utils.HorzAlignment.Near, DevExpress.Data.ColumnSortOrder.Ascending)});
			this.cmbGrupo.Properties.DisplayMember = "Grupo";
			this.cmbGrupo.Properties.NullText = "";
			this.cmbGrupo.Properties.PopupWidth = 300;
			this.cmbGrupo.Properties.ShowFooter = false;
			this.cmbGrupo.Properties.ShowHeader = false;
			this.cmbGrupo.Properties.ValueMember = "idRolGrupo";
			this.cmbGrupo.Size = new System.Drawing.Size(176, 20);
			this.cmbGrupo.TabIndex = 62;
			this.cmbGrupo.EditValueChanged += new System.EventHandler(this.cmbGrupo_EditValueChanged);
			// 
			// btArchivoEmpleado
			// 
			this.btArchivoEmpleado.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btArchivoEmpleado.Location = new System.Drawing.Point(112, 296);
			this.btArchivoEmpleado.Name = "btArchivoEmpleado";
			this.btArchivoEmpleado.Size = new System.Drawing.Size(136, 24);
			this.btArchivoEmpleado.TabIndex = 8;
			this.btArchivoEmpleado.Text = "Archivo empleados Txt";
			this.btArchivoEmpleado.Click += new System.EventHandler(this.btArchivoEmpleado_Click);
			// 
			// xtraTabPage2
			// 
			this.xtraTabPage2.Controls.Add(this.grdRdep);
			this.xtraTabPage2.Controls.Add(this.panelControl1);
			this.xtraTabPage2.Name = "xtraTabPage2";
			this.xtraTabPage2.Size = new System.Drawing.Size(815, 330);
			this.xtraTabPage2.Text = "RDEP Xml";
			// 
			// grdRdep
			// 
			this.grdRdep.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			// 
			// grdRdep.EmbeddedNavigator
			// 
			this.grdRdep.EmbeddedNavigator.Name = "";
			this.grdRdep.Location = new System.Drawing.Point(8, 8);
			this.grdRdep.MainView = this.grvRdep;
			this.grdRdep.Name = "grdRdep";
			this.grdRdep.Size = new System.Drawing.Size(800, 273);
			this.grdRdep.TabIndex = 20;
			this.grdRdep.ViewCollection.AddRange(new DevExpress.XtraGrid.Views.Base.BaseView[] {
																																													 this.grvRdep,
																																													 this.gridView2});
			this.grdRdep.DoubleClick += new System.EventHandler(this.grdRdep_DoubleClick);
			// 
			// grvRdep
			// 
			this.grvRdep.GridControl = this.grdRdep;
			this.grvRdep.Name = "grvRdep";
			this.grvRdep.OptionsSelection.MultiSelect = true;
			this.grvRdep.OptionsSelection.MultiSelectMode = DevExpress.XtraGrid.Views.Grid.GridMultiSelectMode.CellSelect;
			this.grvRdep.OptionsView.ColumnAutoWidth = false;
			this.grvRdep.OptionsView.ShowAutoFilterRow = true;
			// 
			// gridView2
			// 
			this.gridView2.GridControl = this.grdRdep;
			this.gridView2.Name = "gridView2";
			// 
			// panelControl1
			// 
			this.panelControl1.Controls.Add(this.btExcelRDep);
			this.panelControl1.Controls.Add(this.btPesonal);
			this.panelControl1.Controls.Add(this.btDetalle);
			this.panelControl1.Controls.Add(this.btGeneraRDEP);
			this.panelControl1.Controls.Add(this.btImprimir);
			this.panelControl1.Controls.Add(this.chkResumen);
			this.panelControl1.Controls.Add(this.cmbRolGrupo);
			this.panelControl1.Controls.Add(this.dtFecha);
			this.panelControl1.Dock = System.Windows.Forms.DockStyle.Bottom;
			this.panelControl1.Location = new System.Drawing.Point(0, 290);
			this.panelControl1.Name = "panelControl1";
			this.panelControl1.Size = new System.Drawing.Size(815, 40);
			this.panelControl1.TabIndex = 0;
			this.panelControl1.Text = "panelControl1";
			// 
			// btExcelRDep
			// 
			this.btExcelRDep.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btExcelRDep.Location = new System.Drawing.Point(488, 8);
			this.btExcelRDep.Name = "btExcelRDep";
			this.btExcelRDep.Size = new System.Drawing.Size(72, 25);
			this.btExcelRDep.TabIndex = 23;
			this.btExcelRDep.Text = "&Excel";
			this.btExcelRDep.Click += new System.EventHandler(this.btExcelRDep_Click);
			// 
			// btPesonal
			// 
			this.btPesonal.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btPesonal.Location = new System.Drawing.Point(568, 8);
			this.btPesonal.Name = "btPesonal";
			this.btPesonal.Size = new System.Drawing.Size(72, 25);
			this.btPesonal.TabIndex = 18;
			this.btPesonal.Text = "&Personal";
			this.btPesonal.Click += new System.EventHandler(this.btPesonal_Click);
			// 
			// btDetalle
			// 
			this.btDetalle.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btDetalle.Location = new System.Drawing.Point(728, 8);
			this.btDetalle.Name = "btDetalle";
			this.btDetalle.Size = new System.Drawing.Size(72, 25);
			this.btDetalle.TabIndex = 17;
			this.btDetalle.Text = "&Detalle";
			this.btDetalle.Click += new System.EventHandler(this.btDetalle_Click);
			// 
			// btGeneraRDEP
			// 
			this.btGeneraRDEP.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btGeneraRDEP.Location = new System.Drawing.Point(112, 8);
			this.btGeneraRDEP.Name = "btGeneraRDEP";
			this.btGeneraRDEP.Size = new System.Drawing.Size(72, 25);
			this.btGeneraRDEP.TabIndex = 1;
			this.btGeneraRDEP.Text = "&Generar";
			this.btGeneraRDEP.Click += new System.EventHandler(this.btGeneraRDEP_Click);
			this.btGeneraRDEP.MouseUp += new System.Windows.Forms.MouseEventHandler(this.btGeneraRDEP_MouseUp);
			// 
			// btImprimir
			// 
			this.btImprimir.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btImprimir.Location = new System.Drawing.Point(648, 8);
			this.btImprimir.Name = "btImprimir";
			this.btImprimir.Size = new System.Drawing.Size(72, 25);
			this.btImprimir.TabIndex = 7;
			this.btImprimir.Text = "&Imprimir";
			this.btImprimir.Click += new System.EventHandler(this.btImprimir_Click);
			this.btImprimir.MouseDown += new System.Windows.Forms.MouseEventHandler(this.btImprimir_MouseDown);
			// 
			// xtraTabPage3
			// 
			this.xtraTabPage3.Controls.Add(this.btExcelDetalle);
			this.xtraTabPage3.Controls.Add(this.grdDetalle);
			this.xtraTabPage3.Name = "xtraTabPage3";
			this.xtraTabPage3.Size = new System.Drawing.Size(815, 330);
			this.xtraTabPage3.Text = "Detalle";
			// 
			// btExcelDetalle
			// 
			this.btExcelDetalle.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btExcelDetalle.Location = new System.Drawing.Point(16, 289);
			this.btExcelDetalle.Name = "btExcelDetalle";
			this.btExcelDetalle.Size = new System.Drawing.Size(72, 25);
			this.btExcelDetalle.TabIndex = 22;
			this.btExcelDetalle.Text = "&Excel";
			this.btExcelDetalle.Click += new System.EventHandler(this.btExcelDetalle_Click);
			// 
			// grdDetalle
			// 
			this.grdDetalle.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			// 
			// grdDetalle.EmbeddedNavigator
			// 
			this.grdDetalle.EmbeddedNavigator.Name = "";
			this.grdDetalle.Location = new System.Drawing.Point(8, 8);
			this.grdDetalle.MainView = this.grvDetalle;
			this.grdDetalle.Name = "grdDetalle";
			this.grdDetalle.Size = new System.Drawing.Size(800, 273);
			this.grdDetalle.TabIndex = 21;
			this.grdDetalle.ViewCollection.AddRange(new DevExpress.XtraGrid.Views.Base.BaseView[] {
																																															this.grvDetalle,
																																															this.gridView3});
			// 
			// grvDetalle
			// 
			this.grvDetalle.GridControl = this.grdDetalle;
			this.grvDetalle.Name = "grvDetalle";
			this.grvDetalle.OptionsSelection.MultiSelect = true;
			this.grvDetalle.OptionsSelection.MultiSelectMode = DevExpress.XtraGrid.Views.Grid.GridMultiSelectMode.CellSelect;
			this.grvDetalle.OptionsView.ShowAutoFilterRow = true;
			// 
			// gridView3
			// 
			this.gridView3.GridControl = this.grdDetalle;
			this.gridView3.Name = "gridView3";
			// 
			// RolAnexo
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 14);
			this.ClientSize = new System.Drawing.Size(848, 385);
			this.Controls.Add(this.xtraTabControl1);
			this.Name = "RolAnexo";
			this.Text = "Anexo de Relacion de Dependencia";
			this.Closing += new System.ComponentModel.CancelEventHandler(this.RolAnexo_Closing);
			this.Load += new System.EventHandler(this.RolAnexo_Load);
			((System.ComponentModel.ISupportInitialize)(this.ultraGrid1)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsPersonal)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsCiudad)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsProvincia)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdvCiudad)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdvProvincia)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbCiudad)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbProvincia)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.dtFecha)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.udsTipo)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.udsSistema)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbTipo)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbSalario)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdvRolPago)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsRolGrupo)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbRolGrupo)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.xtraTabControl1)).EndInit();
			this.xtraTabControl1.ResumeLayout(false);
			this.xtraTabPage1.ResumeLayout(false);
			((System.ComponentModel.ISupportInitialize)(this.cmbGrupo.Properties)).EndInit();
			this.xtraTabPage2.ResumeLayout(false);
			((System.ComponentModel.ISupportInitialize)(this.grdRdep)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.grvRdep)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.gridView2)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.panelControl1)).EndInit();
			this.panelControl1.ResumeLayout(false);
			this.xtraTabPage3.ResumeLayout(false);
			((System.ComponentModel.ISupportInitialize)(this.grdDetalle)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.grvDetalle)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.gridView3)).EndInit();
			this.ResumeLayout(false);

		}
		#endregion

		private string LlenaCampo(string stCampo, int iEspacio)
		{
			stCampo = stCampo.Trim();
			if (stCampo.Length > iEspacio) stCampo = stCampo.Substring(0, iEspacio);
			stCampo = stCampo.PadLeft(iEspacio);
			return stCampo;
		}


		private string ValidosSolo(string stMensaje)
		{
			stMensaje = stMensaje.Replace(".", "");
			stMensaje = stMensaje.Replace("á", "a");
			stMensaje = stMensaje.Replace("é", "e");
			stMensaje = stMensaje.Replace("í", "i");
			stMensaje = stMensaje.Replace("ó", "o");
			stMensaje = stMensaje.Replace("ú", "u");
			stMensaje = stMensaje.Replace("ñ", "u");
			stMensaje = stMensaje.Replace("Ñ", "u");
			stMensaje = stMensaje.Replace("&", "");
			stMensaje = stMensaje.Replace("(", "");
			stMensaje = stMensaje.Replace("°", "");
			stMensaje = stMensaje.Replace("!", "");
			stMensaje = stMensaje.Replace("#", "");
			stMensaje = stMensaje.Replace("$", "");
			stMensaje = stMensaje.Replace("%", "");
			stMensaje = stMensaje.Replace("&", "");
			stMensaje = stMensaje.Replace("/", "");
			stMensaje = stMensaje.Replace("=", "");
			stMensaje = stMensaje.Replace("?", "");
			stMensaje = stMensaje.Replace("¡", "");
			stMensaje = stMensaje.Replace("+", "");
			stMensaje = stMensaje.Replace("}", "");
			stMensaje = stMensaje.Replace("{", "");
			stMensaje = stMensaje.Replace("-", "");
			stMensaje = stMensaje.Replace(".", "");
			stMensaje = stMensaje.Replace("-", "");
			return stMensaje;
		}

		private void cdsPersonal_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			this.cdsPersonal.Schema.Connections[0].ConnectionString = 
				miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);
		}

		private void cdsCiudad_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			this.cdsCiudad.Schema.Connections[0].ConnectionString = 
				miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);
		}

		private void cdsProvincia_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			this.cdsProvincia.Schema.Connections[0].ConnectionString = 
				miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);
		}
		CultureInfo us = new CultureInfo("en-US");

		private void RolAnexo_Load(object sender, System.EventArgs e)
		{
			string stAudita = "Exec AuditaCrear 53, 6, 'Rol Anexo'";
			Funcion.EjecutaSQL(cdsCiudad, stAudita, true);
      dtFecha.Value = DateTime.Today.AddYears(-1);
			cmbGrupo.Properties.DataSource = Funcion.dvProcedimiento(cdsCiudad, "Select * From RolGrupo");
	}

		private void ultraButton2_Click(object sender, System.EventArgs e)
		{
			if (DialogResult.No == MessageBox.Show("Desea grabar los cambios en Personal",
				"Información", MessageBoxButtons.YesNo, MessageBoxIcon.Question))
				return;
			try
			{
				cdsPersonal.Update();
			}
			catch(System.Exception ex)
			{
				MessageBox.Show(ex.Message, "Error al Grabar");
				return;
			}
			MessageBox.Show("Datos Guardados", "Información");
		}

		private void RolAnexo_Closing(object sender, System.ComponentModel.CancelEventArgs e)
		{
			string stAudita = "Exec AuditaCrear 53, 7, 'Rol Anexo'";
			Funcion.EjecutaSQL(cdsCiudad, stAudita, true);
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

		private void btImprimir_Click(object sender, System.EventArgs e)
		{
			DateTime dtAnio = (DateTime) dtFecha.Value;
			Cursor = Cursors.WaitCursor;
			int iAnio = dtAnio.Year;
			ParameterFields paramFields = new ParameterFields ();
			ParameterField paramField = new ParameterField ();
			ParameterDiscreteValue discreteVal = new ParameterDiscreteValue ();
			paramField.ParameterFieldName = "@Anio";
			discreteVal.Value = iAnio;
			paramField.CurrentValues.Add (discreteVal);
			paramFields.Add (paramField);
		
			Reporte miRep1 = new Reporte("RolAnexo.Rpt");
			miRep1.crVista.ParameterFieldInfo = paramFields;
			miRep1.MdiParent = this.MdiParent;
			if (chkResumen.Checked) miRep1.Resumen();
			miRep1.Titulo(iAnio.ToString());
			miRep1.Show();
			Cursor = Cursors.Default;
		}

		private void dtFecha_BeforeDropDown(object sender, System.ComponentModel.CancelEventArgs e)
		{
		
		}

		private void udsSistema_CellDataRequested(object sender, Infragistics.Win.UltraWinDataSource.CellDataRequestedEventArgs e)
		{
		
		}

		private void udsTipo_CellDataRequested(object sender, Infragistics.Win.UltraWinDataSource.CellDataRequestedEventArgs e)
		{
		
		}

		private void btImprimir_MouseDown(object sender, System.Windows.Forms.MouseEventArgs e)
		{
			if (e.Button != MouseButtons.Right) return;

			DateTime dtAnio = (DateTime) dtFecha.Value;
			Cursor = Cursors.WaitCursor;
			int iAnio = dtAnio.Year;
			ParameterFields paramFields = new ParameterFields ();
			ParameterField paramField = new ParameterField ();
			ParameterDiscreteValue discreteVal = new ParameterDiscreteValue ();
			paramField.ParameterFieldName = "@Anio";
			discreteVal.Value = iAnio;
			paramField.CurrentValues.Add (discreteVal);
			paramFields.Add (paramField);
		
			Reporte miRep = new Reporte("RolAnexoResumen.Rpt");
			miRep.crVista.ParameterFieldInfo = paramFields;
			miRep.MdiParent = this.MdiParent;
			if (chkResumen.Checked) miRep.Resumen();
			miRep.Titulo(iAnio.ToString());
			miRep.Show();
			Cursor = Cursors.Default;
		}

		private void ConfiguraGrilla(string stTipo)
		{
			DevExpress.XtraGrid.Views.Base.ColumnView View = null;
			if (stTipo == "RDEP")
			{
				View = grdRdep.MainView as DevExpress.XtraGrid.Views.Base.ColumnView;
			}
			if (stTipo == "Detalle")
			{
				View = grdDetalle.MainView as DevExpress.XtraGrid.Views.Base.ColumnView;
			}
			foreach (DevExpress.XtraGrid.Columns.GridColumn col in View.Columns)
			{
				col.OptionsColumn.AllowEdit = false;
				if (col.FieldName.StartsWith("id"))
					col.Visible = false;

				if (!(col.FieldName.StartsWith("Ced") || col.FieldName.StartsWith("Calle")
					|| col.FieldName.StartsWith("Numero") || col.FieldName.StartsWith("Ciudad")
					|| col.FieldName.StartsWith("Sistema") || col.FieldName.StartsWith("Año")
					|| col.FieldName.StartsWith("Provincia") || col.FieldName.StartsWith("Telefono")) )
				{
					col.DisplayFormat.FormatString = "n2";
					col.DisplayFormat.FormatType = DevExpress.Utils.FormatType.Numeric;
				}
			}
		}

		private string GeneraGrilla()
		{
			DateTime dtFecha1 = (DateTime) dtFecha.Value;

			string stGrupo = "";
			//			string stGrupo = ", 0";
			//			if (cmbRolGrupo.Value != null && (int) cmbRolGrupo.Value > 0)
			//				stGrupo = ", " + cmbRolGrupo.Value.ToString();
			string stExec = "Exec PersonalDatos " + dtFecha1.ToString("yyyy") + stGrupo;
			grvRdep.Columns.Clear();
			grdRdep.DataSource = Funcion.dvProcedimiento(cdsCiudad, stExec);
			ConfiguraGrilla("RDEP");
			ConfiguraGrilla("RDEP");
			return stExec;
		}

		private void btGeneraRDEP_Click(object sender, System.EventArgs e)
		{
			#region Variables
			Cursor = Cursors.WaitCursor;

			DateTime dtFecha1 = (DateTime) dtFecha.Value;
			#endregion Variables
			#region Creacion de Archivo
			StreamWriter sr = null;
			string stDir = @"C:\XML_ANEXO_TRANSACCIONAL\";
			string stNombreArchivo = "";
			try // Si no existe la carpeta C:\XML_ANEXO_TRANSACCIONAL\ se crea en la carpeta que inicio el sistema
			{
				if (!Directory.Exists(stDir)) stDir = MenuLatinium.stDirInicio + @"\";
				stNombreArchivo = @stDir + "RDEP" + dtFecha1.ToString("yyyy") + ".Xml";
				sr= File.CreateText(stNombreArchivo);
			}
			catch
			{
				MessageBox.Show("Esta abierto el archivo RDEP" + dtFecha1.ToString("yyyy") + ".Xml");
				Cursor = Cursors.Default;
			}
			#endregion Creacion de Archivo
			string stExec = GeneraGrilla();

			#region Encabezado
			string stRucEmpresa = "Select Ruc From seteo";
			stRucEmpresa = Funcion.sEscalarSQL(cdsPersonal, stRucEmpresa, true);
			sr.WriteLine("<?xml version=\"1.0\" encoding=\"ISO-8859-1\" standalone=\"yes\"?>");
			sr.WriteLine("<rdep>");
			sr.WriteLine("<numRuc>" + stRucEmpresa + "</numRuc> ");
			sr.WriteLine("<anio>" + dtFecha1.ToString("yyyy") + "</anio>");
			sr.WriteLine("<retRelDep>");
			#endregion Encabezado

			#region Lectura de Datos
			SqlDataReader miReader = Funcion.rEscalarSQL(cdsCiudad, stExec, true);
			while (miReader.Read())
			{
				#region Cedula
				int iCedulaP = miReader.GetInt32(1);
				string stCedula = miReader.GetString(2).Trim();
				int idPersonal = miReader.GetInt32(36);
				string stVerifica = Funcion.VerificaCedula(stCedula, false);
				if (stVerifica.Length > 0 && iCedulaP == 2)
				{
					miReader.Close();
					sr.Close();
					MessageBox.Show("Error en cédula: " + stCedula, stVerifica);
					RolPersonal miRol = new RolPersonal(idPersonal);
					miRol.ShowDialog();
					Cursor = Cursors.Default;
					return;
				}
				#endregion Cedula
				#region Datos Personales
				string stDireccion = miReader.GetString(3).Trim();
				string stNumero = miReader.GetString(4).Trim();
				string stCiudad = miReader.GetString(5).Trim();
				string stProvCodigo = miReader.GetString(6).Trim();
				string stTelefono = miReader.GetString(7).Trim();
				#endregion Datos Generales
				#region Verificacion de Datos
				string stErrorDato = "";
				if (stDireccion.Length < 2) stErrorDato = "Ingrese Direccion\n";
				if (stNumero.Length < 1) stErrorDato += "Ingrese Numero de la calle\n";
				if (stCiudad.Length < 1) stErrorDato += "Ingrese Ciudad\n";
				if (stProvCodigo.Length < 1) stErrorDato += "Ingrese Provincia\n";
				if (stTelefono.Length != 9) stErrorDato += "Telefono debe ser de 9 digitos\n";
				if (!(iCedulaP == 2 || iCedulaP == 3)) stErrorDato += "Ingrese Tipo Cedula o Pasaporte\n";
				if (stErrorDato.Length > 0)
				{
					miReader.Close();
					sr.Close();
					MessageBox.Show(stErrorDato, "Error en datos del Personal");
					RolPersonal miRol = new RolPersonal(idPersonal);
					miRol.ShowDialog();
					Cursor = Cursors.Default;
					return;
				}
				#endregion Verificaciono de Datos
				#region Valores de Roles
				decimal dSueldo = miReader.GetDecimal(9);
				decimal dSobreSueldo = miReader.GetDecimal(10);
				decimal dDecimoXIII = miReader.GetDecimal(11);
				decimal dDecimoXIV = miReader.GetDecimal(12);
				decimal dFondoReserva = miReader.GetDecimal(13);
				decimal dSalarioDigno = miReader.GetDecimal(14);
				decimal dUtilidades = miReader.GetDecimal(15);
				decimal dgDesahucio = miReader.GetDecimal(16);
				decimal dAporteIess = miReader.GetDecimal(17);
				double dVivienda = miReader.GetDouble(18);
				double dSalud = miReader.GetDouble(19);
				double dEducacion = miReader.GetDouble(20);
				double dAlimentacion = miReader.GetDouble(21);
				double dVestimenta = miReader.GetDouble(22);

				double dgDiscapacidad = miReader.GetDouble(23);
				double dgRebajaIIIEdad = miReader.GetDouble(24);
				decimal ImpRentaAsumido = miReader.GetDecimal(25);
				decimal dSubTotal = miReader.GetDecimal(26);
				int iCuentaRet = miReader.GetInt32(27);
				int iMeses = miReader.GetInt32(28);
				double intGravGen = miReader.GetDouble(29);

				double DeduccionOE = miReader.GetDouble(30);
				double OtrasRebOE = miReader.GetDouble(31);

				decimal dBaseRetenido = miReader.GetDecimal(32);
				decimal dImpRentaCaus = miReader.GetDecimal(33);
				decimal IngGravOE  = miReader.GetDecimal(34);
				decimal dValRet = miReader.GetDecimal(35);
				double dRentaRetOE = 0;

				#endregion Valores de Roles
				#region Grabacion de Linea
				sr.WriteLine("<datRetRelDep>");
				int iSistemaSalario = miReader.GetInt32(8);
				sr.WriteLine("<tipIdRet>" + iCedulaP.ToString() + "</tipIdRet>");
				sr.WriteLine("<idRet>" + stCedula + "</idRet>");
				if (stDireccion.Length >= 20)
					stDireccion = ValidosSolo(stDireccion.Substring(0, 20));
				sr.WriteLine("<dirCal>" + stDireccion + "</dirCal>");
				sr.WriteLine("<dirNum>" + stNumero + "</dirNum>");
				sr.WriteLine("<dirCiu>" + stCiudad + "</dirCiu>");
				sr.WriteLine("<dirProv>" + stProvCodigo + "</dirProv>");
				sr.WriteLine("<tel>" + stTelefono + "</tel>");
				sr.WriteLine("<sisSalNet>" + iSistemaSalario.ToString() + "</sisSalNet>");
				sr.WriteLine("<suelSal>" + dSueldo.ToString("0.00", us) + "</suelSal>");
				sr.WriteLine("<sobSuelComRemu>" + dSobreSueldo.ToString("0.00", us) + "</sobSuelComRemu>");
				sr.WriteLine("<decimTer>" + dDecimoXIII.ToString("0.00", us) + "</decimTer>");
				sr.WriteLine("<decimCuar>" + dDecimoXIV.ToString("0.00", us) + "</decimCuar>");
				sr.WriteLine("<fondoReserva>" + dFondoReserva.ToString("0.00", us) + "</fondoReserva>");
				sr.WriteLine("<salarioDigno>" + dSalarioDigno.ToString("0.00", us) + "</salarioDigno>");

				sr.WriteLine("<partUtil>" + dUtilidades.ToString("0.00", us) + "</partUtil>");
				sr.WriteLine("<desauOtras>" + dgDesahucio.ToString("0.00", us) + "</desauOtras>");
				sr.WriteLine("<apoPerIess>" + dAporteIess.ToString("0.00", us) + "</apoPerIess>");
				sr.WriteLine("<deducVivienda>" + dVivienda.ToString("0.00", us) + "</deducVivienda>");
				sr.WriteLine("<deducSalud>" + dSalud.ToString("0.00", us) + "</deducSalud>");
				sr.WriteLine("<deducEduca>" + dEducacion.ToString("0.00", us) + "</deducEduca>");
				sr.WriteLine("<deducAliement>" + dAlimentacion.ToString("0.00", us) + "</deducAliement>");
				sr.WriteLine("<deducVestim>" + dVestimenta.ToString("0.00", us) + "</deducVestim>");
				sr.WriteLine("<rebEspDiscap>" + dgDiscapacidad.ToString("0.00", us) + "</rebEspDiscap>");

				sr.WriteLine("<rebEspTerEd>" + dgRebajaIIIEdad.ToString("0.00", us) + "</rebEspTerEd>");
				sr.WriteLine("<impRentEmpl>" + ImpRentaAsumido.ToString("0.00", us) + "</impRentEmpl>");
				sr.WriteLine("<subTotal>" + dSubTotal.ToString("0.00", us) + "</subTotal>");
				sr.WriteLine("<numRet>" + iCuentaRet.ToString() + "</numRet>");
				sr.WriteLine("<numMesEmplead>" + iMeses.ToString("0.00", us) + "</numMesEmplead>");
				sr.WriteLine("<intGrabGen>" + IngGravOE.ToString("0.00", us) + "</intGrabGen>");
				sr.WriteLine("<deduccGastosOtrEmpl>" + DeduccionOE.ToString("0.00", us) + "</deduccGastosOtrEmpl>");
				sr.WriteLine("<otrRebjOtrEmpl>" + OtrasRebOE.ToString("0.00", us) + "</otrRebjOtrEmpl>");

				sr.WriteLine("<basImp>" + dBaseRetenido.ToString("0.00", us) + "</basImp>");
				sr.WriteLine("<impRentCaus>" + dImpRentaCaus.ToString("0.00", us) + "</impRentCaus>");
				sr.WriteLine("<valRet>" + dValRet.ToString("0.00", us) + "</valRet>");
				sr.WriteLine("<valorImpempAnter>" + dRentaRetOE.ToString("0.00", us) + "</valorImpempAnter>");
				sr.WriteLine("<anioRet>" + dtFecha1.ToString("yyyy") + "</anioRet>");
				sr.WriteLine("</datRetRelDep>");
				#endregion Grabacion de Linea
			}
			sr.WriteLine("</retRelDep>");
			sr.WriteLine("</rdep>");
			#endregion Lectura de Datos

			#region Cierre
			miReader.Close();
			sr.Flush();
			sr.Close();
			Cursor = Cursors.Default;
			MessageBox.Show("Archivo " + stNombreArchivo + " generado", "Información");
			#endregion Cierre
		}

		private void btArchivoEmpleado_Click(object sender, System.EventArgs e)
		{
			string stDir = @"C:\XML_ANEXO_TRANSACCIONAL\";
			string stNombreArchivo = "";
			try // Si no existe la carpeta C:\XML_ANEXO_TRANSACCIONAL\ se crea en la carpeta que inicio el sistema
			{
				if (!Directory.Exists(stDir)) stDir = MenuLatinium.stDirInicio + @"\";
				stNombreArchivo = @stDir + "Empleados.txt";
			}
			catch
			{
				MessageBox.Show("Esta abierto el archivo Empleados.txt");
				Cursor = Cursors.Default;
			}

			string stMensaje = "";
			int iAnio = ((DateTime) dtFecha.Value).Year;
			int idRolGrupo = 0;
			string stGrupo = "";
			if (cmbGrupo.Text.Length > 0)
			{
				idRolGrupo = (int) cmbGrupo.EditValue;
				stGrupo = "del grupo " + cmbGrupo.Text.Trim();
			}
			if (DialogResult.No == MessageBox.Show(String.Format("Desea Generar Listado de Pesonal del Año {0} {1}",
				iAnio, stGrupo), "Confirmación", MessageBoxButtons.YesNo, MessageBoxIcon.Question)) return;
			Cursor = Cursors.WaitCursor;
			StreamWriter sw = new StreamWriter(stNombreArchivo);

			foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in ultraGrid1.Rows.All)
			{
				string stTipo = dr.Cells["Tipo"].Value != DBNull.Value ? dr.Cells["Tipo"].Value.ToString() : "";
				string stCedula = dr.Cells["Cedula"].Value != DBNull.Value ? dr.Cells["Cedula"].Value.ToString() : "";
				string stNombre = dr.Cells["Nombre"].Value != DBNull.Value ? dr.Cells["Nombre"].Value.ToString().ToUpper() : "";
				string stTelefono = dr.Cells["Telefono"].Value != DBNull.Value ? dr.Cells["Telefono"].Value.ToString() : "";
				string stDireccion = dr.Cells["Direccion"].Value != DBNull.Value ? dr.Cells["Direccion"].Value.ToString() : "";
				string stNumero = dr.Cells["Numero"].Value != DBNull.Value ? dr.Cells["Numero"].Value.ToString() : "";
				string stProvincia = dr.Cells["idProvincia"].Value != DBNull.Value ? dr.Cells["idProvincia"].Value.ToString() : "";
				string stCiudad = dr.Cells["idCiudad"].Value != DBNull.Value ? dr.Cells["idCiudad"].Value.ToString() : "";

				if (stTipo == "") stMensaje += "Ingrese Tipo Como Cedula o Pasaporte";
				if (stCedula == "") stMensaje += "Ingrese Cedula\n";
				if (stTipo == "2" && stCedula.Length != 10) stMensaje += "Cedula debe tener 10 digitos";
				if (stNombre == "") stMensaje += "Ingrese Apellidos y Nombres";
				if (stTelefono.Length != 9) stMensaje += "Telefono debe tener 9 digitos";
				if (stDireccion == "") stMensaje += "Ingrese Direccion";
				if (stNumero == "") stMensaje += "Ingrese Numero de la calle";
				if (stProvincia == "") stMensaje += "Ingrese Provincia";
				if (stCiudad == "") stMensaje += "Ingrese Ciudad";

				if (stMensaje.Length > 0)
				{
					ultraGrid1.ActiveRow = dr;
					sw.Close();
					MessageBox.Show(stMensaje, "Error en: " + stNombre, 
						MessageBoxButtons.OK, MessageBoxIcon.Warning);
					return;
				}

				string stSelect = String.Format("Exec PersonalVerificaAnio {0}, '{1}', {2}", 
					iAnio, stCedula, idRolGrupo);
				if (Funcion.iEscalarSQL(cdsCiudad, stSelect) != 1) continue;

				if (stCedula.Length > 10) stCedula = stCedula.Substring(0, 10);
				if (stCedula.Length < 10) stCedula = stCedula + "0000000000".Substring(0, 10 - stCedula.Length);
				stNombre = stNombre.Replace('Á', 'A');
				stNombre = stNombre.Replace('É', 'E');
				stNombre = stNombre.Replace('Í', 'I');
				stNombre = stNombre.Replace('Ó', 'O');
				stNombre = stNombre.Replace('Ú', 'A');
				stNombre = stNombre.Replace('Ñ', 'n');

				if (stDireccion.Length > 20) stDireccion = stDireccion.Substring(0, 20);
				stSelect = "Select RTRIM(Codigo) From Provincia Where idProvincia = " + stProvincia;
				stProvincia = Funcion.sEscalarSQL(cdsCiudad, stSelect);
				stSelect = "Select RTRIM(Codigo) From ClienteCiudad Where idCiudad = " + stCiudad;
				stCiudad = Funcion.sEscalarSQL(cdsCiudad, stSelect);

				sw.WriteLine(String.Format("{0}\t{1}\t{2}\t{3}\t{4}\t{5}\t{6}\t{7}",
					stTipo, stCedula, stNombre, stTelefono, stDireccion, stNumero, stProvincia, stCiudad));
			}
			sw.Flush();
			sw.Close();
			Cursor = Cursors.Default;
			MessageBox.Show("Archivo " + stNombreArchivo + " generado", "Informacion", 
				MessageBoxButtons.OK, MessageBoxIcon.Information);
		}

		private void btDetalle_Click(object sender, System.EventArgs e)
		{
			int iFila = grvRdep.FocusedRowHandle;
			if (iFila < 0)
			{
				MessageBox.Show("Seleccione la línea que desa ver la Transacción", "Información");
				return;
			}
		
			int idPersonal = (int) grvRdep.GetRowCellValue(iFila, "idPersonal");
			if (idPersonal == 0)
			{
				MessageBox.Show("Fila no contiene Transacción Asociados",	"Información");
				return;
			}
			
			int iAnio = ((DateTime) dtFecha.Value).Year;
			string stExec = String.Format("Exec PersonalDatosDetalle {0}, {1}",
				iAnio, idPersonal);

			grvDetalle.Columns.Clear();
			grdDetalle.DataSource = Funcion.dvProcedimiento(cdsCiudad, stExec);
			ConfiguraGrilla("Detalle");
			ConfiguraGrilla("Detalle");
			xtraTabControl1.SelectedTabPage = xtraTabPage3;
		}

		private void btExcelRDep_Click(object sender, System.EventArgs e)
		{
			string stDir = MenuLatinium.stDirInicio + "\\Rdep.xls";
			grdRdep.ExportToXls(stDir, true);
			if (DialogResult.Yes == MessageBox.Show(
				"¿Desea Ver Archivo Generado?", 
				"Información", MessageBoxButtons.YesNo, MessageBoxIcon.Question))
				System.Diagnostics.Process.Start(stDir);		
		}

		private void btExcelDetalle_Click(object sender, System.EventArgs e)
		{
			string stDir = MenuLatinium.stDirInicio + "\\RdepDetalle.xls";
			grdDetalle.ExportToXls(stDir, true);
			if (DialogResult.Yes == MessageBox.Show(
				"¿Desea Ver Archivo Generado?", 
				"Información", MessageBoxButtons.YesNo, MessageBoxIcon.Question))
				System.Diagnostics.Process.Start(stDir);		
		}

		private void btPesonal_Click(object sender, System.EventArgs e)
		{
			int iFila = grvRdep.FocusedRowHandle;
			if (iFila < 0)
			{
				MessageBox.Show("Seleccione la línea que desa ver la Transacción", "Información");
				return;
			}
		
			int idPersonal = (int) grvRdep.GetRowCellValue(iFila, "idPersonal");
			if (idPersonal == 0)
			{
				MessageBox.Show("Fila no contiene Transacción Asociados",	"Información");
				return;
			}

			RolPersonal miPersona = new RolPersonal(idPersonal);
			miPersona.MdiParent = this.MdiParent;
			miPersona.Show();
		}

		private void btGeneraRDEP_MouseUp(object sender, System.Windows.Forms.MouseEventArgs e)
		{
			if (e.Button != MouseButtons.Right) return;
			GeneraGrilla();
		}

		private void grdRdep_DoubleClick(object sender, System.EventArgs e)
		{
			btDetalle_Click(this, e);
		}

		private void cmbGrupo_EditValueChanged(object sender, System.EventArgs e)
		{
		
		}

		private void xtraTabControl1_Click(object sender, System.EventArgs e)
		{
		
		}

		private void chkResumen_CheckedChanged(object sender, System.EventArgs e)
		{
		
		}
	}
}
