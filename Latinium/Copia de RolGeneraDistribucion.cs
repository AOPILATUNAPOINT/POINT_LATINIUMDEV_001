using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using System.Globalization;
using C1.C1Excel;
using System.Security.Cryptography;  
using System.Text;  




namespace Latinium
{
	/// <summary>
	/// Descripción breve de RolGeneraDistribucion.
	/// </summary>
	public class RolGeneraDistribucion : DevExpress.XtraEditors.XtraForm
	{
	//	private C1.Data.C1DataSet cdsRolDist;
		private C1.Data.C1DataSet cdsPersonal;
		private C1.Data.C1DataView cdvPersonalFecha;
		private C1.Data.C1DataSet cdsSubProyecto;
		private C1.Data.C1DataSet cdsProyecto;
		private C1.Data.C1DataView cdvProyecto;
		private Infragistics.Win.Misc.UltraButton ultraButton1;
		private System.Windows.Forms.ColorDialog colorDialog1;
		private Infragistics.Win.Misc.UltraButton btArchivo;
		private System.Windows.Forms.Button button1;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtExcel;
		private System.Windows.Forms.OpenFileDialog openFileDialog;
		private C1.C1Excel.C1XLBook c1XLBook;
		private System.Windows.Forms.TextBox txtidRolDis;
		private C1.Data.C1DataSet cdsRolConcepto;
		private Infragistics.Win.UltraWinGrid.ExcelExport.UltraGridExcelExporter ultraGridExcelExporter1;
		private Infragistics.Win.Misc.UltraButton btExcel;
		private Infragistics.Win.UltraWinSchedule.UltraCalendarCombo dtFecha;
		private BarraDatoSQL.BarraDatoSQL barraDato1;
		private C1.Data.C1DataSet cdsDistribucionG;
		private Infragistics.Win.UltraWinGrid.UltraDropDown cmbPersonal;
		private Infragistics.Win.UltraWinGrid.UltraDropDown cmbProyecto1;
		private Infragistics.Win.UltraWinGrid.UltraGrid ultraGrid1;
		/// <summary>
		/// Variable del diseñador requerida.
		/// </summary>
		private System.ComponentModel.Container components = null;

		public RolGeneraDistribucion()
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
			Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton dateButton1 = new Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton();
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
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand2 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Proyecto", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn92 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idProyecto");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn93 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idPadre");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn94 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn95 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Lugar");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn96 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Responsable");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn97 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Notas");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn98 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("CentroCosto");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn99 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("FechaIngreso");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn100 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Visible");
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand3 = new Infragistics.Win.UltraWinGrid.UltraGridBand("_RolDistribucion", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn101 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idRolDistribucion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn102 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idDistribucionG");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn103 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idPersonal", -1, "cmbPersonal");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn104 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idRolTipo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn105 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idProyecto", -1, "cmbProyecto1");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn106 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idSubproyecto");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn107 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Fecha");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn108 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Anio");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn109 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Mes");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn110 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("dias");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn111 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("porcentaje");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn112 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("posicion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn113 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("valor");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn114 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("referencia");
			this.cdsPersonal = new C1.Data.C1DataSet();
			this.cdvPersonalFecha = new C1.Data.C1DataView();
			this.cdsSubProyecto = new C1.Data.C1DataSet();
			this.cdsProyecto = new C1.Data.C1DataSet();
			this.cdvProyecto = new C1.Data.C1DataView();
			this.ultraButton1 = new Infragistics.Win.Misc.UltraButton();
			this.colorDialog1 = new System.Windows.Forms.ColorDialog();
			this.txtidRolDis = new System.Windows.Forms.TextBox();
			this.btArchivo = new Infragistics.Win.Misc.UltraButton();
			this.button1 = new System.Windows.Forms.Button();
			this.txtExcel = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.openFileDialog = new System.Windows.Forms.OpenFileDialog();
			this.c1XLBook = new C1.C1Excel.C1XLBook();
			this.cdsRolConcepto = new C1.Data.C1DataSet();
			this.ultraGridExcelExporter1 = new Infragistics.Win.UltraWinGrid.ExcelExport.UltraGridExcelExporter();
			this.btExcel = new Infragistics.Win.Misc.UltraButton();
			this.dtFecha = new Infragistics.Win.UltraWinSchedule.UltraCalendarCombo();
			this.barraDato1 = new BarraDatoSQL.BarraDatoSQL();
			this.cdsDistribucionG = new C1.Data.C1DataSet();
			this.cmbPersonal = new Infragistics.Win.UltraWinGrid.UltraDropDown();
			this.cmbProyecto1 = new Infragistics.Win.UltraWinGrid.UltraDropDown();
			this.ultraGrid1 = new Infragistics.Win.UltraWinGrid.UltraGrid();
			((System.ComponentModel.ISupportInitialize)(this.cdsPersonal)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdvPersonalFecha)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSubProyecto)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsProyecto)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdvProyecto)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtExcel)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsRolConcepto)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.dtFecha)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsDistribucionG)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbPersonal)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbProyecto1)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraGrid1)).BeginInit();
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
			// cdvPersonalFecha
			// 
			this.cdvPersonalFecha.BindingContextCtrl = this;
			this.cdvPersonalFecha.DataSet = this.cdsPersonal;
			this.cdvPersonalFecha.Sort = "Codigo";
			this.cdvPersonalFecha.TableName = "";
			this.cdvPersonalFecha.TableViewName = "Cuenta";
			// 
			// cdsSubProyecto
			// 
			this.cdsSubProyecto.BindingContextCtrl = this;
			this.cdsSubProyecto.DataLibrary = "LibContabilidad";
			this.cdsSubProyecto.DataLibraryUrl = "";
			this.cdsSubProyecto.DataSetDef = "dsProyecto";
			this.cdsSubProyecto.Locale = new System.Globalization.CultureInfo("es-EC");
			this.cdsSubProyecto.Name = "cdsSubProyecto";
			this.cdsSubProyecto.SchemaClassName = "LibContabilidad.DataClass";
			this.cdsSubProyecto.SchemaDef = null;
			// 
			// cdsProyecto
			// 
			this.cdsProyecto.BindingContextCtrl = this;
			this.cdsProyecto.DataLibrary = "LibContabilidad";
			this.cdsProyecto.DataLibraryUrl = "";
			this.cdsProyecto.DataSetDef = "dsProyecto";
			this.cdsProyecto.Locale = new System.Globalization.CultureInfo("es-EC");
			this.cdsProyecto.Name = "cdsProyecto";
			this.cdsProyecto.SchemaClassName = "LibContabilidad.DataClass";
			this.cdsProyecto.SchemaDef = null;
			this.cdsProyecto.BeforeFill += new C1.Data.FillEventHandler(this.cdsProyecto_BeforeFill);
			// 
			// cdvProyecto
			// 
			this.cdvProyecto.BindingContextCtrl = this;
			this.cdvProyecto.DataSet = this.cdsProyecto;
			this.cdvProyecto.RowFilter = "Visible =1";
			this.cdvProyecto.Sort = "Nombre";
			this.cdvProyecto.TableName = "";
			this.cdvProyecto.TableViewName = "Proyecto";
			// 
			// ultraButton1
			// 
			this.ultraButton1.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			this.ultraButton1.Location = new System.Drawing.Point(552, 8);
			this.ultraButton1.Name = "ultraButton1";
			this.ultraButton1.Size = new System.Drawing.Size(120, 24);
			this.ultraButton1.TabIndex = 120;
			this.ultraButton1.Text = "Genera Distribucion ";
			this.ultraButton1.Click += new System.EventHandler(this.ultraButton1_Click);
			// 
			// txtidRolDis
			// 
			this.txtidRolDis.DataBindings.Add(new System.Windows.Forms.Binding("Text", this.cdsDistribucionG, "_RolDistribucionG.idRolDistribucionG"));
			this.txtidRolDis.Location = new System.Drawing.Point(608, 248);
			this.txtidRolDis.Name = "txtidRolDis";
			this.txtidRolDis.Size = new System.Drawing.Size(40, 20);
			this.txtidRolDis.TabIndex = 126;
			this.txtidRolDis.Text = "";
			// 
			// btArchivo
			// 
			this.btArchivo.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			this.btArchivo.Location = new System.Drawing.Point(240, 8);
			this.btArchivo.Name = "btArchivo";
			this.btArchivo.Size = new System.Drawing.Size(24, 24);
			this.btArchivo.TabIndex = 130;
			this.btArchivo.Text = "...";
			this.btArchivo.Click += new System.EventHandler(this.btArchivo_Click);
			// 
			// button1
			// 
			this.button1.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			this.button1.Location = new System.Drawing.Point(272, 8);
			this.button1.Name = "button1";
			this.button1.Size = new System.Drawing.Size(88, 24);
			this.button1.TabIndex = 129;
			this.button1.Text = "Importar datos ";
			this.button1.Click += new System.EventHandler(this.button1_Click);
			// 
			// txtExcel
			// 
			this.txtExcel.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.txtExcel.Location = new System.Drawing.Point(16, 8);
			this.txtExcel.Name = "txtExcel";
			this.txtExcel.Size = new System.Drawing.Size(224, 21);
			this.txtExcel.TabIndex = 128;
			this.txtExcel.Text = "C:\\Roles\\Roles.xls";
			// 
			// cdsRolConcepto
			// 
			this.cdsRolConcepto.BindingContextCtrl = this;
			this.cdsRolConcepto.DataLibrary = "LibPersonal";
			this.cdsRolConcepto.DataLibraryUrl = "";
			this.cdsRolConcepto.DataSetDef = "dsRolGrupo";
			this.cdsRolConcepto.Locale = new System.Globalization.CultureInfo("es-EC");
			this.cdsRolConcepto.Name = "cdsRolConcepto";
			this.cdsRolConcepto.SchemaClassName = "LibPersonal.DataClass";
			this.cdsRolConcepto.SchemaDef = null;
			// 
			// btExcel
			// 
			this.btExcel.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			this.btExcel.Location = new System.Drawing.Point(360, 8);
			this.btExcel.Name = "btExcel";
			this.btExcel.Size = new System.Drawing.Size(64, 24);
			this.btExcel.TabIndex = 132;
			this.btExcel.Text = "Excel";
			this.btExcel.Click += new System.EventHandler(this.btExcel_Click);
			// 
			// dtFecha
			// 
			appearance1.ForeColorDisabled = System.Drawing.Color.Black;
			this.dtFecha.Appearance = appearance1;
			this.dtFecha.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			dateButton1.Caption = "Today";
			this.dtFecha.DateButtons.Add(dateButton1);
			this.dtFecha.Format = "MMM/yyyy";
			this.dtFecha.Location = new System.Drawing.Point(424, 8);
			this.dtFecha.Name = "dtFecha";
			this.dtFecha.NonAutoSizeHeight = 23;
			this.dtFecha.Size = new System.Drawing.Size(112, 21);
			this.dtFecha.SpinButtonsVisible = true;
			this.dtFecha.TabIndex = 133;
			this.dtFecha.Value = new System.DateTime(2005, 6, 1, 0, 0, 0, 0);
			// 
			// barraDato1
			// 
			this.barraDato1.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.barraDato1.BarraMovimiento = true;
			this.barraDato1.DataMember = "RolDistribucionG";
			this.barraDato1.DataNombreId = "idRolDistribucionG";
			this.barraDato1.DataOrden = "";
			this.barraDato1.DataSource = this.cdsDistribucionG;
			this.barraDato1.DataTabla = "RolDistribucionG";
			this.barraDato1.DataTablaHija = "";
			this.barraDato1.Location = new System.Drawing.Point(16, 248);
			this.barraDato1.Name = "barraDato1";
			this.barraDato1.Size = new System.Drawing.Size(304, 24);
			this.barraDato1.TabIndex = 127;
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
			this.barraDato1.Editar += new System.EventHandler(this.barraDato1_Editar);
			this.barraDato1.Grabar += new System.EventHandler(this.barraDato1_Grabar);
			// 
			// cdsDistribucionG
			// 
			this.cdsDistribucionG.BindingContextCtrl = this;
			this.cdsDistribucionG.DataLibrary = "LibPersonal";
			this.cdsDistribucionG.DataLibraryUrl = "";
			this.cdsDistribucionG.DataSetDef = "dsRolDistribucion";
			this.cdsDistribucionG.EnforceConstraints = false;
			this.cdsDistribucionG.Locale = new System.Globalization.CultureInfo("es-EC");
			this.cdsDistribucionG.Name = "cdsDistribucionG";
			this.cdsDistribucionG.SchemaClassName = "LibPersonal.DataClass";
			this.cdsDistribucionG.SchemaDef = null;
			this.cdsDistribucionG.BeforeFill += new C1.Data.FillEventHandler(this.cdsDistribucionG_BeforeFill_1);
			// 
			// cmbPersonal
			// 
			this.cmbPersonal.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbPersonal.DataMember = "Personal";
			this.cmbPersonal.DataSource = this.cdsPersonal;
			ultraGridBand1.AddButtonCaption = "Proyecto";
			ultraGridColumn1.Header.VisiblePosition = 0;
			ultraGridColumn1.Hidden = true;
			ultraGridColumn2.Header.VisiblePosition = 2;
			ultraGridColumn2.Hidden = true;
			ultraGridColumn3.Header.VisiblePosition = 3;
			ultraGridColumn3.Hidden = true;
			ultraGridColumn4.Header.VisiblePosition = 5;
			ultraGridColumn4.Hidden = true;
			ultraGridColumn5.Header.VisiblePosition = 7;
			ultraGridColumn5.Hidden = true;
			ultraGridColumn6.Header.VisiblePosition = 8;
			ultraGridColumn6.Hidden = true;
			ultraGridColumn7.Header.VisiblePosition = 10;
			ultraGridColumn7.Hidden = true;
			ultraGridColumn8.Header.VisiblePosition = 11;
			ultraGridColumn8.Hidden = true;
			ultraGridColumn9.Header.VisiblePosition = 1;
			ultraGridColumn9.Width = 234;
			ultraGridColumn10.Header.VisiblePosition = 12;
			ultraGridColumn10.Hidden = true;
			ultraGridColumn11.Header.VisiblePosition = 13;
			ultraGridColumn11.Width = 73;
			ultraGridColumn12.Header.VisiblePosition = 14;
			ultraGridColumn12.Hidden = true;
			ultraGridColumn13.Header.VisiblePosition = 15;
			ultraGridColumn13.Hidden = true;
			ultraGridColumn14.Header.VisiblePosition = 16;
			ultraGridColumn14.Hidden = true;
			ultraGridColumn15.Header.VisiblePosition = 17;
			ultraGridColumn15.Hidden = true;
			ultraGridColumn16.Header.VisiblePosition = 18;
			ultraGridColumn16.Hidden = true;
			ultraGridColumn17.Header.VisiblePosition = 19;
			ultraGridColumn17.Hidden = true;
			ultraGridColumn18.Header.VisiblePosition = 20;
			ultraGridColumn18.Hidden = true;
			ultraGridColumn19.Header.VisiblePosition = 21;
			ultraGridColumn19.Hidden = true;
			ultraGridColumn20.Header.VisiblePosition = 22;
			ultraGridColumn20.Hidden = true;
			ultraGridColumn21.Header.VisiblePosition = 23;
			ultraGridColumn21.Hidden = true;
			ultraGridColumn22.Header.VisiblePosition = 24;
			ultraGridColumn22.Hidden = true;
			ultraGridColumn23.Header.VisiblePosition = 25;
			ultraGridColumn23.Hidden = true;
			ultraGridColumn24.Header.VisiblePosition = 26;
			ultraGridColumn24.Hidden = true;
			ultraGridColumn25.Header.VisiblePosition = 27;
			ultraGridColumn25.Hidden = true;
			ultraGridColumn26.Header.VisiblePosition = 28;
			ultraGridColumn26.Hidden = true;
			ultraGridColumn27.Header.VisiblePosition = 9;
			ultraGridColumn27.Hidden = true;
			ultraGridColumn28.Header.VisiblePosition = 29;
			ultraGridColumn28.Hidden = true;
			ultraGridColumn29.Header.VisiblePosition = 30;
			ultraGridColumn29.Hidden = true;
			ultraGridColumn30.Header.VisiblePosition = 31;
			ultraGridColumn30.Hidden = true;
			ultraGridColumn31.Header.VisiblePosition = 32;
			ultraGridColumn31.Hidden = true;
			ultraGridColumn32.Header.VisiblePosition = 33;
			ultraGridColumn32.Hidden = true;
			ultraGridColumn33.Header.VisiblePosition = 34;
			ultraGridColumn33.Hidden = true;
			ultraGridColumn34.Header.VisiblePosition = 35;
			ultraGridColumn34.Hidden = true;
			ultraGridColumn35.Header.VisiblePosition = 36;
			ultraGridColumn35.Hidden = true;
			ultraGridColumn36.Header.VisiblePosition = 37;
			ultraGridColumn36.Hidden = true;
			ultraGridColumn37.Header.VisiblePosition = 38;
			ultraGridColumn37.Hidden = true;
			ultraGridColumn38.Header.VisiblePosition = 39;
			ultraGridColumn38.Hidden = true;
			ultraGridColumn39.Header.VisiblePosition = 40;
			ultraGridColumn39.Hidden = true;
			ultraGridColumn40.Header.VisiblePosition = 41;
			ultraGridColumn40.Hidden = true;
			ultraGridColumn41.Header.VisiblePosition = 42;
			ultraGridColumn41.Hidden = true;
			ultraGridColumn42.Header.VisiblePosition = 43;
			ultraGridColumn42.Hidden = true;
			ultraGridColumn43.Header.VisiblePosition = 44;
			ultraGridColumn43.Hidden = true;
			ultraGridColumn44.Header.VisiblePosition = 45;
			ultraGridColumn44.Hidden = true;
			ultraGridColumn45.Header.VisiblePosition = 46;
			ultraGridColumn45.Hidden = true;
			ultraGridColumn46.Header.VisiblePosition = 47;
			ultraGridColumn46.Hidden = true;
			ultraGridColumn47.Header.VisiblePosition = 48;
			ultraGridColumn47.Hidden = true;
			ultraGridColumn48.Header.VisiblePosition = 49;
			ultraGridColumn48.Hidden = true;
			ultraGridColumn49.Header.VisiblePosition = 50;
			ultraGridColumn49.Hidden = true;
			ultraGridColumn50.Header.VisiblePosition = 51;
			ultraGridColumn50.Hidden = true;
			ultraGridColumn51.Header.VisiblePosition = 52;
			ultraGridColumn51.Hidden = true;
			ultraGridColumn52.Header.VisiblePosition = 53;
			ultraGridColumn52.Hidden = true;
			ultraGridColumn53.Header.VisiblePosition = 54;
			ultraGridColumn53.Hidden = true;
			ultraGridColumn54.Header.VisiblePosition = 4;
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
			ultraGridColumn63.Header.VisiblePosition = 6;
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
																										 ultraGridColumn91});
			this.cmbPersonal.DisplayLayout.BandsSerializer.Add(ultraGridBand1);
			this.cmbPersonal.DisplayMember = "Nombre";
			this.cmbPersonal.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbPersonal.Location = new System.Drawing.Point(24, 96);
			this.cmbPersonal.Name = "cmbPersonal";
			this.cmbPersonal.Size = new System.Drawing.Size(336, 80);
			this.cmbPersonal.TabIndex = 125;
			this.cmbPersonal.ValueMember = "idPersonal";
			this.cmbPersonal.Visible = false;
			this.cmbPersonal.InitializeLayout += new Infragistics.Win.UltraWinGrid.InitializeLayoutEventHandler(this.cmbPersonal_InitializeLayout);
			// 
			// cmbProyecto1
			// 
			this.cmbProyecto1.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbProyecto1.DataSource = this.cdvProyecto;
			ultraGridColumn92.Header.VisiblePosition = 2;
			ultraGridColumn92.Hidden = true;
			ultraGridColumn93.Header.VisiblePosition = 6;
			ultraGridColumn93.Hidden = true;
			ultraGridColumn94.Header.VisiblePosition = 0;
			ultraGridColumn94.Width = 144;
			ultraGridColumn95.Header.VisiblePosition = 3;
			ultraGridColumn95.Hidden = true;
			ultraGridColumn96.Header.VisiblePosition = 7;
			ultraGridColumn96.Hidden = true;
			ultraGridColumn97.Header.VisiblePosition = 4;
			ultraGridColumn97.Hidden = true;
			ultraGridColumn98.Header.VisiblePosition = 1;
			ultraGridColumn99.Header.VisiblePosition = 5;
			ultraGridColumn99.Hidden = true;
			ultraGridColumn100.Header.VisiblePosition = 8;
			ultraGridColumn100.Hidden = true;
			ultraGridBand2.Columns.AddRange(new object[] {
																										 ultraGridColumn92,
																										 ultraGridColumn93,
																										 ultraGridColumn94,
																										 ultraGridColumn95,
																										 ultraGridColumn96,
																										 ultraGridColumn97,
																										 ultraGridColumn98,
																										 ultraGridColumn99,
																										 ultraGridColumn100});
			this.cmbProyecto1.DisplayLayout.BandsSerializer.Add(ultraGridBand2);
			this.cmbProyecto1.DisplayMember = "Nombre";
			this.cmbProyecto1.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbProyecto1.Location = new System.Drawing.Point(384, 96);
			this.cmbProyecto1.Name = "cmbProyecto1";
			this.cmbProyecto1.Size = new System.Drawing.Size(264, 80);
			this.cmbProyecto1.TabIndex = 123;
			this.cmbProyecto1.ValueMember = "idProyecto";
			this.cmbProyecto1.Visible = false;
			this.cmbProyecto1.BeforeDropDown += new System.ComponentModel.CancelEventHandler(this.cmbProyecto1_BeforeDropDown);
			this.cmbProyecto1.InitializeLayout += new Infragistics.Win.UltraWinGrid.InitializeLayoutEventHandler(this.cmbProyecto1_InitializeLayout);
			// 
			// ultraGrid1
			// 
			this.ultraGrid1.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.ultraGrid1.Cursor = System.Windows.Forms.Cursors.Default;
			this.ultraGrid1.DataMember = "_RolDistribucionG.RolDistribucionG - RolDistribucion";
			this.ultraGrid1.DataSource = this.cdsDistribucionG;
			ultraGridColumn101.Header.VisiblePosition = 0;
			ultraGridColumn101.Hidden = true;
			ultraGridColumn102.Header.VisiblePosition = 1;
			ultraGridColumn102.Hidden = true;
			ultraGridColumn103.Header.Caption = "Personal";
			ultraGridColumn103.Header.VisiblePosition = 2;
			ultraGridColumn103.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.DropDownList;
			ultraGridColumn104.Header.VisiblePosition = 3;
			ultraGridColumn104.Hidden = true;
			ultraGridColumn105.Header.Caption = "Proyecto";
			ultraGridColumn105.Header.VisiblePosition = 4;
			ultraGridColumn105.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.DropDownList;
			ultraGridColumn106.Header.VisiblePosition = 9;
			ultraGridColumn106.Hidden = true;
			ultraGridColumn107.Header.VisiblePosition = 5;
			ultraGridColumn107.Hidden = true;
			ultraGridColumn108.Header.VisiblePosition = 6;
			ultraGridColumn109.Header.VisiblePosition = 7;
			ultraGridColumn110.Header.VisiblePosition = 8;
			ultraGridColumn111.Header.VisiblePosition = 10;
			ultraGridColumn111.Hidden = true;
			ultraGridColumn112.Header.VisiblePosition = 11;
			ultraGridColumn113.Header.VisiblePosition = 12;
			ultraGridColumn114.Header.VisiblePosition = 13;
			ultraGridBand3.Columns.AddRange(new object[] {
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
																										 ultraGridColumn114});
			this.ultraGrid1.DisplayLayout.BandsSerializer.Add(ultraGridBand3);
			this.ultraGrid1.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.ultraGrid1.Location = new System.Drawing.Point(16, 48);
			this.ultraGrid1.Name = "ultraGrid1";
			this.ultraGrid1.Size = new System.Drawing.Size(656, 184);
			this.ultraGrid1.TabIndex = 122;
			this.ultraGrid1.InitializeLayout += new Infragistics.Win.UltraWinGrid.InitializeLayoutEventHandler(this.ultraGrid1_InitializeLayout);
			// 
			// RolGeneraDistribucion
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 13);
			this.ClientSize = new System.Drawing.Size(688, 277);
			this.Controls.Add(this.dtFecha);
			this.Controls.Add(this.btExcel);
			this.Controls.Add(this.btArchivo);
			this.Controls.Add(this.button1);
			this.Controls.Add(this.txtExcel);
			this.Controls.Add(this.barraDato1);
			this.Controls.Add(this.txtidRolDis);
			this.Controls.Add(this.cmbPersonal);
			this.Controls.Add(this.cmbProyecto1);
			this.Controls.Add(this.ultraGrid1);
			this.Controls.Add(this.ultraButton1);
			this.Name = "RolGeneraDistribucion";
			this.Text = "RolGeneraDistribucion";
			this.Load += new System.EventHandler(this.RolGeneraDistribucion_Load);
			((System.ComponentModel.ISupportInitialize)(this.cdsPersonal)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdvPersonalFecha)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSubProyecto)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsProyecto)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdvProyecto)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtExcel)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsRolConcepto)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.dtFecha)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsDistribucionG)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbPersonal)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbProyecto1)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraGrid1)).EndInit();
			this.ResumeLayout(false);

		}
		#endregion

		private void cmbCentroCosto_InitializeLayout(object sender, Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs e)
		{
		
		}

		private void ultraDropDown2_InitializeLayout(object sender, Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs e)
		{
		
		}

		private void ultraButton1_Click(object sender, System.EventArgs e)
		{

			DateTime dtHasta = (DateTime) dtFecha.Value;
			if (MessageBox.Show("¿Desea Subir Valores de la Grilla al rol del mes de "
				+dtHasta.ToString("MM")+ " y anio "+ dtHasta.ToString("yyyy"),
				"Confirmación", MessageBoxButtons.YesNo, 
				MessageBoxIcon.Question) == DialogResult.No) return;
			if(dtFecha == null)
			{
				MessageBox.Show("Verifique Fecha","Información");
				return;
			}
			string stExec = "Exec RolGeneraDistribucion "+ txtidRolDis.Text+",'"
				+ dtHasta.ToString("yyyyMMdd") + "'";
			MessageBox.Show("Distribucion Generada", "Información");
     Funcion.EjecutaSQL(cdsDistribucionG,stExec);
		}

		private void cmbProyecto1_InitializeLayout(object sender, Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs e)
		{
		
		}

		private void ultraButton2_Click(object sender, System.EventArgs e)
		{
			try
			{
				cdsDistribucionG.Update();
			}
			catch (System.Exception ex)
			{
				MessageBox.Show(ex.Message, "Grabar Distribuciones");
				return;
			}
			this.Close();

		}

		private void barraDato1_Crear(object sender, System.EventArgs e)
		{
			this.barraDato1.CrearRegistro();
		
		}

		private void barraDato1_Grabar(object sender, System.EventArgs e)
		{
			this.barraDato1.GrabaRegistro();
		
		}

		private void barraDato1_Refresca(object sender, System.EventArgs e)
		{
	
		}
//int idRolD = 0;
		private void RolGeneraDistribucion_Load(object sender, System.EventArgs e)
		{
			this.barraDato1.ProximoId(4);
		}

		private void barraDato1_Editar(object sender, System.EventArgs e)
		{
			this.barraDato1.EditarRegistro();
		}

		private void ultraGrid1_InitializeLayout(object sender, Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs e)
		{
		
		}

		private void btArchivo_Click(object sender, System.EventArgs e)
		{
			this.openFileDialog.Filter = "xls (*.xls)|*.xls";
			this.openFileDialog.InitialDirectory = @"c:\Roles\";
			if (DialogResult.OK == this.openFileDialog.ShowDialog())
			{
				string stArchivo = this.openFileDialog.FileName.ToString();
				txtExcel.Text = stArchivo;
			}
		}

		private void button1_Click(object sender, System.EventArgs e)
		{
			
			if (MessageBox.Show("¿Desea Subir Valores de Roles de Excel?",
				"Confirmación", MessageBoxButtons.YesNo, 
				MessageBoxIcon.Question) == DialogResult.No) return;

			C1XLBook book1 = new C1XLBook();
			try
			{
				book1.Load(txtExcel.Text);
			}
			catch (Exception exception1)
			{
				MessageBox.Show(exception1.Message);
				return;
			}
			XLSheet sheet1 = book1.Sheets[0];
			
			int iFila = 1;
			string stCedula = "";
			string stNombre = "";
			double iDias = 0;
			double dValor = 0.00;
			string codProyecto = "";
//			string iRoldistri = txtidRolDis.Text;
			while (iFila < sheet1.Rows.Count && sheet1.GetCell(iFila, 1)!= null)
//																						 && 
//					sheet1.GetCell(iFila, 1).Value != null)
			{
			#region Subida_Excel

			stCedula = sheet1.GetCell(iFila, 0).Value.ToString();;
			stNombre = sheet1.GetCell(iFila, 1).Value.ToString();;
			iDias = (double)sheet1.GetCell(iFila, 2).Value;;
			dValor = (double)sheet1.GetCell(iFila, 3).Value;;
			codProyecto = sheet1.GetCell(iFila, 4).Value.ToString();;
			stCedula = stCedula.Replace("-", "");
			stNombre = stNombre.Replace("#", "Nro.");
			stNombre = stNombre.Replace("'", "''");
			string siExec = string.Format("exec RolDistribSubida '{0}','{1}',{2},{3},'{4}',{5}",
			stCedula	,	stNombre, iDias, dValor, codProyecto,int.Parse(txtidRolDis.Text));
			Funcion.EjecutaSQL(cdsDistribucionG,siExec.ToString(),true);
				//					int idCompra = Funcion.iEscalarSQL(cdsCliente,iidCompra.ToString(),true);
			iFila += 1;
			#endregion Subida_Excel
			}
		MessageBox.Show("Archivo Subido","Información");
		}

		private void barraDato1_Restaurar(object sender, System.EventArgs e)
		{
		
		}

		private void barraDato1_Borrar(object sender, System.EventArgs e)
		{
			if (DialogResult.No == MessageBox.Show(
				"¿Está seguro que desea Anular el Registro?", 
				"Confirmación", MessageBoxButtons.YesNo, MessageBoxIcon.Warning, 
				MessageBoxDefaultButton.Button2)) return;

			this.barraDato1.BorraRegistro();
		}

		private void barraDato1_Buscar(object sender, System.EventArgs e)
		{
			
			

		}

		private void barraDato1_Deshacer(object sender, System.EventArgs e)
		{
			this.barraDato1.DeshacerRegistro();
		}

		private void barraDato1_Imprimir(object sender, System.EventArgs e)
		{
		
		}

		private void btExcel_Click(object sender, System.EventArgs e)
		{
			string stDir = MenuLatinium.stDirInicio + "\\"; 
			ultraGridExcelExporter1.Export(ultraGrid1, "RolDistribucion.xls");
			if (DialogResult.Yes == MessageBox.Show(
				"¿Desea Ver Archivo Generado?", 
				"Información", MessageBoxButtons.YesNo, MessageBoxIcon.Question))
				System.Diagnostics.Process.Start(stDir + "RolDistribucionCC.xls");
		}

		private void cmbPersonal_InitializeLayout(object sender, Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs e)
		{
		
		}

		private void cmbProyecto1_BeforeDropDown(object sender, System.ComponentModel.CancelEventArgs e)
		{
		
		}

		private void cdsPersonal_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			this.cdsPersonal.Schema.Connections[0].ConnectionString = 
				miFuncion.AccesoBase(MenuLatinium.stDirRoles);
		}

		private void cdsDistribucion_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			this.cdsDistribucionG.Schema.Connections[0].ConnectionString = 
				miFuncion.AccesoBase(MenuLatinium.stDirRoles);
		}

		private void barraDato1_Load(object sender, System.EventArgs e)
		{
		
		}

		private void cdsProyecto_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			this.cdsProyecto.Schema.Connections[0].ConnectionString = 
				miFuncion.AccesoBase(MenuLatinium.stDirRoles);		
		}

		private void cdsDistribucionG_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			this.cdsDistribucionG.Schema.Connections[0].ConnectionString = 
				miFuncion.AccesoBase(MenuLatinium.stDirRoles);			
		}

		private void cdsDistribucionG_BeforeFill_1(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			this.cdsDistribucionG.Schema.Connections[0].ConnectionString = 
				miFuncion.AccesoBase(MenuLatinium.stDirRoles);	
		}
	}
}
