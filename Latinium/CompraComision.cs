using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using Infragistics.Win.UltraWinGrid;

namespace Latinium
{
	/// <summary>
	/// Descripción breve de CompraComision.
	/// </summary>
	public class CompraComision : DevExpress.XtraEditors.XtraForm
	{
		private C1.Data.C1DataSet cdsComision;
		private BarraDatoSQL.BarraDatoSQL barraDatoSQL1;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtIdCompra;
		private C1.Data.C1DataSet cdsCompra;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor ultraTextEditor1;
		private Infragistics.Win.Misc.UltraLabel ultraLabel1;
		private Infragistics.Win.UltraWinSchedule.UltraCalendarCombo ultraCalendarCombo1;
		private Infragistics.Win.Misc.UltraLabel ultraLabel2;
		private Infragistics.Win.UltraWinEditors.UltraCheckEditor chkVerifica;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource udsOrigen;
		private Infragistics.Win.UltraWinGrid.UltraDropDown ultraDropDown1;
		private Infragistics.Win.UltraWinGrid.UltraDropDown ultraDropDown2;
		private C1.Data.C1DataSet cdsOriginador;
		private Infragistics.Win.UltraWinGrid.UltraDropDown ultraDropDown3;
		private C1.Data.C1DataSet cdsPersonal;
		private Infragistics.Win.Misc.UltraButton btCliente;
		private Infragistics.Win.UltraWinGrid.UltraGrid grAsiento;
		private System.Windows.Forms.Label label1;
		private Infragistics.Win.Misc.UltraButton btCalcular;
		/// <summary>
		/// Variable del diseñador requerida.
		/// </summary>
		private System.ComponentModel.Container components = null;

		public CompraComision()
		{
			InitializeComponent();
		}

		int IdCompra = 0;
		public CompraComision(int idCompra)
		{
			InitializeComponent();
			IdCompra=idCompra;
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
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn1 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Trabajo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn2 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idOrigen");
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand1 = new Infragistics.Win.UltraWinGrid.UltraGridBand("CompraComisionOrig", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn1 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idOriginador");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn2 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Descripcion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn3 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("ValorBase");
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand2 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Personal", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn4 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idPersonal");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn5 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idAsiDistGrupo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn6 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCentroCosto");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn7 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idDepartamento");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn8 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCargo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn9 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idRolSucursal");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn10 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idRolGrupo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn11 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Jornal");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn12 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn13 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Apellido");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn14 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Codigo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn15 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Jornada");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn16 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Direccion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn17 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Telefono");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn18 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Celular");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn19 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Cedula");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn20 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Afiliacion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn21 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Reserva");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn22 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Renta");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn23 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Iess");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn24 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Sueldo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn25 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Comision");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn26 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Tipo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn27 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("FechaIngreso");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn28 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("FechaNacimiento");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn29 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("FechaSalida");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn30 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Notas");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn31 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("CargasFamiliares");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn32 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idProvincia");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn33 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCiudad");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn34 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Numero");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn35 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("SistemaSalario");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn36 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Cuenta");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn37 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("CodigoSectorial");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn38 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Condicion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn39 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("FechaNovedad");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn40 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Acceso");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn41 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Vendedor");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn42 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Borrar");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn43 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Cobrador");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn44 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Sexo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn45 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Estado_Civil");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn46 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Banco");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn47 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("CuentaCorr");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn48 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idTipoCuenta");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn49 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("CedulaMilitar");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn50 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idFormaPago");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn51 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Quincena");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn52 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("CodIngreso");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn53 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("CodTelefono");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn54 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("TarjEdificio");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn55 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("TarjetaOficina");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn56 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("email");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn57 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idProyecto");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn58 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCuenta");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn59 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Bono");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn60 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Bono1");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn61 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Desc1");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn62 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Desc2");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn63 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("TieneQuincena");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn64 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Fotos");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn65 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Convenio");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn66 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idSubProyecto");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn67 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Sindicato");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn68 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Asociacion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn69 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idRolHorario");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn70 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("FechaContrato");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn71 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Educacion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn72 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("NumeroDias");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn73 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idRelacionT");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn74 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("gSalud");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn75 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("gEducacion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn76 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("gVivienda");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn77 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("gVestimenta");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn78 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("gAlimentacion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn79 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("gTotal");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn80 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("BaseComision");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn81 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("PorcentajeComision");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn82 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("gDesahucio");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn83 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("gDiscapacidad");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn84 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("IngGravOE");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn85 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("DeduccionOE");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn86 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("OtrasRebOE");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn87 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("ImpRentaAsumido");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn88 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("gRebajaIIIEdad");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn89 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("RentaRetenidaOE");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn90 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("RentaRetOE");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn91 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("TipoSangre");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn92 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("RegEntrada");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn93 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("RegSalida");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn94 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idTipoContrato");
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand3 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn95 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Trabajo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn96 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idOrigen");
			Infragistics.Win.Appearance appearance1 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand4 = new Infragistics.Win.UltraWinGrid.UltraGridBand("CompraComision", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn97 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCompraComision");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn98 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCompra");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn99 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn100 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Porcentaje");
			Infragistics.Win.Appearance appearance2 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn101 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Valor");
			Infragistics.Win.Appearance appearance3 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn102 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Total");
			Infragistics.Win.Appearance appearance4 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn103 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idOriginador", -1, "ultraDropDown3");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn104 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idVendedor", -1, "ultraDropDown2");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn105 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idOriginadorT", -1, "ultraDropDown1");
			Infragistics.Win.UltraWinGrid.SummarySettings summarySettings1 = new Infragistics.Win.UltraWinGrid.SummarySettings("", Infragistics.Win.UltraWinGrid.SummaryType.Sum, null, "Total", 5, true, "CompraComision", 0, Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn, null, -1, false);
			Infragistics.Win.Appearance appearance5 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance6 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance7 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance8 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance9 = new Infragistics.Win.Appearance();
			this.cdsComision = new C1.Data.C1DataSet();
			this.barraDatoSQL1 = new BarraDatoSQL.BarraDatoSQL();
			this.cdsCompra = new C1.Data.C1DataSet();
			this.txtIdCompra = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.ultraTextEditor1 = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.ultraLabel1 = new Infragistics.Win.Misc.UltraLabel();
			this.ultraCalendarCombo1 = new Infragistics.Win.UltraWinSchedule.UltraCalendarCombo();
			this.ultraLabel2 = new Infragistics.Win.Misc.UltraLabel();
			this.chkVerifica = new Infragistics.Win.UltraWinEditors.UltraCheckEditor();
			this.udsOrigen = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.ultraDropDown1 = new Infragistics.Win.UltraWinGrid.UltraDropDown();
			this.ultraDropDown2 = new Infragistics.Win.UltraWinGrid.UltraDropDown();
			this.cdsPersonal = new C1.Data.C1DataSet();
			this.cdsOriginador = new C1.Data.C1DataSet();
			this.ultraDropDown3 = new Infragistics.Win.UltraWinGrid.UltraDropDown();
			this.btCliente = new Infragistics.Win.Misc.UltraButton();
			this.grAsiento = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.label1 = new System.Windows.Forms.Label();
			this.btCalcular = new Infragistics.Win.Misc.UltraButton();
			((System.ComponentModel.ISupportInitialize)(this.cdsComision)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsCompra)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIdCompra)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraTextEditor1)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraCalendarCombo1)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.udsOrigen)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDropDown1)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDropDown2)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsPersonal)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsOriginador)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDropDown3)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.grAsiento)).BeginInit();
			this.SuspendLayout();
			// 
			// cdsComision
			// 
			this.cdsComision.BindingContextCtrl = this;
			this.cdsComision.DataLibrary = "LibFacturacion";
			this.cdsComision.DataLibraryUrl = "";
			this.cdsComision.DataSetDef = "dsComision";
			this.cdsComision.EnforceConstraints = false;
			this.cdsComision.FillOnRequest = false;
			this.cdsComision.Locale = new System.Globalization.CultureInfo("es-EC");
			this.cdsComision.Name = "cdsComision";
			this.cdsComision.SchemaClassName = "LibFacturacion.DataClass";
			this.cdsComision.SchemaDef = null;
			// 
			// barraDatoSQL1
			// 
			this.barraDatoSQL1.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.barraDatoSQL1.BarraMovimiento = false;
			this.barraDatoSQL1.DataMember = "Compra";
			this.barraDatoSQL1.DataNombreId = "idCompra";
			this.barraDatoSQL1.DataOrden = "";
			this.barraDatoSQL1.DataSource = this.cdsCompra;
			this.barraDatoSQL1.DataTabla = "Compra";
			this.barraDatoSQL1.DataTablaHija = "";
			this.barraDatoSQL1.Location = new System.Drawing.Point(16, 223);
			this.barraDatoSQL1.Name = "barraDatoSQL1";
			this.barraDatoSQL1.Size = new System.Drawing.Size(96, 24);
			this.barraDatoSQL1.TabIndex = 1;
			this.barraDatoSQL1.VisibleBorrar = false;
			this.barraDatoSQL1.VisibleBuscar = false;
			this.barraDatoSQL1.VisibleEditar = true;
			this.barraDatoSQL1.VisibleImprimir = true;
			this.barraDatoSQL1.VisibleNuevo = true;
			this.barraDatoSQL1.Buscar += new System.EventHandler(this.barraDatoSQL1_Buscar);
			this.barraDatoSQL1.Borrar += new System.EventHandler(this.barraDatoSQL1_Borrar);
			this.barraDatoSQL1.Imprimir += new System.EventHandler(this.barraDatoSQL1_Imprimir);
			this.barraDatoSQL1.Restaurar += new System.EventHandler(this.barraDatoSQL1_Restaurar);
			this.barraDatoSQL1.Refresca += new System.EventHandler(this.barraDatoSQL1_Refresca);
			this.barraDatoSQL1.Crear += new System.EventHandler(this.barraDatoSQL1_Crear);
			this.barraDatoSQL1.Deshacer += new System.EventHandler(this.barraDatoSQL1_Deshacer);
			this.barraDatoSQL1.Copiar += new System.EventHandler(this.barraDatoSQL1_Copiar);
			this.barraDatoSQL1.Impresora += new System.EventHandler(this.barraDatoSQL1_Impresora);
			this.barraDatoSQL1.Editar += new System.EventHandler(this.barraDatoSQL1_Editar);
			this.barraDatoSQL1.Grabar += new System.EventHandler(this.barraDatoSQL1_Grabar);
			// 
			// cdsCompra
			// 
			this.cdsCompra.BindingContextCtrl = this;
			this.cdsCompra.DataLibrary = "LibFacturacion";
			this.cdsCompra.DataLibraryUrl = "";
			this.cdsCompra.DataSetDef = "dsCompra";
			this.cdsCompra.EnforceConstraints = false;
			this.cdsCompra.FillOnRequest = false;
			this.cdsCompra.Locale = new System.Globalization.CultureInfo("es-EC");
			this.cdsCompra.Name = "cdsCompra";
			this.cdsCompra.SchemaClassName = "LibFacturacion.DataClass";
			this.cdsCompra.SchemaDef = null;
			// 
			// txtIdCompra
			// 
			this.txtIdCompra.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsCompra, "Compra.idCompra"));
			this.txtIdCompra.Location = new System.Drawing.Point(352, 8);
			this.txtIdCompra.Name = "txtIdCompra";
			this.txtIdCompra.PromptChar = ' ';
			this.txtIdCompra.Size = new System.Drawing.Size(56, 21);
			this.txtIdCompra.TabIndex = 2;
			// 
			// ultraTextEditor1
			// 
			this.ultraTextEditor1.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsCompra, "Compra.Numero"));
			this.ultraTextEditor1.Location = new System.Drawing.Point(72, 8);
			this.ultraTextEditor1.Name = "ultraTextEditor1";
			this.ultraTextEditor1.ReadOnly = true;
			this.ultraTextEditor1.Size = new System.Drawing.Size(88, 21);
			this.ultraTextEditor1.TabIndex = 3;
			// 
			// ultraLabel1
			// 
			this.ultraLabel1.AutoSize = true;
			this.ultraLabel1.Location = new System.Drawing.Point(16, 8);
			this.ultraLabel1.Name = "ultraLabel1";
			this.ultraLabel1.Size = new System.Drawing.Size(44, 14);
			this.ultraLabel1.TabIndex = 4;
			this.ultraLabel1.Text = "Numero";
			// 
			// ultraCalendarCombo1
			// 
			this.ultraCalendarCombo1.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsCompra, "Compra.Fecha"));
			dateButton1.Caption = "Today";
			this.ultraCalendarCombo1.DateButtons.Add(dateButton1);
			this.ultraCalendarCombo1.Format = "dd/MMM/yyyy";
			this.ultraCalendarCombo1.Location = new System.Drawing.Point(240, 8);
			this.ultraCalendarCombo1.Name = "ultraCalendarCombo1";
			this.ultraCalendarCombo1.NonAutoSizeHeight = 23;
			this.ultraCalendarCombo1.ReadOnly = true;
			this.ultraCalendarCombo1.Size = new System.Drawing.Size(104, 21);
			this.ultraCalendarCombo1.TabIndex = 5;
			this.ultraCalendarCombo1.Value = new System.DateTime(2011, 6, 3, 0, 0, 0, 0);
			// 
			// ultraLabel2
			// 
			this.ultraLabel2.AutoSize = true;
			this.ultraLabel2.Location = new System.Drawing.Point(192, 8);
			this.ultraLabel2.Name = "ultraLabel2";
			this.ultraLabel2.Size = new System.Drawing.Size(36, 14);
			this.ultraLabel2.TabIndex = 6;
			this.ultraLabel2.Text = "Fecha";
			// 
			// chkVerifica
			// 
			this.chkVerifica.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.chkVerifica.Checked = true;
			this.chkVerifica.CheckState = System.Windows.Forms.CheckState.Checked;
			this.chkVerifica.Location = new System.Drawing.Point(144, 223);
			this.chkVerifica.Name = "chkVerifica";
			this.chkVerifica.TabIndex = 9;
			this.chkVerifica.Text = "Verifica Porcentaje";
			// 
			// udsOrigen
			// 
			ultraDataColumn2.DataType = typeof(int);
			this.udsOrigen.Band.Columns.AddRange(new object[] {
																													ultraDataColumn1,
																													ultraDataColumn2});
			this.udsOrigen.Rows.AddRange(new object[] {
																									new Infragistics.Win.UltraWinDataSource.UltraDataRow(new object[] {
																																																											"Trabajo",
																																																											"Trabajo",
																																																											"idOrigen",
																																																											1}, new object[0]),
																									new Infragistics.Win.UltraWinDataSource.UltraDataRow(new object[] {
																																																											"Trabajo",
																																																											"Comision",
																																																											"idOrigen",
																																																											2}, new object[0]),
																									new Infragistics.Win.UltraWinDataSource.UltraDataRow(new object[] {
																																																											"Trabajo",
																																																											"Origen",
																																																											"idOrigen",
																																																											3}, new object[0]),
																									new Infragistics.Win.UltraWinDataSource.UltraDataRow(new object[] {
																																																											"Trabajo",
																																																											"Oficina",
																																																											"idOrigen",
																																																											4}, new object[0]),
																									new Infragistics.Win.UltraWinDataSource.UltraDataRow(new object[] {
																																																											"Trabajo",
																																																											"Inversion",
																																																											"idOrigen",
																																																											5}, new object[0])});
			// 
			// ultraDropDown1
			// 
			this.ultraDropDown1.Cursor = System.Windows.Forms.Cursors.Default;
			this.ultraDropDown1.DataMember = "CompraComisionOrig";
			this.ultraDropDown1.DataSource = this.cdsComision;
			ultraGridBand1.AddButtonCaption = "Band 0";
			ultraGridColumn1.Header.VisiblePosition = 2;
			ultraGridColumn1.Hidden = true;
			ultraGridColumn2.Header.VisiblePosition = 0;
			ultraGridColumn3.Header.VisiblePosition = 1;
			ultraGridBand1.Columns.AddRange(new object[] {
																										 ultraGridColumn1,
																										 ultraGridColumn2,
																										 ultraGridColumn3});
			this.ultraDropDown1.DisplayLayout.BandsSerializer.Add(ultraGridBand1);
			this.ultraDropDown1.DisplayMember = "Descripcion";
			this.ultraDropDown1.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.ultraDropDown1.Location = new System.Drawing.Point(280, 80);
			this.ultraDropDown1.Name = "ultraDropDown1";
			this.ultraDropDown1.Size = new System.Drawing.Size(168, 72);
			this.ultraDropDown1.TabIndex = 10;
			this.ultraDropDown1.ValueMember = "idOriginador";
			this.ultraDropDown1.Visible = false;
			// 
			// ultraDropDown2
			// 
			this.ultraDropDown2.Cursor = System.Windows.Forms.Cursors.Default;
			this.ultraDropDown2.DataMember = "Personal";
			this.ultraDropDown2.DataSource = this.cdsPersonal;
			ultraGridBand2.AddButtonCaption = "Band 0";
			ultraGridColumn4.Header.VisiblePosition = 0;
			ultraGridColumn4.Hidden = true;
			ultraGridColumn5.Header.VisiblePosition = 1;
			ultraGridColumn5.Hidden = true;
			ultraGridColumn6.Header.VisiblePosition = 2;
			ultraGridColumn6.Hidden = true;
			ultraGridColumn7.Header.VisiblePosition = 3;
			ultraGridColumn7.Hidden = true;
			ultraGridColumn8.Header.VisiblePosition = 4;
			ultraGridColumn8.Hidden = true;
			ultraGridColumn9.Header.VisiblePosition = 5;
			ultraGridColumn9.Hidden = true;
			ultraGridColumn10.Header.VisiblePosition = 6;
			ultraGridColumn10.Hidden = true;
			ultraGridColumn11.Header.VisiblePosition = 7;
			ultraGridColumn11.Hidden = true;
			ultraGridColumn12.Header.VisiblePosition = 8;
			ultraGridColumn13.Header.VisiblePosition = 9;
			ultraGridColumn13.Hidden = true;
			ultraGridColumn14.Header.VisiblePosition = 10;
			ultraGridColumn14.Hidden = true;
			ultraGridColumn15.Header.VisiblePosition = 11;
			ultraGridColumn15.Hidden = true;
			ultraGridColumn16.Header.VisiblePosition = 12;
			ultraGridColumn16.Hidden = true;
			ultraGridColumn17.Header.VisiblePosition = 13;
			ultraGridColumn17.Hidden = true;
			ultraGridColumn18.Header.VisiblePosition = 14;
			ultraGridColumn18.Hidden = true;
			ultraGridColumn19.Header.VisiblePosition = 15;
			ultraGridColumn19.Hidden = true;
			ultraGridColumn20.Header.VisiblePosition = 16;
			ultraGridColumn20.Hidden = true;
			ultraGridColumn21.Header.VisiblePosition = 17;
			ultraGridColumn21.Hidden = true;
			ultraGridColumn22.Header.VisiblePosition = 18;
			ultraGridColumn22.Hidden = true;
			ultraGridColumn23.Header.VisiblePosition = 19;
			ultraGridColumn23.Hidden = true;
			ultraGridColumn24.Header.VisiblePosition = 20;
			ultraGridColumn24.Hidden = true;
			ultraGridColumn25.Header.VisiblePosition = 21;
			ultraGridColumn25.Hidden = true;
			ultraGridColumn26.Header.VisiblePosition = 22;
			ultraGridColumn26.Hidden = true;
			ultraGridColumn27.Header.VisiblePosition = 23;
			ultraGridColumn27.Hidden = true;
			ultraGridColumn28.Header.VisiblePosition = 24;
			ultraGridColumn28.Hidden = true;
			ultraGridColumn29.Header.VisiblePosition = 25;
			ultraGridColumn29.Hidden = true;
			ultraGridColumn30.Header.VisiblePosition = 26;
			ultraGridColumn30.Hidden = true;
			ultraGridColumn31.Header.VisiblePosition = 27;
			ultraGridColumn31.Hidden = true;
			ultraGridColumn32.Header.VisiblePosition = 28;
			ultraGridColumn32.Hidden = true;
			ultraGridColumn33.Header.VisiblePosition = 29;
			ultraGridColumn33.Hidden = true;
			ultraGridColumn34.Header.VisiblePosition = 30;
			ultraGridColumn34.Hidden = true;
			ultraGridColumn35.Header.VisiblePosition = 31;
			ultraGridColumn35.Hidden = true;
			ultraGridColumn36.Header.VisiblePosition = 32;
			ultraGridColumn36.Hidden = true;
			ultraGridColumn37.Header.VisiblePosition = 33;
			ultraGridColumn37.Hidden = true;
			ultraGridColumn38.Header.VisiblePosition = 34;
			ultraGridColumn38.Hidden = true;
			ultraGridColumn39.Header.VisiblePosition = 35;
			ultraGridColumn39.Hidden = true;
			ultraGridColumn40.Header.VisiblePosition = 36;
			ultraGridColumn40.Hidden = true;
			ultraGridColumn41.Header.VisiblePosition = 37;
			ultraGridColumn41.Hidden = true;
			ultraGridColumn42.Header.VisiblePosition = 38;
			ultraGridColumn42.Hidden = true;
			ultraGridColumn43.Header.VisiblePosition = 39;
			ultraGridColumn43.Hidden = true;
			ultraGridColumn44.Header.VisiblePosition = 40;
			ultraGridColumn44.Hidden = true;
			ultraGridColumn45.Header.VisiblePosition = 41;
			ultraGridColumn45.Hidden = true;
			ultraGridColumn46.Header.VisiblePosition = 42;
			ultraGridColumn46.Hidden = true;
			ultraGridColumn47.Header.VisiblePosition = 43;
			ultraGridColumn47.Hidden = true;
			ultraGridColumn48.Header.VisiblePosition = 44;
			ultraGridColumn48.Hidden = true;
			ultraGridColumn49.Header.VisiblePosition = 45;
			ultraGridColumn49.Hidden = true;
			ultraGridColumn50.Header.VisiblePosition = 46;
			ultraGridColumn50.Hidden = true;
			ultraGridColumn51.Header.VisiblePosition = 47;
			ultraGridColumn51.Hidden = true;
			ultraGridColumn52.Header.VisiblePosition = 48;
			ultraGridColumn52.Hidden = true;
			ultraGridColumn53.Header.VisiblePosition = 49;
			ultraGridColumn53.Hidden = true;
			ultraGridColumn54.Header.VisiblePosition = 50;
			ultraGridColumn54.Hidden = true;
			ultraGridColumn55.Header.VisiblePosition = 51;
			ultraGridColumn55.Hidden = true;
			ultraGridColumn56.Header.VisiblePosition = 52;
			ultraGridColumn56.Hidden = true;
			ultraGridColumn57.Header.VisiblePosition = 53;
			ultraGridColumn57.Hidden = true;
			ultraGridColumn58.Header.VisiblePosition = 54;
			ultraGridColumn58.Hidden = true;
			ultraGridColumn59.Header.VisiblePosition = 55;
			ultraGridColumn59.Hidden = true;
			ultraGridColumn60.Header.VisiblePosition = 56;
			ultraGridColumn60.Hidden = true;
			ultraGridColumn61.Header.VisiblePosition = 57;
			ultraGridColumn61.Hidden = true;
			ultraGridColumn62.Header.VisiblePosition = 58;
			ultraGridColumn62.Hidden = true;
			ultraGridColumn63.Header.VisiblePosition = 59;
			ultraGridColumn63.Hidden = true;
			ultraGridColumn64.Header.VisiblePosition = 60;
			ultraGridColumn64.Hidden = true;
			ultraGridColumn65.Header.VisiblePosition = 61;
			ultraGridColumn65.Hidden = true;
			ultraGridColumn66.Header.VisiblePosition = 62;
			ultraGridColumn66.Hidden = true;
			ultraGridColumn67.Header.VisiblePosition = 63;
			ultraGridColumn67.Hidden = true;
			ultraGridColumn68.Header.VisiblePosition = 64;
			ultraGridColumn68.Hidden = true;
			ultraGridColumn69.Header.VisiblePosition = 65;
			ultraGridColumn69.Hidden = true;
			ultraGridColumn70.Header.VisiblePosition = 66;
			ultraGridColumn70.Hidden = true;
			ultraGridColumn71.Header.VisiblePosition = 67;
			ultraGridColumn71.Hidden = true;
			ultraGridColumn72.Header.VisiblePosition = 68;
			ultraGridColumn72.Hidden = true;
			ultraGridColumn73.Header.VisiblePosition = 69;
			ultraGridColumn73.Hidden = true;
			ultraGridColumn74.Header.VisiblePosition = 70;
			ultraGridColumn74.Hidden = true;
			ultraGridColumn75.Header.VisiblePosition = 71;
			ultraGridColumn75.Hidden = true;
			ultraGridColumn76.Header.VisiblePosition = 72;
			ultraGridColumn76.Hidden = true;
			ultraGridColumn77.Header.VisiblePosition = 73;
			ultraGridColumn77.Hidden = true;
			ultraGridColumn78.Header.VisiblePosition = 74;
			ultraGridColumn78.Hidden = true;
			ultraGridColumn79.Header.VisiblePosition = 75;
			ultraGridColumn79.Hidden = true;
			ultraGridColumn80.Header.VisiblePosition = 76;
			ultraGridColumn80.Hidden = true;
			ultraGridColumn81.Header.VisiblePosition = 77;
			ultraGridColumn81.Hidden = true;
			ultraGridColumn82.Header.VisiblePosition = 78;
			ultraGridColumn82.Hidden = true;
			ultraGridColumn83.Header.VisiblePosition = 79;
			ultraGridColumn83.Hidden = true;
			ultraGridColumn84.Header.VisiblePosition = 80;
			ultraGridColumn84.Hidden = true;
			ultraGridColumn85.Header.VisiblePosition = 81;
			ultraGridColumn85.Hidden = true;
			ultraGridColumn86.Header.VisiblePosition = 82;
			ultraGridColumn86.Hidden = true;
			ultraGridColumn87.Header.VisiblePosition = 83;
			ultraGridColumn87.Hidden = true;
			ultraGridColumn88.Header.VisiblePosition = 84;
			ultraGridColumn88.Hidden = true;
			ultraGridColumn89.Header.VisiblePosition = 85;
			ultraGridColumn89.Hidden = true;
			ultraGridColumn90.Header.VisiblePosition = 86;
			ultraGridColumn90.Hidden = true;
			ultraGridColumn91.Header.VisiblePosition = 87;
			ultraGridColumn91.Hidden = true;
			ultraGridColumn92.Header.VisiblePosition = 88;
			ultraGridColumn92.Hidden = true;
			ultraGridColumn93.Header.VisiblePosition = 89;
			ultraGridColumn93.Hidden = true;
			ultraGridColumn94.Header.VisiblePosition = 90;
			ultraGridColumn94.Hidden = true;
			ultraGridBand2.Columns.AddRange(new object[] {
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
																										 ultraGridColumn94});
			this.ultraDropDown2.DisplayLayout.BandsSerializer.Add(ultraGridBand2);
			this.ultraDropDown2.DisplayMember = "Nombre";
			this.ultraDropDown2.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.ultraDropDown2.Location = new System.Drawing.Point(96, 72);
			this.ultraDropDown2.Name = "ultraDropDown2";
			this.ultraDropDown2.Size = new System.Drawing.Size(152, 72);
			this.ultraDropDown2.TabIndex = 11;
			this.ultraDropDown2.ValueMember = "idPersonal";
			this.ultraDropDown2.Visible = false;
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
			// cdsOriginador
			// 
			this.cdsOriginador.BindingContextCtrl = this;
			this.cdsOriginador.DataLibrary = "LibFacturacion";
			this.cdsOriginador.DataLibraryUrl = "";
			this.cdsOriginador.DataSetDef = "dsComision";
			this.cdsOriginador.EnforceConstraints = false;
			this.cdsOriginador.FillOnRequest = false;
			this.cdsOriginador.Locale = new System.Globalization.CultureInfo("es-EC");
			this.cdsOriginador.Name = "cdsOriginador";
			this.cdsOriginador.SchemaClassName = "LibFacturacion.DataClass";
			this.cdsOriginador.SchemaDef = null;
			this.cdsOriginador.BeforeFill += new C1.Data.FillEventHandler(this.cdsOriginador_BeforeFill);
			// 
			// ultraDropDown3
			// 
			this.ultraDropDown3.Cursor = System.Windows.Forms.Cursors.Default;
			this.ultraDropDown3.DataMember = "Band 0";
			this.ultraDropDown3.DataSource = this.udsOrigen;
			ultraGridColumn95.Header.VisiblePosition = 0;
			ultraGridColumn96.Header.VisiblePosition = 1;
			ultraGridColumn96.Hidden = true;
			ultraGridBand3.Columns.AddRange(new object[] {
																										 ultraGridColumn95,
																										 ultraGridColumn96});
			this.ultraDropDown3.DisplayLayout.BandsSerializer.Add(ultraGridBand3);
			this.ultraDropDown3.DisplayMember = "Trabajo";
			this.ultraDropDown3.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.ultraDropDown3.Location = new System.Drawing.Point(384, 112);
			this.ultraDropDown3.Name = "ultraDropDown3";
			this.ultraDropDown3.Size = new System.Drawing.Size(152, 72);
			this.ultraDropDown3.TabIndex = 12;
			this.ultraDropDown3.ValueMember = "idOrigen";
			this.ultraDropDown3.Visible = false;
			// 
			// btCliente
			// 
			this.btCliente.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.btCliente.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			this.btCliente.Location = new System.Drawing.Point(616, 8);
			this.btCliente.Name = "btCliente";
			this.btCliente.Size = new System.Drawing.Size(24, 23);
			this.btCliente.TabIndex = 77;
			this.btCliente.Text = "...";
			this.btCliente.Visible = false;
			this.btCliente.Click += new System.EventHandler(this.btCliente_Click);
			// 
			// grAsiento
			// 
			this.grAsiento.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.grAsiento.Cursor = System.Windows.Forms.Cursors.Default;
			this.grAsiento.DataMember = "CompraComision";
			this.grAsiento.DataSource = this.cdsComision;
			appearance1.BackColor = System.Drawing.Color.White;
			this.grAsiento.DisplayLayout.Appearance = appearance1;
			ultraGridColumn97.Header.VisiblePosition = 6;
			ultraGridColumn97.Hidden = true;
			ultraGridColumn98.Header.VisiblePosition = 5;
			ultraGridColumn98.Hidden = true;
			ultraGridColumn99.Header.VisiblePosition = 0;
			ultraGridColumn99.Width = 183;
			appearance2.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn100.CellAppearance = appearance2;
			ultraGridColumn100.Format = "#,##0.00";
			ultraGridColumn100.Header.VisiblePosition = 4;
			appearance3.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn101.CellAppearance = appearance3;
			ultraGridColumn101.Format = "#,##0.00";
			ultraGridColumn101.Header.VisiblePosition = 7;
			ultraGridColumn102.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			appearance4.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn102.CellAppearance = appearance4;
			ultraGridColumn102.Format = "#,##0.00";
			ultraGridColumn102.Header.VisiblePosition = 8;
			ultraGridColumn103.Header.Caption = "Originador";
			ultraGridColumn103.Header.VisiblePosition = 1;
			ultraGridColumn103.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.DropDownValidate;
			ultraGridColumn104.Header.Caption = "Nombre";
			ultraGridColumn104.Header.VisiblePosition = 2;
			ultraGridColumn104.Hidden = true;
			ultraGridColumn104.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.DropDownValidate;
			ultraGridColumn104.Width = 186;
			ultraGridColumn105.Header.Caption = "Tipo Abogado";
			ultraGridColumn105.Header.VisiblePosition = 3;
			ultraGridColumn105.Hidden = true;
			ultraGridColumn105.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.DropDownValidate;
			ultraGridBand4.Columns.AddRange(new object[] {
																										 ultraGridColumn97,
																										 ultraGridColumn98,
																										 ultraGridColumn99,
																										 ultraGridColumn100,
																										 ultraGridColumn101,
																										 ultraGridColumn102,
																										 ultraGridColumn103,
																										 ultraGridColumn104,
																										 ultraGridColumn105});
			appearance5.TextHAlign = Infragistics.Win.HAlign.Right;
			summarySettings1.Appearance = appearance5;
			summarySettings1.DisplayFormat = "{0: #,##0.00}";
			summarySettings1.ShowCalculatingText = Infragistics.Win.DefaultableBoolean.False;
			ultraGridBand4.Summaries.AddRange(new Infragistics.Win.UltraWinGrid.SummarySettings[] {
																																															summarySettings1});
			ultraGridBand4.SummaryFooterCaption = "";
			this.grAsiento.DisplayLayout.BandsSerializer.Add(ultraGridBand4);
			this.grAsiento.DisplayLayout.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.TemplateOnBottom;
			this.grAsiento.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.True;
			this.grAsiento.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.True;
			appearance6.BackColor = System.Drawing.Color.Transparent;
			this.grAsiento.DisplayLayout.Override.CardAreaAppearance = appearance6;
			appearance7.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance7.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance7.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance7.FontData.BoldAsString = "True";
			appearance7.FontData.Name = "Arial";
			appearance7.FontData.SizeInPoints = 10F;
			appearance7.ForeColor = System.Drawing.Color.White;
			appearance7.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.grAsiento.DisplayLayout.Override.HeaderAppearance = appearance7;
			appearance8.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance8.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance8.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.grAsiento.DisplayLayout.Override.RowSelectorAppearance = appearance8;
			appearance9.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance9.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance9.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.grAsiento.DisplayLayout.Override.SelectedRowAppearance = appearance9;
			this.grAsiento.Location = new System.Drawing.Point(16, 40);
			this.grAsiento.Name = "grAsiento";
			this.grAsiento.Size = new System.Drawing.Size(624, 167);
			this.grAsiento.TabIndex = 0;
			this.grAsiento.InitializeRow += new Infragistics.Win.UltraWinGrid.InitializeRowEventHandler(this.grAsiento_InitializeRow);
			this.grAsiento.AfterRowInsert += new Infragistics.Win.UltraWinGrid.RowEventHandler(this.ultraGrid1_AfterRowInsert);
			this.grAsiento.BeforeRowInsert += new Infragistics.Win.UltraWinGrid.BeforeRowInsertEventHandler(this.grAsiento_BeforeRowInsert);
			this.grAsiento.KeyDown += new System.Windows.Forms.KeyEventHandler(this.grAsiento_KeyDown);
			this.grAsiento.AfterCellUpdate += new Infragistics.Win.UltraWinGrid.CellEventHandler(this.grAsiento_AfterCellUpdate);
			// 
			// label1
			// 
			this.label1.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.label1.Location = new System.Drawing.Point(512, 8);
			this.label1.Name = "label1";
			this.label1.TabIndex = 80;
			this.label1.Text = "Tipo Abogado";
			// 
			// btCalcular
			// 
			this.btCalcular.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.btCalcular.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			this.btCalcular.Location = new System.Drawing.Point(584, 216);
			this.btCalcular.Name = "btCalcular";
			this.btCalcular.Size = new System.Drawing.Size(56, 23);
			this.btCalcular.TabIndex = 81;
			this.btCalcular.Text = "&Calcular";
			this.btCalcular.Visible = false;
			this.btCalcular.Click += new System.EventHandler(this.btCalcular_Click);
			// 
			// CompraComision
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 13);
			this.ClientSize = new System.Drawing.Size(664, 261);
			this.Controls.Add(this.btCalcular);
			this.Controls.Add(this.label1);
			this.Controls.Add(this.btCliente);
			this.Controls.Add(this.ultraDropDown3);
			this.Controls.Add(this.ultraDropDown2);
			this.Controls.Add(this.ultraDropDown1);
			this.Controls.Add(this.chkVerifica);
			this.Controls.Add(this.ultraLabel2);
			this.Controls.Add(this.ultraCalendarCombo1);
			this.Controls.Add(this.ultraLabel1);
			this.Controls.Add(this.ultraTextEditor1);
			this.Controls.Add(this.txtIdCompra);
			this.Controls.Add(this.barraDatoSQL1);
			this.Controls.Add(this.grAsiento);
			this.Name = "CompraComision";
			this.Text = "Comisiones por factura";
			this.Load += new System.EventHandler(this.CompraComision_Load);
			((System.ComponentModel.ISupportInitialize)(this.cdsComision)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsCompra)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIdCompra)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraTextEditor1)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraCalendarCombo1)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.udsOrigen)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDropDown1)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDropDown2)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsPersonal)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsOriginador)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDropDown3)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.grAsiento)).EndInit();
			this.ResumeLayout(false);

		}
		#endregion

		private void barraDatoSQL1_Borrar(object sender, System.EventArgs e)
		{
		
		}

		private void barraDatoSQL1_Buscar(object sender, System.EventArgs e)
		{
		
		}

		private void barraDatoSQL1_Copiar(object sender, System.EventArgs e)
		{
		
		}

		private void barraDatoSQL1_Crear(object sender, System.EventArgs e)
		{
			MessageBox.Show("No se puede crear comisiones", "Información");
		}

		private void barraDatoSQL1_Deshacer(object sender, System.EventArgs e)
		{
			barraDatoSQL1.DeshacerRegistro();
		}

		private void barraDatoSQL1_Editar(object sender, System.EventArgs e)
		{
			barraDatoSQL1.EditarRegistro();
		}

		private bool Verifica()
		{
			bool bOk = true;
			if (!chkVerifica.Checked) return true;
			double dPorcentaje = 0;
			foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in grAsiento.Rows.All)
			{
				if (dr.Cells["Porcentaje"].Value != System.DBNull.Value)
					dPorcentaje += (double) dr.Cells["Porcentaje"].Value;
			}
			if (Math.Abs(Math.Round(dPorcentaje, 2) - 100.0) >= 0.01)
			{
				MessageBox.Show("Suma de Porcentaje es: " 
					+ dPorcentaje.ToString("0.00") +" debe ser 100", "Información");
				bOk = false;
			}
			return bOk;
		}

		private void barraDatoSQL1_Grabar(object sender, System.EventArgs e)
		{
			if (!Verifica()) return;
			if (Funcion.bEjecutaSQL(cdsComision, "Exec SeteoGExiste 'FVCLEXIM'")) 				
			{
				ReCalculo();
			}
			
			try
			{
				this.cdsComision.Update();
				barraDatoSQL1.GrabaRegistro();
			}
			catch (System.Exception ex)
			{
				MessageBox.Show(ex.Message, "Error al Grabar");
			}
		}

		private void barraDatoSQL1_Impresora(object sender, System.EventArgs e)
		{
		
		}

		private void barraDatoSQL1_Imprimir(object sender, System.EventArgs e)
		{

		if (Funcion.bEjecutaSQL(cdsComision, "Exec SeteoGExiste 'FVCLEXIM'")) 				
		{
			
		Funcion.EjecutaSQL(cdsComision,"Exec P_gananciaBuffete");

		}

			CompraComImprime miImprime = new CompraComImprime(int.Parse(txtIdCompra.Value.ToString()));
			miImprime.MdiParent=this.MdiParent;
			miImprime.Show();
		}

		private void barraDatoSQL1_Refresca(object sender, System.EventArgs e)
		{
			try
			{
				C1.Data.FilterCondition fcPadre = new 
					C1.Data.FilterCondition(cdsComision, "CompraComision", " [idCompra] = " + txtIdCompra.Value.ToString());
				C1.Data.FilterConditions fcTotal = new C1.Data.FilterConditions();
				fcTotal.Add(fcPadre);

				cdsComision.Fill(fcTotal, false);
			}
			catch(System.Exception ex)
			{
				MessageBox.Show(ex.Message, "error refresca");
			}
		}

		private void barraDatoSQL1_Restaurar(object sender, System.EventArgs e)
		{
		}

		private void CompraComision_Load(object sender, System.EventArgs e)
		{
						if (Funcion.bEjecutaSQL(cdsComision, "Exec SeteoGExiste 'FVCLEXIM'")) 				
						{								
							this.grAsiento.DisplayLayout.Bands[0].Columns["idOriginador"].Hidden = true;										
							this.grAsiento.DisplayLayout.Bands[0].Columns["Nombre"].Hidden = true;				
							this.grAsiento.DisplayLayout.Bands[0].Columns["idOriginadorT"].Hidden = false;										
							this.grAsiento.DisplayLayout.Bands[0].Columns["idVendedor"].Hidden = false;
							btCliente.Visible=true;
							btCalcular.Visible=true;
							chkVerifica.Checked = false;
						}				

			txtIdCompra.Width=0;
			if (IdCompra>0)
			{
				//				string stSelect = "Select Count(*) From CompraComision Where idCompra = " + IdCompra.ToString();
				//				int iCuenta = Funcion.iEscalarSQL(cdsComision, stSelect, true);
				//				if (iCuenta==0)
				//				{
				//					stSelect
				//				}
				barraDatoSQL1.IdRegistro=IdCompra;
				barraDatoSQL1.ProximoId(5);
			}
			else
				barraDatoSQL1.ProximoId(4);

			barraDatoSQL1.HabilitaControles(false);
		}

		private void ultraGrid1_AfterRowInsert(object sender, Infragistics.Win.UltraWinGrid.RowEventArgs e)
		{
			e.Row.Cells["idCompra"].Value = (int) txtIdCompra.Value;			
			e.Row.Cells["Valor"].Value=0;
			e.Row.Cells["Porcentaje"].Value=0;		
		}

		private void grAsiento_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.Alt) 
				return;
			if (grAsiento.ActiveCell == null) return;
			switch(e.KeyValue)
			{
				case (int) Keys.PageDown:
					if(grAsiento.ActiveCell.DroppedDown == false)
						grAsiento.ActiveCell.DroppedDown = true;
					break;

				case (int) Keys.Up:
					if (grAsiento.ActiveCell.DroppedDown) return;

					grAsiento.PerformAction(UltraGridAction.ExitEditMode, false, false);
					grAsiento.PerformAction(UltraGridAction.AboveCell, false, false);
					e.Handled = true;
					grAsiento.PerformAction(UltraGridAction.EnterEditMode, false, false);
					break;
				case (int) Keys.Down:
					if (grAsiento.ActiveCell.DroppedDown) return;
			
					grAsiento.PerformAction(UltraGridAction.ExitEditMode, false, false);
					grAsiento.PerformAction(UltraGridAction.BelowCell, false, false);
					e.Handled = true;
					grAsiento.PerformAction(UltraGridAction.EnterEditMode, false, false);
					break;
				case (int) Keys.Enter:
					grAsiento.PerformAction(UltraGridAction.ExitEditMode, false, false);
					grAsiento.PerformAction(UltraGridAction.NextCellByTab, false, false);
					e.Handled = true;
					grAsiento.PerformAction(UltraGridAction.EnterEditMode, false, false);
					break;
			}		
		}

		private void grAsiento_InitializeRow(object sender, Infragistics.Win.UltraWinGrid.InitializeRowEventArgs e)
		{
			if (e.Row.Cells["Valor"].Value != DBNull.Value && e.Row.Cells["Porcentaje"].Value != DBNull.Value)
			e.Row.Cells["Total"].Value = (double) e.Row.Cells["Valor"].Value * (double) e.Row.Cells["Porcentaje"].Value / 100.0;
		}

		private void cmbTotal_Click(object sender, System.EventArgs e)
		{
//			double dTotal = 0;
//			foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in grAsiento.Rows.All)
//			{
//				if (dr.Cells["Total"].Value != System.DBNull.Value)
//					dTotal += (double) dr.Cells["Total"].Value;
//			}
//			txtTotal.Value = dTotal;
		}

		private void grAsiento_BeforeRowInsert(object sender, Infragistics.Win.UltraWinGrid.BeforeRowInsertEventArgs e)
		{
		}	
		private void cdsOriginador_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			this.cdsOriginador.Schema.Connections[0].ConnectionString = 
				miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);
		}

		private void cdsPersonal_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			cdsPersonal.Schema.Connections[0].ConnectionString = 
				miFuncion.AccesoBase(MenuLatinium.stDirRoles);		
		}

		private void btCliente_Click(object sender, System.EventArgs e)
		{
			CompraComisionOrg cco = new CompraComisionOrg();
			cco.ShowDialog();
			this.cdsOriginador.Clear();
			this.cdsOriginador.Fill();
			this.ultraDropDown1.Refresh();
		}
		
		private void grAsiento_AfterCellUpdate(object sender, Infragistics.Win.UltraWinGrid.CellEventArgs e)
		{			
			if (Funcion.bEjecutaSQL(cdsComision, "Exec SeteoGExiste 'FVCLEXIM'")) 				
			{	
				string stSelect = "select SUM(SubtotalExcento+SubtotalIva) from Compra where idCompra = "
					+txtIdCompra.Value;
				double dbSuma = Funcion.dEscalarSQL(cdsComision,stSelect);

				int CValorP=0;
				int CTotal=0;
	try{
				if (e.Cell.IsInEditMode)
				{
					//if (!chkVerifica.Checked) return;										
					if (e.Cell.Column.ToString() == "Porcentaje")
					{
						int idOrig = int.Parse(grAsiento.ActiveRow.Cells["idOriginadorT"].Value.ToString());												
						foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in grAsiento.Rows.All)
						{
							if(int.Parse(dr.Cells["idOriginadorT"].Value.ToString()) == idOrig)
								CValorP = int.Parse(dr.Cells["Porcentaje"].Value.ToString()) + CValorP;

							if(int.Parse(dr.Cells["idOriginadorT"].Value.ToString()) == 1)
								CTotal = int.Parse(dr.Cells["Total"].Value.ToString()) + CTotal;

						}
						if (grAsiento.ActiveRow.Cells["idOriginadorT"].Value.ToString() == "1")
						{
							grAsiento.ActiveRow.Cells["Valor"].Value = dbSuma;
						}
						if (grAsiento.ActiveRow.Cells["idOriginadorT"].Value.ToString() == "2")
						{
							grAsiento.ActiveRow.Cells["Valor"].Value = dbSuma - CTotal;
						}

						string sqlProced = "select ValorBase from compraComisionorig where idOriginador ="+ idOrig;
						int valorRetorna = Funcion.iEscalarSQL(cdsComision,sqlProced);
						if(valorRetorna > 0)
						{
							if(CValorP > valorRetorna)
							{
								if(idOrig ==1)
									MessageBox.Show("Valor Excede Base Asociado: \n\n Base: "+valorRetorna+" ValorAsignado: "+CValorP);

								if(idOrig ==2)
									MessageBox.Show("Valor Excede Base Socio:  \n\n Base: "+valorRetorna+" ValorAsignado: "+CValorP);
								return;
							}
						}	
						CValorP = 0;	
						CTotal = 0;
					}
				}
			}catch{}
			}
		}

		private void btCalcular_Click(object sender, System.EventArgs e)
		{
			ReCalculo();
		}
		private void ReCalculo()
		{
			double Acumula=0.0;
			double Valor=0.0;
			double dbSuma = 0.0;
			string stSelect = "select SUM(SubtotalExcento+SubtotalIva) from Compra where idCompra = "
				+txtIdCompra.Value;
			dbSuma = Funcion.dEscalarSQL(cdsComision,stSelect);			
			foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in grAsiento.Rows.All)
			{
				if (dr.Cells["idOriginadorT"].Value != System.DBNull.Value)
				{
					if (int.Parse( dr.Cells["idOriginadorT"].Value.ToString()) == 1)
					{
						dr.Cells["Valor"].Value = dbSuma;
						dr.Cells["Total"].Value = dbSuma * (double) dr.Cells["Porcentaje"].Value / 100.0;
						Acumula =Acumula+(double) dr.Cells["Valor"].Value * (double) dr.Cells["Porcentaje"].Value / 100.0;
					}				
				}
			}			
			Valor = dbSuma-Acumula;
			foreach(Infragistics.Win.UltraWinGrid.UltraGridRow dr in grAsiento.Rows.All)
			{	
				if (dr.Cells["idOriginadorT"].Value != System.DBNull.Value)
				{
					if (int.Parse( dr.Cells["idOriginadorT"].Value.ToString()) == 2)
					{		
						dr.Cells["Valor"].Value = Valor;					
						dr.Cells["Total"].Value = Valor * (double) dr.Cells["Porcentaje"].Value / 100.0;
					}
				}
			}
		}

	}
}
