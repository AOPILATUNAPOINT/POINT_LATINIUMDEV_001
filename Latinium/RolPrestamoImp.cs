using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;

namespace Latinium
{
	/// <summary>
	/// Descripción breve de RolPrestamoImp.
	/// </summary>
	public class RolPrestamoImp : DevExpress.XtraEditors.XtraForm
	{
		private Infragistics.Win.UltraWinSchedule.UltraCalendarCombo cmbFechaIni;
		private System.Windows.Forms.Label label1;
		private System.Windows.Forms.Label label2;
		private C1.Data.C1DataSet cdsTipoPrestamo;
		private C1.Data.C1DataSet cdsPersonal;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbTipo;
		private C1.Data.C1DataView cdvPersonal;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbPersonal;
		private Infragistics.Win.UltraWinEditors.UltraCheckEditor chkTodoPersonal;
		private Infragistics.Win.Misc.UltraButton btAceptar;
		private Infragistics.Win.Misc.UltraButton btCancelar;
		private Infragistics.Win.UltraWinEditors.UltraCheckEditor chkTodoTipo;
		private System.Windows.Forms.ErrorProvider errorProvider;
		private Infragistics.Win.UltraWinSchedule.UltraCalendarCombo cmbFechaFin;
		private Infragistics.Win.UltraWinEditors.UltraCheckEditor chkFecha;
		/// <summary>
		/// Variable del diseñador requerida.
		/// </summary>
		private System.ComponentModel.Container components = null;

		public RolPrestamoImp()
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
			Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton dateButton2 = new Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand1 = new Infragistics.Win.UltraWinGrid.UltraGridBand("RolTipoPrestamo", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn1 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idRolTipoPrest");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn2 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("TipoPrestamo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn3 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Campo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn4 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCuenta");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn5 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idRolGrupo");
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand2 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Personal", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn6 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idPersonal");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn7 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idAsiDistGrupo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn8 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCentroCosto");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn9 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idDepartamento");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn10 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCargo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn11 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idRolSucursal");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn12 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idRolGrupo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn13 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Jornal");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn14 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn15 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Apellido");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn16 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Codigo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn17 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Jornada");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn18 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Direccion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn19 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Telefono");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn20 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Celular");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn21 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Cedula");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn22 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Afiliacion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn23 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Reserva");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn24 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Renta");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn25 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Iess");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn26 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Sueldo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn27 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Comision");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn28 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Tipo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn29 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("FechaIngreso");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn30 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("FechaNacimiento");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn31 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("FechaSalida");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn32 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Notas");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn33 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("CargasFamiliares");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn34 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idProvincia");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn35 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCiudad");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn36 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Numero");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn37 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("SistemaSalario");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn38 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Cuenta");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn39 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("CodigoSectorial");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn40 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Condicion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn41 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("FechaNovedad");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn42 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Acceso");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn43 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Vendedor");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn44 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Borrar");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn45 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Cobrador");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn46 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Sexo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn47 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Estado_Civil");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn48 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Banco");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn49 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("CuentaCorr");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn50 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idTipoCuenta");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn51 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("CedulaMilitar");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn52 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idFormaPago");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn53 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Quincena");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn54 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("CodIngreso");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn55 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("CodTelefono");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn56 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("TarjEdificio");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn57 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("TarjetaOficina");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn58 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("email");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn59 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idProyecto");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn60 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCuenta");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn61 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Bono");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn62 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Bono1");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn63 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Desc1");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn64 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Desc2");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn65 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("TieneQuincena");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn66 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Fotos");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn67 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Convenio");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn68 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idSubProyecto");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn69 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Sindicato");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn70 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Asociacion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn71 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idRolHorario");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn72 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("FechaContrato");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn73 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Educacion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn74 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("NumeroDias");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn75 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idRelacionT");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn76 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("gSalud");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn77 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("gEducacion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn78 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("gVivienda");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn79 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("gVestimenta");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn80 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("gAlimentacion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn81 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("gTotal");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn82 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("BaseComision");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn83 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("PorcentajeComision");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn84 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("gDesahucio");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn85 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("gDiscapacidad");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn86 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("IngGravOE");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn87 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("DeduccionOE");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn88 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("OtrasRebOE");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn89 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("ImpRentaAsumido");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn90 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("gRebajaIIIEdad");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn91 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("RentaRetenidaOE");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn92 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("RentaRetOE");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn93 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("TipoSangre");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn94 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("RegEntrada");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn95 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("RegSalida");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn96 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idTipoContrato");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn97 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("CodVendedor");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn98 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Clave");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn99 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idSupervisor");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn100 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("bSupervisor");
			this.cmbFechaIni = new Infragistics.Win.UltraWinSchedule.UltraCalendarCombo();
			this.label1 = new System.Windows.Forms.Label();
			this.label2 = new System.Windows.Forms.Label();
			this.cmbFechaFin = new Infragistics.Win.UltraWinSchedule.UltraCalendarCombo();
			this.cdsTipoPrestamo = new C1.Data.C1DataSet();
			this.cmbTipo = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.cdsPersonal = new C1.Data.C1DataSet();
			this.cdvPersonal = new C1.Data.C1DataView();
			this.cmbPersonal = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.chkTodoPersonal = new Infragistics.Win.UltraWinEditors.UltraCheckEditor();
			this.chkTodoTipo = new Infragistics.Win.UltraWinEditors.UltraCheckEditor();
			this.btAceptar = new Infragistics.Win.Misc.UltraButton();
			this.btCancelar = new Infragistics.Win.Misc.UltraButton();
			this.errorProvider = new System.Windows.Forms.ErrorProvider();
			this.chkFecha = new Infragistics.Win.UltraWinEditors.UltraCheckEditor();
			((System.ComponentModel.ISupportInitialize)(this.cmbFechaIni)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbFechaFin)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsTipoPrestamo)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbTipo)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsPersonal)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdvPersonal)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbPersonal)).BeginInit();
			this.SuspendLayout();
			// 
			// cmbFechaIni
			// 
			dateButton1.Caption = "Today";
			this.cmbFechaIni.DateButtons.Add(dateButton1);
			this.cmbFechaIni.Format = "dd/MMM/yyyy";
			this.cmbFechaIni.Location = new System.Drawing.Point(112, 52);
			this.cmbFechaIni.Name = "cmbFechaIni";
			this.cmbFechaIni.NonAutoSizeHeight = 8;
			this.cmbFechaIni.Size = new System.Drawing.Size(128, 21);
			this.cmbFechaIni.SpinButtonsVisible = true;
			this.cmbFechaIni.TabIndex = 0;
			// 
			// label1
			// 
			this.label1.AutoSize = true;
			this.label1.Location = new System.Drawing.Point(56, 52);
			this.label1.Name = "label1";
			this.label1.Size = new System.Drawing.Size(35, 16);
			this.label1.TabIndex = 1;
			this.label1.Text = "Desde";
			// 
			// label2
			// 
			this.label2.AutoSize = true;
			this.label2.Location = new System.Drawing.Point(56, 86);
			this.label2.Name = "label2";
			this.label2.Size = new System.Drawing.Size(33, 16);
			this.label2.TabIndex = 3;
			this.label2.Text = "Hasta";
			// 
			// cmbFechaFin
			// 
			dateButton2.Caption = "Today";
			this.cmbFechaFin.DateButtons.Add(dateButton2);
			this.cmbFechaFin.Format = "dd/MMM/yyyy";
			this.cmbFechaFin.Location = new System.Drawing.Point(112, 86);
			this.cmbFechaFin.Name = "cmbFechaFin";
			this.cmbFechaFin.NonAutoSizeHeight = 8;
			this.cmbFechaFin.Size = new System.Drawing.Size(128, 21);
			this.cmbFechaFin.SpinButtonsVisible = true;
			this.cmbFechaFin.TabIndex = 2;
			// 
			// cdsTipoPrestamo
			// 
			this.cdsTipoPrestamo.BindingContextCtrl = this;
			this.cdsTipoPrestamo.DataLibrary = "LibPersonal";
			this.cdsTipoPrestamo.DataLibraryUrl = "";
			this.cdsTipoPrestamo.DataSetDef = "dsRolTipoPrestamo";
			this.cdsTipoPrestamo.Locale = new System.Globalization.CultureInfo("es-EC");
			this.cdsTipoPrestamo.Name = "cdsTipoPrestamo";
			this.cdsTipoPrestamo.SchemaClassName = "LibPersonal.DataClass";
			this.cdsTipoPrestamo.SchemaDef = null;
			// 
			// cmbTipo
			// 
			this.cmbTipo.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbTipo.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbTipo.DataMember = "RolTipoPrestamo";
			this.cmbTipo.DataSource = this.cdsTipoPrestamo;
			ultraGridColumn1.Header.VisiblePosition = 1;
			ultraGridColumn1.Hidden = true;
			ultraGridColumn2.Header.VisiblePosition = 0;
			ultraGridColumn2.Width = 262;
			ultraGridColumn3.Header.VisiblePosition = 2;
			ultraGridColumn3.Hidden = true;
			ultraGridColumn4.Header.VisiblePosition = 3;
			ultraGridColumn4.Hidden = true;
			ultraGridColumn5.Header.VisiblePosition = 4;
			ultraGridColumn5.Hidden = true;
			ultraGridBand1.Columns.AddRange(new object[] {
																										 ultraGridColumn1,
																										 ultraGridColumn2,
																										 ultraGridColumn3,
																										 ultraGridColumn4,
																										 ultraGridColumn5});
			this.cmbTipo.DisplayLayout.BandsSerializer.Add(ultraGridBand1);
			this.cmbTipo.DisplayMember = "TipoPrestamo";
			this.cmbTipo.Enabled = false;
			this.cmbTipo.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbTipo.Location = new System.Drawing.Point(136, 121);
			this.cmbTipo.Name = "cmbTipo";
			this.cmbTipo.Size = new System.Drawing.Size(128, 21);
			this.cmbTipo.TabIndex = 4;
			this.cmbTipo.ValueMember = "idRolTipoPrest";
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
			// cdvPersonal
			// 
			this.cdvPersonal.BindingContextCtrl = this;
			this.cdvPersonal.DataSet = this.cdsPersonal;
			this.cdvPersonal.Sort = "Apellido";
			this.cdvPersonal.TableName = "";
			this.cdvPersonal.TableViewName = "Personal";
			// 
			// cmbPersonal
			// 
			this.cmbPersonal.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbPersonal.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbPersonal.DataSource = this.cdvPersonal;
			ultraGridColumn6.Header.VisiblePosition = 21;
			ultraGridColumn6.Hidden = true;
			ultraGridColumn7.Header.VisiblePosition = 59;
			ultraGridColumn7.Hidden = true;
			ultraGridColumn8.Header.VisiblePosition = 19;
			ultraGridColumn8.Hidden = true;
			ultraGridColumn9.Header.VisiblePosition = 20;
			ultraGridColumn9.Hidden = true;
			ultraGridColumn10.Header.VisiblePosition = 18;
			ultraGridColumn10.Hidden = true;
			ultraGridColumn11.Header.VisiblePosition = 22;
			ultraGridColumn11.Hidden = true;
			ultraGridColumn12.Header.VisiblePosition = 36;
			ultraGridColumn12.Hidden = true;
			ultraGridColumn13.Header.VisiblePosition = 25;
			ultraGridColumn13.Hidden = true;
			ultraGridColumn14.Header.VisiblePosition = 0;
			ultraGridColumn14.Width = 339;
			ultraGridColumn15.Header.VisiblePosition = 3;
			ultraGridColumn15.Hidden = true;
			ultraGridColumn16.Header.VisiblePosition = 35;
			ultraGridColumn16.Hidden = true;
			ultraGridColumn17.Header.VisiblePosition = 24;
			ultraGridColumn17.Hidden = true;
			ultraGridColumn18.Header.VisiblePosition = 13;
			ultraGridColumn18.Hidden = true;
			ultraGridColumn19.Header.VisiblePosition = 32;
			ultraGridColumn19.Hidden = true;
			ultraGridColumn20.Header.VisiblePosition = 7;
			ultraGridColumn20.Hidden = true;
			ultraGridColumn21.Header.VisiblePosition = 6;
			ultraGridColumn21.Hidden = true;
			ultraGridColumn22.Header.VisiblePosition = 2;
			ultraGridColumn22.Hidden = true;
			ultraGridColumn23.Header.VisiblePosition = 29;
			ultraGridColumn23.Hidden = true;
			ultraGridColumn24.Header.VisiblePosition = 28;
			ultraGridColumn24.Hidden = true;
			ultraGridColumn25.Header.VisiblePosition = 23;
			ultraGridColumn25.Hidden = true;
			ultraGridColumn26.Header.VisiblePosition = 31;
			ultraGridColumn26.Hidden = true;
			ultraGridColumn27.Header.VisiblePosition = 10;
			ultraGridColumn27.Hidden = true;
			ultraGridColumn28.Header.VisiblePosition = 33;
			ultraGridColumn28.Hidden = true;
			ultraGridColumn29.Header.VisiblePosition = 14;
			ultraGridColumn29.Hidden = true;
			ultraGridColumn30.Header.VisiblePosition = 15;
			ultraGridColumn30.Hidden = true;
			ultraGridColumn31.Header.VisiblePosition = 17;
			ultraGridColumn31.Hidden = true;
			ultraGridColumn32.Header.VisiblePosition = 26;
			ultraGridColumn32.Hidden = true;
			ultraGridColumn33.Header.VisiblePosition = 5;
			ultraGridColumn33.Hidden = true;
			ultraGridColumn34.Header.VisiblePosition = 37;
			ultraGridColumn34.Hidden = true;
			ultraGridColumn35.Header.VisiblePosition = 38;
			ultraGridColumn35.Hidden = true;
			ultraGridColumn36.Header.VisiblePosition = 27;
			ultraGridColumn36.Hidden = true;
			ultraGridColumn37.Header.VisiblePosition = 30;
			ultraGridColumn37.Hidden = true;
			ultraGridColumn38.Header.VisiblePosition = 12;
			ultraGridColumn38.Hidden = true;
			ultraGridColumn39.Header.VisiblePosition = 9;
			ultraGridColumn39.Hidden = true;
			ultraGridColumn40.Header.VisiblePosition = 11;
			ultraGridColumn40.Hidden = true;
			ultraGridColumn41.Header.VisiblePosition = 16;
			ultraGridColumn41.Hidden = true;
			ultraGridColumn42.Header.VisiblePosition = 1;
			ultraGridColumn42.Hidden = true;
			ultraGridColumn43.Header.VisiblePosition = 34;
			ultraGridColumn43.Hidden = true;
			ultraGridColumn44.Header.VisiblePosition = 4;
			ultraGridColumn44.Hidden = true;
			ultraGridColumn45.Header.VisiblePosition = 8;
			ultraGridColumn45.Hidden = true;
			ultraGridColumn46.Header.VisiblePosition = 39;
			ultraGridColumn46.Hidden = true;
			ultraGridColumn47.Header.VisiblePosition = 40;
			ultraGridColumn47.Hidden = true;
			ultraGridColumn48.Header.VisiblePosition = 41;
			ultraGridColumn48.Hidden = true;
			ultraGridColumn49.Header.VisiblePosition = 42;
			ultraGridColumn49.Hidden = true;
			ultraGridColumn50.Header.VisiblePosition = 43;
			ultraGridColumn50.Hidden = true;
			ultraGridColumn51.Header.VisiblePosition = 44;
			ultraGridColumn51.Hidden = true;
			ultraGridColumn52.Header.VisiblePosition = 45;
			ultraGridColumn52.Hidden = true;
			ultraGridColumn53.Header.VisiblePosition = 46;
			ultraGridColumn53.Hidden = true;
			ultraGridColumn54.Header.VisiblePosition = 47;
			ultraGridColumn54.Hidden = true;
			ultraGridColumn55.Header.VisiblePosition = 48;
			ultraGridColumn55.Hidden = true;
			ultraGridColumn56.Header.VisiblePosition = 49;
			ultraGridColumn56.Hidden = true;
			ultraGridColumn57.Header.VisiblePosition = 50;
			ultraGridColumn57.Hidden = true;
			ultraGridColumn58.Header.VisiblePosition = 51;
			ultraGridColumn58.Hidden = true;
			ultraGridColumn59.Header.VisiblePosition = 52;
			ultraGridColumn59.Hidden = true;
			ultraGridColumn60.Header.VisiblePosition = 53;
			ultraGridColumn60.Hidden = true;
			ultraGridColumn61.Header.VisiblePosition = 54;
			ultraGridColumn61.Hidden = true;
			ultraGridColumn62.Header.VisiblePosition = 55;
			ultraGridColumn62.Hidden = true;
			ultraGridColumn63.Header.VisiblePosition = 56;
			ultraGridColumn63.Hidden = true;
			ultraGridColumn64.Header.VisiblePosition = 57;
			ultraGridColumn64.Hidden = true;
			ultraGridColumn65.Header.VisiblePosition = 58;
			ultraGridColumn65.Hidden = true;
			ultraGridColumn66.Header.VisiblePosition = 60;
			ultraGridColumn66.Hidden = true;
			ultraGridColumn67.Header.VisiblePosition = 61;
			ultraGridColumn67.Hidden = true;
			ultraGridColumn68.Header.VisiblePosition = 62;
			ultraGridColumn68.Hidden = true;
			ultraGridColumn69.Header.VisiblePosition = 63;
			ultraGridColumn69.Hidden = true;
			ultraGridColumn70.Header.VisiblePosition = 64;
			ultraGridColumn70.Hidden = true;
			ultraGridColumn71.Header.VisiblePosition = 65;
			ultraGridColumn71.Hidden = true;
			ultraGridColumn72.Header.VisiblePosition = 66;
			ultraGridColumn72.Hidden = true;
			ultraGridColumn73.Header.VisiblePosition = 67;
			ultraGridColumn73.Hidden = true;
			ultraGridColumn74.Header.VisiblePosition = 68;
			ultraGridColumn74.Hidden = true;
			ultraGridColumn75.Header.VisiblePosition = 69;
			ultraGridColumn75.Hidden = true;
			ultraGridColumn76.Header.VisiblePosition = 70;
			ultraGridColumn76.Hidden = true;
			ultraGridColumn77.Header.VisiblePosition = 71;
			ultraGridColumn77.Hidden = true;
			ultraGridColumn78.Header.VisiblePosition = 72;
			ultraGridColumn78.Hidden = true;
			ultraGridColumn79.Header.VisiblePosition = 73;
			ultraGridColumn79.Hidden = true;
			ultraGridColumn80.Header.VisiblePosition = 74;
			ultraGridColumn80.Hidden = true;
			ultraGridColumn81.Header.VisiblePosition = 75;
			ultraGridColumn81.Hidden = true;
			ultraGridColumn82.Header.VisiblePosition = 76;
			ultraGridColumn82.Hidden = true;
			ultraGridColumn83.Header.VisiblePosition = 77;
			ultraGridColumn83.Hidden = true;
			ultraGridColumn84.Header.VisiblePosition = 78;
			ultraGridColumn84.Hidden = true;
			ultraGridColumn85.Header.VisiblePosition = 79;
			ultraGridColumn85.Hidden = true;
			ultraGridColumn86.Header.VisiblePosition = 80;
			ultraGridColumn86.Hidden = true;
			ultraGridColumn87.Header.VisiblePosition = 81;
			ultraGridColumn87.Hidden = true;
			ultraGridColumn88.Header.VisiblePosition = 82;
			ultraGridColumn88.Hidden = true;
			ultraGridColumn89.Header.VisiblePosition = 83;
			ultraGridColumn89.Hidden = true;
			ultraGridColumn90.Header.VisiblePosition = 84;
			ultraGridColumn90.Hidden = true;
			ultraGridColumn91.Header.VisiblePosition = 85;
			ultraGridColumn91.Hidden = true;
			ultraGridColumn92.Header.VisiblePosition = 86;
			ultraGridColumn92.Hidden = true;
			ultraGridColumn93.Header.VisiblePosition = 87;
			ultraGridColumn93.Hidden = true;
			ultraGridColumn94.Header.VisiblePosition = 88;
			ultraGridColumn94.Hidden = true;
			ultraGridColumn95.Header.VisiblePosition = 89;
			ultraGridColumn95.Hidden = true;
			ultraGridColumn96.Header.VisiblePosition = 90;
			ultraGridColumn96.Hidden = true;
			ultraGridColumn97.Header.VisiblePosition = 91;
			ultraGridColumn97.Hidden = true;
			ultraGridColumn98.Header.VisiblePosition = 92;
			ultraGridColumn98.Hidden = true;
			ultraGridColumn99.Header.VisiblePosition = 93;
			ultraGridColumn99.Hidden = true;
			ultraGridColumn100.Header.VisiblePosition = 94;
			ultraGridColumn100.Hidden = true;
			ultraGridBand2.Columns.AddRange(new object[] {
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
																										 ultraGridColumn100});
			this.cmbPersonal.DisplayLayout.BandsSerializer.Add(ultraGridBand2);
			this.cmbPersonal.DisplayMember = "Nombre";
			this.cmbPersonal.Enabled = false;
			this.cmbPersonal.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbPersonal.Location = new System.Drawing.Point(136, 164);
			this.cmbPersonal.Name = "cmbPersonal";
			this.cmbPersonal.Size = new System.Drawing.Size(128, 21);
			this.cmbPersonal.TabIndex = 6;
			this.cmbPersonal.ValueMember = "idPersonal";
			// 
			// chkTodoPersonal
			// 
			this.chkTodoPersonal.Checked = true;
			this.chkTodoPersonal.CheckState = System.Windows.Forms.CheckState.Checked;
			this.chkTodoPersonal.Location = new System.Drawing.Point(24, 164);
			this.chkTodoPersonal.Name = "chkTodoPersonal";
			this.chkTodoPersonal.Size = new System.Drawing.Size(96, 26);
			this.chkTodoPersonal.TabIndex = 7;
			this.chkTodoPersonal.Text = "Todo Personal";
			this.chkTodoPersonal.CheckedChanged += new System.EventHandler(this.chkTodoPersonal_CheckedChanged);
			// 
			// chkTodoTipo
			// 
			this.chkTodoTipo.Checked = true;
			this.chkTodoTipo.CheckState = System.Windows.Forms.CheckState.Checked;
			this.chkTodoTipo.Location = new System.Drawing.Point(24, 121);
			this.chkTodoTipo.Name = "chkTodoTipo";
			this.chkTodoTipo.Size = new System.Drawing.Size(96, 25);
			this.chkTodoTipo.TabIndex = 8;
			this.chkTodoTipo.Text = "Todo Tipo";
			this.chkTodoTipo.CheckedChanged += new System.EventHandler(this.chkTodoTipo_CheckedChanged);
			// 
			// btAceptar
			// 
			this.btAceptar.Location = new System.Drawing.Point(40, 215);
			this.btAceptar.Name = "btAceptar";
			this.btAceptar.Size = new System.Drawing.Size(75, 25);
			this.btAceptar.TabIndex = 9;
			this.btAceptar.Text = "&Aceptar";
			this.btAceptar.Click += new System.EventHandler(this.btAceptar_Click);
			// 
			// btCancelar
			// 
			this.btCancelar.DialogResult = System.Windows.Forms.DialogResult.Cancel;
			this.btCancelar.Location = new System.Drawing.Point(176, 215);
			this.btCancelar.Name = "btCancelar";
			this.btCancelar.Size = new System.Drawing.Size(75, 25);
			this.btCancelar.TabIndex = 10;
			this.btCancelar.Text = "&Cancelar";
			this.btCancelar.Click += new System.EventHandler(this.btCancelar_Click);
			// 
			// errorProvider
			// 
			this.errorProvider.ContainerControl = this;
			// 
			// chkFecha
			// 
			this.chkFecha.Checked = true;
			this.chkFecha.CheckState = System.Windows.Forms.CheckState.Checked;
			this.chkFecha.Location = new System.Drawing.Point(48, 17);
			this.chkFecha.Name = "chkFecha";
			this.chkFecha.Size = new System.Drawing.Size(216, 26);
			this.chkFecha.TabIndex = 11;
			this.chkFecha.Text = "Fecha de Pagos / Fecha de Prestamo";
			// 
			// RolPrestamoImp
			// 
			this.AcceptButton = this.btAceptar;
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 14);
			this.CancelButton = this.btCancelar;
			this.ClientSize = new System.Drawing.Size(292, 264);
			this.Controls.Add(this.chkFecha);
			this.Controls.Add(this.btCancelar);
			this.Controls.Add(this.btAceptar);
			this.Controls.Add(this.chkTodoTipo);
			this.Controls.Add(this.chkTodoPersonal);
			this.Controls.Add(this.cmbPersonal);
			this.Controls.Add(this.cmbTipo);
			this.Controls.Add(this.label2);
			this.Controls.Add(this.label1);
			this.Controls.Add(this.cmbFechaFin);
			this.Controls.Add(this.cmbFechaIni);
			this.Name = "RolPrestamoImp";
			this.Text = "Reportes de Prestamos";
			this.Closing += new System.ComponentModel.CancelEventHandler(this.RolPrestamoImp_Closing);
			this.Load += new System.EventHandler(this.RolPrestamoImp_Load);
			((System.ComponentModel.ISupportInitialize)(this.cmbFechaIni)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbFechaFin)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsTipoPrestamo)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbTipo)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsPersonal)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdvPersonal)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbPersonal)).EndInit();
			this.ResumeLayout(false);

		}
		#endregion

		private void chkTodoTipo_CheckedChanged(object sender, System.EventArgs e)
		{
			cmbTipo.Enabled = !chkTodoTipo.Checked;
			if (!chkTodoTipo.Checked) cmbTipo.Select();
		}

		private void chkTodoPersonal_CheckedChanged(object sender, System.EventArgs e)
		{
			cmbPersonal.Enabled = ! chkTodoPersonal.Checked;
			if (!chkTodoPersonal.Checked) cmbPersonal.Select();
		}

		private bool Verifica()
		{
			bool bOk = true;
			errorProvider.SetError(cmbTipo, "");
			errorProvider.SetError(cmbPersonal, "");
			if (!chkTodoPersonal.Checked)
				if (cmbPersonal.Value == null)
				{
					bOk = false;
					errorProvider.SetError(cmbPersonal, "Ingrese Personal");
				}
			if (!chkTodoTipo.Checked)
				if (cmbTipo.Value == null)
				{
					bOk = false;
					errorProvider.SetError(cmbTipo, "Ingrese Tipo de Prestamo");
				}
			return bOk;
		}

		private string Filtro()
		{
			DateTime dtFechaIni = (DateTime) cmbFechaIni.Value;
			DateTime dtFechaFin = (DateTime) cmbFechaFin.Value;
			dtFechaFin = dtFechaFin.AddDays(1);

			string stTabla = "RolPrestamo.FechaEntrega";
			if (chkFecha.Checked) stTabla = "RolPrestamoDetalle.Fechapropuesta";
			string stFiltro = "{" + stTabla + "} >= #" + dtFechaIni.ToString("MM/dd/yyyy")
				+ "# And {" + stTabla + "} < #" + dtFechaFin.ToString("MM/dd/yyyy") + "#";
			if (!chkTodoPersonal.Checked)
				stFiltro += " And {RolPrestamo.idPersonal} = " + cmbPersonal.Value.ToString();
			if (!chkTodoTipo.Checked)
				stFiltro += " And {RolPrestamo.idRolTipoPrestamo} = " + cmbTipo.Value.ToString();
			return stFiltro;
		}

		private void btAceptar_Click(object sender, System.EventArgs e)
		{
			if (!Verifica()) return;
			Cursor = Cursors.WaitCursor;
      string stFiltro = Filtro();
			Reporte miReporte =  new Reporte("Prestamo.Rpt", stFiltro);
			miReporte.MdiParent = this.MdiParent;
			miReporte.Show();
			Cursor = Cursors.Default;
		}

		private void btCancelar_Click(object sender, System.EventArgs e)
		{
			Close();
		}

		private void RolPrestamoImp_Load(object sender, System.EventArgs e)
		{
		}

		private void RolPrestamoImp_Closing(object sender, System.ComponentModel.CancelEventArgs e)
		{
		}
	}
}
