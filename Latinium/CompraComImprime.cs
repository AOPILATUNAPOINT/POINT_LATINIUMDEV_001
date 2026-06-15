using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using CrystalDecisions.Shared;

namespace Latinium
{
	/// <summary>
	/// Descripción breve de CompraComImprime.
	/// </summary>
	public class CompraComImprime : DevExpress.XtraEditors.XtraForm
	{
		private Infragistics.Win.UltraWinSchedule.UltraCalendarCombo cmbDesde;
		private Infragistics.Win.Misc.UltraLabel ultraLabel1;
		private Infragistics.Win.Misc.UltraLabel ultraLabel2;
		private Infragistics.Win.UltraWinSchedule.UltraCalendarCombo cmbHasta;
		private Infragistics.Win.UltraWinEditors.UltraCheckEditor chkUsuarios;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtCodigo;
		private Infragistics.Win.Misc.UltraButton btAceptar;
		private Infragistics.Win.Misc.UltraButton btCancelar;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource udsOrigen;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbOriginador;
		private Infragistics.Win.UltraWinEditors.UltraCheckEditor chkOriginador;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbVendedor;
		private C1.Data.C1DataView cdvPersonal;
		private C1.Data.C1DataSet cdsPersonal;
		private Infragistics.Win.Misc.UltraButton btGrid;
		private Infragistics.Win.Misc.UltraButton ultraButton1;
		/// <summary>
		/// Variable del diseñador requerida.
		/// </summary>
		private System.ComponentModel.Container components = null;

		public CompraComImprime()
		{
			//
			// Necesario para admitir el Diseñador de Windows Forms
			//
			InitializeComponent();

			//
			// TODO: agregar código de constructor después de llamar a InitializeComponent
			//
		}

private int idC=0;
		public CompraComImprime(int idCompra)
		{
			//
			// Necesario para admitir el Diseñador de Windows Forms
			//
			idC = idCompra;
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
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn1 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Trabajo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn2 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idOrigen");
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand1 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn1 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Trabajo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn2 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idOrigen");
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand2 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Personal", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn3 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idPersonal");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn4 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idAsiDistGrupo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn5 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCentroCosto");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn6 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idDepartamento");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn7 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCargo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn8 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idRolSucursal");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn9 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idRolGrupo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn10 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Jornal");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn11 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn12 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Apellido");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn13 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Codigo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn14 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Jornada");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn15 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Direccion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn16 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Telefono");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn17 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Celular");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn18 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Cedula");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn19 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Afiliacion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn20 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Reserva");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn21 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Renta");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn22 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Iess");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn23 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Sueldo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn24 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Comision");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn25 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Tipo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn26 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("FechaIngreso");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn27 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("FechaNacimiento");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn28 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("FechaSalida");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn29 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Notas");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn30 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("CargasFamiliares");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn31 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idProvincia");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn32 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCiudad");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn33 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Numero");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn34 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("SistemaSalario");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn35 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Cuenta");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn36 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("CodigoSectorial");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn37 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Condicion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn38 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("FechaNovedad");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn39 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Acceso");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn40 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Vendedor");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn41 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Borrar");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn42 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Cobrador");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn43 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Sexo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn44 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Estado_Civil");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn45 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Banco");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn46 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("CuentaCorr");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn47 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idTipoCuenta");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn48 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("CedulaMilitar");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn49 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idFormaPago");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn50 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Quincena");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn51 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("CodIngreso");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn52 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("CodTelefono");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn53 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("TarjEdificio");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn54 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("TarjetaOficina");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn55 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("email");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn56 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idProyecto");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn57 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCuenta");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn58 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Bono");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn59 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Bono1");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn60 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Desc1");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn61 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Desc2");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn62 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("TieneQuincena");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn63 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Fotos");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn64 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Convenio");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn65 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idSubProyecto");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn66 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Sindicato");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn67 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Asociacion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn68 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idRolHorario");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn69 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("FechaContrato");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn70 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Educacion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn71 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("NumeroDias");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn72 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idRelacionT");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn73 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("gSalud");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn74 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("gEducacion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn75 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("gVivienda");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn76 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("gVestimenta");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn77 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("gAlimentacion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn78 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("gTotal");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn79 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("BaseComision");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn80 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("PorcentajeComision");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn81 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("gDesahucio");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn82 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("gDiscapacidad");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn83 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("IngGravOE");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn84 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("DeduccionOE");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn85 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("OtrasRebOE");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn86 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("ImpRentaAsumido");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn87 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("gRebajaIIIEdad");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn88 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("RentaRetenidaOE");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn89 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("RentaRetOE");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn90 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("TipoSangre");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn91 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("RegEntrada");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn92 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("RegSalida");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn93 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idTipoContrato");
			this.cmbDesde = new Infragistics.Win.UltraWinSchedule.UltraCalendarCombo();
			this.ultraLabel1 = new Infragistics.Win.Misc.UltraLabel();
			this.ultraLabel2 = new Infragistics.Win.Misc.UltraLabel();
			this.cmbHasta = new Infragistics.Win.UltraWinSchedule.UltraCalendarCombo();
			this.chkUsuarios = new Infragistics.Win.UltraWinEditors.UltraCheckEditor();
			this.txtCodigo = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.btAceptar = new Infragistics.Win.Misc.UltraButton();
			this.btCancelar = new Infragistics.Win.Misc.UltraButton();
			this.udsOrigen = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.cmbOriginador = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.chkOriginador = new Infragistics.Win.UltraWinEditors.UltraCheckEditor();
			this.cmbVendedor = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.cdvPersonal = new C1.Data.C1DataView();
			this.cdsPersonal = new C1.Data.C1DataSet();
			this.btGrid = new Infragistics.Win.Misc.UltraButton();
			this.ultraButton1 = new Infragistics.Win.Misc.UltraButton();
			((System.ComponentModel.ISupportInitialize)(this.cmbDesde)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbHasta)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtCodigo)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.udsOrigen)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbOriginador)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbVendedor)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdvPersonal)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsPersonal)).BeginInit();
			this.SuspendLayout();
			// 
			// cmbDesde
			// 
			dateButton1.Caption = "Today";
			this.cmbDesde.DateButtons.Add(dateButton1);
			this.cmbDesde.Format = "dd/MMM/yyyy";
			this.cmbDesde.Location = new System.Drawing.Point(112, 32);
			this.cmbDesde.Name = "cmbDesde";
			this.cmbDesde.NonAutoSizeHeight = 23;
			this.cmbDesde.Size = new System.Drawing.Size(120, 21);
			this.cmbDesde.TabIndex = 0;
			this.cmbDesde.Value = new System.DateTime(2011, 5, 13, 0, 0, 0, 0);
			// 
			// ultraLabel1
			// 
			this.ultraLabel1.AutoSize = true;
			this.ultraLabel1.Location = new System.Drawing.Point(48, 32);
			this.ultraLabel1.Name = "ultraLabel1";
			this.ultraLabel1.Size = new System.Drawing.Size(37, 14);
			this.ultraLabel1.TabIndex = 1;
			this.ultraLabel1.Text = "Desde";
			// 
			// ultraLabel2
			// 
			this.ultraLabel2.AutoSize = true;
			this.ultraLabel2.Location = new System.Drawing.Point(48, 64);
			this.ultraLabel2.Name = "ultraLabel2";
			this.ultraLabel2.Size = new System.Drawing.Size(34, 14);
			this.ultraLabel2.TabIndex = 2;
			this.ultraLabel2.Text = "Hasta";
			// 
			// cmbHasta
			// 
			dateButton2.Caption = "Today";
			this.cmbHasta.DateButtons.Add(dateButton2);
			this.cmbHasta.Format = "dd/MMM/yyyy";
			this.cmbHasta.Location = new System.Drawing.Point(112, 64);
			this.cmbHasta.Name = "cmbHasta";
			this.cmbHasta.NonAutoSizeHeight = 23;
			this.cmbHasta.Size = new System.Drawing.Size(120, 21);
			this.cmbHasta.TabIndex = 3;
			this.cmbHasta.Value = new System.DateTime(2011, 5, 13, 0, 0, 0, 0);
			// 
			// chkUsuarios
			// 
			this.chkUsuarios.Checked = true;
			this.chkUsuarios.CheckState = System.Windows.Forms.CheckState.Checked;
			this.chkUsuarios.Location = new System.Drawing.Point(16, 104);
			this.chkUsuarios.Name = "chkUsuarios";
			this.chkUsuarios.Size = new System.Drawing.Size(96, 20);
			this.chkUsuarios.TabIndex = 4;
			this.chkUsuarios.Text = "Todo Codigo";
			this.chkUsuarios.CheckedChanged += new System.EventHandler(this.chkUsuarios_CheckedChanged);
			// 
			// txtCodigo
			// 
			this.txtCodigo.Location = new System.Drawing.Point(136, 104);
			this.txtCodigo.Name = "txtCodigo";
			this.txtCodigo.Size = new System.Drawing.Size(136, 21);
			this.txtCodigo.TabIndex = 5;
			// 
			// btAceptar
			// 
			this.btAceptar.Location = new System.Drawing.Point(24, 176);
			this.btAceptar.Name = "btAceptar";
			this.btAceptar.Size = new System.Drawing.Size(56, 23);
			this.btAceptar.TabIndex = 6;
			this.btAceptar.Text = "&Aceptar";
			this.btAceptar.Click += new System.EventHandler(this.btAceptar_Click);
			// 
			// btCancelar
			// 
			this.btCancelar.Location = new System.Drawing.Point(208, 176);
			this.btCancelar.Name = "btCancelar";
			this.btCancelar.TabIndex = 7;
			this.btCancelar.Text = "&Cancelar";
			this.btCancelar.Click += new System.EventHandler(this.btCancelar_Click);
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
			// cmbOriginador
			// 
			this.cmbOriginador.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbOriginador.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbOriginador.DataMember = "Band 0";
			this.cmbOriginador.DataSource = this.udsOrigen;
			ultraGridColumn1.Header.VisiblePosition = 0;
			ultraGridColumn2.Header.VisiblePosition = 1;
			ultraGridBand1.Columns.AddRange(new object[] {
																										 ultraGridColumn1,
																										 ultraGridColumn2});
			this.cmbOriginador.DisplayLayout.BandsSerializer.Add(ultraGridBand1);
			this.cmbOriginador.DisplayMember = "Trabajo";
			this.cmbOriginador.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbOriginador.Enabled = false;
			this.cmbOriginador.Location = new System.Drawing.Point(136, 136);
			this.cmbOriginador.Name = "cmbOriginador";
			this.cmbOriginador.Size = new System.Drawing.Size(136, 21);
			this.cmbOriginador.TabIndex = 8;
			this.cmbOriginador.ValueMember = "idOrigen";
			// 
			// chkOriginador
			// 
			this.chkOriginador.Checked = true;
			this.chkOriginador.CheckState = System.Windows.Forms.CheckState.Checked;
			this.chkOriginador.Location = new System.Drawing.Point(16, 136);
			this.chkOriginador.Name = "chkOriginador";
			this.chkOriginador.Size = new System.Drawing.Size(112, 20);
			this.chkOriginador.TabIndex = 9;
			this.chkOriginador.Text = "Todo Originador";
			this.chkOriginador.CheckedChanged += new System.EventHandler(this.chkOriginador_CheckedChanged);
			// 
			// cmbVendedor
			// 
			this.cmbVendedor.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbVendedor.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbVendedor.DataSource = this.cdvPersonal;
			ultraGridBand2.AddButtonCaption = "Band 0";
			ultraGridColumn3.Header.VisiblePosition = 0;
			ultraGridColumn3.Hidden = true;
			ultraGridColumn4.Header.VisiblePosition = 1;
			ultraGridColumn4.Hidden = true;
			ultraGridColumn5.Header.VisiblePosition = 2;
			ultraGridColumn5.Hidden = true;
			ultraGridColumn6.Header.VisiblePosition = 3;
			ultraGridColumn6.Hidden = true;
			ultraGridColumn7.Header.VisiblePosition = 4;
			ultraGridColumn7.Hidden = true;
			ultraGridColumn8.Header.VisiblePosition = 5;
			ultraGridColumn8.Hidden = true;
			ultraGridColumn9.Header.VisiblePosition = 6;
			ultraGridColumn9.Hidden = true;
			ultraGridColumn10.Header.VisiblePosition = 7;
			ultraGridColumn10.Hidden = true;
			ultraGridColumn11.Header.VisiblePosition = 8;
			ultraGridColumn12.Header.VisiblePosition = 9;
			ultraGridColumn12.Hidden = true;
			ultraGridColumn13.Header.VisiblePosition = 10;
			ultraGridColumn13.Hidden = true;
			ultraGridColumn14.Header.VisiblePosition = 11;
			ultraGridColumn14.Hidden = true;
			ultraGridColumn15.Header.VisiblePosition = 12;
			ultraGridColumn15.Hidden = true;
			ultraGridColumn16.Header.VisiblePosition = 13;
			ultraGridColumn16.Hidden = true;
			ultraGridColumn17.Header.VisiblePosition = 14;
			ultraGridColumn17.Hidden = true;
			ultraGridColumn18.Header.VisiblePosition = 15;
			ultraGridColumn18.Hidden = true;
			ultraGridColumn19.Header.VisiblePosition = 16;
			ultraGridColumn19.Hidden = true;
			ultraGridColumn20.Header.VisiblePosition = 17;
			ultraGridColumn20.Hidden = true;
			ultraGridColumn21.Header.VisiblePosition = 18;
			ultraGridColumn21.Hidden = true;
			ultraGridColumn22.Header.VisiblePosition = 19;
			ultraGridColumn22.Hidden = true;
			ultraGridColumn23.Header.VisiblePosition = 20;
			ultraGridColumn23.Hidden = true;
			ultraGridColumn24.Header.VisiblePosition = 21;
			ultraGridColumn24.Hidden = true;
			ultraGridColumn25.Header.VisiblePosition = 22;
			ultraGridColumn25.Hidden = true;
			ultraGridColumn26.Header.VisiblePosition = 23;
			ultraGridColumn26.Hidden = true;
			ultraGridColumn27.Header.VisiblePosition = 24;
			ultraGridColumn27.Hidden = true;
			ultraGridColumn28.Header.VisiblePosition = 25;
			ultraGridColumn28.Hidden = true;
			ultraGridColumn29.Header.VisiblePosition = 26;
			ultraGridColumn29.Hidden = true;
			ultraGridColumn30.Header.VisiblePosition = 27;
			ultraGridColumn30.Hidden = true;
			ultraGridColumn31.Header.VisiblePosition = 28;
			ultraGridColumn31.Hidden = true;
			ultraGridColumn32.Header.VisiblePosition = 29;
			ultraGridColumn32.Hidden = true;
			ultraGridColumn33.Header.VisiblePosition = 30;
			ultraGridColumn33.Hidden = true;
			ultraGridColumn34.Header.VisiblePosition = 31;
			ultraGridColumn34.Hidden = true;
			ultraGridColumn35.Header.VisiblePosition = 32;
			ultraGridColumn35.Hidden = true;
			ultraGridColumn36.Header.VisiblePosition = 33;
			ultraGridColumn36.Hidden = true;
			ultraGridColumn37.Header.VisiblePosition = 34;
			ultraGridColumn37.Hidden = true;
			ultraGridColumn38.Header.VisiblePosition = 35;
			ultraGridColumn38.Hidden = true;
			ultraGridColumn39.Header.VisiblePosition = 36;
			ultraGridColumn39.Hidden = true;
			ultraGridColumn40.Header.VisiblePosition = 37;
			ultraGridColumn40.Hidden = true;
			ultraGridColumn41.Header.VisiblePosition = 38;
			ultraGridColumn41.Hidden = true;
			ultraGridColumn42.Header.VisiblePosition = 39;
			ultraGridColumn42.Hidden = true;
			ultraGridColumn43.Header.VisiblePosition = 40;
			ultraGridColumn43.Hidden = true;
			ultraGridColumn44.Header.VisiblePosition = 41;
			ultraGridColumn44.Hidden = true;
			ultraGridColumn45.Header.VisiblePosition = 42;
			ultraGridColumn45.Hidden = true;
			ultraGridColumn46.Header.VisiblePosition = 43;
			ultraGridColumn46.Hidden = true;
			ultraGridColumn47.Header.VisiblePosition = 44;
			ultraGridColumn47.Hidden = true;
			ultraGridColumn48.Header.VisiblePosition = 45;
			ultraGridColumn48.Hidden = true;
			ultraGridColumn49.Header.VisiblePosition = 46;
			ultraGridColumn49.Hidden = true;
			ultraGridColumn50.Header.VisiblePosition = 47;
			ultraGridColumn50.Hidden = true;
			ultraGridColumn51.Header.VisiblePosition = 48;
			ultraGridColumn51.Hidden = true;
			ultraGridColumn52.Header.VisiblePosition = 49;
			ultraGridColumn52.Hidden = true;
			ultraGridColumn53.Header.VisiblePosition = 50;
			ultraGridColumn53.Hidden = true;
			ultraGridColumn54.Header.VisiblePosition = 51;
			ultraGridColumn54.Hidden = true;
			ultraGridColumn55.Header.VisiblePosition = 52;
			ultraGridColumn55.Hidden = true;
			ultraGridColumn56.Header.VisiblePosition = 53;
			ultraGridColumn56.Hidden = true;
			ultraGridColumn57.Header.VisiblePosition = 54;
			ultraGridColumn57.Hidden = true;
			ultraGridColumn58.Header.VisiblePosition = 55;
			ultraGridColumn58.Hidden = true;
			ultraGridColumn59.Header.VisiblePosition = 56;
			ultraGridColumn59.Hidden = true;
			ultraGridColumn60.Header.VisiblePosition = 57;
			ultraGridColumn60.Hidden = true;
			ultraGridColumn61.Header.VisiblePosition = 58;
			ultraGridColumn61.Hidden = true;
			ultraGridColumn62.Header.VisiblePosition = 59;
			ultraGridColumn62.Hidden = true;
			ultraGridColumn63.Header.VisiblePosition = 60;
			ultraGridColumn63.Hidden = true;
			ultraGridColumn64.Header.VisiblePosition = 61;
			ultraGridColumn64.Hidden = true;
			ultraGridColumn65.Header.VisiblePosition = 62;
			ultraGridColumn65.Hidden = true;
			ultraGridColumn66.Header.VisiblePosition = 63;
			ultraGridColumn66.Hidden = true;
			ultraGridColumn67.Header.VisiblePosition = 64;
			ultraGridColumn67.Hidden = true;
			ultraGridColumn68.Header.VisiblePosition = 65;
			ultraGridColumn68.Hidden = true;
			ultraGridColumn69.Header.VisiblePosition = 66;
			ultraGridColumn69.Hidden = true;
			ultraGridColumn70.Header.VisiblePosition = 67;
			ultraGridColumn70.Hidden = true;
			ultraGridColumn71.Header.VisiblePosition = 68;
			ultraGridColumn71.Hidden = true;
			ultraGridColumn72.Header.VisiblePosition = 69;
			ultraGridColumn72.Hidden = true;
			ultraGridColumn73.Header.VisiblePosition = 70;
			ultraGridColumn73.Hidden = true;
			ultraGridColumn74.Header.VisiblePosition = 71;
			ultraGridColumn74.Hidden = true;
			ultraGridColumn75.Header.VisiblePosition = 72;
			ultraGridColumn75.Hidden = true;
			ultraGridColumn76.Header.VisiblePosition = 73;
			ultraGridColumn76.Hidden = true;
			ultraGridColumn77.Header.VisiblePosition = 74;
			ultraGridColumn77.Hidden = true;
			ultraGridColumn78.Header.VisiblePosition = 75;
			ultraGridColumn78.Hidden = true;
			ultraGridColumn79.Header.VisiblePosition = 76;
			ultraGridColumn79.Hidden = true;
			ultraGridColumn80.Header.VisiblePosition = 77;
			ultraGridColumn80.Hidden = true;
			ultraGridColumn81.Header.VisiblePosition = 78;
			ultraGridColumn81.Hidden = true;
			ultraGridColumn82.Header.VisiblePosition = 79;
			ultraGridColumn82.Hidden = true;
			ultraGridColumn83.Header.VisiblePosition = 80;
			ultraGridColumn83.Hidden = true;
			ultraGridColumn84.Header.VisiblePosition = 81;
			ultraGridColumn84.Hidden = true;
			ultraGridColumn85.Header.VisiblePosition = 82;
			ultraGridColumn85.Hidden = true;
			ultraGridColumn86.Header.VisiblePosition = 83;
			ultraGridColumn86.Hidden = true;
			ultraGridColumn87.Header.VisiblePosition = 84;
			ultraGridColumn87.Hidden = true;
			ultraGridColumn88.Header.VisiblePosition = 85;
			ultraGridColumn88.Hidden = true;
			ultraGridColumn89.Header.VisiblePosition = 86;
			ultraGridColumn89.Hidden = true;
			ultraGridColumn90.Header.VisiblePosition = 87;
			ultraGridColumn90.Hidden = true;
			ultraGridColumn91.Header.VisiblePosition = 88;
			ultraGridColumn91.Hidden = true;
			ultraGridColumn92.Header.VisiblePosition = 89;
			ultraGridColumn92.Hidden = true;
			ultraGridColumn93.Header.VisiblePosition = 90;
			ultraGridColumn93.Hidden = true;
			ultraGridBand2.Columns.AddRange(new object[] {
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
																										 ultraGridColumn93});
			this.cmbVendedor.DisplayLayout.BandsSerializer.Add(ultraGridBand2);
			this.cmbVendedor.DisplayMember = "Nombre";
			this.cmbVendedor.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbVendedor.Enabled = false;
			this.cmbVendedor.Location = new System.Drawing.Point(136, 104);
			this.cmbVendedor.Name = "cmbVendedor";
			this.cmbVendedor.Size = new System.Drawing.Size(136, 21);
			this.cmbVendedor.TabIndex = 10;
			this.cmbVendedor.ValueMember = "idPersonal";
			this.cmbVendedor.Visible = false;
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
			// btGrid
			// 
			this.btGrid.Location = new System.Drawing.Point(160, 176);
			this.btGrid.Name = "btGrid";
			this.btGrid.Size = new System.Drawing.Size(40, 23);
			this.btGrid.TabIndex = 11;
			this.btGrid.Text = "&Grid";
			this.btGrid.Visible = false;
			this.btGrid.Click += new System.EventHandler(this.btGrid_Click);
			// 
			// ultraButton1
			// 
			this.ultraButton1.Location = new System.Drawing.Point(88, 176);
			this.ultraButton1.Name = "ultraButton1";
			this.ultraButton1.Size = new System.Drawing.Size(56, 23);
			this.ultraButton1.TabIndex = 12;
			this.ultraButton1.Text = "&Lineas";
			this.ultraButton1.Visible = false;
			this.ultraButton1.Click += new System.EventHandler(this.ultraButton1_Click);
			// 
			// CompraComImprime
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 13);
			this.ClientSize = new System.Drawing.Size(292, 222);
			this.Controls.Add(this.ultraButton1);
			this.Controls.Add(this.btGrid);
			this.Controls.Add(this.cmbVendedor);
			this.Controls.Add(this.chkOriginador);
			this.Controls.Add(this.cmbOriginador);
			this.Controls.Add(this.btCancelar);
			this.Controls.Add(this.btAceptar);
			this.Controls.Add(this.txtCodigo);
			this.Controls.Add(this.chkUsuarios);
			this.Controls.Add(this.cmbHasta);
			this.Controls.Add(this.ultraLabel2);
			this.Controls.Add(this.ultraLabel1);
			this.Controls.Add(this.cmbDesde);
			this.Name = "CompraComImprime";
			this.Text = "Impresion de Comisiones";
			this.Load += new System.EventHandler(this.CompraComImprime_Load);
			((System.ComponentModel.ISupportInitialize)(this.cmbDesde)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbHasta)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtCodigo)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.udsOrigen)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbOriginador)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbVendedor)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdvPersonal)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsPersonal)).EndInit();
			this.ResumeLayout(false);

		}
		#endregion

		private void btCancelar_Click(object sender, System.EventArgs e)
		{
			Close();
		}

		private void btAceptar_Click(object sender, System.EventArgs e)
		{
			string reporte= "Comision.Rpt";
			string nombre = txtCodigo.Text.Trim();	

			if (Funcion.bEjecutaSQL(cdsPersonal, "Exec SeteoGExiste 'FVCLEXIM'")) 				
			{											
			reporte = "ComisionLexim.Rpt";			
			}
				DateTime dtDesde = (DateTime)cmbDesde.Value;
				DateTime dtHasta = (DateTime)cmbHasta.Value;
				string stTitulo = "Comision desde " + dtDesde.ToString("dd/MMM/yyyy")
					+ " hasta " + dtHasta.ToString("dd/MMM/yyyy");
				string stFiltro = "{Compra.Fecha} >= #" + dtDesde.ToString("MM/dd/yyyy")
					+ "# And {Compra.Fecha} < #" + dtHasta.AddDays(1).ToString("MM/dd/yyyy") + "#";

			if (Funcion.bEjecutaSQL(cdsPersonal, "Exec SeteoGExiste 'FVCLEXIM'")) 				
			{			
				if (!chkUsuarios.Checked)
				{
					stFiltro += " And {CompraComision.idVendedor} ="
						+ cmbVendedor.Value.ToString() ;
					stTitulo += " Nombre " + nombre;
				}
			}
			else
			{
				if (!chkUsuarios.Checked)
				{
					stFiltro += " And {CompraComision.Nombre} = '"
						+ nombre + "'";
					stTitulo += " Nombre " + nombre;
				}
			}
				if (!chkOriginador.Checked)
				{
					stFiltro += " And {CompraComision.idOriginador} = "
						+ cmbOriginador.Value.ToString().Trim();
					stTitulo += " Tipo " + cmbOriginador.Value.ToString().Trim();
				}
				Reporte miRep = new Reporte(reporte, stFiltro);
				miRep.Titulo(stTitulo);
				miRep.MdiParent = this.MdiParent;
				miRep.Show();
			

			
		}

		private void CompraComImprime_Load(object sender, System.EventArgs e)
		{
			if (Funcion.bEjecutaSQL(cdsPersonal, "Exec SeteoGExiste 'FVCLEXIM'")) 				
			{
				cmbVendedor.Visible = true;
				txtCodigo.Visible = false;
				ultraButton1.Visible = true;
				btGrid.Visible = true;
				cmbOriginador.Visible = false;
				chkOriginador.Visible = false;
			}
		}

		private void chkOriginador_CheckedChanged(object sender, System.EventArgs e)
		{
			if (!chkOriginador.Checked)
			{
				cmbOriginador.Enabled=true;
			}
			else
			{
			cmbOriginador.Enabled=false;
			}
		}

		private void chkUsuarios_CheckedChanged(object sender, System.EventArgs e)
		{
			if (!chkUsuarios.Checked)
			{
				cmbVendedor.Enabled=true;
			}
			else
			{
				cmbVendedor.Enabled=false;
			}
		}

		

		private void cdsPersonal_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			this.cdsPersonal.Schema.Connections[0].ConnectionString = 
				miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);
		}

		private void btGrid_Click(object sender, System.EventArgs e)
		{
			DateTime dtDesde;
			DateTime dtHasta;
			
			dtDesde = (DateTime) cmbDesde.Value;
			dtHasta = (DateTime) cmbHasta.Value;

			string stProcd = "exec RepComisionesLEXIM '"+dtDesde.ToString("yyyyMMdd")+
												"', '"+dtHasta.ToString("yyyyMMdd")+"', 0";

			CompraComisionGrilla ccg = new CompraComisionGrilla(stProcd);
			ccg.Show();

		}

		private void ultraButton1_Click(object sender, System.EventArgs e)
		{
			string reporte= "Comision.Rpt";
			string nombre = txtCodigo.Text.Trim();	
			if (Funcion.bEjecutaSQL(cdsPersonal, "Exec SeteoGExiste 'FVCLEXIM'")) 				
			{				
				string stSelect = "update CompraComision set Nombre = p.Nombre " + 
					" from CompraComision c inner join Personal p on p.idPersonal = c.idVendedor " +
					" where c.idCompra = "+idC;
				Funcion.EjecutaSQL(cdsPersonal, stSelect, true);
				//MessageBox.Show("Datos Modificados", "Información");

				reporte = "ComisionLeximL.Rpt";
				
				if (!chkUsuarios.Checked)
				{
					nombre = cmbVendedor.Text.Trim();
				}		

				DateTime dtDesde;
				DateTime dtHasta;
			
				dtDesde = (DateTime) cmbDesde.Value;
				dtHasta = (DateTime) cmbHasta.Value;
									
			
				ParameterFields paramFields = new ParameterFields ();
				ParameterField paramField = new ParameterField ();
				ParameterDiscreteValue discreteVal = new ParameterDiscreteValue ();
				paramField.ParameterFieldName = "@FechaDesde";
				discreteVal.Value = dtDesde.ToString("yyyy-MM-dd");
				paramField.CurrentValues.Add (discreteVal);
				paramFields.Add (paramField);

				ParameterField paramField1 = new ParameterField ();
				ParameterDiscreteValue discreteVal1 = new ParameterDiscreteValue ();
				paramField1.ParameterFieldName = "@FechaHasta";
				discreteVal1.Value = dtHasta.ToString("yyyy-MM-dd");
				paramField1.CurrentValues.Add (discreteVal1);
				paramFields.Add (paramField1);

				ParameterField paramField2 = new ParameterField ();
				ParameterDiscreteValue discreteVal2 = new ParameterDiscreteValue ();
				paramField2.ParameterFieldName = "@idVendedor";
				if (chkUsuarios.Checked)
					discreteVal2.Value = (int) 0;
				else
					discreteVal2.Value = (int) cmbVendedor.Value;
				paramField2.CurrentValues.Add (discreteVal2);
				paramFields.Add (paramField2);

				string stReporte = reporte;				
				Reporte miReporte = new Reporte(stReporte, "");
				miReporte.MdiParent = this.MdiParent;
				miReporte.crVista.ParameterFieldInfo = paramFields;
				
				miReporte.Show();					


			}
		
		}
	}
}
