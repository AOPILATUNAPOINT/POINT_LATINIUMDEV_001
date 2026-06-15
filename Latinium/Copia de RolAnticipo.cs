using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;

namespace Latinium
{
	/// <summary>
	/// Descripción breve de RolAnticipo.
	/// </summary>
	public class RolAnticipo : System.Windows.Forms.Form
	{
		private C1.Data.C1DataSet cdsRolAnticipo;
		private C1.Data.C1DataView cdvRolAnticipo;
		private System.Windows.Forms.Label label1;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtConcepto;
		private System.Windows.Forms.Label label2;
		private C1.Data.C1DataSet cdsPersonal;
		private C1.Data.C1DataView cdvPersonal;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtValor;
		private System.Windows.Forms.Label label3;
		private Infragistics.Win.UltraWinSchedule.UltraCalendarCombo cmbFechaEntrega;
		private System.Windows.Forms.Label label4;
		private System.Windows.Forms.Label label5;
		private Infragistics.Win.UltraWinSchedule.UltraCalendarCombo cmbFechaCobro;
		private Infragistics.Win.Misc.UltraButton btEgreso;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtIdAsiento;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtIdRolAnticipo;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbNombre;
		private System.Windows.Forms.ErrorProvider errorProvider;
		private C1.Data.C1DataSet cdsAsiento;
		private BarraDatoSQL.BarraDatoSQL barraDato1;
		/// <summary>
		/// Variable del diseñador requerida.
		/// </summary>
		private System.ComponentModel.Container components = null;

		public RolAnticipo()
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
			Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton dateButton1 = new Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton();
			Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton dateButton2 = new Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton();
			this.cdsRolAnticipo = new C1.Data.C1DataSet();
			this.cdvRolAnticipo = new C1.Data.C1DataView();
			this.label1 = new System.Windows.Forms.Label();
			this.txtConcepto = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.label2 = new System.Windows.Forms.Label();
			this.cmbNombre = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.cdvPersonal = new C1.Data.C1DataView();
			this.cdsPersonal = new C1.Data.C1DataSet();
			this.txtValor = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.label3 = new System.Windows.Forms.Label();
			this.cmbFechaEntrega = new Infragistics.Win.UltraWinSchedule.UltraCalendarCombo();
			this.label4 = new System.Windows.Forms.Label();
			this.label5 = new System.Windows.Forms.Label();
			this.cmbFechaCobro = new Infragistics.Win.UltraWinSchedule.UltraCalendarCombo();
			this.btEgreso = new Infragistics.Win.Misc.UltraButton();
			this.txtIdAsiento = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.txtIdRolAnticipo = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.errorProvider = new System.Windows.Forms.ErrorProvider();
			this.cdsAsiento = new C1.Data.C1DataSet();
			this.barraDato1 = new BarraDatoSQL.BarraDatoSQL();
			((System.ComponentModel.ISupportInitialize)(this.cdsRolAnticipo)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdvRolAnticipo)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtConcepto)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbNombre)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdvPersonal)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsPersonal)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtValor)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbFechaEntrega)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbFechaCobro)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIdAsiento)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIdRolAnticipo)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsAsiento)).BeginInit();
			this.SuspendLayout();
			// 
			// cdsRolAnticipo
			// 
			this.cdsRolAnticipo.BindingContextCtrl = this;
			this.cdsRolAnticipo.DataLibrary = "LibPersonal";
			this.cdsRolAnticipo.DataLibraryUrl = "";
			this.cdsRolAnticipo.DataSetDef = "dsRolGrupo";
			this.cdsRolAnticipo.Locale = new System.Globalization.CultureInfo("es-EC");
			this.cdsRolAnticipo.Name = "cdsRolAnticipo";
			this.cdsRolAnticipo.SchemaClassName = "LibPersonal.DataClass";
			this.cdsRolAnticipo.SchemaDef = null;
			this.cdsRolAnticipo.BeforeFill += new C1.Data.FillEventHandler(this.cdsRolAnticipo_BeforeFill);
			// 
			// cdvRolAnticipo
			// 
			this.cdvRolAnticipo.BindingContextCtrl = this;
			this.cdvRolAnticipo.DataSet = this.cdsRolAnticipo;
			this.cdvRolAnticipo.Sort = "idRolAnticipo";
			this.cdvRolAnticipo.TableName = "";
			this.cdvRolAnticipo.TableViewName = "RolAnticipo";
			// 
			// label1
			// 
			this.label1.AutoSize = true;
			this.label1.Location = new System.Drawing.Point(24, 120);
			this.label1.Name = "label1";
			this.label1.Size = new System.Drawing.Size(56, 16);
			this.label1.TabIndex = 1;
			this.label1.Text = "Concepto:";
			// 
			// txtConcepto
			// 
			this.txtConcepto.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsRolAnticipo, "RolAnticipo.Concepto"));
			this.txtConcepto.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtConcepto.Location = new System.Drawing.Point(88, 120);
			this.txtConcepto.Multiline = true;
			this.txtConcepto.Name = "txtConcepto";
			this.txtConcepto.Size = new System.Drawing.Size(272, 48);
			this.txtConcepto.TabIndex = 2;
			// 
			// label2
			// 
			this.label2.AutoSize = true;
			this.label2.Location = new System.Drawing.Point(24, 24);
			this.label2.Name = "label2";
			this.label2.Size = new System.Drawing.Size(48, 16);
			this.label2.TabIndex = 3;
			this.label2.Text = "Nombre:";
			// 
			// cmbNombre
			// 
			this.cmbNombre.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbNombre.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbNombre.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsRolAnticipo, "RolAnticipo.idPersonal"));
			this.cmbNombre.DataSource = this.cdvPersonal;
			this.cmbNombre.DisplayLayout.AutoFitColumns = true;
			ultraGridColumn1.Header.VisiblePosition = 21;
			ultraGridColumn1.Hidden = true;
			ultraGridColumn2.Header.VisiblePosition = 59;
			ultraGridColumn2.Hidden = true;
			ultraGridColumn2.Width = 30;
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
			ultraGridColumn9.Width = 97;
			ultraGridColumn10.Header.VisiblePosition = 1;
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
			ultraGridColumn17.Header.VisiblePosition = 3;
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
			ultraGridColumn29.Header.VisiblePosition = 36;
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
			ultraGridColumn37.Header.VisiblePosition = 2;
			ultraGridColumn37.Hidden = true;
			ultraGridColumn38.Header.VisiblePosition = 34;
			ultraGridColumn38.Hidden = true;
			ultraGridColumn39.Header.VisiblePosition = 4;
			ultraGridColumn39.Hidden = true;
			ultraGridColumn40.Header.VisiblePosition = 8;
			ultraGridColumn40.Hidden = true;
			ultraGridColumn41.Header.VisiblePosition = 39;
			ultraGridColumn41.Hidden = true;
			ultraGridColumn41.Width = 8;
			ultraGridColumn42.Header.VisiblePosition = 40;
			ultraGridColumn42.Hidden = true;
			ultraGridColumn42.Width = 8;
			ultraGridColumn43.Header.VisiblePosition = 41;
			ultraGridColumn43.Hidden = true;
			ultraGridColumn43.Width = 8;
			ultraGridColumn44.Header.VisiblePosition = 42;
			ultraGridColumn44.Hidden = true;
			ultraGridColumn44.Width = 8;
			ultraGridColumn45.Header.VisiblePosition = 43;
			ultraGridColumn45.Hidden = true;
			ultraGridColumn45.Width = 8;
			ultraGridColumn46.Header.VisiblePosition = 44;
			ultraGridColumn46.Hidden = true;
			ultraGridColumn46.Width = 8;
			ultraGridColumn47.Header.VisiblePosition = 45;
			ultraGridColumn47.Hidden = true;
			ultraGridColumn47.Width = 8;
			ultraGridColumn48.Header.VisiblePosition = 46;
			ultraGridColumn48.Hidden = true;
			ultraGridColumn48.Width = 8;
			ultraGridColumn49.Header.VisiblePosition = 47;
			ultraGridColumn49.Hidden = true;
			ultraGridColumn49.Width = 8;
			ultraGridColumn50.Header.VisiblePosition = 48;
			ultraGridColumn50.Hidden = true;
			ultraGridColumn50.Width = 8;
			ultraGridColumn51.Header.VisiblePosition = 49;
			ultraGridColumn51.Hidden = true;
			ultraGridColumn51.Width = 8;
			ultraGridColumn52.Header.VisiblePosition = 50;
			ultraGridColumn52.Hidden = true;
			ultraGridColumn52.Width = 8;
			ultraGridColumn53.Header.VisiblePosition = 51;
			ultraGridColumn53.Hidden = true;
			ultraGridColumn53.Width = 8;
			ultraGridColumn54.Header.VisiblePosition = 52;
			ultraGridColumn54.Hidden = true;
			ultraGridColumn54.Width = 8;
			ultraGridColumn55.Header.VisiblePosition = 53;
			ultraGridColumn55.Hidden = true;
			ultraGridColumn55.Width = 8;
			ultraGridColumn56.Header.VisiblePosition = 54;
			ultraGridColumn56.Hidden = true;
			ultraGridColumn56.Width = 8;
			ultraGridColumn57.Header.VisiblePosition = 55;
			ultraGridColumn57.Hidden = true;
			ultraGridColumn57.Width = 8;
			ultraGridColumn58.Header.VisiblePosition = 56;
			ultraGridColumn58.Hidden = true;
			ultraGridColumn58.Width = 23;
			ultraGridColumn59.Header.VisiblePosition = 57;
			ultraGridColumn59.Hidden = true;
			ultraGridColumn59.Width = 26;
			ultraGridColumn60.Header.VisiblePosition = 58;
			ultraGridColumn60.Hidden = true;
			ultraGridColumn60.Width = 30;
			ultraGridColumn61.Header.VisiblePosition = 60;
			ultraGridColumn61.Hidden = true;
			ultraGridColumn61.Width = 18;
			ultraGridColumn62.Header.VisiblePosition = 61;
			ultraGridColumn62.Hidden = true;
			ultraGridColumn62.Width = 36;
			ultraGridColumn63.Header.VisiblePosition = 62;
			ultraGridColumn63.Hidden = true;
			ultraGridColumn63.Width = 40;
			ultraGridColumn64.Header.VisiblePosition = 63;
			ultraGridColumn64.Hidden = true;
			ultraGridColumn64.Width = 28;
			ultraGridColumn65.Header.VisiblePosition = 64;
			ultraGridColumn65.Hidden = true;
			ultraGridColumn65.Width = 35;
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
																										 ultraGridColumn65});
			this.cmbNombre.DisplayLayout.BandsSerializer.Add(ultraGridBand1);
			this.cmbNombre.DisplayMember = "Nombre";
			this.cmbNombre.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbNombre.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbNombre.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbNombre.Location = new System.Drawing.Point(88, 24);
			this.cmbNombre.Name = "cmbNombre";
			this.cmbNombre.Size = new System.Drawing.Size(272, 21);
			this.cmbNombre.TabIndex = 5;
			this.cmbNombre.ValueMember = "idPersonal";
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
			// txtValor
			// 
			this.txtValor.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsRolAnticipo, "RolAnticipo.Valor"));
			this.txtValor.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtValor.FormatString = "#,##0.00";
			this.txtValor.Location = new System.Drawing.Point(88, 56);
			this.txtValor.MaskInput = "{LOC}nn,nnn,nnn.nn";
			this.txtValor.Name = "txtValor";
			this.txtValor.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtValor.PromptChar = ' ';
			this.txtValor.Size = new System.Drawing.Size(104, 21);
			this.txtValor.SpinButtonDisplayStyle = Infragistics.Win.ButtonDisplayStyle.Always;
			this.txtValor.TabIndex = 6;
			// 
			// label3
			// 
			this.label3.AutoSize = true;
			this.label3.Location = new System.Drawing.Point(24, 56);
			this.label3.Name = "label3";
			this.label3.Size = new System.Drawing.Size(34, 16);
			this.label3.TabIndex = 7;
			this.label3.Text = "Valor:";
			// 
			// cmbFechaEntrega
			// 
			this.cmbFechaEntrega.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			this.cmbFechaEntrega.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsRolAnticipo, "RolAnticipo.FechaEntrega"));
			dateButton1.Caption = "Today";
			this.cmbFechaEntrega.DateButtons.Add(dateButton1);
			this.cmbFechaEntrega.Format = "dd/MMM/yyyy";
			this.cmbFechaEntrega.Location = new System.Drawing.Point(88, 88);
			this.cmbFechaEntrega.Name = "cmbFechaEntrega";
			this.cmbFechaEntrega.NonAutoSizeHeight = 23;
			this.cmbFechaEntrega.Size = new System.Drawing.Size(104, 21);
			this.cmbFechaEntrega.TabIndex = 8;
			// 
			// label4
			// 
			this.label4.AutoSize = true;
			this.label4.Location = new System.Drawing.Point(24, 88);
			this.label4.Name = "label4";
			this.label4.Size = new System.Drawing.Size(44, 16);
			this.label4.TabIndex = 9;
			this.label4.Text = "Entrega";
			// 
			// label5
			// 
			this.label5.AutoSize = true;
			this.label5.Location = new System.Drawing.Point(208, 88);
			this.label5.Name = "label5";
			this.label5.Size = new System.Drawing.Size(35, 16);
			this.label5.TabIndex = 11;
			this.label5.Text = "Cobro";
			// 
			// cmbFechaCobro
			// 
			this.cmbFechaCobro.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			this.cmbFechaCobro.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsRolAnticipo, "RolAnticipo.FechaCobro"));
			dateButton2.Caption = "Today";
			this.cmbFechaCobro.DateButtons.Add(dateButton2);
			this.cmbFechaCobro.Format = "dd/MMM/yyyy";
			this.cmbFechaCobro.Location = new System.Drawing.Point(256, 88);
			this.cmbFechaCobro.Name = "cmbFechaCobro";
			this.cmbFechaCobro.NonAutoSizeHeight = 23;
			this.cmbFechaCobro.Size = new System.Drawing.Size(104, 21);
			this.cmbFechaCobro.TabIndex = 10;
			// 
			// btEgreso
			// 
			this.btEgreso.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			this.btEgreso.Location = new System.Drawing.Point(328, 192);
			this.btEgreso.Name = "btEgreso";
			this.btEgreso.Size = new System.Drawing.Size(56, 23);
			this.btEgreso.TabIndex = 12;
			this.btEgreso.Text = "&Egreso";
			this.btEgreso.Click += new System.EventHandler(this.btEgreso_Click);
			this.btEgreso.MouseUp += new System.Windows.Forms.MouseEventHandler(this.btEgreso_MouseDown);
			// 
			// txtIdAsiento
			// 
			this.txtIdAsiento.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsRolAnticipo, "RolAnticipo.idAsiento"));
			this.txtIdAsiento.Location = new System.Drawing.Point(328, 168);
			this.txtIdAsiento.Name = "txtIdAsiento";
			this.txtIdAsiento.Size = new System.Drawing.Size(40, 21);
			this.txtIdAsiento.TabIndex = 13;
			// 
			// txtIdRolAnticipo
			// 
			this.txtIdRolAnticipo.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.cdsRolAnticipo, "RolAnticipo.idRolAnticipo"));
			this.txtIdRolAnticipo.Location = new System.Drawing.Point(280, 168);
			this.txtIdRolAnticipo.Name = "txtIdRolAnticipo";
			this.txtIdRolAnticipo.Size = new System.Drawing.Size(40, 21);
			this.txtIdRolAnticipo.TabIndex = 14;
			// 
			// errorProvider
			// 
			this.errorProvider.ContainerControl = this;
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
			// barraDato1
			// 
			this.barraDato1.BarraMovimiento = true;
			this.barraDato1.DataMember = "RolAnticipo";
			this.barraDato1.DataNombreId = "idRolAnticipo";
			this.barraDato1.DataOrden = "";
			this.barraDato1.DataSource = this.cdsRolAnticipo;
			this.barraDato1.DataTabla = "RolAnticipo";
			this.barraDato1.DataTablaHija = "";
			this.barraDato1.Location = new System.Drawing.Point(24, 192);
			this.barraDato1.Name = "barraDato1";
			this.barraDato1.Size = new System.Drawing.Size(304, 24);
			this.barraDato1.TabIndex = 15;
			this.barraDato1.VisibleBorrar = true;
			this.barraDato1.VisibleBuscar = true;
			this.barraDato1.VisibleEditar = true;
			this.barraDato1.VisibleImprimir = true;
			this.barraDato1.VisibleNuevo = true;
			this.barraDato1.Buscar += new System.EventHandler(this.barraDato1_Buscar);
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
			// RolAnticipo
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 13);
			this.ClientSize = new System.Drawing.Size(408, 237);
			this.Controls.Add(this.barraDato1);
			this.Controls.Add(this.txtIdRolAnticipo);
			this.Controls.Add(this.txtIdAsiento);
			this.Controls.Add(this.btEgreso);
			this.Controls.Add(this.label5);
			this.Controls.Add(this.label4);
			this.Controls.Add(this.label3);
			this.Controls.Add(this.label2);
			this.Controls.Add(this.label1);
			this.Controls.Add(this.cmbFechaCobro);
			this.Controls.Add(this.cmbFechaEntrega);
			this.Controls.Add(this.txtValor);
			this.Controls.Add(this.cmbNombre);
			this.Controls.Add(this.txtConcepto);
			this.MaximizeBox = false;
			this.Name = "RolAnticipo";
			this.Text = "Anticipos";
			this.Closing += new System.ComponentModel.CancelEventHandler(this.RolAnticipo_Closing);
			this.Load += new System.EventHandler(this.RolAnticipo_Load);
			((System.ComponentModel.ISupportInitialize)(this.cdsRolAnticipo)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdvRolAnticipo)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtConcepto)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbNombre)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdvPersonal)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsPersonal)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtValor)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbFechaEntrega)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbFechaCobro)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIdAsiento)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIdRolAnticipo)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsAsiento)).EndInit();
			this.ResumeLayout(false);

		}
		#endregion

		private void RolAnticipo_Load(object sender, System.EventArgs e)
		{
			string stAudita = "Exec AuditaCrear 53, 6, 'Rol Anticipo'";
			Funcion.EjecutaSQL(cdsAsiento, stAudita, true);
			Cursor = Cursors.Default;
			this.txtIdAsiento.Width = 0;
			this.txtIdRolAnticipo.Width = 0;
			this.barraDato1.PosUltima();
			this.barraDato1.HabilitaControles(false);
		}

		private void barraDato1_Borrar(object sender, System.EventArgs e)
		{
			if (DialogResult.No == MessageBox.Show("Desea Borrar este Registro",
				"Informacion", MessageBoxButtons.YesNo, MessageBoxIcon.Question,
				MessageBoxDefaultButton.Button2)) return;
			this.barraDato1.BorraRegistro();
		}

		private void barraDato1_Buscar(object sender, System.EventArgs e)
		{
			using (BuscaAnticipo miBusqueda = new BuscaAnticipo())
			{
				if (miBusqueda.ShowDialog() == DialogResult.OK)
				{
					if (miBusqueda.idBusca.Text.Length > 0)
						AnticipoBuscado(int.Parse(miBusqueda.idBusca.Text));
				}
			}
		}

		private void AnticipoBuscado(int id)
		{
			int idx = this.cdvRolAnticipo.Find(id);
			if (idx == -1) 
				MessageBox.Show("IdRolConcepto no encontrado " + id.ToString(), "Error");
			else BindingContext[this.cdsRolAnticipo,"RolAnticipo"].Position = idx;
			this.barraDato1.ProximoId(5);
		}
		private void barraDato1_Copiar(object sender, System.EventArgs e)
		{
		
		}

		private void barraDato1_Crear(object sender, System.EventArgs e)
		{
			this.barraDato1.CrearRegistro();
			C1.Data.C1DataRow drAnticipo = 
				this.cdsRolAnticipo.TableViews["RolAnticipo"].Rows.Find(int.Parse(this.txtIdRolAnticipo.Value.ToString()));
			drAnticipo["FechaEntrega"] = DateTime.Today;
			drAnticipo["FechaCobro"] = DateTime.Today;
		}

		private void barraDato1_Deshacer(object sender, System.EventArgs e)
		{
			this.BorraErrores();
			this.barraDato1.DeshacerRegistro();
		}

		private void barraDato1_Editar(object sender, System.EventArgs e)
		{
			this.barraDato1.EditarRegistro();
		}

		private void BorraErrores()
		{
			this.errorProvider.SetError(this.cmbNombre, "");
			this.errorProvider.SetError(this.txtValor, "");
			this.errorProvider.SetError(this.cmbFechaCobro, "");
			this.errorProvider.SetError(this.cmbFechaEntrega, "");
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
				this.errorProvider.SetError(this.txtValor, "Ingreso Valor del anticipo");
				bOk = false;
			}
			if (this.cmbFechaCobro.Value == System.DBNull.Value)
			{
				this.errorProvider.SetError(this.cmbFechaCobro, "Ingreso Fecha de Cobro");
				bOk = false;
			}
			if (this.cmbFechaEntrega.Value == System.DBNull.Value)
			{
				this.errorProvider.SetError(this.cmbFechaEntrega, "Ingreso Fecha de Entrega");
				bOk = false;
			}
			return bOk;
		}

		private void barraDato1_Grabar(object sender, System.EventArgs e)
		{
			if (!Verifica()) return;
			this.barraDato1.GrabaRegistro();
		}

		private void barraDato1_Impresora(object sender, System.EventArgs e)
		{
		
		}

		private void barraDato1_Imprimir(object sender, System.EventArgs e)
		{
		
		}

		private void barraDato1_Restaurar(object sender, System.EventArgs e)
		{
		
		}

		private void cdsRolAnticipo_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			this.cdsRolAnticipo.Schema.Connections[0].ConnectionString = 
				miFuncion.AccesoBase(MenuLatinium.stDirRoles);
		}

		private void cdsPersonal_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			this.cdsPersonal.Schema.Connections[0].ConnectionString = 
				miFuncion.AccesoBase(MenuLatinium.stDirRoles);
		}

		private void btEgreso_Click(object sender, System.EventArgs e)
		{
			if (int.Parse(this.txtIdAsiento.Value.ToString()) > 0) 
			{
				MessageBox.Show("Asiento ya creado", "Informacion",
					MessageBoxButtons.OK, MessageBoxIcon.Information);
				return;
			}
			if (DialogResult.No == MessageBox.Show("Desea Generar un Egreso de este anticipo?",
				"Confirmacion", MessageBoxButtons.YesNo, MessageBoxIcon.Question))
				return;
      C1.Data.C1DataRow drAsi = cdsAsiento.TableViews["Asiento"].AddNew();
			drAsi["Fecha"] = DateTime.Today;
			drAsi["CodAsiento"] = Contabilidad.NumeroAsiento(cdsAsiento, 3);
			drAsi["Descripcion"] = "Anticipo de " + cmbNombre.Text;
			drAsi["Total"] = txtValor.Value;
			drAsi["TotalH"] = txtValor.Value;
			drAsi["idTipoAsiento"] = 3;
			cdsAsiento.Update();

			int IdAsi = (int) drAsi["idAsiento"];
			C1.Data.C1DataRow drDetAsi = cdsAsiento.TableViews["AsientoDetalle"].AddNew();
			drDetAsi["Debe"] = txtValor.Value;
			drDetAsi["idAsiento"] = IdAsi;
			string stSelect = "Select idCuenta From AsignaCuenta Where idAsignacion = 10";
			drDetAsi["idCuenta"] = Funcion.iEscalarSQL(cdsAsiento, stSelect);
			
			drDetAsi = cdsAsiento.TableViews["AsientoDetalle"].AddNew();
			drDetAsi["idAsiento"] = IdAsi;
			drDetAsi["Haber"] = txtValor.Value;
			stSelect = "Select idCuenta From AsignaCuenta Where idAsignacion = 19";
			drDetAsi["idCuenta"] = Funcion.iEscalarSQL(cdsAsiento, stSelect, false);
			cdsAsiento.Update();

			stSelect = "Update RolAnticipo SET idAsiento = "
				+ IdAsi.ToString() + " Where idRolAnticipo = " 
				+ txtIdRolAnticipo.Value.ToString();
			txtIdAsiento.Value = IdAsi;
			Funcion.EjecutaSQL(cdsRolAnticipo, stSelect);
		}

		private void btEgreso_MouseDown(object sender, System.Windows.Forms.MouseEventArgs e)
		{
			if (e.Button != MouseButtons.Right) return;
			if (int.Parse(this.txtIdAsiento.Value.ToString()) > 0) 
			{
				Cursor = Cursors.WaitCursor;
				Asiento miAsiento = new Asiento(int.Parse(this.txtIdAsiento.Value.ToString()));
				miAsiento.MdiParent = this.MdiParent;
				miAsiento.Show();
				Cursor = Cursors.Default;
			}
			else
			{
				MessageBox.Show("Asiento NO se ha creado", "Informacion",
					MessageBoxButtons.OK, MessageBoxIcon.Information);
				return;
			}
		}

		private void barraDato1_Refresca(object sender, System.EventArgs e)
		{
		
		}

		private void RolAnticipo_Closing(object sender, System.ComponentModel.CancelEventArgs e)
		{
			string stAudita = "Exec AuditaCrear 53, 7, 'Rol Anticipo'";
			Funcion.EjecutaSQL(cdsAsiento, stAudita, true);
		}
	}
}
