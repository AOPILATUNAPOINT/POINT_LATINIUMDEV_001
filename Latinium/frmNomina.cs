using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using System.Data;
using System.Data.SqlClient;

namespace Latinium
{
	/// <summary>
	/// Descripción breve de frmNomina.
	/// </summary>
	public class frmNomina : System.Windows.Forms.Form
	{
		private System.Windows.Forms.Label label8;
		private System.Windows.Forms.Label label7;
		private System.Windows.Forms.Label label1;
		private Infragistics.Win.UltraWinTabControl.UltraTabControl utcNomina;
		private Infragistics.Win.UltraWinTabControl.UltraTabSharedControlsPage ultraTabSharedControlsPage1;
		private Infragistics.Win.UltraWinTabControl.UltraTabPageControl utpEstudios;
		private System.Windows.Forms.Label label22;
		private System.Windows.Forms.Label label21;
		private System.Windows.Forms.Label label10;
		private System.Windows.Forms.Label label64;
		private System.Windows.Forms.Label label69;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtPorcentajeDiscapacidad;
		private System.Windows.Forms.Label label67;
		private System.Windows.Forms.Label label68;
		private System.Windows.Forms.Label label2;
		private System.Windows.Forms.Label label5;
		private System.Windows.Forms.Label label3;
		private System.Windows.Forms.Label label4;
		private System.Windows.Forms.Label label6;
		private System.Windows.Forms.Label label9;
		private System.Windows.Forms.Label label32;
		private System.Windows.Forms.Label label20;
		private System.Windows.Forms.Label label19;
		private System.Windows.Forms.Label label18;
		private System.Windows.Forms.Label label70;
		private System.Windows.Forms.Label label11;
		private System.Windows.Forms.Button btnNuevo;
		private System.Windows.Forms.Button btnConsultar;
		private System.Windows.Forms.Button btnEditar;
		private System.Windows.Forms.Button btnSalir;
		public System.Windows.Forms.Button btnImprimir;
		private System.Windows.Forms.Button btnGuardar;
		private System.Windows.Forms.Button btnCancelar;
		private System.Windows.Forms.Label label33;
		private System.Windows.Forms.Label label12;
		private System.Windows.Forms.Label label13;
		private System.Windows.Forms.Label label14;
		private System.Windows.Forms.Label label63;
		private System.Windows.Forms.Label label15;
		private Infragistics.Win.UltraWinEditors.UltraCheckEditor chkFondoReserva;
		private System.Windows.Forms.Label label16;
		private System.Windows.Forms.Label label17;
		private System.Windows.Forms.Label label23;
		private System.Windows.Forms.Label label24;
		private System.Windows.Forms.Label label25;
		private System.Windows.Forms.Label label26;
		private System.Windows.Forms.Label label50;
		private System.Windows.Forms.Label label49;
		private System.Windows.Forms.Label label48;
		private System.Windows.Forms.Label label47;
		private System.Windows.Forms.Label label46;
		private System.Windows.Forms.Label label27;
		private System.Windows.Forms.Label label28;
		private System.Windows.Forms.Label label29;
		private System.Windows.Forms.Label label30;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtTipoDiscapacidad;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtGastoSalud;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtGastoVivienda;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtAfiliacion;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtGastoEducacion;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtGastoAlimentacion;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtRetencionJudicial;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtGastoVestimenta;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtNEIdentificacion;
		private Infragistics.Win.UltraWinSchedule.UltraCalendarCombo dtpNEFechaNacimiento;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbNEEstadoCivil;
		private Infragistics.Win.UltraWinEditors.UltraCheckEditor chkNEDiscapacidad;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtNETipoDeSangre;
		private Infragistics.Win.UltraWinEditors.UltraOptionSet optNESexo;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtNECargasFamiliares;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbNEParroquia;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbNECanton;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtNEDireccion;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbNEBarrio;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtNEEMail;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtNECelular;
		private Infragistics.Win.UltraWinEditors.UltraOptionSet optNEVivienda;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbNECargo;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbNEDepartamento;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbNECentroCosto;
		private Infragistics.Win.UltraWinEditors.UltraOptionSet optNESistema;
		private Infragistics.Win.UltraWinEditors.UltraOptionSet optNERolesProservicios;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbNECodigoSectorial;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtNESueldo;
		private Infragistics.Win.UltraWinEditors.UltraCheckEditor chkNEAporteIESS;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbNETipoContrato;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtNEMovilizacion;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtNEHorasExtras;
		private Infragistics.Win.UltraWinEditors.UltraOptionSet optNETipoDePago;
		private Infragistics.Win.UltraWinEditors.UltraCheckEditor chkNERetencionJudicial;
		private Infragistics.Win.UltraWinEditors.UltraCheckEditor chkNEImpuestoRenta;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtNEBono;
		private Infragistics.Win.UltraWinEditors.UltraCheckEditor chkNETCC;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtNEComisiones;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtNEResponsabilidades;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtNEFunciones;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtConadis;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtCuenta;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtCodigo;
		private C1.Data.C1DataSet cdsSeteoF;
		private Infragistics.Win.UltraWinTabControl.UltraTabPageControl utcDatosGenerales;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbNEIdentificacion;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource1;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource2;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource3;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbNEProvincia;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource4;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource5;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource6;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource7;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource8;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource9;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbNEBancos;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource10;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource11;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource12;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource13;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource14;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource15;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource16;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource17;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtEstado;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtIdNomina;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource18;
		private System.Windows.Forms.Label label31;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource19;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource20;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource21;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource22;
		private System.Windows.Forms.Label lblEstado;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtNEOrdinarias;
		private System.Windows.Forms.Label label34;
		private System.Windows.Forms.Label label35;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtNEBonos;
		private System.Windows.Forms.Label label36;
		private System.Windows.Forms.Label label37;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtNEApellidoPaterno;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtNEPrimerNombre;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtNEApellidoMaterno;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtNESegundoNombre;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtNEEMailEmpresarial;
		private System.Windows.Forms.Label label38;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtNETiempoDeVivienda;
		private System.Windows.Forms.Label label39;
		private System.Windows.Forms.Label label40;
		private System.Windows.Forms.GroupBox groupBox1;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtNEVive_Con;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtPredial;
		private Infragistics.Win.UltraWinTabControl.UltraTabPageControl utpImagenes;
		private System.Windows.Forms.PictureBox pbFoto;
		private System.Windows.Forms.Label label42;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource23;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtTipoDeLicencia;
		private Infragistics.Win.UltraWinEditors.UltraCheckEditor chkNELicencia;
		private System.Windows.Forms.Label label44;
		private Infragistics.Win.UltraWinSchedule.UltraCalendarCombo dtpFechaCaducidad;
		private System.Windows.Forms.Label label45;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource24;
		private Infragistics.Win.UltraWinEditors.UltraCheckEditor chkNEVendedor;
		private Infragistics.Win.UltraWinEditors.UltraCheckEditor chkNEComisiona;
		private Infragistics.Win.UltraWinEditors.UltraOptionSet optNEMesesAños;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource25;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbNENacionalidad;
		private System.Windows.Forms.Label label54;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource26;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbProvinciaNacimiento;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbCantonNacimiento;
		private Infragistics.Win.UltraWinEditors.UltraCheckEditor chkNERevisado;
		private Infragistics.Win.UltraWinEditors.UltraCheckEditor chkNEAplicaRol;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtQuincena;
		private Infragistics.Win.UltraWinEditors.UltraCheckEditor chkNEQuincena;
		private Infragistics.Win.UltraWinTabControl.UltraTabControl utcVarios;
		private Infragistics.Win.UltraWinTabControl.UltraTabSharedControlsPage ultraTabSharedControlsPage2;
		private Infragistics.Win.UltraWinTabControl.UltraTabPageControl ultraTabPageControl1;
		private Infragistics.Win.UltraWinGrid.UltraGrid uGridExperiencia;
		private Infragistics.Win.UltraWinTabControl.UltraTabPageControl ultraTabPageControl2;
		private Infragistics.Win.UltraWinGrid.UltraDropDown cmbPeriodosTiempo;
		private Infragistics.Win.UltraWinGrid.UltraDropDown cmbEstadoEstudios;
		private Infragistics.Win.UltraWinGrid.UltraDropDown cmbNivel;
		private Infragistics.Win.UltraWinGrid.UltraGrid uGridEstudios;
		private Infragistics.Win.UltraWinTabControl.UltraTabPageControl ultraTabPageControl3;
		private Infragistics.Win.UltraWinGrid.UltraDropDown cmbParentesco;
		private Infragistics.Win.UltraWinGrid.UltraGrid uGridFamiliares;
		private Infragistics.Win.UltraWinTabControl.UltraTabPageControl ultraTabPageControl4;
		private Infragistics.Win.UltraWinGrid.UltraGrid uGridVacaciones;
		private Infragistics.Win.UltraWinTabControl.UltraTabPageControl ultraTabPageControl5;
		private Infragistics.Win.UltraWinGrid.UltraDropDown cmbPersonalAutVerifica;
		private Infragistics.Win.UltraWinGrid.UltraDropDown cmbMotivosDeSalida;
		private Infragistics.Win.UltraWinGrid.UltraGrid uGridEntradasSalidas;
		private Infragistics.Win.UltraWinTabControl.UltraTabPageControl ultraTabPageControl6;
		private Infragistics.Win.UltraWinGrid.UltraGrid uGridEquipos;
		private Infragistics.Win.UltraWinTabControl.UltraTabPageControl ultraTabPageControl7;
		private Infragistics.Win.UltraWinGrid.UltraGrid uGridUniformes;
		private Infragistics.Win.UltraWinTabControl.UltraTabPageControl utpAdicional;
		private Infragistics.Win.UltraWinTabControl.UltraTabPageControl utpFaltas;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtIdPersonal;
		private System.Windows.Forms.Label label51;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbRangos;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource27;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtTotalIngresos;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource28;
		private Infragistics.Win.UltraWinGrid.UltraDropDown cmbTipoDeSalida;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource29;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource30;
		private Infragistics.Win.UltraWinGrid.UltraDropDown cmbTiposDeParto;
		private Infragistics.Win.UltraWinGrid.UltraDropDown cmbTiposDeFaltas;
		private Infragistics.Win.UltraWinGrid.UltraGrid uGridFaltas;
		private Infragistics.Win.UltraWinTabControl.UltraTabPageControl utpOtros;
		private Infragistics.Win.UltraWinGrid.UltraGrid uGridOtros;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource31;
		private Infragistics.Win.UltraWinGrid.UltraDropDown cmbMotivos;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource32;
		private Infragistics.Win.UltraWinEditors.UltraCheckEditor chkDerechoFondosDeReserva;
		private Infragistics.Win.UltraWinTabControl.UltraTabPageControl ultraTabPageControl8;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource33;
		private Infragistics.Win.UltraWinGrid.UltraGrid uGridObservaciones;
		private System.Windows.Forms.Button btnImportar;
		private Infragistics.Win.UltraWinSchedule.UltraCalendarCombo dtFecha;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtNETelefono;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbRegion;
		private System.Windows.Forms.Label label41;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource34;
		private Infragistics.Win.UltraWinEditors.UltraCheckEditor chkDecTerAcum;
		private Infragistics.Win.UltraWinEditors.UltraCheckEditor chkDecCuartoAcum;
		private Infragistics.Win.UltraWinTabControl.UltraTabPageControl utpCargas;
		private Infragistics.Win.UltraWinGrid.UltraGrid uGridCargas;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource35;
		private System.Windows.Forms.Button btnAPP;
		/// <summary>
		/// Variable del diseñador requerida.
		/// </summary>
		private System.ComponentModel.Container components = null;
		public frmNomina()
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
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand1 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn1 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idParentesco");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn2 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Parentesco");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn1 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idParentesco");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn2 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Parentesco");
			Infragistics.Win.Appearance appearance1 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand2 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn3 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idNominaFamiliar");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn4 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idNomina");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn5 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idParentesco", -1, "cmbParentesco");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn6 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn7 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("FechaNacimiento");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn8 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Telefono");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn9 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Celular");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn10 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("LugarTrabajo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn11 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("TiempoDeTrabajo");
			Infragistics.Win.Appearance appearance2 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance3 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance4 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance5 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance6 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn3 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idNominaFamiliar");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn4 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idNomina");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn5 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idParentesco");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn6 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Nombre");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn7 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("FechaNacimiento");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn8 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Telefono");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn9 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Celular");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn10 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("LugarTrabajo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn11 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("TiempoDeTrabajo");
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand3 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn12 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idPeriodosDeTiempo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn13 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Periodo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn12 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idPeriodosDeTiempo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn13 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Periodo");
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand4 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn14 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idEstadoEstudios");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn15 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Estado");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn14 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idEstadoEstudios");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn15 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Estado");
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand5 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn16 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idNivel");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn17 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nivel");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn16 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idNivel");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn17 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Nivel");
			Infragistics.Win.Appearance appearance7 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand6 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn18 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idNominaEstudios");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn19 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idNomina");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn20 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idNivel", -1, "cmbNivel");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn21 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Titulo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn22 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Institucion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn23 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("AñoFinalizacion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn24 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Tiempo");
			Infragistics.Win.Appearance appearance8 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn25 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idPeriodosDeTiempo", -1, "cmbPeriodosTiempo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn26 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idEstadoEstudios", -1, "cmbEstadoEstudios");
			Infragistics.Win.Appearance appearance9 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance10 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance11 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance12 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn18 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idNominaEstudios");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn19 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idNomina");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn20 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idNivel");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn21 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Titulo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn22 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Institucion");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn23 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("AñoFinalizacion");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn24 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Tiempo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn25 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idPeriodosDeTiempo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn26 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idEstadoEstudios");
			Infragistics.Win.Appearance appearance13 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand7 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn27 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idNominaExperiencia");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn28 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idNomina");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn29 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Empresa");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn30 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("PersonasACargo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn31 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Actividades");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn32 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Logros");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn33 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Años");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn34 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idPeriodosDeTiempo", -1, "cmbPeriodosTiempo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn35 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("AfiliadoIESS");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn36 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Aportaciones");
			Infragistics.Win.Appearance appearance14 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn37 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("SinProblemas");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn38 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("ConProblemas");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn39 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("DescripcionProblema");
			Infragistics.Win.Appearance appearance15 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance16 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance17 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance18 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn27 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idNominaExperiencia");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn28 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idNomina");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn29 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Empresa");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn30 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("PersonasACargo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn31 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Actividades");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn32 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Logros");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn33 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Años");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn34 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idPeriodosDeTiempo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn35 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("AfiliadoIESS");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn36 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Aportaciones");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn37 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("SinProblemas");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn38 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("ConProblemas");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn39 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("DescripcionProblema");
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand8 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn40 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idNomina");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn41 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombres");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn40 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idNomina");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn41 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Nombres");
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand9 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn42 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idNominaMotivosDeSalida");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn43 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Motivo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn42 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idNominaMotivosDeSalida");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn43 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Motivo");
			Infragistics.Win.Appearance appearance19 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand10 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn44 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idNominaEntrada");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn45 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idNomina");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn46 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("FechaEntrada");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn47 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("APrueba");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn48 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("FechaTerminaPrueba");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn49 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("NEntrada");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn50 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idVerificadoPor", -1, "cmbPersonalAutVerifica");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn51 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idAutorizadoPor", -1, "cmbPersonalAutVerifica");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn52 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("FechaSalida");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn53 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("NSalida");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn54 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idMotivoSalida", -1, "cmbMotivosDeSalida");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn55 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Descripcion");
			Infragistics.Win.Appearance appearance20 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance21 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance22 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance23 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance24 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn44 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idNominaEntrada");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn45 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idNomina");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn46 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("FechaEntrada");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn47 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("APrueba");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn48 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("FechaTerminaPrueba");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn49 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("NEntrada");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn50 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idVerificadoPor");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn51 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idAutorizadoPor");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn52 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("FechaSalida");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn53 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("NSalida");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn54 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idMotivoSalida");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn55 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Descripcion");
			Infragistics.Win.Appearance appearance25 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand11 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn56 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idNominaEquipos");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn57 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idNomina");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn58 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Articulo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn59 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Serial");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn60 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("FechaEntrega");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn61 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("ObservacionesEntrega");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn62 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("FechaDevolucion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn63 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("ObservacionesDevolucion");
			Infragistics.Win.Appearance appearance26 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance27 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance28 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance29 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn56 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idNominaEquipos");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn57 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idNomina");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn58 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Articulo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn59 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Serial");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn60 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("FechaEntrega");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn61 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("ObservacionesEntrega");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn62 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("FechaDevolucion");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn63 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("ObservacionesDevolucion");
			Infragistics.Win.Appearance appearance30 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand12 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn64 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idNominaUniformes");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn65 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idNomina");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn66 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("FechaEntrega");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn67 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("DetalleEntregado");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn68 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("FechaDevolucion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn69 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("DetalleDevolucion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn70 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Precio");
			Infragistics.Win.Appearance appearance31 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn71 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Cuotas");
			Infragistics.Win.Appearance appearance32 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn72 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("AplicaRol");
			Infragistics.Win.Appearance appearance33 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance34 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance35 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance36 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance37 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn64 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idNominaUniformes");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn65 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idNomina");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn66 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("FechaEntrega");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn67 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("DetalleEntregado");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn68 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("FechaDevolucion");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn69 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("DetalleDevolucion");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn70 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Precio");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn71 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Cuotas");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn72 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("AplicaRol");
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand13 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn73 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idNominaTiposDeParto");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn74 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Descripcion");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn73 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idNominaTiposDeParto");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn74 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Descripcion");
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand14 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn75 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idNominaTiposDeSalidas");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn76 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Descripcion");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn75 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idNominaTiposDeSalidas");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn76 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Descripcion");
			Infragistics.Win.Appearance appearance38 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand15 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn77 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idNominaSalidas");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn78 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idNomina");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn79 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idNominaTiposDeSalidas", -1, "cmbTipoDeSalida");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn80 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idNominaTiposDeParto", -1, "cmbTiposDeParto");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn81 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Desde");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn82 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Hasta");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn83 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Observaciones");
			Infragistics.Win.Appearance appearance39 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance40 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance41 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance42 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance43 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn77 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idNominaSalidas");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn78 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idNomina");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn79 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idNominaTiposDeSalidas");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn80 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idNominaTiposDeParto");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn81 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Desde");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn82 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Hasta");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn83 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Observaciones");
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand16 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn84 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idNominaTiposDeSalidas");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn85 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Descripcion");
			Infragistics.Win.Appearance appearance44 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand17 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn86 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idNominaSalidas");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn87 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idNomina");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn88 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idNominaTiposDeSalidas", -1, "cmbTiposDeFaltas");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn89 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Fecha");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn90 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Observaciones");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn91 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Estado");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn92 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("EstadoM");
			Infragistics.Win.Appearance appearance45 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance46 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance47 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance48 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance49 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn84 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idNominaSalidas");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn85 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idNomina");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn86 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idNominaTiposDeSalidas");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn87 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Fecha");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn88 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Observaciones");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn89 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Estado");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn90 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("EstadoM");
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand18 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn93 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idNominaMotivosOtros");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn94 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Motivo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn95 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Valor");
			Infragistics.Win.Appearance appearance50 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn96 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Observacion");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn91 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idNominaMotivosOtros");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn92 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Motivo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn93 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Valor");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn94 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Observacion");
			Infragistics.Win.Appearance appearance51 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand19 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn97 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idNominaOtros");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn98 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idNomina");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn99 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idNominaMotivosOtros", -1, "cmbMotivos");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn100 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Fecha");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn101 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Valor");
			Infragistics.Win.Appearance appearance52 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn102 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Observacion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn103 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Estado");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn104 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("EstadoM");
			Infragistics.Win.Appearance appearance53 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance54 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance55 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance56 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance57 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn95 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idNominaOtros");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn96 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idNomina");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn97 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idNominaMotivosOtros");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn98 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Fecha");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn99 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Valor");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn100 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Observacion");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn101 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Estado");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn102 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("EstadoM");
			Infragistics.Win.Appearance appearance58 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand20 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn105 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idNominaCargas");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn106 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idNomina");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn107 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn108 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("FechaNacimiento");
			Infragistics.Win.Appearance appearance59 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance60 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance61 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance62 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance63 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance64 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance65 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn103 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idNominaCargas");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn104 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idNomina");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn105 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Nombre");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn106 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("FechaNacimiento");
			Infragistics.Win.Appearance appearance66 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand21 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn109 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idNominaPruebaObservaciones");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn110 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Observaciones");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn111 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Fecha");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn112 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Usuario");
			Infragistics.Win.Appearance appearance67 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance68 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance69 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance70 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance71 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance72 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance73 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn107 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idNominaPruebaObservaciones");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn108 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Observaciones");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn109 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Fecha");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn110 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Usuario");
			Infragistics.Win.Appearance appearance74 = new Infragistics.Win.Appearance();
			Infragistics.Win.ValueListItem valueListItem1 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.ValueListItem valueListItem2 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand22 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn113 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idProvincia");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn114 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn111 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idProvincia");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn112 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Nombre");
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand23 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn115 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idProvincia");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn116 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.Appearance appearance75 = new Infragistics.Win.Appearance();
			Infragistics.Win.ValueListItem valueListItem3 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.ValueListItem valueListItem4 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.ValueListItem valueListItem5 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand24 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn117 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idBarrio");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn118 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn113 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idBarrio");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn114 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Nombre");
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand25 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn119 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idParroquia");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn120 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn115 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idParroquia");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn116 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Nombre");
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand26 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn121 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCanton");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn122 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn117 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idCanton");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn118 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Nombre");
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand27 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn123 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idEdoCivil");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn124 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn119 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idEdoCivil");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn120 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Nombre");
			Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton dateButton1 = new Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton();
			System.Resources.ResourceManager resources = new System.Resources.ResourceManager(typeof(frmNomina));
			Infragistics.Win.Appearance appearance76 = new Infragistics.Win.Appearance();
			Infragistics.Win.ValueListItem valueListItem6 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.ValueListItem valueListItem7 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand28 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn125 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCanton");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn126 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton dateButton2 = new Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand29 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn127 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idRegion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn128 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Region");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn121 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idRegion");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn122 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Region");
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand30 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn129 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCom_Rango");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn130 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Rango");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn123 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idCom_Rango");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn124 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Rango");
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand31 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn131 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idBanco");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn132 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn125 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idBanco");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn126 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Nombre");
			Infragistics.Win.Appearance appearance77 = new Infragistics.Win.Appearance();
			Infragistics.Win.ValueListItem valueListItem8 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.ValueListItem valueListItem9 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.ValueListItem valueListItem10 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.ValueListItem valueListItem11 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand32 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn133 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idTipoContrato");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn134 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("TipoContrato");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn127 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idTipoContrato");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn128 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("TipoContrato");
			Infragistics.Win.Appearance appearance78 = new Infragistics.Win.Appearance();
			Infragistics.Win.ValueListItem valueListItem12 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.Appearance appearance79 = new Infragistics.Win.Appearance();
			Infragistics.Win.ValueListItem valueListItem13 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.Appearance appearance80 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance81 = new Infragistics.Win.Appearance();
			Infragistics.Win.ValueListItem valueListItem14 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.Appearance appearance82 = new Infragistics.Win.Appearance();
			Infragistics.Win.ValueListItem valueListItem15 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.Appearance appearance83 = new Infragistics.Win.Appearance();
			Infragistics.Win.ValueListItem valueListItem16 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.Appearance appearance84 = new Infragistics.Win.Appearance();
			Infragistics.Win.ValueListItem valueListItem17 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.Appearance appearance85 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand33 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn135 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idNominaSectorial");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn136 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Codigo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn137 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Actividad");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn129 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idNominaSectorial");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn130 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Codigo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn131 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Actividad");
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand34 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn138 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCargo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn139 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Descripcion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn140 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Com_IdCargo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn132 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idCargo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn133 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Descripcion");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn134 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Com_IdCargo");
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand35 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn141 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idDepartamento");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn142 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Descripcion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn143 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idModulo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn135 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idDepartamento");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn136 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Descripcion");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn137 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idModulo");
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand36 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn144 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idProyecto");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn145 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn138 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idProyecto");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn139 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Nombre");
			Infragistics.Win.Appearance appearance86 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinTabControl.UltraTab ultraTab1 = new Infragistics.Win.UltraWinTabControl.UltraTab();
			Infragistics.Win.UltraWinTabControl.UltraTab ultraTab2 = new Infragistics.Win.UltraWinTabControl.UltraTab();
			Infragistics.Win.UltraWinTabControl.UltraTab ultraTab3 = new Infragistics.Win.UltraWinTabControl.UltraTab();
			Infragistics.Win.UltraWinTabControl.UltraTab ultraTab4 = new Infragistics.Win.UltraWinTabControl.UltraTab();
			Infragistics.Win.UltraWinTabControl.UltraTab ultraTab5 = new Infragistics.Win.UltraWinTabControl.UltraTab();
			Infragistics.Win.UltraWinTabControl.UltraTab ultraTab6 = new Infragistics.Win.UltraWinTabControl.UltraTab();
			Infragistics.Win.UltraWinTabControl.UltraTab ultraTab7 = new Infragistics.Win.UltraWinTabControl.UltraTab();
			Infragistics.Win.UltraWinTabControl.UltraTab ultraTab8 = new Infragistics.Win.UltraWinTabControl.UltraTab();
			Infragistics.Win.UltraWinTabControl.UltraTab ultraTab9 = new Infragistics.Win.UltraWinTabControl.UltraTab();
			Infragistics.Win.UltraWinTabControl.UltraTab ultraTab10 = new Infragistics.Win.UltraWinTabControl.UltraTab();
			Infragistics.Win.UltraWinTabControl.UltraTab ultraTab11 = new Infragistics.Win.UltraWinTabControl.UltraTab();
			Infragistics.Win.Appearance appearance87 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinTabControl.UltraTab ultraTab12 = new Infragistics.Win.UltraWinTabControl.UltraTab();
			Infragistics.Win.UltraWinTabControl.UltraTab ultraTab13 = new Infragistics.Win.UltraWinTabControl.UltraTab();
			Infragistics.Win.UltraWinTabControl.UltraTab ultraTab14 = new Infragistics.Win.UltraWinTabControl.UltraTab();
			Infragistics.Win.UltraWinTabControl.UltraTab ultraTab15 = new Infragistics.Win.UltraWinTabControl.UltraTab();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand37 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn146 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idTipoDoc");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn147 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn140 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idTipoDoc");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn141 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Nombre");
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand38 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn148 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idNacionalidad");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn149 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn142 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idNacionalidad");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn143 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Nombre");
			Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton dateButton3 = new Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton();
			this.ultraTabPageControl3 = new Infragistics.Win.UltraWinTabControl.UltraTabPageControl();
			this.cmbParentesco = new Infragistics.Win.UltraWinGrid.UltraDropDown();
			this.ultraDataSource18 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.uGridFamiliares = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.ultraDataSource12 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.ultraTabPageControl2 = new Infragistics.Win.UltraWinTabControl.UltraTabPageControl();
			this.cmbPeriodosTiempo = new Infragistics.Win.UltraWinGrid.UltraDropDown();
			this.ultraDataSource23 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.cmbEstadoEstudios = new Infragistics.Win.UltraWinGrid.UltraDropDown();
			this.ultraDataSource20 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.cmbNivel = new Infragistics.Win.UltraWinGrid.UltraDropDown();
			this.ultraDataSource19 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.uGridEstudios = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.ultraDataSource13 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.ultraTabPageControl1 = new Infragistics.Win.UltraWinTabControl.UltraTabPageControl();
			this.uGridExperiencia = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.ultraDataSource14 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.ultraTabPageControl5 = new Infragistics.Win.UltraWinTabControl.UltraTabPageControl();
			this.cmbPersonalAutVerifica = new Infragistics.Win.UltraWinGrid.UltraDropDown();
			this.ultraDataSource24 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.cmbMotivosDeSalida = new Infragistics.Win.UltraWinGrid.UltraDropDown();
			this.ultraDataSource21 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.uGridEntradasSalidas = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.ultraDataSource15 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.ultraTabPageControl6 = new Infragistics.Win.UltraWinTabControl.UltraTabPageControl();
			this.uGridEquipos = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.ultraDataSource17 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.ultraTabPageControl7 = new Infragistics.Win.UltraWinTabControl.UltraTabPageControl();
			this.uGridUniformes = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.ultraDataSource16 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.ultraTabPageControl4 = new Infragistics.Win.UltraWinTabControl.UltraTabPageControl();
			this.cmbTiposDeParto = new Infragistics.Win.UltraWinGrid.UltraDropDown();
			this.ultraDataSource30 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.cmbTipoDeSalida = new Infragistics.Win.UltraWinGrid.UltraDropDown();
			this.ultraDataSource29 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.uGridVacaciones = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.ultraDataSource25 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.utpFaltas = new Infragistics.Win.UltraWinTabControl.UltraTabPageControl();
			this.cmbTiposDeFaltas = new Infragistics.Win.UltraWinGrid.UltraDropDown();
			this.uGridFaltas = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.ultraDataSource28 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.utpOtros = new Infragistics.Win.UltraWinTabControl.UltraTabPageControl();
			this.cmbMotivos = new Infragistics.Win.UltraWinGrid.UltraDropDown();
			this.ultraDataSource32 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.uGridOtros = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.ultraDataSource31 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.utpCargas = new Infragistics.Win.UltraWinTabControl.UltraTabPageControl();
			this.uGridCargas = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.ultraDataSource35 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.ultraTabPageControl8 = new Infragistics.Win.UltraWinTabControl.UltraTabPageControl();
			this.uGridObservaciones = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.ultraDataSource33 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.utcDatosGenerales = new Infragistics.Win.UltraWinTabControl.UltraTabPageControl();
			this.optNEMesesAños = new Infragistics.Win.UltraWinEditors.UltraOptionSet();
			this.label45 = new System.Windows.Forms.Label();
			this.txtNEVive_Con = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.label40 = new System.Windows.Forms.Label();
			this.txtPredial = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.label39 = new System.Windows.Forms.Label();
			this.txtNETiempoDeVivienda = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtNEEMailEmpresarial = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.label38 = new System.Windows.Forms.Label();
			this.cmbNEProvincia = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.ultraDataSource2 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.cmbProvinciaNacimiento = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.label11 = new System.Windows.Forms.Label();
			this.label70 = new System.Windows.Forms.Label();
			this.optNEVivienda = new Infragistics.Win.UltraWinEditors.UltraOptionSet();
			this.txtNEEMail = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.label32 = new System.Windows.Forms.Label();
			this.txtNECelular = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.label20 = new System.Windows.Forms.Label();
			this.txtNETelefono = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.label19 = new System.Windows.Forms.Label();
			this.txtNEDireccion = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.label18 = new System.Windows.Forms.Label();
			this.cmbNEBarrio = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.ultraDataSource6 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.cmbNEParroquia = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.ultraDataSource5 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.label6 = new System.Windows.Forms.Label();
			this.label9 = new System.Windows.Forms.Label();
			this.cmbNECanton = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.ultraDataSource3 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.label3 = new System.Windows.Forms.Label();
			this.label4 = new System.Windows.Forms.Label();
			this.txtNECargasFamiliares = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.label5 = new System.Windows.Forms.Label();
			this.cmbNEEstadoCivil = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.ultraDataSource4 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.label2 = new System.Windows.Forms.Label();
			this.txtConadis = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.label69 = new System.Windows.Forms.Label();
			this.txtPorcentajeDiscapacidad = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.chkNEDiscapacidad = new Infragistics.Win.UltraWinEditors.UltraCheckEditor();
			this.label67 = new System.Windows.Forms.Label();
			this.label68 = new System.Windows.Forms.Label();
			this.label10 = new System.Windows.Forms.Label();
			this.dtpNEFechaNacimiento = new Infragistics.Win.UltraWinSchedule.UltraCalendarCombo();
			this.txtNETipoDeSangre = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.label64 = new System.Windows.Forms.Label();
			this.optNESexo = new Infragistics.Win.UltraWinEditors.UltraOptionSet();
			this.cmbCantonNacimiento = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.label22 = new System.Windows.Forms.Label();
			this.label21 = new System.Windows.Forms.Label();
			this.txtTipoDiscapacidad = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.groupBox1 = new System.Windows.Forms.GroupBox();
			this.chkNELicencia = new Infragistics.Win.UltraWinEditors.UltraCheckEditor();
			this.label44 = new System.Windows.Forms.Label();
			this.dtpFechaCaducidad = new Infragistics.Win.UltraWinSchedule.UltraCalendarCombo();
			this.txtTipoDeLicencia = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.lblEstado = new System.Windows.Forms.Label();
			this.utpEstudios = new Infragistics.Win.UltraWinTabControl.UltraTabPageControl();
			this.chkDecCuartoAcum = new Infragistics.Win.UltraWinEditors.UltraCheckEditor();
			this.cmbRegion = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.ultraDataSource34 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.cmbRangos = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.ultraDataSource27 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.label51 = new System.Windows.Forms.Label();
			this.txtQuincena = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.chkNEQuincena = new Infragistics.Win.UltraWinEditors.UltraCheckEditor();
			this.chkNEAplicaRol = new Infragistics.Win.UltraWinEditors.UltraCheckEditor();
			this.chkDerechoFondosDeReserva = new Infragistics.Win.UltraWinEditors.UltraCheckEditor();
			this.chkNEVendedor = new Infragistics.Win.UltraWinEditors.UltraCheckEditor();
			this.chkNEComisiona = new Infragistics.Win.UltraWinEditors.UltraCheckEditor();
			this.label35 = new System.Windows.Forms.Label();
			this.txtNEBonos = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtNEOrdinarias = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.label34 = new System.Windows.Forms.Label();
			this.cmbNEBancos = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.ultraDataSource10 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.txtNEResponsabilidades = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.label30 = new System.Windows.Forms.Label();
			this.txtNEFunciones = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.label29 = new System.Windows.Forms.Label();
			this.txtNEComisiones = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.label28 = new System.Windows.Forms.Label();
			this.txtNEBono = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.label27 = new System.Windows.Forms.Label();
			this.chkNETCC = new Infragistics.Win.UltraWinEditors.UltraCheckEditor();
			this.txtRetencionJudicial = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.chkNERetencionJudicial = new Infragistics.Win.UltraWinEditors.UltraCheckEditor();
			this.txtGastoVestimenta = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtGastoEducacion = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtGastoAlimentacion = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtGastoSalud = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtGastoVivienda = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.chkNEImpuestoRenta = new Infragistics.Win.UltraWinEditors.UltraCheckEditor();
			this.txtCuenta = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.label25 = new System.Windows.Forms.Label();
			this.label26 = new System.Windows.Forms.Label();
			this.optNETipoDePago = new Infragistics.Win.UltraWinEditors.UltraOptionSet();
			this.txtTotalIngresos = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.label23 = new System.Windows.Forms.Label();
			this.txtNEHorasExtras = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.label24 = new System.Windows.Forms.Label();
			this.txtNEMovilizacion = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.label17 = new System.Windows.Forms.Label();
			this.txtNESueldo = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.label16 = new System.Windows.Forms.Label();
			this.chkDecTerAcum = new Infragistics.Win.UltraWinEditors.UltraCheckEditor();
			this.chkFondoReserva = new Infragistics.Win.UltraWinEditors.UltraCheckEditor();
			this.label15 = new System.Windows.Forms.Label();
			this.chkNEAporteIESS = new Infragistics.Win.UltraWinEditors.UltraCheckEditor();
			this.label63 = new System.Windows.Forms.Label();
			this.cmbNETipoContrato = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.ultraDataSource11 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.optNESistema = new Infragistics.Win.UltraWinEditors.UltraOptionSet();
			this.optNERolesProservicios = new Infragistics.Win.UltraWinEditors.UltraOptionSet();
			this.cmbNECodigoSectorial = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.ultraDataSource22 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.label14 = new System.Windows.Forms.Label();
			this.cmbNECargo = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.ultraDataSource9 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.label13 = new System.Windows.Forms.Label();
			this.cmbNEDepartamento = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.ultraDataSource8 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.label12 = new System.Windows.Forms.Label();
			this.cmbNECentroCosto = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.ultraDataSource7 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.label33 = new System.Windows.Forms.Label();
			this.txtAfiliacion = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.label48 = new System.Windows.Forms.Label();
			this.label49 = new System.Windows.Forms.Label();
			this.label50 = new System.Windows.Forms.Label();
			this.label47 = new System.Windows.Forms.Label();
			this.label46 = new System.Windows.Forms.Label();
			this.label41 = new System.Windows.Forms.Label();
			this.utpAdicional = new Infragistics.Win.UltraWinTabControl.UltraTabPageControl();
			this.utcVarios = new Infragistics.Win.UltraWinTabControl.UltraTabControl();
			this.ultraTabSharedControlsPage2 = new Infragistics.Win.UltraWinTabControl.UltraTabSharedControlsPage();
			this.utpImagenes = new Infragistics.Win.UltraWinTabControl.UltraTabPageControl();
			this.label42 = new System.Windows.Forms.Label();
			this.pbFoto = new System.Windows.Forms.PictureBox();
			this.label8 = new System.Windows.Forms.Label();
			this.txtCodigo = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.label7 = new System.Windows.Forms.Label();
			this.txtNEIdentificacion = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.txtNEApellidoPaterno = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.label1 = new System.Windows.Forms.Label();
			this.utcNomina = new Infragistics.Win.UltraWinTabControl.UltraTabControl();
			this.ultraTabSharedControlsPage1 = new Infragistics.Win.UltraWinTabControl.UltraTabSharedControlsPage();
			this.btnNuevo = new System.Windows.Forms.Button();
			this.btnConsultar = new System.Windows.Forms.Button();
			this.btnEditar = new System.Windows.Forms.Button();
			this.btnSalir = new System.Windows.Forms.Button();
			this.btnImprimir = new System.Windows.Forms.Button();
			this.btnGuardar = new System.Windows.Forms.Button();
			this.btnCancelar = new System.Windows.Forms.Button();
			this.cdsSeteoF = new C1.Data.C1DataSet();
			this.cmbNEIdentificacion = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.ultraDataSource1 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.txtEstado = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtIdNomina = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtNEPrimerNombre = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.label31 = new System.Windows.Forms.Label();
			this.label36 = new System.Windows.Forms.Label();
			this.txtNEApellidoMaterno = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.label37 = new System.Windows.Forms.Label();
			this.txtNESegundoNombre = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.cmbNENacionalidad = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.ultraDataSource26 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.label54 = new System.Windows.Forms.Label();
			this.chkNERevisado = new Infragistics.Win.UltraWinEditors.UltraCheckEditor();
			this.txtIdPersonal = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.btnImportar = new System.Windows.Forms.Button();
			this.dtFecha = new Infragistics.Win.UltraWinSchedule.UltraCalendarCombo();
			this.btnAPP = new System.Windows.Forms.Button();
			this.ultraTabPageControl3.SuspendLayout();
			((System.ComponentModel.ISupportInitialize)(this.cmbParentesco)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource18)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.uGridFamiliares)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource12)).BeginInit();
			this.ultraTabPageControl2.SuspendLayout();
			((System.ComponentModel.ISupportInitialize)(this.cmbPeriodosTiempo)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource23)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbEstadoEstudios)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource20)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbNivel)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource19)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.uGridEstudios)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource13)).BeginInit();
			this.ultraTabPageControl1.SuspendLayout();
			((System.ComponentModel.ISupportInitialize)(this.uGridExperiencia)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource14)).BeginInit();
			this.ultraTabPageControl5.SuspendLayout();
			((System.ComponentModel.ISupportInitialize)(this.cmbPersonalAutVerifica)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource24)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbMotivosDeSalida)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource21)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.uGridEntradasSalidas)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource15)).BeginInit();
			this.ultraTabPageControl6.SuspendLayout();
			((System.ComponentModel.ISupportInitialize)(this.uGridEquipos)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource17)).BeginInit();
			this.ultraTabPageControl7.SuspendLayout();
			((System.ComponentModel.ISupportInitialize)(this.uGridUniformes)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource16)).BeginInit();
			this.ultraTabPageControl4.SuspendLayout();
			((System.ComponentModel.ISupportInitialize)(this.cmbTiposDeParto)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource30)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbTipoDeSalida)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource29)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.uGridVacaciones)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource25)).BeginInit();
			this.utpFaltas.SuspendLayout();
			((System.ComponentModel.ISupportInitialize)(this.cmbTiposDeFaltas)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.uGridFaltas)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource28)).BeginInit();
			this.utpOtros.SuspendLayout();
			((System.ComponentModel.ISupportInitialize)(this.cmbMotivos)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource32)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.uGridOtros)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource31)).BeginInit();
			this.utpCargas.SuspendLayout();
			((System.ComponentModel.ISupportInitialize)(this.uGridCargas)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource35)).BeginInit();
			this.ultraTabPageControl8.SuspendLayout();
			((System.ComponentModel.ISupportInitialize)(this.uGridObservaciones)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource33)).BeginInit();
			this.utcDatosGenerales.SuspendLayout();
			((System.ComponentModel.ISupportInitialize)(this.optNEMesesAños)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNEVive_Con)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtPredial)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNETiempoDeVivienda)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNEEMailEmpresarial)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbNEProvincia)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource2)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbProvinciaNacimiento)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.optNEVivienda)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNEEMail)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNECelular)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNETelefono)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNEDireccion)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbNEBarrio)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource6)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbNEParroquia)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource5)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbNECanton)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource3)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNECargasFamiliares)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbNEEstadoCivil)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource4)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtConadis)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtPorcentajeDiscapacidad)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.dtpNEFechaNacimiento)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNETipoDeSangre)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.optNESexo)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbCantonNacimiento)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtTipoDiscapacidad)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.dtpFechaCaducidad)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtTipoDeLicencia)).BeginInit();
			this.utpEstudios.SuspendLayout();
			((System.ComponentModel.ISupportInitialize)(this.cmbRegion)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource34)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbRangos)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource27)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtQuincena)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNEBonos)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNEOrdinarias)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbNEBancos)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource10)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNEResponsabilidades)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNEFunciones)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNEComisiones)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNEBono)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtRetencionJudicial)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtGastoVestimenta)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtGastoEducacion)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtGastoAlimentacion)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtGastoSalud)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtGastoVivienda)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtCuenta)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.optNETipoDePago)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtTotalIngresos)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNEHorasExtras)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNEMovilizacion)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNESueldo)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbNETipoContrato)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource11)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.optNESistema)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.optNERolesProservicios)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbNECodigoSectorial)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource22)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbNECargo)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource9)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbNEDepartamento)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource8)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbNECentroCosto)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource7)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtAfiliacion)).BeginInit();
			this.utpAdicional.SuspendLayout();
			((System.ComponentModel.ISupportInitialize)(this.utcVarios)).BeginInit();
			this.utcVarios.SuspendLayout();
			this.utpImagenes.SuspendLayout();
			((System.ComponentModel.ISupportInitialize)(this.txtCodigo)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNEIdentificacion)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNEApellidoPaterno)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.utcNomina)).BeginInit();
			this.utcNomina.SuspendLayout();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbNEIdentificacion)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource1)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtEstado)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIdNomina)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNEPrimerNombre)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNEApellidoMaterno)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNESegundoNombre)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbNENacionalidad)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource26)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIdPersonal)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.dtFecha)).BeginInit();
			this.SuspendLayout();
			// 
			// ultraTabPageControl3
			// 
			this.ultraTabPageControl3.Controls.Add(this.cmbParentesco);
			this.ultraTabPageControl3.Controls.Add(this.uGridFamiliares);
			this.ultraTabPageControl3.Location = new System.Drawing.Point(1, 20);
			this.ultraTabPageControl3.Name = "ultraTabPageControl3";
			this.ultraTabPageControl3.Size = new System.Drawing.Size(1120, 283);
			// 
			// cmbParentesco
			// 
			this.cmbParentesco.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbParentesco.DataSource = this.ultraDataSource18;
			ultraGridBand1.AddButtonCaption = "PersParentesco";
			ultraGridColumn1.Header.VisiblePosition = 0;
			ultraGridColumn1.Hidden = true;
			ultraGridColumn1.Width = 150;
			ultraGridColumn2.Header.VisiblePosition = 1;
			ultraGridColumn2.Width = 150;
			ultraGridBand1.Columns.AddRange(new object[] {
																										 ultraGridColumn1,
																										 ultraGridColumn2});
			this.cmbParentesco.DisplayLayout.BandsSerializer.Add(ultraGridBand1);
			this.cmbParentesco.DisplayMember = "Parentesco";
			this.cmbParentesco.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbParentesco.Location = new System.Drawing.Point(56, 136);
			this.cmbParentesco.Name = "cmbParentesco";
			this.cmbParentesco.Size = new System.Drawing.Size(173, 104);
			this.cmbParentesco.TabIndex = 5;
			this.cmbParentesco.ValueMember = "idParentesco";
			this.cmbParentesco.Visible = false;
			// 
			// ultraDataSource18
			// 
			ultraDataColumn1.DataType = typeof(int);
			this.ultraDataSource18.Band.Columns.AddRange(new object[] {
																																	ultraDataColumn1,
																																	ultraDataColumn2});
			// 
			// uGridFamiliares
			// 
			this.uGridFamiliares.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.uGridFamiliares.Cursor = System.Windows.Forms.Cursors.Default;
			this.uGridFamiliares.DataSource = this.ultraDataSource12;
			appearance1.BackColor = System.Drawing.Color.White;
			this.uGridFamiliares.DisplayLayout.Appearance = appearance1;
			this.uGridFamiliares.DisplayLayout.AutoFitColumns = true;
			ultraGridColumn3.Header.VisiblePosition = 0;
			ultraGridColumn3.Hidden = true;
			ultraGridColumn3.Width = 125;
			ultraGridColumn4.Header.VisiblePosition = 1;
			ultraGridColumn4.Hidden = true;
			ultraGridColumn4.Width = 73;
			ultraGridColumn5.Header.Caption = "Parentesco";
			ultraGridColumn5.Header.VisiblePosition = 2;
			ultraGridColumn5.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.DropDownList;
			ultraGridColumn5.Width = 150;
			ultraGridColumn6.Case = Infragistics.Win.UltraWinGrid.Case.Upper;
			ultraGridColumn6.Header.Caption = "Apellidos y Nombres";
			ultraGridColumn6.Header.VisiblePosition = 3;
			ultraGridColumn6.Width = 309;
			ultraGridColumn7.Header.Caption = "F. de Nac";
			ultraGridColumn7.Header.VisiblePosition = 4;
			ultraGridColumn7.Hidden = true;
			ultraGridColumn7.Width = 70;
			ultraGridColumn8.FieldLen = 9;
			ultraGridColumn8.Header.Caption = "Telf. Domicilio";
			ultraGridColumn8.Header.VisiblePosition = 5;
			ultraGridColumn8.Width = 97;
			ultraGridColumn9.FieldLen = 10;
			ultraGridColumn9.Header.VisiblePosition = 6;
			ultraGridColumn9.Width = 108;
			ultraGridColumn10.Case = Infragistics.Win.UltraWinGrid.Case.Upper;
			ultraGridColumn10.Header.Caption = "Lugar De Trabajo";
			ultraGridColumn10.Header.VisiblePosition = 7;
			ultraGridColumn10.Width = 306;
			ultraGridColumn11.Header.Caption = "T. Trabajo (meses)";
			ultraGridColumn11.Header.VisiblePosition = 8;
			ultraGridColumn11.Width = 116;
			ultraGridBand2.Columns.AddRange(new object[] {
																										 ultraGridColumn3,
																										 ultraGridColumn4,
																										 ultraGridColumn5,
																										 ultraGridColumn6,
																										 ultraGridColumn7,
																										 ultraGridColumn8,
																										 ultraGridColumn9,
																										 ultraGridColumn10,
																										 ultraGridColumn11});
			appearance2.FontData.SizeInPoints = 8F;
			ultraGridBand2.Override.HeaderAppearance = appearance2;
			this.uGridFamiliares.DisplayLayout.BandsSerializer.Add(ultraGridBand2);
			this.uGridFamiliares.DisplayLayout.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.TemplateOnBottom;
			this.uGridFamiliares.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.True;
			this.uGridFamiliares.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.True;
			appearance3.BackColor = System.Drawing.Color.Transparent;
			this.uGridFamiliares.DisplayLayout.Override.CardAreaAppearance = appearance3;
			appearance4.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance4.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance4.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance4.FontData.BoldAsString = "True";
			appearance4.FontData.Name = "Arial";
			appearance4.FontData.SizeInPoints = 10F;
			appearance4.ForeColor = System.Drawing.Color.White;
			appearance4.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.uGridFamiliares.DisplayLayout.Override.HeaderAppearance = appearance4;
			appearance5.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance5.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance5.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridFamiliares.DisplayLayout.Override.RowSelectorAppearance = appearance5;
			appearance6.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance6.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance6.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridFamiliares.DisplayLayout.Override.SelectedRowAppearance = appearance6;
			this.uGridFamiliares.Font = new System.Drawing.Font("Tahoma", 7.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.uGridFamiliares.Location = new System.Drawing.Point(7, 5);
			this.uGridFamiliares.Name = "uGridFamiliares";
			this.uGridFamiliares.Size = new System.Drawing.Size(1107, 272);
			this.uGridFamiliares.TabIndex = 4;
			this.uGridFamiliares.AfterRowInsert += new Infragistics.Win.UltraWinGrid.RowEventHandler(this.uGridFamiliares_AfterRowInsert);
			this.uGridFamiliares.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.uGridFamiliares_KeyPress);
			this.uGridFamiliares.BeforeRowsDeleted += new Infragistics.Win.UltraWinGrid.BeforeRowsDeletedEventHandler(this.uGridFamiliares_BeforeRowsDeleted);
			this.uGridFamiliares.KeyDown += new System.Windows.Forms.KeyEventHandler(this.uGridFamiliares_KeyDown);
			this.uGridFamiliares.InitializeLayout += new Infragistics.Win.UltraWinGrid.InitializeLayoutEventHandler(this.uGridFamiliares_InitializeLayout);
			// 
			// ultraDataSource12
			// 
			ultraDataColumn3.DataType = typeof(int);
			ultraDataColumn4.DataType = typeof(int);
			ultraDataColumn5.DataType = typeof(int);
			ultraDataColumn7.DataType = typeof(System.DateTime);
			ultraDataColumn11.DataType = typeof(int);
			this.ultraDataSource12.Band.Columns.AddRange(new object[] {
																																	ultraDataColumn3,
																																	ultraDataColumn4,
																																	ultraDataColumn5,
																																	ultraDataColumn6,
																																	ultraDataColumn7,
																																	ultraDataColumn8,
																																	ultraDataColumn9,
																																	ultraDataColumn10,
																																	ultraDataColumn11});
			// 
			// ultraTabPageControl2
			// 
			this.ultraTabPageControl2.Controls.Add(this.cmbPeriodosTiempo);
			this.ultraTabPageControl2.Controls.Add(this.cmbEstadoEstudios);
			this.ultraTabPageControl2.Controls.Add(this.cmbNivel);
			this.ultraTabPageControl2.Controls.Add(this.uGridEstudios);
			this.ultraTabPageControl2.Location = new System.Drawing.Point(-10000, -10000);
			this.ultraTabPageControl2.Name = "ultraTabPageControl2";
			this.ultraTabPageControl2.Size = new System.Drawing.Size(1121, 283);
			// 
			// cmbPeriodosTiempo
			// 
			this.cmbPeriodosTiempo.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbPeriodosTiempo.DataSource = this.ultraDataSource23;
			ultraGridBand3.AddButtonCaption = "PersParentesco";
			ultraGridColumn12.Header.VisiblePosition = 0;
			ultraGridColumn12.Hidden = true;
			ultraGridColumn13.Header.VisiblePosition = 1;
			ultraGridBand3.Columns.AddRange(new object[] {
																										 ultraGridColumn12,
																										 ultraGridColumn13});
			this.cmbPeriodosTiempo.DisplayLayout.BandsSerializer.Add(ultraGridBand3);
			this.cmbPeriodosTiempo.DisplayMember = "Periodo";
			this.cmbPeriodosTiempo.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbPeriodosTiempo.Location = new System.Drawing.Point(384, 128);
			this.cmbPeriodosTiempo.Name = "cmbPeriodosTiempo";
			this.cmbPeriodosTiempo.Size = new System.Drawing.Size(173, 104);
			this.cmbPeriodosTiempo.TabIndex = 158;
			this.cmbPeriodosTiempo.ValueMember = "idPeriodosDeTiempo";
			this.cmbPeriodosTiempo.Visible = false;
			// 
			// ultraDataSource23
			// 
			ultraDataColumn12.DataType = typeof(int);
			this.ultraDataSource23.Band.Columns.AddRange(new object[] {
																																	ultraDataColumn12,
																																	ultraDataColumn13});
			// 
			// cmbEstadoEstudios
			// 
			this.cmbEstadoEstudios.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbEstadoEstudios.DataSource = this.ultraDataSource20;
			ultraGridColumn14.Header.VisiblePosition = 0;
			ultraGridColumn14.Hidden = true;
			ultraGridColumn15.Header.VisiblePosition = 1;
			ultraGridColumn15.Width = 160;
			ultraGridBand4.Columns.AddRange(new object[] {
																										 ultraGridColumn14,
																										 ultraGridColumn15});
			this.cmbEstadoEstudios.DisplayLayout.BandsSerializer.Add(ultraGridBand4);
			this.cmbEstadoEstudios.DisplayMember = "Estado";
			this.cmbEstadoEstudios.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbEstadoEstudios.Location = new System.Drawing.Point(576, 104);
			this.cmbEstadoEstudios.Name = "cmbEstadoEstudios";
			this.cmbEstadoEstudios.Size = new System.Drawing.Size(184, 95);
			this.cmbEstadoEstudios.TabIndex = 155;
			this.cmbEstadoEstudios.ValueMember = "idEstadoEstudios";
			this.cmbEstadoEstudios.Visible = false;
			// 
			// ultraDataSource20
			// 
			ultraDataColumn14.DataType = typeof(int);
			this.ultraDataSource20.Band.Columns.AddRange(new object[] {
																																	ultraDataColumn14,
																																	ultraDataColumn15});
			// 
			// cmbNivel
			// 
			this.cmbNivel.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbNivel.DataSource = this.ultraDataSource19;
			ultraGridColumn16.Header.VisiblePosition = 0;
			ultraGridColumn16.Hidden = true;
			ultraGridColumn17.Header.VisiblePosition = 1;
			ultraGridColumn17.Width = 150;
			ultraGridBand5.Columns.AddRange(new object[] {
																										 ultraGridColumn16,
																										 ultraGridColumn17});
			this.cmbNivel.DisplayLayout.BandsSerializer.Add(ultraGridBand5);
			this.cmbNivel.DisplayMember = "Nivel";
			this.cmbNivel.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbNivel.Location = new System.Drawing.Point(128, 128);
			this.cmbNivel.Name = "cmbNivel";
			this.cmbNivel.Size = new System.Drawing.Size(192, 95);
			this.cmbNivel.TabIndex = 154;
			this.cmbNivel.ValueMember = "idNivel";
			this.cmbNivel.Visible = false;
			// 
			// ultraDataSource19
			// 
			ultraDataColumn16.DataType = typeof(int);
			this.ultraDataSource19.Band.Columns.AddRange(new object[] {
																																	ultraDataColumn16,
																																	ultraDataColumn17});
			// 
			// uGridEstudios
			// 
			this.uGridEstudios.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.uGridEstudios.Cursor = System.Windows.Forms.Cursors.Default;
			this.uGridEstudios.DataSource = this.ultraDataSource13;
			appearance7.BackColor = System.Drawing.Color.White;
			this.uGridEstudios.DisplayLayout.Appearance = appearance7;
			ultraGridColumn18.Header.VisiblePosition = 0;
			ultraGridColumn18.Hidden = true;
			ultraGridColumn19.Header.VisiblePosition = 1;
			ultraGridColumn19.Hidden = true;
			ultraGridColumn19.Width = 73;
			ultraGridColumn20.Header.Caption = "Nivel de Estudios";
			ultraGridColumn20.Header.VisiblePosition = 2;
			ultraGridColumn20.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.DropDownList;
			ultraGridColumn20.Width = 150;
			ultraGridColumn21.Case = Infragistics.Win.UltraWinGrid.Case.Upper;
			ultraGridColumn21.Header.VisiblePosition = 3;
			ultraGridColumn21.Width = 241;
			ultraGridColumn22.Case = Infragistics.Win.UltraWinGrid.Case.Upper;
			ultraGridColumn22.Header.VisiblePosition = 4;
			ultraGridColumn22.Width = 241;
			ultraGridColumn23.Format = "####";
			ultraGridColumn23.Header.Caption = "Año Fin";
			ultraGridColumn23.Header.VisiblePosition = 8;
			ultraGridColumn23.Width = 65;
			appearance8.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn24.CellAppearance = appearance8;
			ultraGridColumn24.Header.VisiblePosition = 5;
			ultraGridColumn24.Width = 60;
			ultraGridColumn25.Header.Caption = "";
			ultraGridColumn25.Header.VisiblePosition = 6;
			ultraGridColumn25.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.DropDownList;
			ultraGridColumn25.Width = 80;
			ultraGridColumn26.Header.Caption = "Estado Estudios";
			ultraGridColumn26.Header.VisiblePosition = 7;
			ultraGridColumn26.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.DropDownList;
			ultraGridColumn26.Width = 150;
			ultraGridBand6.Columns.AddRange(new object[] {
																										 ultraGridColumn18,
																										 ultraGridColumn19,
																										 ultraGridColumn20,
																										 ultraGridColumn21,
																										 ultraGridColumn22,
																										 ultraGridColumn23,
																										 ultraGridColumn24,
																										 ultraGridColumn25,
																										 ultraGridColumn26});
			this.uGridEstudios.DisplayLayout.BandsSerializer.Add(ultraGridBand6);
			this.uGridEstudios.DisplayLayout.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.TemplateOnBottom;
			this.uGridEstudios.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.True;
			this.uGridEstudios.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.True;
			appearance9.BackColor = System.Drawing.Color.Transparent;
			this.uGridEstudios.DisplayLayout.Override.CardAreaAppearance = appearance9;
			appearance10.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance10.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance10.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance10.FontData.BoldAsString = "True";
			appearance10.FontData.Name = "Arial";
			appearance10.FontData.SizeInPoints = 8F;
			appearance10.ForeColor = System.Drawing.Color.White;
			appearance10.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.uGridEstudios.DisplayLayout.Override.HeaderAppearance = appearance10;
			appearance11.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance11.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance11.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridEstudios.DisplayLayout.Override.RowSelectorAppearance = appearance11;
			appearance12.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance12.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance12.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridEstudios.DisplayLayout.Override.SelectedRowAppearance = appearance12;
			this.uGridEstudios.Font = new System.Drawing.Font("Tahoma", 7.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.uGridEstudios.Location = new System.Drawing.Point(7, 5);
			this.uGridEstudios.Name = "uGridEstudios";
			this.uGridEstudios.Size = new System.Drawing.Size(1108, 272);
			this.uGridEstudios.TabIndex = 161;
			this.uGridEstudios.AfterRowInsert += new Infragistics.Win.UltraWinGrid.RowEventHandler(this.uGridEstudios_AfterRowInsert);
			this.uGridEstudios.BeforeRowsDeleted += new Infragistics.Win.UltraWinGrid.BeforeRowsDeletedEventHandler(this.uGridEstudios_BeforeRowsDeleted);
			this.uGridEstudios.KeyDown += new System.Windows.Forms.KeyEventHandler(this.uGridEstudios_KeyDown);
			this.uGridEstudios.AfterCellUpdate += new Infragistics.Win.UltraWinGrid.CellEventHandler(this.uGridEstudios_AfterCellUpdate);
			// 
			// ultraDataSource13
			// 
			ultraDataColumn18.DataType = typeof(int);
			ultraDataColumn19.DataType = typeof(int);
			ultraDataColumn20.DataType = typeof(int);
			ultraDataColumn23.DataType = typeof(int);
			ultraDataColumn24.DataType = typeof(int);
			ultraDataColumn25.DataType = typeof(int);
			ultraDataColumn25.DefaultValue = 0;
			ultraDataColumn26.DataType = typeof(int);
			ultraDataColumn26.DefaultValue = 0;
			this.ultraDataSource13.Band.Columns.AddRange(new object[] {
																																	ultraDataColumn18,
																																	ultraDataColumn19,
																																	ultraDataColumn20,
																																	ultraDataColumn21,
																																	ultraDataColumn22,
																																	ultraDataColumn23,
																																	ultraDataColumn24,
																																	ultraDataColumn25,
																																	ultraDataColumn26});
			// 
			// ultraTabPageControl1
			// 
			this.ultraTabPageControl1.Controls.Add(this.uGridExperiencia);
			this.ultraTabPageControl1.Location = new System.Drawing.Point(-10000, -10000);
			this.ultraTabPageControl1.Name = "ultraTabPageControl1";
			this.ultraTabPageControl1.Size = new System.Drawing.Size(1121, 283);
			// 
			// uGridExperiencia
			// 
			this.uGridExperiencia.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.uGridExperiencia.Cursor = System.Windows.Forms.Cursors.Default;
			this.uGridExperiencia.DataSource = this.ultraDataSource14;
			appearance13.BackColor = System.Drawing.Color.White;
			this.uGridExperiencia.DisplayLayout.Appearance = appearance13;
			ultraGridColumn27.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn27.Header.VisiblePosition = 0;
			ultraGridColumn27.Hidden = true;
			ultraGridColumn28.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn28.Header.VisiblePosition = 1;
			ultraGridColumn28.Hidden = true;
			ultraGridColumn29.Case = Infragistics.Win.UltraWinGrid.Case.Upper;
			ultraGridColumn29.Header.Caption = "Empresa o Institución";
			ultraGridColumn29.Header.VisiblePosition = 2;
			ultraGridColumn29.Width = 200;
			ultraGridColumn30.Case = Infragistics.Win.UltraWinGrid.Case.Upper;
			ultraGridColumn30.Header.Caption = "Personas a Cargo";
			ultraGridColumn30.Header.VisiblePosition = 3;
			ultraGridColumn30.Width = 105;
			ultraGridColumn31.Case = Infragistics.Win.UltraWinGrid.Case.Upper;
			ultraGridColumn31.Header.VisiblePosition = 4;
			ultraGridColumn31.Width = 200;
			ultraGridColumn32.Case = Infragistics.Win.UltraWinGrid.Case.Upper;
			ultraGridColumn32.Header.VisiblePosition = 5;
			ultraGridColumn32.Width = 200;
			ultraGridColumn33.Header.Caption = "Tiempo";
			ultraGridColumn33.Header.VisiblePosition = 6;
			ultraGridColumn33.Width = 60;
			ultraGridColumn34.Header.Caption = "";
			ultraGridColumn34.Header.VisiblePosition = 7;
			ultraGridColumn34.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.DropDownList;
			ultraGridColumn34.Width = 100;
			ultraGridColumn35.Header.Caption = "Afiliado IESS";
			ultraGridColumn35.Header.VisiblePosition = 8;
			ultraGridColumn35.Width = 90;
			appearance14.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn36.CellAppearance = appearance14;
			ultraGridColumn36.Header.VisiblePosition = 9;
			ultraGridColumn36.Hidden = true;
			ultraGridColumn36.Width = 90;
			ultraGridColumn37.Header.Caption = "Sin Problemas";
			ultraGridColumn37.Header.VisiblePosition = 10;
			ultraGridColumn37.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.CheckBox;
			ultraGridColumn37.Width = 90;
			ultraGridColumn38.Header.Caption = "Con Problemas";
			ultraGridColumn38.Header.VisiblePosition = 11;
			ultraGridColumn38.Width = 90;
			ultraGridColumn39.Case = Infragistics.Win.UltraWinGrid.Case.Upper;
			ultraGridColumn39.Header.Caption = "Descripción Problema";
			ultraGridColumn39.Header.VisiblePosition = 12;
			ultraGridColumn39.Width = 200;
			ultraGridBand7.Columns.AddRange(new object[] {
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
																										 ultraGridColumn39});
			this.uGridExperiencia.DisplayLayout.BandsSerializer.Add(ultraGridBand7);
			this.uGridExperiencia.DisplayLayout.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.TemplateOnBottom;
			this.uGridExperiencia.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.True;
			this.uGridExperiencia.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.True;
			appearance15.BackColor = System.Drawing.Color.Transparent;
			this.uGridExperiencia.DisplayLayout.Override.CardAreaAppearance = appearance15;
			appearance16.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance16.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance16.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance16.FontData.BoldAsString = "True";
			appearance16.FontData.Name = "Arial";
			appearance16.FontData.SizeInPoints = 8F;
			appearance16.ForeColor = System.Drawing.Color.White;
			appearance16.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.uGridExperiencia.DisplayLayout.Override.HeaderAppearance = appearance16;
			appearance17.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance17.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance17.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridExperiencia.DisplayLayout.Override.RowSelectorAppearance = appearance17;
			appearance18.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance18.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance18.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridExperiencia.DisplayLayout.Override.SelectedRowAppearance = appearance18;
			this.uGridExperiencia.Font = new System.Drawing.Font("Tahoma", 7.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.uGridExperiencia.Location = new System.Drawing.Point(7, 5);
			this.uGridExperiencia.Name = "uGridExperiencia";
			this.uGridExperiencia.Size = new System.Drawing.Size(1108, 275);
			this.uGridExperiencia.TabIndex = 4;
			this.uGridExperiencia.AfterRowInsert += new Infragistics.Win.UltraWinGrid.RowEventHandler(this.uGridExperiencia_AfterRowInsert);
			this.uGridExperiencia.BeforeRowsDeleted += new Infragistics.Win.UltraWinGrid.BeforeRowsDeletedEventHandler(this.uGridExperiencia_BeforeRowsDeleted);
			this.uGridExperiencia.KeyDown += new System.Windows.Forms.KeyEventHandler(this.uGridExperiencia_KeyDown);
			this.uGridExperiencia.AfterCellUpdate += new Infragistics.Win.UltraWinGrid.CellEventHandler(this.uGridExperiencia_AfterCellUpdate);
			// 
			// ultraDataSource14
			// 
			ultraDataColumn27.DataType = typeof(int);
			ultraDataColumn28.DataType = typeof(int);
			ultraDataColumn30.DataType = typeof(int);
			ultraDataColumn33.DataType = typeof(int);
			ultraDataColumn33.DefaultValue = 0;
			ultraDataColumn34.DataType = typeof(int);
			ultraDataColumn34.DefaultValue = 0;
			ultraDataColumn35.DataType = typeof(bool);
			ultraDataColumn35.DefaultValue = false;
			ultraDataColumn36.DataType = typeof(int);
			ultraDataColumn37.DataType = typeof(bool);
			ultraDataColumn37.DefaultValue = false;
			ultraDataColumn38.DataType = typeof(bool);
			ultraDataColumn38.DefaultValue = false;
			this.ultraDataSource14.Band.Columns.AddRange(new object[] {
																																	ultraDataColumn27,
																																	ultraDataColumn28,
																																	ultraDataColumn29,
																																	ultraDataColumn30,
																																	ultraDataColumn31,
																																	ultraDataColumn32,
																																	ultraDataColumn33,
																																	ultraDataColumn34,
																																	ultraDataColumn35,
																																	ultraDataColumn36,
																																	ultraDataColumn37,
																																	ultraDataColumn38,
																																	ultraDataColumn39});
			// 
			// ultraTabPageControl5
			// 
			this.ultraTabPageControl5.Controls.Add(this.cmbPersonalAutVerifica);
			this.ultraTabPageControl5.Controls.Add(this.cmbMotivosDeSalida);
			this.ultraTabPageControl5.Controls.Add(this.uGridEntradasSalidas);
			this.ultraTabPageControl5.Location = new System.Drawing.Point(-10000, -10000);
			this.ultraTabPageControl5.Name = "ultraTabPageControl5";
			this.ultraTabPageControl5.Size = new System.Drawing.Size(1121, 283);
			// 
			// cmbPersonalAutVerifica
			// 
			this.cmbPersonalAutVerifica.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbPersonalAutVerifica.DataSource = this.ultraDataSource24;
			ultraGridColumn40.Header.VisiblePosition = 0;
			ultraGridColumn40.Hidden = true;
			ultraGridColumn41.Header.VisiblePosition = 1;
			ultraGridColumn41.Width = 250;
			ultraGridBand8.Columns.AddRange(new object[] {
																										 ultraGridColumn40,
																										 ultraGridColumn41});
			this.cmbPersonalAutVerifica.DisplayLayout.BandsSerializer.Add(ultraGridBand8);
			this.cmbPersonalAutVerifica.DisplayMember = "Nombres";
			this.cmbPersonalAutVerifica.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbPersonalAutVerifica.Location = new System.Drawing.Point(808, 136);
			this.cmbPersonalAutVerifica.Name = "cmbPersonalAutVerifica";
			this.cmbPersonalAutVerifica.Size = new System.Drawing.Size(192, 95);
			this.cmbPersonalAutVerifica.TabIndex = 156;
			this.cmbPersonalAutVerifica.ValueMember = "idNomina";
			this.cmbPersonalAutVerifica.Visible = false;
			// 
			// ultraDataSource24
			// 
			ultraDataColumn40.DataType = typeof(int);
			this.ultraDataSource24.Band.Columns.AddRange(new object[] {
																																	ultraDataColumn40,
																																	ultraDataColumn41});
			// 
			// cmbMotivosDeSalida
			// 
			this.cmbMotivosDeSalida.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbMotivosDeSalida.DataSource = this.ultraDataSource21;
			ultraGridColumn42.Header.VisiblePosition = 0;
			ultraGridColumn42.Hidden = true;
			ultraGridColumn43.Header.VisiblePosition = 1;
			ultraGridColumn43.Width = 150;
			ultraGridBand9.Columns.AddRange(new object[] {
																										 ultraGridColumn42,
																										 ultraGridColumn43});
			this.cmbMotivosDeSalida.DisplayLayout.BandsSerializer.Add(ultraGridBand9);
			this.cmbMotivosDeSalida.DisplayMember = "Motivo";
			this.cmbMotivosDeSalida.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbMotivosDeSalida.Location = new System.Drawing.Point(608, 120);
			this.cmbMotivosDeSalida.Name = "cmbMotivosDeSalida";
			this.cmbMotivosDeSalida.Size = new System.Drawing.Size(192, 95);
			this.cmbMotivosDeSalida.TabIndex = 155;
			this.cmbMotivosDeSalida.ValueMember = "idNominaMotivosDeSalida";
			this.cmbMotivosDeSalida.Visible = false;
			// 
			// ultraDataSource21
			// 
			ultraDataColumn42.DataType = typeof(int);
			this.ultraDataSource21.Band.Columns.AddRange(new object[] {
																																	ultraDataColumn42,
																																	ultraDataColumn43});
			// 
			// uGridEntradasSalidas
			// 
			this.uGridEntradasSalidas.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.uGridEntradasSalidas.Cursor = System.Windows.Forms.Cursors.Default;
			this.uGridEntradasSalidas.DataSource = this.ultraDataSource15;
			appearance19.BackColor = System.Drawing.Color.White;
			this.uGridEntradasSalidas.DisplayLayout.Appearance = appearance19;
			ultraGridColumn44.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn44.Header.VisiblePosition = 0;
			ultraGridColumn44.Hidden = true;
			ultraGridColumn45.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn45.Header.VisiblePosition = 1;
			ultraGridColumn45.Hidden = true;
			ultraGridColumn46.Header.Caption = "Fecha Entrada";
			ultraGridColumn46.Header.VisiblePosition = 2;
			ultraGridColumn46.Width = 80;
			ultraGridColumn47.Header.Caption = "A Prueba";
			ultraGridColumn47.Header.VisiblePosition = 3;
			ultraGridColumn47.Hidden = true;
			ultraGridColumn47.Width = 60;
			ultraGridColumn48.Header.Caption = "Termina Prueba";
			ultraGridColumn48.Header.VisiblePosition = 4;
			ultraGridColumn48.Width = 100;
			ultraGridColumn49.Header.Caption = "Reg. Entrada";
			ultraGridColumn49.Header.VisiblePosition = 5;
			ultraGridColumn49.Width = 90;
			ultraGridColumn50.Case = Infragistics.Win.UltraWinGrid.Case.Upper;
			ultraGridColumn50.Header.Caption = "Verificado Por";
			ultraGridColumn50.Header.VisiblePosition = 6;
			ultraGridColumn50.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.DropDownList;
			ultraGridColumn50.Width = 160;
			ultraGridColumn51.Case = Infragistics.Win.UltraWinGrid.Case.Upper;
			ultraGridColumn51.Header.Caption = "Autorizado Por";
			ultraGridColumn51.Header.VisiblePosition = 7;
			ultraGridColumn51.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.DropDownList;
			ultraGridColumn51.Width = 160;
			ultraGridColumn52.Header.Caption = "Fecha de Salida";
			ultraGridColumn52.Header.VisiblePosition = 8;
			ultraGridColumn52.Width = 100;
			ultraGridColumn53.Header.Caption = "Reg. Sálida";
			ultraGridColumn53.Header.VisiblePosition = 9;
			ultraGridColumn53.Width = 80;
			ultraGridColumn54.Header.Caption = "Mótivo de Sálida";
			ultraGridColumn54.Header.VisiblePosition = 10;
			ultraGridColumn54.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.DropDownList;
			ultraGridColumn54.Width = 120;
			ultraGridColumn55.Case = Infragistics.Win.UltraWinGrid.Case.Upper;
			ultraGridColumn55.Header.Caption = "Recuento de Sálida";
			ultraGridColumn55.Header.VisiblePosition = 11;
			ultraGridColumn55.Width = 300;
			ultraGridBand10.Columns.AddRange(new object[] {
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
																											ultraGridColumn55});
			this.uGridEntradasSalidas.DisplayLayout.BandsSerializer.Add(ultraGridBand10);
			this.uGridEntradasSalidas.DisplayLayout.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.TemplateOnBottom;
			this.uGridEntradasSalidas.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.False;
			this.uGridEntradasSalidas.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.True;
			appearance20.BackColor = System.Drawing.Color.Transparent;
			this.uGridEntradasSalidas.DisplayLayout.Override.CardAreaAppearance = appearance20;
			appearance21.ForeColor = System.Drawing.Color.Black;
			appearance21.ForeColorDisabled = System.Drawing.Color.Black;
			this.uGridEntradasSalidas.DisplayLayout.Override.CellAppearance = appearance21;
			appearance22.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance22.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance22.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance22.FontData.BoldAsString = "True";
			appearance22.FontData.Name = "Arial";
			appearance22.FontData.SizeInPoints = 8F;
			appearance22.ForeColor = System.Drawing.Color.White;
			appearance22.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.uGridEntradasSalidas.DisplayLayout.Override.HeaderAppearance = appearance22;
			appearance23.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance23.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance23.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridEntradasSalidas.DisplayLayout.Override.RowSelectorAppearance = appearance23;
			appearance24.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance24.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance24.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridEntradasSalidas.DisplayLayout.Override.SelectedRowAppearance = appearance24;
			this.uGridEntradasSalidas.Font = new System.Drawing.Font("Tahoma", 7.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.uGridEntradasSalidas.Location = new System.Drawing.Point(7, 5);
			this.uGridEntradasSalidas.Name = "uGridEntradasSalidas";
			this.uGridEntradasSalidas.Size = new System.Drawing.Size(1108, 275);
			this.uGridEntradasSalidas.TabIndex = 4;
			this.uGridEntradasSalidas.AfterRowInsert += new Infragistics.Win.UltraWinGrid.RowEventHandler(this.uGridEntradasSalidas_AfterRowInsert);
			this.uGridEntradasSalidas.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.uGridEntradasSalidas_KeyPress);
			this.uGridEntradasSalidas.BeforeRowsDeleted += new Infragistics.Win.UltraWinGrid.BeforeRowsDeletedEventHandler(this.uGridEntradasSalidas_BeforeRowsDeleted);
			this.uGridEntradasSalidas.KeyDown += new System.Windows.Forms.KeyEventHandler(this.uGridEntradasSalidas_KeyDown);
			this.uGridEntradasSalidas.AfterCellUpdate += new Infragistics.Win.UltraWinGrid.CellEventHandler(this.uGridEntradasSalidas_AfterCellUpdate);
			this.uGridEntradasSalidas.InitializeLayout += new Infragistics.Win.UltraWinGrid.InitializeLayoutEventHandler(this.uGridEntradasSalidas_InitializeLayout);
			// 
			// ultraDataSource15
			// 
			ultraDataColumn44.DataType = typeof(int);
			ultraDataColumn45.DataType = typeof(int);
			ultraDataColumn46.DataType = typeof(System.DateTime);
			ultraDataColumn46.DefaultValue = new System.DateTime(((long)(0)));
			ultraDataColumn47.DataType = typeof(bool);
			ultraDataColumn47.DefaultValue = false;
			ultraDataColumn48.DataType = typeof(System.DateTime);
			ultraDataColumn48.DefaultValue = new System.DateTime(((long)(0)));
			ultraDataColumn50.DataType = typeof(int);
			ultraDataColumn51.DataType = typeof(int);
			ultraDataColumn52.DataType = typeof(System.DateTime);
			ultraDataColumn52.DefaultValue = new System.DateTime(((long)(0)));
			ultraDataColumn54.DataType = typeof(int);
			this.ultraDataSource15.Band.Columns.AddRange(new object[] {
																																	ultraDataColumn44,
																																	ultraDataColumn45,
																																	ultraDataColumn46,
																																	ultraDataColumn47,
																																	ultraDataColumn48,
																																	ultraDataColumn49,
																																	ultraDataColumn50,
																																	ultraDataColumn51,
																																	ultraDataColumn52,
																																	ultraDataColumn53,
																																	ultraDataColumn54,
																																	ultraDataColumn55});
			// 
			// ultraTabPageControl6
			// 
			this.ultraTabPageControl6.Controls.Add(this.uGridEquipos);
			this.ultraTabPageControl6.Location = new System.Drawing.Point(-10000, -10000);
			this.ultraTabPageControl6.Name = "ultraTabPageControl6";
			this.ultraTabPageControl6.Size = new System.Drawing.Size(1121, 283);
			// 
			// uGridEquipos
			// 
			this.uGridEquipos.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.uGridEquipos.Cursor = System.Windows.Forms.Cursors.Default;
			this.uGridEquipos.DataSource = this.ultraDataSource17;
			appearance25.BackColor = System.Drawing.Color.White;
			this.uGridEquipos.DisplayLayout.Appearance = appearance25;
			ultraGridColumn56.Header.VisiblePosition = 0;
			ultraGridColumn56.Hidden = true;
			ultraGridColumn57.Header.VisiblePosition = 1;
			ultraGridColumn57.Hidden = true;
			ultraGridColumn58.Case = Infragistics.Win.UltraWinGrid.Case.Upper;
			ultraGridColumn58.Header.VisiblePosition = 2;
			ultraGridColumn58.Width = 250;
			ultraGridColumn59.Case = Infragistics.Win.UltraWinGrid.Case.Upper;
			ultraGridColumn59.Header.VisiblePosition = 3;
			ultraGridColumn59.Width = 150;
			ultraGridColumn60.Header.Caption = "F. Entrega";
			ultraGridColumn60.Header.VisiblePosition = 4;
			ultraGridColumn60.Width = 80;
			ultraGridColumn61.Case = Infragistics.Win.UltraWinGrid.Case.Upper;
			ultraGridColumn61.Header.Caption = "Notas de Entrega";
			ultraGridColumn61.Header.VisiblePosition = 5;
			ultraGridColumn61.Width = 350;
			ultraGridColumn62.Header.Caption = "F. Devolucion";
			ultraGridColumn62.Header.VisiblePosition = 6;
			ultraGridColumn62.Width = 100;
			ultraGridColumn63.Case = Infragistics.Win.UltraWinGrid.Case.Upper;
			ultraGridColumn63.Header.Caption = "Notas de Devolución";
			ultraGridColumn63.Header.VisiblePosition = 7;
			ultraGridColumn63.Width = 350;
			ultraGridBand11.Columns.AddRange(new object[] {
																											ultraGridColumn56,
																											ultraGridColumn57,
																											ultraGridColumn58,
																											ultraGridColumn59,
																											ultraGridColumn60,
																											ultraGridColumn61,
																											ultraGridColumn62,
																											ultraGridColumn63});
			this.uGridEquipos.DisplayLayout.BandsSerializer.Add(ultraGridBand11);
			this.uGridEquipos.DisplayLayout.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.TemplateOnBottom;
			this.uGridEquipos.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.True;
			this.uGridEquipos.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.True;
			appearance26.BackColor = System.Drawing.Color.Transparent;
			this.uGridEquipos.DisplayLayout.Override.CardAreaAppearance = appearance26;
			appearance27.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance27.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance27.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance27.FontData.BoldAsString = "True";
			appearance27.FontData.Name = "Arial";
			appearance27.FontData.SizeInPoints = 8F;
			appearance27.ForeColor = System.Drawing.Color.White;
			appearance27.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.uGridEquipos.DisplayLayout.Override.HeaderAppearance = appearance27;
			appearance28.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance28.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance28.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridEquipos.DisplayLayout.Override.RowSelectorAppearance = appearance28;
			appearance29.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance29.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance29.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridEquipos.DisplayLayout.Override.SelectedRowAppearance = appearance29;
			this.uGridEquipos.Font = new System.Drawing.Font("Tahoma", 7.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.uGridEquipos.Location = new System.Drawing.Point(8, 4);
			this.uGridEquipos.Name = "uGridEquipos";
			this.uGridEquipos.Size = new System.Drawing.Size(1107, 275);
			this.uGridEquipos.TabIndex = 3;
			this.uGridEquipos.AfterRowInsert += new Infragistics.Win.UltraWinGrid.RowEventHandler(this.uGridEquipos_AfterRowInsert);
			this.uGridEquipos.BeforeRowsDeleted += new Infragistics.Win.UltraWinGrid.BeforeRowsDeletedEventHandler(this.uGridEquipos_BeforeRowsDeleted);
			this.uGridEquipos.KeyDown += new System.Windows.Forms.KeyEventHandler(this.uGridEquipos_KeyDown);
			// 
			// ultraDataSource17
			// 
			ultraDataColumn56.DataType = typeof(int);
			ultraDataColumn57.DataType = typeof(int);
			ultraDataColumn60.DataType = typeof(System.DateTime);
			ultraDataColumn62.DataType = typeof(System.DateTime);
			this.ultraDataSource17.Band.Columns.AddRange(new object[] {
																																	ultraDataColumn56,
																																	ultraDataColumn57,
																																	ultraDataColumn58,
																																	ultraDataColumn59,
																																	ultraDataColumn60,
																																	ultraDataColumn61,
																																	ultraDataColumn62,
																																	ultraDataColumn63});
			// 
			// ultraTabPageControl7
			// 
			this.ultraTabPageControl7.Controls.Add(this.uGridUniformes);
			this.ultraTabPageControl7.Location = new System.Drawing.Point(-10000, -10000);
			this.ultraTabPageControl7.Name = "ultraTabPageControl7";
			this.ultraTabPageControl7.Size = new System.Drawing.Size(1121, 283);
			// 
			// uGridUniformes
			// 
			this.uGridUniformes.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.uGridUniformes.Cursor = System.Windows.Forms.Cursors.Default;
			this.uGridUniformes.DataSource = this.ultraDataSource16;
			appearance30.BackColor = System.Drawing.Color.White;
			this.uGridUniformes.DisplayLayout.Appearance = appearance30;
			this.uGridUniformes.DisplayLayout.AutoFitColumns = true;
			ultraGridBand12.AddButtonCaption = "PersFamiliar";
			ultraGridColumn64.Header.VisiblePosition = 0;
			ultraGridColumn64.Hidden = true;
			ultraGridColumn64.Width = 154;
			ultraGridColumn65.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn65.Header.VisiblePosition = 1;
			ultraGridColumn65.Hidden = true;
			ultraGridColumn66.Header.Caption = "Entregado";
			ultraGridColumn66.Header.VisiblePosition = 2;
			ultraGridColumn66.Width = 82;
			ultraGridColumn67.Case = Infragistics.Win.UltraWinGrid.Case.Upper;
			ultraGridColumn67.Header.Caption = "Detalle Entregado";
			ultraGridColumn67.Header.VisiblePosition = 3;
			ultraGridColumn67.Width = 374;
			ultraGridColumn68.Header.Caption = "Devuelto";
			ultraGridColumn68.Header.VisiblePosition = 4;
			ultraGridColumn68.Width = 70;
			ultraGridColumn69.Case = Infragistics.Win.UltraWinGrid.Case.Upper;
			ultraGridColumn69.Header.Caption = "Detalle Devolucion";
			ultraGridColumn69.Header.VisiblePosition = 5;
			ultraGridColumn69.Width = 383;
			appearance31.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn70.CellAppearance = appearance31;
			ultraGridColumn70.Formula = "#,##0.00";
			ultraGridColumn70.Header.VisiblePosition = 6;
			ultraGridColumn70.Width = 65;
			appearance32.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn71.CellAppearance = appearance32;
			ultraGridColumn71.Formula = "#0";
			ultraGridColumn71.Header.VisiblePosition = 7;
			ultraGridColumn71.Width = 48;
			ultraGridColumn72.Header.Caption = "Aplica Rol";
			ultraGridColumn72.Header.VisiblePosition = 8;
			ultraGridColumn72.Width = 65;
			ultraGridBand12.Columns.AddRange(new object[] {
																											ultraGridColumn64,
																											ultraGridColumn65,
																											ultraGridColumn66,
																											ultraGridColumn67,
																											ultraGridColumn68,
																											ultraGridColumn69,
																											ultraGridColumn70,
																											ultraGridColumn71,
																											ultraGridColumn72});
			this.uGridUniformes.DisplayLayout.BandsSerializer.Add(ultraGridBand12);
			this.uGridUniformes.DisplayLayout.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.TemplateOnBottom;
			this.uGridUniformes.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.True;
			this.uGridUniformes.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.True;
			appearance33.BackColor = System.Drawing.Color.Transparent;
			this.uGridUniformes.DisplayLayout.Override.CardAreaAppearance = appearance33;
			appearance34.ForeColor = System.Drawing.Color.Black;
			appearance34.ForeColorDisabled = System.Drawing.Color.Black;
			this.uGridUniformes.DisplayLayout.Override.CellAppearance = appearance34;
			appearance35.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance35.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance35.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance35.FontData.BoldAsString = "True";
			appearance35.FontData.Name = "Arial";
			appearance35.FontData.SizeInPoints = 8F;
			appearance35.ForeColor = System.Drawing.Color.White;
			appearance35.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.uGridUniformes.DisplayLayout.Override.HeaderAppearance = appearance35;
			appearance36.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance36.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance36.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridUniformes.DisplayLayout.Override.RowSelectorAppearance = appearance36;
			appearance37.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance37.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance37.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridUniformes.DisplayLayout.Override.SelectedRowAppearance = appearance37;
			this.uGridUniformes.Font = new System.Drawing.Font("Tahoma", 7.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.uGridUniformes.Location = new System.Drawing.Point(8, 8);
			this.uGridUniformes.Name = "uGridUniformes";
			this.uGridUniformes.Size = new System.Drawing.Size(1107, 272);
			this.uGridUniformes.TabIndex = 3;
			this.uGridUniformes.AfterRowInsert += new Infragistics.Win.UltraWinGrid.RowEventHandler(this.uGridUniformes_AfterRowInsert);
			this.uGridUniformes.BeforeRowsDeleted += new Infragistics.Win.UltraWinGrid.BeforeRowsDeletedEventHandler(this.uGridUniformes_BeforeRowsDeleted);
			this.uGridUniformes.KeyDown += new System.Windows.Forms.KeyEventHandler(this.uGridUniformes_KeyDown);
			this.uGridUniformes.InitializeLayout += new Infragistics.Win.UltraWinGrid.InitializeLayoutEventHandler(this.uGridUniformes_InitializeLayout);
			// 
			// ultraDataSource16
			// 
			ultraDataColumn64.DataType = typeof(int);
			ultraDataColumn65.DataType = typeof(int);
			ultraDataColumn66.DataType = typeof(System.DateTime);
			ultraDataColumn68.DataType = typeof(System.DateTime);
			ultraDataColumn70.DataType = typeof(System.Decimal);
			ultraDataColumn71.DataType = typeof(int);
			ultraDataColumn72.DataType = typeof(bool);
			this.ultraDataSource16.Band.Columns.AddRange(new object[] {
																																	ultraDataColumn64,
																																	ultraDataColumn65,
																																	ultraDataColumn66,
																																	ultraDataColumn67,
																																	ultraDataColumn68,
																																	ultraDataColumn69,
																																	ultraDataColumn70,
																																	ultraDataColumn71,
																																	ultraDataColumn72});
			// 
			// ultraTabPageControl4
			// 
			this.ultraTabPageControl4.Controls.Add(this.cmbTiposDeParto);
			this.ultraTabPageControl4.Controls.Add(this.cmbTipoDeSalida);
			this.ultraTabPageControl4.Controls.Add(this.uGridVacaciones);
			this.ultraTabPageControl4.Location = new System.Drawing.Point(-10000, -10000);
			this.ultraTabPageControl4.Name = "ultraTabPageControl4";
			this.ultraTabPageControl4.Size = new System.Drawing.Size(1121, 283);
			// 
			// cmbTiposDeParto
			// 
			this.cmbTiposDeParto.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbTiposDeParto.DataSource = this.ultraDataSource30;
			ultraGridBand13.AddButtonCaption = "PersParentesco";
			ultraGridColumn73.Header.VisiblePosition = 0;
			ultraGridColumn73.Hidden = true;
			ultraGridColumn74.Header.VisiblePosition = 1;
			ultraGridColumn74.Width = 250;
			ultraGridBand13.Columns.AddRange(new object[] {
																											ultraGridColumn73,
																											ultraGridColumn74});
			this.cmbTiposDeParto.DisplayLayout.BandsSerializer.Add(ultraGridBand13);
			this.cmbTiposDeParto.DisplayMember = "Descripcion";
			this.cmbTiposDeParto.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbTiposDeParto.Location = new System.Drawing.Point(640, 96);
			this.cmbTiposDeParto.Name = "cmbTiposDeParto";
			this.cmbTiposDeParto.Size = new System.Drawing.Size(173, 104);
			this.cmbTiposDeParto.TabIndex = 7;
			this.cmbTiposDeParto.ValueMember = "idNominaTiposDeParto";
			this.cmbTiposDeParto.Visible = false;
			// 
			// ultraDataSource30
			// 
			ultraDataColumn73.DataType = typeof(int);
			this.ultraDataSource30.Band.Columns.AddRange(new object[] {
																																	ultraDataColumn73,
																																	ultraDataColumn74});
			// 
			// cmbTipoDeSalida
			// 
			this.cmbTipoDeSalida.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbTipoDeSalida.DataSource = this.ultraDataSource29;
			ultraGridBand14.AddButtonCaption = "PersParentesco";
			ultraGridColumn75.Header.VisiblePosition = 0;
			ultraGridColumn75.Hidden = true;
			ultraGridColumn76.Header.VisiblePosition = 1;
			ultraGridColumn76.Width = 250;
			ultraGridBand14.Columns.AddRange(new object[] {
																											ultraGridColumn75,
																											ultraGridColumn76});
			this.cmbTipoDeSalida.DisplayLayout.BandsSerializer.Add(ultraGridBand14);
			this.cmbTipoDeSalida.DisplayMember = "Descripcion";
			this.cmbTipoDeSalida.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbTipoDeSalida.Location = new System.Drawing.Point(425, 85);
			this.cmbTipoDeSalida.Name = "cmbTipoDeSalida";
			this.cmbTipoDeSalida.Size = new System.Drawing.Size(173, 104);
			this.cmbTipoDeSalida.TabIndex = 6;
			this.cmbTipoDeSalida.ValueMember = "idNominaTiposDeSalidas";
			this.cmbTipoDeSalida.Visible = false;
			// 
			// ultraDataSource29
			// 
			ultraDataColumn75.DataType = typeof(int);
			this.ultraDataSource29.Band.Columns.AddRange(new object[] {
																																	ultraDataColumn75,
																																	ultraDataColumn76});
			// 
			// uGridVacaciones
			// 
			this.uGridVacaciones.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.uGridVacaciones.Cursor = System.Windows.Forms.Cursors.Default;
			this.uGridVacaciones.DataSource = this.ultraDataSource25;
			appearance38.BackColor = System.Drawing.Color.White;
			this.uGridVacaciones.DisplayLayout.Appearance = appearance38;
			ultraGridBand15.AddButtonCaption = "PersFamiliar";
			ultraGridColumn77.Header.VisiblePosition = 0;
			ultraGridColumn77.Hidden = true;
			ultraGridColumn78.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn78.Header.VisiblePosition = 1;
			ultraGridColumn78.Hidden = true;
			ultraGridColumn79.Header.Caption = "Tipo";
			ultraGridColumn79.Header.VisiblePosition = 2;
			ultraGridColumn79.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.DropDown;
			ultraGridColumn79.Width = 207;
			ultraGridColumn80.Header.Caption = "Tipo De Parto";
			ultraGridColumn80.Header.VisiblePosition = 3;
			ultraGridColumn80.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.DropDown;
			ultraGridColumn80.Width = 200;
			ultraGridColumn81.Header.VisiblePosition = 4;
			ultraGridColumn81.Width = 90;
			ultraGridColumn82.Header.VisiblePosition = 5;
			ultraGridColumn82.Width = 90;
			ultraGridColumn83.Case = Infragistics.Win.UltraWinGrid.Case.Upper;
			ultraGridColumn83.Header.VisiblePosition = 6;
			ultraGridColumn83.Width = 400;
			ultraGridBand15.Columns.AddRange(new object[] {
																											ultraGridColumn77,
																											ultraGridColumn78,
																											ultraGridColumn79,
																											ultraGridColumn80,
																											ultraGridColumn81,
																											ultraGridColumn82,
																											ultraGridColumn83});
			this.uGridVacaciones.DisplayLayout.BandsSerializer.Add(ultraGridBand15);
			this.uGridVacaciones.DisplayLayout.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.TemplateOnBottom;
			this.uGridVacaciones.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.True;
			this.uGridVacaciones.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.True;
			appearance39.BackColor = System.Drawing.Color.Transparent;
			this.uGridVacaciones.DisplayLayout.Override.CardAreaAppearance = appearance39;
			appearance40.ForeColor = System.Drawing.Color.Black;
			appearance40.ForeColorDisabled = System.Drawing.Color.Black;
			this.uGridVacaciones.DisplayLayout.Override.CellAppearance = appearance40;
			appearance41.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance41.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance41.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance41.FontData.BoldAsString = "True";
			appearance41.FontData.Name = "Arial";
			appearance41.FontData.SizeInPoints = 8F;
			appearance41.ForeColor = System.Drawing.Color.White;
			appearance41.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.uGridVacaciones.DisplayLayout.Override.HeaderAppearance = appearance41;
			appearance42.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance42.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance42.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridVacaciones.DisplayLayout.Override.RowSelectorAppearance = appearance42;
			appearance43.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance43.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance43.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridVacaciones.DisplayLayout.Override.SelectedRowAppearance = appearance43;
			this.uGridVacaciones.Font = new System.Drawing.Font("Tahoma", 7.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.uGridVacaciones.Location = new System.Drawing.Point(8, 8);
			this.uGridVacaciones.Name = "uGridVacaciones";
			this.uGridVacaciones.Size = new System.Drawing.Size(1107, 272);
			this.uGridVacaciones.TabIndex = 2;
			this.uGridVacaciones.AfterRowInsert += new Infragistics.Win.UltraWinGrid.RowEventHandler(this.uGridVacaciones_AfterRowInsert);
			this.uGridVacaciones.BeforeCellUpdate += new Infragistics.Win.UltraWinGrid.BeforeCellUpdateEventHandler(this.uGridVacaciones_BeforeCellUpdate);
			this.uGridVacaciones.BeforeRowsDeleted += new Infragistics.Win.UltraWinGrid.BeforeRowsDeletedEventHandler(this.uGridVacaciones_BeforeRowsDeleted);
			this.uGridVacaciones.KeyDown += new System.Windows.Forms.KeyEventHandler(this.uGridVacaciones_KeyDown);
			this.uGridVacaciones.AfterCellUpdate += new Infragistics.Win.UltraWinGrid.CellEventHandler(this.uGridVacaciones_AfterCellUpdate);
			this.uGridVacaciones.InitializeLayout += new Infragistics.Win.UltraWinGrid.InitializeLayoutEventHandler(this.uGridVacaciones_InitializeLayout);
			// 
			// ultraDataSource25
			// 
			ultraDataColumn77.DataType = typeof(int);
			ultraDataColumn78.DataType = typeof(int);
			ultraDataColumn79.DataType = typeof(int);
			ultraDataColumn80.DataType = typeof(int);
			ultraDataColumn81.DataType = typeof(System.DateTime);
			ultraDataColumn81.DefaultValue = new System.DateTime(((long)(0)));
			ultraDataColumn82.DataType = typeof(System.DateTime);
			ultraDataColumn82.DefaultValue = new System.DateTime(((long)(0)));
			this.ultraDataSource25.Band.Columns.AddRange(new object[] {
																																	ultraDataColumn77,
																																	ultraDataColumn78,
																																	ultraDataColumn79,
																																	ultraDataColumn80,
																																	ultraDataColumn81,
																																	ultraDataColumn82,
																																	ultraDataColumn83});
			// 
			// utpFaltas
			// 
			this.utpFaltas.Controls.Add(this.cmbTiposDeFaltas);
			this.utpFaltas.Controls.Add(this.uGridFaltas);
			this.utpFaltas.Location = new System.Drawing.Point(-10000, -10000);
			this.utpFaltas.Name = "utpFaltas";
			this.utpFaltas.Size = new System.Drawing.Size(1121, 283);
			// 
			// cmbTiposDeFaltas
			// 
			this.cmbTiposDeFaltas.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbTiposDeFaltas.DataSource = this.ultraDataSource29;
			ultraGridBand16.AddButtonCaption = "PersParentesco";
			ultraGridColumn84.Header.VisiblePosition = 0;
			ultraGridColumn84.Hidden = true;
			ultraGridColumn85.Header.VisiblePosition = 1;
			ultraGridColumn85.Width = 250;
			ultraGridBand16.Columns.AddRange(new object[] {
																											ultraGridColumn84,
																											ultraGridColumn85});
			this.cmbTiposDeFaltas.DisplayLayout.BandsSerializer.Add(ultraGridBand16);
			this.cmbTiposDeFaltas.DisplayMember = "Descripcion";
			this.cmbTiposDeFaltas.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbTiposDeFaltas.Location = new System.Drawing.Point(425, 85);
			this.cmbTiposDeFaltas.Name = "cmbTiposDeFaltas";
			this.cmbTiposDeFaltas.Size = new System.Drawing.Size(173, 104);
			this.cmbTiposDeFaltas.TabIndex = 7;
			this.cmbTiposDeFaltas.ValueMember = "idNominaTiposDeSalidas";
			this.cmbTiposDeFaltas.Visible = false;
			// 
			// uGridFaltas
			// 
			this.uGridFaltas.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.uGridFaltas.Cursor = System.Windows.Forms.Cursors.Default;
			this.uGridFaltas.DataSource = this.ultraDataSource28;
			appearance44.BackColor = System.Drawing.Color.White;
			this.uGridFaltas.DisplayLayout.Appearance = appearance44;
			this.uGridFaltas.DisplayLayout.AutoFitColumns = true;
			ultraGridBand17.AddButtonCaption = "PersFamiliar";
			ultraGridColumn86.Header.VisiblePosition = 0;
			ultraGridColumn86.Hidden = true;
			ultraGridColumn87.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn87.Header.VisiblePosition = 1;
			ultraGridColumn87.Hidden = true;
			ultraGridColumn88.Header.Caption = "Tipo";
			ultraGridColumn88.Header.VisiblePosition = 2;
			ultraGridColumn88.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.DropDown;
			ultraGridColumn88.Width = 249;
			ultraGridColumn89.Header.VisiblePosition = 3;
			ultraGridColumn89.Width = 80;
			ultraGridColumn90.Case = Infragistics.Win.UltraWinGrid.Case.Upper;
			ultraGridColumn90.Header.VisiblePosition = 6;
			ultraGridColumn90.Width = 660;
			ultraGridColumn91.Header.VisiblePosition = 4;
			ultraGridColumn91.Hidden = true;
			ultraGridColumn92.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			ultraGridColumn92.Header.Caption = "Estado";
			ultraGridColumn92.Header.VisiblePosition = 5;
			ultraGridColumn92.Width = 98;
			ultraGridBand17.Columns.AddRange(new object[] {
																											ultraGridColumn86,
																											ultraGridColumn87,
																											ultraGridColumn88,
																											ultraGridColumn89,
																											ultraGridColumn90,
																											ultraGridColumn91,
																											ultraGridColumn92});
			this.uGridFaltas.DisplayLayout.BandsSerializer.Add(ultraGridBand17);
			this.uGridFaltas.DisplayLayout.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.TemplateOnBottom;
			this.uGridFaltas.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.True;
			this.uGridFaltas.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.True;
			appearance45.BackColor = System.Drawing.Color.Transparent;
			this.uGridFaltas.DisplayLayout.Override.CardAreaAppearance = appearance45;
			appearance46.ForeColor = System.Drawing.Color.Black;
			appearance46.ForeColorDisabled = System.Drawing.Color.Black;
			this.uGridFaltas.DisplayLayout.Override.CellAppearance = appearance46;
			appearance47.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance47.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance47.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance47.FontData.BoldAsString = "True";
			appearance47.FontData.Name = "Arial";
			appearance47.FontData.SizeInPoints = 8F;
			appearance47.ForeColor = System.Drawing.Color.White;
			appearance47.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.uGridFaltas.DisplayLayout.Override.HeaderAppearance = appearance47;
			appearance48.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance48.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance48.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridFaltas.DisplayLayout.Override.RowSelectorAppearance = appearance48;
			appearance49.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance49.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance49.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridFaltas.DisplayLayout.Override.SelectedRowAppearance = appearance49;
			this.uGridFaltas.Font = new System.Drawing.Font("Tahoma", 7.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.uGridFaltas.Location = new System.Drawing.Point(7, 5);
			this.uGridFaltas.Name = "uGridFaltas";
			this.uGridFaltas.Size = new System.Drawing.Size(1107, 272);
			this.uGridFaltas.TabIndex = 3;
			this.uGridFaltas.AfterRowInsert += new Infragistics.Win.UltraWinGrid.RowEventHandler(this.uGridFaltas_AfterRowInsert);
			this.uGridFaltas.BeforeCellUpdate += new Infragistics.Win.UltraWinGrid.BeforeCellUpdateEventHandler(this.uGridFaltas_BeforeCellUpdate);
			this.uGridFaltas.BeforeRowsDeleted += new Infragistics.Win.UltraWinGrid.BeforeRowsDeletedEventHandler(this.uGridFaltas_BeforeRowsDeleted);
			this.uGridFaltas.KeyDown += new System.Windows.Forms.KeyEventHandler(this.uGridFaltas_KeyDown);
			this.uGridFaltas.InitializeLayout += new Infragistics.Win.UltraWinGrid.InitializeLayoutEventHandler(this.uGridFaltas_InitializeLayout);
			// 
			// ultraDataSource28
			// 
			ultraDataColumn84.DataType = typeof(int);
			ultraDataColumn85.DataType = typeof(int);
			ultraDataColumn86.DataType = typeof(int);
			ultraDataColumn87.DataType = typeof(System.DateTime);
			ultraDataColumn87.DefaultValue = new System.DateTime(((long)(0)));
			ultraDataColumn89.DataType = typeof(int);
			this.ultraDataSource28.Band.Columns.AddRange(new object[] {
																																	ultraDataColumn84,
																																	ultraDataColumn85,
																																	ultraDataColumn86,
																																	ultraDataColumn87,
																																	ultraDataColumn88,
																																	ultraDataColumn89,
																																	ultraDataColumn90});
			// 
			// utpOtros
			// 
			this.utpOtros.Controls.Add(this.cmbMotivos);
			this.utpOtros.Controls.Add(this.uGridOtros);
			this.utpOtros.Location = new System.Drawing.Point(-10000, -10000);
			this.utpOtros.Name = "utpOtros";
			this.utpOtros.Size = new System.Drawing.Size(1121, 283);
			// 
			// cmbMotivos
			// 
			this.cmbMotivos.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbMotivos.DataSource = this.ultraDataSource32;
			ultraGridBand18.AddButtonCaption = "PersParentesco";
			ultraGridColumn93.Header.VisiblePosition = 0;
			ultraGridColumn93.Hidden = true;
			ultraGridColumn94.Header.VisiblePosition = 1;
			ultraGridColumn94.Width = 300;
			appearance50.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn95.CellAppearance = appearance50;
			ultraGridColumn95.Format = "#,##0.00";
			ultraGridColumn95.Header.VisiblePosition = 2;
			ultraGridColumn95.Width = 70;
			ultraGridColumn96.Header.VisiblePosition = 3;
			ultraGridColumn96.Width = 400;
			ultraGridBand18.Columns.AddRange(new object[] {
																											ultraGridColumn93,
																											ultraGridColumn94,
																											ultraGridColumn95,
																											ultraGridColumn96});
			this.cmbMotivos.DisplayLayout.BandsSerializer.Add(ultraGridBand18);
			this.cmbMotivos.DisplayMember = "Motivo";
			this.cmbMotivos.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbMotivos.Location = new System.Drawing.Point(425, 89);
			this.cmbMotivos.Name = "cmbMotivos";
			this.cmbMotivos.Size = new System.Drawing.Size(173, 104);
			this.cmbMotivos.TabIndex = 8;
			this.cmbMotivos.ValueMember = "idNominaMotivosOtros";
			this.cmbMotivos.Visible = false;
			this.cmbMotivos.InitializeLayout += new Infragistics.Win.UltraWinGrid.InitializeLayoutEventHandler(this.cmbMotivos_InitializeLayout);
			// 
			// ultraDataSource32
			// 
			ultraDataColumn91.DataType = typeof(int);
			ultraDataColumn93.DataType = typeof(System.Decimal);
			this.ultraDataSource32.Band.Columns.AddRange(new object[] {
																																	ultraDataColumn91,
																																	ultraDataColumn92,
																																	ultraDataColumn93,
																																	ultraDataColumn94});
			// 
			// uGridOtros
			// 
			this.uGridOtros.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.uGridOtros.Cursor = System.Windows.Forms.Cursors.Default;
			this.uGridOtros.DataSource = this.ultraDataSource31;
			appearance51.BackColor = System.Drawing.Color.White;
			this.uGridOtros.DisplayLayout.Appearance = appearance51;
			ultraGridBand19.AddButtonCaption = "PersFamiliar";
			ultraGridColumn97.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			ultraGridColumn97.Header.VisiblePosition = 0;
			ultraGridColumn97.Hidden = true;
			ultraGridColumn97.Width = 136;
			ultraGridColumn98.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			ultraGridColumn98.Header.VisiblePosition = 1;
			ultraGridColumn98.Hidden = true;
			ultraGridColumn99.Header.Caption = "Motivo";
			ultraGridColumn99.Header.VisiblePosition = 2;
			ultraGridColumn99.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.DropDown;
			ultraGridColumn99.Width = 250;
			ultraGridColumn100.Header.VisiblePosition = 3;
			ultraGridColumn100.Width = 80;
			appearance52.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn101.CellAppearance = appearance52;
			ultraGridColumn101.Format = "#,##0.00";
			ultraGridColumn101.Header.VisiblePosition = 4;
			ultraGridColumn101.Width = 70;
			ultraGridColumn102.Case = Infragistics.Win.UltraWinGrid.Case.Upper;
			ultraGridColumn102.Header.Caption = "Observaciones";
			ultraGridColumn102.Header.VisiblePosition = 5;
			ultraGridColumn102.Width = 487;
			ultraGridColumn103.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			ultraGridColumn103.Header.VisiblePosition = 6;
			ultraGridColumn103.Hidden = true;
			ultraGridColumn103.Width = 135;
			ultraGridColumn104.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			ultraGridColumn104.Header.Caption = "Estado";
			ultraGridColumn104.Header.VisiblePosition = 7;
			ultraGridColumn104.Width = 100;
			ultraGridBand19.Columns.AddRange(new object[] {
																											ultraGridColumn97,
																											ultraGridColumn98,
																											ultraGridColumn99,
																											ultraGridColumn100,
																											ultraGridColumn101,
																											ultraGridColumn102,
																											ultraGridColumn103,
																											ultraGridColumn104});
			this.uGridOtros.DisplayLayout.BandsSerializer.Add(ultraGridBand19);
			this.uGridOtros.DisplayLayout.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.TemplateOnBottom;
			this.uGridOtros.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.True;
			this.uGridOtros.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.True;
			appearance53.BackColor = System.Drawing.Color.Transparent;
			this.uGridOtros.DisplayLayout.Override.CardAreaAppearance = appearance53;
			appearance54.ForeColor = System.Drawing.Color.Black;
			appearance54.ForeColorDisabled = System.Drawing.Color.Black;
			this.uGridOtros.DisplayLayout.Override.CellAppearance = appearance54;
			appearance55.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance55.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance55.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance55.FontData.BoldAsString = "True";
			appearance55.FontData.Name = "Arial";
			appearance55.FontData.SizeInPoints = 8F;
			appearance55.ForeColor = System.Drawing.Color.White;
			appearance55.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.uGridOtros.DisplayLayout.Override.HeaderAppearance = appearance55;
			appearance56.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance56.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance56.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridOtros.DisplayLayout.Override.RowSelectorAppearance = appearance56;
			appearance57.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance57.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance57.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridOtros.DisplayLayout.Override.SelectedRowAppearance = appearance57;
			this.uGridOtros.Font = new System.Drawing.Font("Tahoma", 7.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.uGridOtros.Location = new System.Drawing.Point(7, 5);
			this.uGridOtros.Name = "uGridOtros";
			this.uGridOtros.Size = new System.Drawing.Size(1107, 272);
			this.uGridOtros.TabIndex = 4;
			this.uGridOtros.AfterRowInsert += new Infragistics.Win.UltraWinGrid.RowEventHandler(this.uGridOtros_AfterRowInsert);
			this.uGridOtros.BeforeCellUpdate += new Infragistics.Win.UltraWinGrid.BeforeCellUpdateEventHandler(this.uGridOtros_BeforeCellUpdate);
			this.uGridOtros.BeforeRowsDeleted += new Infragistics.Win.UltraWinGrid.BeforeRowsDeletedEventHandler(this.uGridOtros_BeforeRowsDeleted);
			this.uGridOtros.KeyDown += new System.Windows.Forms.KeyEventHandler(this.uGridOtros_KeyDown);
			this.uGridOtros.AfterCellUpdate += new Infragistics.Win.UltraWinGrid.CellEventHandler(this.uGridOtros_AfterCellUpdate);
			this.uGridOtros.InitializeLayout += new Infragistics.Win.UltraWinGrid.InitializeLayoutEventHandler(this.uGridOtros_InitializeLayout);
			// 
			// ultraDataSource31
			// 
			ultraDataColumn95.DataType = typeof(int);
			ultraDataColumn96.DataType = typeof(int);
			ultraDataColumn97.DataType = typeof(int);
			ultraDataColumn98.DataType = typeof(System.DateTime);
			ultraDataColumn99.DataType = typeof(System.Decimal);
			ultraDataColumn101.DataType = typeof(int);
			this.ultraDataSource31.Band.Columns.AddRange(new object[] {
																																	ultraDataColumn95,
																																	ultraDataColumn96,
																																	ultraDataColumn97,
																																	ultraDataColumn98,
																																	ultraDataColumn99,
																																	ultraDataColumn100,
																																	ultraDataColumn101,
																																	ultraDataColumn102});
			// 
			// utpCargas
			// 
			this.utpCargas.Controls.Add(this.uGridCargas);
			this.utpCargas.Location = new System.Drawing.Point(-10000, -10000);
			this.utpCargas.Name = "utpCargas";
			this.utpCargas.Size = new System.Drawing.Size(1121, 283);
			this.utpCargas.Paint += new System.Windows.Forms.PaintEventHandler(this.ultraTabPageControl9_Paint);
			// 
			// uGridCargas
			// 
			this.uGridCargas.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.uGridCargas.Cursor = System.Windows.Forms.Cursors.Default;
			this.uGridCargas.DataSource = this.ultraDataSource35;
			appearance58.BackColor = System.Drawing.Color.White;
			this.uGridCargas.DisplayLayout.Appearance = appearance58;
			this.uGridCargas.DisplayLayout.AutoFitColumns = true;
			ultraGridBand20.AddButtonCaption = "SeteoF";
			ultraGridColumn105.Header.VisiblePosition = 0;
			ultraGridColumn105.Hidden = true;
			ultraGridColumn105.Width = 273;
			ultraGridColumn106.Header.VisiblePosition = 1;
			ultraGridColumn106.Hidden = true;
			ultraGridColumn106.Width = 171;
			ultraGridColumn107.Header.Caption = "Apellidos y Nombres";
			ultraGridColumn107.Header.VisiblePosition = 2;
			ultraGridColumn107.Width = 922;
			ultraGridColumn108.Format = "dd/MM/yyyy";
			ultraGridColumn108.Header.Caption = "Fecha De Nacimiento";
			ultraGridColumn108.Header.VisiblePosition = 3;
			ultraGridColumn108.Width = 165;
			ultraGridBand20.Columns.AddRange(new object[] {
																											ultraGridColumn105,
																											ultraGridColumn106,
																											ultraGridColumn107,
																											ultraGridColumn108});
			ultraGridBand20.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.False;
			ultraGridBand20.Override.CellClickAction = Infragistics.Win.UltraWinGrid.CellClickAction.Edit;
			this.uGridCargas.DisplayLayout.BandsSerializer.Add(ultraGridBand20);
			appearance59.ForeColor = System.Drawing.Color.Black;
			appearance59.ForeColorDisabled = System.Drawing.Color.Black;
			this.uGridCargas.DisplayLayout.Override.ActiveRowAppearance = appearance59;
			this.uGridCargas.DisplayLayout.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.TemplateOnBottom;
			appearance60.BackColor = System.Drawing.Color.Transparent;
			this.uGridCargas.DisplayLayout.Override.CardAreaAppearance = appearance60;
			appearance61.ForeColor = System.Drawing.Color.Black;
			appearance61.ForeColorDisabled = System.Drawing.Color.Black;
			this.uGridCargas.DisplayLayout.Override.CellAppearance = appearance61;
			this.uGridCargas.DisplayLayout.Override.CellClickAction = Infragistics.Win.UltraWinGrid.CellClickAction.Edit;
			appearance62.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance62.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance62.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance62.FontData.BoldAsString = "True";
			appearance62.FontData.Name = "Arial";
			appearance62.FontData.SizeInPoints = 8F;
			appearance62.ForeColor = System.Drawing.Color.White;
			appearance62.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.uGridCargas.DisplayLayout.Override.HeaderAppearance = appearance62;
			appearance63.BackColor = System.Drawing.SystemColors.ControlLight;
			appearance63.BackColor2 = System.Drawing.Color.Gainsboro;
			appearance63.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridCargas.DisplayLayout.Override.RowAlternateAppearance = appearance63;
			appearance64.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance64.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance64.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridCargas.DisplayLayout.Override.RowSelectorAppearance = appearance64;
			appearance65.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance65.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance65.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridCargas.DisplayLayout.Override.SelectedRowAppearance = appearance65;
			this.uGridCargas.Font = new System.Drawing.Font("Microsoft Sans Serif", 7.5F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.uGridCargas.Location = new System.Drawing.Point(7, 6);
			this.uGridCargas.Name = "uGridCargas";
			this.uGridCargas.Size = new System.Drawing.Size(1107, 271);
			this.uGridCargas.TabIndex = 415;
			this.uGridCargas.AfterRowInsert += new Infragistics.Win.UltraWinGrid.RowEventHandler(this.uGridCargas_AfterRowInsert);
			this.uGridCargas.BeforeRowsDeleted += new Infragistics.Win.UltraWinGrid.BeforeRowsDeletedEventHandler(this.uGridCargas_BeforeRowsDeleted);
			this.uGridCargas.KeyDown += new System.Windows.Forms.KeyEventHandler(this.uGridCargas_KeyDown);
			// 
			// ultraDataSource35
			// 
			ultraDataColumn103.DataType = typeof(int);
			ultraDataColumn104.DataType = typeof(int);
			ultraDataColumn106.DataType = typeof(System.DateTime);
			this.ultraDataSource35.Band.Columns.AddRange(new object[] {
																																	ultraDataColumn103,
																																	ultraDataColumn104,
																																	ultraDataColumn105,
																																	ultraDataColumn106});
			// 
			// ultraTabPageControl8
			// 
			this.ultraTabPageControl8.Controls.Add(this.uGridObservaciones);
			this.ultraTabPageControl8.Location = new System.Drawing.Point(-10000, -10000);
			this.ultraTabPageControl8.Name = "ultraTabPageControl8";
			this.ultraTabPageControl8.Size = new System.Drawing.Size(1121, 283);
			// 
			// uGridObservaciones
			// 
			this.uGridObservaciones.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.uGridObservaciones.Cursor = System.Windows.Forms.Cursors.Default;
			this.uGridObservaciones.DataSource = this.ultraDataSource33;
			appearance66.BackColor = System.Drawing.Color.White;
			this.uGridObservaciones.DisplayLayout.Appearance = appearance66;
			this.uGridObservaciones.DisplayLayout.AutoFitColumns = true;
			ultraGridBand21.AddButtonCaption = "SeteoF";
			ultraGridColumn109.Header.VisiblePosition = 0;
			ultraGridColumn109.Hidden = true;
			ultraGridColumn110.Case = Infragistics.Win.UltraWinGrid.Case.Upper;
			ultraGridColumn110.Header.VisiblePosition = 1;
			ultraGridColumn110.Width = 868;
			ultraGridColumn111.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			ultraGridColumn111.Format = "dd/MM/yyyy HH:mm";
			ultraGridColumn111.Header.VisiblePosition = 2;
			ultraGridColumn111.Width = 110;
			ultraGridColumn112.CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled;
			ultraGridColumn112.Header.VisiblePosition = 3;
			ultraGridColumn112.Width = 109;
			ultraGridBand21.Columns.AddRange(new object[] {
																											ultraGridColumn109,
																											ultraGridColumn110,
																											ultraGridColumn111,
																											ultraGridColumn112});
			ultraGridBand21.Override.CellClickAction = Infragistics.Win.UltraWinGrid.CellClickAction.Edit;
			this.uGridObservaciones.DisplayLayout.BandsSerializer.Add(ultraGridBand21);
			appearance67.ForeColor = System.Drawing.Color.Black;
			appearance67.ForeColorDisabled = System.Drawing.Color.Black;
			this.uGridObservaciones.DisplayLayout.Override.ActiveRowAppearance = appearance67;
			this.uGridObservaciones.DisplayLayout.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.No;
			this.uGridObservaciones.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.False;
			this.uGridObservaciones.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.False;
			appearance68.BackColor = System.Drawing.Color.Transparent;
			this.uGridObservaciones.DisplayLayout.Override.CardAreaAppearance = appearance68;
			appearance69.ForeColor = System.Drawing.Color.Black;
			appearance69.ForeColorDisabled = System.Drawing.Color.Black;
			this.uGridObservaciones.DisplayLayout.Override.CellAppearance = appearance69;
			this.uGridObservaciones.DisplayLayout.Override.CellClickAction = Infragistics.Win.UltraWinGrid.CellClickAction.Edit;
			appearance70.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance70.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance70.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance70.FontData.BoldAsString = "True";
			appearance70.FontData.Name = "Arial";
			appearance70.FontData.SizeInPoints = 8F;
			appearance70.ForeColor = System.Drawing.Color.White;
			appearance70.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.uGridObservaciones.DisplayLayout.Override.HeaderAppearance = appearance70;
			appearance71.BackColor = System.Drawing.SystemColors.ControlLight;
			appearance71.BackColor2 = System.Drawing.Color.Gainsboro;
			appearance71.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridObservaciones.DisplayLayout.Override.RowAlternateAppearance = appearance71;
			appearance72.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance72.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance72.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridObservaciones.DisplayLayout.Override.RowSelectorAppearance = appearance72;
			appearance73.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance73.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance73.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridObservaciones.DisplayLayout.Override.SelectedRowAppearance = appearance73;
			this.uGridObservaciones.Font = new System.Drawing.Font("Microsoft Sans Serif", 7.5F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.uGridObservaciones.Location = new System.Drawing.Point(8, 8);
			this.uGridObservaciones.Name = "uGridObservaciones";
			this.uGridObservaciones.Size = new System.Drawing.Size(1107, 271);
			this.uGridObservaciones.TabIndex = 414;
			// 
			// ultraDataSource33
			// 
			ultraDataColumn107.DataType = typeof(int);
			ultraDataColumn107.DefaultValue = 0;
			ultraDataColumn109.DataType = typeof(System.DateTime);
			ultraDataColumn109.DefaultValue = new System.DateTime(((long)(0)));
			this.ultraDataSource33.Band.Columns.AddRange(new object[] {
																																	ultraDataColumn107,
																																	ultraDataColumn108,
																																	ultraDataColumn109,
																																	ultraDataColumn110});
			// 
			// utcDatosGenerales
			// 
			this.utcDatosGenerales.Controls.Add(this.optNEMesesAños);
			this.utcDatosGenerales.Controls.Add(this.label45);
			this.utcDatosGenerales.Controls.Add(this.txtNEVive_Con);
			this.utcDatosGenerales.Controls.Add(this.label40);
			this.utcDatosGenerales.Controls.Add(this.txtPredial);
			this.utcDatosGenerales.Controls.Add(this.label39);
			this.utcDatosGenerales.Controls.Add(this.txtNETiempoDeVivienda);
			this.utcDatosGenerales.Controls.Add(this.txtNEEMailEmpresarial);
			this.utcDatosGenerales.Controls.Add(this.label38);
			this.utcDatosGenerales.Controls.Add(this.cmbNEProvincia);
			this.utcDatosGenerales.Controls.Add(this.cmbProvinciaNacimiento);
			this.utcDatosGenerales.Controls.Add(this.label11);
			this.utcDatosGenerales.Controls.Add(this.label70);
			this.utcDatosGenerales.Controls.Add(this.optNEVivienda);
			this.utcDatosGenerales.Controls.Add(this.txtNEEMail);
			this.utcDatosGenerales.Controls.Add(this.label32);
			this.utcDatosGenerales.Controls.Add(this.txtNECelular);
			this.utcDatosGenerales.Controls.Add(this.label20);
			this.utcDatosGenerales.Controls.Add(this.txtNETelefono);
			this.utcDatosGenerales.Controls.Add(this.label19);
			this.utcDatosGenerales.Controls.Add(this.txtNEDireccion);
			this.utcDatosGenerales.Controls.Add(this.label18);
			this.utcDatosGenerales.Controls.Add(this.cmbNEBarrio);
			this.utcDatosGenerales.Controls.Add(this.cmbNEParroquia);
			this.utcDatosGenerales.Controls.Add(this.label6);
			this.utcDatosGenerales.Controls.Add(this.label9);
			this.utcDatosGenerales.Controls.Add(this.cmbNECanton);
			this.utcDatosGenerales.Controls.Add(this.label3);
			this.utcDatosGenerales.Controls.Add(this.label4);
			this.utcDatosGenerales.Controls.Add(this.txtNECargasFamiliares);
			this.utcDatosGenerales.Controls.Add(this.label5);
			this.utcDatosGenerales.Controls.Add(this.cmbNEEstadoCivil);
			this.utcDatosGenerales.Controls.Add(this.label2);
			this.utcDatosGenerales.Controls.Add(this.txtConadis);
			this.utcDatosGenerales.Controls.Add(this.label69);
			this.utcDatosGenerales.Controls.Add(this.txtPorcentajeDiscapacidad);
			this.utcDatosGenerales.Controls.Add(this.chkNEDiscapacidad);
			this.utcDatosGenerales.Controls.Add(this.label67);
			this.utcDatosGenerales.Controls.Add(this.label68);
			this.utcDatosGenerales.Controls.Add(this.label10);
			this.utcDatosGenerales.Controls.Add(this.dtpNEFechaNacimiento);
			this.utcDatosGenerales.Controls.Add(this.txtNETipoDeSangre);
			this.utcDatosGenerales.Controls.Add(this.label64);
			this.utcDatosGenerales.Controls.Add(this.optNESexo);
			this.utcDatosGenerales.Controls.Add(this.cmbCantonNacimiento);
			this.utcDatosGenerales.Controls.Add(this.label22);
			this.utcDatosGenerales.Controls.Add(this.label21);
			this.utcDatosGenerales.Controls.Add(this.txtTipoDiscapacidad);
			this.utcDatosGenerales.Controls.Add(this.groupBox1);
			this.utcDatosGenerales.Controls.Add(this.chkNELicencia);
			this.utcDatosGenerales.Controls.Add(this.label44);
			this.utcDatosGenerales.Controls.Add(this.dtpFechaCaducidad);
			this.utcDatosGenerales.Controls.Add(this.txtTipoDeLicencia);
			this.utcDatosGenerales.Controls.Add(this.lblEstado);
			this.utcDatosGenerales.Location = new System.Drawing.Point(-10000, -10000);
			this.utcDatosGenerales.Name = "utcDatosGenerales";
			this.utcDatosGenerales.Size = new System.Drawing.Size(1136, 315);
			this.utcDatosGenerales.Paint += new System.Windows.Forms.PaintEventHandler(this.utcDatosGenerales_Paint);
			// 
			// optNEMesesAños
			// 
			this.optNEMesesAños.CheckedIndex = 0;
			this.optNEMesesAños.Enabled = false;
			this.optNEMesesAños.ItemAppearance = appearance74;
			valueListItem1.DataValue = 1;
			valueListItem1.DisplayText = "Meses";
			valueListItem2.DataValue = 2;
			valueListItem2.DisplayText = "Años";
			this.optNEMesesAños.Items.Add(valueListItem1);
			this.optNEMesesAños.Items.Add(valueListItem2);
			this.optNEMesesAños.Location = new System.Drawing.Point(376, 278);
			this.optNEMesesAños.Name = "optNEMesesAños";
			this.optNEMesesAños.Size = new System.Drawing.Size(104, 24);
			this.optNEMesesAños.TabIndex = 351;
			this.optNEMesesAños.Text = "Meses";
			this.optNEMesesAños.KeyDown += new System.Windows.Forms.KeyEventHandler(this.optNEMesesAños_KeyDown);
			this.optNEMesesAños.ValueChanged += new System.EventHandler(this.optNEMesesAños_ValueChanged);
			// 
			// label45
			// 
			this.label45.AutoSize = true;
			this.label45.BackColor = System.Drawing.Color.Transparent;
			this.label45.Location = new System.Drawing.Point(584, 98);
			this.label45.Name = "label45";
			this.label45.Size = new System.Drawing.Size(26, 17);
			this.label45.TabIndex = 355;
			this.label45.Text = "Tipo";
			this.label45.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// txtNEVive_Con
			// 
			this.txtNEVive_Con.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtNEVive_Con.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtNEVive_Con.Enabled = false;
			this.txtNEVive_Con.Location = new System.Drawing.Point(488, 280);
			this.txtNEVive_Con.MaxLength = 60;
			this.txtNEVive_Con.Name = "txtNEVive_Con";
			this.txtNEVive_Con.Size = new System.Drawing.Size(304, 22);
			this.txtNEVive_Con.TabIndex = 194;
			this.txtNEVive_Con.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtNEVive_Con_KeyDown);
			this.txtNEVive_Con.ValueChanged += new System.EventHandler(this.txtNEVive_Con_ValueChanged);
			// 
			// label40
			// 
			this.label40.AutoSize = true;
			this.label40.BackColor = System.Drawing.SystemColors.Control;
			this.label40.Location = new System.Drawing.Point(488, 256);
			this.label40.Name = "label40";
			this.label40.Size = new System.Drawing.Size(48, 17);
			this.label40.TabIndex = 193;
			this.label40.Text = "Vive Con";
			this.label40.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// txtPredial
			// 
			this.txtPredial.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtPredial.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtPredial.Enabled = false;
			this.txtPredial.Location = new System.Drawing.Point(800, 280);
			this.txtPredial.MaxLength = 60;
			this.txtPredial.Name = "txtPredial";
			this.txtPredial.Size = new System.Drawing.Size(144, 22);
			this.txtPredial.TabIndex = 192;
			this.txtPredial.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtPredial_KeyDown);
			this.txtPredial.ValueChanged += new System.EventHandler(this.txtPredial_ValueChanged);
			// 
			// label39
			// 
			this.label39.AutoSize = true;
			this.label39.BackColor = System.Drawing.SystemColors.Control;
			this.label39.Location = new System.Drawing.Point(800, 256);
			this.label39.Name = "label39";
			this.label39.Size = new System.Drawing.Size(38, 17);
			this.label39.TabIndex = 191;
			this.label39.Text = "Predial";
			this.label39.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// txtNETiempoDeVivienda
			// 
			this.txtNETiempoDeVivienda.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtNETiempoDeVivienda.Enabled = false;
			this.txtNETiempoDeVivienda.Location = new System.Drawing.Point(296, 280);
			this.txtNETiempoDeVivienda.MinValue = 0;
			this.txtNETiempoDeVivienda.Name = "txtNETiempoDeVivienda";
			this.txtNETiempoDeVivienda.PromptChar = ' ';
			this.txtNETiempoDeVivienda.Size = new System.Drawing.Size(69, 22);
			this.txtNETiempoDeVivienda.SpinButtonDisplayStyle = Infragistics.Win.ButtonDisplayStyle.Always;
			this.txtNETiempoDeVivienda.TabIndex = 190;
			this.txtNETiempoDeVivienda.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtNETiempoDeVivienda_KeyDown);
			this.txtNETiempoDeVivienda.Click += new System.EventHandler(this.txtNETiempoDeVivienda_Click);
			this.txtNETiempoDeVivienda.Validated += new System.EventHandler(this.txtNETiempoDeVivienda_Validated);
			this.txtNETiempoDeVivienda.ValueChanged += new System.EventHandler(this.txtNETiempoDeVivienda_ValueChanged);
			this.txtNETiempoDeVivienda.Enter += new System.EventHandler(this.txtNETiempoDeVivienda_Enter);
			// 
			// txtNEEMailEmpresarial
			// 
			this.txtNEEMailEmpresarial.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtNEEMailEmpresarial.Enabled = false;
			this.txtNEEMailEmpresarial.Location = new System.Drawing.Point(800, 224);
			this.txtNEEMailEmpresarial.MaxLength = 60;
			this.txtNEEMailEmpresarial.Name = "txtNEEMailEmpresarial";
			this.txtNEEMailEmpresarial.Size = new System.Drawing.Size(144, 22);
			this.txtNEEMailEmpresarial.TabIndex = 189;
			this.txtNEEMailEmpresarial.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtNEEMailEmpresarial_KeyDown);
			// 
			// label38
			// 
			this.label38.AutoSize = true;
			this.label38.BackColor = System.Drawing.SystemColors.Control;
			this.label38.Location = new System.Drawing.Point(800, 200);
			this.label38.Name = "label38";
			this.label38.Size = new System.Drawing.Size(96, 17);
			this.label38.TabIndex = 188;
			this.label38.Text = "E-Mail Empresarial";
			this.label38.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// cmbNEProvincia
			// 
			this.cmbNEProvincia.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbNEProvincia.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbNEProvincia.DataSource = this.ultraDataSource2;
			ultraGridColumn113.Header.VisiblePosition = 0;
			ultraGridColumn113.Hidden = true;
			ultraGridColumn114.Header.VisiblePosition = 1;
			ultraGridColumn114.Width = 208;
			ultraGridBand22.Columns.AddRange(new object[] {
																											ultraGridColumn113,
																											ultraGridColumn114});
			this.cmbNEProvincia.DisplayLayout.BandsSerializer.Add(ultraGridBand22);
			this.cmbNEProvincia.DisplayMember = "Nombre";
			this.cmbNEProvincia.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbNEProvincia.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbNEProvincia.Enabled = false;
			this.cmbNEProvincia.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbNEProvincia.Location = new System.Drawing.Point(91, 168);
			this.cmbNEProvincia.Name = "cmbNEProvincia";
			this.cmbNEProvincia.Size = new System.Drawing.Size(208, 22);
			this.cmbNEProvincia.TabIndex = 186;
			this.cmbNEProvincia.ValueMember = "idProvincia";
			this.cmbNEProvincia.KeyDown += new System.Windows.Forms.KeyEventHandler(this.cmbNEProvincia_KeyDown);
			this.cmbNEProvincia.ValueChanged += new System.EventHandler(this.cmbNEProvincia_ValueChanged);
			this.cmbNEProvincia.InitializeLayout += new Infragistics.Win.UltraWinGrid.InitializeLayoutEventHandler(this.cmbNEProvincia_InitializeLayout);
			// 
			// ultraDataSource2
			// 
			ultraDataColumn111.DataType = typeof(int);
			this.ultraDataSource2.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn111,
																																 ultraDataColumn112});
			// 
			// cmbProvinciaNacimiento
			// 
			this.cmbProvinciaNacimiento.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbProvinciaNacimiento.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbProvinciaNacimiento.DataSource = this.ultraDataSource2;
			ultraGridBand23.AddButtonCaption = "RolSucursal";
			ultraGridColumn115.Header.VisiblePosition = 0;
			ultraGridColumn115.Hidden = true;
			ultraGridColumn116.Header.VisiblePosition = 1;
			ultraGridColumn116.Width = 208;
			ultraGridBand23.Columns.AddRange(new object[] {
																											ultraGridColumn115,
																											ultraGridColumn116});
			this.cmbProvinciaNacimiento.DisplayLayout.BandsSerializer.Add(ultraGridBand23);
			this.cmbProvinciaNacimiento.DisplayMember = "Nombre";
			this.cmbProvinciaNacimiento.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbProvinciaNacimiento.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbProvinciaNacimiento.Enabled = false;
			this.cmbProvinciaNacimiento.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbProvinciaNacimiento.Location = new System.Drawing.Point(16, 32);
			this.cmbProvinciaNacimiento.Name = "cmbProvinciaNacimiento";
			this.cmbProvinciaNacimiento.Size = new System.Drawing.Size(208, 22);
			this.cmbProvinciaNacimiento.TabIndex = 185;
			this.cmbProvinciaNacimiento.ValueMember = "idProvincia";
			this.cmbProvinciaNacimiento.KeyDown += new System.Windows.Forms.KeyEventHandler(this.cmbNEProvinciaNacimiento_KeyDown);
			this.cmbProvinciaNacimiento.ValueChanged += new System.EventHandler(this.cmbNEProvinciaNacimiento_ValueChanged);
			this.cmbProvinciaNacimiento.InitializeLayout += new Infragistics.Win.UltraWinGrid.InitializeLayoutEventHandler(this.cmbNEProvinciaNacimiento_InitializeLayout);
			// 
			// label11
			// 
			this.label11.AutoSize = true;
			this.label11.BackColor = System.Drawing.SystemColors.Control;
			this.label11.Location = new System.Drawing.Point(296, 256);
			this.label11.Name = "label11";
			this.label11.Size = new System.Drawing.Size(103, 17);
			this.label11.TabIndex = 182;
			this.label11.Text = "Tiempo de Vivienda";
			this.label11.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// label70
			// 
			this.label70.AutoSize = true;
			this.label70.BackColor = System.Drawing.SystemColors.Control;
			this.label70.Location = new System.Drawing.Point(91, 256);
			this.label70.Name = "label70";
			this.label70.Size = new System.Drawing.Size(87, 17);
			this.label70.TabIndex = 181;
			this.label70.Text = "Tipo de Vivienda";
			this.label70.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// optNEVivienda
			// 
			this.optNEVivienda.CheckedIndex = 0;
			this.optNEVivienda.Enabled = false;
			this.optNEVivienda.ItemAppearance = appearance75;
			valueListItem3.DataValue = 1;
			valueListItem3.DisplayText = "Propia";
			valueListItem4.DataValue = 2;
			valueListItem4.DisplayText = "Arrendada";
			valueListItem5.DataValue = 3;
			valueListItem5.DisplayText = "Familiar";
			this.optNEVivienda.Items.Add(valueListItem3);
			this.optNEVivienda.Items.Add(valueListItem4);
			this.optNEVivienda.Items.Add(valueListItem5);
			this.optNEVivienda.Location = new System.Drawing.Point(91, 278);
			this.optNEVivienda.Name = "optNEVivienda";
			this.optNEVivienda.Size = new System.Drawing.Size(197, 24);
			this.optNEVivienda.TabIndex = 180;
			this.optNEVivienda.Text = "Propia";
			this.optNEVivienda.KeyDown += new System.Windows.Forms.KeyEventHandler(this.optNEVivienda_KeyDown);
			this.optNEVivienda.ValueChanged += new System.EventHandler(this.optNEVivienda_ValueChanged);
			// 
			// txtNEEMail
			// 
			this.txtNEEMail.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtNEEMail.Enabled = false;
			this.txtNEEMail.Location = new System.Drawing.Point(648, 224);
			this.txtNEEMail.MaxLength = 60;
			this.txtNEEMail.Name = "txtNEEMail";
			this.txtNEEMail.Size = new System.Drawing.Size(144, 22);
			this.txtNEEMail.TabIndex = 179;
			this.txtNEEMail.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtNEEMail_KeyDown);
			// 
			// label32
			// 
			this.label32.AutoSize = true;
			this.label32.BackColor = System.Drawing.SystemColors.Control;
			this.label32.Location = new System.Drawing.Point(648, 200);
			this.label32.Name = "label32";
			this.label32.Size = new System.Drawing.Size(80, 17);
			this.label32.TabIndex = 178;
			this.label32.Text = "E-Mail Personal";
			this.label32.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// txtNECelular
			// 
			this.txtNECelular.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtNECelular.Enabled = false;
			this.txtNECelular.Location = new System.Drawing.Point(536, 224);
			this.txtNECelular.MaxLength = 10;
			this.txtNECelular.Name = "txtNECelular";
			this.txtNECelular.Size = new System.Drawing.Size(104, 22);
			this.txtNECelular.TabIndex = 176;
			this.txtNECelular.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtNECelular_KeyDown);
			this.txtNECelular.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.txtNECelular_KeyPress);
			// 
			// label20
			// 
			this.label20.AutoSize = true;
			this.label20.BackColor = System.Drawing.SystemColors.Control;
			this.label20.Location = new System.Drawing.Point(536, 200);
			this.label20.Name = "label20";
			this.label20.Size = new System.Drawing.Size(38, 17);
			this.label20.TabIndex = 175;
			this.label20.Text = "Celular";
			this.label20.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// txtNETelefono
			// 
			this.txtNETelefono.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtNETelefono.Enabled = false;
			this.txtNETelefono.Location = new System.Drawing.Point(432, 224);
			this.txtNETelefono.MaxLength = 9;
			this.txtNETelefono.Name = "txtNETelefono";
			this.txtNETelefono.Size = new System.Drawing.Size(96, 22);
			this.txtNETelefono.TabIndex = 174;
			this.txtNETelefono.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtNETelefono_KeyDown);
			this.txtNETelefono.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.txtNETelefono_KeyPress);
			// 
			// label19
			// 
			this.label19.AutoSize = true;
			this.label19.BackColor = System.Drawing.SystemColors.Control;
			this.label19.Location = new System.Drawing.Point(432, 200);
			this.label19.Name = "label19";
			this.label19.Size = new System.Drawing.Size(48, 17);
			this.label19.TabIndex = 173;
			this.label19.Text = "Teléfono";
			this.label19.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// txtNEDireccion
			// 
			this.txtNEDireccion.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtNEDireccion.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtNEDireccion.Enabled = false;
			this.txtNEDireccion.Location = new System.Drawing.Point(91, 224);
			this.txtNEDireccion.Name = "txtNEDireccion";
			this.txtNEDireccion.Size = new System.Drawing.Size(333, 22);
			this.txtNEDireccion.TabIndex = 172;
			this.txtNEDireccion.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtNEDireccion_KeyDown);
			// 
			// label18
			// 
			this.label18.AutoSize = true;
			this.label18.BackColor = System.Drawing.SystemColors.Control;
			this.label18.Location = new System.Drawing.Point(91, 200);
			this.label18.Name = "label18";
			this.label18.Size = new System.Drawing.Size(50, 17);
			this.label18.TabIndex = 171;
			this.label18.Text = "Dirección";
			this.label18.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// cmbNEBarrio
			// 
			this.cmbNEBarrio.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbNEBarrio.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbNEBarrio.DataSource = this.ultraDataSource6;
			ultraGridBand24.AddButtonCaption = "RolSucursal";
			ultraGridColumn117.Header.VisiblePosition = 0;
			ultraGridColumn117.Hidden = true;
			ultraGridColumn118.Header.VisiblePosition = 1;
			ultraGridColumn118.Width = 208;
			ultraGridBand24.Columns.AddRange(new object[] {
																											ultraGridColumn117,
																											ultraGridColumn118});
			this.cmbNEBarrio.DisplayLayout.BandsSerializer.Add(ultraGridBand24);
			this.cmbNEBarrio.DisplayLayout.Override.RowSelectors = Infragistics.Win.DefaultableBoolean.True;
			this.cmbNEBarrio.DisplayMember = "Nombre";
			this.cmbNEBarrio.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbNEBarrio.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbNEBarrio.Enabled = false;
			this.cmbNEBarrio.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbNEBarrio.Location = new System.Drawing.Point(739, 168);
			this.cmbNEBarrio.Name = "cmbNEBarrio";
			this.cmbNEBarrio.Size = new System.Drawing.Size(208, 22);
			this.cmbNEBarrio.TabIndex = 170;
			this.cmbNEBarrio.ValueMember = "idBarrio";
			this.cmbNEBarrio.KeyDown += new System.Windows.Forms.KeyEventHandler(this.cmbNEBarrio_KeyDown);
			// 
			// ultraDataSource6
			// 
			ultraDataColumn113.DataType = typeof(int);
			this.ultraDataSource6.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn113,
																																 ultraDataColumn114});
			// 
			// cmbNEParroquia
			// 
			this.cmbNEParroquia.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbNEParroquia.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbNEParroquia.DataSource = this.ultraDataSource5;
			ultraGridBand25.AddButtonCaption = "RolSucursal";
			ultraGridColumn119.Header.VisiblePosition = 0;
			ultraGridColumn119.Hidden = true;
			ultraGridColumn120.Header.VisiblePosition = 1;
			ultraGridColumn120.Width = 208;
			ultraGridBand25.Columns.AddRange(new object[] {
																											ultraGridColumn119,
																											ultraGridColumn120});
			this.cmbNEParroquia.DisplayLayout.BandsSerializer.Add(ultraGridBand25);
			this.cmbNEParroquia.DisplayMember = "Nombre";
			this.cmbNEParroquia.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbNEParroquia.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbNEParroquia.Enabled = false;
			this.cmbNEParroquia.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbNEParroquia.Location = new System.Drawing.Point(523, 168);
			this.cmbNEParroquia.Name = "cmbNEParroquia";
			this.cmbNEParroquia.Size = new System.Drawing.Size(208, 22);
			this.cmbNEParroquia.TabIndex = 169;
			this.cmbNEParroquia.ValueMember = "idParroquia";
			this.cmbNEParroquia.KeyDown += new System.Windows.Forms.KeyEventHandler(this.cmbNEParroquia_KeyDown);
			this.cmbNEParroquia.ValueChanged += new System.EventHandler(this.cmbNEParroquia_ValueChanged);
			this.cmbNEParroquia.InitializeLayout += new Infragistics.Win.UltraWinGrid.InitializeLayoutEventHandler(this.cmbNEParroquia_InitializeLayout);
			// 
			// ultraDataSource5
			// 
			ultraDataColumn115.DataType = typeof(int);
			this.ultraDataSource5.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn115,
																																 ultraDataColumn116});
			// 
			// label6
			// 
			this.label6.AutoSize = true;
			this.label6.BackColor = System.Drawing.SystemColors.Control;
			this.label6.Location = new System.Drawing.Point(523, 144);
			this.label6.Name = "label6";
			this.label6.Size = new System.Drawing.Size(52, 17);
			this.label6.TabIndex = 168;
			this.label6.Text = "Parroquia";
			this.label6.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// label9
			// 
			this.label9.AutoSize = true;
			this.label9.BackColor = System.Drawing.SystemColors.Control;
			this.label9.Location = new System.Drawing.Point(739, 144);
			this.label9.Name = "label9";
			this.label9.Size = new System.Drawing.Size(34, 17);
			this.label9.TabIndex = 167;
			this.label9.Text = "Barrio";
			this.label9.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// cmbNECanton
			// 
			this.cmbNECanton.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbNECanton.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbNECanton.DataSource = this.ultraDataSource3;
			ultraGridBand26.AddButtonCaption = "RolSucursal";
			ultraGridColumn121.Header.VisiblePosition = 0;
			ultraGridColumn121.Hidden = true;
			ultraGridColumn122.Header.VisiblePosition = 1;
			ultraGridColumn122.Width = 208;
			ultraGridBand26.Columns.AddRange(new object[] {
																											ultraGridColumn121,
																											ultraGridColumn122});
			this.cmbNECanton.DisplayLayout.BandsSerializer.Add(ultraGridBand26);
			this.cmbNECanton.DisplayLayout.Override.RowSelectors = Infragistics.Win.DefaultableBoolean.True;
			this.cmbNECanton.DisplayMember = "Nombre";
			this.cmbNECanton.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbNECanton.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbNECanton.Enabled = false;
			this.cmbNECanton.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbNECanton.Location = new System.Drawing.Point(307, 168);
			this.cmbNECanton.Name = "cmbNECanton";
			this.cmbNECanton.Size = new System.Drawing.Size(208, 22);
			this.cmbNECanton.TabIndex = 166;
			this.cmbNECanton.ValueMember = "idCanton";
			this.cmbNECanton.KeyDown += new System.Windows.Forms.KeyEventHandler(this.cmbNECanton_KeyDown);
			this.cmbNECanton.ValueChanged += new System.EventHandler(this.cmbNECanton_ValueChanged);
			this.cmbNECanton.InitializeLayout += new Infragistics.Win.UltraWinGrid.InitializeLayoutEventHandler(this.cmbNECanton_InitializeLayout);
			// 
			// ultraDataSource3
			// 
			ultraDataColumn117.DataType = typeof(int);
			this.ultraDataSource3.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn117,
																																 ultraDataColumn118});
			// 
			// label3
			// 
			this.label3.AutoSize = true;
			this.label3.BackColor = System.Drawing.SystemColors.Control;
			this.label3.Location = new System.Drawing.Point(91, 144);
			this.label3.Name = "label3";
			this.label3.Size = new System.Drawing.Size(49, 17);
			this.label3.TabIndex = 164;
			this.label3.Text = "Provincia";
			this.label3.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// label4
			// 
			this.label4.AutoSize = true;
			this.label4.BackColor = System.Drawing.SystemColors.Control;
			this.label4.Location = new System.Drawing.Point(307, 144);
			this.label4.Name = "label4";
			this.label4.Size = new System.Drawing.Size(39, 17);
			this.label4.TabIndex = 163;
			this.label4.Text = "Canton";
			this.label4.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// txtNECargasFamiliares
			// 
			this.txtNECargasFamiliares.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtNECargasFamiliares.Enabled = false;
			this.txtNECargasFamiliares.Location = new System.Drawing.Point(936, 32);
			this.txtNECargasFamiliares.MinValue = 0;
			this.txtNECargasFamiliares.Name = "txtNECargasFamiliares";
			this.txtNECargasFamiliares.PromptChar = ' ';
			this.txtNECargasFamiliares.Size = new System.Drawing.Size(80, 22);
			this.txtNECargasFamiliares.SpinButtonDisplayStyle = Infragistics.Win.ButtonDisplayStyle.Always;
			this.txtNECargasFamiliares.TabIndex = 162;
			this.txtNECargasFamiliares.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtNECargasFamiliares_KeyDown);
			this.txtNECargasFamiliares.Click += new System.EventHandler(this.txtNECargasFamiliares_Click);
			this.txtNECargasFamiliares.Validated += new System.EventHandler(this.txtNECargasFamiliares_Validated);
			this.txtNECargasFamiliares.Enter += new System.EventHandler(this.txtNECargasFamiliares_Enter);
			// 
			// label5
			// 
			this.label5.AutoSize = true;
			this.label5.BackColor = System.Drawing.Color.Transparent;
			this.label5.Location = new System.Drawing.Point(936, 8);
			this.label5.Name = "label5";
			this.label5.Size = new System.Drawing.Size(92, 17);
			this.label5.TabIndex = 161;
			this.label5.Text = "Cargas Familiares";
			this.label5.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// cmbNEEstadoCivil
			// 
			this.cmbNEEstadoCivil.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbNEEstadoCivil.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbNEEstadoCivil.DataSource = this.ultraDataSource4;
			ultraGridBand27.AddButtonCaption = "RolSucursal";
			ultraGridColumn123.Header.VisiblePosition = 0;
			ultraGridColumn123.Hidden = true;
			ultraGridColumn124.Header.VisiblePosition = 1;
			ultraGridColumn124.Width = 208;
			ultraGridBand27.Columns.AddRange(new object[] {
																											ultraGridColumn123,
																											ultraGridColumn124});
			this.cmbNEEstadoCivil.DisplayLayout.BandsSerializer.Add(ultraGridBand27);
			this.cmbNEEstadoCivil.DisplayMember = "Nombre";
			this.cmbNEEstadoCivil.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbNEEstadoCivil.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbNEEstadoCivil.Enabled = false;
			this.cmbNEEstadoCivil.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbNEEstadoCivil.Location = new System.Drawing.Point(728, 32);
			this.cmbNEEstadoCivil.Name = "cmbNEEstadoCivil";
			this.cmbNEEstadoCivil.Size = new System.Drawing.Size(200, 22);
			this.cmbNEEstadoCivil.TabIndex = 160;
			this.cmbNEEstadoCivil.ValueMember = "idEdoCivil";
			this.cmbNEEstadoCivil.KeyDown += new System.Windows.Forms.KeyEventHandler(this.cmbNEEstadoCivil_KeyDown);
			// 
			// ultraDataSource4
			// 
			ultraDataColumn119.DataType = typeof(int);
			this.ultraDataSource4.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn119,
																																 ultraDataColumn120});
			// 
			// label2
			// 
			this.label2.AutoSize = true;
			this.label2.BackColor = System.Drawing.Color.Transparent;
			this.label2.Location = new System.Drawing.Point(728, 8);
			this.label2.Name = "label2";
			this.label2.Size = new System.Drawing.Size(61, 17);
			this.label2.TabIndex = 159;
			this.label2.Text = "Estado Civil";
			this.label2.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// txtConadis
			// 
			this.txtConadis.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtConadis.Enabled = false;
			this.txtConadis.Location = new System.Drawing.Point(136, 96);
			this.txtConadis.MaxLength = 15;
			this.txtConadis.Name = "txtConadis";
			this.txtConadis.Size = new System.Drawing.Size(104, 22);
			this.txtConadis.TabIndex = 158;
			this.txtConadis.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtConadis_KeyDown);
			// 
			// label69
			// 
			this.label69.AutoSize = true;
			this.label69.BackColor = System.Drawing.Color.Transparent;
			this.label69.Location = new System.Drawing.Point(232, 66);
			this.label69.Name = "label69";
			this.label69.Size = new System.Drawing.Size(94, 17);
			this.label69.TabIndex = 157;
			this.label69.Text = "Tipo Discapacidad";
			this.label69.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// txtPorcentajeDiscapacidad
			// 
			this.txtPorcentajeDiscapacidad.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtPorcentajeDiscapacidad.Enabled = false;
			this.txtPorcentajeDiscapacidad.Location = new System.Drawing.Point(32, 96);
			this.txtPorcentajeDiscapacidad.MaxValue = 100;
			this.txtPorcentajeDiscapacidad.MinValue = 0;
			this.txtPorcentajeDiscapacidad.Name = "txtPorcentajeDiscapacidad";
			this.txtPorcentajeDiscapacidad.PromptChar = ' ';
			this.txtPorcentajeDiscapacidad.Size = new System.Drawing.Size(72, 22);
			this.txtPorcentajeDiscapacidad.SpinButtonDisplayStyle = Infragistics.Win.ButtonDisplayStyle.Always;
			this.txtPorcentajeDiscapacidad.TabIndex = 156;
			this.txtPorcentajeDiscapacidad.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtPorcentajeDiscapacidad_KeyDown);
			this.txtPorcentajeDiscapacidad.Click += new System.EventHandler(this.txtPorcentajeDiscapacidad_Click);
			this.txtPorcentajeDiscapacidad.Validated += new System.EventHandler(this.txtPorcentajeDiscapacidad_Validated);
			this.txtPorcentajeDiscapacidad.Enter += new System.EventHandler(this.txtPorcentajeDiscapacidad_Enter);
			// 
			// chkNEDiscapacidad
			// 
			this.chkNEDiscapacidad.BackColor = System.Drawing.Color.Transparent;
			this.chkNEDiscapacidad.Enabled = false;
			this.chkNEDiscapacidad.Location = new System.Drawing.Point(16, 64);
			this.chkNEDiscapacidad.Name = "chkNEDiscapacidad";
			this.chkNEDiscapacidad.Size = new System.Drawing.Size(96, 21);
			this.chkNEDiscapacidad.TabIndex = 155;
			this.chkNEDiscapacidad.Text = "Discapacidad";
			this.chkNEDiscapacidad.KeyDown += new System.Windows.Forms.KeyEventHandler(this.chkNEDiscapacidad_KeyDown);
			this.chkNEDiscapacidad.CheckedChanged += new System.EventHandler(this.chkNEDiscapacidad_CheckedChanged);
			// 
			// label67
			// 
			this.label67.AutoSize = true;
			this.label67.BackColor = System.Drawing.Color.Transparent;
			this.label67.Location = new System.Drawing.Point(120, 66);
			this.label67.Name = "label67";
			this.label67.Size = new System.Drawing.Size(105, 17);
			this.label67.TabIndex = 153;
			this.label67.Text = "Numero del Conadis";
			this.label67.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// label68
			// 
			this.label68.AutoSize = true;
			this.label68.BackColor = System.Drawing.Color.Transparent;
			this.label68.Location = new System.Drawing.Point(112, 98);
			this.label68.Name = "label68";
			this.label68.Size = new System.Drawing.Size(15, 17);
			this.label68.TabIndex = 154;
			this.label68.Text = "%";
			this.label68.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// label10
			// 
			this.label10.AutoSize = true;
			this.label10.BackColor = System.Drawing.Color.Transparent;
			this.label10.Location = new System.Drawing.Point(448, 8);
			this.label10.Name = "label10";
			this.label10.Size = new System.Drawing.Size(112, 17);
			this.label10.TabIndex = 152;
			this.label10.Text = "Fecha de Nacimiento:";
			this.label10.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// dtpNEFechaNacimiento
			// 
			this.dtpNEFechaNacimiento.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			dateButton1.Caption = "Today";
			this.dtpNEFechaNacimiento.DateButtons.Add(dateButton1);
			this.dtpNEFechaNacimiento.Enabled = false;
			this.dtpNEFechaNacimiento.Format = "dd/MM/yyyy";
			this.dtpNEFechaNacimiento.Location = new System.Drawing.Point(448, 32);
			this.dtpNEFechaNacimiento.Name = "dtpNEFechaNacimiento";
			this.dtpNEFechaNacimiento.NonAutoSizeHeight = 23;
			this.dtpNEFechaNacimiento.Size = new System.Drawing.Size(104, 21);
			this.dtpNEFechaNacimiento.TabIndex = 151;
			this.dtpNEFechaNacimiento.Value = ((object)(resources.GetObject("dtpNEFechaNacimiento.Value")));
			this.dtpNEFechaNacimiento.KeyDown += new System.Windows.Forms.KeyEventHandler(this.dtpNEFechaNacimiento_KeyDown);
			// 
			// txtNETipoDeSangre
			// 
			this.txtNETipoDeSangre.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtNETipoDeSangre.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtNETipoDeSangre.Enabled = false;
			this.txtNETipoDeSangre.Location = new System.Drawing.Point(648, 32);
			this.txtNETipoDeSangre.Name = "txtNETipoDeSangre";
			this.txtNETipoDeSangre.Size = new System.Drawing.Size(72, 22);
			this.txtNETipoDeSangre.TabIndex = 150;
			this.txtNETipoDeSangre.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtNETipoDeSangre_KeyDown);
			// 
			// label64
			// 
			this.label64.AutoSize = true;
			this.label64.BackColor = System.Drawing.Color.Transparent;
			this.label64.Location = new System.Drawing.Point(648, 8);
			this.label64.Name = "label64";
			this.label64.Size = new System.Drawing.Size(64, 17);
			this.label64.TabIndex = 149;
			this.label64.Text = "Tipo Sangre";
			this.label64.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// optNESexo
			// 
			this.optNESexo.CheckedIndex = 0;
			this.optNESexo.Enabled = false;
			this.optNESexo.ItemAppearance = appearance76;
			valueListItem6.DataValue = 1;
			valueListItem6.DisplayText = "Hombre";
			valueListItem7.DataValue = 2;
			valueListItem7.DisplayText = "Mujer";
			this.optNESexo.Items.Add(valueListItem6);
			this.optNESexo.Items.Add(valueListItem7);
			this.optNESexo.Location = new System.Drawing.Point(568, 8);
			this.optNESexo.Name = "optNESexo";
			this.optNESexo.Size = new System.Drawing.Size(72, 40);
			this.optNESexo.TabIndex = 148;
			this.optNESexo.Text = "Hombre";
			this.optNESexo.KeyDown += new System.Windows.Forms.KeyEventHandler(this.optNESexo_KeyDown);
			this.optNESexo.ValueChanged += new System.EventHandler(this.optNESexo_ValueChanged);
			// 
			// cmbCantonNacimiento
			// 
			this.cmbCantonNacimiento.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbCantonNacimiento.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbCantonNacimiento.DataSource = this.ultraDataSource3;
			ultraGridBand28.AddButtonCaption = "RolSucursal";
			ultraGridColumn125.Header.VisiblePosition = 0;
			ultraGridColumn125.Hidden = true;
			ultraGridColumn126.Header.VisiblePosition = 1;
			ultraGridColumn126.Width = 208;
			ultraGridBand28.Columns.AddRange(new object[] {
																											ultraGridColumn125,
																											ultraGridColumn126});
			this.cmbCantonNacimiento.DisplayLayout.BandsSerializer.Add(ultraGridBand28);
			this.cmbCantonNacimiento.DisplayLayout.Override.RowSelectors = Infragistics.Win.DefaultableBoolean.True;
			this.cmbCantonNacimiento.DisplayMember = "Nombre";
			this.cmbCantonNacimiento.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbCantonNacimiento.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbCantonNacimiento.Enabled = false;
			this.cmbCantonNacimiento.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbCantonNacimiento.Location = new System.Drawing.Point(232, 32);
			this.cmbCantonNacimiento.Name = "cmbCantonNacimiento";
			this.cmbCantonNacimiento.Size = new System.Drawing.Size(208, 22);
			this.cmbCantonNacimiento.TabIndex = 147;
			this.cmbCantonNacimiento.ValueMember = "idCanton";
			this.cmbCantonNacimiento.KeyDown += new System.Windows.Forms.KeyEventHandler(this.cmbNECantonNacimiento_KeyDown);
			// 
			// label22
			// 
			this.label22.AutoSize = true;
			this.label22.BackColor = System.Drawing.Color.Transparent;
			this.label22.Location = new System.Drawing.Point(16, 8);
			this.label22.Name = "label22";
			this.label22.Size = new System.Drawing.Size(124, 17);
			this.label22.TabIndex = 145;
			this.label22.Text = "Provincia de Nacimiento";
			this.label22.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// label21
			// 
			this.label21.AutoSize = true;
			this.label21.BackColor = System.Drawing.Color.Transparent;
			this.label21.Location = new System.Drawing.Point(232, 8);
			this.label21.Name = "label21";
			this.label21.Size = new System.Drawing.Size(114, 17);
			this.label21.TabIndex = 144;
			this.label21.Text = "Canton de Nacimiento";
			this.label21.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// txtTipoDiscapacidad
			// 
			this.txtTipoDiscapacidad.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtTipoDiscapacidad.Enabled = false;
			this.txtTipoDiscapacidad.Location = new System.Drawing.Point(248, 96);
			this.txtTipoDiscapacidad.Name = "txtTipoDiscapacidad";
			this.txtTipoDiscapacidad.Size = new System.Drawing.Size(304, 22);
			this.txtTipoDiscapacidad.TabIndex = 147;
			this.txtTipoDiscapacidad.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtTipoDiscapacidad_KeyDown);
			this.txtTipoDiscapacidad.ValueChanged += new System.EventHandler(this.txtTipoDiscapacidad_ValueChanged);
			// 
			// groupBox1
			// 
			this.groupBox1.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.groupBox1.Location = new System.Drawing.Point(15, 120);
			this.groupBox1.Name = "groupBox1";
			this.groupBox1.Size = new System.Drawing.Size(1008, 192);
			this.groupBox1.TabIndex = 195;
			this.groupBox1.TabStop = false;
			this.groupBox1.Text = "Dirección del Domicilio";
			this.groupBox1.Enter += new System.EventHandler(this.groupBox1_Enter);
			// 
			// chkNELicencia
			// 
			this.chkNELicencia.BackColor = System.Drawing.Color.Transparent;
			this.chkNELicencia.Enabled = false;
			this.chkNELicencia.Location = new System.Drawing.Point(568, 64);
			this.chkNELicencia.Name = "chkNELicencia";
			this.chkNELicencia.Size = new System.Drawing.Size(128, 21);
			this.chkNELicencia.TabIndex = 353;
			this.chkNELicencia.Text = "Licencia de Conducir";
			this.chkNELicencia.KeyDown += new System.Windows.Forms.KeyEventHandler(this.chkNELicencia_KeyDown);
			this.chkNELicencia.CheckedChanged += new System.EventHandler(this.chkNELicencia_CheckedChanged);
			// 
			// label44
			// 
			this.label44.AutoSize = true;
			this.label44.BackColor = System.Drawing.Color.Transparent;
			this.label44.Location = new System.Drawing.Point(728, 66);
			this.label44.Name = "label44";
			this.label44.Size = new System.Drawing.Size(104, 17);
			this.label44.TabIndex = 352;
			this.label44.Text = "Fecha de Caducidad";
			this.label44.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// dtpFechaCaducidad
			// 
			this.dtpFechaCaducidad.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			dateButton2.Caption = "Today";
			this.dtpFechaCaducidad.DateButtons.Add(dateButton2);
			this.dtpFechaCaducidad.Enabled = false;
			this.dtpFechaCaducidad.Format = "dd/MM/yyyy";
			this.dtpFechaCaducidad.Location = new System.Drawing.Point(728, 96);
			this.dtpFechaCaducidad.Name = "dtpFechaCaducidad";
			this.dtpFechaCaducidad.NonAutoSizeHeight = 23;
			this.dtpFechaCaducidad.Size = new System.Drawing.Size(104, 21);
			this.dtpFechaCaducidad.TabIndex = 351;
			this.dtpFechaCaducidad.Value = ((object)(resources.GetObject("dtpFechaCaducidad.Value")));
			// 
			// txtTipoDeLicencia
			// 
			this.txtTipoDeLicencia.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtTipoDeLicencia.Enabled = false;
			this.txtTipoDeLicencia.Location = new System.Drawing.Point(632, 96);
			this.txtTipoDeLicencia.MaxLength = 15;
			this.txtTipoDeLicencia.Name = "txtTipoDeLicencia";
			this.txtTipoDeLicencia.Size = new System.Drawing.Size(88, 22);
			this.txtTipoDeLicencia.TabIndex = 354;
			// 
			// lblEstado
			// 
			this.lblEstado.AutoSize = true;
			this.lblEstado.BackColor = System.Drawing.Color.Transparent;
			this.lblEstado.Font = new System.Drawing.Font("Microsoft Sans Serif", 11F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.lblEstado.ForeColor = System.Drawing.Color.Firebrick;
			this.lblEstado.Location = new System.Drawing.Point(840, 96);
			this.lblEstado.Name = "lblEstado";
			this.lblEstado.Size = new System.Drawing.Size(0, 20);
			this.lblEstado.TabIndex = 187;
			this.lblEstado.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// utpEstudios
			// 
			this.utpEstudios.Controls.Add(this.chkDecCuartoAcum);
			this.utpEstudios.Controls.Add(this.cmbRegion);
			this.utpEstudios.Controls.Add(this.cmbRangos);
			this.utpEstudios.Controls.Add(this.label51);
			this.utpEstudios.Controls.Add(this.txtQuincena);
			this.utpEstudios.Controls.Add(this.chkNEQuincena);
			this.utpEstudios.Controls.Add(this.chkNEAplicaRol);
			this.utpEstudios.Controls.Add(this.chkDerechoFondosDeReserva);
			this.utpEstudios.Controls.Add(this.chkNEVendedor);
			this.utpEstudios.Controls.Add(this.chkNEComisiona);
			this.utpEstudios.Controls.Add(this.label35);
			this.utpEstudios.Controls.Add(this.txtNEBonos);
			this.utpEstudios.Controls.Add(this.txtNEOrdinarias);
			this.utpEstudios.Controls.Add(this.label34);
			this.utpEstudios.Controls.Add(this.cmbNEBancos);
			this.utpEstudios.Controls.Add(this.txtNEResponsabilidades);
			this.utpEstudios.Controls.Add(this.label30);
			this.utpEstudios.Controls.Add(this.txtNEFunciones);
			this.utpEstudios.Controls.Add(this.label29);
			this.utpEstudios.Controls.Add(this.txtNEComisiones);
			this.utpEstudios.Controls.Add(this.label28);
			this.utpEstudios.Controls.Add(this.txtNEBono);
			this.utpEstudios.Controls.Add(this.label27);
			this.utpEstudios.Controls.Add(this.chkNETCC);
			this.utpEstudios.Controls.Add(this.txtRetencionJudicial);
			this.utpEstudios.Controls.Add(this.chkNERetencionJudicial);
			this.utpEstudios.Controls.Add(this.txtGastoVestimenta);
			this.utpEstudios.Controls.Add(this.txtGastoEducacion);
			this.utpEstudios.Controls.Add(this.txtGastoAlimentacion);
			this.utpEstudios.Controls.Add(this.txtGastoSalud);
			this.utpEstudios.Controls.Add(this.txtGastoVivienda);
			this.utpEstudios.Controls.Add(this.chkNEImpuestoRenta);
			this.utpEstudios.Controls.Add(this.txtCuenta);
			this.utpEstudios.Controls.Add(this.label25);
			this.utpEstudios.Controls.Add(this.label26);
			this.utpEstudios.Controls.Add(this.optNETipoDePago);
			this.utpEstudios.Controls.Add(this.txtTotalIngresos);
			this.utpEstudios.Controls.Add(this.label23);
			this.utpEstudios.Controls.Add(this.txtNEHorasExtras);
			this.utpEstudios.Controls.Add(this.label24);
			this.utpEstudios.Controls.Add(this.txtNEMovilizacion);
			this.utpEstudios.Controls.Add(this.label17);
			this.utpEstudios.Controls.Add(this.txtNESueldo);
			this.utpEstudios.Controls.Add(this.label16);
			this.utpEstudios.Controls.Add(this.chkDecTerAcum);
			this.utpEstudios.Controls.Add(this.chkFondoReserva);
			this.utpEstudios.Controls.Add(this.label15);
			this.utpEstudios.Controls.Add(this.chkNEAporteIESS);
			this.utpEstudios.Controls.Add(this.label63);
			this.utpEstudios.Controls.Add(this.cmbNETipoContrato);
			this.utpEstudios.Controls.Add(this.optNESistema);
			this.utpEstudios.Controls.Add(this.optNERolesProservicios);
			this.utpEstudios.Controls.Add(this.cmbNECodigoSectorial);
			this.utpEstudios.Controls.Add(this.label14);
			this.utpEstudios.Controls.Add(this.cmbNECargo);
			this.utpEstudios.Controls.Add(this.label13);
			this.utpEstudios.Controls.Add(this.cmbNEDepartamento);
			this.utpEstudios.Controls.Add(this.label12);
			this.utpEstudios.Controls.Add(this.cmbNECentroCosto);
			this.utpEstudios.Controls.Add(this.label33);
			this.utpEstudios.Controls.Add(this.txtAfiliacion);
			this.utpEstudios.Controls.Add(this.label48);
			this.utpEstudios.Controls.Add(this.label49);
			this.utpEstudios.Controls.Add(this.label50);
			this.utpEstudios.Controls.Add(this.label47);
			this.utpEstudios.Controls.Add(this.label46);
			this.utpEstudios.Controls.Add(this.label41);
			this.utpEstudios.Location = new System.Drawing.Point(1, 20);
			this.utpEstudios.Name = "utpEstudios";
			this.utpEstudios.Size = new System.Drawing.Size(1136, 315);
			this.utpEstudios.Paint += new System.Windows.Forms.PaintEventHandler(this.utpEstudios_Paint);
			// 
			// chkDecCuartoAcum
			// 
			this.chkDecCuartoAcum.BackColor = System.Drawing.Color.Transparent;
			this.chkDecCuartoAcum.Enabled = false;
			this.chkDecCuartoAcum.Location = new System.Drawing.Point(968, 128);
			this.chkDecCuartoAcum.Name = "chkDecCuartoAcum";
			this.chkDecCuartoAcum.Size = new System.Drawing.Size(136, 22);
			this.chkDecCuartoAcum.TabIndex = 394;
			this.chkDecCuartoAcum.Text = "14 Acumulado";
			// 
			// cmbRegion
			// 
			this.cmbRegion.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbRegion.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbRegion.DataSource = this.ultraDataSource34;
			ultraGridColumn127.Header.VisiblePosition = 0;
			ultraGridColumn127.Hidden = true;
			ultraGridColumn128.Header.VisiblePosition = 1;
			ultraGridBand29.Columns.AddRange(new object[] {
																											ultraGridColumn127,
																											ultraGridColumn128});
			this.cmbRegion.DisplayLayout.BandsSerializer.Add(ultraGridBand29);
			this.cmbRegion.DisplayMember = "Region";
			this.cmbRegion.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbRegion.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbRegion.Enabled = false;
			this.cmbRegion.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbRegion.Location = new System.Drawing.Point(632, 168);
			this.cmbRegion.Name = "cmbRegion";
			this.cmbRegion.Size = new System.Drawing.Size(176, 22);
			this.cmbRegion.TabIndex = 392;
			this.cmbRegion.ValueMember = "idRegion";
			// 
			// ultraDataSource34
			// 
			ultraDataColumn122.DataType = typeof(int);
			this.ultraDataSource34.Band.Columns.AddRange(new object[] {
																																	ultraDataColumn121,
																																	ultraDataColumn122});
			// 
			// cmbRangos
			// 
			this.cmbRangos.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbRangos.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbRangos.DataSource = this.ultraDataSource27;
			ultraGridColumn129.Header.VisiblePosition = 0;
			ultraGridColumn129.Hidden = true;
			ultraGridColumn130.Header.VisiblePosition = 1;
			ultraGridColumn130.Width = 150;
			ultraGridBand30.Columns.AddRange(new object[] {
																											ultraGridColumn129,
																											ultraGridColumn130});
			this.cmbRangos.DisplayLayout.BandsSerializer.Add(ultraGridBand30);
			this.cmbRangos.DisplayMember = "Rango";
			this.cmbRangos.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbRangos.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbRangos.Enabled = false;
			this.cmbRangos.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbRangos.Location = new System.Drawing.Point(79, 48);
			this.cmbRangos.Name = "cmbRangos";
			this.cmbRangos.Size = new System.Drawing.Size(144, 22);
			this.cmbRangos.TabIndex = 391;
			this.cmbRangos.ValueMember = "idCom_Rango";
			this.cmbRangos.KeyDown += new System.Windows.Forms.KeyEventHandler(this.cmbRangos_KeyDown);
			this.cmbRangos.InitializeLayout += new Infragistics.Win.UltraWinGrid.InitializeLayoutEventHandler(this.ultraCombo1_InitializeLayout);
			// 
			// ultraDataSource27
			// 
			ultraDataColumn123.DataType = typeof(int);
			this.ultraDataSource27.Band.Columns.AddRange(new object[] {
																																	ultraDataColumn123,
																																	ultraDataColumn124});
			// 
			// label51
			// 
			this.label51.AutoSize = true;
			this.label51.BackColor = System.Drawing.Color.Transparent;
			this.label51.Location = new System.Drawing.Point(16, 50);
			this.label51.Name = "label51";
			this.label51.Size = new System.Drawing.Size(52, 17);
			this.label51.TabIndex = 390;
			this.label51.Text = "Categoria";
			// 
			// txtQuincena
			// 
			this.txtQuincena.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtQuincena.Enabled = false;
			this.txtQuincena.FormatString = "#,##0.00";
			this.txtQuincena.Location = new System.Drawing.Point(727, 96);
			this.txtQuincena.MaxValue = 3600;
			this.txtQuincena.MinValue = 0;
			this.txtQuincena.Name = "txtQuincena";
			this.txtQuincena.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtQuincena.PromptChar = ' ';
			this.txtQuincena.Size = new System.Drawing.Size(88, 22);
			this.txtQuincena.TabIndex = 389;
			this.txtQuincena.Visible = false;
			this.txtQuincena.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtQuincena_KeyDown);
			this.txtQuincena.ValueChanged += new System.EventHandler(this.txtQuincena_ValueChanged);
			// 
			// chkNEQuincena
			// 
			this.chkNEQuincena.BackColor = System.Drawing.Color.Transparent;
			this.chkNEQuincena.Enabled = false;
			this.chkNEQuincena.Location = new System.Drawing.Point(711, 78);
			this.chkNEQuincena.Name = "chkNEQuincena";
			this.chkNEQuincena.Size = new System.Drawing.Size(72, 21);
			this.chkNEQuincena.TabIndex = 388;
			this.chkNEQuincena.Text = "Quincena";
			this.chkNEQuincena.Visible = false;
			this.chkNEQuincena.KeyDown += new System.Windows.Forms.KeyEventHandler(this.chkNEQuincena_KeyDown);
			this.chkNEQuincena.CheckedChanged += new System.EventHandler(this.chkNEQuincena_CheckedChanged);
			// 
			// chkNEAplicaRol
			// 
			this.chkNEAplicaRol.BackColor = System.Drawing.Color.Transparent;
			this.chkNEAplicaRol.Enabled = false;
			this.chkNEAplicaRol.Location = new System.Drawing.Point(840, 216);
			this.chkNEAplicaRol.Name = "chkNEAplicaRol";
			this.chkNEAplicaRol.Size = new System.Drawing.Size(96, 21);
			this.chkNEAplicaRol.TabIndex = 387;
			this.chkNEAplicaRol.Text = "Aplica a Rol";
			this.chkNEAplicaRol.KeyDown += new System.Windows.Forms.KeyEventHandler(this.chkNEAplicaRol_KeyDown);
			this.chkNEAplicaRol.CheckedChanged += new System.EventHandler(this.chkNEAplicaRol_CheckedChanged);
			// 
			// chkDerechoFondosDeReserva
			// 
			this.chkDerechoFondosDeReserva.BackColor = System.Drawing.Color.Transparent;
			this.chkDerechoFondosDeReserva.Enabled = false;
			this.chkDerechoFondosDeReserva.Location = new System.Drawing.Point(840, 160);
			this.chkDerechoFondosDeReserva.Name = "chkDerechoFondosDeReserva";
			this.chkDerechoFondosDeReserva.Size = new System.Drawing.Size(176, 21);
			this.chkDerechoFondosDeReserva.TabIndex = 386;
			this.chkDerechoFondosDeReserva.Text = "Derecho a Fondos de Reserva";
			this.chkDerechoFondosDeReserva.KeyDown += new System.Windows.Forms.KeyEventHandler(this.chkDerechoFondosDeReserva_KeyDown);
			// 
			// chkNEVendedor
			// 
			this.chkNEVendedor.BackColor = System.Drawing.Color.Transparent;
			this.chkNEVendedor.Enabled = false;
			this.chkNEVendedor.Location = new System.Drawing.Point(936, 216);
			this.chkNEVendedor.Name = "chkNEVendedor";
			this.chkNEVendedor.Size = new System.Drawing.Size(96, 21);
			this.chkNEVendedor.TabIndex = 385;
			this.chkNEVendedor.Text = "Vendedor";
			this.chkNEVendedor.KeyDown += new System.Windows.Forms.KeyEventHandler(this.chkNEVendedor_KeyDown);
			this.chkNEVendedor.CheckedChanged += new System.EventHandler(this.chkNEVendedor_CheckedChanged);
			// 
			// chkNEComisiona
			// 
			this.chkNEComisiona.BackColor = System.Drawing.Color.Transparent;
			this.chkNEComisiona.Enabled = false;
			this.chkNEComisiona.Location = new System.Drawing.Point(1032, 216);
			this.chkNEComisiona.Name = "chkNEComisiona";
			this.chkNEComisiona.Size = new System.Drawing.Size(96, 21);
			this.chkNEComisiona.TabIndex = 383;
			this.chkNEComisiona.Text = "Comisiona";
			this.chkNEComisiona.KeyDown += new System.Windows.Forms.KeyEventHandler(this.chkNEComisiona_KeyDown);
			// 
			// label35
			// 
			this.label35.AutoSize = true;
			this.label35.BackColor = System.Drawing.Color.Transparent;
			this.label35.Location = new System.Drawing.Point(111, 80);
			this.label35.Name = "label35";
			this.label35.Size = new System.Drawing.Size(29, 17);
			this.label35.TabIndex = 382;
			this.label35.Text = "Bono";
			this.label35.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// txtNEBonos
			// 
			this.txtNEBonos.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			this.txtNEBonos.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtNEBonos.Enabled = false;
			this.txtNEBonos.FormatString = "#,##0.00";
			this.txtNEBonos.Location = new System.Drawing.Point(111, 96);
			this.txtNEBonos.Name = "txtNEBonos";
			this.txtNEBonos.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtNEBonos.PromptChar = ' ';
			this.txtNEBonos.Size = new System.Drawing.Size(88, 22);
			this.txtNEBonos.TabIndex = 381;
			this.txtNEBonos.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtNEBonos_KeyDown);
			this.txtNEBonos.Click += new System.EventHandler(this.txtNEBonos_Click);
			this.txtNEBonos.Validated += new System.EventHandler(this.txtNEBonos_Validated);
			this.txtNEBonos.Enter += new System.EventHandler(this.txtNEBonos_Enter);
			// 
			// txtNEOrdinarias
			// 
			this.txtNEOrdinarias.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			this.txtNEOrdinarias.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtNEOrdinarias.Enabled = false;
			this.txtNEOrdinarias.FormatString = "#,##0.00";
			this.txtNEOrdinarias.Location = new System.Drawing.Point(303, 96);
			this.txtNEOrdinarias.Name = "txtNEOrdinarias";
			this.txtNEOrdinarias.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtNEOrdinarias.PromptChar = ' ';
			this.txtNEOrdinarias.Size = new System.Drawing.Size(88, 22);
			this.txtNEOrdinarias.TabIndex = 380;
			this.txtNEOrdinarias.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtNEOrdinarias_KeyDown);
			this.txtNEOrdinarias.Click += new System.EventHandler(this.txtNEOrdinarias_Click);
			this.txtNEOrdinarias.Validated += new System.EventHandler(this.txtNEOrdinarias_Validated);
			this.txtNEOrdinarias.Enter += new System.EventHandler(this.txtNEOrdinarias_Enter);
			// 
			// label34
			// 
			this.label34.AutoSize = true;
			this.label34.BackColor = System.Drawing.Color.Transparent;
			this.label34.Location = new System.Drawing.Point(399, 80);
			this.label34.Name = "label34";
			this.label34.Size = new System.Drawing.Size(81, 17);
			this.label34.TabIndex = 379;
			this.label34.Text = "Horas ExtraOrd";
			this.label34.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// cmbNEBancos
			// 
			this.cmbNEBancos.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbNEBancos.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbNEBancos.DataSource = this.ultraDataSource10;
			ultraGridColumn131.Header.VisiblePosition = 0;
			ultraGridColumn131.Hidden = true;
			ultraGridColumn132.Header.VisiblePosition = 1;
			ultraGridColumn132.Width = 208;
			ultraGridBand31.Columns.AddRange(new object[] {
																											ultraGridColumn131,
																											ultraGridColumn132});
			this.cmbNEBancos.DisplayLayout.BandsSerializer.Add(ultraGridBand31);
			this.cmbNEBancos.DisplayMember = "Nombre";
			this.cmbNEBancos.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbNEBancos.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbNEBancos.Enabled = false;
			this.cmbNEBancos.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbNEBancos.Location = new System.Drawing.Point(432, 122);
			this.cmbNEBancos.Name = "cmbNEBancos";
			this.cmbNEBancos.Size = new System.Drawing.Size(192, 22);
			this.cmbNEBancos.TabIndex = 378;
			this.cmbNEBancos.ValueMember = "idBanco";
			this.cmbNEBancos.KeyDown += new System.Windows.Forms.KeyEventHandler(this.cmbNEBancos_KeyDown);
			// 
			// ultraDataSource10
			// 
			ultraDataColumn125.DataType = typeof(int);
			this.ultraDataSource10.Band.Columns.AddRange(new object[] {
																																	ultraDataColumn125,
																																	ultraDataColumn126});
			// 
			// txtNEResponsabilidades
			// 
			this.txtNEResponsabilidades.BorderStyle = Infragistics.Win.UIElementBorderStyle.Solid;
			this.txtNEResponsabilidades.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtNEResponsabilidades.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtNEResponsabilidades.Enabled = false;
			this.txtNEResponsabilidades.Location = new System.Drawing.Point(536, 256);
			this.txtNEResponsabilidades.MaxLength = 1000;
			this.txtNEResponsabilidades.Multiline = true;
			this.txtNEResponsabilidades.Name = "txtNEResponsabilidades";
			this.txtNEResponsabilidades.Size = new System.Drawing.Size(280, 56);
			this.txtNEResponsabilidades.TabIndex = 377;
			this.txtNEResponsabilidades.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtNEResponsabilidades_KeyDown);
			this.txtNEResponsabilidades.ValueChanged += new System.EventHandler(this.txtNEResponsabilidades_ValueChanged);
			// 
			// label30
			// 
			this.label30.AutoSize = true;
			this.label30.BackColor = System.Drawing.Color.Transparent;
			this.label30.Location = new System.Drawing.Point(424, 256);
			this.label30.Name = "label30";
			this.label30.Size = new System.Drawing.Size(95, 17);
			this.label30.TabIndex = 376;
			this.label30.Text = "Responsabilidades";
			this.label30.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// txtNEFunciones
			// 
			this.txtNEFunciones.BorderStyle = Infragistics.Win.UIElementBorderStyle.Solid;
			this.txtNEFunciones.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtNEFunciones.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtNEFunciones.Enabled = false;
			this.txtNEFunciones.Location = new System.Drawing.Point(80, 256);
			this.txtNEFunciones.MaxLength = 1000;
			this.txtNEFunciones.Multiline = true;
			this.txtNEFunciones.Name = "txtNEFunciones";
			this.txtNEFunciones.Size = new System.Drawing.Size(328, 56);
			this.txtNEFunciones.TabIndex = 375;
			this.txtNEFunciones.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtNEFunciones_KeyDown);
			// 
			// label29
			// 
			this.label29.AutoSize = true;
			this.label29.BackColor = System.Drawing.Color.Transparent;
			this.label29.Location = new System.Drawing.Point(15, 256);
			this.label29.Name = "label29";
			this.label29.Size = new System.Drawing.Size(54, 17);
			this.label29.TabIndex = 374;
			this.label29.Text = "Funciones";
			this.label29.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// txtNEComisiones
			// 
			this.txtNEComisiones.BorderStyle = Infragistics.Win.UIElementBorderStyle.Solid;
			this.txtNEComisiones.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtNEComisiones.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtNEComisiones.Enabled = false;
			this.txtNEComisiones.Location = new System.Drawing.Point(536, 192);
			this.txtNEComisiones.MaxLength = 1000;
			this.txtNEComisiones.Multiline = true;
			this.txtNEComisiones.Name = "txtNEComisiones";
			this.txtNEComisiones.Size = new System.Drawing.Size(280, 56);
			this.txtNEComisiones.TabIndex = 373;
			this.txtNEComisiones.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtNEComisiones_KeyDown);
			// 
			// label28
			// 
			this.label28.AutoSize = true;
			this.label28.BackColor = System.Drawing.Color.Transparent;
			this.label28.Location = new System.Drawing.Point(424, 192);
			this.label28.Name = "label28";
			this.label28.Size = new System.Drawing.Size(60, 17);
			this.label28.TabIndex = 372;
			this.label28.Text = "Comisiones";
			this.label28.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// txtNEBono
			// 
			this.txtNEBono.BorderStyle = Infragistics.Win.UIElementBorderStyle.Solid;
			this.txtNEBono.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtNEBono.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtNEBono.Enabled = false;
			this.txtNEBono.Location = new System.Drawing.Point(80, 192);
			this.txtNEBono.MaxLength = 1000;
			this.txtNEBono.Multiline = true;
			this.txtNEBono.Name = "txtNEBono";
			this.txtNEBono.Size = new System.Drawing.Size(328, 56);
			this.txtNEBono.TabIndex = 371;
			this.txtNEBono.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtNEBono_KeyDown);
			// 
			// label27
			// 
			this.label27.AutoSize = true;
			this.label27.BackColor = System.Drawing.Color.Transparent;
			this.label27.Location = new System.Drawing.Point(16, 192);
			this.label27.Name = "label27";
			this.label27.Size = new System.Drawing.Size(29, 17);
			this.label27.TabIndex = 370;
			this.label27.Text = "Bono";
			this.label27.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// chkNETCC
			// 
			this.chkNETCC.BackColor = System.Drawing.Color.Transparent;
			this.chkNETCC.Enabled = false;
			this.chkNETCC.Location = new System.Drawing.Point(840, 240);
			this.chkNETCC.Name = "chkNETCC";
			this.chkNETCC.Size = new System.Drawing.Size(176, 21);
			this.chkNETCC.TabIndex = 369;
			this.chkNETCC.Text = "Tarjeta de Crédito Corporativa";
			this.chkNETCC.KeyDown += new System.Windows.Forms.KeyEventHandler(this.chkNETCC_KeyDown);
			this.chkNETCC.CheckedChanged += new System.EventHandler(this.chkNETCC_CheckedChanged);
			// 
			// txtRetencionJudicial
			// 
			this.txtRetencionJudicial.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtRetencionJudicial.Enabled = false;
			this.txtRetencionJudicial.FormatString = "#,##0.00";
			this.txtRetencionJudicial.Location = new System.Drawing.Point(607, 96);
			this.txtRetencionJudicial.MaxValue = 3600;
			this.txtRetencionJudicial.MinValue = 0;
			this.txtRetencionJudicial.Name = "txtRetencionJudicial";
			this.txtRetencionJudicial.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtRetencionJudicial.PromptChar = ' ';
			this.txtRetencionJudicial.Size = new System.Drawing.Size(88, 22);
			this.txtRetencionJudicial.TabIndex = 368;
			this.txtRetencionJudicial.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtRetencionJudicial_KeyDown);
			this.txtRetencionJudicial.Click += new System.EventHandler(this.txtRetencionJudicial_Click);
			this.txtRetencionJudicial.Validated += new System.EventHandler(this.txtRetencionJudicial_Validated);
			this.txtRetencionJudicial.ValueChanged += new System.EventHandler(this.txtRetencionJudicial_ValueChanged);
			this.txtRetencionJudicial.Enter += new System.EventHandler(this.txtRetencionJudicial_Enter);
			// 
			// chkNERetencionJudicial
			// 
			this.chkNERetencionJudicial.BackColor = System.Drawing.Color.Transparent;
			this.chkNERetencionJudicial.Enabled = false;
			this.chkNERetencionJudicial.Location = new System.Drawing.Point(591, 78);
			this.chkNERetencionJudicial.Name = "chkNERetencionJudicial";
			this.chkNERetencionJudicial.Size = new System.Drawing.Size(120, 21);
			this.chkNERetencionJudicial.TabIndex = 367;
			this.chkNERetencionJudicial.Text = "Retención Judicial";
			this.chkNERetencionJudicial.KeyDown += new System.Windows.Forms.KeyEventHandler(this.chkNERetencionJudicial_KeyDown);
			this.chkNERetencionJudicial.CheckedChanged += new System.EventHandler(this.chkNERetencionJudicial_CheckedChanged);
			// 
			// txtGastoVestimenta
			// 
			this.txtGastoVestimenta.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			this.txtGastoVestimenta.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtGastoVestimenta.Enabled = false;
			this.txtGastoVestimenta.FormatString = "#,##0.00";
			this.txtGastoVestimenta.Location = new System.Drawing.Point(536, 168);
			this.txtGastoVestimenta.Name = "txtGastoVestimenta";
			this.txtGastoVestimenta.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtGastoVestimenta.PromptChar = ' ';
			this.txtGastoVestimenta.Size = new System.Drawing.Size(88, 22);
			this.txtGastoVestimenta.TabIndex = 366;
			this.txtGastoVestimenta.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtGastoVestimenta_KeyDown);
			this.txtGastoVestimenta.Click += new System.EventHandler(this.txtGastoVestimenta_Click);
			this.txtGastoVestimenta.Validated += new System.EventHandler(this.txtGastoVestimenta_Validated);
			this.txtGastoVestimenta.Enter += new System.EventHandler(this.txtGastoVestimenta_Enter);
			// 
			// txtGastoEducacion
			// 
			this.txtGastoEducacion.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			this.txtGastoEducacion.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtGastoEducacion.Enabled = false;
			this.txtGastoEducacion.FormatString = "#,##0.00";
			this.txtGastoEducacion.Location = new System.Drawing.Point(440, 168);
			this.txtGastoEducacion.Name = "txtGastoEducacion";
			this.txtGastoEducacion.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtGastoEducacion.PromptChar = ' ';
			this.txtGastoEducacion.Size = new System.Drawing.Size(88, 22);
			this.txtGastoEducacion.TabIndex = 365;
			this.txtGastoEducacion.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtGastoEducacion_KeyDown);
			this.txtGastoEducacion.Click += new System.EventHandler(this.txtGastoEducacion_Click);
			this.txtGastoEducacion.Validated += new System.EventHandler(this.txtGastoEducacion_Validated);
			this.txtGastoEducacion.Enter += new System.EventHandler(this.txtGastoEducacion_Enter);
			// 
			// txtGastoAlimentacion
			// 
			this.txtGastoAlimentacion.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			this.txtGastoAlimentacion.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtGastoAlimentacion.Enabled = false;
			this.txtGastoAlimentacion.FormatString = "#,##0.00";
			this.txtGastoAlimentacion.Location = new System.Drawing.Point(344, 168);
			this.txtGastoAlimentacion.Name = "txtGastoAlimentacion";
			this.txtGastoAlimentacion.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtGastoAlimentacion.PromptChar = ' ';
			this.txtGastoAlimentacion.Size = new System.Drawing.Size(88, 22);
			this.txtGastoAlimentacion.TabIndex = 364;
			this.txtGastoAlimentacion.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtGastoAlimentacion_KeyDown);
			this.txtGastoAlimentacion.Click += new System.EventHandler(this.txtGastoAlimentacion_Click);
			this.txtGastoAlimentacion.Validated += new System.EventHandler(this.txtGastoAlimentacion_Validated);
			this.txtGastoAlimentacion.Enter += new System.EventHandler(this.txtGastoAlimentacion_Enter);
			// 
			// txtGastoSalud
			// 
			this.txtGastoSalud.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			this.txtGastoSalud.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtGastoSalud.Enabled = false;
			this.txtGastoSalud.FormatString = "#,##0.00";
			this.txtGastoSalud.Location = new System.Drawing.Point(248, 168);
			this.txtGastoSalud.Name = "txtGastoSalud";
			this.txtGastoSalud.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtGastoSalud.PromptChar = ' ';
			this.txtGastoSalud.Size = new System.Drawing.Size(88, 22);
			this.txtGastoSalud.TabIndex = 363;
			this.txtGastoSalud.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtGastoSalud_KeyDown);
			this.txtGastoSalud.Click += new System.EventHandler(this.txtGastoSalud_Click);
			this.txtGastoSalud.Validated += new System.EventHandler(this.txtGastoSalud_Validated);
			this.txtGastoSalud.Enter += new System.EventHandler(this.txtGastoSalud_Enter);
			// 
			// txtGastoVivienda
			// 
			this.txtGastoVivienda.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			this.txtGastoVivienda.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtGastoVivienda.Enabled = false;
			this.txtGastoVivienda.FormatString = "#,##0.00";
			this.txtGastoVivienda.Location = new System.Drawing.Point(152, 168);
			this.txtGastoVivienda.Name = "txtGastoVivienda";
			this.txtGastoVivienda.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtGastoVivienda.PromptChar = ' ';
			this.txtGastoVivienda.Size = new System.Drawing.Size(88, 22);
			this.txtGastoVivienda.TabIndex = 362;
			this.txtGastoVivienda.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtGastoVivienda_KeyDown);
			this.txtGastoVivienda.Click += new System.EventHandler(this.txtGastoVivienda_Click);
			this.txtGastoVivienda.Validated += new System.EventHandler(this.txtGastoVivienda_Validated);
			this.txtGastoVivienda.Enter += new System.EventHandler(this.txtGastoVivienda_Enter);
			// 
			// chkNEImpuestoRenta
			// 
			this.chkNEImpuestoRenta.BackColor = System.Drawing.Color.Transparent;
			this.chkNEImpuestoRenta.Enabled = false;
			this.chkNEImpuestoRenta.Location = new System.Drawing.Point(15, 144);
			this.chkNEImpuestoRenta.Name = "chkNEImpuestoRenta";
			this.chkNEImpuestoRenta.Size = new System.Drawing.Size(136, 21);
			this.chkNEImpuestoRenta.TabIndex = 361;
			this.chkNEImpuestoRenta.Text = "Impuesto a la Renta";
			this.chkNEImpuestoRenta.KeyDown += new System.Windows.Forms.KeyEventHandler(this.chkNEImpuestoRenta_KeyDown);
			this.chkNEImpuestoRenta.CheckedChanged += new System.EventHandler(this.chkNEImpuestoRenta_CheckedChanged);
			// 
			// txtCuenta
			// 
			this.txtCuenta.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtCuenta.Enabled = false;
			this.txtCuenta.Location = new System.Drawing.Point(696, 122);
			this.txtCuenta.MaxLength = 15;
			this.txtCuenta.Name = "txtCuenta";
			this.txtCuenta.Size = new System.Drawing.Size(120, 22);
			this.txtCuenta.TabIndex = 358;
			this.txtCuenta.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtCuenta_KeyDown);
			this.txtCuenta.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.txtCuenta_KeyPress);
			// 
			// label25
			// 
			this.label25.AutoSize = true;
			this.label25.BackColor = System.Drawing.Color.Transparent;
			this.label25.Location = new System.Drawing.Point(632, 124);
			this.label25.Name = "label25";
			this.label25.Size = new System.Drawing.Size(54, 17);
			this.label25.TabIndex = 357;
			this.label25.Text = "Cuenta N.";
			this.label25.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// label26
			// 
			this.label26.AutoSize = true;
			this.label26.BackColor = System.Drawing.Color.Transparent;
			this.label26.Location = new System.Drawing.Point(384, 124);
			this.label26.Name = "label26";
			this.label26.Size = new System.Drawing.Size(34, 17);
			this.label26.TabIndex = 356;
			this.label26.Text = "Banco";
			this.label26.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// optNETipoDePago
			// 
			this.optNETipoDePago.CheckedIndex = 0;
			this.optNETipoDePago.Enabled = false;
			this.optNETipoDePago.ItemAppearance = appearance77;
			valueListItem8.DataValue = 1;
			valueListItem8.DisplayText = "Transf. Ahorros";
			valueListItem9.DataValue = 2;
			valueListItem9.DisplayText = "Transf. Corriente";
			valueListItem10.DataValue = 3;
			valueListItem10.DisplayText = "Transf. Virtual";
			valueListItem11.DataValue = 4;
			valueListItem11.DisplayText = "Cheque";
			this.optNETipoDePago.Items.Add(valueListItem8);
			this.optNETipoDePago.Items.Add(valueListItem9);
			this.optNETipoDePago.Items.Add(valueListItem10);
			this.optNETipoDePago.Items.Add(valueListItem11);
			this.optNETipoDePago.Location = new System.Drawing.Point(16, 120);
			this.optNETipoDePago.Name = "optNETipoDePago";
			this.optNETipoDePago.Size = new System.Drawing.Size(360, 24);
			this.optNETipoDePago.TabIndex = 355;
			this.optNETipoDePago.Text = "Transf. Ahorros";
			this.optNETipoDePago.KeyDown += new System.Windows.Forms.KeyEventHandler(this.optNETipoDePago_KeyDown);
			this.optNETipoDePago.ValueChanged += new System.EventHandler(this.optNETipoDePago_ValueChanged);
			// 
			// txtTotalIngresos
			// 
			this.txtTotalIngresos.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			this.txtTotalIngresos.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtTotalIngresos.Enabled = false;
			this.txtTotalIngresos.FormatString = "#,##0.00";
			this.txtTotalIngresos.Location = new System.Drawing.Point(495, 96);
			this.txtTotalIngresos.Name = "txtTotalIngresos";
			this.txtTotalIngresos.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtTotalIngresos.PromptChar = ' ';
			this.txtTotalIngresos.Size = new System.Drawing.Size(88, 22);
			this.txtTotalIngresos.TabIndex = 354;
			// 
			// label23
			// 
			this.label23.AutoSize = true;
			this.label23.BackColor = System.Drawing.Color.Transparent;
			this.label23.Location = new System.Drawing.Point(495, 80);
			this.label23.Name = "label23";
			this.label23.Size = new System.Drawing.Size(92, 17);
			this.label23.TabIndex = 353;
			this.label23.Text = "Total de Ingresos";
			this.label23.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// txtNEHorasExtras
			// 
			this.txtNEHorasExtras.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			this.txtNEHorasExtras.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtNEHorasExtras.Enabled = false;
			this.txtNEHorasExtras.FormatString = "#,##0.00";
			this.txtNEHorasExtras.Location = new System.Drawing.Point(399, 96);
			this.txtNEHorasExtras.Name = "txtNEHorasExtras";
			this.txtNEHorasExtras.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtNEHorasExtras.PromptChar = ' ';
			this.txtNEHorasExtras.Size = new System.Drawing.Size(88, 22);
			this.txtNEHorasExtras.TabIndex = 352;
			this.txtNEHorasExtras.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtNEHorasExtras_KeyDown);
			this.txtNEHorasExtras.Click += new System.EventHandler(this.txtNEHorasExtras_Click);
			this.txtNEHorasExtras.Validated += new System.EventHandler(this.txtNEHorasExtras_Validated);
			this.txtNEHorasExtras.Enter += new System.EventHandler(this.txtNEHorasExtras_Enter);
			// 
			// label24
			// 
			this.label24.AutoSize = true;
			this.label24.BackColor = System.Drawing.Color.Transparent;
			this.label24.Location = new System.Drawing.Point(303, 80);
			this.label24.Name = "label24";
			this.label24.Size = new System.Drawing.Size(87, 17);
			this.label24.TabIndex = 351;
			this.label24.Text = "Horas Ordinarias";
			this.label24.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// txtNEMovilizacion
			// 
			this.txtNEMovilizacion.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			this.txtNEMovilizacion.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtNEMovilizacion.Enabled = false;
			this.txtNEMovilizacion.FormatString = "#,##0.00";
			this.txtNEMovilizacion.Location = new System.Drawing.Point(207, 96);
			this.txtNEMovilizacion.Name = "txtNEMovilizacion";
			this.txtNEMovilizacion.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtNEMovilizacion.PromptChar = ' ';
			this.txtNEMovilizacion.Size = new System.Drawing.Size(88, 22);
			this.txtNEMovilizacion.TabIndex = 350;
			this.txtNEMovilizacion.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtNEMovilizacion_KeyDown);
			this.txtNEMovilizacion.Click += new System.EventHandler(this.txtNEMovilizacion_Click);
			this.txtNEMovilizacion.Validated += new System.EventHandler(this.txtNEMovilizacion_Validated);
			this.txtNEMovilizacion.Enter += new System.EventHandler(this.txtNEMovilizacion_Enter);
			// 
			// label17
			// 
			this.label17.AutoSize = true;
			this.label17.BackColor = System.Drawing.Color.Transparent;
			this.label17.Location = new System.Drawing.Point(207, 80);
			this.label17.Name = "label17";
			this.label17.Size = new System.Drawing.Size(64, 17);
			this.label17.TabIndex = 349;
			this.label17.Text = "Movilización";
			this.label17.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// txtNESueldo
			// 
			this.txtNESueldo.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			this.txtNESueldo.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtNESueldo.Enabled = false;
			this.txtNESueldo.FormatString = "#,##0.00";
			this.txtNESueldo.Location = new System.Drawing.Point(15, 96);
			this.txtNESueldo.Name = "txtNESueldo";
			this.txtNESueldo.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtNESueldo.PromptChar = ' ';
			this.txtNESueldo.Size = new System.Drawing.Size(88, 22);
			this.txtNESueldo.TabIndex = 348;
			this.txtNESueldo.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtNESueldo_KeyDown);
			this.txtNESueldo.Click += new System.EventHandler(this.txtNESueldo_Click);
			this.txtNESueldo.Validated += new System.EventHandler(this.txtNESueldo_Validated);
			this.txtNESueldo.Enter += new System.EventHandler(this.txtNESueldo_Enter);
			// 
			// label16
			// 
			this.label16.AutoSize = true;
			this.label16.BackColor = System.Drawing.Color.Transparent;
			this.label16.Location = new System.Drawing.Point(15, 80);
			this.label16.Name = "label16";
			this.label16.Size = new System.Drawing.Size(38, 17);
			this.label16.TabIndex = 347;
			this.label16.Text = "Sueldo";
			this.label16.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// chkDecTerAcum
			// 
			this.chkDecTerAcum.BackColor = System.Drawing.Color.Transparent;
			this.chkDecTerAcum.Enabled = false;
			this.chkDecTerAcum.Location = new System.Drawing.Point(840, 128);
			this.chkDecTerAcum.Name = "chkDecTerAcum";
			this.chkDecTerAcum.Size = new System.Drawing.Size(136, 22);
			this.chkDecTerAcum.TabIndex = 346;
			this.chkDecTerAcum.Text = "13 Acumulado";
			this.chkDecTerAcum.KeyDown += new System.Windows.Forms.KeyEventHandler(this.chkDecimos_KeyDown);
			// 
			// chkFondoReserva
			// 
			this.chkFondoReserva.BackColor = System.Drawing.Color.Transparent;
			this.chkFondoReserva.Enabled = false;
			this.chkFondoReserva.Location = new System.Drawing.Point(840, 184);
			this.chkFondoReserva.Name = "chkFondoReserva";
			this.chkFondoReserva.Size = new System.Drawing.Size(184, 22);
			this.chkFondoReserva.TabIndex = 345;
			this.chkFondoReserva.Text = "Fondo de Reserva Acumulados";
			this.chkFondoReserva.KeyDown += new System.Windows.Forms.KeyEventHandler(this.chkFondoReserva_KeyDown);
			// 
			// label15
			// 
			this.label15.AutoSize = true;
			this.label15.BackColor = System.Drawing.Color.Transparent;
			this.label15.Location = new System.Drawing.Point(840, 107);
			this.label15.Name = "label15";
			this.label15.Size = new System.Drawing.Size(49, 17);
			this.label15.TabIndex = 94;
			this.label15.Text = "Afiliación";
			// 
			// chkNEAporteIESS
			// 
			this.chkNEAporteIESS.BackColor = System.Drawing.Color.Transparent;
			this.chkNEAporteIESS.Enabled = false;
			this.chkNEAporteIESS.Location = new System.Drawing.Point(824, 80);
			this.chkNEAporteIESS.Name = "chkNEAporteIESS";
			this.chkNEAporteIESS.Size = new System.Drawing.Size(96, 21);
			this.chkNEAporteIESS.TabIndex = 93;
			this.chkNEAporteIESS.Text = "Aporte al IESS";
			this.chkNEAporteIESS.KeyDown += new System.Windows.Forms.KeyEventHandler(this.chkNEAporteIESS_KeyDown);
			this.chkNEAporteIESS.CheckedChanged += new System.EventHandler(this.chkNEAporteIESS_CheckedChanged);
			// 
			// label63
			// 
			this.label63.AutoSize = true;
			this.label63.BackColor = System.Drawing.Color.Transparent;
			this.label63.Location = new System.Drawing.Point(688, 50);
			this.label63.Name = "label63";
			this.label63.Size = new System.Drawing.Size(72, 17);
			this.label63.TabIndex = 88;
			this.label63.Text = "Tipo Contrato";
			this.label63.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// cmbNETipoContrato
			// 
			this.cmbNETipoContrato.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbNETipoContrato.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbNETipoContrato.DataSource = this.ultraDataSource11;
			ultraGridBand32.AddButtonCaption = "RolSucursal";
			ultraGridColumn133.Header.VisiblePosition = 1;
			ultraGridColumn133.Hidden = true;
			ultraGridColumn134.Header.VisiblePosition = 0;
			ultraGridColumn134.Width = 150;
			ultraGridBand32.Columns.AddRange(new object[] {
																											ultraGridColumn133,
																											ultraGridColumn134});
			this.cmbNETipoContrato.DisplayLayout.BandsSerializer.Add(ultraGridBand32);
			this.cmbNETipoContrato.DisplayMember = "TipoContrato";
			this.cmbNETipoContrato.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbNETipoContrato.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbNETipoContrato.Enabled = false;
			this.cmbNETipoContrato.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbNETipoContrato.Location = new System.Drawing.Point(815, 48);
			this.cmbNETipoContrato.Name = "cmbNETipoContrato";
			this.cmbNETipoContrato.Size = new System.Drawing.Size(208, 22);
			this.cmbNETipoContrato.TabIndex = 87;
			this.cmbNETipoContrato.ValueMember = "idTipoContrato";
			this.cmbNETipoContrato.KeyDown += new System.Windows.Forms.KeyEventHandler(this.cmbNETipoContrato_KeyDown);
			// 
			// ultraDataSource11
			// 
			ultraDataColumn127.DataType = typeof(int);
			this.ultraDataSource11.Band.Columns.AddRange(new object[] {
																																	ultraDataColumn127,
																																	ultraDataColumn128});
			// 
			// optNESistema
			// 
			this.optNESistema.CheckedIndex = 0;
			this.optNESistema.Enabled = false;
			this.optNESistema.ItemAppearance = appearance78;
			appearance79.TextVAlign = Infragistics.Win.VAlign.Middle;
			valueListItem12.Appearance = appearance79;
			valueListItem12.DataValue = 1;
			valueListItem12.DisplayText = "Sin Sistema de Salario Neto";
			appearance80.TextVAlign = Infragistics.Win.VAlign.Middle;
			valueListItem13.Appearance = appearance80;
			valueListItem13.DataValue = 2;
			valueListItem13.DisplayText = "Con Sistema de Salario Neto";
			this.optNESistema.Items.Add(valueListItem12);
			this.optNESistema.Items.Add(valueListItem13);
			this.optNESistema.Location = new System.Drawing.Point(312, 46);
			this.optNESistema.Name = "optNESistema";
			this.optNESistema.Size = new System.Drawing.Size(336, 24);
			this.optNESistema.TabIndex = 86;
			this.optNESistema.Text = "Sin Sistema de Salario Neto";
			this.optNESistema.KeyDown += new System.Windows.Forms.KeyEventHandler(this.optNESistema_KeyDown);
			this.optNESistema.ValueChanged += new System.EventHandler(this.optNESistema_ValueChanged);
			// 
			// optNERolesProservicios
			// 
			this.optNERolesProservicios.Enabled = false;
			this.optNERolesProservicios.ItemAppearance = appearance81;
			appearance82.TextVAlign = Infragistics.Win.VAlign.Middle;
			valueListItem14.Appearance = appearance82;
			valueListItem14.DataValue = 1;
			valueListItem14.DisplayText = "Administrador";
			appearance83.TextVAlign = Infragistics.Win.VAlign.Middle;
			valueListItem15.Appearance = appearance83;
			valueListItem15.DataValue = 2;
			valueListItem15.DisplayText = "Revisor";
			appearance84.TextVAlign = Infragistics.Win.VAlign.Middle;
			valueListItem16.Appearance = appearance84;
			valueListItem16.DataValue = 3;
			valueListItem16.DisplayText = "Camarera";
			appearance85.TextVAlign = Infragistics.Win.VAlign.Middle;
			valueListItem17.Appearance = appearance85;
			valueListItem17.DataValue = 0;
			valueListItem17.DisplayText = "Ninguno";
			this.optNERolesProservicios.Items.Add(valueListItem14);
			this.optNERolesProservicios.Items.Add(valueListItem15);
			this.optNERolesProservicios.Items.Add(valueListItem16);
			this.optNERolesProservicios.Items.Add(valueListItem17);
			this.optNERolesProservicios.Location = new System.Drawing.Point(840, 272);
			this.optNERolesProservicios.Name = "optNERolesProservicios";
			this.optNERolesProservicios.Size = new System.Drawing.Size(288, 24);
			this.optNERolesProservicios.SupportThemes = false;
			this.optNERolesProservicios.TabIndex = 87;
			this.optNERolesProservicios.ValueChanged += new System.EventHandler(this.optNERolesProservicios_ValueChanged);
			// 
			// cmbNECodigoSectorial
			// 
			this.cmbNECodigoSectorial.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbNECodigoSectorial.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbNECodigoSectorial.DataSource = this.ultraDataSource22;
			ultraGridColumn135.Header.VisiblePosition = 0;
			ultraGridColumn135.Hidden = true;
			ultraGridColumn136.Header.VisiblePosition = 1;
			ultraGridColumn137.Header.VisiblePosition = 2;
			ultraGridColumn137.Width = 300;
			ultraGridBand33.Columns.AddRange(new object[] {
																											ultraGridColumn135,
																											ultraGridColumn136,
																											ultraGridColumn137});
			this.cmbNECodigoSectorial.DisplayLayout.BandsSerializer.Add(ultraGridBand33);
			this.cmbNECodigoSectorial.DisplayMember = "Actividad";
			this.cmbNECodigoSectorial.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbNECodigoSectorial.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbNECodigoSectorial.Enabled = false;
			this.cmbNECodigoSectorial.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbNECodigoSectorial.Location = new System.Drawing.Point(687, 24);
			this.cmbNECodigoSectorial.Name = "cmbNECodigoSectorial";
			this.cmbNECodigoSectorial.Size = new System.Drawing.Size(336, 22);
			this.cmbNECodigoSectorial.TabIndex = 34;
			this.cmbNECodigoSectorial.ValueMember = "idNominaSectorial";
			this.cmbNECodigoSectorial.KeyDown += new System.Windows.Forms.KeyEventHandler(this.cmbNECodigoSectorial_KeyDown);
			this.cmbNECodigoSectorial.InitializeLayout += new Infragistics.Win.UltraWinGrid.InitializeLayoutEventHandler(this.cmbNECodigoSectorial_InitializeLayout);
			// 
			// ultraDataSource22
			// 
			ultraDataColumn129.DataType = typeof(int);
			this.ultraDataSource22.Band.Columns.AddRange(new object[] {
																																	ultraDataColumn129,
																																	ultraDataColumn130,
																																	ultraDataColumn131});
			// 
			// label14
			// 
			this.label14.AutoSize = true;
			this.label14.BackColor = System.Drawing.Color.Transparent;
			this.label14.Location = new System.Drawing.Point(687, 8);
			this.label14.Name = "label14";
			this.label14.Size = new System.Drawing.Size(85, 17);
			this.label14.TabIndex = 33;
			this.label14.Text = "Codigo Sectorial";
			// 
			// cmbNECargo
			// 
			this.cmbNECargo.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbNECargo.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbNECargo.DataSource = this.ultraDataSource9;
			ultraGridBand34.AddButtonCaption = "RolTipo";
			ultraGridColumn138.Header.VisiblePosition = 0;
			ultraGridColumn138.Hidden = true;
			ultraGridColumn139.Header.VisiblePosition = 1;
			ultraGridColumn139.Width = 208;
			ultraGridColumn140.Header.VisiblePosition = 2;
			ultraGridColumn140.Hidden = true;
			ultraGridBand34.Columns.AddRange(new object[] {
																											ultraGridColumn138,
																											ultraGridColumn139,
																											ultraGridColumn140});
			this.cmbNECargo.DisplayLayout.BandsSerializer.Add(ultraGridBand34);
			this.cmbNECargo.DisplayMember = "Descripcion";
			this.cmbNECargo.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbNECargo.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbNECargo.Enabled = false;
			this.cmbNECargo.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbNECargo.Location = new System.Drawing.Point(463, 24);
			this.cmbNECargo.Name = "cmbNECargo";
			this.cmbNECargo.Size = new System.Drawing.Size(208, 22);
			this.cmbNECargo.TabIndex = 32;
			this.cmbNECargo.ValueMember = "idCargo";
			this.cmbNECargo.KeyDown += new System.Windows.Forms.KeyEventHandler(this.cmbNECargo_KeyDown);
			this.cmbNECargo.ValueChanged += new System.EventHandler(this.cmbNECargo_ValueChanged);
			this.cmbNECargo.InitializeLayout += new Infragistics.Win.UltraWinGrid.InitializeLayoutEventHandler(this.cmbNECargo_InitializeLayout);
			// 
			// ultraDataSource9
			// 
			ultraDataColumn132.DataType = typeof(int);
			ultraDataColumn134.DataType = typeof(int);
			this.ultraDataSource9.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn132,
																																 ultraDataColumn133,
																																 ultraDataColumn134});
			// 
			// label13
			// 
			this.label13.AutoSize = true;
			this.label13.BackColor = System.Drawing.Color.Transparent;
			this.label13.Location = new System.Drawing.Point(463, 8);
			this.label13.Name = "label13";
			this.label13.Size = new System.Drawing.Size(62, 17);
			this.label13.TabIndex = 31;
			this.label13.Text = "Cargo Point";
			// 
			// cmbNEDepartamento
			// 
			this.cmbNEDepartamento.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbNEDepartamento.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbNEDepartamento.DataSource = this.ultraDataSource8;
			ultraGridBand35.AddButtonCaption = "RolTipo";
			ultraGridColumn141.Header.VisiblePosition = 0;
			ultraGridColumn141.Hidden = true;
			ultraGridColumn142.Header.VisiblePosition = 1;
			ultraGridColumn142.Width = 208;
			ultraGridColumn143.Header.VisiblePosition = 2;
			ultraGridColumn143.Hidden = true;
			ultraGridBand35.Columns.AddRange(new object[] {
																											ultraGridColumn141,
																											ultraGridColumn142,
																											ultraGridColumn143});
			this.cmbNEDepartamento.DisplayLayout.BandsSerializer.Add(ultraGridBand35);
			this.cmbNEDepartamento.DisplayMember = "Descripcion";
			this.cmbNEDepartamento.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbNEDepartamento.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbNEDepartamento.Enabled = false;
			this.cmbNEDepartamento.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbNEDepartamento.Location = new System.Drawing.Point(240, 24);
			this.cmbNEDepartamento.Name = "cmbNEDepartamento";
			this.cmbNEDepartamento.Size = new System.Drawing.Size(208, 22);
			this.cmbNEDepartamento.TabIndex = 30;
			this.cmbNEDepartamento.ValueMember = "idDepartamento";
			this.cmbNEDepartamento.KeyDown += new System.Windows.Forms.KeyEventHandler(this.cmbNEDepartamento_KeyDown);
			this.cmbNEDepartamento.ValueChanged += new System.EventHandler(this.cmbNEDepartamento_ValueChanged);
			this.cmbNEDepartamento.InitializeLayout += new Infragistics.Win.UltraWinGrid.InitializeLayoutEventHandler(this.cmbNEDepartamento_InitializeLayout);
			// 
			// ultraDataSource8
			// 
			ultraDataColumn135.DataType = typeof(int);
			ultraDataColumn137.DataType = typeof(int);
			this.ultraDataSource8.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn135,
																																 ultraDataColumn136,
																																 ultraDataColumn137});
			// 
			// label12
			// 
			this.label12.AutoSize = true;
			this.label12.BackColor = System.Drawing.Color.Transparent;
			this.label12.Location = new System.Drawing.Point(239, 8);
			this.label12.Name = "label12";
			this.label12.Size = new System.Drawing.Size(76, 17);
			this.label12.TabIndex = 29;
			this.label12.Text = "Departamento";
			// 
			// cmbNECentroCosto
			// 
			this.cmbNECentroCosto.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbNECentroCosto.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbNECentroCosto.DataSource = this.ultraDataSource7;
			ultraGridBand36.AddButtonCaption = "RolTipo";
			ultraGridColumn144.Header.VisiblePosition = 1;
			ultraGridColumn144.Hidden = true;
			ultraGridColumn144.Width = 85;
			ultraGridColumn145.Header.VisiblePosition = 0;
			ultraGridColumn145.Width = 208;
			ultraGridBand36.Columns.AddRange(new object[] {
																											ultraGridColumn144,
																											ultraGridColumn145});
			this.cmbNECentroCosto.DisplayLayout.BandsSerializer.Add(ultraGridBand36);
			this.cmbNECentroCosto.DisplayMember = "Nombre";
			this.cmbNECentroCosto.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbNECentroCosto.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbNECentroCosto.Enabled = false;
			this.cmbNECentroCosto.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbNECentroCosto.Location = new System.Drawing.Point(15, 24);
			this.cmbNECentroCosto.Name = "cmbNECentroCosto";
			this.cmbNECentroCosto.Size = new System.Drawing.Size(208, 22);
			this.cmbNECentroCosto.TabIndex = 28;
			this.cmbNECentroCosto.ValueMember = "idProyecto";
			this.cmbNECentroCosto.KeyDown += new System.Windows.Forms.KeyEventHandler(this.cmbNECentroCosto_KeyDown);
			// 
			// ultraDataSource7
			// 
			ultraDataColumn138.DataType = typeof(int);
			this.ultraDataSource7.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn138,
																																 ultraDataColumn139});
			// 
			// label33
			// 
			this.label33.AutoSize = true;
			this.label33.BackColor = System.Drawing.Color.Transparent;
			this.label33.Location = new System.Drawing.Point(15, 8);
			this.label33.Name = "label33";
			this.label33.Size = new System.Drawing.Size(86, 17);
			this.label33.TabIndex = 27;
			this.label33.Text = "Centro De Costo";
			// 
			// txtAfiliacion
			// 
			this.txtAfiliacion.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtAfiliacion.Enabled = false;
			this.txtAfiliacion.Location = new System.Drawing.Point(896, 104);
			this.txtAfiliacion.Name = "txtAfiliacion";
			this.txtAfiliacion.Size = new System.Drawing.Size(120, 22);
			this.txtAfiliacion.TabIndex = 344;
			this.txtAfiliacion.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtAfiliacion_KeyDown);
			// 
			// label48
			// 
			this.label48.AutoSize = true;
			this.label48.BackColor = System.Drawing.Color.Transparent;
			this.label48.Location = new System.Drawing.Point(248, 152);
			this.label48.Name = "label48";
			this.label48.Size = new System.Drawing.Size(32, 17);
			this.label48.TabIndex = 51;
			this.label48.Text = "Salud";
			// 
			// label49
			// 
			this.label49.AutoSize = true;
			this.label49.BackColor = System.Drawing.Color.Transparent;
			this.label49.Location = new System.Drawing.Point(440, 152);
			this.label49.Name = "label49";
			this.label49.Size = new System.Drawing.Size(55, 17);
			this.label49.TabIndex = 53;
			this.label49.Text = "Educacion";
			this.label49.TextAlign = System.Drawing.ContentAlignment.TopCenter;
			// 
			// label50
			// 
			this.label50.AutoSize = true;
			this.label50.BackColor = System.Drawing.Color.Transparent;
			this.label50.Location = new System.Drawing.Point(536, 152);
			this.label50.Name = "label50";
			this.label50.Size = new System.Drawing.Size(60, 17);
			this.label50.TabIndex = 55;
			this.label50.Text = "Vestimenta";
			// 
			// label47
			// 
			this.label47.AutoSize = true;
			this.label47.BackColor = System.Drawing.Color.Transparent;
			this.label47.Location = new System.Drawing.Point(344, 152);
			this.label47.Name = "label47";
			this.label47.Size = new System.Drawing.Size(68, 17);
			this.label47.TabIndex = 49;
			this.label47.Text = "Alimentacion";
			// 
			// label46
			// 
			this.label46.AutoSize = true;
			this.label46.BackColor = System.Drawing.Color.Transparent;
			this.label46.Location = new System.Drawing.Point(152, 152);
			this.label46.Name = "label46";
			this.label46.Size = new System.Drawing.Size(46, 17);
			this.label46.TabIndex = 47;
			this.label46.Text = "Vivienda";
			// 
			// label41
			// 
			this.label41.AutoSize = true;
			this.label41.BackColor = System.Drawing.Color.Transparent;
			this.label41.Location = new System.Drawing.Point(632, 152);
			this.label41.Name = "label41";
			this.label41.Size = new System.Drawing.Size(38, 17);
			this.label41.TabIndex = 393;
			this.label41.Text = "Región";
			this.label41.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.label41.Click += new System.EventHandler(this.label41_Click);
			// 
			// utpAdicional
			// 
			this.utpAdicional.Controls.Add(this.utcVarios);
			this.utpAdicional.Location = new System.Drawing.Point(-10000, -10000);
			this.utpAdicional.Name = "utpAdicional";
			this.utpAdicional.Size = new System.Drawing.Size(1136, 315);
			// 
			// utcVarios
			// 
			this.utcVarios.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			appearance86.ForeColorDisabled = System.Drawing.Color.Black;
			appearance86.TextVAlign = Infragistics.Win.VAlign.Middle;
			this.utcVarios.Appearance = appearance86;
			this.utcVarios.Controls.Add(this.ultraTabSharedControlsPage2);
			this.utcVarios.Controls.Add(this.ultraTabPageControl1);
			this.utcVarios.Controls.Add(this.ultraTabPageControl2);
			this.utcVarios.Controls.Add(this.ultraTabPageControl3);
			this.utcVarios.Controls.Add(this.ultraTabPageControl4);
			this.utcVarios.Controls.Add(this.ultraTabPageControl5);
			this.utcVarios.Controls.Add(this.ultraTabPageControl6);
			this.utcVarios.Controls.Add(this.ultraTabPageControl7);
			this.utcVarios.Controls.Add(this.utpFaltas);
			this.utcVarios.Controls.Add(this.utpOtros);
			this.utcVarios.Controls.Add(this.ultraTabPageControl8);
			this.utcVarios.Controls.Add(this.utpCargas);
			this.utcVarios.Location = new System.Drawing.Point(7, 8);
			this.utcVarios.Name = "utcVarios";
			this.utcVarios.SharedControlsPage = this.ultraTabSharedControlsPage2;
			this.utcVarios.Size = new System.Drawing.Size(1122, 304);
			this.utcVarios.Style = Infragistics.Win.UltraWinTabControl.UltraTabControlStyle.Excel;
			this.utcVarios.TabIndex = 185;
			ultraTab1.TabPage = this.ultraTabPageControl3;
			ultraTab1.Text = "Referencias Personales y Familiares";
			ultraTab2.TabPage = this.ultraTabPageControl2;
			ultraTab2.Text = "Estudios Realizados";
			ultraTab3.TabPage = this.ultraTabPageControl1;
			ultraTab3.Text = "Experiencia Laboral";
			ultraTab4.TabPage = this.ultraTabPageControl5;
			ultraTab4.Text = "Entrada - Salida";
			ultraTab5.TabPage = this.ultraTabPageControl6;
			ultraTab5.Text = "Equipos Entregados";
			ultraTab6.TabPage = this.ultraTabPageControl7;
			ultraTab6.Text = "Uniformes";
			ultraTab7.TabPage = this.ultraTabPageControl4;
			ultraTab7.Text = "Permisos - Vacaciones";
			ultraTab8.TabPage = this.utpFaltas;
			ultraTab8.Text = "Faltas";
			ultraTab9.TabPage = this.utpOtros;
			ultraTab9.Text = "Otros";
			ultraTab10.TabPage = this.utpCargas;
			ultraTab10.Text = "Cargas Familiares";
			ultraTab11.TabPage = this.ultraTabPageControl8;
			ultraTab11.Text = "Notas Periodo de Prueba";
			this.utcVarios.Tabs.AddRange(new Infragistics.Win.UltraWinTabControl.UltraTab[] {
																																												ultraTab1,
																																												ultraTab2,
																																												ultraTab3,
																																												ultraTab4,
																																												ultraTab5,
																																												ultraTab6,
																																												ultraTab7,
																																												ultraTab8,
																																												ultraTab9,
																																												ultraTab10,
																																												ultraTab11});
			this.utcVarios.ViewStyle = Infragistics.Win.UltraWinTabControl.ViewStyle.Office2003;
			this.utcVarios.SelectedTabChanged += new Infragistics.Win.UltraWinTabControl.SelectedTabChangedEventHandler(this.utcVarios_SelectedTabChanged_1);
			// 
			// ultraTabSharedControlsPage2
			// 
			this.ultraTabSharedControlsPage2.Location = new System.Drawing.Point(-10000, -10000);
			this.ultraTabSharedControlsPage2.Name = "ultraTabSharedControlsPage2";
			this.ultraTabSharedControlsPage2.Size = new System.Drawing.Size(1120, 283);
			// 
			// utpImagenes
			// 
			this.utpImagenes.Controls.Add(this.label42);
			this.utpImagenes.Controls.Add(this.pbFoto);
			this.utpImagenes.Location = new System.Drawing.Point(-10000, -10000);
			this.utpImagenes.Name = "utpImagenes";
			this.utpImagenes.Size = new System.Drawing.Size(1136, 315);
			// 
			// label42
			// 
			this.label42.AutoSize = true;
			this.label42.BackColor = System.Drawing.Color.Transparent;
			this.label42.Location = new System.Drawing.Point(24, 48);
			this.label42.Name = "label42";
			this.label42.Size = new System.Drawing.Size(26, 17);
			this.label42.TabIndex = 132;
			this.label42.Text = "Foto";
			this.label42.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// pbFoto
			// 
			this.pbFoto.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
			this.pbFoto.Location = new System.Drawing.Point(24, 72);
			this.pbFoto.Name = "pbFoto";
			this.pbFoto.Size = new System.Drawing.Size(160, 190);
			this.pbFoto.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage;
			this.pbFoto.TabIndex = 2;
			this.pbFoto.TabStop = false;
			// 
			// label8
			// 
			this.label8.AutoSize = true;
			this.label8.BackColor = System.Drawing.Color.Transparent;
			this.label8.Location = new System.Drawing.Point(16, 10);
			this.label8.Name = "label8";
			this.label8.Size = new System.Drawing.Size(71, 17);
			this.label8.TabIndex = 130;
			this.label8.Text = "Identificación";
			this.label8.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// txtCodigo
			// 
			this.txtCodigo.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtCodigo.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtCodigo.Enabled = false;
			this.txtCodigo.Location = new System.Drawing.Point(432, 8);
			this.txtCodigo.Name = "txtCodigo";
			this.txtCodigo.Size = new System.Drawing.Size(144, 22);
			this.txtCodigo.TabIndex = 128;
			// 
			// label7
			// 
			this.label7.AutoSize = true;
			this.label7.BackColor = System.Drawing.Color.Transparent;
			this.label7.Location = new System.Drawing.Point(376, 10);
			this.label7.Name = "label7";
			this.label7.Size = new System.Drawing.Size(38, 17);
			this.label7.TabIndex = 127;
			this.label7.Text = "Codigo";
			this.label7.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// txtNEIdentificacion
			// 
			this.txtNEIdentificacion.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtNEIdentificacion.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtNEIdentificacion.Enabled = false;
			this.txtNEIdentificacion.Location = new System.Drawing.Point(240, 8);
			this.txtNEIdentificacion.Name = "txtNEIdentificacion";
			this.txtNEIdentificacion.Size = new System.Drawing.Size(120, 22);
			this.txtNEIdentificacion.TabIndex = 126;
			this.txtNEIdentificacion.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtNEIdentificacion_KeyDown);
			this.txtNEIdentificacion.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.txtNEIdentificacion_KeyPress);
			this.txtNEIdentificacion.Validated += new System.EventHandler(this.txtNEIdentificacion_Validated);
			this.txtNEIdentificacion.ValueChanged += new System.EventHandler(this.txtNEIdentificacion_ValueChanged);
			// 
			// txtNEApellidoPaterno
			// 
			this.txtNEApellidoPaterno.BorderStyle = Infragistics.Win.UIElementBorderStyle.Solid;
			this.txtNEApellidoPaterno.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtNEApellidoPaterno.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtNEApellidoPaterno.Enabled = false;
			this.txtNEApellidoPaterno.Location = new System.Drawing.Point(112, 32);
			this.txtNEApellidoPaterno.Name = "txtNEApellidoPaterno";
			this.txtNEApellidoPaterno.Size = new System.Drawing.Size(145, 20);
			this.txtNEApellidoPaterno.TabIndex = 125;
			this.txtNEApellidoPaterno.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtNEApellidoPaterno_KeyDown);
			this.txtNEApellidoPaterno.Validated += new System.EventHandler(this.txtNEApellidoPaterno_Validated);
			this.txtNEApellidoPaterno.ValueChanged += new System.EventHandler(this.txtNEApellidoPaterno_ValueChanged);
			// 
			// label1
			// 
			this.label1.AutoSize = true;
			this.label1.BackColor = System.Drawing.Color.Transparent;
			this.label1.Location = new System.Drawing.Point(16, 33);
			this.label1.Name = "label1";
			this.label1.Size = new System.Drawing.Size(85, 17);
			this.label1.TabIndex = 124;
			this.label1.Text = "Apellido Paterno";
			this.label1.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// utcNomina
			// 
			this.utcNomina.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			appearance87.ForeColorDisabled = System.Drawing.Color.Black;
			this.utcNomina.Appearance = appearance87;
			this.utcNomina.Controls.Add(this.ultraTabSharedControlsPage1);
			this.utcNomina.Controls.Add(this.utcDatosGenerales);
			this.utcNomina.Controls.Add(this.utpEstudios);
			this.utcNomina.Controls.Add(this.utpImagenes);
			this.utcNomina.Controls.Add(this.utpAdicional);
			this.utcNomina.Location = new System.Drawing.Point(8, 56);
			this.utcNomina.Name = "utcNomina";
			this.utcNomina.SharedControlsPage = this.ultraTabSharedControlsPage1;
			this.utcNomina.Size = new System.Drawing.Size(1138, 336);
			this.utcNomina.Style = Infragistics.Win.UltraWinTabControl.UltraTabControlStyle.Excel;
			this.utcNomina.TabIndex = 146;
			ultraTab12.TabPage = this.utcDatosGenerales;
			ultraTab12.Text = "Datos Generales";
			ultraTab13.TabPage = this.utpEstudios;
			ultraTab13.Text = "Contratación";
			ultraTab14.TabPage = this.utpAdicional;
			ultraTab14.Text = "Información Adicional";
			ultraTab15.TabPage = this.utpImagenes;
			ultraTab15.Text = "Imagenes";
			this.utcNomina.Tabs.AddRange(new Infragistics.Win.UltraWinTabControl.UltraTab[] {
																																												ultraTab12,
																																												ultraTab13,
																																												ultraTab14,
																																												ultraTab15});
			this.utcNomina.ViewStyle = Infragistics.Win.UltraWinTabControl.ViewStyle.Office2003;
			this.utcNomina.SelectedTabChanged += new Infragistics.Win.UltraWinTabControl.SelectedTabChangedEventHandler(this.utcNomina_SelectedTabChanged);
			// 
			// ultraTabSharedControlsPage1
			// 
			this.ultraTabSharedControlsPage1.Location = new System.Drawing.Point(-10000, -10000);
			this.ultraTabSharedControlsPage1.Name = "ultraTabSharedControlsPage1";
			this.ultraTabSharedControlsPage1.Size = new System.Drawing.Size(1136, 315);
			// 
			// btnNuevo
			// 
			this.btnNuevo.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btnNuevo.CausesValidation = false;
			this.btnNuevo.Image = ((System.Drawing.Image)(resources.GetObject("btnNuevo.Image")));
			this.btnNuevo.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnNuevo.Location = new System.Drawing.Point(8, 402);
			this.btnNuevo.Name = "btnNuevo";
			this.btnNuevo.Size = new System.Drawing.Size(72, 23);
			this.btnNuevo.TabIndex = 340;
			this.btnNuevo.Text = "Nuevo";
			this.btnNuevo.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnNuevo.Click += new System.EventHandler(this.btnNuevo_Click);
			// 
			// btnConsultar
			// 
			this.btnConsultar.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btnConsultar.CausesValidation = false;
			this.btnConsultar.Image = ((System.Drawing.Image)(resources.GetObject("btnConsultar.Image")));
			this.btnConsultar.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnConsultar.Location = new System.Drawing.Point(82, 402);
			this.btnConsultar.Name = "btnConsultar";
			this.btnConsultar.Size = new System.Drawing.Size(78, 23);
			this.btnConsultar.TabIndex = 341;
			this.btnConsultar.Text = "Consultar";
			this.btnConsultar.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnConsultar.Click += new System.EventHandler(this.btnConsultar_Click);
			// 
			// btnEditar
			// 
			this.btnEditar.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btnEditar.CausesValidation = false;
			this.btnEditar.Enabled = false;
			this.btnEditar.Image = ((System.Drawing.Image)(resources.GetObject("btnEditar.Image")));
			this.btnEditar.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnEditar.Location = new System.Drawing.Point(162, 402);
			this.btnEditar.Name = "btnEditar";
			this.btnEditar.Size = new System.Drawing.Size(76, 23);
			this.btnEditar.TabIndex = 342;
			this.btnEditar.Text = "Editar";
			this.btnEditar.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnEditar.Click += new System.EventHandler(this.btnEditar_Click);
			// 
			// btnSalir
			// 
			this.btnSalir.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btnSalir.CausesValidation = false;
			this.btnSalir.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnSalir.ImageIndex = 62;
			this.btnSalir.Location = new System.Drawing.Point(471, 402);
			this.btnSalir.Name = "btnSalir";
			this.btnSalir.TabIndex = 343;
			this.btnSalir.Text = "Salir";
			this.btnSalir.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnSalir.Click += new System.EventHandler(this.btnSalir_Click);
			// 
			// btnImprimir
			// 
			this.btnImprimir.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btnImprimir.Enabled = false;
			this.btnImprimir.Image = ((System.Drawing.Image)(resources.GetObject("btnImprimir.Image")));
			this.btnImprimir.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnImprimir.Location = new System.Drawing.Point(318, 402);
			this.btnImprimir.Name = "btnImprimir";
			this.btnImprimir.Size = new System.Drawing.Size(77, 23);
			this.btnImprimir.TabIndex = 337;
			this.btnImprimir.Text = "&Imprimir";
			this.btnImprimir.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnImprimir.Click += new System.EventHandler(this.btnImprimir_Click);
			// 
			// btnGuardar
			// 
			this.btnGuardar.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btnGuardar.Enabled = false;
			this.btnGuardar.Image = ((System.Drawing.Image)(resources.GetObject("btnGuardar.Image")));
			this.btnGuardar.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnGuardar.Location = new System.Drawing.Point(240, 402);
			this.btnGuardar.Name = "btnGuardar";
			this.btnGuardar.Size = new System.Drawing.Size(76, 23);
			this.btnGuardar.TabIndex = 336;
			this.btnGuardar.Text = "&Guardar";
			this.btnGuardar.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnGuardar.Click += new System.EventHandler(this.btnGuardar_Click);
			// 
			// btnCancelar
			// 
			this.btnCancelar.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btnCancelar.CausesValidation = false;
			this.btnCancelar.Enabled = false;
			this.btnCancelar.Image = ((System.Drawing.Image)(resources.GetObject("btnCancelar.Image")));
			this.btnCancelar.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnCancelar.Location = new System.Drawing.Point(397, 402);
			this.btnCancelar.Name = "btnCancelar";
			this.btnCancelar.Size = new System.Drawing.Size(72, 23);
			this.btnCancelar.TabIndex = 339;
			this.btnCancelar.Text = "&Cancelar";
			this.btnCancelar.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnCancelar.Click += new System.EventHandler(this.btnCancelar_Click);
			// 
			// cdsSeteoF
			// 
			this.cdsSeteoF.BindingContextCtrl = this;
			this.cdsSeteoF.DataLibrary = "LibFacturacion";
			this.cdsSeteoF.DataLibraryUrl = "";
			this.cdsSeteoF.DataSetDef = "dsSeteoF";
			this.cdsSeteoF.Locale = new System.Globalization.CultureInfo("es-EC");
			this.cdsSeteoF.Name = "cdsSeteoF";
			this.cdsSeteoF.SchemaClassName = "LibFacturacion.DataClass";
			this.cdsSeteoF.SchemaDef = null;
			this.cdsSeteoF.BeforeFill += new C1.Data.FillEventHandler(this.cdsSeteoF_BeforeFill);
			// 
			// cmbNEIdentificacion
			// 
			this.cmbNEIdentificacion.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbNEIdentificacion.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbNEIdentificacion.DataSource = this.ultraDataSource1;
			ultraGridBand37.AddButtonCaption = "RolSucursal";
			ultraGridColumn146.Header.VisiblePosition = 0;
			ultraGridColumn146.Hidden = true;
			ultraGridColumn147.Header.VisiblePosition = 1;
			ultraGridColumn147.Width = 120;
			ultraGridBand37.Columns.AddRange(new object[] {
																											ultraGridColumn146,
																											ultraGridColumn147});
			this.cmbNEIdentificacion.DisplayLayout.BandsSerializer.Add(ultraGridBand37);
			this.cmbNEIdentificacion.DisplayMember = "Nombre";
			this.cmbNEIdentificacion.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbNEIdentificacion.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbNEIdentificacion.Enabled = false;
			this.cmbNEIdentificacion.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbNEIdentificacion.Location = new System.Drawing.Point(112, 8);
			this.cmbNEIdentificacion.Name = "cmbNEIdentificacion";
			this.cmbNEIdentificacion.Size = new System.Drawing.Size(120, 22);
			this.cmbNEIdentificacion.TabIndex = 185;
			this.cmbNEIdentificacion.ValueMember = "idTipoDoc";
			this.cmbNEIdentificacion.KeyDown += new System.Windows.Forms.KeyEventHandler(this.cmbNEIdentificacion_KeyDown);
			this.cmbNEIdentificacion.ValueChanged += new System.EventHandler(this.cmbNEIdentificacion_ValueChanged);
			// 
			// ultraDataSource1
			// 
			ultraDataColumn140.DataType = typeof(int);
			this.ultraDataSource1.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn140,
																																 ultraDataColumn141});
			// 
			// txtEstado
			// 
			this.txtEstado.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtEstado.Enabled = false;
			this.txtEstado.Location = new System.Drawing.Point(864, 400);
			this.txtEstado.MinValue = 0;
			this.txtEstado.Name = "txtEstado";
			this.txtEstado.PromptChar = ' ';
			this.txtEstado.Size = new System.Drawing.Size(80, 22);
			this.txtEstado.SpinButtonDisplayStyle = Infragistics.Win.ButtonDisplayStyle.Always;
			this.txtEstado.TabIndex = 187;
			this.txtEstado.Visible = false;
			// 
			// txtIdNomina
			// 
			this.txtIdNomina.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtIdNomina.Enabled = false;
			this.txtIdNomina.Location = new System.Drawing.Point(768, 400);
			this.txtIdNomina.MinValue = 0;
			this.txtIdNomina.Name = "txtIdNomina";
			this.txtIdNomina.PromptChar = ' ';
			this.txtIdNomina.Size = new System.Drawing.Size(80, 22);
			this.txtIdNomina.SpinButtonDisplayStyle = Infragistics.Win.ButtonDisplayStyle.Always;
			this.txtIdNomina.TabIndex = 344;
			this.txtIdNomina.Visible = false;
			// 
			// txtNEPrimerNombre
			// 
			this.txtNEPrimerNombre.BorderStyle = Infragistics.Win.UIElementBorderStyle.Solid;
			this.txtNEPrimerNombre.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtNEPrimerNombre.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtNEPrimerNombre.Enabled = false;
			this.txtNEPrimerNombre.Location = new System.Drawing.Point(632, 32);
			this.txtNEPrimerNombre.Name = "txtNEPrimerNombre";
			this.txtNEPrimerNombre.Size = new System.Drawing.Size(145, 20);
			this.txtNEPrimerNombre.TabIndex = 345;
			this.txtNEPrimerNombre.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtNEPrimerNombre_KeyDown);
			this.txtNEPrimerNombre.Validated += new System.EventHandler(this.txtNEPrimerNombre_Validated);
			this.txtNEPrimerNombre.ValueChanged += new System.EventHandler(this.txtNEPrimerNombre_ValueChanged);
			// 
			// label31
			// 
			this.label31.AutoSize = true;
			this.label31.BackColor = System.Drawing.Color.Transparent;
			this.label31.Location = new System.Drawing.Point(536, 33);
			this.label31.Name = "label31";
			this.label31.Size = new System.Drawing.Size(80, 17);
			this.label31.TabIndex = 346;
			this.label31.Text = "Primer Nombre";
			this.label31.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// label36
			// 
			this.label36.AutoSize = true;
			this.label36.BackColor = System.Drawing.Color.Transparent;
			this.label36.Location = new System.Drawing.Point(272, 33);
			this.label36.Name = "label36";
			this.label36.Size = new System.Drawing.Size(88, 17);
			this.label36.TabIndex = 347;
			this.label36.Text = "Apellido Materno";
			this.label36.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// txtNEApellidoMaterno
			// 
			this.txtNEApellidoMaterno.BorderStyle = Infragistics.Win.UIElementBorderStyle.Solid;
			this.txtNEApellidoMaterno.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtNEApellidoMaterno.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtNEApellidoMaterno.Enabled = false;
			this.txtNEApellidoMaterno.Location = new System.Drawing.Point(376, 32);
			this.txtNEApellidoMaterno.Name = "txtNEApellidoMaterno";
			this.txtNEApellidoMaterno.Size = new System.Drawing.Size(145, 20);
			this.txtNEApellidoMaterno.TabIndex = 348;
			this.txtNEApellidoMaterno.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtNEApellidoMaterno_KeyDown);
			// 
			// label37
			// 
			this.label37.AutoSize = true;
			this.label37.BackColor = System.Drawing.Color.Transparent;
			this.label37.Location = new System.Drawing.Point(792, 33);
			this.label37.Name = "label37";
			this.label37.Size = new System.Drawing.Size(91, 17);
			this.label37.TabIndex = 350;
			this.label37.Text = "Segundo Nombre";
			this.label37.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// txtNESegundoNombre
			// 
			this.txtNESegundoNombre.BorderStyle = Infragistics.Win.UIElementBorderStyle.Solid;
			this.txtNESegundoNombre.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtNESegundoNombre.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtNESegundoNombre.Enabled = false;
			this.txtNESegundoNombre.Location = new System.Drawing.Point(896, 32);
			this.txtNESegundoNombre.Name = "txtNESegundoNombre";
			this.txtNESegundoNombre.Size = new System.Drawing.Size(145, 20);
			this.txtNESegundoNombre.TabIndex = 349;
			this.txtNESegundoNombre.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtNESegundoNombre_KeyDown);
			// 
			// cmbNENacionalidad
			// 
			this.cmbNENacionalidad.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbNENacionalidad.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbNENacionalidad.DataSource = this.ultraDataSource26;
			ultraGridBand38.AddButtonCaption = "RolSucursal";
			ultraGridColumn148.Header.VisiblePosition = 0;
			ultraGridColumn148.Hidden = true;
			ultraGridColumn149.Header.VisiblePosition = 1;
			ultraGridColumn149.Width = 300;
			ultraGridBand38.Columns.AddRange(new object[] {
																											ultraGridColumn148,
																											ultraGridColumn149});
			this.cmbNENacionalidad.DisplayLayout.BandsSerializer.Add(ultraGridBand38);
			this.cmbNENacionalidad.DisplayMember = "Nombre";
			this.cmbNENacionalidad.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbNENacionalidad.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbNENacionalidad.Enabled = false;
			this.cmbNENacionalidad.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbNENacionalidad.Location = new System.Drawing.Point(680, 8);
			this.cmbNENacionalidad.Name = "cmbNENacionalidad";
			this.cmbNENacionalidad.Size = new System.Drawing.Size(216, 22);
			this.cmbNENacionalidad.TabIndex = 352;
			this.cmbNENacionalidad.ValueMember = "idNacionalidad";
			this.cmbNENacionalidad.ValueChanged += new System.EventHandler(this.cmbNENacionalidad_ValueChanged);
			this.cmbNENacionalidad.InitializeLayout += new Infragistics.Win.UltraWinGrid.InitializeLayoutEventHandler(this.cmbNENacionalidad_InitializeLayout);
			// 
			// ultraDataSource26
			// 
			ultraDataColumn142.DataType = typeof(int);
			this.ultraDataSource26.Band.Columns.AddRange(new object[] {
																																	ultraDataColumn142,
																																	ultraDataColumn143});
			// 
			// label54
			// 
			this.label54.AutoSize = true;
			this.label54.BackColor = System.Drawing.Color.Transparent;
			this.label54.Location = new System.Drawing.Point(592, 10);
			this.label54.Name = "label54";
			this.label54.Size = new System.Drawing.Size(68, 17);
			this.label54.TabIndex = 351;
			this.label54.Text = "Nacionalidad";
			this.label54.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// chkNERevisado
			// 
			this.chkNERevisado.BackColor = System.Drawing.Color.Transparent;
			this.chkNERevisado.Enabled = false;
			this.chkNERevisado.Location = new System.Drawing.Point(904, 8);
			this.chkNERevisado.Name = "chkNERevisado";
			this.chkNERevisado.Size = new System.Drawing.Size(136, 21);
			this.chkNERevisado.TabIndex = 356;
			this.chkNERevisado.Text = "Revisado";
			// 
			// txtIdPersonal
			// 
			this.txtIdPersonal.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtIdPersonal.Enabled = false;
			this.txtIdPersonal.Location = new System.Drawing.Point(952, 400);
			this.txtIdPersonal.MinValue = 0;
			this.txtIdPersonal.Name = "txtIdPersonal";
			this.txtIdPersonal.PromptChar = ' ';
			this.txtIdPersonal.Size = new System.Drawing.Size(80, 22);
			this.txtIdPersonal.SpinButtonDisplayStyle = Infragistics.Win.ButtonDisplayStyle.Always;
			this.txtIdPersonal.TabIndex = 357;
			this.txtIdPersonal.Visible = false;
			// 
			// btnImportar
			// 
			this.btnImportar.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			this.btnImportar.CausesValidation = false;
			this.btnImportar.Image = ((System.Drawing.Image)(resources.GetObject("btnImportar.Image")));
			this.btnImportar.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnImportar.Location = new System.Drawing.Point(568, 400);
			this.btnImportar.Name = "btnImportar";
			this.btnImportar.Size = new System.Drawing.Size(208, 23);
			this.btnImportar.TabIndex = 746;
			this.btnImportar.Text = "Actualizar Estado Fondos de Reserva";
			this.btnImportar.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnImportar.Click += new System.EventHandler(this.btnImportar_Click);
			// 
			// dtFecha
			// 
			this.dtFecha.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			dateButton3.Caption = "Today";
			this.dtFecha.DateButtons.Add(dateButton3);
			this.dtFecha.Enabled = false;
			this.dtFecha.Format = "dd/MM/yyyy";
			this.dtFecha.Location = new System.Drawing.Point(1040, 400);
			this.dtFecha.Name = "dtFecha";
			this.dtFecha.NonAutoSizeHeight = 23;
			this.dtFecha.Size = new System.Drawing.Size(104, 21);
			this.dtFecha.TabIndex = 152;
			this.dtFecha.Value = ((object)(resources.GetObject("dtFecha.Value")));
			this.dtFecha.Visible = false;
			// 
			// btnAPP
			// 
			this.btnAPP.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btnAPP.CausesValidation = false;
			this.btnAPP.Enabled = false;
			this.btnAPP.Image = ((System.Drawing.Image)(resources.GetObject("btnAPP.Image")));
			this.btnAPP.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnAPP.Location = new System.Drawing.Point(800, 400);
			this.btnAPP.Name = "btnAPP";
			this.btnAPP.Size = new System.Drawing.Size(56, 23);
			this.btnAPP.TabIndex = 1079;
			this.btnAPP.Text = "&APP";
			this.btnAPP.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnAPP.Click += new System.EventHandler(this.btnAPP_Click);
			// 
			// frmNomina
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 14);
			this.ClientSize = new System.Drawing.Size(1154, 432);
			this.Controls.Add(this.btnAPP);
			this.Controls.Add(this.btnImportar);
			this.Controls.Add(this.txtIdPersonal);
			this.Controls.Add(this.cmbNENacionalidad);
			this.Controls.Add(this.label54);
			this.Controls.Add(this.label37);
			this.Controls.Add(this.label36);
			this.Controls.Add(this.label31);
			this.Controls.Add(this.label8);
			this.Controls.Add(this.label1);
			this.Controls.Add(this.label7);
			this.Controls.Add(this.txtNESegundoNombre);
			this.Controls.Add(this.txtNEApellidoMaterno);
			this.Controls.Add(this.txtNEPrimerNombre);
			this.Controls.Add(this.txtIdNomina);
			this.Controls.Add(this.btnNuevo);
			this.Controls.Add(this.btnConsultar);
			this.Controls.Add(this.btnEditar);
			this.Controls.Add(this.btnSalir);
			this.Controls.Add(this.btnImprimir);
			this.Controls.Add(this.btnGuardar);
			this.Controls.Add(this.btnCancelar);
			this.Controls.Add(this.utcNomina);
			this.Controls.Add(this.txtNEIdentificacion);
			this.Controls.Add(this.txtNEApellidoPaterno);
			this.Controls.Add(this.txtCodigo);
			this.Controls.Add(this.cmbNEIdentificacion);
			this.Controls.Add(this.txtEstado);
			this.Controls.Add(this.chkNERevisado);
			this.Controls.Add(this.dtFecha);
			this.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog;
			this.KeyPreview = true;
			this.MaximizeBox = false;
			this.Name = "frmNomina";
			this.ShowInTaskbar = false;
			this.Text = "Nomina";
			this.KeyDown += new System.Windows.Forms.KeyEventHandler(this.frmNomina_KeyDown);
			this.Closing += new System.ComponentModel.CancelEventHandler(this.frmNomina_Closing);
			this.Load += new System.EventHandler(this.frmNomina_Load);
			this.ultraTabPageControl3.ResumeLayout(false);
			((System.ComponentModel.ISupportInitialize)(this.cmbParentesco)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource18)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.uGridFamiliares)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource12)).EndInit();
			this.ultraTabPageControl2.ResumeLayout(false);
			((System.ComponentModel.ISupportInitialize)(this.cmbPeriodosTiempo)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource23)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbEstadoEstudios)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource20)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbNivel)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource19)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.uGridEstudios)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource13)).EndInit();
			this.ultraTabPageControl1.ResumeLayout(false);
			((System.ComponentModel.ISupportInitialize)(this.uGridExperiencia)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource14)).EndInit();
			this.ultraTabPageControl5.ResumeLayout(false);
			((System.ComponentModel.ISupportInitialize)(this.cmbPersonalAutVerifica)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource24)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbMotivosDeSalida)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource21)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.uGridEntradasSalidas)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource15)).EndInit();
			this.ultraTabPageControl6.ResumeLayout(false);
			((System.ComponentModel.ISupportInitialize)(this.uGridEquipos)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource17)).EndInit();
			this.ultraTabPageControl7.ResumeLayout(false);
			((System.ComponentModel.ISupportInitialize)(this.uGridUniformes)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource16)).EndInit();
			this.ultraTabPageControl4.ResumeLayout(false);
			((System.ComponentModel.ISupportInitialize)(this.cmbTiposDeParto)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource30)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbTipoDeSalida)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource29)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.uGridVacaciones)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource25)).EndInit();
			this.utpFaltas.ResumeLayout(false);
			((System.ComponentModel.ISupportInitialize)(this.cmbTiposDeFaltas)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.uGridFaltas)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource28)).EndInit();
			this.utpOtros.ResumeLayout(false);
			((System.ComponentModel.ISupportInitialize)(this.cmbMotivos)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource32)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.uGridOtros)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource31)).EndInit();
			this.utpCargas.ResumeLayout(false);
			((System.ComponentModel.ISupportInitialize)(this.uGridCargas)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource35)).EndInit();
			this.ultraTabPageControl8.ResumeLayout(false);
			((System.ComponentModel.ISupportInitialize)(this.uGridObservaciones)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource33)).EndInit();
			this.utcDatosGenerales.ResumeLayout(false);
			((System.ComponentModel.ISupportInitialize)(this.optNEMesesAños)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNEVive_Con)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtPredial)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNETiempoDeVivienda)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNEEMailEmpresarial)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbNEProvincia)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource2)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbProvinciaNacimiento)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.optNEVivienda)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNEEMail)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNECelular)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNETelefono)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNEDireccion)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbNEBarrio)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource6)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbNEParroquia)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource5)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbNECanton)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource3)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNECargasFamiliares)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbNEEstadoCivil)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource4)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtConadis)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtPorcentajeDiscapacidad)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.dtpNEFechaNacimiento)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNETipoDeSangre)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.optNESexo)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbCantonNacimiento)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtTipoDiscapacidad)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.dtpFechaCaducidad)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtTipoDeLicencia)).EndInit();
			this.utpEstudios.ResumeLayout(false);
			((System.ComponentModel.ISupportInitialize)(this.cmbRegion)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource34)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbRangos)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource27)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtQuincena)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNEBonos)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNEOrdinarias)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbNEBancos)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource10)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNEResponsabilidades)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNEFunciones)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNEComisiones)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNEBono)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtRetencionJudicial)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtGastoVestimenta)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtGastoEducacion)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtGastoAlimentacion)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtGastoSalud)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtGastoVivienda)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtCuenta)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.optNETipoDePago)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtTotalIngresos)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNEHorasExtras)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNEMovilizacion)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNESueldo)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbNETipoContrato)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource11)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.optNESistema)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.optNERolesProservicios)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbNECodigoSectorial)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource22)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbNECargo)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource9)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbNEDepartamento)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource8)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbNECentroCosto)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource7)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtAfiliacion)).EndInit();
			this.utpAdicional.ResumeLayout(false);
			((System.ComponentModel.ISupportInitialize)(this.utcVarios)).EndInit();
			this.utcVarios.ResumeLayout(false);
			this.utpImagenes.ResumeLayout(false);
			((System.ComponentModel.ISupportInitialize)(this.txtCodigo)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNEIdentificacion)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNEApellidoPaterno)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.utcNomina)).EndInit();
			this.utcNomina.ResumeLayout(false);
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbNEIdentificacion)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource1)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtEstado)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIdNomina)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNEPrimerNombre)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNEApellidoMaterno)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNESegundoNombre)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbNENacionalidad)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource26)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIdPersonal)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.dtFecha)).EndInit();
			this.ResumeLayout(false);

		}
		#endregion

		#region Variables 
		private Acceso miAcceso;		
		bool bNuevo = false;
		bool bEdicion = false;
		bool bEstado = false;
		int idCentroDeCostoA = 0; 
		int idDepartamentoA = 0; 
		int idCargoA = 0;
		decimal SueldoA = 0.00m;
		decimal BonoA = 0.00m; 
		decimal MovilizacionA = 0.00m;
		int idCom_Estado = 0;
		#endregion Variables 

		private void UnloadMe()
		{
			this.Close();
		}

		private void ConsultaTablas()
		{
			this.ConsultaNominaTablas(1, this.uGridFamiliares);
			this.ConsultaNominaTablas(2, this.uGridEstudios);
			this.ConsultaNominaTablas(3, this.uGridExperiencia);
			this.ConsultaNominaTablas(4, this.uGridEntradasSalidas);
			this.ConsultaNominaTablas(5, this.uGridEquipos);
			this.ConsultaNominaTablas(6, this.uGridVacaciones);
			this.ConsultaNominaTablas(7, this.uGridUniformes);
			this.ConsultaNominaTablas(8, this.uGridFaltas);
			this.ConsultaNominaTablas(9, this.uGridOtros);
			this.ConsultaNominaTablas(10, this.uGridObservaciones);
			this.ConsultaNominaTablas(11, this.uGridCargas);
		}

		private void ConsultaNominaTablas(int iNumero, Infragistics.Win.UltraWinGrid.UltraGrid uGrid)
		{
			string sSQL = string.Format("Exec ConsultaNominaTablas {0}, {1}", iNumero, (int)this.txtIdNomina.Value);
			uGrid.DataSource = Funcion.dvProcedimiento(cdsSeteoF, sSQL);
		}

		private void Consulta (int idNomina, bool bValida)
		{
			try
			{	
				if (!bValida)
				{
					this.btnNuevo.Enabled = false;
					this.btnConsultar.Enabled = false;
					this.btnEditar.Enabled = false;
					this.btnGuardar.Enabled = false;
					this.btnImprimir.Enabled = false;
					this.btnCancelar.Enabled = false;
					 this.btnAPP.Enabled = false;
				}

				int idProvinciaNac = 0;
				int idCantonNac = 0;
				int idProvinciaDom = 0;
				int idCantonDom = 0;
				int idParroquiaDom = 0;
				int idBarrioDom = 0;

				#region Consulta
				string sSQL = string.Format("Exec ConsultaNominaIndividual {0}", idNomina);
				SqlDataReader drConsulta = Funcion.rEscalarSQL(cdsSeteoF, sSQL, true);
				drConsulta.Read();
				if (drConsulta.HasRows)
				{
					#region Datos Generales
					
					this.txtIdNomina.Value = drConsulta.GetInt32(0);
					this.txtIdPersonal.Value = drConsulta.GetInt32(74);
					this.txtEstado.Value = drConsulta.GetInt32(76);
					this.lblEstado.Text = drConsulta.GetString(77);
					bEstado = drConsulta.GetBoolean(1);
					this.cmbNEIdentificacion.Value = drConsulta.GetInt32(2);
					this.txtNEIdentificacion.Text = drConsulta.GetString(3);
					this.txtNEApellidoPaterno.Text = drConsulta.GetString(4);
					this.txtNEApellidoMaterno.Text = drConsulta.GetString(5);
					this.txtNEPrimerNombre.Text = drConsulta.GetString(60);
					this.txtNESegundoNombre.Text = drConsulta.GetString(61);
					this.txtCodigo.Text = drConsulta.GetString(6);

					this.cmbNEBancos.Value = drConsulta.GetInt32(45);
					this.txtCuenta.Text = drConsulta.GetString(46);

					this.cmbNENacionalidad.Value = drConsulta.GetInt32(69);

					this.chkNERevisado.Checked = drConsulta.GetBoolean(70);
					if (miAcceso.BVistas) 
					{
						idProvinciaNac = drConsulta.GetInt32(7);
						idCantonNac = drConsulta.GetInt32(8);
						this.dtpNEFechaNacimiento.Value = drConsulta.GetDateTime(9);
						this.optNESexo.Value = drConsulta.GetInt32(10);
						this.txtNETipoDeSangre.Value = drConsulta.GetString(11);
						this.chkNEDiscapacidad.Checked = drConsulta.GetBoolean(12);
						this.txtPorcentajeDiscapacidad.Value = drConsulta.GetInt32(13);
						this.txtConadis.Text = drConsulta.GetString(14);
						this.txtTipoDiscapacidad.Text = drConsulta.GetString(15);
						this.cmbNEEstadoCivil.Value = drConsulta.GetInt32(16); 
						this.txtNECargasFamiliares.Value = drConsulta.GetInt32(17);
						idProvinciaDom = drConsulta.GetInt32(18);
						idCantonDom = drConsulta.GetInt32(19);
						idParroquiaDom = drConsulta.GetInt32(20);
						idBarrioDom = drConsulta.GetInt32(21);
						this.txtNEDireccion.Text = drConsulta.GetString(22);
						this.txtNETelefono.Text = drConsulta.GetString(23);
						this.txtNECelular.Text = drConsulta.GetString(24);
						this.txtNEEMail.Text = drConsulta.GetString(25);	
						this.txtNEEMailEmpresarial.Text = drConsulta.GetString(62);
						this.optNEVivienda.Value = drConsulta.GetInt32(26);
						this.txtNETiempoDeVivienda.Value = drConsulta.GetInt32(27);
						this.optNEMesesAños.Value = drConsulta.GetInt32(68);
						this.txtNEVive_Con.Text = drConsulta.GetString(64);					
						this.txtPredial.Text = drConsulta.GetString(63);
						//						this.cmbNENacionalidad.Value = drConsulta.GetInt32(69);
						if (drConsulta.GetValue(78) != System.DBNull.Value) this.cmbRegion.Value  = drConsulta.GetInt32(78);
						#endregion Datos Generales
						#region Contratacion
						this.cmbNECentroCosto.Value = drConsulta.GetInt32(28);
						this.cmbNEDepartamento.Value = drConsulta.GetInt32(29);
						this.cmbNECargo.Value = drConsulta.GetInt32(30);
						this.cmbNECodigoSectorial.Value = drConsulta.GetInt32(31);
						this.optNESistema.Value = drConsulta.GetInt32(32);
						this.cmbNETipoContrato.Value = drConsulta.GetInt32(33);
						this.chkNEAporteIESS.Checked = drConsulta.GetBoolean(34);
						this.txtAfiliacion.Text = drConsulta.GetString(35);
						this.chkFondoReserva.Checked = drConsulta.GetBoolean(36);
						//this.chkDecimos.Checked = drConsulta.GetBoolean(37);
						this.txtNESueldo.Value = drConsulta.GetDecimal(38);
						this.txtNEBonos.Value = drConsulta.GetDecimal(39);
						this.txtNEMovilizacion.Value = drConsulta.GetDecimal(40);
						this.txtNEOrdinarias.Value = drConsulta.GetDecimal(41);					
						this.txtNEHorasExtras.Value = drConsulta.GetDecimal(42);
						this.txtTotalIngresos.Value = drConsulta.GetDecimal(43);
						this.optNETipoDePago.Value = drConsulta.GetInt32(44);
						//					  this.cmbNEBancos.Value = drConsulta.GetInt32(45);
						//					  this.txtCuenta.Text = drConsulta.GetString(46);
						this.chkNEImpuestoRenta.Checked = drConsulta.GetBoolean(47);
						this.txtGastoVivienda.Value = drConsulta.GetDecimal(48);
						this.txtGastoSalud.Value = drConsulta.GetDecimal(49);
						this.txtGastoAlimentacion.Value = drConsulta.GetDecimal(50);
						this.txtGastoEducacion.Value = drConsulta.GetDecimal(51);
						this.txtGastoVestimenta.Value = drConsulta.GetDecimal(52);
						this.chkNERetencionJudicial.Checked = drConsulta.GetBoolean(53);
						this.txtRetencionJudicial.Value = drConsulta.GetDecimal(54);
						this.chkNETCC.Checked = drConsulta.GetBoolean(55);
						this.txtNEBono.Text = drConsulta.GetString(56);
						this.txtNEComisiones.Text = drConsulta.GetString(57);
						this.txtNEFunciones.Text = drConsulta.GetString(58);
						this.txtNEResponsabilidades.Text = drConsulta.GetString(59);
						this.chkNEComisiona.Checked = drConsulta.GetBoolean(65);
						this.chkNEVendedor.Checked = drConsulta.GetBoolean(66);
						this.chkDerechoFondosDeReserva.Checked = drConsulta.GetBoolean(67);
						//						this.chkNERevisado.Checked = drConsulta.GetBoolean(70);
						this.chkNEAplicaRol.Checked = drConsulta.GetBoolean(71);
						this.chkNEQuincena.Checked = drConsulta.GetBoolean(72);
						this.txtQuincena.Value = drConsulta.GetDecimal(73);
						this.cmbRangos.Value = drConsulta.GetInt32(75);
						this.chkDecTerAcum.Checked = drConsulta.GetBoolean(79);
						this.chkDecCuartoAcum.Checked = drConsulta.GetBoolean(80);	

						bool tieneAdministrador = !drConsulta.IsDBNull(81) && drConsulta.GetBoolean(81);
						bool tieneRevisor = !drConsulta.IsDBNull(82) && drConsulta.GetBoolean(82);
						bool tieneCamarera = !drConsulta.IsDBNull(83) && drConsulta.GetBoolean(83);

						if (tieneAdministrador)
							this.optNERolesProservicios.Value = 1;
						else if (tieneRevisor)
							this.optNERolesProservicios.Value = 2;
						else if (tieneCamarera)
							this.optNERolesProservicios.Value = 3;
						else
							this.optNERolesProservicios.Value = 0; // ninguno seleccionado


					}			
					#endregion Contratacion		
				}
				drConsulta.Close();
				#endregion Consulta

				FuncionesProcedimientos.NominaFoto(@"\\192.168.1.3\Fotosnomina\", this.txtNEIdentificacion.Text.ToString(), this.pbFoto);
      
				this.cmbProvinciaNacimiento.Value = idProvinciaNac;
				this.cmbCantonNacimiento.DataSource = Funcion.dvProcedimiento (cdsSeteoF, string.Format("Exec Listas 1, {0}", Convert.ToInt32(this.cmbProvinciaNacimiento.Value)));
				this.cmbCantonNacimiento.Value = idCantonNac;
				
				this.cmbNEProvincia.Value = idProvinciaDom;
				this.cmbNECanton.DataSource = Funcion.dvProcedimiento(cdsSeteoF, string.Format("Exec Listas 1, {0}", Convert.ToInt32(this.cmbNEProvincia.Value)));
				this.cmbNECanton.Value = idCantonDom;
				this.cmbNEParroquia.DataSource = Funcion.dvProcedimiento (cdsSeteoF, string.Format("Exec Listas 2, {0}", Convert.ToInt32(this.cmbNECanton.Value)));
				this.cmbNEParroquia.Value = idParroquiaDom;
				this.cmbNEBarrio.DataSource = Funcion.dvProcedimiento (cdsSeteoF, string.Format("Exec Listas 68, {0}", Convert.ToInt32(this.cmbNEParroquia.Value)));
				this.cmbNEBarrio.Value = idBarrioDom;

				//				if (bEstado) this.lblEstado.Text = "ACTIVO";
				//				if (!bEstado) this.lblEstado.Text = "INACTIVO";
				if (miAcceso.BVistas) 
				{
					this.ConsultaTablas();
				}

				if (!bValida)
				{
					if (miAcceso.Nuevo) this.btnNuevo.Enabled = true;
					if (miAcceso.Buscar) this.btnConsultar.Enabled = true;
					if (miAcceso.BReliquidar) this.btnAPP.Enabled = true;
					if (miAcceso.Editar) this.btnEditar.Enabled = true;
					this.btnGuardar.Enabled = false;
					if (miAcceso.Imprimir) this.btnImprimir.Enabled = true;
					this.btnCancelar.Enabled = true;
				}
			}
			catch(Exception Exc)
			{
				MessageBox.Show(string.Format("Error al Consultar: {0}", Exc), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Error);
			}
		}

		private string[] ApellidosNombres(string sNombre)
		{
			string[] sCadena = new string[5];

			char[] sDelimitador = { ' ', '.' };

			string[] sComodin = { "DE", "DEL", "LA", "LAS", "LOS", "LO", "EL" };

			string[] sPalabras = sNombre.Split(sDelimitador);

			int j = 0;
			bool bComodin = false;

			foreach (string sDivision in sPalabras)
			{
				if (!bComodin)
				{
					j++;
					if (j < 5) sCadena[j] = sDivision;                    
				}
				else
				{
					if (j < 6) sCadena[j] = sCadena[j] + " " + sDivision;
				}

				//var comodines = from a in sComodin where a.Contains(sDivision) select a;

				if (Array.IndexOf(sComodin, sDivision, 0, 7) > 0) bComodin = true; else bComodin = false;
				//if (comodines.Count() == 0) bComodin = false; else bComodin = true;
			}

			return sCadena;
		}

		private void EstadoGrids(bool bEstado)
		{
			FuncionesProcedimientos.EstadoGrids(bEstado, this.uGridFamiliares);
			FuncionesProcedimientos.EstadoGrids(bEstado, this.uGridEstudios);
			FuncionesProcedimientos.EstadoGrids(bEstado, this.uGridExperiencia);
			FuncionesProcedimientos.EstadoGrids(bEstado, this.uGridEntradasSalidas);
			FuncionesProcedimientos.EstadoGrids(bEstado, this.uGridEquipos);
			FuncionesProcedimientos.EstadoGrids(bEstado, this.uGridVacaciones);
			FuncionesProcedimientos.EstadoGrids(bEstado, this.uGridUniformes);
			FuncionesProcedimientos.EstadoGrids(bEstado, this.uGridFaltas);
			FuncionesProcedimientos.EstadoGrids(bEstado, this.uGridOtros);
			FuncionesProcedimientos.EstadoGrids(bEstado, this.uGridCargas);
		}

		private void TotalIngresos()
		{
			this.txtTotalIngresos.Value =
				Convert.ToDecimal(this.txtNESueldo.Value) +
				Convert.ToDecimal(this.txtNEBonos.Value) +
				Convert.ToDecimal(this.txtNEMovilizacion.Value);
		}
		
		private void frmNomina_Load(object sender, System.EventArgs e)
		{			
			if (!Funcion.Permiso("79", cdsSeteoF))
			{
				MessageBox.Show("No puede ingresar a Nomina", "Informacion");
				BeginInvoke(new MethodInvoker(UnloadMe));
				return;
			}

			string stTipo = "0501";							
			miAcceso = new Acceso(cdsSeteoF, stTipo);
			
			if (miAcceso.Nuevo) this.btnNuevo.Enabled = true;
			if (miAcceso.Buscar) this.btnConsultar.Enabled = true;

			if (MenuLatinium.bComisionQuincena)
			{
				this.chkNEQuincena.Visible = true;
				this.txtQuincena.Visible = true;
			}

			FuncionesProcedimientos.RetornaFechaServidor(this.dtFecha, cdsSeteoF, true);

			#region Listas
			this.cmbNEIdentificacion.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "Exec Listas 44, 0");
			this.cmbNENacionalidad.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "Exec Listas 11, 0");
			this.cmbProvinciaNacimiento.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "Exec Listas 0, 0");
			this.cmbNEEstadoCivil.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "Exec Listas 12, 0");
			this.cmbNEProvincia.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "Exec Listas 0, 0");
			this.cmbNECentroCosto.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "Exec Listas 45, 0");
			this.cmbNETipoContrato.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "Exec Listas 46, 0");
			this.cmbNEBancos.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "Exec Listas 47, 0");
			this.cmbNEDepartamento.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "Exec Listas 48, 0");
			this.cmbNECargo.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "Exec Listas 49, 0");
			this.cmbParentesco.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "Exec Listas 50, 0");
			this.cmbNivel.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "Exec Listas 51, 0");
			//this.cmbTitulo.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "Exec Listas 52, 0");
			this.cmbMotivosDeSalida.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "Exec Listas 53, 0");
			this.cmbNECodigoSectorial.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "Exec Listas 55, 0");
			this.cmbEstadoEstudios.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "Exec Listas 56, 0");
			this.cmbPeriodosTiempo.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "Exec Listas 57, 0");
			this.cmbPersonalAutVerifica.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "Exec Listas 58, 0");
			this.cmbRangos.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "Exec Listas 60, 0");
			this.cmbTipoDeSalida.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "Exec Listas 62, 0");
			this.cmbTiposDeParto.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "Exec Listas 63, 0");
			this.cmbTiposDeFaltas.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "Exec Listas 62, 1");
			this.cmbMotivos.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "Exec Listas 64, 0");
			this.cmbRegion.DataSource = Funcion.dvProcedimiento(cdsSeteoF,"select idRegion, Region from Region ");
			#endregion Listas

			this.dtpNEFechaNacimiento.CalendarInfo.MaxDate = DateTime.Today.AddYears(-18);
			this.dtpFechaCaducidad.CalendarInfo.MinDate = DateTime.Today;

			EstadoGrids(false);
		}

		private void btnSalir_Click(object sender, System.EventArgs e)
		{
			this.Close();
		}

		private void btnCancelar_Click(object sender, System.EventArgs e)
		{
			bNuevo = false;
			bEdicion = false;

			FuncionesProcedimientos.EstadoControles(this, false, 4);
			FuncionesProcedimientos.EstadoControlesTabControl(this.utcNomina, false, 4);	
			FuncionesProcedimientos.RetornaFechaServidor(this.dtFecha, cdsSeteoF, true);
      	this.cmbRegion.Enabled = false;
				this.cmbRegion.Value= System.DBNull.Value;

			ConsultaTablas();

			if (miAcceso.Nuevo) this.btnNuevo.Enabled = true;
			if (miAcceso.Buscar) this.btnConsultar.Enabled = true;
			this.btnEditar.Enabled = false;
			this.btnGuardar.Enabled = false;
			this.btnImprimir.Enabled = false;
			this.btnCancelar.Enabled = false;
			this.btnAPP.Enabled = false;

			this.utcVarios.Tabs[6].Enabled = true;
			this.utcVarios.Tabs[7].Enabled = true;
			this.utcVarios.Tabs[8].Enabled = true;

			this.pbFoto.Image = null;

			this.lblEstado.Text = "";		
			bEstado = false;
			idCom_Estado = 0;

			EstadoGrids(false);
		}

		private void btnNuevo_Click(object sender, System.EventArgs e)
		{
			this.btnCancelar_Click(sender, e);

			FuncionesProcedimientos.EstadoControles(this, true, 1);
			FuncionesProcedimientos.EstadoControlesTabControl(this.utcNomina, true, 1);
			
			EstadoGrids(true);
			this.cmbRegion.Enabled = true;

			this.btnNuevo.Enabled = false;
			this.btnConsultar.Enabled = false;
			this.btnAPP.Enabled = false;
			this.btnEditar.Enabled = false;
			this.btnGuardar.Enabled = true;
			this.btnImprimir.Enabled = false;
			this.btnCancelar.Enabled = true;

			bNuevo = true;
			bEdicion = true;

			bEstado = true;
			this.txtEstado.Value = 3;
			this.cmbNEIdentificacion.Value = 1;
			this.cmbNENacionalidad.Value = 54;
			this.cmbNETipoContrato.Value = 1;
			this.chkNEAporteIESS.Checked = true;
			this.chkDecTerAcum.Checked = false;
			this.chkDecCuartoAcum.Checked = false;
			//this.chkFondoReserva.Checked = true;
			this.optNETipoDePago.Value = 1;
			this.optNESistema.Value = 1;
			this.chkNEAplicaRol.Checked = true;	
			this.chkNEComisiona.Checked = true;	

			this.chkNEAplicaRol.Enabled = true;	
			this.chkNEComisiona.Enabled = false;	
			
			this.utcVarios.Tabs[6].Enabled = false;
			this.utcVarios.Tabs[7].Enabled = false;
			this.utcVarios.Tabs[8].Enabled = false;

			this.txtNEIdentificacion.Focus();
			FuncionesProcedimientos.RetornaFechaServidor(this.dtFecha, cdsSeteoF, true);
		}

		private void cdsSeteoF_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			cdsSeteoF.Schema.Connections[0].ConnectionString = miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);	
		}

		private void txtNEIdentificacion_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
			if ((int)this.cmbNEIdentificacion.Value == 1)
				Funcion.CamposNumericos(sender, e);
		}

		private void txtCuenta_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
			Funcion.CamposNumericos(sender, e);
		}

		private void txtNETelefono_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
			Funcion.CamposNumericos(sender, e);
		}

		private void txtNECelular_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
			Funcion.CamposNumericos(sender, e);
		}

		private void btnEditar_Click(object sender, System.EventArgs e)
		{
			FuncionesProcedimientos.RetornaFechaServidor(this.dtFecha, cdsSeteoF, true);

			if (Funcion.bEscalarSQL(cdsSeteoF, string.Format("Exec NominaValidaRolPagado '{0}'", Convert.ToDateTime(this.dtFecha.Value).ToString("yyyyMMdd")), true))
			{
				MessageBox.Show("El Rol del Periodo se encuentra generado. No se pueden realizar cambios mientras la nómina está generada.", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				//return;
			}

			FuncionesProcedimientos.EstadoControles(this, true, 2);
			FuncionesProcedimientos.EstadoControlesTabControl(this.utcNomina, true, 2);
			
			bNuevo = false;
			bEdicion = true;

			#region Habilita Controles Condicionales
			if (this.chkNEDiscapacidad.Checked)
			{
				this.txtPorcentajeDiscapacidad.Enabled = true;
				this.txtConadis.Enabled = true;
				this.txtTipoDiscapacidad.Enabled = true;
			}

			if (this.chkNELicencia.Checked)
			{
				this.txtTipoDeLicencia.Enabled = true;
				this.dtpFechaCaducidad.Enabled = true;				
			}

			if ((int)this.optNEVivienda.Value == 1) this.txtPredial.Enabled = true;

//			if ((int)this.cmbNEDepartamento.ActiveRow.Cells["idModulo"].Value == 1) this.cmbRangos.Enabled = true;

			if (this.chkNEAporteIESS.Checked)
			{
				this.chkDecTerAcum.Enabled = true;
				this.chkDecCuartoAcum.Enabled = true;
				this.chkFondoReserva.Enabled = true;
				this.txtAfiliacion.Enabled = true;
				this.chkDerechoFondosDeReserva.Enabled = true;
			}

			if ((int)this.optNETipoDePago.Value != 4) 
			{
				this.txtCuenta.Enabled = true;
				this.cmbNEBancos.Enabled = true;
			}

			if (this.chkNEImpuestoRenta.Checked)
			{
				this.txtGastoVivienda.Enabled = true;
				this.txtGastoSalud.Enabled = true;
				this.txtGastoAlimentacion.Enabled = true;
				this.txtGastoEducacion.Enabled = true;
				this.txtGastoVestimenta.Enabled = true;
			}

			if (this.chkNERetencionJudicial.Checked) this.txtRetencionJudicial.Enabled = true;
			if (this.chkNEQuincena.Checked) this.txtQuincena.Enabled = true;

			if ((int)this.cmbNENacionalidad.Value == 54)
			{
				this.cmbProvinciaNacimiento.Enabled = true;
				this.cmbCantonNacimiento.Enabled = true;
			}			
			#endregion Habilita Controles Condicionales
      
			this.cmbRegion.Enabled = true;
			idCom_Estado = (int)this.txtEstado.Value;
      this.chkNERetencionJudicial.Enabled = true;
			EstadoGrids(true);
      /*editar campos especiales*/
			if (miAcceso.BExportar) 
			{
				this.txtNESueldo.Enabled = true;
				this.txtNEBonos.Enabled = true;
				this.txtNEMovilizacion.Enabled = true;
				this.txtNEOrdinarias.Enabled = true;
				this.txtNEHorasExtras.Enabled = true;
				this.txtTotalIngresos.Enabled = true;
				this.chkNEAplicaRol.Enabled = true;				
			}
			else
			{
				this.txtNESueldo.Enabled = false;
				this.txtNEBonos.Enabled = false;
				this.txtNEMovilizacion.Enabled = false;
				this.txtNEOrdinarias.Enabled = false;
				this.txtNEHorasExtras.Enabled = false;
				this.txtTotalIngresos.Enabled = false;
//				this.chkNERetencionJudicial.Enabled = false;
				this.chkNEAplicaRol.Enabled = false;				
			}
       /*editar salidas*/

/* EDITAR CARGO Y DEPARTAMENTO */
			if (miAcceso.EditarCuentaNombre) 
			{
				this.cmbNEDepartamento.Enabled = true;
				this.cmbNECargo.Enabled = true;
			}
			else
			{
				this.cmbNEDepartamento.Enabled = false;
				this.cmbNECargo.Enabled = false;
			}
			/* FIN EDITAR CARGO Y DEPARTAMENTO */

			if (miAcceso.BAPAutorizar)
			{
				FuncionesProcedimientos.EstadoGrids(true, this.uGridEntradasSalidas);
			}
			else
			{
				FuncionesProcedimientos.EstadoGrids(false, this.uGridEntradasSalidas);
			}

      int idCargo =0;
			idCargo = Funcion.iEscalarSQL(cdsSeteoF, string.Format(" Select count (rc.idCargo) From  RolCargo rc Inner Join Com_CargosDeVentas cv On cv.idCargo = rc.Com_IdCargo Where cv.AV = 1 and RC.idCargo =  {0}",
				this.cmbNECargo.Value));	
				
//			if (idCargo > 0)
//			{
//				FuncionesProcedimientos.EstadoGrids(false, this.uGridVacaciones);
//			}
//			else
//			{
//				FuncionesProcedimientos.EstadoGrids(true, this.uGridVacaciones);
//			}
      
//			//this.chkNEComisiona.Enabled = false;
			this.btnNuevo.Enabled = false;
			this.btnConsultar.Enabled = false;
			 this.btnAPP.Enabled = false;
			this.btnEditar.Enabled = false;
			this.btnGuardar.Enabled = true;
			this.btnImprimir.Enabled = false;
			this.btnCancelar.Enabled = true;

			this.txtNEIdentificacion.Focus();
		}

		private void cmbNEProvinciaNacimiento_ValueChanged(object sender, System.EventArgs e)
		{
			if (bEdicion)
			{		
				int idProvNac = 0;
				if (this.cmbProvinciaNacimiento.ActiveRow != null) idProvNac = (int)this.cmbProvinciaNacimiento.Value;

				this.cmbCantonNacimiento.DataSource = Funcion.dvProcedimiento (cdsSeteoF, string.Format("Exec Listas 1, {0}", idProvNac));
			}
		}

		private void cmbNEProvincia_ValueChanged(object sender, System.EventArgs e)
		{
			if (bEdicion)
			{
				this.cmbNECanton.DataSource = Funcion.dvProcedimiento(cdsSeteoF, string.Format("Exec Listas 1, {0}", Convert.ToInt32(this.cmbNEProvincia.Value)));
			}
		}

		private void cmbNECanton_ValueChanged(object sender, System.EventArgs e)
		{
			if (bEdicion)
			{
				this.cmbNEParroquia.DataSource = Funcion.dvProcedimiento (cdsSeteoF, string.Format("Exec Listas 2, {0}", Convert.ToInt32(this.cmbNECanton.Value)));
			}
		}

		private void cmbNEParroquia_ValueChanged(object sender, System.EventArgs e)
		{
			if (bEdicion)
			{
				this.cmbNEBarrio.DataSource = Funcion.dvProcedimiento (cdsSeteoF, string.Format("Exec Listas 68, {0}", Convert.ToInt32(this.cmbNEParroquia.Value)));
			}
		}

		private void chkNEDiscapacidad_CheckedChanged(object sender, System.EventArgs e)
		{
			if (bEdicion)
			{
				this.txtPorcentajeDiscapacidad.Enabled = this.chkNEDiscapacidad.Checked;
				this.txtConadis.Enabled = this.chkNEDiscapacidad.Checked;
				this.txtTipoDiscapacidad.Enabled = this.chkNEDiscapacidad.Checked;

				if (!this.chkNEDiscapacidad.Checked)
				{
					this.txtPorcentajeDiscapacidad.Value = 0;
					this.txtConadis.Text = "";
					this.txtTipoDiscapacidad.Text = "";
				}				
			}
		}

		private void chkNEAporteIESS_CheckedChanged(object sender, System.EventArgs e)
		{
			if (bEdicion)
			{
				this.chkDecTerAcum.Enabled = this.chkNEAporteIESS.Checked;
				this.chkDecCuartoAcum.Enabled = this.chkNEAporteIESS.Checked;
				this.chkFondoReserva.Enabled = this.chkNEAporteIESS.Checked;
				this.chkDerechoFondosDeReserva.Enabled = this.chkNEAporteIESS.Checked;
				this.txtAfiliacion.Enabled = this.chkNEAporteIESS.Checked;

				if (!this.chkNEAporteIESS.Checked)
				{
					this.chkDecTerAcum.Checked = false;
					this.chkDecCuartoAcum.Checked = false;
					this.chkFondoReserva.Checked = false;
					this.chkDerechoFondosDeReserva.Checked = false;
					this.txtAfiliacion.Text = "";
				}
			}
		}

		private void chkNEImpuestoRenta_CheckedChanged(object sender, System.EventArgs e)
		{
			if (bEdicion)
			{
				this.txtGastoVivienda.Enabled = this.chkNEImpuestoRenta.Checked;
				this.txtGastoSalud.Enabled = this.chkNEImpuestoRenta.Checked;
				this.txtGastoAlimentacion.Enabled = this.chkNEImpuestoRenta.Checked;
				this.txtGastoEducacion.Enabled = this.chkNEImpuestoRenta.Checked;
				this.txtGastoVestimenta.Enabled = this.chkNEImpuestoRenta.Checked;
			
				if (!this.chkNEImpuestoRenta.Checked)
				{
					this.txtGastoVivienda.Value = 0;
					this.txtGastoSalud.Value = 0;
					this.txtGastoAlimentacion.Value = 0;
					this.txtGastoEducacion.Value = 0;
					this.txtGastoVestimenta.Value = 0;
				}
			}
		}

		private void chkNERetencionJudicial_CheckedChanged(object sender, System.EventArgs e)
		{
			if (bEdicion)
			{
				this.txtRetencionJudicial.Enabled = this.chkNERetencionJudicial.Checked;

				if (!this.chkNERetencionJudicial.Checked) this.txtRetencionJudicial.Value = 0;
			}
		}

		private void cmbNEIdentificacion_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Enter) this.txtNEIdentificacion.Focus();
		}

		private void txtNEIdentificacion_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Enter) this.txtNEApellidoPaterno.Focus();
		}

		private void cmbNEProvinciaNacimiento_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Enter) this.cmbCantonNacimiento.Focus();
		}

		private void cmbNECantonNacimiento_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Enter) this.dtpNEFechaNacimiento.Focus();
		}

		private void dtpNEFechaNacimiento_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Enter) this.optNESexo.Focus();
		}

		private void optNESexo_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Enter) this.txtNETipoDeSangre.Focus();
		}

		private void txtNETipoDeSangre_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Enter) this.cmbNEEstadoCivil.Focus();
		}

		private void txtPorcentajeDiscapacidad_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Enter) this.txtConadis.Focus();
		}

		private void txtConadis_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Enter) this.txtTipoDiscapacidad.Focus();
		}

		private void txtTipoDiscapacidad_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Enter) this.chkNELicencia.Focus();
		}

		private void cmbNEEstadoCivil_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Enter) this.txtNECargasFamiliares.Focus();
		}

		private void txtNECargasFamiliares_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Enter) this.chkNEDiscapacidad.Focus();
		}

		private void cmbNEProvincia_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Enter) this.cmbNECanton.Focus();
		}

		private void cmbNECanton_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Enter) this.cmbNEParroquia.Focus();
		}

		private void cmbNEParroquia_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Enter) this.cmbNEBarrio.Focus();
		}

		private void cmbNEBarrio_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Enter) this.txtNEDireccion.Focus();
		}

		private void txtNEDireccion_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Enter) this.txtNETelefono.Focus();
		}

		private void txtNETelefono_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Enter) this.txtNECelular.Focus();
		}

		private void txtNECelular_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Enter) this.txtNEEMail.Focus();
		}

		private void txtNEEMail_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Enter) this.txtNEEMailEmpresarial.Focus();
		}

		private void optNEVivienda_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Enter) this.txtNETiempoDeVivienda.Focus();
		}

		private void cmbNECentroCosto_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Enter) this.cmbNEDepartamento.Focus();
		}

		private void cmbNEDepartamento_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Enter) this.cmbNECargo.Focus();
		}

		private void cmbNECargo_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Enter) this.cmbNECodigoSectorial.Focus();
		}

		private void cmbNECodigoSectorial_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Enter) 
			{
				if (this.cmbRangos.Enabled) this.cmbRangos.Focus(); else this.optNESistema.Focus();
			}
		}

		private void optNESistema_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Enter) this.cmbNETipoContrato.Focus();
		}

		private void cmbNETipoContrato_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Enter) this.txtNESueldo.Focus();
		}

		private void txtAfiliacion_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Enter) this.txtNESueldo.Focus();
		}

		private void txtNESueldo_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Enter) this.txtNEBonos.Focus();
		}

		private void txtNEMovilizacion_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Enter) this.txtNEOrdinarias.Focus();
		}

		private void txtNEHorasExtras_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Enter) this.chkNERetencionJudicial.Focus();
		}

		private void txtNETotalIngresos_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Enter) this.optNETipoDePago.Focus();
		}

		private void optNETipoDePago_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Enter) this.cmbNEBancos.Focus();
		}

		private void txtCuenta_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Enter) this.chkNEImpuestoRenta.Focus();
		}

		private void txtGastoVivienda_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Enter) this.txtGastoSalud.Focus();
		}

		private void txtGastoSalud_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Enter) this.txtGastoAlimentacion.Focus();
		}

		private void txtGastoAlimentacion_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Enter) this.txtGastoEducacion.Focus();
		}

		private void txtGastoEducacion_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Enter) this.txtGastoVestimenta.Focus();
		}

		private void txtGastoVestimenta_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Enter) this.txtNEBono.Focus();
		}

		private void txtRetencionJudicial_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Enter) 
			{
				if (MenuLatinium.bComisionQuincena) this.chkNEQuincena.Focus(); else this.optNETipoDePago.Focus();				
			}
		}

		private void chkNETCC_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Enter) this.chkNEAplicaRol.Focus();
		}

		private void txtNEBono_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Enter) this.txtNEComisiones.Focus();
		}

		private void txtNEComisiones_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Enter) this.txtNEFunciones.Focus();
		}

		private void txtNEFunciones_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Enter) this.txtNEResponsabilidades.Focus();
		}

		private void utcVarios_SelectedTabChanged(object sender, Infragistics.Win.UltraWinTabControl.SelectedTabChangedEventArgs e)
		{
		
		}

		private void uGridFamiliares_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			FuncionesProcedimientos.DesplazamientoGrids(sender, e, this.uGridFamiliares);
		}

		private void btnGuardar_Click(object sender, System.EventArgs e)
		{
			#region Validacion
			#region Datos Generales
			#region Valida identificación
			if (this.txtNEIdentificacion.Text.ToString().Length == 0)
			{
				MessageBox.Show("Esciba el numero de identificacion", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				this.utcNomina.Tabs[0].Selected = true;
				this.txtNEIdentificacion.Focus();
				return;
			}

			if ((int)this.cmbNEIdentificacion.Value == 1)
			{
				string stExec = string.Format("Select dbo.ValidaIdentificacion ('{0}', {1})", this.txtNEIdentificacion.Text.ToString(), 2);
				string VlCed = Funcion.sEscalarSQL(cdsSeteoF, stExec);
				
				if (VlCed.StartsWith("Error"))
				{
					MessageBox.Show(VlCed, "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
					this.utcNomina.Tabs[0].Selected = true;
					this.txtNEIdentificacion.Focus();
					return;
				}
			}
			#endregion Valida identificación
			#region Apellidos - Nombres
			if (this.txtNEApellidoPaterno.Text.ToString().Length == 0)
			{
				MessageBox.Show("Ingrese el Apellido", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				this.utcNomina.Tabs[0].Selected = true;
				this.txtNEApellidoPaterno.Focus();
				return;
			}			
			if (this.txtNEPrimerNombre.Text.ToString().Length == 0)
			{
				MessageBox.Show("Ingrese el Nombre", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				this.utcNomina.Tabs[0].Selected = true;
				this.txtNEPrimerNombre.Focus();
				return;
			}
			#endregion Apellidos - Nombres
			#region Nacimiento
			if (this.cmbNENacionalidad.ActiveRow == null)
			{
				MessageBox.Show("Seleccione la Nacionalidad", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);				
				this.cmbNENacionalidad.Focus();
				return;
			}
			if ((int)this.cmbNENacionalidad.Value == 54)
			{
				if (this.cmbProvinciaNacimiento.ActiveRow == null)
				{
					MessageBox.Show("Seleccione la Provincia de Nacimiento", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
					this.utcNomina.Tabs[0].Selected = true;
					this.cmbProvinciaNacimiento.Focus();
					return;
				}
				if (this.cmbCantonNacimiento.ActiveRow == null)
				{
					MessageBox.Show("Seleccione el Canton de Nacimiento", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
					this.utcNomina.Tabs[0].Selected = true;
					this.cmbCantonNacimiento.Focus();
					return;
				}
			}
			if (this.dtpNEFechaNacimiento.Value == DBNull.Value)
			{
				MessageBox.Show("Ingrese la Fecha de Nacimiento", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				this.utcNomina.Tabs[0].Selected = true;
				this.dtpNEFechaNacimiento.Focus();
				return;
			}
			DateTime dtValFechaNac = DateTime.Today.AddYears(-18);
			if (Convert.ToDateTime(this.dtpNEFechaNacimiento.Value) > dtValFechaNac)
			{
				MessageBox.Show("La edad debe ser mayor o igual a 18 Años", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				this.utcNomina.Tabs[0].Selected = true;
				this.dtpNEFechaNacimiento.Focus();
				return;
			}
			if (this.optNESexo.Value == System.DBNull.Value)
			{
				MessageBox.Show("Seleccione Hombre/Mujer", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				this.utcNomina.Tabs[0].Selected = true;
				this.optNESexo.Focus();
				return;
			}
			if (this.txtNETipoDeSangre.Text.ToString().Length == 0)
			{
				MessageBox.Show("Ingrese el Tipo de Sangre", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				this.utcNomina.Tabs[0].Selected = true;
				this.txtNETipoDeSangre.Focus();
				return;
			}
			#endregion Nacimiento
			#region Discapacidad 
			if (this.chkNEDiscapacidad.Checked)
			{
				if ((int)this.txtPorcentajeDiscapacidad.Value == 0)
				{
					MessageBox.Show("ingrese el Porcentaje de Discapacidad", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
					this.utcNomina.Tabs[0].Selected = true;
					this.txtPorcentajeDiscapacidad.Focus();
					return;				
				}				
				if (this.txtConadis.Text.ToString().Length == 0)
				{
					MessageBox.Show("Esciba el Apellido", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
					this.utcNomina.Tabs[0].Selected = true;
					this.txtConadis.Focus();
					return;
				}			
				if (this.txtTipoDiscapacidad.Text.ToString().Length == 0)
				{
					MessageBox.Show("Esciba el Apellido", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
					this.utcNomina.Tabs[0].Selected = true;
					this.txtTipoDiscapacidad.Focus();
					return;
				}				
			}
			#endregion Discapacidad
			#region Estado civil
			if (this.cmbNEEstadoCivil.ActiveRow == null)
			{
				MessageBox.Show("Seleccione el Estado Civil", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				this.utcNomina.Tabs[0].Selected = true;
				this.cmbNEEstadoCivil.Focus();
				return;
			}					
			#endregion Estado civil
			#region Vivienda
			if (this.cmbNEProvincia.ActiveRow == null)
			{
				MessageBox.Show("Seleccione la Provincia de Domicilio", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				this.utcNomina.Tabs[0].Selected = true;
				this.cmbProvinciaNacimiento.Focus();
				return;
			}
			if (this.cmbNECanton.ActiveRow == null)
			{
				MessageBox.Show("Seleccione el Canton de Domicilio", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				this.utcNomina.Tabs[0].Selected = true;
				this.cmbNECanton.Focus();
				return;
			}
			if (this.cmbNEParroquia.ActiveRow == null)
			{
				MessageBox.Show("Seleccione la Parroquia de Domicilio", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				this.utcNomina.Tabs[0].Selected = true;
				this.cmbNEParroquia.Focus();
				return;
			}
			if (this.cmbNEBarrio.ActiveRow == null)
			{
				MessageBox.Show("Seleccione el Barrio de Domicilio", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				this.utcNomina.Tabs[0].Selected = true;
				this.cmbNEBarrio.Focus();
				return;
			}
			if (this.txtNEDireccion.Text.ToString().Length == 0)
			{
				MessageBox.Show("Ingrese la Direccion de domicilio", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				this.utcNomina.Tabs[0].Selected = true;
				this.txtNEDireccion.Focus();
				return;
			}
			if (this.txtNETelefono.Text.ToString().Length == 0 && this.txtNECelular.Text.ToString().Length == 0)
			{
				MessageBox.Show("Ingrese un Numero de Telefono Convencional o Celular", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				this.utcNomina.Tabs[0].Selected = true;
				this.txtNETelefono.Focus();
				return;
			}
			if (this.txtNETelefono.Text.ToString().Length > 0)
			{
				if (this.txtNETelefono.Text.ToString().Length != 9)
				{
					MessageBox.Show("El Numero de Telefono Convencional debe tener 9 digitos anteponiendo el codigo de provincia", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
					this.utcNomina.Tabs[0].Selected = true;
					this.txtNETelefono.Focus();
					return;
				}
			}
			if (this.txtNECelular.Text.ToString().Length > 0)
			{
				if (this.txtNECelular.Text.ToString().Length != 10)
				{
					MessageBox.Show("El Numero Celular debe tener 10 digitos", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
					this.utcNomina.Tabs[0].Selected = true;
					this.txtNECelular.Focus();
					return;
				}
			}
			if (this.optNEVivienda.Value == System.DBNull.Value)
			{
				MessageBox.Show("Seleccione el Tipo de vivienda", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				this.utcNomina.Tabs[0].Selected = true;
				this.optNEVivienda.Focus();
				return;
			}
			if ((int)this.txtNETiempoDeVivienda.Value == 0)
			{
				MessageBox.Show("Ingrese el Tiempo de vivienda", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				this.utcNomina.Tabs[0].Selected = true;
				this.txtNETiempoDeVivienda.Focus();
				return;
			}
			if (this.optNEMesesAños.Value == System.DBNull.Value)
			{
				MessageBox.Show("Seleccione el Tiempo de vivienda (Meses o Años)", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				this.utcNomina.Tabs[0].Selected = true;
				this.optNEMesesAños.Focus();
				return;
			}	
			if (this.txtNEVive_Con.Text.ToString().Length == 0)
			{
				MessageBox.Show("Debe registrar con quien vive", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				this.utcNomina.Tabs[0].Selected = true;
				this.txtNEVive_Con.Focus();
				return;
			}
			if ((int)this.optNEVivienda.Value == 1)
			{
				if (this.txtPredial.Text.ToString().Length == 0)
				{
					MessageBox.Show("Si la vivienda es propia debe registrar el Predial", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
					this.utcNomina.Tabs[0].Selected = true;
					this.txtPredial.Focus();
					return;
				}
			}			
			#endregion Vivienda
			#region Email

			if (!Validacion.vbTexto(this.txtNEEMail, 1, 100, " E-Mail personal")) 
			{
				this.txtNEEMail.Focus(); 
				return;
			}
			if (this.txtNEEMail.Text.ToString().Trim().Length > 0 )
			{
				if (!Validacion.bEMail(this.txtNEEMail.Text.ToString())) 
				{
					this.txtNEEMail.Focus();
					return;
				}
			
			}
			
			#endregion Email
			#endregion Datos Generales
			if(MenuLatinium.stIdDB == 1 || MenuLatinium.stIdDB == 33)
			{
				if (!Validacion.vbComboVacio(this.cmbRegion, "Seleccione la Región")) return;
			}

			#region Contratacion
			#region Ubicacion
			if (this.cmbNECentroCosto.ActiveRow == null)
			{
				MessageBox.Show("Seleccione el Centro de Costo", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				this.utcNomina.Tabs[1].Selected = true;
				this.cmbNECentroCosto.Focus();
				return;
			}
			if (this.cmbNEDepartamento.ActiveRow == null)
			{
				MessageBox.Show("Seleccione el Departamento", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				this.utcNomina.Tabs[1].Selected = true;
				this.cmbNEDepartamento.Focus();
				return;
			}
			if (this.cmbNECargo.ActiveRow == null)
			{
				MessageBox.Show("Seleccione el Cargo", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				this.utcNomina.Tabs[1].Selected = true;
				this.cmbNECargo.Focus();
				return;
			}
			if (this.cmbNECodigoSectorial.ActiveRow == null)
			{
				MessageBox.Show("Seleccione el Codigo Sectorial", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				this.utcNomina.Tabs[1].Selected = true;
				this.cmbNECodigoSectorial.Focus();
				return;
			}
			#endregion Ubicacion
			#region Tipo de Contrato
			if (this.cmbNETipoContrato.ActiveRow == null)
			{
				MessageBox.Show("Seleccione el Tipo de Contrato", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				this.utcNomina.Tabs[1].Selected = true;
				this.cmbNETipoContrato.Focus();
				return;
			}
			if (this.chkNEAporteIESS.Checked)
			{
				if (this.txtAfiliacion.Text.ToString().Length == 0)
				{
					MessageBox.Show("Ingrese el Numero de Afiliacion", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
					this.utcNomina.Tabs[1].Selected = true;
					this.txtAfiliacion.Focus();
					return;
				}
			}
			#endregion Tipo de Contrato
			#region Ingresos
			if (Convert.ToDecimal(this.txtNESueldo.Value) == 0)	
			{
				MessageBox.Show("Ingrese el Sueldo", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				this.utcNomina.Tabs[1].Selected = true;
				this.txtNESueldo.Focus();
				return;
			}

//			decimal dSueldoBasico = Funcion.decEscalarSQL(cdsSeteoF, "SELECT SueldoBasico FROM Seteo");
//			
//			if (Convert.ToDecimal(this.txtNESueldo.Value) < dSueldoBasico)
//			{
//				MessageBox.Show(string.Format("El Sueldo no puede ser menor a {0}", dSueldoBasico), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
//				this.utcNomina.Tabs[1].Selected = true;
//				this.txtNESueldo.Focus();
//				return;
//			}

			if (this.optNETipoDePago.Value == System.DBNull.Value)
			{
				MessageBox.Show("Seleccione un Tipo de Pago", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				this.utcNomina.Tabs[1].Selected = true;
				this.optNETipoDePago.Focus();
				return;
			}			

			if ((int)this.optNETipoDePago.Value != 4)
			{
				if (this.cmbNEBancos.ActiveRow == null)
				{
					MessageBox.Show("Seleccione un Banco", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
					this.utcNomina.Tabs[1].Selected = true;
					this.cmbNEBancos.Focus();
					return;
				}
				if (this.txtCuenta.Text.ToString().Length == 0)
				{
					MessageBox.Show("Ingrese el Numero de Cuenta", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
					this.utcNomina.Tabs[1].Selected = true;
					this.txtCuenta.Focus();
					return;
				}
			}
			if (MenuLatinium.bComisionQuincena)
			{
				if (this.chkNEQuincena.Checked)
				{
					if (Convert.ToDecimal(this.txtQuincena.Value) == 0.00m)
					{
						MessageBox.Show("Ingrese el Valor de la Quincena", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
						this.utcNomina.Tabs[1].Selected = true;
						this.txtQuincena.Focus();
						return;
					}
				}
			}
			#endregion Ingresos
			#region Funciones - Responsabilidades	
			if (this.txtNEFunciones.Text.ToString().Length == 0)
			{
				MessageBox.Show("Ingrese las Funciones del Empleado", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				this.utcNomina.Tabs[1].Selected = true;
				this.txtNEFunciones.Focus();
				return;
			}			
			if (this.txtNEResponsabilidades.Text.ToString().Length == 0)
			{
				MessageBox.Show("Ingrese las Responsabilidades del Empleado", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				this.utcNomina.Tabs[1].Selected = true;
				this.txtNEResponsabilidades.Focus();
				return;
			}			
			#endregion Funciones - Responsabilidades
			#endregion Contratacion			

			DateTime dtValFecha = DateTime.Today;

			#region Grids
			#region Familiares
			if (this.uGridFamiliares.Rows.Count < 1)
			{
				MessageBox.Show("Debe Ingresar al menos 1 referencia", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				this.utcNomina.Tabs[2].Selected = true;
				this.utcVarios.Tabs[0].Selected = true;
				this.uGridFamiliares.Focus();
				return;
			}

			foreach(Infragistics.Win.UltraWinGrid.UltraGridRow drValidaFamiliares in this.uGridFamiliares.Rows.All)
			{	
				#region Parentesco
				int idParentesco = 0;
				if (drValidaFamiliares.Cells["idParentesco"].Value != System.DBNull.Value) idParentesco = (int)drValidaFamiliares.Cells["idParentesco"].Value;	

				if (idParentesco == 0)
				{
					MessageBox.Show("Seleccione el Parentesco de la Referencia Familiar Seleccionada", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
					this.utcNomina.Tabs[2].Selected = true;
					this.utcVarios.Tabs[0].Selected = true;
					this.uGridFamiliares.ActiveRow = drValidaFamiliares;
					this.uGridFamiliares.ActiveRow.Selected = true;
					return;
				}
				#endregion Parentesco

				#region Nombres
				if (drValidaFamiliares.Cells["Nombre"].Value.ToString().Trim().Length == 0)
				{
					MessageBox.Show("Ingrese el Nombre de la Referencia Familiar Seleccionada", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
					this.utcNomina.Tabs[2].Selected = true;
					this.utcVarios.Tabs[0].Selected = true;
					this.uGridFamiliares.ActiveRow = drValidaFamiliares;
					this.uGridFamiliares.ActiveRow.Selected = true;
					return;
				}
				#endregion Nombres

				#region Fecha Nacimiento
//				if (drValidaFamiliares.Cells["FechaNacimiento"].Value == System.DBNull.Value)
//				{
//					MessageBox.Show("Ingrese la Fecha de Nacimiento de la Referencia Familiar Seleccionada", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
//					this.utcNomina.Tabs[2].Selected = true;
//					this.utcVarios.Tabs[0].Selected = true;
//					this.uGridFamiliares.ActiveRow = drValidaFamiliares;
//					this.uGridFamiliares.ActiveRow.Selected = true;
//					return;
//				}
				#endregion Fecha Nacimiento

				#region Telefonos Vacios
				if (drValidaFamiliares.Cells["Telefono"].Value.ToString().Trim().Length == 0 && drValidaFamiliares.Cells["Celular"].Value.ToString().Trim().Length == 0)
				{
					MessageBox.Show("Debe Ingresar el Telefono Convencional o el Telefono Celular de la Referencia Familiar Seleccionada", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
					this.utcNomina.Tabs[2].Selected = true;
					this.utcVarios.Tabs[0].Selected = true;
					this.uGridFamiliares.ActiveRow = drValidaFamiliares;
					this.uGridFamiliares.ActiveRow.Selected = true;
					return;
				}
				#endregion Telefonos Vacios

				#region Telefono
				if (drValidaFamiliares.Cells["Telefono"].Value.ToString().Trim().Length > 0)
				{
					if (drValidaFamiliares.Cells["Telefono"].Value.ToString().Trim().Length != 9)
					{
						MessageBox.Show("El Telefono Convencional debe tener 9 Digitos.\n\nLos dos primeros corresponden al Codigo de Provincia", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
						this.utcNomina.Tabs[2].Selected = true;
						this.utcVarios.Tabs[0].Selected = true;
						this.uGridFamiliares.ActiveRow = drValidaFamiliares;
						this.uGridFamiliares.ActiveRow.Selected = true;
						return;
					}
				}
				#endregion Telefono

				#region Celular
				if (drValidaFamiliares.Cells["Celular"].Value.ToString().Trim().Length > 0)
				{
					if (drValidaFamiliares.Cells["Celular"].Value.ToString().Trim().Length != 10)
					{
						MessageBox.Show("El Telefono Celular debe tener 10 digitos", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
						this.utcNomina.Tabs[2].Selected = true;
						this.utcVarios.Tabs[0].Selected = true;
						this.uGridFamiliares.ActiveRow = drValidaFamiliares;
						this.uGridFamiliares.ActiveRow.Selected = true;
						return;
					}
				}
				#endregion Celular
			}
			#endregion Familiares

			#region Estudios
			if (this.uGridEstudios.Rows.Count == 0)
			{
				MessageBox.Show("Ingrese al menos un Estudio Realizado", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				this.utcNomina.Tabs[2].Selected = true;
				this.utcVarios.Tabs[1].Selected = true;
				this.uGridEstudios.Focus();
				return;
			}

			foreach(Infragistics.Win.UltraWinGrid.UltraGridRow drValidaEstudios in this.uGridEstudios.Rows.All)
			{
				#region Nivel
				int idNivel = 0;
				if (drValidaEstudios.Cells["idNivel"].Value != System.DBNull.Value) idNivel = (int)drValidaEstudios.Cells["idNivel"].Value;	

				if (idNivel == 0)
				{
					MessageBox.Show("Seleccione el Nivel de Estudios", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
					this.utcNomina.Tabs[2].Selected = true;
					this.utcVarios.Tabs[1].Selected = true;
					this.uGridEstudios.ActiveRow = drValidaEstudios;
					this.uGridEstudios.ActiveRow.Selected = true;
					return;
				}
				#endregion Nivel

				#region Titulo
				if (drValidaEstudios.Cells["Titulo"].Value.ToString().Trim().Length == 0)
				{
					MessageBox.Show("Ingrese el Título Obtenido", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
					this.utcNomina.Tabs[2].Selected = true;
					this.utcVarios.Tabs[1].Selected = true;
					this.uGridEstudios.ActiveRow = drValidaEstudios;
					this.uGridEstudios.ActiveRow.Selected = true;
					return;
				}
				#endregion Titulo

				#region Institucion
				if (drValidaEstudios.Cells["Institucion"].Value.ToString().Trim().Length == 0)
				{
					MessageBox.Show("Ingrese el Nombre de la Institución", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
					this.utcNomina.Tabs[2].Selected = true;
					this.utcVarios.Tabs[1].Selected = true;
					this.uGridEstudios.ActiveRow = drValidaEstudios;
					this.uGridEstudios.ActiveRow.Selected = true;
					return;
				}
				#endregion Institucion

				#region Tiempo
				if (drValidaEstudios.Cells["Tiempo"].Value == System.DBNull.Value)
				{
					MessageBox.Show("Ingrese el Tiempo de Estudios", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
					this.utcNomina.Tabs[2].Selected = true;
					this.utcVarios.Tabs[1].Selected = true;
					this.uGridEstudios.ActiveRow = drValidaEstudios;
					this.uGridEstudios.ActiveRow.Selected = true;
					return;
				}

				if ((int)drValidaEstudios.Cells["Tiempo"].Value == 0)
				{
					MessageBox.Show("El Tiempo de Estudios debe ser mayor a Cero", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
					this.utcNomina.Tabs[2].Selected = true;
					this.utcVarios.Tabs[1].Selected = true;
					this.uGridEstudios.ActiveRow = drValidaEstudios;
					this.uGridEstudios.ActiveRow.Selected = true;
					return;
				}
				#endregion Tiempo

				#region Periodos
				int idPeriodosDeTiempo = 0;
				if (drValidaEstudios.Cells["idPeriodosDeTiempo"].Value != System.DBNull.Value) idPeriodosDeTiempo = (int)drValidaEstudios.Cells["idPeriodosDeTiempo"].Value;	

				if (idPeriodosDeTiempo == 0)
				{
					MessageBox.Show("Seleccione el Periodo", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
					this.utcNomina.Tabs[2].Selected = true;
					this.utcVarios.Tabs[1].Selected = true;
					this.uGridEstudios.ActiveRow = drValidaEstudios;
					this.uGridEstudios.ActiveRow.Selected = true;
					return;
				}
				#endregion Periodos

				#region Estado
				int idEstadoEstudios = 0;
				if (drValidaEstudios.Cells["idEstadoEstudios"].Value != System.DBNull.Value) idEstadoEstudios = (int)drValidaEstudios.Cells["idEstadoEstudios"].Value;	

				if (idEstadoEstudios == 0)
				{
					MessageBox.Show("Seleccione el Estado de los Estudios", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
					this.utcNomina.Tabs[2].Selected = true;
					this.utcVarios.Tabs[1].Selected = true;
					this.uGridEstudios.ActiveRow = drValidaEstudios;
					this.uGridEstudios.ActiveRow.Selected = true;
					return;
				}
				#endregion Estado

				#region Año Finalizacion
				if (idEstadoEstudios != 2)
				{
					if (drValidaEstudios.Cells["AñoFinalizacion"].Value == System.DBNull.Value)
					{
						MessageBox.Show("Ingrese el Año de Finalización de los Estudios", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
						this.utcNomina.Tabs[2].Selected = true;
						this.utcVarios.Tabs[1].Selected = true;
						this.uGridEstudios.ActiveRow = drValidaEstudios;
						this.uGridEstudios.ActiveRow.Selected = true;
						return;
					}

					int iAñoNacimiento = Convert.ToDateTime(this.dtpNEFechaNacimiento.Value).Year;
					int iAñoActual = DateTime.Today.Year;
					iAñoNacimiento = iAñoNacimiento + 15;

					if ((int)drValidaEstudios.Cells["AñoFinalizacion"].Value < iAñoNacimiento)
					{
						MessageBox.Show(string.Format("El Año de Finalización no puede ser menor a {0}", iAñoNacimiento), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
						this.utcNomina.Tabs[2].Selected = true;
						this.utcVarios.Tabs[1].Selected = true;
						this.uGridEstudios.ActiveRow = drValidaEstudios;
						this.uGridEstudios.ActiveRow.Selected = true;
						return;
					}

					if ((int)drValidaEstudios.Cells["AñoFinalizacion"].Value > iAñoActual)
					{
						MessageBox.Show(string.Format("El Año de Finaclización no puede ser mayor a {0}", iAñoActual), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
						this.utcNomina.Tabs[2].Selected = true;
						this.utcVarios.Tabs[1].Selected = true;
						this.uGridEstudios.ActiveRow = drValidaEstudios;
						this.uGridEstudios.ActiveRow.Selected = true;
						return;
					}
				}
				#endregion Año Finalizacion
			}
			#endregion Estudios

			#region Experiencia
			if (this.uGridExperiencia.Rows.Count > 0)
			{
				foreach(Infragistics.Win.UltraWinGrid.UltraGridRow drValidaExperiencia in this.uGridExperiencia.Rows.All)
				{
					#region Empresa
					if (drValidaExperiencia.Cells["Empresa"].Value.ToString().Trim().Length == 0)
					{
						MessageBox.Show("Ingrese La Empresa o Institución en la Ficha de Experiencia", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
						this.utcNomina.Tabs[2].Selected = true;
						this.utcVarios.Tabs[2].Selected = true;
						this.uGridExperiencia.ActiveRow = drValidaExperiencia;
						this.uGridExperiencia.ActiveRow.Selected = true;
						return;
					}
					#endregion Empresa

					#region Actividades
					if (drValidaExperiencia.Cells["Actividades"].Value.ToString().Trim().Length == 0)
					{
						MessageBox.Show(string.Format("Ingrese las Actividades que realizaba en la Empresa '{0}'", drValidaExperiencia.Cells["Empresa"].Value.ToString().Trim()), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
						this.utcNomina.Tabs[2].Selected = true;
						this.utcVarios.Tabs[2].Selected = true;
						this.uGridExperiencia.ActiveRow = drValidaExperiencia;
						this.uGridExperiencia.ActiveRow.Selected = true;
						return;
					}
					#endregion Actividades

					#region Afiliado IESS
//					if ((bool)drValidaExperiencia.Cells["AfiliadoIESS"].Value)
//					{
//						if ((int)drValidaExperiencia.Cells["Aportaciones"].Value == 0)
//						{
//							MessageBox.Show(string.Format("Ingrese el número de aportaciones al I.E.S.S. en la Empresa '{0}'", drValidaExperiencia.Cells["Empresa"].Value.ToString().Trim()), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
//							this.utcNomina.Tabs[2].Selected = true;
//							this.utcVarios.Tabs[2].Selected = true;
//							this.uGridExperiencia.ActiveRow = drValidaExperiencia;
//							this.uGridExperiencia.ActiveRow.Selected = true;
//							return;
//						}
//					}
					#endregion Afiliado IESS

					#region Sin Problemas - Con Problemas
					if (!(bool)drValidaExperiencia.Cells["SinProblemas"].Value && !(bool)drValidaExperiencia.Cells["ConProblemas"].Value)
					{
						MessageBox.Show(string.Format("Debe seleccionar si Salió con o sin problemas de la Empresa '{0}'", drValidaExperiencia.Cells["Empresa"].Value.ToString().Trim()), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
						this.utcNomina.Tabs[2].Selected = true;
						this.utcVarios.Tabs[2].Selected = true;
						this.uGridExperiencia.ActiveRow = drValidaExperiencia;
						this.uGridExperiencia.ActiveRow.Selected = true;
						return;
					}

					if ((bool)drValidaExperiencia.Cells["SinProblemas"].Value && (bool)drValidaExperiencia.Cells["ConProblemas"].Value)
					{
						MessageBox.Show(string.Format("No puede seleccionar las dos opciones, confirme si salió Con Problemas o Sin Problemas de la Empresa '{0}'", drValidaExperiencia.Cells["Empresa"].Value.ToString().Trim()), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
						this.utcNomina.Tabs[2].Selected = true;
						this.utcVarios.Tabs[2].Selected = true;
						this.uGridExperiencia.ActiveRow = drValidaExperiencia;
						this.uGridExperiencia.ActiveRow.Selected = true;
						return;
					}

					if ((bool)drValidaExperiencia.Cells["ConProblemas"].Value)
					{
						if (drValidaExperiencia.Cells["DescripcionProblema"].Value.ToString().Trim().Length == 0)
						{
							MessageBox.Show(string.Format("Si Salió Con Problemas de la Empresa '{0}', Ingrese la descripción del problema", drValidaExperiencia.Cells["Empresa"].Value.ToString().Trim()), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
							this.utcNomina.Tabs[2].Selected = true;
							this.utcVarios.Tabs[2].Selected = true;
							this.uGridExperiencia.ActiveRow = drValidaExperiencia;
							this.uGridExperiencia.ActiveRow.Selected = true;
							return;
						}
					}
					#endregion Sin Problemas - Con Problemas
				}
			}
			#endregion Experiencia

			#region Entrada - Salida
			if (this.uGridEntradasSalidas.Rows.Count == 0)
			{
				MessageBox.Show("Ingrese los datos de Entrada", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				this.utcNomina.Tabs[2].Selected = true;
				this.utcVarios.Tabs[3].Selected = true;
				this.uGridEntradasSalidas.Focus();
				return;
			}
			
			int iCont = 0;
			foreach(Infragistics.Win.UltraWinGrid.UltraGridRow drValidaEntradaSalida in this.uGridEntradasSalidas.Rows.All)
			{
				#region Fecha de Entrada
				if (drValidaEntradaSalida.Cells["FechaEntrada"].Value == System.DBNull.Value)
				{
					MessageBox.Show("Ingrese la Fecha de Entrada", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
					this.utcNomina.Tabs[2].Selected = true;
					this.utcVarios.Tabs[3].Selected = true;
					this.uGridEntradasSalidas.ActiveRow = drValidaEntradaSalida;
					this.uGridEntradasSalidas.ActiveRow.Selected = true;
					return;
				}
				
//				if (Convert.ToDateTime(drValidaEntradaSalida.Cells["FechaEntrada"].Value) > dtValFecha)
//				{
//					MessageBox.Show("La Fecha de Entrada no puede ser mayor a Hoy", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
//					this.utcNomina.Tabs[2].Selected = true;
//					this.utcVarios.Tabs[3].Selected = true;
//					this.uGridEntradasSalidas.ActiveRow = drValidaEntradaSalida;
//					this.uGridEntradasSalidas.ActiveRow.Selected = true;
//					return;
//				}

				if (drValidaEntradaSalida.Cells["FechaSalida"].Value != System.DBNull.Value)
				{
					if (Convert.ToDateTime(drValidaEntradaSalida.Cells["FechaEntrada"].Value) > Convert.ToDateTime(drValidaEntradaSalida.Cells["FechaSalida"].Value))
					{
						MessageBox.Show("La Fecha de Entrada no puede ser mayor a la Fecha de Salida", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
						this.utcNomina.Tabs[2].Selected = true;
						this.utcVarios.Tabs[3].Selected = true;
						this.uGridEntradasSalidas.ActiveRow = drValidaEntradaSalida;
						this.uGridEntradasSalidas.ActiveRow.Selected = true;
						return;
					}
				}

				if (drValidaEntradaSalida.Cells["FechaTerminaPrueba"].Value != System.DBNull.Value)
				{
					if (Convert.ToDateTime(drValidaEntradaSalida.Cells["FechaEntrada"].Value) > Convert.ToDateTime(drValidaEntradaSalida.Cells["FechaTerminaPrueba"].Value))
					{
						MessageBox.Show("La Fecha de Entrada no puede ser mayor a la Fecha de Termino de Pruebas", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
						this.utcNomina.Tabs[2].Selected = true;
						this.utcVarios.Tabs[3].Selected = true;
						this.uGridEntradasSalidas.ActiveRow = drValidaEntradaSalida;
						this.uGridEntradasSalidas.ActiveRow.Selected = true;
						return;
					}
				}
				#endregion Fecha de Entrada

				#region Fecha Termina Prueba
				if (drValidaEntradaSalida.Cells["FechaTerminaPrueba"].Value == System.DBNull.Value)
				{
					MessageBox.Show("Ingrese la Fecha de Término de Prueba", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
					this.utcNomina.Tabs[2].Selected = true;
					this.utcVarios.Tabs[3].Selected = true;
					this.uGridEntradasSalidas.ActiveRow = drValidaEntradaSalida;
					this.uGridEntradasSalidas.ActiveRow.Selected = true;
					return;
				}
				
				dtValFecha = dtValFecha.AddDays(90);
//				if (Convert.ToDateTime(drValidaEntradaSalida.Cells["FechaTerminaPrueba"].Value) > dtValFecha)
//				{
//					MessageBox.Show(string.Format("La Fecha de Término de Prueba no puede ser mayor a '{0}'", dtValFecha.ToString("dd/MM/yyyy")), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
//					this.utcNomina.Tabs[2].Selected = true;
//					this.utcVarios.Tabs[3].Selected = true;
//					this.uGridEntradasSalidas.ActiveRow = drValidaEntradaSalida;
//					this.uGridEntradasSalidas.ActiveRow.Selected = true;
//					return;
//				}

				if (drValidaEntradaSalida.Cells["FechaEntrada"].Value != System.DBNull.Value)
				{
					if (Convert.ToDateTime(drValidaEntradaSalida.Cells["FechaEntrada"].Value) > Convert.ToDateTime(drValidaEntradaSalida.Cells["FechaTerminaPrueba"].Value))
					{
						MessageBox.Show("La Fecha de Término de Prueba no puede ser menor a la Fecha de Entrada", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
						this.utcNomina.Tabs[2].Selected = true;
						this.utcVarios.Tabs[3].Selected = true;
						this.uGridEntradasSalidas.ActiveRow = drValidaEntradaSalida;
						this.uGridEntradasSalidas.ActiveRow.Selected = true;
						return;
					}
				}
				#endregion Fecha Termina Prueba

				#region Registro de Entrada
				if (!(bool)drValidaEntradaSalida.Cells["APrueba"].Value)
				{
					if (drValidaEntradaSalida.Cells["NEntrada"].Value.ToString().Trim().Length == 0)
					{
						MessageBox.Show("Ingrese el Registro de Entrada", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
						this.utcNomina.Tabs[2].Selected = true;
						this.utcVarios.Tabs[3].Selected = true;
						this.uGridEntradasSalidas.ActiveRow = drValidaEntradaSalida;
						this.uGridEntradasSalidas.ActiveRow.Selected = true;
						return;
					}					
				}
				#endregion Registro de Entrada


				int iValidaFechaPrueba = Funcion.iEscalarSQL(cdsSeteoF, string.Format(" Exec [ValidaFechaPeriodoPrueba] '{0}', '{1}'",
					Convert.ToDateTime(drValidaEntradaSalida.Cells["FechaEntrada"].Value).ToString("yyyyMMdd") , Convert.ToDateTime(drValidaEntradaSalida.Cells["FechaTerminaPrueba"].Value).ToString("yyyyMMdd") ));	
	
				if(iValidaFechaPrueba == 1)
				{
					MessageBox.Show("La Fecha de termina prueba debe ser minimo 60 dias de la Fecha de Entrada ", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
					this.utcNomina.Tabs[2].Selected = true;
					this.utcVarios.Tabs[3].Selected = true;
					this.uGridEntradasSalidas.ActiveRow = drValidaEntradaSalida;
					this.uGridEntradasSalidas.ActiveRow.Selected = true;
					return;
				}
				if (iValidaFechaPrueba == 2)
				{
					MessageBox.Show("La Fecha de termina prueba debe ser maximo 90 dias de la Fecha de Entrada ", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
					this.utcNomina.Tabs[2].Selected = true;
					this.utcVarios.Tabs[3].Selected = true;
					this.uGridEntradasSalidas.ActiveRow = drValidaEntradaSalida;
					this.uGridEntradasSalidas.ActiveRow.Selected = true;
					return;
				}

				#region Verificado por
//				int idVerificado = 0;
//				if (drValidaEntradaSalida.Cells["idVerificadoPor"].Value != System.DBNull.Value) idVerificado = (int)drValidaEntradaSalida.Cells["idVerificadoPor"].Value;	
//
//				if (idVerificado == 0)
//				{
//					MessageBox.Show("Seleccione el Nombre de quien verifica la entrada", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
//					this.utcNomina.Tabs[2].Selected = true;
//					this.utcVarios.Tabs[3].Selected = true;
//					this.uGridEntradasSalidas.ActiveRow = drValidaEntradaSalida;
//					this.uGridEntradasSalidas.ActiveRow.Selected = true;
//					return;
//				}
				#endregion Verificado por

				#region Autorizado por
//				int idAutorizado = 0;
//				if (drValidaEntradaSalida.Cells["idAutorizadoPor"].Value != System.DBNull.Value) idAutorizado = (int)drValidaEntradaSalida.Cells["idAutorizadoPor"].Value;	
//
//				if (idAutorizado == 0)
//				{
//					MessageBox.Show("Seleccione el Nombre de quien autorizo la entrada", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
//					this.utcNomina.Tabs[2].Selected = true;
//					this.utcVarios.Tabs[3].Selected = true;
//					this.uGridEntradasSalidas.ActiveRow = drValidaEntradaSalida;
//					this.uGridEntradasSalidas.ActiveRow.Selected = true;
//					return;
//				}
				#endregion Autorizado por

				dtValFecha = DateTime.Today;

				#region Salida
				if (drValidaEntradaSalida.Cells["FechaSalida"].Value != System.DBNull.Value)
				{									
					#region Fecha de Salida
					if (Convert.ToDateTime(drValidaEntradaSalida.Cells["FechaSalida"].Value) > dtValFecha)
					{
						MessageBox.Show("La Fecha de Salida no puede ser mayor a hoy", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
						this.utcNomina.Tabs[2].Selected = true;
						this.utcVarios.Tabs[3].Selected = true;
						this.uGridEntradasSalidas.ActiveRow = drValidaEntradaSalida;
						this.uGridEntradasSalidas.ActiveRow.Selected = true;
						return;
					}
					
					if (Convert.ToDateTime(drValidaEntradaSalida.Cells["FechaEntrada"].Value) > Convert.ToDateTime(drValidaEntradaSalida.Cells["FechaSalida"].Value))
					{
						MessageBox.Show("La Fecha de Salida no puede ser menor a la Fecha de Entrada", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
						this.utcNomina.Tabs[2].Selected = true;
						this.utcVarios.Tabs[3].Selected = true;
						this.uGridEntradasSalidas.ActiveRow = drValidaEntradaSalida;
						this.uGridEntradasSalidas.ActiveRow.Selected = true;
						return;
					}				
					#endregion Fecha de Salida

					#region Registro de Salida
					if (drValidaEntradaSalida.Cells["NSalida"].Value.ToString().Trim().Length == 0)
					{
						MessageBox.Show("Ingrese el Registro de Salida", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
						this.utcNomina.Tabs[2].Selected = true;
						this.utcVarios.Tabs[3].Selected = true;
						this.uGridEntradasSalidas.ActiveRow = drValidaEntradaSalida;
						this.uGridEntradasSalidas.ActiveRow.Selected = true;
						return;
					}					
					#endregion Registro de Salida

					#region Motivo de Salida
					int idMotivoSalida = 0;
					if (drValidaEntradaSalida.Cells["idMotivoSalida"].Value != System.DBNull.Value) idMotivoSalida = (int)drValidaEntradaSalida.Cells["idMotivoSalida"].Value;	

					if (idMotivoSalida == 0)
					{
						MessageBox.Show("Seleccione el Motivo de Sálida", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
						this.utcNomina.Tabs[2].Selected = true;
						this.utcVarios.Tabs[3].Selected = true;
						this.uGridEntradasSalidas.ActiveRow = drValidaEntradaSalida;
						this.uGridEntradasSalidas.ActiveRow.Selected = true;
						return;
					}
					#endregion Motivo de Salida

					#region Recuento de Salida
					if (drValidaEntradaSalida.Cells["Descripcion"].Value.ToString().Trim().Length == 0)
					{
						MessageBox.Show("Ingrese el Recuento de Salida", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
						this.utcNomina.Tabs[2].Selected = true;
						this.utcVarios.Tabs[3].Selected = true;
						this.uGridEntradasSalidas.ActiveRow = drValidaEntradaSalida;
						this.uGridEntradasSalidas.ActiveRow.Selected = true;
						return;
					}					
					#endregion Recuento de Salida
				}
				#endregion Salida

				if (drValidaEntradaSalida.Cells["FechaSalida"].Value == System.DBNull.Value) iCont++;				
			}

			if (!bNuevo)
			{
				if (iCont == 0) 
				{
					bEstado = false; 
					this.txtEstado.Value = 2;

//					this.chkDerechoFondosDeReserva.Checked = false;
//					this.chkFondoReserva.Checked = false;
				}
				else 
				{
					if ((int)this.txtEstado.Value == 2) this.txtEstado.Value = 3;// else this.txtEstado.Value = 1;
					bEstado = true;
				}
			}

			if (iCont > 1)
			{
				MessageBox.Show("No puede tener mas de un registro de entrada sin el registro de salida", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				this.utcNomina.Tabs[2].Selected = true;
				this.utcVarios.Tabs[3].Selected = true;
				this.uGridEntradasSalidas.Focus();
				return;
			}
			#endregion Entrada - Salida

			#region Equipos
			if (this.uGridEquipos.Rows.Count > 0)
			{
				foreach(Infragistics.Win.UltraWinGrid.UltraGridRow drValidaEquipos in this.uGridEquipos.Rows.All)
				{
					#region Articulo
					if (drValidaEquipos.Cells["Articulo"].Value.ToString().Trim().Length == 0)
					{
						MessageBox.Show("Ingrese el nombre del articulo entregado", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
						this.utcNomina.Tabs[2].Selected = true;
						this.utcVarios.Tabs[4].Selected = true;
						this.uGridEquipos.ActiveRow = drValidaEquipos;
						this.uGridEquipos.ActiveRow.Selected = true;
						return;
					}
					#endregion Articulo

					#region Serial
					if (drValidaEquipos.Cells["Serial"].Value.ToString().Trim().Length == 0)
					{
						MessageBox.Show(string.Format("Ingrese la serial del articulo {0}", drValidaEquipos.Cells["Articulo"].Value.ToString().Trim()), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
						this.utcNomina.Tabs[2].Selected = true;
						this.utcVarios.Tabs[4].Selected = true;
						this.uGridEquipos.ActiveRow = drValidaEquipos;
						this.uGridEquipos.ActiveRow.Selected = true;
						return;
					}
					#endregion Serial

					#region Fecha de Entrega
					if (drValidaEquipos.Cells["FechaEntrega"].Value == System.DBNull.Value)
					{
						MessageBox.Show(string.Format("Ingrese la fecha de entrega del articulo", drValidaEquipos.Cells["Articulo"].Value.ToString().Trim()), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
						this.utcNomina.Tabs[2].Selected = true;
						this.utcVarios.Tabs[4].Selected = true;
						this.uGridEquipos.ActiveRow = drValidaEquipos;
						this.uGridEquipos.ActiveRow.Selected = true;
						return;
					}

					if (Convert.ToDateTime(drValidaEquipos.Cells["FechaEntrega"].Value) > dtValFecha)
					{
						MessageBox.Show(string.Format("La fecha de entrega del articulo no puede ser mayor a hoy", drValidaEquipos.Cells["Articulo"].Value.ToString().Trim()), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
						this.utcNomina.Tabs[2].Selected = true;
						this.utcVarios.Tabs[4].Selected = true;
						this.uGridEquipos.ActiveRow = drValidaEquipos;
						this.uGridEquipos.ActiveRow.Selected = true;
						return;
					}
					#endregion Fecha de Entrega

					#region Notas de entrega
					if (drValidaEquipos.Cells["ObservacionesEntrega"].Value.ToString().Trim().Length == 0)
					{
						MessageBox.Show(string.Format("Ingrese una nota de entrega del articulo {0}", drValidaEquipos.Cells["Articulo"].Value.ToString().Trim()), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
						this.utcNomina.Tabs[2].Selected = true;
						this.utcVarios.Tabs[4].Selected = true;
						this.uGridEquipos.ActiveRow = drValidaEquipos;
						this.uGridEquipos.ActiveRow.Selected = true;
						return;
					}
					#endregion Notas de entrega

					#region Fecha de Devolucion
					if (drValidaEquipos.Cells["FechaDevolucion"].Value != System.DBNull.Value)
					{
						if (Convert.ToDateTime(drValidaEquipos.Cells["FechaDevolucion"].Value) > dtValFecha)
						{
							MessageBox.Show(string.Format("La fecha de devolución del articulo no puede ser mayor a hoy", drValidaEquipos.Cells["Articulo"].Value.ToString().Trim()), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
							this.utcNomina.Tabs[2].Selected = true;
							this.utcVarios.Tabs[4].Selected = true;
							this.uGridEquipos.ActiveRow = drValidaEquipos;
							this.uGridEquipos.ActiveRow.Selected = true;
							return;
						}

						if (Convert.ToDateTime(drValidaEquipos.Cells["FechaDevolucion"].Value) < Convert.ToDateTime(drValidaEquipos.Cells["FechaEntrega"].Value))
						{
							MessageBox.Show(string.Format("La fecha de devolución del articulo no puede ser menor a la fecha de entrega del articulo", drValidaEquipos.Cells["Articulo"].Value.ToString().Trim()), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
							this.utcNomina.Tabs[2].Selected = true;
							this.utcVarios.Tabs[4].Selected = true;
							this.uGridEquipos.ActiveRow = drValidaEquipos;
							this.uGridEquipos.ActiveRow.Selected = true;
							return;
						}

						#region Notas de entrega
						if (drValidaEquipos.Cells["ObservacionesDevolucion"].Value.ToString().Trim().Length == 0)
						{
							MessageBox.Show(string.Format("Ingrese una nota de devolución del articulo {0}", drValidaEquipos.Cells["Articulo"].Value.ToString().Trim()), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
							this.utcNomina.Tabs[2].Selected = true;
							this.utcVarios.Tabs[4].Selected = true;
							this.uGridEquipos.ActiveRow = drValidaEquipos;
							this.uGridEquipos.ActiveRow.Selected = true;
							return;
						}
						#endregion Notas de entrega
					}
					#endregion Fecha de Devolucion
				}
			}
			#endregion Equipos
			
			#region Uniformes
			if (this.uGridUniformes.Rows.Count > 0)
			{			
				foreach(Infragistics.Win.UltraWinGrid.UltraGridRow drValidaUniformes in this.uGridUniformes.Rows.All)
				{
					#region Fecha de Entrega
					if (drValidaUniformes.Cells["FechaEntrega"].Value == System.DBNull.Value)
					{
						MessageBox.Show("Ingrese la fecha de entrega del uniforme", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
						this.utcNomina.Tabs[2].Selected = true;
						this.utcVarios.Tabs[5].Selected = true;
						this.uGridUniformes.ActiveRow = drValidaUniformes;
						this.uGridUniformes.ActiveRow.Selected = true;
						return;
					}
					#endregion Fecha de Entrega
					
					#region Detalle de Entrega
					if (drValidaUniformes.Cells["DetalleEntregado"].Value == System.DBNull.Value)
					{
						MessageBox.Show("Ingrese el detalle del uniforme entregado", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
						this.utcNomina.Tabs[2].Selected = true;
						this.utcVarios.Tabs[5].Selected = true;
						this.uGridUniformes.ActiveRow = drValidaUniformes;
						this.uGridUniformes.ActiveRow.Selected = true;
						return;
					}
					#endregion Detalle de Entrega

					#region Fecha de Devolucion
					if (drValidaUniformes.Cells["FechaDevolucion"].Value != System.DBNull.Value)
					{
						if (Convert.ToDateTime(drValidaUniformes.Cells["FechaEntrega"].Value) > Convert.ToDateTime(drValidaUniformes.Cells["FechaDevolucion"].Value))
						{
							MessageBox.Show("La fecha de entrega no puede ser mayor a la fecha de devolucion del uniforme", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
							this.utcNomina.Tabs[2].Selected = true;
							this.utcVarios.Tabs[5].Selected = true;
							this.uGridUniformes.ActiveRow = drValidaUniformes;
							this.uGridUniformes.ActiveRow.Selected = true;
							return;
						}

						if (drValidaUniformes.Cells["DetalleDevolucion"].Value == System.DBNull.Value)
						{
							MessageBox.Show("Ingrese el detalle del uniforme devuelto", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
							this.utcNomina.Tabs[2].Selected = true;
							this.utcVarios.Tabs[5].Selected = true;
							this.uGridUniformes.ActiveRow = drValidaUniformes;
							this.uGridUniformes.ActiveRow.Selected = true;
							return;
						}
					}					
					#endregion Fecha de Devolucion
				}
			}
			#endregion Uniformes

			#region Vacaciones - Permisos
			if (this.uGridVacaciones.Rows.Count > 0)
			{			
				foreach(Infragistics.Win.UltraWinGrid.UltraGridRow drValidaVacaciones in this.uGridVacaciones.Rows.All)
				{
					#region Tipo De Salida
					if (drValidaVacaciones.Cells["idNominaTiposDeSalidas"].Value == System.DBNull.Value)
					{
						MessageBox.Show("Seleccione un Tipo de Permiso", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
						this.utcNomina.Tabs[2].Selected = true;
						this.utcVarios.Tabs[6].Selected = true;
						this.uGridVacaciones.ActiveRow = drValidaVacaciones;
						this.uGridVacaciones.ActiveRow.Selected = true;
						return;
					}					
					#endregion Tipo De Salida

					#region Tipo De Parto
					if ((int)drValidaVacaciones.Cells["idNominaTiposDeSalidas"].Value == 4 || (int)drValidaVacaciones.Cells["idNominaTiposDeSalidas"].Value == 5)
					{
						int idNominaTiposDeParto = 0;
						if (drValidaVacaciones.Cells["idNominaTiposDeParto"].Value != System.DBNull.Value) idNominaTiposDeParto = (int)drValidaVacaciones.Cells["idNominaTiposDeParto"].Value;	
					
						if (idNominaTiposDeParto == 0)
						{
							MessageBox.Show("Seleccione un Tipo de Parto", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
							this.utcNomina.Tabs[2].Selected = true;
							this.utcVarios.Tabs[6].Selected = true;
							this.uGridVacaciones.ActiveRow = drValidaVacaciones;
							this.uGridVacaciones.ActiveRow.Selected = true;
							return;
						}
					}
					#endregion Tipo De Parto

					#region Desde
					if (drValidaVacaciones.Cells["Desde"].Value == System.DBNull.Value)
					{
						MessageBox.Show("Ingrese la fecha de Inicio de Permiso", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
						this.utcNomina.Tabs[2].Selected = true;
						this.utcVarios.Tabs[6].Selected = true;
						this.uGridVacaciones.ActiveRow = drValidaVacaciones;
						this.uGridVacaciones.ActiveRow.Selected = true;
						return;
					}
					#endregion Desde

					#region Hasta
					if (drValidaVacaciones.Cells["Hasta"].Value == System.DBNull.Value)
					{
						MessageBox.Show("Ingrese la fecha de termino de permiso", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
						this.utcNomina.Tabs[2].Selected = true;
						this.utcVarios.Tabs[6].Selected = true;
						this.uGridVacaciones.ActiveRow = drValidaVacaciones;
						this.uGridVacaciones.ActiveRow.Selected = true;
						return;
					}

					if (Convert.ToDateTime(drValidaVacaciones.Cells["Desde"].Value) > Convert.ToDateTime(drValidaVacaciones.Cells["Hasta"].Value))
					{
						MessageBox.Show("La fecha inicial no puede ser mayor a la fecha final del periodo de vacaciones", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
						this.utcNomina.Tabs[2].Selected = true;
						this.utcVarios.Tabs[6].Selected = true;
						this.uGridVacaciones.ActiveRow = drValidaVacaciones;
						this.uGridVacaciones.ActiveRow.Selected = true;
						return;
					}
					#endregion Hasta					

					#region Valida Permisos
//					iCont = 0;
//					string sSQLValida = string.Format("Exec Com_ValidaMovimientoPermisos {0}, '{1}', {2}", 
//						(int)this.txtIdPersonal.Value, Convert.ToDateTime(drValidaVacaciones.Cells["Desde"].Value).ToString("yyyyMMdd"), (int)drValidaVacaciones.Cells["idNominaSalidas"].Value);
//					iCont = Funcion.iEscalarSQL(cdsSeteoF, sSQLValida);
//				
//					if (iCont > 0)
//					{
//						MessageBox.Show("El empleado registra permiso (VACACIONES - MATERNIDAD - PATERNIDAD - ENFERMEDAD - FALTAS) en la fecha inicial seleccionada", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
//						this.utcNomina.Tabs[2].Selected = true;
//						this.utcVarios.Tabs[6].Selected = true;
//						this.uGridVacaciones.ActiveRow = drValidaVacaciones;
//						this.uGridVacaciones.ActiveRow.Selected = true;						
//						return;
//					}

//					iCont = 0;
//					sSQLValida = string.Format("Exec Com_ValidaMovimientoPermisos {0}, '{1}', {2}", 
//						(int)this.txtIdPersonal.Value, Convert.ToDateTime(drValidaVacaciones.Cells["Hasta"].Value).ToString("yyyyMMdd"), (int)drValidaVacaciones.Cells["idNominaSalidas"].Value);
//					iCont = Funcion.iEscalarSQL(cdsSeteoF, sSQLValida);
//				
//					if (iCont > 0)
//					{
//						MessageBox.Show("El empleado registra permiso (VACACIONES - MATERNIDAD - PATERNIDAD - ENFERMEDAD - FALTAS) en la fecha final seleccionada", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
//						this.utcNomina.Tabs[2].Selected = true;
//						this.utcVarios.Tabs[6].Selected = true;
//						this.uGridVacaciones.ActiveRow = drValidaVacaciones;
//						this.uGridVacaciones.ActiveRow.Selected = true;
//						return;
//					}

//					iCont = 0;
//					sSQLValida = string.Format("Exec Com_ValidaMovimientosVentas {0}, 0, '{1}', '{2}', 0", 
//						(int)this.txtIdPersonal.Value, Convert.ToDateTime(drValidaVacaciones.Cells["Desde"].Value).ToString("yyyyMMdd"), Convert.ToDateTime(drValidaVacaciones.Cells["Hasta"].Value).ToString("yyyyMMdd"));
//					iCont = Funcion.iEscalarSQL(cdsSeteoF, sSQLValida);
//
//					if (iCont > 0)
//					{
//						MessageBox.Show("El empleado registra Ventas en el rango de fechas seleccionado", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
//						this.utcNomina.Tabs[2].Selected = true;
//						this.utcVarios.Tabs[6].Selected = true;
//						this.uGridVacaciones.ActiveRow = drValidaVacaciones;
//						this.uGridVacaciones.ActiveRow.Selected = true;
//						return;
//					}
					#endregion Valida Permisos
				}
			}
			#endregion Vacaciones - Permisos		

			#region Faltas
			if (this.uGridFaltas.Rows.Count > 0)
			{			
				foreach(Infragistics.Win.UltraWinGrid.UltraGridRow drValidaFaltas in this.uGridFaltas.Rows.All)
				{
					#region Tipo De Salida
					if (drValidaFaltas.Cells["idNominaTiposDeSalidas"].Value == System.DBNull.Value)
					{
						MessageBox.Show("Seleccione un Tipo de Falta", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
						this.utcNomina.Tabs[2].Selected = true;
						this.utcVarios.Tabs[7].Selected = true;
						this.uGridFaltas.ActiveRow = drValidaFaltas;
						this.uGridFaltas.ActiveRow.Selected = true;
						return;
					}					
					#endregion Tipo De Salida
					
					#region Fecha
					if (drValidaFaltas.Cells["Fecha"].Value == System.DBNull.Value)
					{
						MessageBox.Show("Ingrese la fecha de falta", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
						this.utcNomina.Tabs[2].Selected = true;
						this.utcVarios.Tabs[7].Selected = true;
						this.uGridFaltas.ActiveRow = drValidaFaltas;
						this.uGridFaltas.ActiveRow.Selected = true;
						return;
					}

					if (Convert.ToDateTime(drValidaFaltas.Cells["Fecha"].Value) > DateTime.Today)
					{
						MessageBox.Show("La fecha de falta no puede ser mayor al dia de Hoy", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
						this.utcNomina.Tabs[2].Selected = true;
						this.utcVarios.Tabs[7].Selected = true;
						this.uGridFaltas.ActiveRow = drValidaFaltas;
						this.uGridFaltas.ActiveRow.Selected = true;
						return;
					}				

//					if ((int)drValidaFaltas.Cells["idNominaSalidas"].Value == 0)
//					{
//						string sSQL = string.Format("Exec ValidaFechaEnElMes '{0}'", Convert.ToDateTime(drValidaFaltas.Cells["Fecha"].Value).ToString("yyyyMMdd"));
//						if (!Funcion.bEscalarSQL(cdsSeteoF, sSQL, true))
//						{
//							MessageBox.Show("La fecha ingresada esta fuera del periodo", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
//							this.utcNomina.Tabs[2].Selected = true;
//							this.utcVarios.Tabs[7].Selected = true;
//							this.uGridFaltas.ActiveRow = drValidaFaltas;
//							this.uGridFaltas.ActiveRow.Selected = true;
//							return;
//						}
//					}	
					#endregion Fecha					

					#region Valida Permisos
//					iCont = 0;
//					string sSQLValida = string.Format("Exec Com_ValidaMovimientoPermisos {0}, '{1}', {2}", 
//						(int)this.txtIdPersonal.Value, Convert.ToDateTime(drValidaFaltas.Cells["Fecha"].Value).ToString("yyyyMMdd"), (int)drValidaFaltas.Cells["idNominaSalidas"].Value);
//					iCont = Funcion.iEscalarSQL(cdsSeteoF, sSQLValida);
//				
//					if (iCont > 0)
//					{
//						MessageBox.Show("El empleado registra permiso (VACACIONES - MATERNIDAD - PATERNIDAD - ENFERMEDAD - FALTAS) en la fecha seleccionada", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
//						this.utcNomina.Tabs[2].Selected = true;
//						this.utcVarios.Tabs[7].Selected = true;
//						this.uGridFaltas.ActiveRow = drValidaFaltas;
//						this.uGridFaltas.ActiveRow.Selected = true;						
//						return;
//					}
//					
//					iCont = 0;
//					sSQLValida = string.Format("Exec Com_ValidaMovimientosVentas {0}, 0, '{1}', '{2}', 0", 
//						(int)this.txtIdPersonal.Value, Convert.ToDateTime(drValidaFaltas.Cells["Fecha"].Value).ToString("yyyyMMdd"), Convert.ToDateTime(drValidaFaltas.Cells["Fecha"].Value).ToString("yyyyMMdd"));
//					iCont = Funcion.iEscalarSQL(cdsSeteoF, sSQLValida);
//
//					if (iCont > 0)
//					{
//						MessageBox.Show("El empleado registra Ventas en la fecha seleccionada", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
//						this.utcNomina.Tabs[2].Selected = true;
//						this.utcVarios.Tabs[7].Selected = true;
//						this.uGridFaltas.ActiveRow = drValidaFaltas;
//						this.uGridFaltas.ActiveRow.Selected = true;
//						return;
//					}
					#endregion Valida Permisos
				}
			}
			#endregion Faltas

			#region Otros
			if (this.uGridOtros.Rows.Count > 0)
			{			
				foreach(Infragistics.Win.UltraWinGrid.UltraGridRow drOtros in this.uGridOtros.Rows.All)
				{
					#region Motivo
					if (drOtros.Cells["idNominaMotivosOtros"].Value == System.DBNull.Value)
					{
						MessageBox.Show("Seleccione un Motivo", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
						this.utcNomina.Tabs[2].Selected = true;
						this.utcVarios.Tabs[8].Selected = true;
						this.uGridOtros.ActiveRow = drOtros;
						this.uGridOtros.ActiveRow.Selected = true;
						return;
					}					
					#endregion Motivo
					
					#region Fecha
					if (drOtros.Cells["Fecha"].Value == System.DBNull.Value)
					{
						MessageBox.Show("Ingrese la fecha de novedad", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
						this.utcNomina.Tabs[2].Selected = true;
						this.utcVarios.Tabs[8].Selected = true;
						this.uGridOtros.ActiveRow = drOtros;
						this.uGridOtros.ActiveRow.Selected = true;
						return;
					}

//					if ((int)drOtros.Cells["Estado"].Value == 0)
//					{
//						string sSQL = string.Format("Exec ValidaFechaEnElMes '{0}'", Convert.ToDateTime(drOtros.Cells["Fecha"].Value).ToString("yyyyMMdd"));
//						if (!Funcion.bEscalarSQL(cdsSeteoF, sSQL, true))
//						{
//							MessageBox.Show("La fecha ingresada esta fuera del periodo", "Point Tewchnology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
//							this.utcNomina.Tabs[2].Selected = true;
//							this.utcVarios.Tabs[8].Selected = true;
//							this.uGridOtros.ActiveRow = drOtros;
//							this.uGridOtros.ActiveRow.Selected = true;
//							return;
//						}
//					}
					#endregion Fecha

					#region Valor
					bool bValorFijo = false;
					decimal dValor = 0.00m;
					SqlDataReader drValores = Funcion.rEscalarSQL(cdsSeteoF, string.Format("Select ValorFijo, Valor From NominaMotivosOtros Where idNominaMotivosOtros = {0}", (int)drOtros.Cells["idNominaMotivosOtros"].Value), true);
					drValores.Read();
					bValorFijo = drValores.GetBoolean(0);
					dValor = drValores.GetDecimal(1);
					drValores.Close();

					if (Convert.ToDecimal(drOtros.Cells["Valor"].Value) != dValor && bValorFijo)
					{
						MessageBox.Show(string.Format("El valor debe ser {0}", dValor), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
						this.utcNomina.Tabs[2].Selected = true;
						this.utcVarios.Tabs[8].Selected = true;
						this.uGridOtros.ActiveRow = drOtros;
						this.uGridOtros.ActiveRow.Selected = true;
						return;
					}	
					#endregion Valor

					#region Observacion
					if (drOtros.Cells["Observacion"].Value.ToString().Length == 0)
					{
						MessageBox.Show("Ingrese una Observacion", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
						this.utcNomina.Tabs[2].Selected = true;
						this.utcVarios.Tabs[8].Selected = true;
						this.uGridOtros.ActiveRow = drOtros;
						this.uGridOtros.ActiveRow.Selected = true;
						return;
					}	
					#endregion Observacion
				}
			}
			#endregion Otros
			#endregion Grids

			#region Cambios
			bool bModifica = false;

			if (idCentroDeCostoA != (int)this.cmbNECentroCosto.Value) bModifica = true;
			if (idDepartamentoA != (int)this.cmbNEDepartamento.Value) bModifica = true;
			if (idCargoA != (int)this.cmbNECargo.Value) bModifica = true;
			if (SueldoA != Convert.ToDecimal(this.txtNESueldo.Value)) bModifica = true;
			if (BonoA != Convert.ToDecimal(this.txtNEBonos.Value)) bModifica = true;
			if (MovilizacionA != Convert.ToDecimal(this.txtNEMovilizacion.Value)) bModifica = true;
			#endregion cambios

			int IdRegion = 0;
			if ( this.cmbRegion.ActiveRow != null) IdRegion  = ( int) this.cmbRegion.Value;

			#endregion Validacion 
			
			#region Guarda Registro
			using (SqlConnection oConexion = new SqlConnection("Password=Bl45o6$9;Persist Security Info=True;User ID= " + MenuLatinium.stUsuario + ";Initial Catalog=" + MenuLatinium.stDirFacturacion + ";Data Source=" + MenuLatinium.stDirServidor + ";Max Pool Size = 1024;"))
			{
				oConexion.Open();

				SqlTransaction oTransaction = null;
				SqlCommand oCmdActualiza = oConexion.CreateCommand();
				oCmdActualiza.Connection = oConexion;
				oCmdActualiza.CommandType = CommandType.Text;
				oCmdActualiza.CommandTimeout = 120;
						
				try
				{
					oTransaction = oConexion.BeginTransaction();
					oCmdActualiza.Transaction = oTransaction;

					#region Valida Lugar de Nacimiento
					int IdProvNac = 0;
					int IdCantNac = 0;
					if ((int)this.cmbNENacionalidad.Value == 54)
					{
						IdProvNac = (int)this.cmbProvinciaNacimiento.Value;
						IdCantNac = (int)this.cmbCantonNacimiento.Value;
					}
					#endregion Valida Lugar de Nacimiento
						
					int idBanco = 0;
					if (this.cmbNEBancos.ActiveRow != null) idBanco = (int)this.cmbNEBancos.Value;
					int idCom_Rango = 0;
					if (this.cmbRangos.ActiveRow != null) idCom_Rango = (int)this.cmbRangos.Value;

					#region Guarda Datos Generales
					int rolSeleccionado = (int)this.optNERolesProservicios.Value;

					string esAdministrador = rolSeleccionado == 1 ? "1" : "NULL";
					string esRevisor = rolSeleccionado == 2 ? "1" : "NULL";
					string esCamarera = rolSeleccionado == 3 ? "1" : "NULL";

					// Si es "Ninguno", los tres deben ser NULL
					if (rolSeleccionado == 0)
					{
						esAdministrador = "NULL";
						esRevisor = "NULL";
						esCamarera = "NULL";
					}

					string sSQLGenerales = string.Format("Exec GuardaNomina {0}, {1}, {2}, '{3}', '{4}', '{5}', '{6}', {7}, {8}, '{9}', {10}, '{11}', {12}, {13}, '{14}', '{15}', {16}, {17}, {18}, {19}, {20}, {21}, '{22}', '{23}', '{24}', '{25}', {26}, {27}, {28}, {29}, {30}, {31}, {32}, {33}, {34}, '{35}', {36}, {37}, {38}, {39}, {40}, {41}, {42}, {43}, '{44}', {45}, {46}, {47}, {48}, {49}, {50}, {51}, {52}, {53}, '{54}', '{55}', '{56}', '{57}', {58}, {59}, '{60}', '{61}', '{62}', '{63}', '{64}', {65}, {66}, {67}, {68}, {69}, {70}, {71}, {72}, {73}, {74}, {75}, {76}, {77}, {78}, {79}, {80}, {81}",
						(int)this.txtIdNomina.Value, //0 @idNomina Int, 
						bEstado, //1 @Estado Bit, 
						(int)this.cmbNEIdentificacion.Value,//2 @idTipoRuc Int,
						this.txtNEIdentificacion.Text.ToString().Trim(), //3 @NIdentificacion Varchar(20),
						this.txtNEApellidoPaterno.Text.ToString().Trim(), //4 @ApellidoPaterno Varchar(100),
						this.txtNEApellidoMaterno.Text.ToString().Trim(), //5 @ApellidoMaterno Varchar(100), 
						this.txtCodigo.Text.ToString(), //6 @Codigo Varchar(20), 
						IdProvNac, //7 @idProvinciaNacimiento Int,
						IdCantNac, //8  @idCantonNacimiento Int, 
						Convert.ToDateTime(this.dtpNEFechaNacimiento.Value).ToString("yyyyMMdd"), //9  @FechaNacimiento Datetime, 
						(int)this.optNESexo.Value, //10 @Sexo Int, 
						this.txtNETipoDeSangre.Text.ToString(),//11 @TipoDeSangre Varchar(10),
						this.chkNEDiscapacidad.Checked, //12 @Discapacidad Bit, 
						(int)this.txtPorcentajeDiscapacidad.Value,//13 @PorcentajeDiscapacidad Int, 
						this.txtConadis.Text.ToString(), //14 @NumeroConadis Varchar(20), 
						this.txtTipoDiscapacidad.Text.ToString(), //15 @TipoDiscapacidad Varchar(100), 
						(int)this.cmbNEEstadoCivil.Value, //16 @idEdoCivil Int, 
						(int)this.txtNECargasFamiliares.Value,//17 @Cargas Int, 
						(int)this.cmbNEProvincia.Value, //18 @idProvincia Int, 
						(int)this.cmbNECanton.Value,//19 @idCanton Int, 
						(int)this.cmbNEParroquia.Value, //20 @idParroquia Int, 
						(int)this.cmbNEBarrio.Value, //21 @idBarrio Int, 
						this.txtNEDireccion.Text.ToString().Trim(), //22 @Direccion Varchar(200), 
						this.txtNETelefono.Text.ToString().Trim(), //23 @Telefono Varchar(20), 
						this.txtNECelular.Text.ToString().Trim(), //24 @Celular Varchar(20), 
						this.txtNEEMail.Text.ToString().Trim(), //25	@EMail Varchar(60),
						(int)this.optNEVivienda.Value, //26 @TipoVivienda Int, 
						(int)this.txtNETiempoDeVivienda.Value,//27 @TiempoDeVivienda Int, 
						(int)this.cmbNECentroCosto.Value, //28 @idProyecto Int, 
						(int)this.cmbNEDepartamento.Value, //29 @idDepartamento Int, 
						(int)this.cmbNECargo.Value, //30 @idCargo Int,
						(int)this.cmbNECodigoSectorial.Value, //31 @idSectorial Int,
						(int)this.optNESistema.Value,//32 @SSN Int,
						(int)this.cmbNETipoContrato.Value,//33 @idTipoContrato Int, 
						this.chkNEAporteIESS.Checked, //34 @AporteIESS Bit,
						this.txtAfiliacion.Text.ToString(), //35 @Afiliacion Varchar(50),
						this.chkFondoReserva.Checked, //36 @FondosDeReservaMensuales Bit,
						false,//37 @DecimosMensuales Bit, 
						Convert.ToDecimal(this.txtNESueldo.Value), //38 @Sueldo Decimal(18, 2),
						Convert.ToDecimal(this.txtNEMovilizacion.Value), //39  @Movilizacion Decimal(18, 2),
						Convert.ToDecimal(this.txtNEHorasExtras.Value), //40 @HorasExtras Decimal(18, 2),
						Convert.ToDecimal(this.txtTotalIngresos.Value), //41 @TotalIngresos Decimal(18, 2), 
						(int)this.optNETipoDePago.Value,//42 @TipoDePago Int,
						idBanco, //43 @idBanco Int,
						this.txtCuenta.Text.ToString().Trim(),//44 @NumeroDeCuenta Varchar(20), 
						this.chkNEImpuestoRenta.Checked, //45 @ImpuestoALaRenta Bit,
						Convert.ToDecimal(this.txtGastoVivienda.Value), //46 @GastosVivienda Decimal (18, 2),
						Convert.ToDecimal(this.txtGastoSalud.Value), //47 @GastosSalud Decimal (18, 2),
						Convert.ToDecimal(this.txtGastoAlimentacion.Value), //48 @GastosAlimentacion Decimal (18, 2),
						Convert.ToDecimal(this.txtGastoEducacion.Value), //49 @GastosEducacion Decimal (18, 2),
						Convert.ToDecimal(this.txtGastoVestimenta.Value), //50 @GastosVestimenta Decimal (18, 2), 
						this.chkNERetencionJudicial.Checked, //51 @RetencionJudicial Bit,
						Convert.ToDecimal(this.txtRetencionJudicial.Value), //52 @ValorRetencionJudicial Decimal(18, 2),
						this.chkNETCC.Checked,//53 @TCC Bit, 
						this.txtNEBono.Text.ToString().Trim(), //54 @Bono Varchar(1000),
						this.txtNEComisiones.Text.ToString().Trim(), //55 @Comisiones Varchar(1000),
						this.txtNEFunciones.Text.ToString().Trim(), //56 @Funciones Varchar(1000),
						this.txtNEResponsabilidades.Text.ToString().Trim(),//57 @Responsabilidades Varchar(1000), 
						Convert.ToDecimal(this.txtNEBonos.Value), //58 @Bonos Decimal(18, 2),
						Convert.ToDecimal(this.txtNEOrdinarias.Value), //59 @Ordinarias Decimal(18, 2),
						this.txtNEPrimerNombre.Text.ToString().Trim(), //60 @PrimerNombre Varchar(100),
						this.txtNESegundoNombre.Text.ToString().Trim(), //61 @SegundoNombre Varchar(100),
						this.txtNEEMailEmpresarial.Text.ToString().Trim(), //62 @EMailEmpresarial Varchar(60),
						this.txtPredial.Text.ToString().Trim(), //63 @Predial Varchar(50),
						this.txtNEVive_Con.Text.ToString(),//64 @Vive_Con Varchar(100)
						this.chkNEComisiona.Checked, // @Comisiona Bit, 
						this.chkNEVendedor.Checked, //@Vendedor Bit, 
						this.chkDerechoFondosDeReserva.Checked, //@DerechoFondosDeReserva Bit, 
						(int)this.optNEMesesAños.Value,
						(int)this.cmbNENacionalidad.Value,
						this.chkNERevisado.Checked, //@Revisado Bit
						this.chkNEAplicaRol.Checked,//@Pago en nomina Bit
						this.chkNEQuincena.Checked,
						Convert.ToDecimal(this.txtQuincena.Value),
						idCom_Rango, 
						(int)this.txtEstado.Value,
						IdRegion,
						this.chkDecTerAcum.Checked,
						this.chkDecCuartoAcum.Checked,
						esAdministrador,
						esRevisor,
						esCamarera
						); 
					oCmdActualiza.CommandText = sSQLGenerales;
					this.txtIdNomina.Value = (int)oCmdActualiza.ExecuteScalar();				
					#endregion Guarda Datos Generales
					
					#region Genera Codigo
					string sSQLCodigo = string.Format("Exec NominaGeneraCodigo {0}, '{1}', '{2}', '{3}', '{4}'",
						(int)this.txtIdNomina.Value, this.txtNEApellidoPaterno.Text.ToString(), this.txtNEApellidoMaterno.Text.ToString(), this.txtNEPrimerNombre.Text.ToString(), this.txtNESegundoNombre.Text.ToString());
					oCmdActualiza.CommandText = sSQLCodigo;
					this.txtCodigo.Text = oCmdActualiza.ExecuteScalar().ToString();
					#endregion Genera Codigo

					#region Grids
					#region Familiares
					foreach(Infragistics.Win.UltraWinGrid.UltraGridRow drFamiliares in this.uGridFamiliares.Rows.All)
					{	
						int idParentesco = 0;
						DateTime FechaNacimiento = DateTime.Today;

						if (drFamiliares.Cells["idParentesco"].Value != System.DBNull.Value) idParentesco = (int)drFamiliares.Cells["idParentesco"].Value;	
						if (drFamiliares.Cells["FechaNacimiento"].Value != System.DBNull.Value) FechaNacimiento = (DateTime)drFamiliares.Cells["FechaNacimiento"].Value;	

						string sSQLFamiliares = string.Format("Exec GuardaNominaFamiliares {0}, {1}, {2}, '{3}', '{4}', '{5}', '{6}', '{7}', {8}",
							(int)drFamiliares.Cells["idNominaFamiliar"].Value, (int)this.txtIdNomina.Value, idParentesco, drFamiliares.Cells["Nombre"].Value.ToString(),
							FechaNacimiento.ToString("yyyyMMdd"), drFamiliares.Cells["Telefono"].Value.ToString(), drFamiliares.Cells["Celular"].Value.ToString(),
							drFamiliares.Cells["LugarTrabajo"].Value.ToString(), drFamiliares.Cells["TiempoDeTrabajo"].Value.ToString());
						oCmdActualiza.CommandText = sSQLFamiliares;
						oCmdActualiza.ExecuteNonQuery();
					}
					#endregion Familiares

					#region Estudios
					foreach(Infragistics.Win.UltraWinGrid.UltraGridRow drEstudios in this.uGridEstudios.Rows.All)
					{						
						int idNivel = 0;
						int idEstadoEstudios = 0;
						int idPeriodosDeTiempo = 0;
						int iAñoFinalizacion = 0;

						if (drEstudios.Cells["idNivel"].Value != System.DBNull.Value) idNivel = (int)drEstudios.Cells["idNivel"].Value;							
						if (drEstudios.Cells["idPeriodosDeTiempo"].Value != System.DBNull.Value) idPeriodosDeTiempo = (int)drEstudios.Cells["idPeriodosDeTiempo"].Value;	
						if (drEstudios.Cells["idEstadoEstudios"].Value != System.DBNull.Value) idEstadoEstudios = (int)drEstudios.Cells["idEstadoEstudios"].Value;	
						if (drEstudios.Cells["AñoFinalizacion"].Value != System.DBNull.Value) iAñoFinalizacion = (int)drEstudios.Cells["AñoFinalizacion"].Value;

						string sSQLEstudios = string.Format("Exec GuardaNominaEstudios {0}, {1}, {2}, '{3}', '{4}', {5}, {6}, {7}, {8}",
							(int)drEstudios.Cells["idNominaEstudios"].Value, (int)this.txtIdNomina.Value, idNivel, drEstudios.Cells["Titulo"].Value.ToString(),
							drEstudios.Cells["Institucion"].Value.ToString(), iAñoFinalizacion, (int)drEstudios.Cells["Tiempo"].Value, idEstadoEstudios, idPeriodosDeTiempo);
						oCmdActualiza.CommandText = sSQLEstudios;
						oCmdActualiza.ExecuteNonQuery();
					}
					#endregion Estudios
					
					#region Experiencia
					foreach(Infragistics.Win.UltraWinGrid.UltraGridRow drExperiencia in this.uGridExperiencia.Rows.All)
					{							
						string sSQLExperiencia = string.Format("Exec GuardaNominaExperiencia {0}, {1}, '{2}', {3}, '{4}', '{5}', {6}, {7}, {8}, {9}, '{10}', {11}, {12}",
							(int)drExperiencia.Cells["idNominaExperiencia"].Value, (int)this.txtIdNomina.Value, 
							drExperiencia.Cells["Empresa"].Value.ToString(), (int)drExperiencia.Cells["PersonasACargo"].Value, 
							drExperiencia.Cells["Actividades"].Value.ToString(), drExperiencia.Cells["Logros"].Value.ToString(),
							(bool)drExperiencia.Cells["AfiliadoIESS"].Value, (int)drExperiencia.Cells["Aportaciones"].Value, 
							(bool)drExperiencia.Cells["SinProblemas"].Value, (bool)drExperiencia.Cells["ConProblemas"].Value, drExperiencia.Cells["DescripcionProblema"].Value.ToString(),
							(int)drExperiencia.Cells["Años"].Value, (int)drExperiencia.Cells["idPeriodosDeTiempo"].Value);
						oCmdActualiza.CommandText = sSQLExperiencia;
						oCmdActualiza.ExecuteNonQuery();
					}
					#endregion Experiencia

					#region Entradas
					foreach(Infragistics.Win.UltraWinGrid.UltraGridRow drEntradasSalidas in this.uGridEntradasSalidas.Rows.All)
					{	
						int idMotivoSalida = 0;
						if (drEntradasSalidas.Cells["idMotivoSalida"].Value != System.DBNull.Value) idMotivoSalida = (int)drEntradasSalidas.Cells["idMotivoSalida"].Value;	
						
						int idVerificadoPorE = 0;
						int idAutorizadoPorE = 0;

						if (drEntradasSalidas.Cells["idVerificadoPor"].Value != System.DBNull.Value) idVerificadoPorE = (int)drEntradasSalidas.Cells["idVerificadoPor"].Value;
						if (drEntradasSalidas.Cells["idAutorizadoPor"].Value != System.DBNull.Value) idAutorizadoPorE = (int)drEntradasSalidas.Cells["idAutorizadoPor"].Value;

						DateTime dtSalida = DateTime.Today;
						bool bFechaSalida = false;

						if (drEntradasSalidas.Cells["FechaSalida"].Value != System.DBNull.Value) 
						{
							dtSalida = Convert.ToDateTime(drEntradasSalidas.Cells["FechaSalida"].Value);
							bFechaSalida = true;
						}
						
						string sSQLEntradas = string.Format("Exec GuardaNominaEntrada {0}, {1}, '{2}', {3}, '{4}', '{5}', {6}, {7}, '{8}', '{9}', {10}, '{11}', {12}",
							(int)drEntradasSalidas.Cells["idNominaEntrada"].Value, (int)this.txtIdNomina.Value, 
							Convert.ToDateTime(drEntradasSalidas.Cells["FechaEntrada"].Value).ToString("yyyyMMdd"), 
							(bool)drEntradasSalidas.Cells["APrueba"].Value, Convert.ToDateTime(drEntradasSalidas.Cells["FechaTerminaPrueba"].Value).ToString("yyyyMMdd"),
							drEntradasSalidas.Cells["NEntrada"].Value.ToString(), 
							idVerificadoPorE, idAutorizadoPorE,
							dtSalida.ToString("yyyyMMdd"), drEntradasSalidas.Cells["NSalida"].Value.ToString(), idMotivoSalida, drEntradasSalidas.Cells["Descripcion"].Value.ToString(), bFechaSalida);
						oCmdActualiza.CommandText = sSQLEntradas;
						oCmdActualiza.ExecuteNonQuery();
					}
					#endregion Entradas

					#region Equipos
					foreach(Infragistics.Win.UltraWinGrid.UltraGridRow drEquipos in this.uGridEquipos.Rows.All)
					{
						DateTime dtDevolucion = DateTime.Today;
						bool bFechaDevolucion = false;

						if (drEquipos.Cells["FechaDevolucion"].Value != System.DBNull.Value) 
						{
							dtDevolucion = Convert.ToDateTime(drEquipos.Cells["FechaDevolucion"].Value);
							bFechaDevolucion = true;
						}

						string sSQLEquipos = string.Format("Exec GuardaNominaEquipos {0}, {1}, '{2}', '{3}', '{4}', '{5}', {6}, '{7}', '{8}'",
							(int)drEquipos.Cells["idNominaEquipos"].Value, (int)this.txtIdNomina.Value, 
							drEquipos.Cells["Articulo"].Value.ToString(), drEquipos.Cells["Serial"].Value.ToString(), 
							Convert.ToDateTime(drEquipos.Cells["FechaEntrega"].Value).ToString("yyyyMMdd"), dtDevolucion.ToString("yyyyMMdd"), bFechaDevolucion,
							drEquipos.Cells["ObservacionesEntrega"].Value.ToString(), drEquipos.Cells["ObservacionesDevolucion"].Value.ToString());
						oCmdActualiza.CommandText = sSQLEquipos;
						oCmdActualiza.ExecuteNonQuery();
					}
					#endregion Equipos

					#region Vacaciones - Permisos
					foreach(Infragistics.Win.UltraWinGrid.UltraGridRow drVacaciones in this.uGridVacaciones.Rows.All)
					{
						int idNominaTiposDeParto = 0;
						if (drVacaciones.Cells["idNominaTiposDeParto"].Value != System.DBNull.Value) idNominaTiposDeParto = (int)drVacaciones.Cells["idNominaTiposDeParto"].Value;	

						string sSQLVacaciones = string.Format("Exec NominaGuardaSalidas {0}, {1}, {2}, {3}, '{4}', '{5}', '{6}', ''",
							(int)drVacaciones.Cells["idNominaSalidas"].Value, (int)this.txtIdNomina.Value, 
							(int)drVacaciones.Cells["idNominaTiposDeSalidas"].Value, idNominaTiposDeParto, 
							Convert.ToDateTime(drVacaciones.Cells["Desde"].Value).ToString("yyyyMMdd"), Convert.ToDateTime(drVacaciones.Cells["Hasta"].Value).ToString("yyyyMMdd"),
							drVacaciones.Cells["Observaciones"].Value.ToString());
						oCmdActualiza.CommandText = sSQLVacaciones;
						oCmdActualiza.ExecuteNonQuery();						
					}
					#endregion Vacaciones - Permisos

					#region Uniformes
					foreach(Infragistics.Win.UltraWinGrid.UltraGridRow drUniformes in this.uGridUniformes.Rows.All)
					{
						DateTime dtFechaDev = DateTime.Today;
						bool bFD = false;

						if (drUniformes.Cells["FechaDevolucion"].Value != System.DBNull.Value) 
						{
							dtFechaDev = (DateTime)drUniformes.Cells["FechaDevolucion"].Value;
							bFD = true;
						}

						string sSQLUniformes = string.Format("Exec GuardaNominaUniformes {0}, {1}, '{2}', '{3}', '{4}', {5}, '{6}',{7},{8},{9}",
							(int)drUniformes.Cells["idNominaUniformes"].Value, 
							(int)this.txtIdNomina.Value, 
							Convert.ToDateTime(drUniformes.Cells["FechaEntrega"].Value).ToString("yyyyMMdd"), 
							drUniformes.Cells["DetalleEntregado"].Value.ToString(),
							dtFechaDev.ToString("yyyyMMdd"),
							bFD, 
							drUniformes.Cells["DetalleDevolucion"].Value.ToString(),
							(decimal)drUniformes.Cells["Precio"].Value,
							(int)drUniformes.Cells["Cuotas"].Value,
							(bool)drUniformes.Cells["AplicaRol"].Value);
						oCmdActualiza.CommandText = sSQLUniformes;
						oCmdActualiza.ExecuteNonQuery();
					}
					#endregion Uniformes

					#region Faltas
					if (this.uGridFaltas.Rows.Count > 0)
					{
						foreach(Infragistics.Win.UltraWinGrid.UltraGridRow drFaltas in this.uGridFaltas.Rows.All)
						{
							string sSQLFaltas = string.Format("Exec NominaGuardaSalidas {0}, {1}, {2}, {3}, '{4}', '{5}', '{6}', ''",
								(int)drFaltas.Cells["idNominaSalidas"].Value, (int)this.txtIdNomina.Value, 
								(int)drFaltas.Cells["idNominaTiposDeSalidas"].Value, 0, 
								Convert.ToDateTime(drFaltas.Cells["Fecha"].Value).ToString("yyyyMMdd"), Convert.ToDateTime(drFaltas.Cells["Fecha"].Value).ToString("yyyyMMdd"),
								drFaltas.Cells["Observaciones"].Value.ToString());
							oCmdActualiza.CommandText = sSQLFaltas;
							oCmdActualiza.ExecuteNonQuery();						
						}
					}
					#endregion Vacaciones - Permisos

					#region Otros
					if (this.uGridOtros.Rows.Count > 0)
					{
						foreach(Infragistics.Win.UltraWinGrid.UltraGridRow drOtros in this.uGridOtros.Rows.All)
						{
							string sSQLOtros = string.Format("Exec GuardaNominaOtros {0}, {1}, {2}, '{3}', {4}, '{5}'",
								(int)drOtros.Cells["idNominaOtros"].Value, (int)this.txtIdNomina.Value, (int)drOtros.Cells["idNominaMotivosOtros"].Value, 
								Convert.ToDateTime(drOtros.Cells["Fecha"].Value).ToString("yyyyMMdd"), Convert.ToDecimal(drOtros.Cells["Valor"].Value), drOtros.Cells["Observacion"].Value.ToString());
							oCmdActualiza.CommandText = sSQLOtros;
							oCmdActualiza.ExecuteNonQuery();
						}
					}
					#endregion Otros

					#region Cargas
					if (this.uGridCargas.Rows.Count > 0)
					{
						foreach(Infragistics.Win.UltraWinGrid.UltraGridRow drCargas in this.uGridCargas.Rows.All)
						{
							string sSQLCargas = string.Format("Exec NominaGuardaCargas {0}, {1}, '{2}', '{3}'",
								(int)drCargas.Cells["idNominaCargas"].Value, (int)this.txtIdNomina.Value, 
								drCargas.Cells["Nombre"].Value.ToString(), 
								Convert.ToDateTime(drCargas.Cells["FechaNacimiento"].Value).ToString("yyyyMMdd"));
							oCmdActualiza.CommandText = sSQLCargas;
							oCmdActualiza.ExecuteNonQuery();
						}
					}
					#endregion Cargas
					#endregion Grids
										
					string sSQLActPersonalNomina = string.Format("Exec ActualizaPersonalDesdeNomina {0}", (int)this.txtIdNomina.Value);
					oCmdActualiza.CommandText = sSQLActPersonalNomina;
					oCmdActualiza.ExecuteNonQuery();

     			#region Genera Nomina Cliente Greta y Pro
					if(MenuLatinium.stIdDB == 29 || MenuLatinium.stIdDB == 34 || MenuLatinium.stIdDB == 36)
					{
						if (bNuevo)
						{
							string sSQLNominaCliente = string.Format("Exec GuardaClienteNomina {0}, '{1}', '{2}', '{3}', '{4}', '{5}', '{6}', '{7}', '{8}', '{9}'", (int)this.cmbNEIdentificacion.Value, 
								this.txtNEIdentificacion.Text.ToString().Trim(), this.txtNEApellidoPaterno.Text.ToString().Trim(), this.txtNEApellidoMaterno.Text.ToString().Trim(), this.txtNEDireccion.Text.ToString().Trim(), 
								this.txtNETelefono.Text.ToString().Trim(), this.txtNECelular.Text.ToString().Trim(), this.txtNEEMail.Text.ToString().Trim(), this.txtNEPrimerNombre.Text.ToString().Trim(), 
								this.txtNESegundoNombre.Text.ToString().Trim());
							oCmdActualiza.CommandText = sSQLNominaCliente;
							oCmdActualiza.ExecuteNonQuery();
						}
					}
					#endregion Genera Nomina Cliente Greta y Pro

					//					if (idCom_Estado != (int)this.txtEstado.Value)//!bEstado)
					//					{
					if ((int)this.txtEstado.Value == 2)
					{
						string sSQLActSalidaAV = string.Format("Exec NominaSalidaPersonalActualizaAsignacionVendedores {0}, {1}", 
							(int)this.txtIdNomina.Value, (int)this.txtEstado.Value);
						oCmdActualiza.CommandText = sSQLActSalidaAV;
						oCmdActualiza.ExecuteNonQuery();	
					
						
						if (DialogResult.Yes == MessageBox.Show("Desea enviar Mail.", "Point Technology", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2)) 
						{
							string sSQLMailSalida = string.Format("Exec Com_MailSalidaPersonal {0}", (int)this.txtIdNomina.Value);
							oCmdActualiza.CommandText = sSQLMailSalida;
							oCmdActualiza.ExecuteNonQuery();		
						}
					}
					//					}

					oTransaction.Commit();	

					#region Controles
					FuncionesProcedimientos.EstadoControles(this, false, 3);
					FuncionesProcedimientos.EstadoControlesTabControl(this.utcNomina, false, 3);
					this.cmbRegion.Enabled = false;

					EstadoGrids(false);

					if (miAcceso.Nuevo) this.btnNuevo.Enabled = true;
					if (miAcceso.Buscar) this.btnConsultar.Enabled = true;
					if (miAcceso.Editar) this.btnEditar.Enabled = true;
					if (miAcceso.BReliquidar)this.btnAPP.Enabled = true;
					this.btnGuardar.Enabled = false;
					if (miAcceso.Imprimir) this.btnImprimir.Enabled = true;
					this.btnCancelar.Enabled = true;

					this.utcVarios.Tabs[6].Enabled = true;
					this.utcVarios.Tabs[7].Enabled = true;
					this.utcVarios.Tabs[8].Enabled = true;

					bNuevo = false;
					bEdicion = false;
				
					this.ConsultaTablas();
					#endregion Controles					
				}
				catch (Exception ex)
				{
					try
					{	
						oTransaction.Rollback();

						MessageBox.Show(string.Format("Commit Exception Type : {0} {1}", ex.GetType(), ex.Message), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Error);

						return;
					}
					catch (Exception ex2)
					{
						oTransaction.Rollback();

						MessageBox.Show(string.Format("Rollback Exception Type: {0} {1}", ex2.GetType(), ex2.Message), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Error);							

						return;
					}
				}
				finally
				{					
					oConexion.Close();					
				}
			}
			#endregion Guarda Registro			

			FuncionesProcedimientos.RetornaFechaServidor(this.dtFecha, cdsSeteoF, true);
		}
				
		private void uGridExperiencia_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			FuncionesProcedimientos.DesplazamientoGrids(sender, e, this.uGridExperiencia);
		}

		private void uGridEntradasSalidas_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			FuncionesProcedimientos.DesplazamientoGrids(sender, e, this.uGridEntradasSalidas);
		}

		private void uGridEquipos_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			FuncionesProcedimientos.DesplazamientoGrids(sender, e, this.uGridEquipos);
		}

		private void utcVarios_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			FuncionesProcedimientos.DesplazamientoGrids(sender, e, this.uGridVacaciones);
		}

		private void btnConsultar_Click(object sender, System.EventArgs e)
		{
			btnCancelar_Click(sender, e);

			using (frmConsultaNomina Busqueda = new frmConsultaNomina())
			{
				if (DialogResult.OK == Busqueda.ShowDialog())
				{
					this.Consulta((int)Busqueda.uGrid.ActiveRow.Cells["idNomina"].Value, false);
				}
			}

			FuncionesProcedimientos.RetornaFechaServidor(this.dtFecha, cdsSeteoF, true);
		}

		private void optNETipoDePago_ValueChanged(object sender, System.EventArgs e)
		{
			if (bEdicion)
			{
				if ((int)this.optNETipoDePago.Value != 4) this.txtCuenta.Enabled = true; 
				else 
				{
					this.cmbNEBancos.Value = System.DBNull.Value;
					this.cmbNEBancos.Enabled = false;
					this.txtCuenta.Enabled = false; 
					this.txtCuenta.Text = "";
				}
			}
		}

		private void frmNomina_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (!bEdicion)
			{
				if (e.KeyCode == Keys.N) if (miAcceso.Nuevo)	if (this.btnNuevo.Enabled) this.btnNuevo_Click(sender, e);
				if (e.KeyCode == Keys.B) if (miAcceso.Buscar) if (this.btnConsultar.Enabled) this.btnConsultar_Click(sender, e);
				if ((int)this.txtIdNomina.Value > 0)
				{
					if (e.KeyCode == Keys.E) if (miAcceso.Editar) if (this.btnEditar.Enabled) this.btnEditar_Click(sender, e);
					if (e.KeyCode == Keys.I) if (miAcceso.Imprimir) if (this.btnImprimir.Enabled) this.btnImprimir_Click(sender, e);																	 
					if (e.KeyCode == Keys.F5)
					{
						miAcceso = new Acceso(cdsSeteoF, "0501");

						this.Consulta((int)this.txtIdNomina.Value, false);					
					}
				}

				if (e.KeyCode == Keys.Escape) this.Close();				
			}
			else
			{
				if (e.KeyCode == Keys.Escape) 
				{
					if (DialogResult.Yes == MessageBox.Show("Esta seguro de Cancelar la Edición", "Point Technology", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2)) this.btnCancelar_Click(sender, e);
				}

				if (e.KeyCode == Keys.F12) this.btnGuardar_Click(sender, e);				
			}
		}

		private void btnImprimir_Click(object sender, System.EventArgs e)
		{
		
		}

		private void txtPorcentajeDiscapacidad_Enter(object sender, System.EventArgs e)
		{
			this.txtPorcentajeDiscapacidad.SelectAll();
		}

		private void txtPorcentajeDiscapacidad_Click(object sender, System.EventArgs e)
		{
			this.txtPorcentajeDiscapacidad.SelectAll();
		}

		private void txtPorcentajeDiscapacidad_Validated(object sender, System.EventArgs e)
		{
			if (this.txtPorcentajeDiscapacidad.Value == System.DBNull.Value) this.txtPorcentajeDiscapacidad.Value = 0;
		}

		private void txtNECargasFamiliares_Validated(object sender, System.EventArgs e)
		{
			if (this.txtNECargasFamiliares.Value == System.DBNull.Value) this.txtNECargasFamiliares.Value = 0;
		}

		private void txtNECargasFamiliares_Click(object sender, System.EventArgs e)
		{
			this.txtNECargasFamiliares.SelectAll();
		}

		private void txtNECargasFamiliares_Enter(object sender, System.EventArgs e)
		{
			this.txtNECargasFamiliares.SelectAll();
		}

		private void txtNESueldo_Click(object sender, System.EventArgs e)
		{
			this.txtNESueldo.SelectAll();
		}

		private void txtNESueldo_Enter(object sender, System.EventArgs e)
		{
			this.txtNESueldo.SelectAll();
		}

		private void txtNEMovilizacion_Click(object sender, System.EventArgs e)
		{
			this.txtNEMovilizacion.SelectAll();
		}

		private void txtNEMovilizacion_Enter(object sender, System.EventArgs e)
		{
			this.txtNEMovilizacion.SelectAll();
		}

		private void txtNEHorasExtras_Click(object sender, System.EventArgs e)
		{
			this.txtNEHorasExtras.SelectAll();
		}

		private void txtNEHorasExtras_Enter(object sender, System.EventArgs e)
		{
			this.txtNEHorasExtras.SelectAll();
		}

		private void txtNETotalIngresos_Click(object sender, System.EventArgs e)
		{
			this.txtTotalIngresos.SelectAll();
		}

		private void txtGastoVivienda_Click(object sender, System.EventArgs e)
		{
			this.txtGastoVivienda.SelectAll();
		}

		private void txtGastoVivienda_Enter(object sender, System.EventArgs e)
		{
			this.txtGastoVivienda.SelectAll();
		}

		private void txtGastoSalud_Click(object sender, System.EventArgs e)
		{
			this.txtGastoSalud.SelectAll();
		}

		private void txtGastoSalud_Enter(object sender, System.EventArgs e)
		{
			this.txtGastoSalud.SelectAll();
		}

		private void txtGastoAlimentacion_Click(object sender, System.EventArgs e)
		{
			this.txtGastoAlimentacion.SelectAll();
		}

		private void txtGastoAlimentacion_Enter(object sender, System.EventArgs e)
		{
			this.txtGastoAlimentacion.SelectAll();
		}

		private void txtGastoEducacion_Click(object sender, System.EventArgs e)
		{
			this.txtGastoEducacion.SelectAll();
		}

		private void txtGastoEducacion_Enter(object sender, System.EventArgs e)
		{
			this.txtGastoEducacion.SelectAll();
		}

		private void txtGastoVestimenta_Click(object sender, System.EventArgs e)
		{
			this.txtGastoVestimenta.SelectAll();
		}

		private void txtGastoVestimenta_Enter(object sender, System.EventArgs e)
		{
			this.txtGastoVestimenta.SelectAll();
		}

		private void txtRetencionJudicial_Click(object sender, System.EventArgs e)
		{
			this.txtRetencionJudicial.SelectAll();
		}

		private void txtRetencionJudicial_Enter(object sender, System.EventArgs e)
		{
			this.txtRetencionJudicial.SelectAll();
		}

		private void txtNESueldo_Validated(object sender, System.EventArgs e)
		{
			if (this.txtNESueldo.Value == System.DBNull.Value) this.txtNESueldo.Value = 0;		

			TotalIngresos();
		}

		private void txtNEMovilizacion_Validated(object sender, System.EventArgs e)
		{
			if (this.txtNEMovilizacion.Value == System.DBNull.Value) this.txtNEMovilizacion.Value = 0;		

			TotalIngresos();
		}

		private void txtNEHorasExtras_Validated(object sender, System.EventArgs e)
		{
			if (this.txtNEHorasExtras.Value == System.DBNull.Value) this.txtNEHorasExtras.Value = 0;		
		}

		private void txtNETotalIngresos_Validated(object sender, System.EventArgs e)
		{
			if (this.txtTotalIngresos.Value == System.DBNull.Value) this.txtTotalIngresos.Value = 0;
		}

		private void txtGastoVivienda_Validated(object sender, System.EventArgs e)
		{
			if (this.txtGastoVivienda.Value == System.DBNull.Value) this.txtGastoVivienda.Value = 0;
		}

		private void txtGastoSalud_Validated(object sender, System.EventArgs e)
		{
			if (this.txtGastoSalud.Value == System.DBNull.Value) this.txtGastoSalud.Value = 0;
		}

		private void txtGastoAlimentacion_Validated(object sender, System.EventArgs e)
		{
			if (this.txtGastoAlimentacion.Value == System.DBNull.Value) this.txtGastoAlimentacion.Value = 0;
		}

		private void txtGastoEducacion_Validated(object sender, System.EventArgs e)
		{
			if (this.txtGastoEducacion.Value == System.DBNull.Value) this.txtGastoEducacion.Value = 0;
		}

		private void txtGastoVestimenta_Validated(object sender, System.EventArgs e)
		{
			if (this.txtGastoVestimenta.Value == System.DBNull.Value) this.txtGastoVestimenta.Value = 0;
		}

		private void txtRetencionJudicial_Validated(object sender, System.EventArgs e)
		{
			if (this.txtRetencionJudicial.Value == System.DBNull.Value) this.txtRetencionJudicial.Value = 0;
		}

		private void uGridFamiliares_AfterRowInsert(object sender, Infragistics.Win.UltraWinGrid.RowEventArgs e)
		{
			if (bEdicion)
			{
				e.Row.Cells["idNominaFamiliar"].Value = 0;
				e.Row.Cells["idNomina"].Value = 0;
				e.Row.Cells["TiempoDeTrabajo"].Value = 0;
				e.Row.Cells["FechaNacimiento"].Value = DateTime.Today;
			}
		}
		
		private void uGridExperiencia_AfterRowInsert(object sender, Infragistics.Win.UltraWinGrid.RowEventArgs e)
		{
			if (bEdicion)
			{
				e.Row.Cells["idNominaExperiencia"].Value = 0;
				e.Row.Cells["idNomina"].Value = 0;				
				e.Row.Cells["PersonasACargo"].Value = 0;
				e.Row.Cells["Años"].Value = 0;				
				e.Row.Cells["idPeriodosDeTiempo"].Value = 0;
				e.Row.Cells["AfiliadoIESS"].Value = false;
				e.Row.Cells["Aportaciones"].Value = 0;
				e.Row.Cells["SinProblemas"].Value = false;
				e.Row.Cells["ConProblemas"].Value = false;				 
			}
		}

		private void uGridEntradasSalidas_AfterRowInsert(object sender, Infragistics.Win.UltraWinGrid.RowEventArgs e)
		{
			if (bEdicion)
			{
				e.Row.Cells["idNominaEntrada"].Value = 0;
				e.Row.Cells["idNomina"].Value = 0;
				e.Row.Cells["APrueba"].Value = false;
				e.Row.Cells["idMotivoSalida"].Value = 0;				
			}
		}

		private void uGridEquipos_AfterRowInsert(object sender, Infragistics.Win.UltraWinGrid.RowEventArgs e)
		{
			if (bEdicion)
			{
				e.Row.Cells["idNominaEquipos"].Value = 0;
				e.Row.Cells["idNomina"].Value = 0;				
			}
		}

		private void uGridVacaciones_AfterRowInsert(object sender, Infragistics.Win.UltraWinGrid.RowEventArgs e)
		{
			if (bEdicion)
			{
				e.Row.Cells["idNominaSalidas"].Value = 0;
				e.Row.Cells["idNomina"].Value = 0;
				e.Row.Cells["idNominaTiposDeSalidas"].Value = 0;
				e.Row.Cells["idNominaTiposDeParto"].Value = 0;
			}
		}

		private void txtNEOrdinarias_Click(object sender, System.EventArgs e)
		{
			this.txtNEOrdinarias.SelectAll();
		}

		private void txtNEOrdinarias_Enter(object sender, System.EventArgs e)
		{
			this.txtNEOrdinarias.SelectAll();
		}

		private void txtNEOrdinarias_Validated(object sender, System.EventArgs e)
		{
			if (this.txtNEOrdinarias.Value == System.DBNull.Value) this.txtNEOrdinarias.Value = 0;
		}

		private void txtNEOrdinarias_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Enter) this.txtNEHorasExtras.Focus();
		}

		private void txtNEBonos_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Enter) this.txtNEMovilizacion.Focus();
		}

		private void txtNEBonos_Validated(object sender, System.EventArgs e)
		{
			if (this.txtNEBonos.Value == System.DBNull.Value) this.txtNEBonos.Value = 0;		

			TotalIngresos();
		}

		private void txtNEBonos_Enter(object sender, System.EventArgs e)
		{
			this.txtNEBonos.SelectAll();
		}

		private void txtNEBonos_Click(object sender, System.EventArgs e)
		{
			this.txtNEBonos.SelectAll();
		}

		private void txtNEResponsabilidades_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Enter) this.chkNEVendedor.Focus();		
		}

		private void txtNETiempoDeVivienda_Enter(object sender, System.EventArgs e)
		{
			this.txtNETiempoDeVivienda.SelectAll();
		}

		private void txtNETiempoDeVivienda_Click(object sender, System.EventArgs e)
		{
			this.txtNETiempoDeVivienda.SelectAll();
		}

		private void txtNETiempoDeVivienda_Validated(object sender, System.EventArgs e)
		{
			if (this.txtNETiempoDeVivienda.Value == System.DBNull.Value) this.txtNETiempoDeVivienda.Value = 0;
		}

		private void txtNEApellidoPaterno_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Enter) this.txtNEApellidoMaterno.Focus();
		}

		private void txtNEApellidoMaterno_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Enter) this.txtNEPrimerNombre.Focus();
		}

		private void txtNESegundoNombre_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Enter) this.cmbProvinciaNacimiento.Focus();
		}

		private void txtNEPrimerNombre_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Enter) this.txtNESegundoNombre.Focus();
		}

		private void txtNEEMailEmpresarial_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Enter) this.optNEVivienda.Focus();
		}

		private void txtNEApellidoPaterno_ValueChanged(object sender, System.EventArgs e)
		{
		
		}

		private void cmbNEBancos_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Enter) this.txtCuenta.Focus();
		}

		private void txtNETiempoDeVivienda_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Enter) this.optNEMesesAños.Focus();
		}

		private void txtNEVive_Con_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Enter) 
			{
				if (this.txtPredial.Enabled) this.txtPredial.Focus(); 
				else 
				{
					this.uGridFamiliares.DisplayLayout.Bands[0].AddNew();
					this.uGridFamiliares.ActiveCell = this.uGridFamiliares.Rows[0].Cells["idParentesco"];
				}
			}
		}

		private void optNEVivienda_ValueChanged(object sender, System.EventArgs e)
		{
			if (bEdicion)
			{
				if ((int)this.optNEVivienda.Value == 1) this.txtPredial.Enabled = true; 
				else 
				{
					this.txtPredial.Enabled = false; 
					this.txtPredial.Text = "";
				}
			}
		}

		private void txtPredial_ValueChanged(object sender, System.EventArgs e)
		{

		}
		
		private void txtNEIdentificacion_ValueChanged(object sender, System.EventArgs e)
		{
		
		}

		private void txtNEIdentificacion_Validated(object sender, System.EventArgs e)
		{
//			pbFoto.Image = Image.FromFile(@"C:/Latinium/Personal/" + this.txtNEIdentificacion.Text.ToString() + "/Foto.jpg");
//			pbCedula.Image = Image.FromFile(@"C:/Latinium/Personal/" + this.txtNEIdentificacion.Text.ToString() + "/Cedula.jpg");
//			pbCroquis.Image = Image.FromFile(@"C:/Latinium/Personal/" + this.txtNEIdentificacion.Text.ToString() + "/Croquis.jpg");			

			if (bEdicion && bNuevo)
			{
				int idNomina = Funcion.iEscalarSQL(cdsSeteoF, string.Format("If Exists(Select idNomina From Nomina Where NIdentificacion = '{0}') Select idNomina From Nomina Where NIdentificacion = '{0}' Else Select CONVERT(Int, 0)", this.txtNEIdentificacion.Text.ToString()));
				if (idNomina > 0)
				{
					this.Consulta(idNomina, true);
					this.bNuevo = false;
					this.bEdicion = false;


				}
			}
		}

		private void chkNEDiscapacidad_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Enter) 
			{
				if (this.chkNEDiscapacidad.Checked) this.txtPorcentajeDiscapacidad.Focus(); else this.chkNELicencia.Focus();
			}
		}

		private void chkNEAporteIESS_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Enter) 
			{
				if (this.chkNEAporteIESS.Checked) this.chkFondoReserva.Focus(); else this.txtNESueldo.Focus();
			}
		}

		private void chkFondoReserva_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Enter) this.chkDecTerAcum.Focus();
		}

		private void chkDecimos_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Enter) this.txtAfiliacion.Focus();
		}

		private void chkNEImpuestoRenta_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Enter) 
			{
				if (this.chkNEImpuestoRenta.Checked) this.txtGastoVivienda.Focus(); else this.chkNERetencionJudicial.Focus();
			}
		}

		private void chkNERetencionJudicial_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Enter) 
			{
				if (this.chkNERetencionJudicial.Checked) this.txtRetencionJudicial.Focus(); 
				else if (MenuLatinium.bComisionQuincena) this.chkNEQuincena.Focus(); 
				else this.optNETipoDePago.Focus();
			}
		}

		private void cmbNEProvinciaNacimiento_InitializeLayout(object sender, Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs e)
		{
		
		}

		private void cmbNEProvincia_InitializeLayout(object sender, Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs e)
		{
		
		}

		private void cmbNECanton_InitializeLayout(object sender, Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs e)
		{
		
		}

		private void cmbNEParroquia_InitializeLayout(object sender, Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs e)
		{
		
		}
		
		private void uGridFamiliares_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
			if (this.uGridFamiliares.ActiveCell.Column.Key.ToString().ToUpper() == "TELEFONO" || this.uGridFamiliares.ActiveCell.Column.Key.ToString().ToUpper() == "CELULAR")
				Funcion.CamposNumericos(sender, e);
		}
		
		private void uGridVacaciones_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			FuncionesProcedimientos.DesplazamientoGrids(sender, e, this.uGridVacaciones);
		}
		
		private void frmNomina_Closing(object sender, System.ComponentModel.CancelEventArgs e)
		{
			if (bEdicion)
			{
				MessageBox.Show("No puede salir mientras edita un Empleado", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				e.Cancel = true;
				return;
			}
		}

		private void txtNEPrimerNombre_ValueChanged(object sender, System.EventArgs e)
		{
			
		}

		private void txtNEApellidoPaterno_Validated(object sender, System.EventArgs e)
		{
//			if (bEdicion)
//			{
//				if (this.txtNEPrimerNombre.Text.Length > 0)	this.txtCodigo.Text = this.txtNEPrimerNombre.Text.ToString().Trim().Substring(0, 1) +  this.txtNEApellidoPaterno.Text.ToString().Trim().Replace(" ", "");
//			}
		}

		private void txtNEPrimerNombre_Validated(object sender, System.EventArgs e)
		{
//			if (bEdicion)
//			{
//				this.txtCodigo.Text = this.txtNEPrimerNombre.Text.ToString().Trim().Substring(0, 1) +  this.txtNEApellidoPaterno.Text.ToString().Trim().Replace(" ", "");
//			}
		}

		private void uGridEstudios_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			FuncionesProcedimientos.DesplazamientoGrids(sender, e, this.uGridEstudios);
		}

		private void uGridEstudios_AfterRowInsert(object sender, Infragistics.Win.UltraWinGrid.RowEventArgs e)
		{
			if (bEdicion)
			{
				e.Row.Cells["idNominaEstudios"].Value = 0;
				e.Row.Cells["idNomina"].Value = 0;
				e.Row.Cells["idNivel"].Value = 0;
				e.Row.Cells["Tiempo"].Value = 0;
				e.Row.Cells["AñoFinalizacion"].Value = 0;
				e.Row.Cells["idEstadoEstudios"].Value = 0;
			}
		}

		private void txtTipoDiscapacidad_ValueChanged(object sender, System.EventArgs e)
		{
		
		}

		private void chkNELicencia_CheckedChanged(object sender, System.EventArgs e)
		{
			if (bEdicion)
			{
				this.txtTipoDeLicencia.Enabled = this.chkNELicencia.Checked;
				this.dtpFechaCaducidad.Enabled = this.chkNELicencia.Checked;				

				if (!this.chkNELicencia.Checked)
				{
					this.txtTipoDeLicencia.Text = "";
					this.dtpFechaCaducidad.Value = System.DBNull.Value;
				}				
			}
		}

		private void chkNELicencia_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Enter) 
			{
				if (this.chkNELicencia.Checked) this.txtTipoDeLicencia.Focus(); else this.cmbNEProvincia.Focus();
			}
		}

		private void txtNEVive_Con_ValueChanged(object sender, System.EventArgs e)
		{
		
		}

		private void txtPredial_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Enter) 
			{
				this.uGridFamiliares.DisplayLayout.Bands[0].AddNew();
				this.uGridFamiliares.ActiveCell = this.uGridFamiliares.Rows[0].Cells["idParentesco"];			
			}
		}

		private void uGridExperiencia_AfterCellUpdate(object sender, Infragistics.Win.UltraWinGrid.CellEventArgs e)
		{
			if (e.Cell.Column.ToString() == "PersonasACargo")
			{
				if (e.Cell.Row.Cells["PersonasACargo"].Value == System.DBNull.Value) e.Cell.Row.Cells["PersonasACargo"].Value = 0;
			}

			if (e.Cell.Column.ToString() == "Aportaciones")
			{
				if (e.Cell.Row.Cells["Aportaciones"].Value == System.DBNull.Value) e.Cell.Row.Cells["Aportaciones"].Value = 0;
			}
		}

		private void uGridEntradasSalidas_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
			if (this.uGridEntradasSalidas.ActiveCell.Column.Key.ToString().ToUpper() == "NENTRADA" || this.uGridEntradasSalidas.ActiveCell.Column.Key.ToString().ToUpper() == "NSALIDA")
				Funcion.CamposNumericos(sender, e);
		}

		private void uGridEstudios_AfterCellUpdate(object sender, Infragistics.Win.UltraWinGrid.CellEventArgs e)
		{
			if (e.Cell.Column.ToString() == "Tiempo")
			{
				if (e.Cell.Row.Cells["Tiempo"].Value == System.DBNull.Value) e.Cell.Row.Cells["Tiempo"].Value = 0;
			}

			if (e.Cell.Column.ToString() == "AñoFinalizacion")
			{
				if (e.Cell.Row.Cells["AñoFinalizacion"].Value == System.DBNull.Value) e.Cell.Row.Cells["AñoFinalizacion"].Value = 0;
			}
		}

		private void uGridVacaciones_AfterCellUpdate(object sender, Infragistics.Win.UltraWinGrid.CellEventArgs e)
		{
//			if (e.Cell.Column.ToString() == "Desde")
//			{
//				if (e.Cell.Row.Cells["Desde"].Value != System.DBNull.Value && e.Cell.Row.Cells["Hasta"].Value != System.DBNull.Value) 
//				{
////					e.Cell.Row.Cells["Dias"].Value = Funcion.iEscalarSQL(cdsSeteoF, string.Format("Select DATEDIFF(DAY, '{0}', '{1}')", 
////						Convert.ToDateTime(e.Cell.Row.Cells["Desde"].Value).ToString("yyyyMMdd"), Convert.ToDateTime(e.Cell.Row.Cells["Hasta"].Value).ToString("yyyyMMdd")));
//
//					//e.Cell.Row.Cells["Dias"].Value = (int)e.Cell.Row.Cells["Dias"].Value + 1;
//				}
//			}
//
//			if (e.Cell.Column.ToString() == "Hasta")
//			{
//				if (e.Cell.Row.Cells["Desde"].Value != System.DBNull.Value && e.Cell.Row.Cells["Hasta"].Value != System.DBNull.Value) 
//				{
////					e.Cell.Row.Cells["Dias"].Value = Funcion.iEscalarSQL(cdsSeteoF, string.Format("Select DATEDIFF(DAY, '{0}', '{1}')", 
////						Convert.ToDateTime(e.Cell.Row.Cells["Desde"].Value).ToString("yyyyMMdd"), Convert.ToDateTime(e.Cell.Row.Cells["Hasta"].Value).ToString("yyyyMMdd")));
//
//					//e.Cell.Row.Cells["Dias"].Value = (int)e.Cell.Row.Cells["Dias"].Value + 1;
//				}
//			}			
		}

		private void txtNETiempoDeVivienda_ValueChanged(object sender, System.EventArgs e)
		{
		
		}

		private void optNEMesesAños_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Enter) this.txtNEVive_Con.Focus();
		}

		private void txtRetencionJudicial_ValueChanged(object sender, System.EventArgs e)
		{
		
		}

		private void txtNEResponsabilidades_ValueChanged(object sender, System.EventArgs e)
		{
		
		}

		private void chkNETCC_CheckedChanged(object sender, System.EventArgs e)
		{
		
		}

		private void chkNEVendedor_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Enter) this.chkNEComisiona.Focus();
		}

		private void chkNEComisiona_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Enter) this.chkDerechoFondosDeReserva.Focus();
		}

		private void uGridUniformes_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			FuncionesProcedimientos.DesplazamientoGrids(sender, e, this.uGridUniformes);
		}

		private void optNEMesesAños_ValueChanged(object sender, System.EventArgs e)
		{
		
		}

		private void optNESexo_ValueChanged(object sender, System.EventArgs e)
		{
		
		}

		private void optNESistema_ValueChanged(object sender, System.EventArgs e)
		{
		
		}		

		private void uGridUniformes_AfterRowInsert(object sender, Infragistics.Win.UltraWinGrid.RowEventArgs e)
		{
			if (bEdicion)
			{
				e.Row.Cells["idNominaUniformes"].Value = 0;
				e.Row.Cells["Precio"].Value = 0.00m;	
				e.Row.Cells["Cuotas"].Value = 0;
				e.Row.Cells["AplicaRol"].Value = false;
			}
		}

		private void cmbNENacionalidad_InitializeLayout(object sender, Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs e)
		{
		
		}

		private void cmbNENacionalidad_ValueChanged(object sender, System.EventArgs e)
		{
			if (bEdicion)
			{
				if ((int)this.cmbNENacionalidad.Value != 54)
				{
					this.cmbProvinciaNacimiento.Enabled = false;
					this.cmbCantonNacimiento.Enabled = false;

					this.cmbProvinciaNacimiento.Value = System.DBNull.Value;
					this.cmbCantonNacimiento.Value = System.DBNull.Value;
				}
				else
				{
					this.cmbProvinciaNacimiento.Enabled = true;
					this.cmbCantonNacimiento.Enabled = true;
				}
			}
		}

		private void chkNEAplicaRol_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Enter) this.btnGuardar.Focus();
		}

		private void chkNEQuincena_CheckedChanged(object sender, System.EventArgs e)
		{
			if (bEdicion)
			{
				this.txtQuincena.Enabled = this.chkNEQuincena.Checked;

				if (!this.chkNEQuincena.Checked) this.txtQuincena.Value = 0;
			}
		}

		private void uGridVacaciones_InitializeLayout(object sender, Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs e)
		{
		
		}

		private void ultraCombo1_InitializeLayout(object sender, Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs e)
		{
			
		}

		private void cmbNEDepartamento_ValueChanged(object sender, System.EventArgs e)
		{
			if (bEdicion)
			{
				if ((int)this.cmbNEDepartamento.ActiveRow.Cells["idModulo"].Value != 1)
				{
					this.cmbNECargo.Value = System.DBNull.Value;
					this.cmbRangos.Value = System.DBNull.Value;
					this.cmbRangos.Enabled = false;
				}
			}
		}

		private void cmbNECodigoSectorial_InitializeLayout(object sender, Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs e)
		{
		}

		private void cmbRangos_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Enter) this.optNESistema.Focus();
		}

		private void txtQuincena_ValueChanged(object sender, System.EventArgs e)
		{
		}

		private void txtQuincena_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Enter) this.optNETipoDePago.Focus();
		}

		private void chkNEQuincena_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Enter) 
			{
				if (this.chkNEQuincena.Checked) this.txtQuincena.Focus(); 				
				else this.optNETipoDePago.Focus();
			}
		}

		private void uGridEntradasSalidas_InitializeLayout(object sender, Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs e)
		{
		
		}
    	bool bActualiza = false;
		private void uGridEntradasSalidas_AfterCellUpdate(object sender, Infragistics.Win.UltraWinGrid.CellEventArgs e)
		{
			#region Variables
			int iMesA = Convert.ToInt32( ( Convert.ToDateTime(this.dtFecha.Value).Month));
			int iYearA = Convert.ToInt32( ( Convert.ToDateTime(this.dtFecha.Value).Year));
			DateTime dtInicio = DateTime.Parse("01/" + iMesA + "/" + iYearA);
			DateTime dComisiones ;

			DateTime dtFinalPeriodo = DateTime.Parse("28/" + iMesA + "/" + iYearA); //29

			int iContR = this.uGridEntradasSalidas.Rows.Count;
			//			DateTime dtFinalPeriodo = DateTime.Parse(DateTime.DaysInMonth(dtInicio.Date.Year, dtInicio.Month).ToString() + "/" + dtInicio.Month.ToString() + "/" + dtInicio.Year.ToString());
			//      int iUdia = Convert.ToInt32( ( Convert.ToDateTime(dtFinalPeriodo).Day));
			//      int iUdiaRes = iUdia - 4;
			#endregion Variables
			if (bActualiza) return;

			#region valida qu no mofifiquen ingresos o salidas ya ingresadas menor al periodo actual
			if (!miAcceso.BLimiteBusquedaFechas) 
			{
				if (e.Cell.Column.ToString() == "FechaEntrada" )
				{
					if ((int)e.Cell.Row.Cells["idNominaEntrada"].Value  > 0)
					{
					
						if (e.Cell.Row.Cells["FechaEntrada"].Value != System.DBNull.Value && e.Cell.Row.Cells["FechaSalida"].Value != System.DBNull.Value  && Convert.ToDateTime (e.Cell.Row.Cells["FechaSalida"].Value) < dtInicio && (int)e.Cell.Row.Cells["idMotivoSalida"].Value >0  )
						{
							bActualiza = true;
							e.Cell.Row.Cells["FechaEntrada"].Value = (DateTime)e.Cell.OriginalValue;	
							bActualiza = false;
							return;
						}

					}
				}

			if (e.Cell.Column.ToString() == "FechaSalida" )
			{
				if ((int)e.Cell.Row.Cells["idNominaEntrada"].Value  > 0)
				{
					if (e.Cell.Row.Cells["FechaEntrada"].Value != System.DBNull.Value && e.Cell.Row.Cells["FechaSalida"].Value != System.DBNull.Value  && Convert.ToDateTime (e.Cell.Row.Cells["FechaSalida"].Value) < dtInicio && (int)e.Cell.Row.Cells["idMotivoSalida"].Value > 0  )
					{
						bActualiza = true;
						e.Cell.Row.Cells["FechaSalida"].Value = (DateTime)e.Cell.OriginalValue;
						bActualiza = false;
						return;
					}

				}
			}
		}
       	#endregion valida qu no mofifiquen ingresos o salidas ya ingresadas menor al periodo actual

			if (e.Cell.Column.ToString() == "FechaEntrada")
			{
				if (!miAcceso.BLimiteBusquedaFechas) 
				{
					if (e.Cell.Row.Cells["FechaEntrada"].Value != System.DBNull.Value)
					{
						bActualiza = true;
					
						if (Convert.ToDateTime (e.Cell.Row.Cells["FechaEntrada"].Value) <  dtInicio)
						{
							MessageBox.Show("La Fecha De Entrada no puede ser menor al periodo actual", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);	
							e.Cell.Row.Cells["FechaEntrada"].Value = (DateTime)e.Cell.OriginalValue;
							bActualiza = false;
							return;
						}
						/*poner seguridad*/
						if (Convert.ToDateTime (this.dtFecha.Value) >  dtFinalPeriodo)
						{
							if (e.Cell.Row.Cells["idNominaEntrada"].Value != System.DBNull.Value)
							{
								
								if (iMesA == 12) 
								{
									iYearA = iYearA + 1; 
									iMesA = 1;
								}
								if((int) e.Cell.Row.Cells["idNominaEntrada"].Value == 0 )
								{
									MessageBox.Show("Los Ingresos son unicamente hasta el 28 de cada mes ", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);	
									e.Cell.Row.Cells["FechaEntrada"].Value = DateTime.Parse("01/" + (iMesA + 1) + "/" + iYearA);
									e.Cell.Row.Cells["FechaTerminaPrueba"].Value = Convert.ToDateTime(e.Cell.Row.Cells["FechaEntrada"].Value).AddDays(90);
									bActualiza = false;
									return;
								}
								else
								{
									MessageBox.Show("La Edición es unicamente hasta 28 de cada mes ", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);	
									e.Cell.Row.Cells["FechaEntrada"].Value = (DateTime)e.Cell.OriginalValue;
									e.Cell.Row.Cells["FechaTerminaPrueba"].Value = (DateTime)e.Cell.OriginalValue;
									bActualiza = false;
									return;
								}
							}
							/*fin seguridad*/
						}
						e.Cell.Row.Cells["FechaTerminaPrueba"].Value = Convert.ToDateTime(e.Cell.Row.Cells["FechaEntrada"].Value).AddDays(90);
					}
					bActualiza = false;
				}
			}
			if (e.Cell.Column.ToString() == "FechaSalida")
			{
				if (!miAcceso.BLimiteBusquedaFechas) 
				{
					
					if (e.Cell.Row.Cells["FechaSalida"].Value != System.DBNull.Value)
					{
						bActualiza = true;
					
						if (Convert.ToDateTime (e.Cell.Row.Cells["FechaSalida"].Value) <  dtInicio)
						{
							MessageBox.Show("La Fecha De Salida no puede ser menor al periodo actual", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);	
							e.Cell.Row.Cells["FechaSalida"].Value = "";
							bActualiza = false;
							return;
						}
						/*poner seguridad*/
						if (Convert.ToDateTime (this.dtFecha.Value) >  dtFinalPeriodo)
						{
							if (e.Cell.Row.Cells["idNominaEntrada"].Value != System.DBNull.Value)
							{
								
								if (iMesA == 12) 
								{
									iYearA = iYearA + 1; 
									iMesA = 1;
								}
								if((int) e.Cell.Row.Cells["idNominaEntrada"].Value == 0 )
								{
									MessageBox.Show("Las Salidas son unicamente hasta el 28 de cada mes ", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);	
									e.Cell.Row.Cells["FechaSalida"].Value = DateTime.Parse("01/" + (iMesA + 1) + "/" + iYearA);
									//									e.Cell.Row.Cells["FechaTerminaPrueba"].Value = Convert.ToDateTime(e.Cell.Row.Cells["FechaEntrada"].Value).AddDays(90);
									bActualiza = false;
									return;
								}
								else
								{
									MessageBox.Show("La Edición es unicamente hasta 28 de cada mes ", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);	
									e.Cell.Row.Cells["FechaSalida"].Value = DateTime.Parse("01/" + (iMesA + 1) + "/" + iYearA);
									//									e.Cell.Row.Cells["FechaTerminaPrueba"].Value = (DateTime)e.Cell.OriginalValue;
									bActualiza = false;
									return;
								}
							}
							/*fin seguridad*/
						}
					}
//							e.Cell.Row.Cells["FechaTerminaPrueba"].Value = Convert.ToDateTime(e.Cell.Row.Cells["FechaEntrada"].Value).AddDays(90);
					bActualiza = false;
				}
			}

			if(MenuLatinium.stDirFacturacion != "POINT" && MenuLatinium.stDirFacturacion != "CREDISOLUTIONS")
			{
				if (e.Cell.Row.Cells["FechaEntrada"].Value != System.DBNull.Value)
				{
					e.Cell.Row.Cells["FechaTerminaPrueba"].Value = Convert.ToDateTime(e.Cell.Row.Cells["FechaEntrada"].Value).AddDays(90);
				}
			}
			else 
			{
//			if (e.Cell.Row.Cells["FechaEntrada"].Value != System.DBNull.Value)
//			{
//				e.Cell.Row.Cells["FechaTerminaPrueba"].Value = Convert.ToDateTime(e.Cell.Row.Cells["FechaEntrada"].Value).AddDays(90);
//			}
			}

			/* validacion para que el personal no salga si ya esta en comisiones*/
			if (e.Cell.Column.ToString() == "FechaSalida")
			{
				if (!miAcceso.BLimiteBusquedaFechas) 
				{
					if (e.Cell.Row.Cells["FechaSalida"].Value != System.DBNull.Value)
					{
						bActualiza = true;
						dComisiones = Funcion.dtEscalarSQL(cdsSeteoF, string.Format("SELECT TOP 1  CONVERT (DATE,c.FechaSistema) FROM Comisiones c INNER JOIN  DetalleComisiones D ON d.idComision = c.idComision WHERE Estado = 2 and idNomina = '{0}'  ORDER BY c.idComision desc", Convert.ToInt32( e.Cell.Row.Cells["idNomina"].Value)));	
						if (Convert.ToDateTime (e.Cell.Row.Cells["FechaSalida"].Value) <  dComisiones)
						{
							MessageBox.Show(string.Format("La Fecha De Salida no puede ser menor al '{0}' lineas", dComisiones), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
							e.Cell.Row.Cells["FechaSalida"].Value = "";
							bActualiza = false;
							return;
						}
					}
					bActualiza = false;
				}
			}
		}

		private void uGridVacaciones_BeforeCellUpdate(object sender, Infragistics.Win.UltraWinGrid.BeforeCellUpdateEventArgs e)
		{
			if (e.Cell.Column.ToString() == "idNominaTiposDeParto")
			{
				if ((int)e.Cell.Row.Cells["idNominaTiposDeSalidas"].Value != 4 && (int)e.Cell.Row.Cells["idNominaTiposDeSalidas"].Value != 5)
				{
					e.Cancel = true;
				}
			}		
		}

		private void uGridFaltas_AfterRowInsert(object sender, Infragistics.Win.UltraWinGrid.RowEventArgs e)
		{
			if (bEdicion)
			{
				e.Row.Cells["idNomina"].Value = 0;				
				e.Row.Cells["Estado"].Value = 0;
				e.Row.Cells["idNominaSalidas"].Value = 0;
				e.Row.Cells["idNominaTiposDeSalidas"].Value = 0;
			}
		}

		private void uGridUniformes_InitializeLayout(object sender, Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs e)
		{
			FuncionesProcedimientos.FormatoGrid(e, "Precio", 2);
			FuncionesProcedimientos.FormatoGrid(e, "Cuotas", 0);
		}

		private void uGridFaltas_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			FuncionesProcedimientos.DesplazamientoGrids(sender, e, this.uGridFaltas);
		}

		private void uGridOtros_BeforeCellUpdate(object sender, Infragistics.Win.UltraWinGrid.BeforeCellUpdateEventArgs e)
		{
			if (e.Cell.Column.ToString() == "idNominaMotivosOtros" || e.Cell.Column.ToString() == "Fecha" || e.Cell.Column.ToString() == "Valor" || e.Cell.Column.ToString() == "Observacion")
			{
				if ((int)e.Cell.Row.Cells["Estado"].Value == 1)
				{
					MessageBox.Show("Este descuento ya esta aplicado no puede modificar", "Point Tewchnology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
					e.Cancel = true;
				}
			}			
		}

		private void uGridOtros_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			FuncionesProcedimientos.DesplazamientoGrids(sender, e, this.uGridOtros);
		}

		private void uGridOtros_AfterRowInsert(object sender, Infragistics.Win.UltraWinGrid.RowEventArgs e)
		{
			if (bEdicion)
			{
				e.Row.Cells["idNominaOtros"].Value = 0;
				e.Row.Cells["idNomina"].Value = 0;
				e.Row.Cells["Estado"].Value = 0;
				e.Row.Cells["idNominaMotivosOtros"].Value = 0;
				e.Row.Cells["Valor"].Value = 0.00m;				
			}
		}

		private void uGridOtros_InitializeLayout(object sender, Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs e)
		{
			FuncionesProcedimientos.FormatoGrid(e, "Valor", 2);
		}

		private void uGridFaltas_BeforeCellUpdate(object sender, Infragistics.Win.UltraWinGrid.BeforeCellUpdateEventArgs e)
		{
//			if (e.Cell.Column.ToString() == "idNominaTiposDeSalidas")
//			{
//				if (e.Cell.Row.Cells["Fecha"].Value != System.DBNull.Value)
//				{
//					if (e.Cell.Row.Cells["idNominaTiposDeSalidas"].Value != System.DBNull.Value)
//					{						
//						if (Convert.ToDateTime(e.Cell.Row.Cells["Fecha"].Value).AddDays(4) < DateTime.Today && (int)e.Cell.Row.Cells["idNominaTiposDeSalidas"].Value == 2)
//						{
//							MessageBox.Show("Se permite modificar de Falta Injustificada a Justificada hasta 3 dias despues de la fecha de la falta", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
//							e.Cancel = true;
//							return;
//						}
//
//						if (Convert.ToDateTime(e.Cell.Row.Cells["Fecha"].Value).AddDays(4) < DateTime.Today && (int)e.Cell.Row.Cells["idNominaTiposDeSalidas"].Value == 1)
//						{
//							MessageBox.Show("Se permite modificar de Falta Justificada a Injustificada hasta 3 dias despues de la fecha de la falta", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
//							e.Cancel = true;
//							return;
//						}
//					}
//				}
//			}	
//		
//			if (e.Cell.Column.ToString() == "Fecha")
//			{
//				if ((int)e.Cell.Row.Cells["idNominaSalidas"].Value > 0)
//				{
//					if (e.Cell.Row.Cells["Fecha"].Value != System.DBNull.Value)
//					{
//						if (e.Cell.Row.Cells["idNominaTiposDeSalidas"].Value != System.DBNull.Value)
//						{
//							if (Funcion.dtEscalarSQL(cdsSeteoF, string.Format("Select Desde From NominaSalidas Where idNominaSalidas = {0}", (int)e.Cell.Row.Cells["idNominaSalidas"].Value)).AddDays(4) < DateTime.Today)
//							{
//								MessageBox.Show("Se permite modificar la fecha de una Falta hasta 3 dias despues de la fecha de la falta", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
//								e.Cancel = true;
//								return;
//							}
//						}
//					}
//				}
//			}
//	
//			if (e.Cell.Column.ToString() == "idNominaTiposDeSalidas" || e.Cell.Column.ToString() == "Fecha")
//			{
//				if (e.Cell.Row.Cells["Estado"].Value != System.DBNull.Value)
//				{
//					if (e.Cell.Row.Cells["Fecha"].Value != System.DBNull.Value)
//					{
//						if ((int)e.Cell.Row.Cells["Estado"].Value == 1)
//						{
//							MessageBox.Show("No puede modificar faltas que ya se han aplicado en Rol de Pagos", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
//							e.Cancel = true;
//							return;
//						}
//
//						if (Funcion.bEscalarSQL(cdsSeteoF, string.Format("Exec NominaValidaRolPagado '{0}'", Convert.ToDateTime(e.Cell.Row.Cells["Fecha"].Value).ToString("yyyyMMdd")), true))
//						{
//							MessageBox.Show("No puede modificar la falta, ya esta generado el Rol Del Periodo", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
//							e.Cancel = true;
//							return;
//						}
//					}
//				}
//			}
		}

		private void uGridFamiliares_InitializeLayout(object sender, Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs e)
		{
			
		}

		private void uGridFamiliares_BeforeRowsDeleted(object sender, Infragistics.Win.UltraWinGrid.BeforeRowsDeletedEventArgs e)
		{
			FuncionesProcedimientos.EliminaFilasGridNomina(e, 1, "idNominaFamiliar", miAcceso.BNominaEliminarExtras, cdsSeteoF);
		}

		private void uGridEstudios_BeforeRowsDeleted(object sender, Infragistics.Win.UltraWinGrid.BeforeRowsDeletedEventArgs e)
		{
			FuncionesProcedimientos.EliminaFilasGridNomina(e, 2, "idNominaEstudios", miAcceso.BNominaEliminarExtras, cdsSeteoF);
		}

		private void uGridExperiencia_BeforeRowsDeleted(object sender, Infragistics.Win.UltraWinGrid.BeforeRowsDeletedEventArgs e)
		{
			FuncionesProcedimientos.EliminaFilasGridNomina(e, 3, "idNominaExperiencia", miAcceso.BNominaEliminarExtras, cdsSeteoF);
		}

		private void uGridEntradasSalidas_BeforeRowsDeleted(object sender, Infragistics.Win.UltraWinGrid.BeforeRowsDeletedEventArgs e)
		{
//			FuncionesProcedimientos.EliminaFilasGridNomina(e, 4, "idNominaEntrada", miAcceso.BNominaEliminarMovimientos, cdsSeteoF);
		}

		private void uGridEquipos_BeforeRowsDeleted(object sender, Infragistics.Win.UltraWinGrid.BeforeRowsDeletedEventArgs e)
		{
			FuncionesProcedimientos.EliminaFilasGridNomina(e, 5, "idNominaEquipos", miAcceso.BNominaEliminarMovimientos, cdsSeteoF);
		}

		private void uGridVacaciones_BeforeRowsDeleted(object sender, Infragistics.Win.UltraWinGrid.BeforeRowsDeletedEventArgs e)
		{
			FuncionesProcedimientos.EliminaFilasGridNomina(e, 6, "idNominaSalidas", miAcceso.BNominaEliminarMovimientos, cdsSeteoF);
		}

		private void uGridUniformes_BeforeRowsDeleted(object sender, Infragistics.Win.UltraWinGrid.BeforeRowsDeletedEventArgs e)
		{
			FuncionesProcedimientos.EliminaFilasGridNomina(e, 7, "idNominaUniformes", miAcceso.BNominaEliminarMovimientos, cdsSeteoF);
		}

		private void uGridFaltas_BeforeRowsDeleted(object sender, Infragistics.Win.UltraWinGrid.BeforeRowsDeletedEventArgs e)
		{
			FuncionesProcedimientos.EliminaFilasGridNomina(e, 8, "idNominaSalidas", miAcceso.BNominaEliminarMovimientos, cdsSeteoF);
		}

		private void uGridOtros_BeforeRowsDeleted(object sender, Infragistics.Win.UltraWinGrid.BeforeRowsDeletedEventArgs e)
		{
			FuncionesProcedimientos.EliminaFilasGridNomina(e, 9, "idNominaOtros", miAcceso.BNominaEliminarMovimientos, cdsSeteoF);
		}

		private void uGridFaltas_InitializeLayout(object sender, Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs e)
		{
		
		}

		private void chkDerechoFondosDeReserva_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Enter) this.chkNETCC.Focus();
		}

		private void uGridOtros_AfterCellUpdate(object sender, Infragistics.Win.UltraWinGrid.CellEventArgs e)
		{
			
		}

		private void utcVarios_SelectedTabChanged_1(object sender, Infragistics.Win.UltraWinTabControl.SelectedTabChangedEventArgs e)
		{
		
		}

		private void cmbMotivos_InitializeLayout(object sender, Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs e)
		{
		
		}

		private void cmbNECargo_ValueChanged(object sender, System.EventArgs e)
		{
			if (bEdicion)			
			{
				if (this.cmbNEDepartamento.ActiveRow != null)
				{
					if ((int)this.cmbNEDepartamento.ActiveRow.Cells["idModulo"].Value == 1) 
					{
						if (bNuevo)					
						{
							if ((int)this.cmbNECargo.ActiveRow.Cells["Com_IdCargo"].Value <= 3 && (int)this.cmbNECargo.ActiveRow.Cells["Com_IdCargo"].Value != 0) 
								this.cmbRangos.Value = 2; else this.cmbRangos.Value = 0; 
						}
						else 
						{
							if ((int)this.cmbNECargo.ActiveRow.Cells["Com_IdCargo"].Value <= 3 && (int)this.cmbNECargo.ActiveRow.Cells["Com_IdCargo"].Value != 0) 
								this.cmbRangos.Value = 2; else this.cmbRangos.Value = 0; 
						}
					}
				}
			}
		}

		private void cmbNECargo_InitializeLayout(object sender, Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs e)
		{
		
		}

		private void cmbNEDepartamento_InitializeLayout(object sender, Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs e)
		{
		
		}

		private void cmbNEIdentificacion_ValueChanged(object sender, System.EventArgs e)
		{
			if (this.cmbNEIdentificacion.ActiveRow != null)
			{
				if ((int)this.cmbNEIdentificacion.Value == 1) this.txtNEIdentificacion.MaxLength = 10;
				if ((int)this.cmbNEIdentificacion.Value == 2) this.txtNEIdentificacion.MaxLength = 13;
			}
		}

		private void utcDatosGenerales_Paint(object sender, System.Windows.Forms.PaintEventArgs e)
		{
		
		}

		private void btnImportar_Click(object sender, System.EventArgs e)
		{
			try
			{
				OpenFileDialog sArchivo = new OpenFileDialog();

				sArchivo.Filter = "Microsoft Excel (*.xls)|*.xls";

				if (sArchivo.ShowDialog() == System.Windows.Forms.DialogResult.OK && sArchivo.FileName.Length > 0)
				{
					this.Cursor = Cursors.WaitCursor;                                  

					int iCont = 0;
					
					#region Sube archivo
					foreach (DataRow row in FuncionesProcedimientos.ImportaDeExcel(sArchivo.FileName, "SELECT * FROM [Nomina$]", "Nomina").Tables[0].Rows)
					{						//@Cedula Varchar(10), @TieneSolicitudde Varchar(2), @Fecha Datetime, @Tiene Varchar(2)
						#region Actualiza Informacion
						DateTime dtFecha = DateTime.Parse("01/01/2000");
						if (row["Fecha"] != System.DBNull.Value) dtFecha = DateTime.Parse(row["Fecha"].ToString());

						string sSQLArchivo = string.Format("Exec NominaActualizaEstadoFondosDeReserva '{0}', '{1}', '{2}', '{3}'",
							row["Cedula"].ToString().Trim(), row["TieneSolicitudde"].ToString().Trim(), dtFecha.ToString("yyyyMMdd"), row["Tiene"].ToString().Trim());
						Funcion.EjecutaSQL(cdsSeteoF, sSQLArchivo);
						#endregion Actualiza Informacion

						iCont++;
					}				
					#endregion Sube archivo

					MessageBox.Show(string.Format("Proceso Finalizado Correctamente, {0} Empleados Actualizados", iCont), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Information);					
					this.Cursor = Cursors.Default;
				}
			}
			catch (Exception ex)
			{
				MessageBox.Show(string.Format("Commit Exception Type : {0} {1}", ex.GetType(), ex.Message), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Error);
				this.Cursor = Cursors.Default;
				return;
			}
		}

		private void utcNomina_SelectedTabChanged(object sender, Infragistics.Win.UltraWinTabControl.SelectedTabChangedEventArgs e)
		{
		
		}

		private void groupBox1_Enter(object sender, System.EventArgs e)
		{
		
		}

		private void utpEstudios_Paint(object sender, System.Windows.Forms.PaintEventArgs e)
		{
		
		}

		private void chkNEAplicaRol_CheckedChanged(object sender, System.EventArgs e)
		{
		
		}

		private void ultraTabPageControl9_Paint(object sender, System.Windows.Forms.PaintEventArgs e)
		{
		
		}

		private void uGridCargas_AfterRowInsert(object sender, Infragistics.Win.UltraWinGrid.RowEventArgs e)
		{
			if (bEdicion)
			{
				e.Row.Cells["idNominaCargas"].Value = 0;
				e.Row.Cells["idNomina"].Value = 0;				
			}
		}

		private void uGridCargas_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			FuncionesProcedimientos.DesplazamientoGrids(sender, e, this.uGridCargas);
		}

		private void uGridCargas_BeforeRowsDeleted(object sender, Infragistics.Win.UltraWinGrid.BeforeRowsDeletedEventArgs e)
		{
			FuncionesProcedimientos.EliminaFilasGridNomina(e, 11, "idNominaCargas", miAcceso.BNominaEliminarMovimientos, cdsSeteoF);
		}

		private void btnAPP_Click(object sender, System.EventArgs e)
		{
			if(!miAcceso.BReliquidar)
			{
				MessageBox.Show("No tiene Acceso Activar Aplicación Movil.", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				return;
			}
			if (this.txtNEIdentificacion.Text.ToString().Length == 0)
			{
				MessageBox.Show("Seleccione un Usuario.", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				this.utcNomina.Tabs[0].Selected = true;
				this.txtNEIdentificacion.Focus();
				return;
			}
			if ((int) this.txtIdNomina.Value == 0)
			{
				MessageBox.Show("Seleccione un Usuario.", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				this.utcNomina.Tabs[0].Selected = true;
				this.txtNEIdentificacion.Focus();
				return;
			}


			if((int)this.txtEstado.Value == 2 )
			{
				MessageBox.Show("El personal Esta inactivo", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				return;
			}
			int iCont = Funcion.iEscalarSQL(cdsSeteoF,string.Format("SELECT COUNT(1) FROM DispositivosAPP WHERE idCom_Estado <> 2 AND CEDULA = '{0}' AND ACTIVO = 1 AND Empresa = 1",this.txtNEIdentificacion.Text.ToString()));
			if( iCont == 0)
			{
				MessageBox.Show("El Usuario no activa la aplicación Movil en su Dispositivo.", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				return;
			}

			if (DialogResult.Yes == MessageBox.Show("¿Desea volver Activar al Usuario?, Validar que no sea un segundo dispositivo.\n La activación es bajo su responsabilidad.", "Point Technology", MessageBoxButtons.YesNo, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button2))
			{
				#region Abre observacion
				using (frmJustAnulaVenta miNota = new frmJustAnulaVenta())
				{
					if (DialogResult.OK == miNota.ShowDialog())
					{	
						#region Guardar      
						using (SqlConnection oConexion = new SqlConnection("Password=Bl45o6$9;Persist Security Info=True;User ID= " + MenuLatinium.stUsuario + ";Initial Catalog=" + MenuLatinium.stDirFacturacion + ";Data Source=" + MenuLatinium.stDirServidor + "; max pool size=1024"))
						{
							oConexion.Open();

							SqlTransaction oTransaction = null;
							SqlCommand oCmdActualiza = oConexion.CreateCommand();
							oCmdActualiza.Connection = oConexion;
							oCmdActualiza.CommandType = CommandType.Text;
							oCmdActualiza.CommandTimeout = 120;

							try
							{
								oTransaction = oConexion.BeginTransaction();
								oCmdActualiza.Transaction = oTransaction;

								string sSQLAutEs = string.Format("Update DispositivosAPP set KeyDispositivo= cast(newid() AS NVARCHAR(50)), iContActivacion = (iContActivacion +1) ,  ACTIVO = 0 where CEDULA = '{0}'",
									this.txtNEIdentificacion.Text);
								oCmdActualiza.CommandText = sSQLAutEs;
								oCmdActualiza.ExecuteNonQuery();					
						

								string sSQLAutEsT = string.Format("iNSERT INTO DispositivosAPPDetalleActivacíon (Cedula, Notas) values ('{0}', '{1}')",
									this.txtNEIdentificacion.Text, miNota.txtNotas.Text);
								oCmdActualiza.CommandText = sSQLAutEsT;
								oCmdActualiza.ExecuteNonQuery();
		
								oTransaction.Commit();
								Cursor = Cursors.Default;	

								MessageBox.Show("Usuario Activado en la Aplicación Movil.", "Información", MessageBoxButtons.OK, MessageBoxIcon.Information);
								//this.btnConsultar_Click(sender,e);
				
							}
							catch(Exception ex)
							{
								#region execpciones
								try
								{ 
									oTransaction.Rollback();

									MessageBox.Show(string.Format("Commit Exception Type : {0} {1}", ex.GetType(), ex.Message), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Error);
								}
								catch (Exception ex2)
								{
									MessageBox.Show(string.Format("Rollback Exception Type: {0} {1}", ex2.GetType(), ex2.Message), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Error);							
								}
								#endregion execpciones
							}
							finally
							{
								oConexion.Close();
							}
						}
				
						#endregion Guardar					
					}
				}
				#endregion Abre observacion

			}
		}

		private void label41_Click(object sender, System.EventArgs e)
		{
		
		}

		private void chkNEVendedor_CheckedChanged(object sender, System.EventArgs e)
		{
		
		}

		private void optNERolesProservicios_ValueChanged(object sender, System.EventArgs e)
		{
		
		}
	}
}
