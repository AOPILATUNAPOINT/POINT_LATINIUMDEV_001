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
	/// Descripción breve de frmCred_GaranteSolicitud.
	/// </summary>

	
	public class frmCred_GaranteSolicitud : System.Windows.Forms.Form
	{
		#region infragistics
		private Infragistics.Win.UltraWinTabControl.UltraTabControl utcSolicitud;
		private Infragistics.Win.UltraWinTabControl.UltraTabSharedControlsPage ultraTabSharedControlsPage1;
		private Infragistics.Win.UltraWinTabControl.UltraTabPageControl utcDatosGenerales;
		private System.Windows.Forms.Label label57;
		private System.Windows.Forms.Label label56;
		private System.Windows.Forms.Label label55;
		private System.Windows.Forms.Label label45;
		private System.Windows.Forms.Label label44;
		private System.Windows.Forms.Label label54;
		private System.Windows.Forms.Label label13;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbNEEstadoCivil;
		private System.Windows.Forms.Label label14;
		private System.Windows.Forms.Label label15;
		private Infragistics.Win.UltraWinSchedule.UltraCalendarCombo dtpNEFechaNacimiento;
		private System.Windows.Forms.Label label22;
		private System.Windows.Forms.Label label21;
		private System.Windows.Forms.GroupBox groupBox1;
		private System.Windows.Forms.Label label65;
		private Infragistics.Win.UltraWinSchedule.UltraCalendarCombo dtpTiempoVivienda;
		private System.Windows.Forms.Label label64;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtNEEMail;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtNECelular;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtNETelefono2;
		private System.Windows.Forms.Label label62;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtNETelefono1;
		private System.Windows.Forms.Label label19;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtCalle2;
		private System.Windows.Forms.Label label61;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtNumcasa;
		private System.Windows.Forms.Label label38;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtCalle1;
		private System.Windows.Forms.Label label7;
		private System.Windows.Forms.Label label10;
		private System.Windows.Forms.Label label12;
		private System.Windows.Forms.Label label8;
		private System.Windows.Forms.Label label9;
		private System.Windows.Forms.Label label32;
		private System.Windows.Forms.Label label20;
		private System.Windows.Forms.Label label11;
		private Infragistics.Win.UltraWinTabControl.UltraTabPageControl utpEstudios;
		private System.Windows.Forms.GroupBox groupBox3;
		private System.Windows.Forms.Label label75;
		private System.Windows.Forms.Label label76;
		private System.Windows.Forms.Label label77;
		private System.Windows.Forms.Label label78;
		private System.Windows.Forms.Label label73;
		private System.Windows.Forms.Label label74;
		private System.Windows.Forms.Label label70;
		private Infragistics.Win.UltraWinSchedule.UltraCalendarCombo dpCFechaNacimiento;
		private System.Windows.Forms.Label label71;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtCSNombre;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtCPNombre;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtCApellidoM;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtCApellidoP;
		private System.Windows.Forms.Label label51;
		private System.Windows.Forms.Label label52;
		private System.Windows.Forms.Label label63;
		private System.Windows.Forms.Label label69;
		public Infragistics.Win.UltraWinEditors.UltraTextEditor txtCDocumento;
		private System.Windows.Forms.Label label50;
		private System.Windows.Forms.Label label49;
		private Infragistics.Win.UltraWinTabControl.UltraTabPageControl utpImagenes;
		private Infragistics.Win.UltraWinEditors.UltraCheckEditor chkSNAfiliacion;
		private Infragistics.Win.UltraWinEditors.UltraCheckEditor chkSNContabilidad;
		private System.Windows.Forms.Label label93;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtSNPasivos;
		private System.Windows.Forms.Label label94;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtSNActivos;
		private System.Windows.Forms.Label label95;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtSNGastos;
		private System.Windows.Forms.Label label96;
		private System.Windows.Forms.Label label97;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtSNReferencia;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtSNIngresos;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtSNEmail;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtSNCelular;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtSNTelefono2;
		private System.Windows.Forms.Label label86;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtSNTelefono1;
		private System.Windows.Forms.Label label87;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtSNCalle2;
		private System.Windows.Forms.Label label88;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtSNNumCasa;
		private System.Windows.Forms.Label label89;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtSNCalle1;
		private System.Windows.Forms.Label label90;
		private System.Windows.Forms.Label label91;
		private System.Windows.Forms.Label label92;
		private System.Windows.Forms.Label label82;
		private System.Windows.Forms.Label label83;
		private System.Windows.Forms.Label label84;
		private System.Windows.Forms.Label label85;
		private System.Windows.Forms.Label label81;
		private System.Windows.Forms.Label label80;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtSNEmpleados;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtSNMetros;
		private System.Windows.Forms.Label label53;
		private System.Windows.Forms.Label label79;
		private System.Windows.Forms.Label label48;
		private Infragistics.Win.UltraWinSchedule.UltraCalendarCombo dtSNFechaInicio;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtSNNombre;
		private System.Windows.Forms.Label label47;
		public Infragistics.Win.UltraWinEditors.UltraTextEditor txtSNRuc;
		private System.Windows.Forms.Label label46;
		private Infragistics.Win.UltraWinTabControl.UltraTabPageControl utpAdicional;
		private Infragistics.Win.UltraWinTabControl.UltraTabPageControl ultraTabPageControl1;
		private System.Windows.Forms.Label label120;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtSDExt;
		private System.Windows.Forms.Label label109;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtSDEmail;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtSDCelular;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtSDTelefono;
		private System.Windows.Forms.Label label110;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtSDCalle2;
		private System.Windows.Forms.Label label111;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtSDNumCasa;
		private System.Windows.Forms.Label label112;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtSDCalle1;
		private System.Windows.Forms.Label label113;
		private System.Windows.Forms.Label label114;
		private System.Windows.Forms.Label label115;
		private System.Windows.Forms.Label label116;
		private System.Windows.Forms.Label label117;
		private System.Windows.Forms.Label label118;
		private System.Windows.Forms.Label label119;
		private Infragistics.Win.UltraWinEditors.UltraCheckEditor chkSDAfiliacion;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtSDPago;
		private System.Windows.Forms.Label label108;
		public Infragistics.Win.UltraWinEditors.UltraTextEditor txtSDDepartamento;
		private System.Windows.Forms.Label label107;
		private System.Windows.Forms.Label label104;
		private System.Windows.Forms.Label label105;
		private System.Windows.Forms.Label label106;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtSDPasivos;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtSDActivos;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtSDGastos;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtSDIngresos;
		private System.Windows.Forms.Label label99;
		private System.Windows.Forms.Label label101;
		private System.Windows.Forms.Label label102;
		private System.Windows.Forms.Label label103;
		public Infragistics.Win.UltraWinEditors.UltraTextEditor txtSDEmpresa;
		private System.Windows.Forms.Label lblempres;
		private System.Windows.Forms.Label label100;
		public Infragistics.Win.UltraWinEditors.UltraTextEditor txtSDRuc;
		private System.Windows.Forms.Label label98;
		private System.Windows.Forms.Label label6;
		private Infragistics.Win.UltraWinSchedule.UltraCalendarCombo dtFecha;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbTipoRuc;
		private System.Windows.Forms.Label lblBodega;
		public Infragistics.Win.UltraWinEditors.UltraTextEditor txtDocumento;
		private System.Windows.Forms.Label label5;
		private System.Windows.Forms.Label label4;
		private System.Windows.Forms.Label label3;
		private System.Windows.Forms.Label label2;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtSnombre;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtPnombre;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtApellidoM;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtApellidoP;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbEstado;
		private System.Windows.Forms.Label label16;
		private System.Windows.Forms.Label label17;
		private System.Windows.Forms.Label label18;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtScore;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbCantonNacimiento;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbProvinciaNacimiento;
		private System.Windows.Forms.Button btnSalir;
		private System.Windows.Forms.Button btnGuardar;
		private Infragistics.Win.UltraWinGrid.UltraDropDown cmbParentesco;
		private System.Windows.Forms.Label label23;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtClienteIdCre_DatosUbicacion;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtClienteIdCre_DatosGenerales;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtSDReferencia;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtClienteIdCre_Negocio;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtNegocioIdCre_DatosUbicacion;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtClienteIdCre_EmpresasTrabajo;
		private C1.Data.C1DataSet cdsSeteoF;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbSituacionLaboral;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbTipoVivienda;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbParroquiaDomicilio;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbCantonDomicilio;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbProvinciaDomicilio;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbCActividadEconomia;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbCSituacionLaboral;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbCalificacion;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbCProvinciaNacimiento;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbCCantonNacimiento;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbSNProvincia;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbSNCanton;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbSNParroquia;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbSNBarrio;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbSDProvincia;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbSDCanton;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbSDParroquia;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbSDBarrio;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbSDCargo;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbSDTipoSueldo;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbSDTipoEmpresa;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtReferencia;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource26;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource7;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource6;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbNENacionalidad;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource1;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource2;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource4;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource19;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbProfesion;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource3;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource5;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource8;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource9;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbBarrioDomicilio;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource10;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource11;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbSCTipoNegocio;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource18;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource12;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbSubActvidadEconomica;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource13;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource14;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource15;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource16;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbCNacionalidad;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbCNivel;
		private Infragistics.Win.UltraWinGrid.UltraDropDown cmbRProvincia;
		private Infragistics.Win.UltraWinGrid.UltraDropDown cmbRCanton;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbCTipoRuc;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbSDTipoContrato;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbNivel;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource17;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbSexo;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbCSexo;
		private System.Windows.Forms.Label label24;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtNECargasFamiliares;
		/// <summary>
		/// Variable del diseñador requerida.
		/// </summary>
		private System.ComponentModel.Container components = null;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource20;
		private Infragistics.Win.UltraWinGrid.UltraGrid uGridParentesco;
		private System.Windows.Forms.Button btnTelf;
		private System.Windows.Forms.Button btnTelf2;
		private System.Windows.Forms.Button btnCelular;
		private System.Windows.Forms.Button btnCTelf;
		private System.Windows.Forms.Button btnCCelular;
		private System.Windows.Forms.Button btnCTelf2;
		private System.Windows.Forms.Button btnSDCelular;
		private System.Windows.Forms.Button bntSDTelf;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtEmpresaIdCre_DatosUbicacion;
		public Infragistics.Win.UltraWinEditors.UltraNumericEditor txtIdCre_Garante;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtIdCre_Conyugue;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtConyugueIdCre_DatosGenerales;
		private System.Windows.Forms.Label label1;
		#endregion

		string sCreSolicitud ;
		int idCre_Solicitud = 0;
		bool bCarga ;
		bool bNuevo = false;
		bool bEdicion = false;
		bool bEstado = false;
		bool bgarante = true;
		bool bPantGarante = true; // si es verdadero significa que el cliente no tiene garante
		int isolicitud = 0;
		public Infragistics.Win.UltraWinEditors.UltraNumericEditor txtIdCre_Solicitud;
		private Infragistics.Win.UltraWinSchedule.UltraCalendarCombo dtSDFecha;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbCProfesion;
		bool bnuevo = true;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbActividadEcomica;
		int ientidad = 0;

		public frmCred_GaranteSolicitud(string screSolicitud, int IdCre_Solicitud, bool bGarante, bool BPantGarante, int iSolicitud, bool bNuevo, int iEntidad)
		{
			//
			// Necesario para admitir el Diseñador de Windows Forms
			//
			InitializeComponent();

			sCreSolicitud = screSolicitud;
			idCre_Solicitud = IdCre_Solicitud;
			bgarante = bGarante;
			bPantGarante= BPantGarante;
			isolicitud = iSolicitud;
			bnuevo = bNuevo;
			ientidad = iEntidad;

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
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn1 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idSexo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn2 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn1 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idSexo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn2 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Nombre");
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand2 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn3 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idNivelEducacion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn4 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn3 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idNivelEducacion");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn4 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Nombre");
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand3 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn5 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idNacionalidad");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn6 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn5 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idNacionalidad");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn6 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Nombre");
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand4 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn7 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCanton");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn8 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn7 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idCanton");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn8 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Nombre");
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand5 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn9 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idProvincia");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn10 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn9 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idProvincia");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn10 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Nombre");
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand6 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn11 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idSituacionLaboral");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn12 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Descripcion");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn11 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idSituacionLaboral");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn12 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Descripcion");
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand7 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn13 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idEdoCivil");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn14 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn13 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idEdoCivil");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn14 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Nombre");
			Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton dateButton1 = new Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton();
			System.Resources.ResourceManager resources = new System.Resources.ResourceManager(typeof(frmCred_GaranteSolicitud));
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand8 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn15 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idBarrio");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn16 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn15 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idBarrio");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn16 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Nombre");
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand9 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn17 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idParroquia");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn18 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn17 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idParroquia");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn18 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Nombre");
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand10 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn19 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCanton");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn20 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand11 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn21 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idProvincia");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn22 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton dateButton2 = new Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand12 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn23 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idTipoVivienda");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn24 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn25 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Codigo", 0);
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn19 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idTipoVivienda");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn20 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Nombre");
			Infragistics.Win.Appearance appearance1 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand13 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn26 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idProfesion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn27 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn21 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idProfesion");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn22 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Nombre");
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand14 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn28 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idActEconomica");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn29 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn23 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idActEconomica");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn24 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Nombre");
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand15 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn30 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idProfesion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn31 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand16 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn32 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idSexo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn33 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand17 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn34 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idNivelEducacion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn35 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand18 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn36 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idNacionalidad");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn37 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand19 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn38 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idActEconomica");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn39 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand20 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn40 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idSituacionLaboral");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn41 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Descripcion");
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand21 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn42 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idProvincia");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn43 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand22 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn44 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCanton");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn45 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton dateButton3 = new Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton();
			Infragistics.Win.Appearance appearance2 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance3 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance4 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance5 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance6 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance7 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand23 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn46 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCre_ReferenciasClientes");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn47 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCre_DatosGenerales");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn48 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Indice");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn49 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idParentesco", -1, "cmbParentesco");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn50 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("ApellidoPaterno");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn51 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("ApellidoMaterno");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn52 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("PrimerNombre");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn53 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("SegundoNombre");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn54 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idProvincia", -1, "cmbRProvincia");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn55 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCanton", -1, "cmbRCanton");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn56 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Direccion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn57 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Telefono");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn58 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Celular");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn59 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("EmpresaTrabajo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn60 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("DireccionTrabajo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn61 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("TelefonoTrabajo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn62 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("ExtensionTrabajo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn63 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("btnTelefono", 0);
			Infragistics.Win.Appearance appearance8 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn64 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("btnCelular", 1);
			Infragistics.Win.Appearance appearance9 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn65 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("btnTelTrabajo", 2);
			Infragistics.Win.Appearance appearance10 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance11 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance12 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance13 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance14 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn25 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idCre_ReferenciasClientes");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn26 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idCre_DatosGenerales");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn27 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Indice");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn28 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idParentesco");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn29 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("ApellidoPaterno");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn30 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("ApellidoMaterno");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn31 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("PrimerNombre");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn32 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("SegundoNombre");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn33 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idProvincia");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn34 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idCanton");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn35 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Direccion");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn36 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Telefono");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn37 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Celular");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn38 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("EmpresaTrabajo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn39 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("DireccionTrabajo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn40 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("TelefonoTrabajo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn41 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("ExtensionTrabajo");
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand24 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn66 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idProvincia");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn67 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand25 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn68 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCanton");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn69 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand26 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn70 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idParentesco");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn71 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn42 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idParentesco");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn43 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Nombre");
			Infragistics.Win.Appearance appearance15 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand27 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn72 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idSubActEcoNegocio");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn73 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn44 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idSubActEcoNegocio");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn45 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Nombre");
			Infragistics.Win.Appearance appearance16 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance17 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance18 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand28 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn74 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idProvincia");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn75 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand29 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn76 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCanton");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn77 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand30 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn78 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idParroquia");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn79 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand31 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn80 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idBarrio");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn81 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.Appearance appearance19 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand32 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn82 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idTipoVivienda");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn83 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.Appearance appearance20 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton dateButton4 = new Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton();
			Infragistics.Win.Appearance appearance21 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance22 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance23 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton dateButton5 = new Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand33 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn84 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idTipoContrato");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn85 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn46 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idTipoContrato");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn47 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Nombre");
			Infragistics.Win.Appearance appearance24 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand34 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn86 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idProvincia");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn87 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand35 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn88 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCanton");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn89 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand36 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn90 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idParroquia");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn91 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand37 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn92 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idBarrio");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn93 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.Appearance appearance25 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance26 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand38 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn94 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCargo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn95 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn96 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Codigo", 0);
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn48 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idCargo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn49 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Nombre");
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand39 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn97 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idTipoSueldo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn98 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn50 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idTipoSueldo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn51 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Nombre");
			Infragistics.Win.Appearance appearance27 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand40 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn99 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idTipoEmpresa");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn100 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Descripción");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn52 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idTipoEmpresa");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn53 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Descripción");
			Infragistics.Win.Appearance appearance28 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance29 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance30 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinTabControl.UltraTab ultraTab1 = new Infragistics.Win.UltraWinTabControl.UltraTab();
			Infragistics.Win.UltraWinTabControl.UltraTab ultraTab2 = new Infragistics.Win.UltraWinTabControl.UltraTab();
			Infragistics.Win.UltraWinTabControl.UltraTab ultraTab3 = new Infragistics.Win.UltraWinTabControl.UltraTab();
			Infragistics.Win.UltraWinTabControl.UltraTab ultraTab4 = new Infragistics.Win.UltraWinTabControl.UltraTab();
			Infragistics.Win.UltraWinTabControl.UltraTab ultraTab5 = new Infragistics.Win.UltraWinTabControl.UltraTab();
			Infragistics.Win.Appearance appearance31 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton dateButton6 = new Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton();
			Infragistics.Win.Appearance appearance32 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance33 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance34 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance35 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance36 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand41 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn101 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idTipoCalificacion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn102 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Codigo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn54 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idTipoCalificacion");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn55 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Codigo");
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand42 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn103 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idEstado");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn104 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Estado");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn56 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idEstado");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn57 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Estado");
			this.utcDatosGenerales = new Infragistics.Win.UltraWinTabControl.UltraTabPageControl();
			this.cmbSexo = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.ultraDataSource17 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.cmbNivel = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.ultraDataSource19 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.cmbNENacionalidad = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.ultraDataSource26 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.cmbCantonNacimiento = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.ultraDataSource2 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.cmbProvinciaNacimiento = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.ultraDataSource1 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.label57 = new System.Windows.Forms.Label();
			this.cmbSituacionLaboral = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.ultraDataSource5 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.label56 = new System.Windows.Forms.Label();
			this.label55 = new System.Windows.Forms.Label();
			this.label45 = new System.Windows.Forms.Label();
			this.label44 = new System.Windows.Forms.Label();
			this.label54 = new System.Windows.Forms.Label();
			this.label13 = new System.Windows.Forms.Label();
			this.cmbNEEstadoCivil = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.ultraDataSource4 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.label14 = new System.Windows.Forms.Label();
			this.label15 = new System.Windows.Forms.Label();
			this.dtpNEFechaNacimiento = new Infragistics.Win.UltraWinSchedule.UltraCalendarCombo();
			this.label22 = new System.Windows.Forms.Label();
			this.label21 = new System.Windows.Forms.Label();
			this.groupBox1 = new System.Windows.Forms.GroupBox();
			this.btnCelular = new System.Windows.Forms.Button();
			this.btnTelf2 = new System.Windows.Forms.Button();
			this.btnTelf = new System.Windows.Forms.Button();
			this.cmbBarrioDomicilio = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.ultraDataSource10 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.cmbParroquiaDomicilio = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.ultraDataSource9 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.cmbCantonDomicilio = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.cmbProvinciaDomicilio = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.label65 = new System.Windows.Forms.Label();
			this.dtpTiempoVivienda = new Infragistics.Win.UltraWinSchedule.UltraCalendarCombo();
			this.cmbTipoVivienda = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.ultraDataSource11 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.label64 = new System.Windows.Forms.Label();
			this.txtReferencia = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.txtNEEMail = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.txtNECelular = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.txtNETelefono2 = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.label62 = new System.Windows.Forms.Label();
			this.txtNETelefono1 = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.label19 = new System.Windows.Forms.Label();
			this.txtCalle2 = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.label61 = new System.Windows.Forms.Label();
			this.txtNumcasa = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.label38 = new System.Windows.Forms.Label();
			this.txtCalle1 = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.label7 = new System.Windows.Forms.Label();
			this.label10 = new System.Windows.Forms.Label();
			this.label12 = new System.Windows.Forms.Label();
			this.label8 = new System.Windows.Forms.Label();
			this.label9 = new System.Windows.Forms.Label();
			this.label32 = new System.Windows.Forms.Label();
			this.label20 = new System.Windows.Forms.Label();
			this.label11 = new System.Windows.Forms.Label();
			this.cmbProfesion = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.ultraDataSource3 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.txtNECargasFamiliares = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.cmbActividadEcomica = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.ultraDataSource8 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.utpEstudios = new Infragistics.Win.UltraWinTabControl.UltraTabPageControl();
			this.groupBox3 = new System.Windows.Forms.GroupBox();
			this.cmbCProfesion = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.cmbCSexo = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.label24 = new System.Windows.Forms.Label();
			this.cmbCTipoRuc = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.cmbCNivel = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.cmbCNacionalidad = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.cmbCActividadEconomia = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.label75 = new System.Windows.Forms.Label();
			this.cmbCSituacionLaboral = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.label76 = new System.Windows.Forms.Label();
			this.label77 = new System.Windows.Forms.Label();
			this.label78 = new System.Windows.Forms.Label();
			this.cmbCProvinciaNacimiento = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.label73 = new System.Windows.Forms.Label();
			this.cmbCCantonNacimiento = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.label74 = new System.Windows.Forms.Label();
			this.label70 = new System.Windows.Forms.Label();
			this.dpCFechaNacimiento = new Infragistics.Win.UltraWinSchedule.UltraCalendarCombo();
			this.label71 = new System.Windows.Forms.Label();
			this.txtCSNombre = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.txtCPNombre = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.txtCApellidoM = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.txtCApellidoP = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.label51 = new System.Windows.Forms.Label();
			this.label52 = new System.Windows.Forms.Label();
			this.label63 = new System.Windows.Forms.Label();
			this.label69 = new System.Windows.Forms.Label();
			this.txtCDocumento = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.label50 = new System.Windows.Forms.Label();
			this.label49 = new System.Windows.Forms.Label();
			this.utpAdicional = new Infragistics.Win.UltraWinTabControl.UltraTabPageControl();
			this.uGridParentesco = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.ultraDataSource20 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.cmbRProvincia = new Infragistics.Win.UltraWinGrid.UltraDropDown();
			this.cmbRCanton = new Infragistics.Win.UltraWinGrid.UltraDropDown();
			this.cmbParentesco = new Infragistics.Win.UltraWinGrid.UltraDropDown();
			this.ultraDataSource18 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.utpImagenes = new Infragistics.Win.UltraWinTabControl.UltraTabPageControl();
			this.btnCCelular = new System.Windows.Forms.Button();
			this.btnCTelf2 = new System.Windows.Forms.Button();
			this.btnCTelf = new System.Windows.Forms.Button();
			this.cmbSubActvidadEconomica = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.ultraDataSource12 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.chkSNAfiliacion = new Infragistics.Win.UltraWinEditors.UltraCheckEditor();
			this.chkSNContabilidad = new Infragistics.Win.UltraWinEditors.UltraCheckEditor();
			this.label93 = new System.Windows.Forms.Label();
			this.txtSNPasivos = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.label94 = new System.Windows.Forms.Label();
			this.txtSNActivos = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.label95 = new System.Windows.Forms.Label();
			this.txtSNGastos = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.label96 = new System.Windows.Forms.Label();
			this.label97 = new System.Windows.Forms.Label();
			this.txtSNReferencia = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.txtSNIngresos = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtSNEmail = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.txtSNCelular = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.txtSNTelefono2 = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.label86 = new System.Windows.Forms.Label();
			this.txtSNTelefono1 = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.label87 = new System.Windows.Forms.Label();
			this.txtSNCalle2 = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.label88 = new System.Windows.Forms.Label();
			this.txtSNNumCasa = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.label89 = new System.Windows.Forms.Label();
			this.txtSNCalle1 = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.label90 = new System.Windows.Forms.Label();
			this.label91 = new System.Windows.Forms.Label();
			this.label92 = new System.Windows.Forms.Label();
			this.label82 = new System.Windows.Forms.Label();
			this.cmbSNProvincia = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.label83 = new System.Windows.Forms.Label();
			this.cmbSNCanton = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.label84 = new System.Windows.Forms.Label();
			this.cmbSNParroquia = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.label85 = new System.Windows.Forms.Label();
			this.cmbSNBarrio = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.label81 = new System.Windows.Forms.Label();
			this.label80 = new System.Windows.Forms.Label();
			this.cmbSCTipoNegocio = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.txtSNEmpleados = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtSNMetros = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.label53 = new System.Windows.Forms.Label();
			this.label79 = new System.Windows.Forms.Label();
			this.label48 = new System.Windows.Forms.Label();
			this.dtSNFechaInicio = new Infragistics.Win.UltraWinSchedule.UltraCalendarCombo();
			this.txtSNNombre = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.label47 = new System.Windows.Forms.Label();
			this.txtSNRuc = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.label46 = new System.Windows.Forms.Label();
			this.ultraTabPageControl1 = new Infragistics.Win.UltraWinTabControl.UltraTabPageControl();
			this.dtSDFecha = new Infragistics.Win.UltraWinSchedule.UltraCalendarCombo();
			this.btnSDCelular = new System.Windows.Forms.Button();
			this.bntSDTelf = new System.Windows.Forms.Button();
			this.cmbSDTipoContrato = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.ultraDataSource13 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.txtSDReferencia = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.label23 = new System.Windows.Forms.Label();
			this.label120 = new System.Windows.Forms.Label();
			this.txtSDExt = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.label109 = new System.Windows.Forms.Label();
			this.txtSDEmail = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.txtSDCelular = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.txtSDTelefono = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.label110 = new System.Windows.Forms.Label();
			this.txtSDCalle2 = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.label111 = new System.Windows.Forms.Label();
			this.txtSDNumCasa = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.label112 = new System.Windows.Forms.Label();
			this.txtSDCalle1 = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.label113 = new System.Windows.Forms.Label();
			this.label114 = new System.Windows.Forms.Label();
			this.label115 = new System.Windows.Forms.Label();
			this.label116 = new System.Windows.Forms.Label();
			this.cmbSDProvincia = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.label117 = new System.Windows.Forms.Label();
			this.cmbSDCanton = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.label118 = new System.Windows.Forms.Label();
			this.cmbSDParroquia = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.label119 = new System.Windows.Forms.Label();
			this.cmbSDBarrio = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.chkSDAfiliacion = new Infragistics.Win.UltraWinEditors.UltraCheckEditor();
			this.txtSDPago = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.label108 = new System.Windows.Forms.Label();
			this.txtSDDepartamento = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.label107 = new System.Windows.Forms.Label();
			this.label104 = new System.Windows.Forms.Label();
			this.cmbSDCargo = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.ultraDataSource15 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.cmbSDTipoSueldo = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.ultraDataSource16 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.label105 = new System.Windows.Forms.Label();
			this.label106 = new System.Windows.Forms.Label();
			this.txtSDPasivos = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtSDActivos = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtSDGastos = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtSDIngresos = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.label99 = new System.Windows.Forms.Label();
			this.label101 = new System.Windows.Forms.Label();
			this.label102 = new System.Windows.Forms.Label();
			this.label103 = new System.Windows.Forms.Label();
			this.txtSDEmpresa = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.cmbSDTipoEmpresa = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.ultraDataSource14 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.lblempres = new System.Windows.Forms.Label();
			this.label100 = new System.Windows.Forms.Label();
			this.txtSDRuc = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.label98 = new System.Windows.Forms.Label();
			this.utcSolicitud = new Infragistics.Win.UltraWinTabControl.UltraTabControl();
			this.ultraTabSharedControlsPage1 = new Infragistics.Win.UltraWinTabControl.UltraTabSharedControlsPage();
			this.label6 = new System.Windows.Forms.Label();
			this.dtFecha = new Infragistics.Win.UltraWinSchedule.UltraCalendarCombo();
			this.cmbTipoRuc = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.lblBodega = new System.Windows.Forms.Label();
			this.txtDocumento = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.label5 = new System.Windows.Forms.Label();
			this.label4 = new System.Windows.Forms.Label();
			this.label3 = new System.Windows.Forms.Label();
			this.label2 = new System.Windows.Forms.Label();
			this.txtSnombre = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.txtPnombre = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.txtApellidoM = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.txtApellidoP = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.cmbCalificacion = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.ultraDataSource6 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.cmbEstado = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.ultraDataSource7 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.label16 = new System.Windows.Forms.Label();
			this.label17 = new System.Windows.Forms.Label();
			this.label18 = new System.Windows.Forms.Label();
			this.txtScore = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.btnSalir = new System.Windows.Forms.Button();
			this.btnGuardar = new System.Windows.Forms.Button();
			this.txtClienteIdCre_DatosUbicacion = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtIdCre_Conyugue = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtClienteIdCre_DatosGenerales = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtNegocioIdCre_DatosUbicacion = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtClienteIdCre_EmpresasTrabajo = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtConyugueIdCre_DatosGenerales = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtClienteIdCre_Negocio = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtEmpresaIdCre_DatosUbicacion = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.cdsSeteoF = new C1.Data.C1DataSet();
			this.txtIdCre_Garante = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.label1 = new System.Windows.Forms.Label();
			this.txtIdCre_Solicitud = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.utcDatosGenerales.SuspendLayout();
			((System.ComponentModel.ISupportInitialize)(this.cmbSexo)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource17)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbNivel)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource19)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbNENacionalidad)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource26)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbCantonNacimiento)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource2)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbProvinciaNacimiento)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource1)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbSituacionLaboral)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource5)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbNEEstadoCivil)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource4)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.dtpNEFechaNacimiento)).BeginInit();
			this.groupBox1.SuspendLayout();
			((System.ComponentModel.ISupportInitialize)(this.cmbBarrioDomicilio)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource10)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbParroquiaDomicilio)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource9)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbCantonDomicilio)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbProvinciaDomicilio)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.dtpTiempoVivienda)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbTipoVivienda)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource11)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtReferencia)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNEEMail)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNECelular)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNETelefono2)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNETelefono1)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtCalle2)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNumcasa)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtCalle1)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbProfesion)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource3)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNECargasFamiliares)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbActividadEcomica)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource8)).BeginInit();
			this.utpEstudios.SuspendLayout();
			this.groupBox3.SuspendLayout();
			((System.ComponentModel.ISupportInitialize)(this.cmbCProfesion)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbCSexo)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbCTipoRuc)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbCNivel)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbCNacionalidad)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbCActividadEconomia)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbCSituacionLaboral)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbCProvinciaNacimiento)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbCCantonNacimiento)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.dpCFechaNacimiento)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtCSNombre)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtCPNombre)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtCApellidoM)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtCApellidoP)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtCDocumento)).BeginInit();
			this.utpAdicional.SuspendLayout();
			((System.ComponentModel.ISupportInitialize)(this.uGridParentesco)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource20)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbRProvincia)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbRCanton)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbParentesco)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource18)).BeginInit();
			this.utpImagenes.SuspendLayout();
			((System.ComponentModel.ISupportInitialize)(this.cmbSubActvidadEconomica)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource12)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtSNPasivos)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtSNActivos)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtSNGastos)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtSNReferencia)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtSNIngresos)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtSNEmail)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtSNCelular)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtSNTelefono2)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtSNTelefono1)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtSNCalle2)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtSNNumCasa)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtSNCalle1)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbSNProvincia)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbSNCanton)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbSNParroquia)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbSNBarrio)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbSCTipoNegocio)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtSNEmpleados)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtSNMetros)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.dtSNFechaInicio)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtSNNombre)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtSNRuc)).BeginInit();
			this.ultraTabPageControl1.SuspendLayout();
			((System.ComponentModel.ISupportInitialize)(this.dtSDFecha)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbSDTipoContrato)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource13)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtSDReferencia)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtSDExt)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtSDEmail)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtSDCelular)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtSDTelefono)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtSDCalle2)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtSDNumCasa)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtSDCalle1)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbSDProvincia)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbSDCanton)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbSDParroquia)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbSDBarrio)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtSDPago)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtSDDepartamento)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbSDCargo)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource15)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbSDTipoSueldo)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource16)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtSDPasivos)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtSDActivos)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtSDGastos)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtSDIngresos)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtSDEmpresa)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbSDTipoEmpresa)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource14)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtSDRuc)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.utcSolicitud)).BeginInit();
			this.utcSolicitud.SuspendLayout();
			((System.ComponentModel.ISupportInitialize)(this.dtFecha)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbTipoRuc)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtDocumento)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtSnombre)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtPnombre)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtApellidoM)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtApellidoP)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbCalificacion)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource6)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbEstado)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource7)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtScore)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtClienteIdCre_DatosUbicacion)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIdCre_Conyugue)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtClienteIdCre_DatosGenerales)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNegocioIdCre_DatosUbicacion)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtClienteIdCre_EmpresasTrabajo)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtConyugueIdCre_DatosGenerales)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtClienteIdCre_Negocio)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtEmpresaIdCre_DatosUbicacion)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIdCre_Garante)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIdCre_Solicitud)).BeginInit();
			this.SuspendLayout();
			// 
			// utcDatosGenerales
			// 
			this.utcDatosGenerales.Controls.Add(this.cmbSexo);
			this.utcDatosGenerales.Controls.Add(this.cmbNivel);
			this.utcDatosGenerales.Controls.Add(this.cmbNENacionalidad);
			this.utcDatosGenerales.Controls.Add(this.cmbCantonNacimiento);
			this.utcDatosGenerales.Controls.Add(this.cmbProvinciaNacimiento);
			this.utcDatosGenerales.Controls.Add(this.label57);
			this.utcDatosGenerales.Controls.Add(this.cmbSituacionLaboral);
			this.utcDatosGenerales.Controls.Add(this.label56);
			this.utcDatosGenerales.Controls.Add(this.label55);
			this.utcDatosGenerales.Controls.Add(this.label45);
			this.utcDatosGenerales.Controls.Add(this.label44);
			this.utcDatosGenerales.Controls.Add(this.label54);
			this.utcDatosGenerales.Controls.Add(this.label13);
			this.utcDatosGenerales.Controls.Add(this.cmbNEEstadoCivil);
			this.utcDatosGenerales.Controls.Add(this.label14);
			this.utcDatosGenerales.Controls.Add(this.label15);
			this.utcDatosGenerales.Controls.Add(this.dtpNEFechaNacimiento);
			this.utcDatosGenerales.Controls.Add(this.label22);
			this.utcDatosGenerales.Controls.Add(this.label21);
			this.utcDatosGenerales.Controls.Add(this.groupBox1);
			this.utcDatosGenerales.Controls.Add(this.cmbProfesion);
			this.utcDatosGenerales.Controls.Add(this.txtNECargasFamiliares);
			this.utcDatosGenerales.Controls.Add(this.cmbActividadEcomica);
			this.utcDatosGenerales.Location = new System.Drawing.Point(-10000, -10000);
			this.utcDatosGenerales.Name = "utcDatosGenerales";
			this.utcDatosGenerales.Size = new System.Drawing.Size(1102, 323);
			// 
			// cmbSexo
			// 
			this.cmbSexo.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbSexo.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbSexo.DataSource = this.ultraDataSource17;
			ultraGridBand1.AddButtonCaption = "RolSucursal";
			ultraGridColumn1.Header.VisiblePosition = 0;
			ultraGridColumn1.Hidden = true;
			ultraGridColumn2.Header.VisiblePosition = 1;
			ultraGridColumn2.Width = 144;
			ultraGridBand1.Columns.AddRange(new object[] {
																										 ultraGridColumn1,
																										 ultraGridColumn2});
			this.cmbSexo.DisplayLayout.BandsSerializer.Add(ultraGridBand1);
			this.cmbSexo.DisplayMember = "Nombre";
			this.cmbSexo.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbSexo.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbSexo.Enabled = false;
			this.cmbSexo.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbSexo.Location = new System.Drawing.Point(352, 24);
			this.cmbSexo.Name = "cmbSexo";
			this.cmbSexo.Size = new System.Drawing.Size(144, 22);
			this.cmbSexo.TabIndex = 808;
			this.cmbSexo.ValueMember = "idSexo";
			this.cmbSexo.KeyDown += new System.Windows.Forms.KeyEventHandler(this.cmbSexo_KeyDown);
			// 
			// ultraDataSource17
			// 
			ultraDataColumn1.DataType = typeof(int);
			this.ultraDataSource17.Band.Columns.AddRange(new object[] {
																																	ultraDataColumn1,
																																	ultraDataColumn2});
			// 
			// cmbNivel
			// 
			this.cmbNivel.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbNivel.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbNivel.DataSource = this.ultraDataSource19;
			ultraGridColumn3.Header.VisiblePosition = 0;
			ultraGridColumn3.Hidden = true;
			ultraGridColumn3.Width = 160;
			ultraGridColumn4.Header.VisiblePosition = 1;
			ultraGridColumn4.Width = 160;
			ultraGridBand2.Columns.AddRange(new object[] {
																										 ultraGridColumn3,
																										 ultraGridColumn4});
			this.cmbNivel.DisplayLayout.BandsSerializer.Add(ultraGridBand2);
			this.cmbNivel.DisplayLayout.Override.RowSelectors = Infragistics.Win.DefaultableBoolean.True;
			this.cmbNivel.DisplayMember = "Nombre";
			this.cmbNivel.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbNivel.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbNivel.Enabled = false;
			this.cmbNivel.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbNivel.Location = new System.Drawing.Point(248, 72);
			this.cmbNivel.Name = "cmbNivel";
			this.cmbNivel.Size = new System.Drawing.Size(160, 22);
			this.cmbNivel.TabIndex = 807;
			this.cmbNivel.ValueMember = "idNivelEducacion";
			this.cmbNivel.KeyDown += new System.Windows.Forms.KeyEventHandler(this.cmbNivel_KeyDown);
			// 
			// ultraDataSource19
			// 
			ultraDataColumn3.DataType = typeof(int);
			this.ultraDataSource19.Band.Columns.AddRange(new object[] {
																																	ultraDataColumn3,
																																	ultraDataColumn4});
			// 
			// cmbNENacionalidad
			// 
			this.cmbNENacionalidad.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbNENacionalidad.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbNENacionalidad.DataSource = this.ultraDataSource26;
			ultraGridBand3.AddButtonCaption = "RolSucursal";
			ultraGridColumn5.Header.VisiblePosition = 0;
			ultraGridColumn5.Hidden = true;
			ultraGridColumn6.Header.VisiblePosition = 1;
			ultraGridColumn6.Width = 144;
			ultraGridBand3.Columns.AddRange(new object[] {
																										 ultraGridColumn5,
																										 ultraGridColumn6});
			this.cmbNENacionalidad.DisplayLayout.BandsSerializer.Add(ultraGridBand3);
			this.cmbNENacionalidad.DisplayMember = "Nombre";
			this.cmbNENacionalidad.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbNENacionalidad.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbNENacionalidad.Enabled = false;
			this.cmbNENacionalidad.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbNENacionalidad.Location = new System.Drawing.Point(16, 24);
			this.cmbNENacionalidad.Name = "cmbNENacionalidad";
			this.cmbNENacionalidad.Size = new System.Drawing.Size(144, 22);
			this.cmbNENacionalidad.TabIndex = 693;
			this.cmbNENacionalidad.ValueMember = "idNacionalidad";
			this.cmbNENacionalidad.KeyDown += new System.Windows.Forms.KeyEventHandler(this.cmbNENacionalidad_KeyDown);
			// 
			// ultraDataSource26
			// 
			ultraDataColumn5.DataType = typeof(int);
			this.ultraDataSource26.Band.Columns.AddRange(new object[] {
																																	ultraDataColumn5,
																																	ultraDataColumn6});
			// 
			// cmbCantonNacimiento
			// 
			this.cmbCantonNacimiento.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbCantonNacimiento.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbCantonNacimiento.DataSource = this.ultraDataSource2;
			ultraGridColumn7.Header.VisiblePosition = 0;
			ultraGridColumn7.Hidden = true;
			ultraGridColumn8.Header.Caption = "Canton";
			ultraGridColumn8.Header.VisiblePosition = 1;
			ultraGridColumn8.Width = 150;
			ultraGridBand4.Columns.AddRange(new object[] {
																										 ultraGridColumn7,
																										 ultraGridColumn8});
			this.cmbCantonNacimiento.DisplayLayout.BandsSerializer.Add(ultraGridBand4);
			this.cmbCantonNacimiento.DisplayLayout.Override.RowSelectors = Infragistics.Win.DefaultableBoolean.True;
			this.cmbCantonNacimiento.DisplayMember = "Nombre";
			this.cmbCantonNacimiento.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbCantonNacimiento.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbCantonNacimiento.Enabled = false;
			this.cmbCantonNacimiento.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbCantonNacimiento.Location = new System.Drawing.Point(672, 24);
			this.cmbCantonNacimiento.Name = "cmbCantonNacimiento";
			this.cmbCantonNacimiento.Size = new System.Drawing.Size(150, 22);
			this.cmbCantonNacimiento.TabIndex = 692;
			this.cmbCantonNacimiento.ValueMember = "idCanton";
			this.cmbCantonNacimiento.KeyDown += new System.Windows.Forms.KeyEventHandler(this.cmbCantonNacimiento_KeyDown);
			// 
			// ultraDataSource2
			// 
			ultraDataColumn7.DataType = typeof(int);
			this.ultraDataSource2.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn7,
																																 ultraDataColumn8});
			// 
			// cmbProvinciaNacimiento
			// 
			this.cmbProvinciaNacimiento.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbProvinciaNacimiento.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbProvinciaNacimiento.DataSource = this.ultraDataSource1;
			ultraGridColumn9.Header.VisiblePosition = 0;
			ultraGridColumn9.Hidden = true;
			ultraGridColumn9.Width = 50;
			ultraGridColumn10.Header.VisiblePosition = 1;
			ultraGridColumn10.Width = 150;
			ultraGridBand5.Columns.AddRange(new object[] {
																										 ultraGridColumn9,
																										 ultraGridColumn10});
			this.cmbProvinciaNacimiento.DisplayLayout.BandsSerializer.Add(ultraGridBand5);
			this.cmbProvinciaNacimiento.DisplayMember = "Nombre";
			this.cmbProvinciaNacimiento.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbProvinciaNacimiento.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbProvinciaNacimiento.Enabled = false;
			this.cmbProvinciaNacimiento.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbProvinciaNacimiento.Location = new System.Drawing.Point(512, 24);
			this.cmbProvinciaNacimiento.Name = "cmbProvinciaNacimiento";
			this.cmbProvinciaNacimiento.Size = new System.Drawing.Size(150, 22);
			this.cmbProvinciaNacimiento.TabIndex = 691;
			this.cmbProvinciaNacimiento.ValueMember = "idProvincia";
			this.cmbProvinciaNacimiento.KeyDown += new System.Windows.Forms.KeyEventHandler(this.cmbProvinciaNacimiento_KeyDown);
			this.cmbProvinciaNacimiento.ValueChanged += new System.EventHandler(this.cmbProvinciaNacimiento_ValueChanged);
			// 
			// ultraDataSource1
			// 
			ultraDataColumn9.DataType = typeof(int);
			this.ultraDataSource1.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn9,
																																 ultraDataColumn10});
			// 
			// label57
			// 
			this.label57.AutoSize = true;
			this.label57.Location = new System.Drawing.Point(792, 48);
			this.label57.Name = "label57";
			this.label57.Size = new System.Drawing.Size(106, 17);
			this.label57.TabIndex = 681;
			this.label57.Text = "Actividad Económica";
			this.label57.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// cmbSituacionLaboral
			// 
			this.cmbSituacionLaboral.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbSituacionLaboral.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbSituacionLaboral.DataSource = this.ultraDataSource5;
			ultraGridColumn11.Header.VisiblePosition = 0;
			ultraGridColumn11.Hidden = true;
			ultraGridColumn12.Header.VisiblePosition = 1;
			ultraGridColumn12.Width = 176;
			ultraGridBand6.Columns.AddRange(new object[] {
																										 ultraGridColumn11,
																										 ultraGridColumn12});
			this.cmbSituacionLaboral.DisplayLayout.BandsSerializer.Add(ultraGridBand6);
			this.cmbSituacionLaboral.DisplayMember = "Descripcion";
			this.cmbSituacionLaboral.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbSituacionLaboral.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbSituacionLaboral.Enabled = false;
			this.cmbSituacionLaboral.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbSituacionLaboral.Location = new System.Drawing.Point(608, 72);
			this.cmbSituacionLaboral.Name = "cmbSituacionLaboral";
			this.cmbSituacionLaboral.Size = new System.Drawing.Size(176, 22);
			this.cmbSituacionLaboral.TabIndex = 680;
			this.cmbSituacionLaboral.ValueMember = "idSituacionLaboral";
			this.cmbSituacionLaboral.KeyDown += new System.Windows.Forms.KeyEventHandler(this.cmbSituacionLaboral_KeyDown);
			this.cmbSituacionLaboral.ValueChanged += new System.EventHandler(this.cmbSituacionLaboral_ValueChanged);
			// 
			// ultraDataSource5
			// 
			ultraDataColumn11.DataType = typeof(int);
			this.ultraDataSource5.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn11,
																																 ultraDataColumn12});
			// 
			// label56
			// 
			this.label56.AutoSize = true;
			this.label56.Location = new System.Drawing.Point(608, 48);
			this.label56.Name = "label56";
			this.label56.Size = new System.Drawing.Size(90, 17);
			this.label56.TabIndex = 678;
			this.label56.Text = "Situacion Laboral";
			this.label56.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// label55
			// 
			this.label55.AutoSize = true;
			this.label55.Location = new System.Drawing.Point(424, 48);
			this.label55.Name = "label55";
			this.label55.Size = new System.Drawing.Size(50, 17);
			this.label55.TabIndex = 676;
			this.label55.Text = "Profesión";
			this.label55.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// label45
			// 
			this.label45.AutoSize = true;
			this.label45.Location = new System.Drawing.Point(248, 48);
			this.label45.Name = "label45";
			this.label45.Size = new System.Drawing.Size(98, 17);
			this.label45.TabIndex = 674;
			this.label45.Text = "Nivel de Educación";
			this.label45.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// label44
			// 
			this.label44.AutoSize = true;
			this.label44.BackColor = System.Drawing.Color.Transparent;
			this.label44.Location = new System.Drawing.Point(352, 8);
			this.label44.Name = "label44";
			this.label44.Size = new System.Drawing.Size(28, 17);
			this.label44.TabIndex = 358;
			this.label44.Text = "Sexo";
			this.label44.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// label54
			// 
			this.label54.AutoSize = true;
			this.label54.BackColor = System.Drawing.Color.Transparent;
			this.label54.Location = new System.Drawing.Point(512, 8);
			this.label54.Name = "label54";
			this.label54.Size = new System.Drawing.Size(124, 17);
			this.label54.TabIndex = 357;
			this.label54.Text = "Provincia de Nacimiento";
			this.label54.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// label13
			// 
			this.label13.AutoSize = true;
			this.label13.BackColor = System.Drawing.Color.Transparent;
			this.label13.Location = new System.Drawing.Point(168, 48);
			this.label13.Name = "label13";
			this.label13.Size = new System.Drawing.Size(72, 17);
			this.label13.TabIndex = 161;
			this.label13.Text = "Dependientes";
			this.label13.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// cmbNEEstadoCivil
			// 
			this.cmbNEEstadoCivil.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbNEEstadoCivil.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbNEEstadoCivil.DataSource = this.ultraDataSource4;
			ultraGridBand7.AddButtonCaption = "RolSucursal";
			ultraGridColumn13.Header.VisiblePosition = 0;
			ultraGridColumn13.Hidden = true;
			ultraGridColumn14.Header.VisiblePosition = 1;
			ultraGridColumn14.Width = 144;
			ultraGridBand7.Columns.AddRange(new object[] {
																										 ultraGridColumn13,
																										 ultraGridColumn14});
			this.cmbNEEstadoCivil.DisplayLayout.BandsSerializer.Add(ultraGridBand7);
			this.cmbNEEstadoCivil.DisplayMember = "Nombre";
			this.cmbNEEstadoCivil.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbNEEstadoCivil.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbNEEstadoCivil.Enabled = false;
			this.cmbNEEstadoCivil.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbNEEstadoCivil.Location = new System.Drawing.Point(16, 72);
			this.cmbNEEstadoCivil.Name = "cmbNEEstadoCivil";
			this.cmbNEEstadoCivil.Size = new System.Drawing.Size(144, 22);
			this.cmbNEEstadoCivil.TabIndex = 160;
			this.cmbNEEstadoCivil.ValueMember = "idEdoCivil";
			this.cmbNEEstadoCivil.KeyDown += new System.Windows.Forms.KeyEventHandler(this.cmbNEEstadoCivil_KeyDown);
			this.cmbNEEstadoCivil.ValueChanged += new System.EventHandler(this.cmbNEEstadoCivil_ValueChanged);
			// 
			// ultraDataSource4
			// 
			ultraDataColumn13.DataType = typeof(int);
			this.ultraDataSource4.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn13,
																																 ultraDataColumn14});
			// 
			// label14
			// 
			this.label14.AutoSize = true;
			this.label14.BackColor = System.Drawing.Color.Transparent;
			this.label14.Location = new System.Drawing.Point(16, 48);
			this.label14.Name = "label14";
			this.label14.Size = new System.Drawing.Size(61, 17);
			this.label14.TabIndex = 159;
			this.label14.Text = "Estado Civil";
			this.label14.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// label15
			// 
			this.label15.AutoSize = true;
			this.label15.BackColor = System.Drawing.Color.Transparent;
			this.label15.Location = new System.Drawing.Point(184, 8);
			this.label15.Name = "label15";
			this.label15.Size = new System.Drawing.Size(112, 17);
			this.label15.TabIndex = 152;
			this.label15.Text = "Fecha de Nacimiento:";
			this.label15.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// dtpNEFechaNacimiento
			// 
			this.dtpNEFechaNacimiento.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			dateButton1.Caption = "Today";
			this.dtpNEFechaNacimiento.DateButtons.Add(dateButton1);
			this.dtpNEFechaNacimiento.Enabled = false;
			this.dtpNEFechaNacimiento.Format = "dd/MM/yyyy";
			this.dtpNEFechaNacimiento.Location = new System.Drawing.Point(184, 24);
			this.dtpNEFechaNacimiento.Name = "dtpNEFechaNacimiento";
			this.dtpNEFechaNacimiento.NonAutoSizeHeight = 23;
			this.dtpNEFechaNacimiento.Size = new System.Drawing.Size(112, 21);
			this.dtpNEFechaNacimiento.TabIndex = 151;
			this.dtpNEFechaNacimiento.Value = ((object)(resources.GetObject("dtpNEFechaNacimiento.Value")));
			this.dtpNEFechaNacimiento.KeyDown += new System.Windows.Forms.KeyEventHandler(this.dtpNEFechaNacimiento_KeyDown);
			// 
			// label22
			// 
			this.label22.AutoSize = true;
			this.label22.BackColor = System.Drawing.Color.Transparent;
			this.label22.Location = new System.Drawing.Point(16, 8);
			this.label22.Name = "label22";
			this.label22.Size = new System.Drawing.Size(68, 17);
			this.label22.TabIndex = 145;
			this.label22.Text = "Nacionalidad";
			this.label22.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// label21
			// 
			this.label21.AutoSize = true;
			this.label21.BackColor = System.Drawing.Color.Transparent;
			this.label21.Location = new System.Drawing.Point(672, 8);
			this.label21.Name = "label21";
			this.label21.Size = new System.Drawing.Size(114, 17);
			this.label21.TabIndex = 144;
			this.label21.Text = "Canton de Nacimiento";
			this.label21.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// groupBox1
			// 
			this.groupBox1.Controls.Add(this.btnCelular);
			this.groupBox1.Controls.Add(this.btnTelf2);
			this.groupBox1.Controls.Add(this.btnTelf);
			this.groupBox1.Controls.Add(this.cmbBarrioDomicilio);
			this.groupBox1.Controls.Add(this.cmbParroquiaDomicilio);
			this.groupBox1.Controls.Add(this.cmbCantonDomicilio);
			this.groupBox1.Controls.Add(this.cmbProvinciaDomicilio);
			this.groupBox1.Controls.Add(this.label65);
			this.groupBox1.Controls.Add(this.dtpTiempoVivienda);
			this.groupBox1.Controls.Add(this.cmbTipoVivienda);
			this.groupBox1.Controls.Add(this.label64);
			this.groupBox1.Controls.Add(this.txtReferencia);
			this.groupBox1.Controls.Add(this.txtNEEMail);
			this.groupBox1.Controls.Add(this.txtNECelular);
			this.groupBox1.Controls.Add(this.txtNETelefono2);
			this.groupBox1.Controls.Add(this.label62);
			this.groupBox1.Controls.Add(this.txtNETelefono1);
			this.groupBox1.Controls.Add(this.label19);
			this.groupBox1.Controls.Add(this.txtCalle2);
			this.groupBox1.Controls.Add(this.label61);
			this.groupBox1.Controls.Add(this.txtNumcasa);
			this.groupBox1.Controls.Add(this.label38);
			this.groupBox1.Controls.Add(this.txtCalle1);
			this.groupBox1.Controls.Add(this.label7);
			this.groupBox1.Controls.Add(this.label10);
			this.groupBox1.Controls.Add(this.label12);
			this.groupBox1.Controls.Add(this.label8);
			this.groupBox1.Controls.Add(this.label9);
			this.groupBox1.Controls.Add(this.label32);
			this.groupBox1.Controls.Add(this.label20);
			this.groupBox1.Controls.Add(this.label11);
			this.groupBox1.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.groupBox1.Location = new System.Drawing.Point(0, 104);
			this.groupBox1.Name = "groupBox1";
			this.groupBox1.Size = new System.Drawing.Size(1096, 216);
			this.groupBox1.TabIndex = 195;
			this.groupBox1.TabStop = false;
			this.groupBox1.Text = "Dirección del Domicilio";
			this.groupBox1.Enter += new System.EventHandler(this.groupBox1_Enter);
			// 
			// btnCelular
			// 
			this.btnCelular.Enabled = false;
			this.btnCelular.Image = ((System.Drawing.Image)(resources.GetObject("btnCelular.Image")));
			this.btnCelular.Location = new System.Drawing.Point(872, 80);
			this.btnCelular.Name = "btnCelular";
			this.btnCelular.Size = new System.Drawing.Size(32, 24);
			this.btnCelular.TabIndex = 691;
			this.btnCelular.Click += new System.EventHandler(this.btnCelular_Click);
			// 
			// btnTelf2
			// 
			this.btnTelf2.Enabled = false;
			this.btnTelf2.Image = ((System.Drawing.Image)(resources.GetObject("btnTelf2.Image")));
			this.btnTelf2.Location = new System.Drawing.Point(744, 80);
			this.btnTelf2.Name = "btnTelf2";
			this.btnTelf2.Size = new System.Drawing.Size(32, 24);
			this.btnTelf2.TabIndex = 690;
			this.btnTelf2.Click += new System.EventHandler(this.btnTelf2_Click);
			// 
			// btnTelf
			// 
			this.btnTelf.Enabled = false;
			this.btnTelf.Image = ((System.Drawing.Image)(resources.GetObject("btnTelf.Image")));
			this.btnTelf.Location = new System.Drawing.Point(616, 80);
			this.btnTelf.Name = "btnTelf";
			this.btnTelf.Size = new System.Drawing.Size(32, 24);
			this.btnTelf.TabIndex = 689;
			this.btnTelf.Click += new System.EventHandler(this.btnTelf_Click);
			// 
			// cmbBarrioDomicilio
			// 
			this.cmbBarrioDomicilio.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbBarrioDomicilio.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbBarrioDomicilio.DataSource = this.ultraDataSource10;
			ultraGridColumn15.Header.VisiblePosition = 0;
			ultraGridColumn15.Hidden = true;
			ultraGridColumn15.Width = 184;
			ultraGridColumn16.Header.Caption = "Provincia";
			ultraGridColumn16.Header.VisiblePosition = 1;
			ultraGridColumn16.Width = 150;
			ultraGridBand8.Columns.AddRange(new object[] {
																										 ultraGridColumn15,
																										 ultraGridColumn16});
			this.cmbBarrioDomicilio.DisplayLayout.BandsSerializer.Add(ultraGridBand8);
			this.cmbBarrioDomicilio.DisplayMember = "Nombre";
			this.cmbBarrioDomicilio.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbBarrioDomicilio.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbBarrioDomicilio.Enabled = false;
			this.cmbBarrioDomicilio.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbBarrioDomicilio.Location = new System.Drawing.Point(496, 40);
			this.cmbBarrioDomicilio.Name = "cmbBarrioDomicilio";
			this.cmbBarrioDomicilio.Size = new System.Drawing.Size(150, 22);
			this.cmbBarrioDomicilio.TabIndex = 688;
			this.cmbBarrioDomicilio.ValueMember = "idBarrio";
			this.cmbBarrioDomicilio.KeyDown += new System.Windows.Forms.KeyEventHandler(this.cmbBarrioDomicilio_KeyDown);
			// 
			// ultraDataSource10
			// 
			ultraDataColumn15.DataType = typeof(int);
			this.ultraDataSource10.Band.Columns.AddRange(new object[] {
																																	ultraDataColumn15,
																																	ultraDataColumn16});
			// 
			// cmbParroquiaDomicilio
			// 
			this.cmbParroquiaDomicilio.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbParroquiaDomicilio.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbParroquiaDomicilio.DataSource = this.ultraDataSource9;
			ultraGridColumn17.Header.VisiblePosition = 0;
			ultraGridColumn17.Hidden = true;
			ultraGridColumn18.Header.Caption = "Provincia";
			ultraGridColumn18.Header.VisiblePosition = 1;
			ultraGridColumn18.Width = 150;
			ultraGridBand9.Columns.AddRange(new object[] {
																										 ultraGridColumn17,
																										 ultraGridColumn18});
			this.cmbParroquiaDomicilio.DisplayLayout.BandsSerializer.Add(ultraGridBand9);
			this.cmbParroquiaDomicilio.DisplayMember = "Nombre";
			this.cmbParroquiaDomicilio.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbParroquiaDomicilio.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbParroquiaDomicilio.Enabled = false;
			this.cmbParroquiaDomicilio.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbParroquiaDomicilio.Location = new System.Drawing.Point(336, 40);
			this.cmbParroquiaDomicilio.Name = "cmbParroquiaDomicilio";
			this.cmbParroquiaDomicilio.Size = new System.Drawing.Size(150, 22);
			this.cmbParroquiaDomicilio.TabIndex = 687;
			this.cmbParroquiaDomicilio.ValueMember = "idParroquia";
			this.cmbParroquiaDomicilio.KeyDown += new System.Windows.Forms.KeyEventHandler(this.cmbParroquiaDomicilio_KeyDown);
			this.cmbParroquiaDomicilio.ValueChanged += new System.EventHandler(this.cmbParroquiaDomicilio_ValueChanged);
			// 
			// ultraDataSource9
			// 
			ultraDataColumn17.DataType = typeof(int);
			this.ultraDataSource9.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn17,
																																 ultraDataColumn18});
			// 
			// cmbCantonDomicilio
			// 
			this.cmbCantonDomicilio.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbCantonDomicilio.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbCantonDomicilio.DataSource = this.ultraDataSource2;
			ultraGridColumn19.Header.VisiblePosition = 0;
			ultraGridColumn19.Hidden = true;
			ultraGridColumn20.Header.Caption = "Canton";
			ultraGridColumn20.Header.VisiblePosition = 1;
			ultraGridColumn20.Width = 150;
			ultraGridBand10.Columns.AddRange(new object[] {
																											ultraGridColumn19,
																											ultraGridColumn20});
			this.cmbCantonDomicilio.DisplayLayout.BandsSerializer.Add(ultraGridBand10);
			this.cmbCantonDomicilio.DisplayLayout.Override.RowSelectors = Infragistics.Win.DefaultableBoolean.True;
			this.cmbCantonDomicilio.DisplayMember = "Nombre";
			this.cmbCantonDomicilio.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbCantonDomicilio.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbCantonDomicilio.Enabled = false;
			this.cmbCantonDomicilio.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbCantonDomicilio.Location = new System.Drawing.Point(176, 40);
			this.cmbCantonDomicilio.Name = "cmbCantonDomicilio";
			this.cmbCantonDomicilio.Size = new System.Drawing.Size(150, 22);
			this.cmbCantonDomicilio.TabIndex = 686;
			this.cmbCantonDomicilio.ValueMember = "idCanton";
			this.cmbCantonDomicilio.KeyDown += new System.Windows.Forms.KeyEventHandler(this.cmbCantonDomicilio_KeyDown);
			this.cmbCantonDomicilio.ValueChanged += new System.EventHandler(this.cmbCantonDomicilio_ValueChanged);
			// 
			// cmbProvinciaDomicilio
			// 
			this.cmbProvinciaDomicilio.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbProvinciaDomicilio.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbProvinciaDomicilio.DataSource = this.ultraDataSource1;
			ultraGridColumn21.Header.VisiblePosition = 0;
			ultraGridColumn21.Hidden = true;
			ultraGridColumn21.Width = 50;
			ultraGridColumn22.Header.VisiblePosition = 1;
			ultraGridColumn22.Width = 150;
			ultraGridBand11.Columns.AddRange(new object[] {
																											ultraGridColumn21,
																											ultraGridColumn22});
			this.cmbProvinciaDomicilio.DisplayLayout.BandsSerializer.Add(ultraGridBand11);
			this.cmbProvinciaDomicilio.DisplayMember = "Nombre";
			this.cmbProvinciaDomicilio.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbProvinciaDomicilio.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbProvinciaDomicilio.Enabled = false;
			this.cmbProvinciaDomicilio.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbProvinciaDomicilio.Location = new System.Drawing.Point(16, 40);
			this.cmbProvinciaDomicilio.Name = "cmbProvinciaDomicilio";
			this.cmbProvinciaDomicilio.Size = new System.Drawing.Size(150, 22);
			this.cmbProvinciaDomicilio.TabIndex = 685;
			this.cmbProvinciaDomicilio.ValueMember = "idProvincia";
			this.cmbProvinciaDomicilio.KeyDown += new System.Windows.Forms.KeyEventHandler(this.cmbProvinciaDomicilio_KeyDown);
			this.cmbProvinciaDomicilio.ValueChanged += new System.EventHandler(this.cmbProvinciaDomicilio_ValueChanged);
			// 
			// label65
			// 
			this.label65.AutoSize = true;
			this.label65.BackColor = System.Drawing.Color.Transparent;
			this.label65.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label65.Location = new System.Drawing.Point(176, 112);
			this.label65.Name = "label65";
			this.label65.Size = new System.Drawing.Size(97, 17);
			this.label65.TabIndex = 678;
			this.label65.Text = "Tiempo Vivienda";
			this.label65.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// dtpTiempoVivienda
			// 
			this.dtpTiempoVivienda.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			dateButton2.Caption = "Today";
			this.dtpTiempoVivienda.DateButtons.Add(dateButton2);
			this.dtpTiempoVivienda.Enabled = false;
			this.dtpTiempoVivienda.Format = "dd/MM/yyyy";
			this.dtpTiempoVivienda.Location = new System.Drawing.Point(176, 136);
			this.dtpTiempoVivienda.Name = "dtpTiempoVivienda";
			this.dtpTiempoVivienda.NonAutoSizeHeight = 23;
			this.dtpTiempoVivienda.Size = new System.Drawing.Size(112, 21);
			this.dtpTiempoVivienda.TabIndex = 677;
			this.dtpTiempoVivienda.Value = ((object)(resources.GetObject("dtpTiempoVivienda.Value")));
			this.dtpTiempoVivienda.KeyDown += new System.Windows.Forms.KeyEventHandler(this.dtpTiempoVivienda_KeyDown);
			// 
			// cmbTipoVivienda
			// 
			this.cmbTipoVivienda.CausesValidation = false;
			this.cmbTipoVivienda.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbTipoVivienda.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbTipoVivienda.DataSource = this.ultraDataSource11;
			ultraGridColumn23.Header.VisiblePosition = 0;
			ultraGridColumn23.Hidden = true;
			ultraGridColumn24.Header.VisiblePosition = 2;
			ultraGridColumn24.Width = 192;
			ultraGridColumn25.Header.VisiblePosition = 1;
			ultraGridColumn25.Hidden = true;
			ultraGridBand12.Columns.AddRange(new object[] {
																											ultraGridColumn23,
																											ultraGridColumn24,
																											ultraGridColumn25});
			this.cmbTipoVivienda.DisplayLayout.BandsSerializer.Add(ultraGridBand12);
			this.cmbTipoVivienda.DisplayMember = "Nombre";
			this.cmbTipoVivienda.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbTipoVivienda.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbTipoVivienda.Enabled = false;
			this.cmbTipoVivienda.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbTipoVivienda.Location = new System.Drawing.Point(16, 136);
			this.cmbTipoVivienda.MaxDropDownItems = 30;
			this.cmbTipoVivienda.Name = "cmbTipoVivienda";
			this.cmbTipoVivienda.Size = new System.Drawing.Size(150, 21);
			this.cmbTipoVivienda.TabIndex = 676;
			this.cmbTipoVivienda.ValueMember = "idTipoVivienda";
			this.cmbTipoVivienda.KeyDown += new System.Windows.Forms.KeyEventHandler(this.cmbTipoVivienda_KeyDown);
			// 
			// ultraDataSource11
			// 
			ultraDataColumn19.DataType = typeof(int);
			this.ultraDataSource11.Band.Columns.AddRange(new object[] {
																																	ultraDataColumn19,
																																	ultraDataColumn20});
			// 
			// label64
			// 
			this.label64.AutoSize = true;
			this.label64.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label64.Location = new System.Drawing.Point(352, 112);
			this.label64.Name = "label64";
			this.label64.Size = new System.Drawing.Size(139, 17);
			this.label64.TabIndex = 652;
			this.label64.Text = "Referencia de Ubicación";
			this.label64.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// txtReferencia
			// 
			appearance1.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtReferencia.Appearance = appearance1;
			this.txtReferencia.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtReferencia.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtReferencia.Enabled = false;
			this.txtReferencia.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.txtReferencia.Location = new System.Drawing.Point(352, 136);
			this.txtReferencia.Multiline = true;
			this.txtReferencia.Name = "txtReferencia";
			this.txtReferencia.Size = new System.Drawing.Size(720, 64);
			this.txtReferencia.TabIndex = 651;
			this.txtReferencia.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtReferencia_KeyDown);
			// 
			// txtNEEMail
			// 
			this.txtNEEMail.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtNEEMail.Enabled = false;
			this.txtNEEMail.Location = new System.Drawing.Point(912, 80);
			this.txtNEEMail.MaxLength = 60;
			this.txtNEEMail.Name = "txtNEEMail";
			this.txtNEEMail.Size = new System.Drawing.Size(160, 23);
			this.txtNEEMail.TabIndex = 392;
			this.txtNEEMail.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtNEEMail_KeyDown);
			// 
			// txtNECelular
			// 
			this.txtNECelular.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtNECelular.Enabled = false;
			this.txtNECelular.Location = new System.Drawing.Point(784, 80);
			this.txtNECelular.MaxLength = 10;
			this.txtNECelular.Name = "txtNECelular";
			this.txtNECelular.Size = new System.Drawing.Size(80, 23);
			this.txtNECelular.TabIndex = 391;
			this.txtNECelular.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtNECelular_KeyDown);
			this.txtNECelular.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.txtNECelular_KeyPress);
			// 
			// txtNETelefono2
			// 
			this.txtNETelefono2.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtNETelefono2.Enabled = false;
			this.txtNETelefono2.Location = new System.Drawing.Point(656, 80);
			this.txtNETelefono2.MaxLength = 9;
			this.txtNETelefono2.Name = "txtNETelefono2";
			this.txtNETelefono2.Size = new System.Drawing.Size(80, 23);
			this.txtNETelefono2.TabIndex = 390;
			this.txtNETelefono2.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtNETelefono2_KeyDown);
			// 
			// label62
			// 
			this.label62.AutoSize = true;
			this.label62.BackColor = System.Drawing.SystemColors.Control;
			this.label62.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label62.Location = new System.Drawing.Point(656, 64);
			this.label62.Name = "label62";
			this.label62.Size = new System.Drawing.Size(95, 17);
			this.label62.TabIndex = 389;
			this.label62.Text = "Teléfono Casa 2";
			this.label62.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// txtNETelefono1
			// 
			this.txtNETelefono1.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtNETelefono1.Enabled = false;
			this.txtNETelefono1.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.txtNETelefono1.Location = new System.Drawing.Point(528, 80);
			this.txtNETelefono1.MaxLength = 9;
			this.txtNETelefono1.Name = "txtNETelefono1";
			this.txtNETelefono1.Size = new System.Drawing.Size(80, 23);
			this.txtNETelefono1.TabIndex = 388;
			this.txtNETelefono1.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtNETelefono1_KeyDown);
			this.txtNETelefono1.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.txtNETelefono1_KeyPress);
			// 
			// label19
			// 
			this.label19.AutoSize = true;
			this.label19.BackColor = System.Drawing.SystemColors.Control;
			this.label19.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label19.Location = new System.Drawing.Point(528, 64);
			this.label19.Name = "label19";
			this.label19.Size = new System.Drawing.Size(95, 17);
			this.label19.TabIndex = 387;
			this.label19.Text = "Teléfono Casa 1";
			this.label19.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// txtCalle2
			// 
			this.txtCalle2.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtCalle2.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtCalle2.Enabled = false;
			this.txtCalle2.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.txtCalle2.Location = new System.Drawing.Point(328, 80);
			this.txtCalle2.Name = "txtCalle2";
			this.txtCalle2.Size = new System.Drawing.Size(192, 23);
			this.txtCalle2.TabIndex = 385;
			this.txtCalle2.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtCalle2_KeyDown);
			this.txtCalle2.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.txtCalle2_KeyPress);
			// 
			// label61
			// 
			this.label61.AutoSize = true;
			this.label61.BackColor = System.Drawing.SystemColors.Control;
			this.label61.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label61.Location = new System.Drawing.Point(328, 64);
			this.label61.Name = "label61";
			this.label61.Size = new System.Drawing.Size(99, 17);
			this.label61.TabIndex = 384;
			this.label61.Text = "Calle Secundaria";
			this.label61.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// txtNumcasa
			// 
			this.txtNumcasa.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtNumcasa.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtNumcasa.Enabled = false;
			this.txtNumcasa.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.txtNumcasa.Location = new System.Drawing.Point(216, 80);
			this.txtNumcasa.Name = "txtNumcasa";
			this.txtNumcasa.Size = new System.Drawing.Size(104, 23);
			this.txtNumcasa.TabIndex = 383;
			this.txtNumcasa.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtNumcasa_KeyDown);
			this.txtNumcasa.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.txtNumcasa_KeyPress);
			// 
			// label38
			// 
			this.label38.AutoSize = true;
			this.label38.BackColor = System.Drawing.SystemColors.Control;
			this.label38.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label38.Location = new System.Drawing.Point(216, 64);
			this.label38.Name = "label38";
			this.label38.Size = new System.Drawing.Size(98, 17);
			this.label38.TabIndex = 382;
			this.label38.Text = "Número de Casa";
			this.label38.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// txtCalle1
			// 
			this.txtCalle1.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtCalle1.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtCalle1.Enabled = false;
			this.txtCalle1.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.txtCalle1.Location = new System.Drawing.Point(16, 80);
			this.txtCalle1.Name = "txtCalle1";
			this.txtCalle1.Size = new System.Drawing.Size(192, 23);
			this.txtCalle1.TabIndex = 381;
			this.txtCalle1.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtCalle1_KeyDown);
			this.txtCalle1.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.txtCalle1_KeyPress);
			// 
			// label7
			// 
			this.label7.AutoSize = true;
			this.label7.BackColor = System.Drawing.SystemColors.Control;
			this.label7.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label7.Location = new System.Drawing.Point(16, 64);
			this.label7.Name = "label7";
			this.label7.Size = new System.Drawing.Size(84, 17);
			this.label7.TabIndex = 380;
			this.label7.Text = "Calle Principal";
			this.label7.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// label10
			// 
			this.label10.AutoSize = true;
			this.label10.BackColor = System.Drawing.SystemColors.Control;
			this.label10.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label10.Location = new System.Drawing.Point(16, 24);
			this.label10.Name = "label10";
			this.label10.Size = new System.Drawing.Size(55, 17);
			this.label10.TabIndex = 164;
			this.label10.Text = "Provincia";
			this.label10.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// label12
			// 
			this.label12.AutoSize = true;
			this.label12.BackColor = System.Drawing.SystemColors.Control;
			this.label12.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label12.Location = new System.Drawing.Point(176, 24);
			this.label12.Name = "label12";
			this.label12.Size = new System.Drawing.Size(44, 17);
			this.label12.TabIndex = 163;
			this.label12.Text = "Canton";
			this.label12.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// label8
			// 
			this.label8.AutoSize = true;
			this.label8.BackColor = System.Drawing.SystemColors.Control;
			this.label8.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label8.Location = new System.Drawing.Point(336, 24);
			this.label8.Name = "label8";
			this.label8.Size = new System.Drawing.Size(58, 17);
			this.label8.TabIndex = 168;
			this.label8.Text = "Parroquia";
			this.label8.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// label9
			// 
			this.label9.AutoSize = true;
			this.label9.BackColor = System.Drawing.SystemColors.Control;
			this.label9.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label9.Location = new System.Drawing.Point(496, 24);
			this.label9.Name = "label9";
			this.label9.Size = new System.Drawing.Size(38, 17);
			this.label9.TabIndex = 167;
			this.label9.Text = "Barrio";
			this.label9.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// label32
			// 
			this.label32.AutoSize = true;
			this.label32.BackColor = System.Drawing.SystemColors.Control;
			this.label32.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label32.Location = new System.Drawing.Point(912, 64);
			this.label32.Name = "label32";
			this.label32.Size = new System.Drawing.Size(43, 17);
			this.label32.TabIndex = 178;
			this.label32.Text = "E-Mail ";
			this.label32.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// label20
			// 
			this.label20.AutoSize = true;
			this.label20.BackColor = System.Drawing.SystemColors.Control;
			this.label20.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label20.Location = new System.Drawing.Point(784, 64);
			this.label20.Name = "label20";
			this.label20.Size = new System.Drawing.Size(44, 17);
			this.label20.TabIndex = 175;
			this.label20.Text = "Celular";
			this.label20.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// label11
			// 
			this.label11.AutoSize = true;
			this.label11.BackColor = System.Drawing.SystemColors.Control;
			this.label11.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label11.Location = new System.Drawing.Point(16, 112);
			this.label11.Name = "label11";
			this.label11.Size = new System.Drawing.Size(96, 17);
			this.label11.TabIndex = 182;
			this.label11.Text = "Tipo de Vivienda";
			this.label11.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// cmbProfesion
			// 
			this.cmbProfesion.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbProfesion.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbProfesion.DataSource = this.ultraDataSource3;
			ultraGridColumn26.Header.VisiblePosition = 0;
			ultraGridColumn26.Hidden = true;
			ultraGridColumn27.Header.VisiblePosition = 1;
			ultraGridColumn27.Width = 176;
			ultraGridBand13.Columns.AddRange(new object[] {
																											ultraGridColumn26,
																											ultraGridColumn27});
			this.cmbProfesion.DisplayLayout.BandsSerializer.Add(ultraGridBand13);
			this.cmbProfesion.DisplayLayout.Override.RowSelectors = Infragistics.Win.DefaultableBoolean.True;
			this.cmbProfesion.DisplayMember = "Nombre";
			this.cmbProfesion.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbProfesion.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbProfesion.Enabled = false;
			this.cmbProfesion.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbProfesion.Location = new System.Drawing.Point(424, 72);
			this.cmbProfesion.Name = "cmbProfesion";
			this.cmbProfesion.Size = new System.Drawing.Size(176, 22);
			this.cmbProfesion.TabIndex = 806;
			this.cmbProfesion.ValueMember = "idProfesion";
			this.cmbProfesion.KeyDown += new System.Windows.Forms.KeyEventHandler(this.cmbProfesion_KeyDown);
			// 
			// ultraDataSource3
			// 
			ultraDataColumn21.DataType = typeof(int);
			this.ultraDataSource3.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn21,
																																 ultraDataColumn22});
			// 
			// txtNECargasFamiliares
			// 
			this.txtNECargasFamiliares.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtNECargasFamiliares.Enabled = false;
			this.txtNECargasFamiliares.FormatString = "#,##0";
			this.txtNECargasFamiliares.Location = new System.Drawing.Point(168, 72);
			this.txtNECargasFamiliares.MaskInput = "nn";
			this.txtNECargasFamiliares.MaxValue = 99;
			this.txtNECargasFamiliares.MinValue = 0;
			this.txtNECargasFamiliares.Name = "txtNECargasFamiliares";
			this.txtNECargasFamiliares.PromptChar = ' ';
			this.txtNECargasFamiliares.Size = new System.Drawing.Size(72, 22);
			this.txtNECargasFamiliares.SpinButtonDisplayStyle = Infragistics.Win.ButtonDisplayStyle.Always;
			this.txtNECargasFamiliares.TabIndex = 806;
			this.txtNECargasFamiliares.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtNECargasFamiliares_KeyDown);
			this.txtNECargasFamiliares.Click += new System.EventHandler(this.txtNECargasFamiliares_Click);
			this.txtNECargasFamiliares.Enter += new System.EventHandler(this.txtNECargasFamiliares_Enter);
			// 
			// cmbActividadEcomica
			// 
			this.cmbActividadEcomica.CausesValidation = false;
			this.cmbActividadEcomica.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbActividadEcomica.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbActividadEcomica.DataSource = this.ultraDataSource8;
			ultraGridColumn28.Header.VisiblePosition = 0;
			ultraGridColumn28.Hidden = true;
			ultraGridColumn28.Width = 304;
			ultraGridColumn29.Header.VisiblePosition = 1;
			ultraGridColumn29.Width = 304;
			ultraGridBand14.Columns.AddRange(new object[] {
																											ultraGridColumn28,
																											ultraGridColumn29});
			this.cmbActividadEcomica.DisplayLayout.BandsSerializer.Add(ultraGridBand14);
			this.cmbActividadEcomica.DisplayMember = "Nombre";
			this.cmbActividadEcomica.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbActividadEcomica.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbActividadEcomica.Enabled = false;
			this.cmbActividadEcomica.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbActividadEcomica.Location = new System.Drawing.Point(800, 72);
			this.cmbActividadEcomica.MaxDropDownItems = 30;
			this.cmbActividadEcomica.Name = "cmbActividadEcomica";
			this.cmbActividadEcomica.Size = new System.Drawing.Size(304, 23);
			this.cmbActividadEcomica.TabIndex = 809;
			this.cmbActividadEcomica.ValueMember = "idActEconomica";
			this.cmbActividadEcomica.KeyDown += new System.Windows.Forms.KeyEventHandler(this.cmbActividadEcomica_KeyDown_1);
			// 
			// ultraDataSource8
			// 
			ultraDataColumn23.DataType = typeof(int);
			this.ultraDataSource8.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn23,
																																 ultraDataColumn24});
			// 
			// utpEstudios
			// 
			this.utpEstudios.Controls.Add(this.groupBox3);
			this.utpEstudios.Location = new System.Drawing.Point(-10000, -10000);
			this.utpEstudios.Name = "utpEstudios";
			this.utpEstudios.Size = new System.Drawing.Size(1102, 323);
			// 
			// groupBox3
			// 
			this.groupBox3.Controls.Add(this.cmbCProfesion);
			this.groupBox3.Controls.Add(this.cmbCSexo);
			this.groupBox3.Controls.Add(this.label24);
			this.groupBox3.Controls.Add(this.cmbCTipoRuc);
			this.groupBox3.Controls.Add(this.cmbCNivel);
			this.groupBox3.Controls.Add(this.cmbCNacionalidad);
			this.groupBox3.Controls.Add(this.cmbCActividadEconomia);
			this.groupBox3.Controls.Add(this.label75);
			this.groupBox3.Controls.Add(this.cmbCSituacionLaboral);
			this.groupBox3.Controls.Add(this.label76);
			this.groupBox3.Controls.Add(this.label77);
			this.groupBox3.Controls.Add(this.label78);
			this.groupBox3.Controls.Add(this.cmbCProvinciaNacimiento);
			this.groupBox3.Controls.Add(this.label73);
			this.groupBox3.Controls.Add(this.cmbCCantonNacimiento);
			this.groupBox3.Controls.Add(this.label74);
			this.groupBox3.Controls.Add(this.label70);
			this.groupBox3.Controls.Add(this.dpCFechaNacimiento);
			this.groupBox3.Controls.Add(this.label71);
			this.groupBox3.Controls.Add(this.txtCSNombre);
			this.groupBox3.Controls.Add(this.txtCPNombre);
			this.groupBox3.Controls.Add(this.txtCApellidoM);
			this.groupBox3.Controls.Add(this.txtCApellidoP);
			this.groupBox3.Controls.Add(this.label51);
			this.groupBox3.Controls.Add(this.label52);
			this.groupBox3.Controls.Add(this.label63);
			this.groupBox3.Controls.Add(this.label69);
			this.groupBox3.Controls.Add(this.txtCDocumento);
			this.groupBox3.Controls.Add(this.label50);
			this.groupBox3.Controls.Add(this.label49);
			this.groupBox3.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.groupBox3.Location = new System.Drawing.Point(8, 8);
			this.groupBox3.Name = "groupBox3";
			this.groupBox3.Size = new System.Drawing.Size(1088, 304);
			this.groupBox3.TabIndex = 197;
			this.groupBox3.TabStop = false;
			this.groupBox3.Text = "Datos Personales";
			this.groupBox3.Enter += new System.EventHandler(this.groupBox3_Enter);
			// 
			// cmbCProfesion
			// 
			this.cmbCProfesion.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbCProfesion.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbCProfesion.DataSource = this.ultraDataSource3;
			ultraGridColumn30.Header.VisiblePosition = 0;
			ultraGridColumn30.Hidden = true;
			ultraGridColumn31.Header.VisiblePosition = 1;
			ultraGridColumn31.Width = 176;
			ultraGridBand15.Columns.AddRange(new object[] {
																											ultraGridColumn30,
																											ultraGridColumn31});
			this.cmbCProfesion.DisplayLayout.BandsSerializer.Add(ultraGridBand15);
			this.cmbCProfesion.DisplayMember = "Nombre";
			this.cmbCProfesion.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbCProfesion.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbCProfesion.Enabled = false;
			this.cmbCProfesion.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbCProfesion.Location = new System.Drawing.Point(192, 136);
			this.cmbCProfesion.Name = "cmbCProfesion";
			this.cmbCProfesion.Size = new System.Drawing.Size(176, 23);
			this.cmbCProfesion.TabIndex = 811;
			this.cmbCProfesion.ValueMember = "idProfesion";
			this.cmbCProfesion.KeyDown += new System.Windows.Forms.KeyEventHandler(this.cmbCProfesion_KeyDown_1);
			// 
			// cmbCSexo
			// 
			this.cmbCSexo.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbCSexo.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbCSexo.DataSource = this.ultraDataSource17;
			ultraGridBand16.AddButtonCaption = "RolSucursal";
			ultraGridColumn32.Header.VisiblePosition = 0;
			ultraGridColumn32.Hidden = true;
			ultraGridColumn33.Header.VisiblePosition = 1;
			ultraGridColumn33.Width = 144;
			ultraGridBand16.Columns.AddRange(new object[] {
																											ultraGridColumn32,
																											ultraGridColumn33});
			this.cmbCSexo.DisplayLayout.BandsSerializer.Add(ultraGridBand16);
			this.cmbCSexo.DisplayMember = "Nombre";
			this.cmbCSexo.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbCSexo.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbCSexo.Enabled = false;
			this.cmbCSexo.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbCSexo.Location = new System.Drawing.Point(328, 80);
			this.cmbCSexo.Name = "cmbCSexo";
			this.cmbCSexo.Size = new System.Drawing.Size(152, 22);
			this.cmbCSexo.TabIndex = 810;
			this.cmbCSexo.ValueMember = "idSexo";
			this.cmbCSexo.KeyDown += new System.Windows.Forms.KeyEventHandler(this.cmbCSexo_KeyDown);
			// 
			// label24
			// 
			this.label24.AutoSize = true;
			this.label24.BackColor = System.Drawing.Color.Transparent;
			this.label24.Location = new System.Drawing.Point(328, 64);
			this.label24.Name = "label24";
			this.label24.Size = new System.Drawing.Size(33, 17);
			this.label24.TabIndex = 809;
			this.label24.Text = "Sexo";
			this.label24.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// cmbCTipoRuc
			// 
			this.cmbCTipoRuc.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbCTipoRuc.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbCTipoRuc.DisplayMember = "Nombre";
			this.cmbCTipoRuc.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbCTipoRuc.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbCTipoRuc.Enabled = false;
			this.cmbCTipoRuc.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbCTipoRuc.Location = new System.Drawing.Point(8, 40);
			this.cmbCTipoRuc.Name = "cmbCTipoRuc";
			this.cmbCTipoRuc.Size = new System.Drawing.Size(152, 22);
			this.cmbCTipoRuc.TabIndex = 808;
			this.cmbCTipoRuc.ValueMember = "idTipoDoc";
			// 
			// cmbCNivel
			// 
			this.cmbCNivel.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbCNivel.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbCNivel.DataSource = this.ultraDataSource19;
			ultraGridColumn34.Header.VisiblePosition = 0;
			ultraGridColumn34.Hidden = true;
			ultraGridColumn35.Header.VisiblePosition = 1;
			ultraGridColumn35.Width = 168;
			ultraGridBand17.Columns.AddRange(new object[] {
																											ultraGridColumn34,
																											ultraGridColumn35});
			this.cmbCNivel.DisplayLayout.BandsSerializer.Add(ultraGridBand17);
			this.cmbCNivel.DisplayLayout.Override.RowSelectors = Infragistics.Win.DefaultableBoolean.True;
			this.cmbCNivel.DisplayMember = "Nombre";
			this.cmbCNivel.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbCNivel.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbCNivel.Enabled = false;
			this.cmbCNivel.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbCNivel.Location = new System.Drawing.Point(8, 136);
			this.cmbCNivel.Name = "cmbCNivel";
			this.cmbCNivel.Size = new System.Drawing.Size(168, 22);
			this.cmbCNivel.TabIndex = 807;
			this.cmbCNivel.ValueMember = "idNivelEducacion";
			this.cmbCNivel.KeyDown += new System.Windows.Forms.KeyEventHandler(this.cmbCNivel_KeyDown);
			// 
			// cmbCNacionalidad
			// 
			this.cmbCNacionalidad.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.cmbCNacionalidad.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbCNacionalidad.DataSource = this.ultraDataSource26;
			ultraGridBand18.AddButtonCaption = "RolSucursal";
			ultraGridColumn36.Header.VisiblePosition = 0;
			ultraGridColumn36.Hidden = true;
			ultraGridColumn37.Header.VisiblePosition = 1;
			ultraGridColumn37.Width = 152;
			ultraGridBand18.Columns.AddRange(new object[] {
																											ultraGridColumn36,
																											ultraGridColumn37});
			this.cmbCNacionalidad.DisplayLayout.BandsSerializer.Add(ultraGridBand18);
			this.cmbCNacionalidad.DisplayMember = "Nombre";
			this.cmbCNacionalidad.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbCNacionalidad.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbCNacionalidad.Enabled = false;
			this.cmbCNacionalidad.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbCNacionalidad.Location = new System.Drawing.Point(8, 80);
			this.cmbCNacionalidad.Name = "cmbCNacionalidad";
			this.cmbCNacionalidad.Size = new System.Drawing.Size(152, 22);
			this.cmbCNacionalidad.TabIndex = 700;
			this.cmbCNacionalidad.ValueMember = "idNacionalidad";
			this.cmbCNacionalidad.KeyDown += new System.Windows.Forms.KeyEventHandler(this.cmbCNacionalidad_KeyDown);
			// 
			// cmbCActividadEconomia
			// 
			this.cmbCActividadEconomia.CausesValidation = false;
			this.cmbCActividadEconomia.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbCActividadEconomia.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbCActividadEconomia.DataSource = this.ultraDataSource8;
			ultraGridColumn38.Header.VisiblePosition = 0;
			ultraGridColumn38.Hidden = true;
			ultraGridColumn38.Width = 376;
			ultraGridColumn39.Header.VisiblePosition = 1;
			ultraGridColumn39.Width = 376;
			ultraGridBand19.Columns.AddRange(new object[] {
																											ultraGridColumn38,
																											ultraGridColumn39});
			this.cmbCActividadEconomia.DisplayLayout.BandsSerializer.Add(ultraGridBand19);
			this.cmbCActividadEconomia.DisplayMember = "Nombre";
			this.cmbCActividadEconomia.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbCActividadEconomia.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbCActividadEconomia.Enabled = false;
			this.cmbCActividadEconomia.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbCActividadEconomia.Location = new System.Drawing.Point(584, 136);
			this.cmbCActividadEconomia.MaxDropDownItems = 30;
			this.cmbCActividadEconomia.Name = "cmbCActividadEconomia";
			this.cmbCActividadEconomia.Size = new System.Drawing.Size(376, 23);
			this.cmbCActividadEconomia.TabIndex = 699;
			this.cmbCActividadEconomia.ValueMember = "idActEconomica";
			this.cmbCActividadEconomia.KeyDown += new System.Windows.Forms.KeyEventHandler(this.cmbCActividadEconomia_KeyDown);
			// 
			// label75
			// 
			this.label75.AutoSize = true;
			this.label75.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label75.Location = new System.Drawing.Point(584, 120);
			this.label75.Name = "label75";
			this.label75.Size = new System.Drawing.Size(119, 17);
			this.label75.TabIndex = 697;
			this.label75.Text = "Actividad Económica";
			this.label75.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// cmbCSituacionLaboral
			// 
			this.cmbCSituacionLaboral.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbCSituacionLaboral.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbCSituacionLaboral.DataSource = this.ultraDataSource5;
			ultraGridColumn40.Header.VisiblePosition = 0;
			ultraGridColumn40.Hidden = true;
			ultraGridColumn41.Header.VisiblePosition = 1;
			ultraGridColumn41.Width = 200;
			ultraGridBand20.Columns.AddRange(new object[] {
																											ultraGridColumn40,
																											ultraGridColumn41});
			this.cmbCSituacionLaboral.DisplayLayout.BandsSerializer.Add(ultraGridBand20);
			this.cmbCSituacionLaboral.DisplayMember = "Descripcion";
			this.cmbCSituacionLaboral.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbCSituacionLaboral.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbCSituacionLaboral.Enabled = false;
			this.cmbCSituacionLaboral.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbCSituacionLaboral.Location = new System.Drawing.Point(376, 136);
			this.cmbCSituacionLaboral.Name = "cmbCSituacionLaboral";
			this.cmbCSituacionLaboral.Size = new System.Drawing.Size(200, 23);
			this.cmbCSituacionLaboral.TabIndex = 696;
			this.cmbCSituacionLaboral.ValueMember = "idSituacionLaboral";
			this.cmbCSituacionLaboral.KeyDown += new System.Windows.Forms.KeyEventHandler(this.cmbCSituacionLaboral_KeyDown);
			this.cmbCSituacionLaboral.ValueChanged += new System.EventHandler(this.cmbCSituacionLaboral_ValueChanged);
			// 
			// label76
			// 
			this.label76.AutoSize = true;
			this.label76.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label76.Location = new System.Drawing.Point(376, 120);
			this.label76.Name = "label76";
			this.label76.TabIndex = 694;
			this.label76.Text = "Situacion Laboral";
			this.label76.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// label77
			// 
			this.label77.AutoSize = true;
			this.label77.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label77.Location = new System.Drawing.Point(192, 120);
			this.label77.Name = "label77";
			this.label77.Size = new System.Drawing.Size(57, 17);
			this.label77.TabIndex = 693;
			this.label77.Text = "Profesión";
			this.label77.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// label78
			// 
			this.label78.AutoSize = true;
			this.label78.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label78.Location = new System.Drawing.Point(8, 120);
			this.label78.Name = "label78";
			this.label78.Size = new System.Drawing.Size(110, 17);
			this.label78.TabIndex = 691;
			this.label78.Text = "Nivel de Educación";
			this.label78.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// cmbCProvinciaNacimiento
			// 
			this.cmbCProvinciaNacimiento.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbCProvinciaNacimiento.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbCProvinciaNacimiento.DataSource = this.ultraDataSource1;
			ultraGridBand21.AddButtonCaption = "RolSucursal";
			ultraGridColumn42.Header.VisiblePosition = 0;
			ultraGridColumn42.Hidden = true;
			ultraGridColumn43.Header.VisiblePosition = 1;
			ultraGridColumn43.Width = 150;
			ultraGridBand21.Columns.AddRange(new object[] {
																											ultraGridColumn42,
																											ultraGridColumn43});
			this.cmbCProvinciaNacimiento.DisplayLayout.BandsSerializer.Add(ultraGridBand21);
			this.cmbCProvinciaNacimiento.DisplayMember = "Nombre";
			this.cmbCProvinciaNacimiento.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbCProvinciaNacimiento.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbCProvinciaNacimiento.Enabled = false;
			this.cmbCProvinciaNacimiento.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbCProvinciaNacimiento.Location = new System.Drawing.Point(488, 80);
			this.cmbCProvinciaNacimiento.Name = "cmbCProvinciaNacimiento";
			this.cmbCProvinciaNacimiento.Size = new System.Drawing.Size(150, 22);
			this.cmbCProvinciaNacimiento.TabIndex = 396;
			this.cmbCProvinciaNacimiento.ValueMember = "idProvincia";
			this.cmbCProvinciaNacimiento.KeyDown += new System.Windows.Forms.KeyEventHandler(this.cmbCProvinciaNacimiento_KeyDown);
			this.cmbCProvinciaNacimiento.ValueChanged += new System.EventHandler(this.cmbCProvinciaNacimiento_ValueChanged);
			// 
			// label73
			// 
			this.label73.AutoSize = true;
			this.label73.BackColor = System.Drawing.Color.Transparent;
			this.label73.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label73.Location = new System.Drawing.Point(488, 64);
			this.label73.Name = "label73";
			this.label73.Size = new System.Drawing.Size(138, 17);
			this.label73.TabIndex = 394;
			this.label73.Text = "Provincia de Nacimiento";
			this.label73.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// cmbCCantonNacimiento
			// 
			this.cmbCCantonNacimiento.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbCCantonNacimiento.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbCCantonNacimiento.DataSource = this.ultraDataSource2;
			ultraGridBand22.AddButtonCaption = "RolSucursal";
			ultraGridColumn44.Header.VisiblePosition = 0;
			ultraGridColumn44.Hidden = true;
			ultraGridColumn45.Header.VisiblePosition = 1;
			ultraGridColumn45.Width = 150;
			ultraGridBand22.Columns.AddRange(new object[] {
																											ultraGridColumn44,
																											ultraGridColumn45});
			this.cmbCCantonNacimiento.DisplayLayout.BandsSerializer.Add(ultraGridBand22);
			this.cmbCCantonNacimiento.DisplayLayout.Override.RowSelectors = Infragistics.Win.DefaultableBoolean.True;
			this.cmbCCantonNacimiento.DisplayMember = "Nombre";
			this.cmbCCantonNacimiento.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbCCantonNacimiento.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbCCantonNacimiento.Enabled = false;
			this.cmbCCantonNacimiento.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbCCantonNacimiento.Location = new System.Drawing.Point(648, 80);
			this.cmbCCantonNacimiento.Name = "cmbCCantonNacimiento";
			this.cmbCCantonNacimiento.Size = new System.Drawing.Size(150, 22);
			this.cmbCCantonNacimiento.TabIndex = 391;
			this.cmbCCantonNacimiento.ValueMember = "idCanton";
			this.cmbCCantonNacimiento.KeyDown += new System.Windows.Forms.KeyEventHandler(this.cmbCCantonNacimiento_KeyDown);
			// 
			// label74
			// 
			this.label74.AutoSize = true;
			this.label74.BackColor = System.Drawing.Color.Transparent;
			this.label74.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label74.Location = new System.Drawing.Point(648, 64);
			this.label74.Name = "label74";
			this.label74.Size = new System.Drawing.Size(127, 17);
			this.label74.TabIndex = 390;
			this.label74.Text = "Canton de Nacimiento";
			this.label74.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// label70
			// 
			this.label70.AutoSize = true;
			this.label70.BackColor = System.Drawing.Color.Transparent;
			this.label70.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label70.Location = new System.Drawing.Point(168, 64);
			this.label70.Name = "label70";
			this.label70.Size = new System.Drawing.Size(125, 17);
			this.label70.TabIndex = 388;
			this.label70.Text = "Fecha de Nacimiento:";
			this.label70.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// dpCFechaNacimiento
			// 
			this.dpCFechaNacimiento.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			dateButton3.Caption = "Today";
			this.dpCFechaNacimiento.DateButtons.Add(dateButton3);
			this.dpCFechaNacimiento.Enabled = false;
			this.dpCFechaNacimiento.Format = "dd/MM/yyyy";
			this.dpCFechaNacimiento.Location = new System.Drawing.Point(168, 80);
			this.dpCFechaNacimiento.Name = "dpCFechaNacimiento";
			this.dpCFechaNacimiento.NonAutoSizeHeight = 23;
			this.dpCFechaNacimiento.Size = new System.Drawing.Size(104, 21);
			this.dpCFechaNacimiento.TabIndex = 387;
			this.dpCFechaNacimiento.Value = ((object)(resources.GetObject("dpCFechaNacimiento.Value")));
			this.dpCFechaNacimiento.KeyDown += new System.Windows.Forms.KeyEventHandler(this.dpCFechaNacimiento_KeyDown);
			// 
			// label71
			// 
			this.label71.AutoSize = true;
			this.label71.BackColor = System.Drawing.Color.Transparent;
			this.label71.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label71.Location = new System.Drawing.Point(8, 64);
			this.label71.Name = "label71";
			this.label71.Size = new System.Drawing.Size(76, 17);
			this.label71.TabIndex = 386;
			this.label71.Text = "Nacionalidad";
			this.label71.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// txtCSNombre
			// 
			appearance2.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtCSNombre.Appearance = appearance2;
			this.txtCSNombre.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtCSNombre.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtCSNombre.Enabled = false;
			this.txtCSNombre.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.txtCSNombre.Location = new System.Drawing.Point(808, 40);
			this.txtCSNombre.Name = "txtCSNombre";
			this.txtCSNombre.Size = new System.Drawing.Size(150, 23);
			this.txtCSNombre.TabIndex = 385;
			this.txtCSNombre.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtCSNombre_KeyDown);
			this.txtCSNombre.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.txtCSNombre_KeyPress);
			// 
			// txtCPNombre
			// 
			appearance3.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtCPNombre.Appearance = appearance3;
			this.txtCPNombre.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtCPNombre.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtCPNombre.Enabled = false;
			this.txtCPNombre.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.txtCPNombre.Location = new System.Drawing.Point(648, 40);
			this.txtCPNombre.Name = "txtCPNombre";
			this.txtCPNombre.Size = new System.Drawing.Size(150, 23);
			this.txtCPNombre.TabIndex = 384;
			this.txtCPNombre.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtCPNombre_KeyDown);
			this.txtCPNombre.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.txtCPNombre_KeyPress);
			// 
			// txtCApellidoM
			// 
			appearance4.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtCApellidoM.Appearance = appearance4;
			this.txtCApellidoM.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtCApellidoM.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtCApellidoM.Enabled = false;
			this.txtCApellidoM.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.txtCApellidoM.Location = new System.Drawing.Point(488, 40);
			this.txtCApellidoM.Name = "txtCApellidoM";
			this.txtCApellidoM.Size = new System.Drawing.Size(150, 23);
			this.txtCApellidoM.TabIndex = 383;
			this.txtCApellidoM.Visible = false;
			this.txtCApellidoM.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtCApellidoM_KeyDown);
			this.txtCApellidoM.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.txtCApellidoM_KeyPress);
			// 
			// txtCApellidoP
			// 
			appearance5.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtCApellidoP.Appearance = appearance5;
			this.txtCApellidoP.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtCApellidoP.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtCApellidoP.Enabled = false;
			this.txtCApellidoP.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.txtCApellidoP.Location = new System.Drawing.Point(328, 40);
			this.txtCApellidoP.Name = "txtCApellidoP";
			this.txtCApellidoP.Size = new System.Drawing.Size(150, 23);
			this.txtCApellidoP.TabIndex = 382;
			this.txtCApellidoP.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtCApellidoP_KeyDown);
			this.txtCApellidoP.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.txtCApellidoP_KeyPress);
			// 
			// label51
			// 
			this.label51.AutoSize = true;
			this.label51.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label51.Location = new System.Drawing.Point(808, 24);
			this.label51.Name = "label51";
			this.label51.Size = new System.Drawing.Size(101, 17);
			this.label51.TabIndex = 381;
			this.label51.Text = "Segundo Nombre";
			this.label51.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// label52
			// 
			this.label52.AutoSize = true;
			this.label52.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label52.Location = new System.Drawing.Point(648, 24);
			this.label52.Name = "label52";
			this.label52.Size = new System.Drawing.Size(88, 17);
			this.label52.TabIndex = 380;
			this.label52.Text = "Primer Nombre";
			this.label52.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// label63
			// 
			this.label63.AutoSize = true;
			this.label63.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label63.Location = new System.Drawing.Point(488, 24);
			this.label63.Name = "label63";
			this.label63.Size = new System.Drawing.Size(97, 17);
			this.label63.TabIndex = 379;
			this.label63.Text = "Apellido Materno";
			this.label63.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// label69
			// 
			this.label69.AutoSize = true;
			this.label69.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label69.Location = new System.Drawing.Point(328, 24);
			this.label69.Name = "label69";
			this.label69.Size = new System.Drawing.Size(95, 17);
			this.label69.TabIndex = 378;
			this.label69.Text = "Apellido Paterno";
			this.label69.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// txtCDocumento
			// 
			appearance6.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtCDocumento.Appearance = appearance6;
			this.txtCDocumento.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtCDocumento.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtCDocumento.Enabled = false;
			this.txtCDocumento.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.txtCDocumento.Location = new System.Drawing.Point(168, 40);
			this.txtCDocumento.Name = "txtCDocumento";
			this.txtCDocumento.Size = new System.Drawing.Size(150, 23);
			this.txtCDocumento.TabIndex = 362;
			this.txtCDocumento.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtCDocumento_KeyDown);
			this.txtCDocumento.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.txtCDocumento_KeyPress);
			this.txtCDocumento.Validating += new System.ComponentModel.CancelEventHandler(this.txtCDocumento_Validating);
			// 
			// label50
			// 
			this.label50.AutoSize = true;
			this.label50.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label50.Location = new System.Drawing.Point(168, 24);
			this.label50.Name = "label50";
			this.label50.Size = new System.Drawing.Size(68, 17);
			this.label50.TabIndex = 361;
			this.label50.Text = "Documento";
			this.label50.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// label49
			// 
			this.label49.AutoSize = true;
			this.label49.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label49.Location = new System.Drawing.Point(8, 24);
			this.label49.Name = "label49";
			this.label49.Size = new System.Drawing.Size(95, 17);
			this.label49.TabIndex = 359;
			this.label49.Text = "Tipo Documento";
			this.label49.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// utpAdicional
			// 
			this.utpAdicional.Controls.Add(this.uGridParentesco);
			this.utpAdicional.Controls.Add(this.cmbRProvincia);
			this.utpAdicional.Controls.Add(this.cmbRCanton);
			this.utpAdicional.Controls.Add(this.cmbParentesco);
			this.utpAdicional.Location = new System.Drawing.Point(1, 20);
			this.utpAdicional.Name = "utpAdicional";
			this.utpAdicional.Size = new System.Drawing.Size(1102, 323);
			// 
			// uGridParentesco
			// 
			this.uGridParentesco.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.uGridParentesco.Cursor = System.Windows.Forms.Cursors.Default;
			this.uGridParentesco.DataSource = this.ultraDataSource20;
			appearance7.BackColor = System.Drawing.Color.White;
			this.uGridParentesco.DisplayLayout.Appearance = appearance7;
			ultraGridColumn46.Header.VisiblePosition = 0;
			ultraGridColumn46.Hidden = true;
			ultraGridColumn47.Header.VisiblePosition = 1;
			ultraGridColumn47.Hidden = true;
			ultraGridColumn48.Header.VisiblePosition = 2;
			ultraGridColumn48.Hidden = true;
			ultraGridColumn49.Header.Caption = "Parentesco";
			ultraGridColumn49.Header.VisiblePosition = 3;
			ultraGridColumn49.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.DropDownList;
			ultraGridColumn49.Width = 150;
			ultraGridColumn50.Case = Infragistics.Win.UltraWinGrid.Case.Upper;
			ultraGridColumn50.Header.Caption = "Apellido Paterno";
			ultraGridColumn50.Header.VisiblePosition = 4;
			ultraGridColumn50.Width = 120;
			ultraGridColumn51.Case = Infragistics.Win.UltraWinGrid.Case.Upper;
			ultraGridColumn51.Header.Caption = "Apellido Materno";
			ultraGridColumn51.Header.VisiblePosition = 5;
			ultraGridColumn51.Hidden = true;
			ultraGridColumn51.Width = 120;
			ultraGridColumn52.Case = Infragistics.Win.UltraWinGrid.Case.Upper;
			ultraGridColumn52.Header.Caption = "Primer Nombre";
			ultraGridColumn52.Header.VisiblePosition = 6;
			ultraGridColumn52.Width = 120;
			ultraGridColumn53.Case = Infragistics.Win.UltraWinGrid.Case.Upper;
			ultraGridColumn53.Header.Caption = "Segundo Nombre";
			ultraGridColumn53.Header.VisiblePosition = 7;
			ultraGridColumn53.Width = 120;
			ultraGridColumn54.Header.Caption = "Provincia";
			ultraGridColumn54.Header.VisiblePosition = 8;
			ultraGridColumn54.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.DropDownList;
			ultraGridColumn54.Width = 150;
			ultraGridColumn55.Header.Caption = "Canton";
			ultraGridColumn55.Header.VisiblePosition = 9;
			ultraGridColumn55.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.DropDownList;
			ultraGridColumn55.Width = 150;
			ultraGridColumn56.Case = Infragistics.Win.UltraWinGrid.Case.Upper;
			ultraGridColumn56.Header.VisiblePosition = 10;
			ultraGridColumn56.Hidden = true;
			ultraGridColumn56.Width = 150;
			ultraGridColumn57.Header.VisiblePosition = 11;
			ultraGridColumn57.Hidden = true;
			ultraGridColumn57.Width = 90;
			ultraGridColumn58.Header.VisiblePosition = 13;
			ultraGridColumn58.Width = 90;
			ultraGridColumn59.Case = Infragistics.Win.UltraWinGrid.Case.Upper;
			ultraGridColumn59.Header.Caption = "Empresa Trabajo";
			ultraGridColumn59.Header.VisiblePosition = 15;
			ultraGridColumn59.Hidden = true;
			ultraGridColumn59.Width = 120;
			ultraGridColumn60.Case = Infragistics.Win.UltraWinGrid.Case.Upper;
			ultraGridColumn60.Header.Caption = "Direccion Trabajo";
			ultraGridColumn60.Header.VisiblePosition = 16;
			ultraGridColumn60.Hidden = true;
			ultraGridColumn60.Width = 150;
			ultraGridColumn61.Header.Caption = "Telefono Trabajo";
			ultraGridColumn61.Header.VisiblePosition = 17;
			ultraGridColumn61.Hidden = true;
			ultraGridColumn62.Header.Caption = "Extension";
			ultraGridColumn62.Header.VisiblePosition = 19;
			ultraGridColumn62.Hidden = true;
			ultraGridColumn62.Width = 60;
			ultraGridColumn63.ButtonDisplayStyle = Infragistics.Win.UltraWinGrid.ButtonDisplayStyle.Always;
			appearance8.BackColorAlpha = Infragistics.Win.Alpha.Opaque;
			appearance8.Image = ((object)(resources.GetObject("appearance8.Image")));
			appearance8.ImageHAlign = Infragistics.Win.HAlign.Center;
			ultraGridColumn63.CellButtonAppearance = appearance8;
			ultraGridColumn63.Header.Caption = "...";
			ultraGridColumn63.Header.VisiblePosition = 12;
			ultraGridColumn63.Hidden = true;
			ultraGridColumn63.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.Button;
			ultraGridColumn63.Width = 30;
			ultraGridColumn64.ButtonDisplayStyle = Infragistics.Win.UltraWinGrid.ButtonDisplayStyle.Always;
			appearance9.BackColorAlpha = Infragistics.Win.Alpha.Opaque;
			appearance9.Image = ((object)(resources.GetObject("appearance9.Image")));
			appearance9.ImageHAlign = Infragistics.Win.HAlign.Center;
			ultraGridColumn64.CellButtonAppearance = appearance9;
			ultraGridColumn64.Header.Caption = "...";
			ultraGridColumn64.Header.VisiblePosition = 14;
			ultraGridColumn64.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.Button;
			ultraGridColumn64.Width = 30;
			ultraGridColumn65.ButtonDisplayStyle = Infragistics.Win.UltraWinGrid.ButtonDisplayStyle.Always;
			appearance10.BackColorAlpha = Infragistics.Win.Alpha.Opaque;
			appearance10.Image = ((object)(resources.GetObject("appearance10.Image")));
			appearance10.ImageHAlign = Infragistics.Win.HAlign.Center;
			ultraGridColumn65.CellButtonAppearance = appearance10;
			ultraGridColumn65.Header.Caption = "...";
			ultraGridColumn65.Header.VisiblePosition = 18;
			ultraGridColumn65.Hidden = true;
			ultraGridColumn65.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.Button;
			ultraGridColumn65.Width = 30;
			ultraGridBand23.Columns.AddRange(new object[] {
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
			this.uGridParentesco.DisplayLayout.BandsSerializer.Add(ultraGridBand23);
			this.uGridParentesco.DisplayLayout.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.TemplateOnBottom;
			this.uGridParentesco.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.True;
			this.uGridParentesco.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.True;
			appearance11.BackColor = System.Drawing.Color.Transparent;
			this.uGridParentesco.DisplayLayout.Override.CardAreaAppearance = appearance11;
			appearance12.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance12.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance12.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance12.FontData.BoldAsString = "True";
			appearance12.FontData.Name = "Arial";
			appearance12.FontData.SizeInPoints = 8F;
			appearance12.ForeColor = System.Drawing.Color.White;
			appearance12.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.uGridParentesco.DisplayLayout.Override.HeaderAppearance = appearance12;
			appearance13.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance13.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance13.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridParentesco.DisplayLayout.Override.RowSelectorAppearance = appearance13;
			appearance14.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance14.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance14.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridParentesco.DisplayLayout.Override.SelectedRowAppearance = appearance14;
			this.uGridParentesco.Enabled = false;
			this.uGridParentesco.Font = new System.Drawing.Font("Tahoma", 7.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.uGridParentesco.Location = new System.Drawing.Point(8, 8);
			this.uGridParentesco.Name = "uGridParentesco";
			this.uGridParentesco.Size = new System.Drawing.Size(1087, 221);
			this.uGridParentesco.TabIndex = 11;
			this.uGridParentesco.AfterRowInsert += new Infragistics.Win.UltraWinGrid.RowEventHandler(this.uGridParentesco_AfterRowInsert);
			this.uGridParentesco.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.uGridParentesco_KeyPress);
			this.uGridParentesco.BeforeRowsDeleted += new Infragistics.Win.UltraWinGrid.BeforeRowsDeletedEventHandler(this.uGridParentesco_BeforeRowsDeleted);
			this.uGridParentesco.KeyDown += new System.Windows.Forms.KeyEventHandler(this.uGridParentesco_KeyDown);
			this.uGridParentesco.ClickCellButton += new Infragistics.Win.UltraWinGrid.CellEventHandler(this.uGridParentesco_ClickCellButton);
			this.uGridParentesco.InitializeLayout += new Infragistics.Win.UltraWinGrid.InitializeLayoutEventHandler(this.uGridParentesco_InitializeLayout);
			// 
			// ultraDataSource20
			// 
			ultraDataColumn25.DataType = typeof(int);
			ultraDataColumn26.DataType = typeof(int);
			ultraDataColumn27.DataType = typeof(int);
			ultraDataColumn28.DataType = typeof(int);
			ultraDataColumn33.DataType = typeof(int);
			ultraDataColumn34.DataType = typeof(int);
			this.ultraDataSource20.Band.Columns.AddRange(new object[] {
																																	ultraDataColumn25,
																																	ultraDataColumn26,
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
																																	ultraDataColumn39,
																																	ultraDataColumn40,
																																	ultraDataColumn41});
			// 
			// cmbRProvincia
			// 
			this.cmbRProvincia.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbRProvincia.DataSource = this.ultraDataSource1;
			ultraGridColumn66.Header.VisiblePosition = 0;
			ultraGridColumn66.Hidden = true;
			ultraGridColumn67.Header.VisiblePosition = 1;
			ultraGridColumn67.Width = 140;
			ultraGridBand24.Columns.AddRange(new object[] {
																											ultraGridColumn66,
																											ultraGridColumn67});
			this.cmbRProvincia.DisplayLayout.BandsSerializer.Add(ultraGridBand24);
			this.cmbRProvincia.DisplayMember = "Nombre";
			this.cmbRProvincia.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbRProvincia.Location = new System.Drawing.Point(192, 248);
			this.cmbRProvincia.Name = "cmbRProvincia";
			this.cmbRProvincia.Size = new System.Drawing.Size(173, 88);
			this.cmbRProvincia.TabIndex = 10;
			this.cmbRProvincia.ValueMember = "idProvincia";
			this.cmbRProvincia.Visible = false;
			this.cmbRProvincia.RowSelected += new Infragistics.Win.UltraWinGrid.RowSelectedEventHandler(this.cmbRProvincia_RowSelected);
			// 
			// cmbRCanton
			// 
			this.cmbRCanton.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbRCanton.DataSource = this.ultraDataSource2;
			ultraGridColumn68.Header.VisiblePosition = 0;
			ultraGridColumn68.Hidden = true;
			ultraGridColumn68.Width = 140;
			ultraGridColumn69.Header.VisiblePosition = 1;
			ultraGridBand25.Columns.AddRange(new object[] {
																											ultraGridColumn68,
																											ultraGridColumn69});
			this.cmbRCanton.DisplayLayout.BandsSerializer.Add(ultraGridBand25);
			this.cmbRCanton.DisplayMember = "Nombre";
			this.cmbRCanton.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbRCanton.Location = new System.Drawing.Point(376, 248);
			this.cmbRCanton.Name = "cmbRCanton";
			this.cmbRCanton.Size = new System.Drawing.Size(173, 88);
			this.cmbRCanton.TabIndex = 9;
			this.cmbRCanton.ValueMember = "idCanton";
			this.cmbRCanton.Visible = false;
			this.cmbRCanton.Validating += new System.ComponentModel.CancelEventHandler(this.cmbRCanton_Validating);
			this.cmbRCanton.Validated += new System.EventHandler(this.cmbRCanton_Validated);
			// 
			// cmbParentesco
			// 
			this.cmbParentesco.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbParentesco.DataSource = this.ultraDataSource18;
			ultraGridColumn70.Header.VisiblePosition = 0;
			ultraGridColumn70.Hidden = true;
			ultraGridColumn71.Header.VisiblePosition = 1;
			ultraGridColumn71.Width = 135;
			ultraGridBand26.Columns.AddRange(new object[] {
																											ultraGridColumn70,
																											ultraGridColumn71});
			this.cmbParentesco.DisplayLayout.BandsSerializer.Add(ultraGridBand26);
			this.cmbParentesco.DisplayMember = "Nombre";
			this.cmbParentesco.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbParentesco.Location = new System.Drawing.Point(8, 248);
			this.cmbParentesco.Name = "cmbParentesco";
			this.cmbParentesco.Size = new System.Drawing.Size(173, 88);
			this.cmbParentesco.TabIndex = 7;
			this.cmbParentesco.ValueMember = "idParentesco";
			this.cmbParentesco.Visible = false;
			// 
			// ultraDataSource18
			// 
			ultraDataColumn42.DataType = typeof(int);
			this.ultraDataSource18.Band.Columns.AddRange(new object[] {
																																	ultraDataColumn42,
																																	ultraDataColumn43});
			// 
			// utpImagenes
			// 
			this.utpImagenes.Controls.Add(this.btnCCelular);
			this.utpImagenes.Controls.Add(this.btnCTelf2);
			this.utpImagenes.Controls.Add(this.btnCTelf);
			this.utpImagenes.Controls.Add(this.cmbSubActvidadEconomica);
			this.utpImagenes.Controls.Add(this.chkSNAfiliacion);
			this.utpImagenes.Controls.Add(this.chkSNContabilidad);
			this.utpImagenes.Controls.Add(this.label93);
			this.utpImagenes.Controls.Add(this.txtSNPasivos);
			this.utpImagenes.Controls.Add(this.label94);
			this.utpImagenes.Controls.Add(this.txtSNActivos);
			this.utpImagenes.Controls.Add(this.label95);
			this.utpImagenes.Controls.Add(this.txtSNGastos);
			this.utpImagenes.Controls.Add(this.label96);
			this.utpImagenes.Controls.Add(this.label97);
			this.utpImagenes.Controls.Add(this.txtSNReferencia);
			this.utpImagenes.Controls.Add(this.txtSNIngresos);
			this.utpImagenes.Controls.Add(this.txtSNEmail);
			this.utpImagenes.Controls.Add(this.txtSNCelular);
			this.utpImagenes.Controls.Add(this.txtSNTelefono2);
			this.utpImagenes.Controls.Add(this.label86);
			this.utpImagenes.Controls.Add(this.txtSNTelefono1);
			this.utpImagenes.Controls.Add(this.label87);
			this.utpImagenes.Controls.Add(this.txtSNCalle2);
			this.utpImagenes.Controls.Add(this.label88);
			this.utpImagenes.Controls.Add(this.txtSNNumCasa);
			this.utpImagenes.Controls.Add(this.label89);
			this.utpImagenes.Controls.Add(this.txtSNCalle1);
			this.utpImagenes.Controls.Add(this.label90);
			this.utpImagenes.Controls.Add(this.label91);
			this.utpImagenes.Controls.Add(this.label92);
			this.utpImagenes.Controls.Add(this.label82);
			this.utpImagenes.Controls.Add(this.cmbSNProvincia);
			this.utpImagenes.Controls.Add(this.label83);
			this.utpImagenes.Controls.Add(this.cmbSNCanton);
			this.utpImagenes.Controls.Add(this.label84);
			this.utpImagenes.Controls.Add(this.cmbSNParroquia);
			this.utpImagenes.Controls.Add(this.label85);
			this.utpImagenes.Controls.Add(this.cmbSNBarrio);
			this.utpImagenes.Controls.Add(this.label81);
			this.utpImagenes.Controls.Add(this.label80);
			this.utpImagenes.Controls.Add(this.cmbSCTipoNegocio);
			this.utpImagenes.Controls.Add(this.txtSNEmpleados);
			this.utpImagenes.Controls.Add(this.txtSNMetros);
			this.utpImagenes.Controls.Add(this.label53);
			this.utpImagenes.Controls.Add(this.label79);
			this.utpImagenes.Controls.Add(this.label48);
			this.utpImagenes.Controls.Add(this.dtSNFechaInicio);
			this.utpImagenes.Controls.Add(this.txtSNNombre);
			this.utpImagenes.Controls.Add(this.label47);
			this.utpImagenes.Controls.Add(this.txtSNRuc);
			this.utpImagenes.Controls.Add(this.label46);
			this.utpImagenes.Location = new System.Drawing.Point(-10000, -10000);
			this.utpImagenes.Name = "utpImagenes";
			this.utpImagenes.Size = new System.Drawing.Size(1102, 323);
			this.utpImagenes.Paint += new System.Windows.Forms.PaintEventHandler(this.utpImagenes_Paint);
			// 
			// btnCCelular
			// 
			this.btnCCelular.Enabled = false;
			this.btnCCelular.Image = ((System.Drawing.Image)(resources.GetObject("btnCCelular.Image")));
			this.btnCCelular.Location = new System.Drawing.Point(864, 112);
			this.btnCCelular.Name = "btnCCelular";
			this.btnCCelular.Size = new System.Drawing.Size(32, 24);
			this.btnCCelular.TabIndex = 739;
			this.btnCCelular.Click += new System.EventHandler(this.btnCCelular_Click);
			// 
			// btnCTelf2
			// 
			this.btnCTelf2.Enabled = false;
			this.btnCTelf2.Image = ((System.Drawing.Image)(resources.GetObject("btnCTelf2.Image")));
			this.btnCTelf2.Location = new System.Drawing.Point(736, 112);
			this.btnCTelf2.Name = "btnCTelf2";
			this.btnCTelf2.Size = new System.Drawing.Size(32, 24);
			this.btnCTelf2.TabIndex = 738;
			this.btnCTelf2.Click += new System.EventHandler(this.btnCTelf2_Click);
			// 
			// btnCTelf
			// 
			this.btnCTelf.Enabled = false;
			this.btnCTelf.Image = ((System.Drawing.Image)(resources.GetObject("btnCTelf.Image")));
			this.btnCTelf.Location = new System.Drawing.Point(608, 112);
			this.btnCTelf.Name = "btnCTelf";
			this.btnCTelf.Size = new System.Drawing.Size(32, 24);
			this.btnCTelf.TabIndex = 737;
			this.btnCTelf.Visible = false;
			this.btnCTelf.Click += new System.EventHandler(this.btnCTelf_Click);
			// 
			// cmbSubActvidadEconomica
			// 
			appearance15.FontData.Name = "Tahoma";
			appearance15.ForeColorDisabled = System.Drawing.Color.Black;
			this.cmbSubActvidadEconomica.Appearance = appearance15;
			this.cmbSubActvidadEconomica.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbSubActvidadEconomica.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbSubActvidadEconomica.DataSource = this.ultraDataSource12;
			ultraGridColumn72.Header.VisiblePosition = 0;
			ultraGridColumn72.Hidden = true;
			ultraGridColumn72.Width = 100;
			ultraGridColumn73.Header.VisiblePosition = 1;
			ultraGridColumn73.Width = 225;
			ultraGridBand27.Columns.AddRange(new object[] {
																											ultraGridColumn72,
																											ultraGridColumn73});
			this.cmbSubActvidadEconomica.DisplayLayout.BandsSerializer.Add(ultraGridBand27);
			this.cmbSubActvidadEconomica.DisplayMember = "Nombre";
			this.cmbSubActvidadEconomica.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbSubActvidadEconomica.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbSubActvidadEconomica.Enabled = false;
			this.cmbSubActvidadEconomica.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbSubActvidadEconomica.Location = new System.Drawing.Point(864, 24);
			this.cmbSubActvidadEconomica.Name = "cmbSubActvidadEconomica";
			this.cmbSubActvidadEconomica.Size = new System.Drawing.Size(224, 22);
			this.cmbSubActvidadEconomica.TabIndex = 736;
			this.cmbSubActvidadEconomica.ValueMember = "idSubActEcoNegocio";
			this.cmbSubActvidadEconomica.Visible = false;
			this.cmbSubActvidadEconomica.KeyDown += new System.Windows.Forms.KeyEventHandler(this.cmbSubActvidadEconomica_KeyDown);
			// 
			// ultraDataSource12
			// 
			ultraDataColumn44.DataType = typeof(int);
			this.ultraDataSource12.Band.Columns.AddRange(new object[] {
																																	ultraDataColumn44,
																																	ultraDataColumn45});
			// 
			// chkSNAfiliacion
			// 
			appearance16.ForeColorDisabled = System.Drawing.Color.Black;
			this.chkSNAfiliacion.Appearance = appearance16;
			this.chkSNAfiliacion.BackColor = System.Drawing.Color.Transparent;
			this.chkSNAfiliacion.CausesValidation = false;
			this.chkSNAfiliacion.Enabled = false;
			this.chkSNAfiliacion.Location = new System.Drawing.Point(808, 192);
			this.chkSNAfiliacion.Name = "chkSNAfiliacion";
			this.chkSNAfiliacion.Size = new System.Drawing.Size(136, 22);
			this.chkSNAfiliacion.TabIndex = 735;
			this.chkSNAfiliacion.Text = "Afiliado Tributario";
			this.chkSNAfiliacion.KeyDown += new System.Windows.Forms.KeyEventHandler(this.chkSNAfiliacion_KeyDown);
			// 
			// chkSNContabilidad
			// 
			appearance17.ForeColorDisabled = System.Drawing.Color.Black;
			this.chkSNContabilidad.Appearance = appearance17;
			this.chkSNContabilidad.BackColor = System.Drawing.Color.Transparent;
			this.chkSNContabilidad.CausesValidation = false;
			this.chkSNContabilidad.Enabled = false;
			this.chkSNContabilidad.Location = new System.Drawing.Point(808, 168);
			this.chkSNContabilidad.Name = "chkSNContabilidad";
			this.chkSNContabilidad.Size = new System.Drawing.Size(192, 22);
			this.chkSNContabilidad.TabIndex = 734;
			this.chkSNContabilidad.Text = "Obligado a Llevar Contabilidad";
			this.chkSNContabilidad.KeyDown += new System.Windows.Forms.KeyEventHandler(this.chkSNContabilidad_KeyDown);
			// 
			// label93
			// 
			this.label93.AutoSize = true;
			this.label93.BackColor = System.Drawing.SystemColors.Control;
			this.label93.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label93.Location = new System.Drawing.Point(736, 144);
			this.label93.Name = "label93";
			this.label93.Size = new System.Drawing.Size(48, 17);
			this.label93.TabIndex = 733;
			this.label93.Text = "Pasivos";
			this.label93.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// txtSNPasivos
			// 
			this.txtSNPasivos.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtSNPasivos.Enabled = false;
			this.txtSNPasivos.Location = new System.Drawing.Point(736, 168);
			this.txtSNPasivos.MinValue = 0;
			this.txtSNPasivos.Name = "txtSNPasivos";
			this.txtSNPasivos.PromptChar = ' ';
			this.txtSNPasivos.Size = new System.Drawing.Size(64, 22);
			this.txtSNPasivos.SpinButtonDisplayStyle = Infragistics.Win.ButtonDisplayStyle.Always;
			this.txtSNPasivos.TabIndex = 732;
			this.txtSNPasivos.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtSNPasivos_KeyDown);
			this.txtSNPasivos.Click += new System.EventHandler(this.txtSNPasivos_Click);
			this.txtSNPasivos.Enter += new System.EventHandler(this.txtSNPasivos_Enter);
			// 
			// label94
			// 
			this.label94.AutoSize = true;
			this.label94.BackColor = System.Drawing.SystemColors.Control;
			this.label94.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label94.Location = new System.Drawing.Point(664, 144);
			this.label94.Name = "label94";
			this.label94.Size = new System.Drawing.Size(44, 17);
			this.label94.TabIndex = 731;
			this.label94.Text = "Activos";
			this.label94.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// txtSNActivos
			// 
			this.txtSNActivos.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtSNActivos.Enabled = false;
			this.txtSNActivos.Location = new System.Drawing.Point(664, 168);
			this.txtSNActivos.MinValue = 0;
			this.txtSNActivos.Name = "txtSNActivos";
			this.txtSNActivos.PromptChar = ' ';
			this.txtSNActivos.Size = new System.Drawing.Size(64, 22);
			this.txtSNActivos.SpinButtonDisplayStyle = Infragistics.Win.ButtonDisplayStyle.Always;
			this.txtSNActivos.TabIndex = 730;
			this.txtSNActivos.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtSNActivos_KeyDown);
			this.txtSNActivos.Click += new System.EventHandler(this.txtSNActivos_Click);
			this.txtSNActivos.Enter += new System.EventHandler(this.txtSNActivos_Enter);
			// 
			// label95
			// 
			this.label95.AutoSize = true;
			this.label95.BackColor = System.Drawing.SystemColors.Control;
			this.label95.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label95.Location = new System.Drawing.Point(592, 144);
			this.label95.Name = "label95";
			this.label95.Size = new System.Drawing.Size(44, 17);
			this.label95.TabIndex = 729;
			this.label95.Text = "Gastos";
			this.label95.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// txtSNGastos
			// 
			this.txtSNGastos.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtSNGastos.Enabled = false;
			this.txtSNGastos.Location = new System.Drawing.Point(592, 168);
			this.txtSNGastos.MaskInput = "nnnn";
			this.txtSNGastos.MaxValue = 9999;
			this.txtSNGastos.MinValue = 0;
			this.txtSNGastos.Name = "txtSNGastos";
			this.txtSNGastos.PromptChar = ' ';
			this.txtSNGastos.Size = new System.Drawing.Size(64, 22);
			this.txtSNGastos.SpinButtonDisplayStyle = Infragistics.Win.ButtonDisplayStyle.Always;
			this.txtSNGastos.TabIndex = 728;
			this.txtSNGastos.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtSNGastos_KeyDown);
			this.txtSNGastos.Click += new System.EventHandler(this.txtSNGastos_Click);
			this.txtSNGastos.Enter += new System.EventHandler(this.txtSNGastos_Enter);
			// 
			// label96
			// 
			this.label96.AutoSize = true;
			this.label96.BackColor = System.Drawing.SystemColors.Control;
			this.label96.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label96.Location = new System.Drawing.Point(520, 144);
			this.label96.Name = "label96";
			this.label96.Size = new System.Drawing.Size(52, 17);
			this.label96.TabIndex = 727;
			this.label96.Text = "Ingresos";
			this.label96.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// label97
			// 
			this.label97.AutoSize = true;
			this.label97.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label97.Location = new System.Drawing.Point(8, 144);
			this.label97.Name = "label97";
			this.label97.Size = new System.Drawing.Size(139, 17);
			this.label97.TabIndex = 726;
			this.label97.Text = "Referencia de Ubicación";
			this.label97.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// txtSNReferencia
			// 
			appearance18.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtSNReferencia.Appearance = appearance18;
			this.txtSNReferencia.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtSNReferencia.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtSNReferencia.Enabled = false;
			this.txtSNReferencia.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.txtSNReferencia.Location = new System.Drawing.Point(8, 168);
			this.txtSNReferencia.Multiline = true;
			this.txtSNReferencia.Name = "txtSNReferencia";
			this.txtSNReferencia.Size = new System.Drawing.Size(504, 104);
			this.txtSNReferencia.TabIndex = 725;
			this.txtSNReferencia.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtSNReferencia_KeyDown);
			this.txtSNReferencia.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.txtSNReferencia_KeyPress);
			// 
			// txtSNIngresos
			// 
			this.txtSNIngresos.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtSNIngresos.Enabled = false;
			this.txtSNIngresos.Location = new System.Drawing.Point(520, 168);
			this.txtSNIngresos.MaskInput = "nnnn";
			this.txtSNIngresos.MaxValue = 9999;
			this.txtSNIngresos.MinValue = 0;
			this.txtSNIngresos.Name = "txtSNIngresos";
			this.txtSNIngresos.PromptChar = ' ';
			this.txtSNIngresos.Size = new System.Drawing.Size(64, 22);
			this.txtSNIngresos.SpinButtonDisplayStyle = Infragistics.Win.ButtonDisplayStyle.Always;
			this.txtSNIngresos.TabIndex = 724;
			this.txtSNIngresos.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtSNIngresos_KeyDown);
			this.txtSNIngresos.Click += new System.EventHandler(this.txtSNIngresos_Click);
			this.txtSNIngresos.Enter += new System.EventHandler(this.txtSNIngresos_Enter);
			// 
			// txtSNEmail
			// 
			this.txtSNEmail.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtSNEmail.Enabled = false;
			this.txtSNEmail.Location = new System.Drawing.Point(904, 112);
			this.txtSNEmail.MaxLength = 60;
			this.txtSNEmail.Name = "txtSNEmail";
			this.txtSNEmail.Size = new System.Drawing.Size(184, 22);
			this.txtSNEmail.TabIndex = 723;
			this.txtSNEmail.Visible = false;
			this.txtSNEmail.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtSNEmail_KeyDown);
			// 
			// txtSNCelular
			// 
			this.txtSNCelular.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtSNCelular.Enabled = false;
			this.txtSNCelular.Location = new System.Drawing.Point(776, 112);
			this.txtSNCelular.MaxLength = 10;
			this.txtSNCelular.Name = "txtSNCelular";
			this.txtSNCelular.Size = new System.Drawing.Size(80, 22);
			this.txtSNCelular.TabIndex = 722;
			this.txtSNCelular.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtSNCelular_KeyDown);
			this.txtSNCelular.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.txtSNCelular_KeyPress);
			// 
			// txtSNTelefono2
			// 
			this.txtSNTelefono2.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtSNTelefono2.Enabled = false;
			this.txtSNTelefono2.Location = new System.Drawing.Point(648, 112);
			this.txtSNTelefono2.MaxLength = 9;
			this.txtSNTelefono2.Name = "txtSNTelefono2";
			this.txtSNTelefono2.Size = new System.Drawing.Size(80, 22);
			this.txtSNTelefono2.TabIndex = 721;
			this.txtSNTelefono2.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtSNTelefono2_KeyDown);
			// 
			// label86
			// 
			this.label86.AutoSize = true;
			this.label86.BackColor = System.Drawing.SystemColors.Control;
			this.label86.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label86.Location = new System.Drawing.Point(648, 96);
			this.label86.Name = "label86";
			this.label86.Size = new System.Drawing.Size(95, 17);
			this.label86.TabIndex = 720;
			this.label86.Text = "Teléfono Casa 2";
			this.label86.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// txtSNTelefono1
			// 
			this.txtSNTelefono1.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtSNTelefono1.Enabled = false;
			this.txtSNTelefono1.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.txtSNTelefono1.Location = new System.Drawing.Point(520, 112);
			this.txtSNTelefono1.MaxLength = 9;
			this.txtSNTelefono1.Name = "txtSNTelefono1";
			this.txtSNTelefono1.Size = new System.Drawing.Size(80, 23);
			this.txtSNTelefono1.TabIndex = 719;
			this.txtSNTelefono1.Visible = false;
			this.txtSNTelefono1.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtSNTelefono1_KeyDown);
			this.txtSNTelefono1.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.txtSNTelefono1_KeyPress);
			// 
			// label87
			// 
			this.label87.AutoSize = true;
			this.label87.BackColor = System.Drawing.SystemColors.Control;
			this.label87.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label87.Location = new System.Drawing.Point(520, 96);
			this.label87.Name = "label87";
			this.label87.Size = new System.Drawing.Size(95, 17);
			this.label87.TabIndex = 718;
			this.label87.Text = "Teléfono Casa 1";
			this.label87.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// txtSNCalle2
			// 
			this.txtSNCalle2.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtSNCalle2.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtSNCalle2.Enabled = false;
			this.txtSNCalle2.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.txtSNCalle2.Location = new System.Drawing.Point(320, 112);
			this.txtSNCalle2.Name = "txtSNCalle2";
			this.txtSNCalle2.Size = new System.Drawing.Size(192, 23);
			this.txtSNCalle2.TabIndex = 717;
			this.txtSNCalle2.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtSNCalle2_KeyDown);
			this.txtSNCalle2.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.txtSNCalle2_KeyPress);
			// 
			// label88
			// 
			this.label88.AutoSize = true;
			this.label88.BackColor = System.Drawing.SystemColors.Control;
			this.label88.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label88.Location = new System.Drawing.Point(320, 96);
			this.label88.Name = "label88";
			this.label88.Size = new System.Drawing.Size(99, 17);
			this.label88.TabIndex = 716;
			this.label88.Text = "Calle Secundaria";
			this.label88.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// txtSNNumCasa
			// 
			this.txtSNNumCasa.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtSNNumCasa.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtSNNumCasa.Enabled = false;
			this.txtSNNumCasa.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.txtSNNumCasa.Location = new System.Drawing.Point(208, 112);
			this.txtSNNumCasa.Name = "txtSNNumCasa";
			this.txtSNNumCasa.Size = new System.Drawing.Size(104, 23);
			this.txtSNNumCasa.TabIndex = 715;
			this.txtSNNumCasa.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtSNNumCasa_KeyDown);
			this.txtSNNumCasa.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.txtSNNumCasa_KeyPress);
			// 
			// label89
			// 
			this.label89.AutoSize = true;
			this.label89.BackColor = System.Drawing.SystemColors.Control;
			this.label89.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label89.Location = new System.Drawing.Point(208, 96);
			this.label89.Name = "label89";
			this.label89.Size = new System.Drawing.Size(98, 17);
			this.label89.TabIndex = 714;
			this.label89.Text = "Número de Casa";
			this.label89.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// txtSNCalle1
			// 
			this.txtSNCalle1.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtSNCalle1.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtSNCalle1.Enabled = false;
			this.txtSNCalle1.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.txtSNCalle1.Location = new System.Drawing.Point(8, 112);
			this.txtSNCalle1.Name = "txtSNCalle1";
			this.txtSNCalle1.Size = new System.Drawing.Size(192, 23);
			this.txtSNCalle1.TabIndex = 713;
			this.txtSNCalle1.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtSNCalle1_KeyDown);
			this.txtSNCalle1.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.txtSNCalle1_KeyPress);
			// 
			// label90
			// 
			this.label90.AutoSize = true;
			this.label90.BackColor = System.Drawing.SystemColors.Control;
			this.label90.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label90.Location = new System.Drawing.Point(8, 96);
			this.label90.Name = "label90";
			this.label90.Size = new System.Drawing.Size(84, 17);
			this.label90.TabIndex = 712;
			this.label90.Text = "Calle Principal";
			this.label90.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// label91
			// 
			this.label91.AutoSize = true;
			this.label91.BackColor = System.Drawing.SystemColors.Control;
			this.label91.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label91.Location = new System.Drawing.Point(904, 96);
			this.label91.Name = "label91";
			this.label91.Size = new System.Drawing.Size(43, 17);
			this.label91.TabIndex = 711;
			this.label91.Text = "E-Mail ";
			this.label91.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// label92
			// 
			this.label92.AutoSize = true;
			this.label92.BackColor = System.Drawing.SystemColors.Control;
			this.label92.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label92.Location = new System.Drawing.Point(776, 96);
			this.label92.Name = "label92";
			this.label92.Size = new System.Drawing.Size(44, 17);
			this.label92.TabIndex = 710;
			this.label92.Text = "Celular";
			this.label92.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// label82
			// 
			this.label82.AutoSize = true;
			this.label82.BackColor = System.Drawing.SystemColors.Control;
			this.label82.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label82.Location = new System.Drawing.Point(8, 48);
			this.label82.Name = "label82";
			this.label82.Size = new System.Drawing.Size(55, 17);
			this.label82.TabIndex = 703;
			this.label82.Text = "Provincia";
			this.label82.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// cmbSNProvincia
			// 
			this.cmbSNProvincia.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbSNProvincia.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbSNProvincia.DataSource = this.ultraDataSource1;
			ultraGridColumn74.Header.VisiblePosition = 0;
			ultraGridColumn74.Hidden = true;
			ultraGridColumn75.Header.VisiblePosition = 1;
			ultraGridColumn75.Width = 256;
			ultraGridBand28.Columns.AddRange(new object[] {
																											ultraGridColumn74,
																											ultraGridColumn75});
			this.cmbSNProvincia.DisplayLayout.BandsSerializer.Add(ultraGridBand28);
			this.cmbSNProvincia.DisplayMember = "Nombre";
			this.cmbSNProvincia.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbSNProvincia.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbSNProvincia.Enabled = false;
			this.cmbSNProvincia.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbSNProvincia.Location = new System.Drawing.Point(8, 72);
			this.cmbSNProvincia.Name = "cmbSNProvincia";
			this.cmbSNProvincia.Size = new System.Drawing.Size(184, 22);
			this.cmbSNProvincia.TabIndex = 709;
			this.cmbSNProvincia.ValueMember = "idProvincia";
			this.cmbSNProvincia.KeyDown += new System.Windows.Forms.KeyEventHandler(this.cmbSNProvincia_KeyDown);
			this.cmbSNProvincia.ValueChanged += new System.EventHandler(this.cmbSNProvincia_ValueChanged);
			// 
			// label83
			// 
			this.label83.AutoSize = true;
			this.label83.BackColor = System.Drawing.SystemColors.Control;
			this.label83.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label83.Location = new System.Drawing.Point(200, 48);
			this.label83.Name = "label83";
			this.label83.Size = new System.Drawing.Size(44, 17);
			this.label83.TabIndex = 702;
			this.label83.Text = "Canton";
			this.label83.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// cmbSNCanton
			// 
			this.cmbSNCanton.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbSNCanton.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbSNCanton.DataSource = this.ultraDataSource2;
			ultraGridBand29.AddButtonCaption = "RolSucursal";
			ultraGridColumn76.Header.VisiblePosition = 0;
			ultraGridColumn76.Hidden = true;
			ultraGridColumn77.Header.VisiblePosition = 1;
			ultraGridColumn77.Width = 224;
			ultraGridBand29.Columns.AddRange(new object[] {
																											ultraGridColumn76,
																											ultraGridColumn77});
			this.cmbSNCanton.DisplayLayout.BandsSerializer.Add(ultraGridBand29);
			this.cmbSNCanton.DisplayLayout.Override.RowSelectors = Infragistics.Win.DefaultableBoolean.True;
			this.cmbSNCanton.DisplayMember = "Nombre";
			this.cmbSNCanton.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbSNCanton.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbSNCanton.Enabled = false;
			this.cmbSNCanton.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbSNCanton.Location = new System.Drawing.Point(200, 72);
			this.cmbSNCanton.Name = "cmbSNCanton";
			this.cmbSNCanton.Size = new System.Drawing.Size(184, 22);
			this.cmbSNCanton.TabIndex = 704;
			this.cmbSNCanton.ValueMember = "idCanton";
			this.cmbSNCanton.KeyDown += new System.Windows.Forms.KeyEventHandler(this.cmbSNCanton_KeyDown);
			this.cmbSNCanton.ValueChanged += new System.EventHandler(this.cmbSNCanton_ValueChanged);
			// 
			// label84
			// 
			this.label84.AutoSize = true;
			this.label84.BackColor = System.Drawing.SystemColors.Control;
			this.label84.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label84.Location = new System.Drawing.Point(392, 48);
			this.label84.Name = "label84";
			this.label84.Size = new System.Drawing.Size(58, 17);
			this.label84.TabIndex = 706;
			this.label84.Text = "Parroquia";
			this.label84.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// cmbSNParroquia
			// 
			this.cmbSNParroquia.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbSNParroquia.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbSNParroquia.DataSource = this.ultraDataSource9;
			ultraGridBand30.AddButtonCaption = "RolSucursal";
			ultraGridColumn78.Header.VisiblePosition = 0;
			ultraGridColumn78.Hidden = true;
			ultraGridColumn79.Header.VisiblePosition = 1;
			ultraGridColumn79.Width = 232;
			ultraGridBand30.Columns.AddRange(new object[] {
																											ultraGridColumn78,
																											ultraGridColumn79});
			this.cmbSNParroquia.DisplayLayout.BandsSerializer.Add(ultraGridBand30);
			this.cmbSNParroquia.DisplayMember = "Nombre";
			this.cmbSNParroquia.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbSNParroquia.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbSNParroquia.Enabled = false;
			this.cmbSNParroquia.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbSNParroquia.Location = new System.Drawing.Point(392, 72);
			this.cmbSNParroquia.Name = "cmbSNParroquia";
			this.cmbSNParroquia.Size = new System.Drawing.Size(184, 22);
			this.cmbSNParroquia.TabIndex = 707;
			this.cmbSNParroquia.ValueMember = "idParroquia";
			this.cmbSNParroquia.KeyDown += new System.Windows.Forms.KeyEventHandler(this.cmbSNParroquia_KeyDown);
			this.cmbSNParroquia.ValueChanged += new System.EventHandler(this.cmbSNParroquia_ValueChanged);
			// 
			// label85
			// 
			this.label85.AutoSize = true;
			this.label85.BackColor = System.Drawing.SystemColors.Control;
			this.label85.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label85.Location = new System.Drawing.Point(584, 48);
			this.label85.Name = "label85";
			this.label85.Size = new System.Drawing.Size(38, 17);
			this.label85.TabIndex = 705;
			this.label85.Text = "Barrio";
			this.label85.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// cmbSNBarrio
			// 
			this.cmbSNBarrio.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbSNBarrio.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbSNBarrio.DataSource = this.ultraDataSource10;
			ultraGridBand31.AddButtonCaption = "RolSucursal";
			ultraGridColumn80.Header.VisiblePosition = 0;
			ultraGridColumn80.Hidden = true;
			ultraGridColumn81.Header.VisiblePosition = 1;
			ultraGridColumn81.Width = 240;
			ultraGridBand31.Columns.AddRange(new object[] {
																											ultraGridColumn80,
																											ultraGridColumn81});
			this.cmbSNBarrio.DisplayLayout.BandsSerializer.Add(ultraGridBand31);
			this.cmbSNBarrio.DisplayLayout.Override.RowSelectors = Infragistics.Win.DefaultableBoolean.True;
			this.cmbSNBarrio.DisplayMember = "Nombre";
			this.cmbSNBarrio.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbSNBarrio.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbSNBarrio.Enabled = false;
			this.cmbSNBarrio.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbSNBarrio.Location = new System.Drawing.Point(584, 72);
			this.cmbSNBarrio.Name = "cmbSNBarrio";
			this.cmbSNBarrio.Size = new System.Drawing.Size(184, 22);
			this.cmbSNBarrio.TabIndex = 708;
			this.cmbSNBarrio.ValueMember = "idBarrio";
			this.cmbSNBarrio.KeyDown += new System.Windows.Forms.KeyEventHandler(this.cmbSNBarrio_KeyDown);
			// 
			// label81
			// 
			this.label81.AutoSize = true;
			this.label81.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label81.Location = new System.Drawing.Point(864, 8);
			this.label81.Name = "label81";
			this.label81.Size = new System.Drawing.Size(140, 17);
			this.label81.TabIndex = 700;
			this.label81.Text = "Subactividad Económica";
			this.label81.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// label80
			// 
			this.label80.AutoSize = true;
			this.label80.Location = new System.Drawing.Point(696, 8);
			this.label80.Name = "label80";
			this.label80.Size = new System.Drawing.Size(29, 17);
			this.label80.TabIndex = 694;
			this.label80.Text = "Tipo ";
			this.label80.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
			// 
			// cmbSCTipoNegocio
			// 
			appearance19.FontData.Name = "Tahoma";
			appearance19.ForeColorDisabled = System.Drawing.Color.Black;
			this.cmbSCTipoNegocio.Appearance = appearance19;
			this.cmbSCTipoNegocio.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbSCTipoNegocio.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbSCTipoNegocio.DataSource = this.ultraDataSource11;
			ultraGridColumn82.Header.VisiblePosition = 0;
			ultraGridColumn82.Hidden = true;
			ultraGridColumn83.Header.VisiblePosition = 1;
			ultraGridColumn83.Width = 160;
			ultraGridBand32.Columns.AddRange(new object[] {
																											ultraGridColumn82,
																											ultraGridColumn83});
			this.cmbSCTipoNegocio.DisplayLayout.BandsSerializer.Add(ultraGridBand32);
			this.cmbSCTipoNegocio.DisplayMember = "Nombre";
			this.cmbSCTipoNegocio.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbSCTipoNegocio.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbSCTipoNegocio.Enabled = false;
			this.cmbSCTipoNegocio.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbSCTipoNegocio.Location = new System.Drawing.Point(696, 24);
			this.cmbSCTipoNegocio.Name = "cmbSCTipoNegocio";
			this.cmbSCTipoNegocio.Size = new System.Drawing.Size(160, 22);
			this.cmbSCTipoNegocio.TabIndex = 693;
			this.cmbSCTipoNegocio.ValueMember = "idTipoVivienda";
			this.cmbSCTipoNegocio.Visible = false;
			this.cmbSCTipoNegocio.KeyDown += new System.Windows.Forms.KeyEventHandler(this.cmbSCTipoNegocio_KeyDown);
			// 
			// txtSNEmpleados
			// 
			this.txtSNEmpleados.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtSNEmpleados.Enabled = false;
			this.txtSNEmpleados.Location = new System.Drawing.Point(592, 24);
			this.txtSNEmpleados.MinValue = 0;
			this.txtSNEmpleados.Name = "txtSNEmpleados";
			this.txtSNEmpleados.PromptChar = ' ';
			this.txtSNEmpleados.Size = new System.Drawing.Size(96, 22);
			this.txtSNEmpleados.SpinButtonDisplayStyle = Infragistics.Win.ButtonDisplayStyle.Always;
			this.txtSNEmpleados.TabIndex = 692;
			this.txtSNEmpleados.Visible = false;
			this.txtSNEmpleados.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtSNEmpleados_KeyDown);
			this.txtSNEmpleados.Click += new System.EventHandler(this.txtSNEmpleados_Click);
			this.txtSNEmpleados.Enter += new System.EventHandler(this.txtSNEmpleados_Enter);
			// 
			// txtSNMetros
			// 
			this.txtSNMetros.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtSNMetros.Enabled = false;
			this.txtSNMetros.Location = new System.Drawing.Point(488, 24);
			this.txtSNMetros.MinValue = 0;
			this.txtSNMetros.Name = "txtSNMetros";
			this.txtSNMetros.PromptChar = ' ';
			this.txtSNMetros.Size = new System.Drawing.Size(96, 22);
			this.txtSNMetros.SpinButtonDisplayStyle = Infragistics.Win.ButtonDisplayStyle.Always;
			this.txtSNMetros.TabIndex = 691;
			this.txtSNMetros.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtSNMetros_KeyDown);
			this.txtSNMetros.Click += new System.EventHandler(this.txtSNMetros_Click);
			this.txtSNMetros.Enter += new System.EventHandler(this.txtSNMetros_Enter);
			// 
			// label53
			// 
			this.label53.AutoSize = true;
			this.label53.BackColor = System.Drawing.SystemColors.Control;
			this.label53.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label53.Location = new System.Drawing.Point(592, 8);
			this.label53.Name = "label53";
			this.label53.Size = new System.Drawing.Size(66, 17);
			this.label53.TabIndex = 690;
			this.label53.Text = "Empleados";
			this.label53.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// label79
			// 
			this.label79.AutoSize = true;
			this.label79.BackColor = System.Drawing.SystemColors.Control;
			this.label79.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label79.Location = new System.Drawing.Point(488, 8);
			this.label79.Name = "label79";
			this.label79.Size = new System.Drawing.Size(42, 17);
			this.label79.TabIndex = 688;
			this.label79.Text = "Metros";
			this.label79.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// label48
			// 
			this.label48.AutoSize = true;
			this.label48.Location = new System.Drawing.Point(368, 8);
			this.label48.Name = "label48";
			this.label48.Size = new System.Drawing.Size(64, 17);
			this.label48.TabIndex = 380;
			this.label48.Text = "Fecha Inicio";
			this.label48.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// dtSNFechaInicio
			// 
			appearance20.ForeColorDisabled = System.Drawing.Color.Black;
			this.dtSNFechaInicio.Appearance = appearance20;
			this.dtSNFechaInicio.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			dateButton4.Caption = "Today";
			this.dtSNFechaInicio.DateButtons.Add(dateButton4);
			this.dtSNFechaInicio.Enabled = false;
			this.dtSNFechaInicio.Format = "dd/MM/yyyy";
			this.dtSNFechaInicio.Location = new System.Drawing.Point(368, 24);
			this.dtSNFechaInicio.Name = "dtSNFechaInicio";
			this.dtSNFechaInicio.NonAutoSizeHeight = 23;
			this.dtSNFechaInicio.Size = new System.Drawing.Size(112, 21);
			this.dtSNFechaInicio.SpinButtonsVisible = true;
			this.dtSNFechaInicio.TabIndex = 379;
			this.dtSNFechaInicio.Value = ((object)(resources.GetObject("dtSNFechaInicio.Value")));
			this.dtSNFechaInicio.KeyDown += new System.Windows.Forms.KeyEventHandler(this.dtSNFechaInicio_KeyDown);
			// 
			// txtSNNombre
			// 
			appearance21.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtSNNombre.Appearance = appearance21;
			this.txtSNNombre.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtSNNombre.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtSNNombre.Enabled = false;
			this.txtSNNombre.Location = new System.Drawing.Point(152, 24);
			this.txtSNNombre.Name = "txtSNNombre";
			this.txtSNNombre.Size = new System.Drawing.Size(208, 22);
			this.txtSNNombre.TabIndex = 378;
			this.txtSNNombre.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtSNNombre_KeyDown);
			this.txtSNNombre.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.txtSNNombre_KeyPress);
			// 
			// label47
			// 
			this.label47.AutoSize = true;
			this.label47.Location = new System.Drawing.Point(152, 8);
			this.label47.Name = "label47";
			this.label47.Size = new System.Drawing.Size(44, 17);
			this.label47.TabIndex = 377;
			this.label47.Text = "Nombre";
			this.label47.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// txtSNRuc
			// 
			appearance22.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtSNRuc.Appearance = appearance22;
			this.txtSNRuc.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtSNRuc.Enabled = false;
			this.txtSNRuc.Location = new System.Drawing.Point(8, 24);
			this.txtSNRuc.MaxLength = 13;
			this.txtSNRuc.Name = "txtSNRuc";
			this.txtSNRuc.Size = new System.Drawing.Size(136, 22);
			this.txtSNRuc.TabIndex = 362;
			this.txtSNRuc.Visible = false;
			this.txtSNRuc.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtSNRuc_KeyDown);
			this.txtSNRuc.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.txtSNRuc_KeyPress);
			this.txtSNRuc.Validating += new System.ComponentModel.CancelEventHandler(this.txtSNRuc_Validating);
			// 
			// label46
			// 
			this.label46.AutoSize = true;
			this.label46.Location = new System.Drawing.Point(8, 8);
			this.label46.Name = "label46";
			this.label46.Size = new System.Drawing.Size(25, 17);
			this.label46.TabIndex = 361;
			this.label46.Text = "RUC";
			this.label46.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// ultraTabPageControl1
			// 
			this.ultraTabPageControl1.Controls.Add(this.dtSDFecha);
			this.ultraTabPageControl1.Controls.Add(this.btnSDCelular);
			this.ultraTabPageControl1.Controls.Add(this.bntSDTelf);
			this.ultraTabPageControl1.Controls.Add(this.cmbSDTipoContrato);
			this.ultraTabPageControl1.Controls.Add(this.txtSDReferencia);
			this.ultraTabPageControl1.Controls.Add(this.label23);
			this.ultraTabPageControl1.Controls.Add(this.label120);
			this.ultraTabPageControl1.Controls.Add(this.txtSDExt);
			this.ultraTabPageControl1.Controls.Add(this.label109);
			this.ultraTabPageControl1.Controls.Add(this.txtSDEmail);
			this.ultraTabPageControl1.Controls.Add(this.txtSDCelular);
			this.ultraTabPageControl1.Controls.Add(this.txtSDTelefono);
			this.ultraTabPageControl1.Controls.Add(this.label110);
			this.ultraTabPageControl1.Controls.Add(this.txtSDCalle2);
			this.ultraTabPageControl1.Controls.Add(this.label111);
			this.ultraTabPageControl1.Controls.Add(this.txtSDNumCasa);
			this.ultraTabPageControl1.Controls.Add(this.label112);
			this.ultraTabPageControl1.Controls.Add(this.txtSDCalle1);
			this.ultraTabPageControl1.Controls.Add(this.label113);
			this.ultraTabPageControl1.Controls.Add(this.label114);
			this.ultraTabPageControl1.Controls.Add(this.label115);
			this.ultraTabPageControl1.Controls.Add(this.label116);
			this.ultraTabPageControl1.Controls.Add(this.cmbSDProvincia);
			this.ultraTabPageControl1.Controls.Add(this.label117);
			this.ultraTabPageControl1.Controls.Add(this.cmbSDCanton);
			this.ultraTabPageControl1.Controls.Add(this.label118);
			this.ultraTabPageControl1.Controls.Add(this.cmbSDParroquia);
			this.ultraTabPageControl1.Controls.Add(this.label119);
			this.ultraTabPageControl1.Controls.Add(this.cmbSDBarrio);
			this.ultraTabPageControl1.Controls.Add(this.chkSDAfiliacion);
			this.ultraTabPageControl1.Controls.Add(this.txtSDPago);
			this.ultraTabPageControl1.Controls.Add(this.label108);
			this.ultraTabPageControl1.Controls.Add(this.txtSDDepartamento);
			this.ultraTabPageControl1.Controls.Add(this.label107);
			this.ultraTabPageControl1.Controls.Add(this.label104);
			this.ultraTabPageControl1.Controls.Add(this.cmbSDCargo);
			this.ultraTabPageControl1.Controls.Add(this.cmbSDTipoSueldo);
			this.ultraTabPageControl1.Controls.Add(this.label105);
			this.ultraTabPageControl1.Controls.Add(this.label106);
			this.ultraTabPageControl1.Controls.Add(this.txtSDPasivos);
			this.ultraTabPageControl1.Controls.Add(this.txtSDActivos);
			this.ultraTabPageControl1.Controls.Add(this.txtSDGastos);
			this.ultraTabPageControl1.Controls.Add(this.txtSDIngresos);
			this.ultraTabPageControl1.Controls.Add(this.label99);
			this.ultraTabPageControl1.Controls.Add(this.label101);
			this.ultraTabPageControl1.Controls.Add(this.label102);
			this.ultraTabPageControl1.Controls.Add(this.label103);
			this.ultraTabPageControl1.Controls.Add(this.txtSDEmpresa);
			this.ultraTabPageControl1.Controls.Add(this.cmbSDTipoEmpresa);
			this.ultraTabPageControl1.Controls.Add(this.lblempres);
			this.ultraTabPageControl1.Controls.Add(this.label100);
			this.ultraTabPageControl1.Controls.Add(this.txtSDRuc);
			this.ultraTabPageControl1.Controls.Add(this.label98);
			this.ultraTabPageControl1.Location = new System.Drawing.Point(-10000, -10000);
			this.ultraTabPageControl1.Name = "ultraTabPageControl1";
			this.ultraTabPageControl1.Size = new System.Drawing.Size(1102, 323);
			this.ultraTabPageControl1.Paint += new System.Windows.Forms.PaintEventHandler(this.ultraTabPageControl1_Paint);
			// 
			// dtSDFecha
			// 
			appearance23.ForeColorDisabled = System.Drawing.Color.Black;
			this.dtSDFecha.Appearance = appearance23;
			this.dtSDFecha.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			dateButton5.Caption = "Today";
			this.dtSDFecha.DateButtons.Add(dateButton5);
			this.dtSDFecha.Enabled = false;
			this.dtSDFecha.Format = "dd/MM/yyyy";
			this.dtSDFecha.Location = new System.Drawing.Point(544, 24);
			this.dtSDFecha.Name = "dtSDFecha";
			this.dtSDFecha.NonAutoSizeHeight = 23;
			this.dtSDFecha.Size = new System.Drawing.Size(112, 21);
			this.dtSDFecha.SpinButtonsVisible = true;
			this.dtSDFecha.TabIndex = 789;
			this.dtSDFecha.Value = ((object)(resources.GetObject("dtSDFecha.Value")));
			this.dtSDFecha.KeyDown += new System.Windows.Forms.KeyEventHandler(this.dtSDFecha_KeyDown);
			// 
			// btnSDCelular
			// 
			this.btnSDCelular.Enabled = false;
			this.btnSDCelular.Image = ((System.Drawing.Image)(resources.GetObject("btnSDCelular.Image")));
			this.btnSDCelular.Location = new System.Drawing.Point(784, 152);
			this.btnSDCelular.Name = "btnSDCelular";
			this.btnSDCelular.Size = new System.Drawing.Size(32, 24);
			this.btnSDCelular.TabIndex = 788;
			this.btnSDCelular.Click += new System.EventHandler(this.btnSDCelular_Click);
			// 
			// bntSDTelf
			// 
			this.bntSDTelf.Enabled = false;
			this.bntSDTelf.Image = ((System.Drawing.Image)(resources.GetObject("bntSDTelf.Image")));
			this.bntSDTelf.Location = new System.Drawing.Point(608, 152);
			this.bntSDTelf.Name = "bntSDTelf";
			this.bntSDTelf.Size = new System.Drawing.Size(32, 24);
			this.bntSDTelf.TabIndex = 787;
			this.bntSDTelf.Click += new System.EventHandler(this.bntSDTelf_Click);
			// 
			// cmbSDTipoContrato
			// 
			this.cmbSDTipoContrato.CausesValidation = false;
			this.cmbSDTipoContrato.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbSDTipoContrato.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbSDTipoContrato.DataSource = this.ultraDataSource13;
			ultraGridColumn84.Header.VisiblePosition = 0;
			ultraGridColumn84.Hidden = true;
			ultraGridColumn85.Header.VisiblePosition = 1;
			ultraGridColumn85.Width = 176;
			ultraGridBand33.Columns.AddRange(new object[] {
																											ultraGridColumn84,
																											ultraGridColumn85});
			this.cmbSDTipoContrato.DisplayLayout.BandsSerializer.Add(ultraGridBand33);
			this.cmbSDTipoContrato.DisplayMember = "Nombre";
			this.cmbSDTipoContrato.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbSDTipoContrato.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbSDTipoContrato.Enabled = false;
			this.cmbSDTipoContrato.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbSDTipoContrato.Location = new System.Drawing.Point(8, 64);
			this.cmbSDTipoContrato.MaxDropDownItems = 30;
			this.cmbSDTipoContrato.Name = "cmbSDTipoContrato";
			this.cmbSDTipoContrato.Size = new System.Drawing.Size(176, 21);
			this.cmbSDTipoContrato.TabIndex = 786;
			this.cmbSDTipoContrato.ValueMember = "idTipoContrato";
			this.cmbSDTipoContrato.KeyDown += new System.Windows.Forms.KeyEventHandler(this.cmbSDTipoContrato_KeyDown);
			// 
			// ultraDataSource13
			// 
			ultraDataColumn46.DataType = typeof(int);
			this.ultraDataSource13.Band.Columns.AddRange(new object[] {
																																	ultraDataColumn46,
																																	ultraDataColumn47});
			// 
			// txtSDReferencia
			// 
			appearance24.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtSDReferencia.Appearance = appearance24;
			this.txtSDReferencia.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtSDReferencia.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtSDReferencia.Enabled = false;
			this.txtSDReferencia.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.txtSDReferencia.Location = new System.Drawing.Point(8, 200);
			this.txtSDReferencia.Multiline = true;
			this.txtSDReferencia.Name = "txtSDReferencia";
			this.txtSDReferencia.Size = new System.Drawing.Size(472, 88);
			this.txtSDReferencia.TabIndex = 785;
			this.txtSDReferencia.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtSDReferencia_KeyDown);
			this.txtSDReferencia.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.txtSDReferencia_KeyPress);
			// 
			// label23
			// 
			this.label23.AutoSize = true;
			this.label23.Location = new System.Drawing.Point(544, 8);
			this.label23.Name = "label23";
			this.label23.Size = new System.Drawing.Size(64, 17);
			this.label23.TabIndex = 784;
			this.label23.Text = "Fecha Inicio";
			this.label23.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// label120
			// 
			this.label120.AutoSize = true;
			this.label120.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label120.Location = new System.Drawing.Point(8, 184);
			this.label120.Name = "label120";
			this.label120.Size = new System.Drawing.Size(139, 17);
			this.label120.TabIndex = 782;
			this.label120.Text = "Referencia de Ubicación";
			this.label120.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// txtSDExt
			// 
			this.txtSDExt.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtSDExt.Enabled = false;
			this.txtSDExt.Location = new System.Drawing.Point(648, 152);
			this.txtSDExt.MaxLength = 9;
			this.txtSDExt.Name = "txtSDExt";
			this.txtSDExt.Size = new System.Drawing.Size(40, 22);
			this.txtSDExt.TabIndex = 780;
			this.txtSDExt.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtSDExt_KeyDown);
			// 
			// label109
			// 
			this.label109.AutoSize = true;
			this.label109.BackColor = System.Drawing.SystemColors.Control;
			this.label109.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label109.Location = new System.Drawing.Point(648, 136);
			this.label109.Name = "label109";
			this.label109.Size = new System.Drawing.Size(26, 17);
			this.label109.TabIndex = 779;
			this.label109.Text = "Ext.";
			this.label109.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// txtSDEmail
			// 
			this.txtSDEmail.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtSDEmail.Enabled = false;
			this.txtSDEmail.Location = new System.Drawing.Point(824, 152);
			this.txtSDEmail.MaxLength = 60;
			this.txtSDEmail.Name = "txtSDEmail";
			this.txtSDEmail.Size = new System.Drawing.Size(176, 22);
			this.txtSDEmail.TabIndex = 778;
			this.txtSDEmail.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtSDEmail_KeyDown);
			// 
			// txtSDCelular
			// 
			this.txtSDCelular.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtSDCelular.Enabled = false;
			this.txtSDCelular.Location = new System.Drawing.Point(696, 152);
			this.txtSDCelular.MaxLength = 10;
			this.txtSDCelular.Name = "txtSDCelular";
			this.txtSDCelular.Size = new System.Drawing.Size(80, 22);
			this.txtSDCelular.TabIndex = 777;
			this.txtSDCelular.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtSDCelular_KeyDown);
			this.txtSDCelular.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.txtSDCelular_KeyPress);
			// 
			// txtSDTelefono
			// 
			this.txtSDTelefono.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtSDTelefono.Enabled = false;
			this.txtSDTelefono.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.txtSDTelefono.Location = new System.Drawing.Point(520, 152);
			this.txtSDTelefono.MaxLength = 9;
			this.txtSDTelefono.Name = "txtSDTelefono";
			this.txtSDTelefono.Size = new System.Drawing.Size(80, 23);
			this.txtSDTelefono.TabIndex = 774;
			this.txtSDTelefono.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtSDTelefono_KeyDown);
			this.txtSDTelefono.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.txtSDTelefono_KeyPress);
			// 
			// label110
			// 
			this.label110.AutoSize = true;
			this.label110.BackColor = System.Drawing.SystemColors.Control;
			this.label110.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label110.Location = new System.Drawing.Point(520, 136);
			this.label110.Name = "label110";
			this.label110.Size = new System.Drawing.Size(63, 17);
			this.label110.TabIndex = 773;
			this.label110.Text = "Teléfono 1";
			this.label110.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// txtSDCalle2
			// 
			this.txtSDCalle2.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtSDCalle2.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtSDCalle2.Enabled = false;
			this.txtSDCalle2.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.txtSDCalle2.Location = new System.Drawing.Point(320, 152);
			this.txtSDCalle2.Name = "txtSDCalle2";
			this.txtSDCalle2.Size = new System.Drawing.Size(192, 23);
			this.txtSDCalle2.TabIndex = 772;
			this.txtSDCalle2.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtSDCalle2_KeyDown);
			this.txtSDCalle2.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.txtSDCalle2_KeyPress);
			// 
			// label111
			// 
			this.label111.AutoSize = true;
			this.label111.BackColor = System.Drawing.SystemColors.Control;
			this.label111.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label111.Location = new System.Drawing.Point(320, 136);
			this.label111.Name = "label111";
			this.label111.Size = new System.Drawing.Size(99, 17);
			this.label111.TabIndex = 771;
			this.label111.Text = "Calle Secundaria";
			this.label111.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// txtSDNumCasa
			// 
			this.txtSDNumCasa.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtSDNumCasa.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtSDNumCasa.Enabled = false;
			this.txtSDNumCasa.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.txtSDNumCasa.Location = new System.Drawing.Point(208, 152);
			this.txtSDNumCasa.Name = "txtSDNumCasa";
			this.txtSDNumCasa.Size = new System.Drawing.Size(104, 23);
			this.txtSDNumCasa.TabIndex = 770;
			this.txtSDNumCasa.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtSDNumCasa_KeyDown);
			this.txtSDNumCasa.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.txtSDNumCasa_KeyPress);
			// 
			// label112
			// 
			this.label112.AutoSize = true;
			this.label112.BackColor = System.Drawing.SystemColors.Control;
			this.label112.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label112.Location = new System.Drawing.Point(208, 136);
			this.label112.Name = "label112";
			this.label112.Size = new System.Drawing.Size(98, 17);
			this.label112.TabIndex = 769;
			this.label112.Text = "Número de Casa";
			this.label112.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// txtSDCalle1
			// 
			this.txtSDCalle1.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtSDCalle1.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtSDCalle1.Enabled = false;
			this.txtSDCalle1.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.txtSDCalle1.Location = new System.Drawing.Point(8, 152);
			this.txtSDCalle1.Name = "txtSDCalle1";
			this.txtSDCalle1.Size = new System.Drawing.Size(192, 23);
			this.txtSDCalle1.TabIndex = 768;
			this.txtSDCalle1.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtSDCalle1_KeyDown);
			this.txtSDCalle1.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.txtSDCalle1_KeyPress);
			// 
			// label113
			// 
			this.label113.AutoSize = true;
			this.label113.BackColor = System.Drawing.SystemColors.Control;
			this.label113.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label113.Location = new System.Drawing.Point(8, 136);
			this.label113.Name = "label113";
			this.label113.Size = new System.Drawing.Size(84, 17);
			this.label113.TabIndex = 767;
			this.label113.Text = "Calle Principal";
			this.label113.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// label114
			// 
			this.label114.AutoSize = true;
			this.label114.BackColor = System.Drawing.SystemColors.Control;
			this.label114.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label114.Location = new System.Drawing.Point(824, 136);
			this.label114.Name = "label114";
			this.label114.Size = new System.Drawing.Size(43, 17);
			this.label114.TabIndex = 766;
			this.label114.Text = "E-Mail ";
			this.label114.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// label115
			// 
			this.label115.AutoSize = true;
			this.label115.BackColor = System.Drawing.SystemColors.Control;
			this.label115.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label115.Location = new System.Drawing.Point(696, 136);
			this.label115.Name = "label115";
			this.label115.Size = new System.Drawing.Size(44, 17);
			this.label115.TabIndex = 765;
			this.label115.Text = "Celular";
			this.label115.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// label116
			// 
			this.label116.AutoSize = true;
			this.label116.BackColor = System.Drawing.SystemColors.Control;
			this.label116.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label116.Location = new System.Drawing.Point(8, 88);
			this.label116.Name = "label116";
			this.label116.Size = new System.Drawing.Size(55, 17);
			this.label116.TabIndex = 758;
			this.label116.Text = "Provincia";
			this.label116.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// cmbSDProvincia
			// 
			this.cmbSDProvincia.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbSDProvincia.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbSDProvincia.DataSource = this.ultraDataSource1;
			ultraGridColumn86.Header.VisiblePosition = 0;
			ultraGridColumn86.Hidden = true;
			ultraGridColumn87.Header.VisiblePosition = 1;
			ultraGridColumn87.Width = 176;
			ultraGridBand34.Columns.AddRange(new object[] {
																											ultraGridColumn86,
																											ultraGridColumn87});
			this.cmbSDProvincia.DisplayLayout.BandsSerializer.Add(ultraGridBand34);
			this.cmbSDProvincia.DisplayMember = "Nombre";
			this.cmbSDProvincia.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbSDProvincia.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbSDProvincia.Enabled = false;
			this.cmbSDProvincia.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbSDProvincia.Location = new System.Drawing.Point(8, 112);
			this.cmbSDProvincia.Name = "cmbSDProvincia";
			this.cmbSDProvincia.Size = new System.Drawing.Size(176, 22);
			this.cmbSDProvincia.TabIndex = 764;
			this.cmbSDProvincia.ValueMember = "idProvincia";
			this.cmbSDProvincia.KeyDown += new System.Windows.Forms.KeyEventHandler(this.cmbSDProvincia_KeyDown);
			this.cmbSDProvincia.ValueChanged += new System.EventHandler(this.cmbSDProvincia_ValueChanged);
			// 
			// label117
			// 
			this.label117.AutoSize = true;
			this.label117.BackColor = System.Drawing.SystemColors.Control;
			this.label117.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label117.Location = new System.Drawing.Point(192, 88);
			this.label117.Name = "label117";
			this.label117.Size = new System.Drawing.Size(44, 17);
			this.label117.TabIndex = 757;
			this.label117.Text = "Canton";
			this.label117.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// cmbSDCanton
			// 
			this.cmbSDCanton.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbSDCanton.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbSDCanton.DataSource = this.ultraDataSource2;
			ultraGridBand35.AddButtonCaption = "RolSucursal";
			ultraGridColumn88.Header.VisiblePosition = 0;
			ultraGridColumn88.Hidden = true;
			ultraGridColumn89.Header.VisiblePosition = 1;
			ultraGridColumn89.Width = 184;
			ultraGridBand35.Columns.AddRange(new object[] {
																											ultraGridColumn88,
																											ultraGridColumn89});
			this.cmbSDCanton.DisplayLayout.BandsSerializer.Add(ultraGridBand35);
			this.cmbSDCanton.DisplayLayout.Override.RowSelectors = Infragistics.Win.DefaultableBoolean.True;
			this.cmbSDCanton.DisplayMember = "Nombre";
			this.cmbSDCanton.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbSDCanton.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbSDCanton.Enabled = false;
			this.cmbSDCanton.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbSDCanton.Location = new System.Drawing.Point(192, 112);
			this.cmbSDCanton.Name = "cmbSDCanton";
			this.cmbSDCanton.Size = new System.Drawing.Size(184, 22);
			this.cmbSDCanton.TabIndex = 759;
			this.cmbSDCanton.ValueMember = "idCanton";
			this.cmbSDCanton.KeyDown += new System.Windows.Forms.KeyEventHandler(this.cmbSDCanton_KeyDown);
			this.cmbSDCanton.ValueChanged += new System.EventHandler(this.cmbSDCanton_ValueChanged);
			// 
			// label118
			// 
			this.label118.AutoSize = true;
			this.label118.BackColor = System.Drawing.SystemColors.Control;
			this.label118.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label118.Location = new System.Drawing.Point(384, 88);
			this.label118.Name = "label118";
			this.label118.Size = new System.Drawing.Size(58, 17);
			this.label118.TabIndex = 761;
			this.label118.Text = "Parroquia";
			this.label118.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// cmbSDParroquia
			// 
			this.cmbSDParroquia.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbSDParroquia.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbSDParroquia.DataSource = this.ultraDataSource9;
			ultraGridBand36.AddButtonCaption = "RolSucursal";
			ultraGridColumn90.Header.VisiblePosition = 0;
			ultraGridColumn90.Hidden = true;
			ultraGridColumn91.Header.VisiblePosition = 1;
			ultraGridColumn91.Width = 184;
			ultraGridBand36.Columns.AddRange(new object[] {
																											ultraGridColumn90,
																											ultraGridColumn91});
			this.cmbSDParroquia.DisplayLayout.BandsSerializer.Add(ultraGridBand36);
			this.cmbSDParroquia.DisplayMember = "Nombre";
			this.cmbSDParroquia.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbSDParroquia.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbSDParroquia.Enabled = false;
			this.cmbSDParroquia.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbSDParroquia.Location = new System.Drawing.Point(384, 112);
			this.cmbSDParroquia.Name = "cmbSDParroquia";
			this.cmbSDParroquia.Size = new System.Drawing.Size(184, 22);
			this.cmbSDParroquia.TabIndex = 762;
			this.cmbSDParroquia.ValueMember = "idParroquia";
			this.cmbSDParroquia.KeyDown += new System.Windows.Forms.KeyEventHandler(this.cmbSDParroquia_KeyDown);
			this.cmbSDParroquia.ValueChanged += new System.EventHandler(this.cmbSDParroquia_ValueChanged);
			// 
			// label119
			// 
			this.label119.AutoSize = true;
			this.label119.BackColor = System.Drawing.SystemColors.Control;
			this.label119.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label119.Location = new System.Drawing.Point(576, 88);
			this.label119.Name = "label119";
			this.label119.Size = new System.Drawing.Size(38, 17);
			this.label119.TabIndex = 760;
			this.label119.Text = "Barrio";
			this.label119.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// cmbSDBarrio
			// 
			this.cmbSDBarrio.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbSDBarrio.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbSDBarrio.DataSource = this.ultraDataSource10;
			ultraGridBand37.AddButtonCaption = "RolSucursal";
			ultraGridColumn92.Header.VisiblePosition = 0;
			ultraGridColumn92.Hidden = true;
			ultraGridColumn93.Header.VisiblePosition = 1;
			ultraGridColumn93.Width = 184;
			ultraGridBand37.Columns.AddRange(new object[] {
																											ultraGridColumn92,
																											ultraGridColumn93});
			this.cmbSDBarrio.DisplayLayout.BandsSerializer.Add(ultraGridBand37);
			this.cmbSDBarrio.DisplayLayout.Override.RowSelectors = Infragistics.Win.DefaultableBoolean.True;
			this.cmbSDBarrio.DisplayMember = "Nombre";
			this.cmbSDBarrio.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbSDBarrio.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbSDBarrio.Enabled = false;
			this.cmbSDBarrio.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbSDBarrio.Location = new System.Drawing.Point(576, 112);
			this.cmbSDBarrio.Name = "cmbSDBarrio";
			this.cmbSDBarrio.Size = new System.Drawing.Size(184, 22);
			this.cmbSDBarrio.TabIndex = 763;
			this.cmbSDBarrio.ValueMember = "idBarrio";
			this.cmbSDBarrio.KeyDown += new System.Windows.Forms.KeyEventHandler(this.cmbSDBarrio_KeyDown);
			// 
			// chkSDAfiliacion
			// 
			appearance25.ForeColorDisabled = System.Drawing.Color.Black;
			this.chkSDAfiliacion.Appearance = appearance25;
			this.chkSDAfiliacion.BackColor = System.Drawing.Color.Transparent;
			this.chkSDAfiliacion.CausesValidation = false;
			this.chkSDAfiliacion.Enabled = false;
			this.chkSDAfiliacion.Location = new System.Drawing.Point(880, 64);
			this.chkSDAfiliacion.Name = "chkSDAfiliacion";
			this.chkSDAfiliacion.Size = new System.Drawing.Size(104, 22);
			this.chkSDAfiliacion.TabIndex = 756;
			this.chkSDAfiliacion.Text = "Afiliado al IESS";
			this.chkSDAfiliacion.KeyDown += new System.Windows.Forms.KeyEventHandler(this.chkSDAfiliacion_KeyDown);
			// 
			// txtSDPago
			// 
			this.txtSDPago.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtSDPago.Enabled = false;
			this.txtSDPago.Location = new System.Drawing.Point(792, 64);
			this.txtSDPago.MaskInput = "nn";
			this.txtSDPago.MaxValue = 31;
			this.txtSDPago.MinValue = 0;
			this.txtSDPago.Name = "txtSDPago";
			this.txtSDPago.PromptChar = ' ';
			this.txtSDPago.Size = new System.Drawing.Size(80, 22);
			this.txtSDPago.SpinButtonDisplayStyle = Infragistics.Win.ButtonDisplayStyle.Always;
			this.txtSDPago.TabIndex = 755;
			this.txtSDPago.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtSDPago_KeyDown);
			this.txtSDPago.Click += new System.EventHandler(this.txtSDPago_Click);
			this.txtSDPago.Enter += new System.EventHandler(this.txtSDPago_Enter);
			// 
			// label108
			// 
			this.label108.AutoSize = true;
			this.label108.BackColor = System.Drawing.SystemColors.Control;
			this.label108.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label108.Location = new System.Drawing.Point(792, 48);
			this.label108.Name = "label108";
			this.label108.Size = new System.Drawing.Size(55, 17);
			this.label108.TabIndex = 754;
			this.label108.Text = "Día Pago";
			this.label108.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// txtSDDepartamento
			// 
			appearance26.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtSDDepartamento.Appearance = appearance26;
			this.txtSDDepartamento.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtSDDepartamento.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtSDDepartamento.Enabled = false;
			this.txtSDDepartamento.Location = new System.Drawing.Point(384, 64);
			this.txtSDDepartamento.Name = "txtSDDepartamento";
			this.txtSDDepartamento.Size = new System.Drawing.Size(248, 22);
			this.txtSDDepartamento.TabIndex = 753;
			this.txtSDDepartamento.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtSDDepartamento_KeyDown);
			// 
			// label107
			// 
			this.label107.AutoSize = true;
			this.label107.BackColor = System.Drawing.SystemColors.Control;
			this.label107.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label107.Location = new System.Drawing.Point(384, 48);
			this.label107.Name = "label107";
			this.label107.Size = new System.Drawing.Size(83, 17);
			this.label107.TabIndex = 752;
			this.label107.Text = "Departamento";
			this.label107.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// label104
			// 
			this.label104.AutoSize = true;
			this.label104.Location = new System.Drawing.Point(0, 48);
			this.label104.Name = "label104";
			this.label104.Size = new System.Drawing.Size(72, 17);
			this.label104.TabIndex = 751;
			this.label104.Text = "Tipo Contrato";
			this.label104.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
			// 
			// cmbSDCargo
			// 
			this.cmbSDCargo.CausesValidation = false;
			this.cmbSDCargo.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbSDCargo.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbSDCargo.DataSource = this.ultraDataSource15;
			ultraGridColumn94.Header.VisiblePosition = 0;
			ultraGridColumn94.Hidden = true;
			ultraGridColumn95.Header.VisiblePosition = 2;
			ultraGridColumn95.Width = 144;
			ultraGridColumn96.Header.VisiblePosition = 1;
			ultraGridColumn96.Hidden = true;
			ultraGridBand38.Columns.AddRange(new object[] {
																											ultraGridColumn94,
																											ultraGridColumn95,
																											ultraGridColumn96});
			this.cmbSDCargo.DisplayLayout.BandsSerializer.Add(ultraGridBand38);
			this.cmbSDCargo.DisplayMember = "Nombre";
			this.cmbSDCargo.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbSDCargo.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbSDCargo.Enabled = false;
			this.cmbSDCargo.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbSDCargo.Location = new System.Drawing.Point(640, 64);
			this.cmbSDCargo.MaxDropDownItems = 30;
			this.cmbSDCargo.Name = "cmbSDCargo";
			this.cmbSDCargo.Size = new System.Drawing.Size(144, 21);
			this.cmbSDCargo.TabIndex = 749;
			this.cmbSDCargo.ValueMember = "idCargo";
			this.cmbSDCargo.KeyDown += new System.Windows.Forms.KeyEventHandler(this.cmbSDCargo_KeyDown);
			// 
			// ultraDataSource15
			// 
			ultraDataColumn48.DataType = typeof(int);
			this.ultraDataSource15.Band.Columns.AddRange(new object[] {
																																	ultraDataColumn48,
																																	ultraDataColumn49});
			// 
			// cmbSDTipoSueldo
			// 
			this.cmbSDTipoSueldo.CausesValidation = false;
			this.cmbSDTipoSueldo.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbSDTipoSueldo.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbSDTipoSueldo.DataSource = this.ultraDataSource16;
			ultraGridColumn97.Header.VisiblePosition = 0;
			ultraGridColumn97.Hidden = true;
			ultraGridColumn98.Header.VisiblePosition = 1;
			ultraGridColumn98.Width = 185;
			ultraGridBand39.Columns.AddRange(new object[] {
																											ultraGridColumn97,
																											ultraGridColumn98});
			this.cmbSDTipoSueldo.DisplayLayout.BandsSerializer.Add(ultraGridBand39);
			this.cmbSDTipoSueldo.DisplayMember = "Nombre";
			this.cmbSDTipoSueldo.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbSDTipoSueldo.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbSDTipoSueldo.Enabled = false;
			this.cmbSDTipoSueldo.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbSDTipoSueldo.Location = new System.Drawing.Point(192, 64);
			this.cmbSDTipoSueldo.MaxDropDownItems = 30;
			this.cmbSDTipoSueldo.Name = "cmbSDTipoSueldo";
			this.cmbSDTipoSueldo.Size = new System.Drawing.Size(184, 21);
			this.cmbSDTipoSueldo.TabIndex = 748;
			this.cmbSDTipoSueldo.ValueMember = "idTipoSueldo";
			this.cmbSDTipoSueldo.KeyDown += new System.Windows.Forms.KeyEventHandler(this.cmbSDTipoSueldo_KeyDown);
			// 
			// ultraDataSource16
			// 
			ultraDataColumn50.DataType = typeof(int);
			this.ultraDataSource16.Band.Columns.AddRange(new object[] {
																																	ultraDataColumn50,
																																	ultraDataColumn51});
			// 
			// label105
			// 
			this.label105.AutoSize = true;
			this.label105.Location = new System.Drawing.Point(640, 48);
			this.label105.Name = "label105";
			this.label105.Size = new System.Drawing.Size(33, 17);
			this.label105.TabIndex = 747;
			this.label105.Text = "Cargo";
			this.label105.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// label106
			// 
			this.label106.AutoSize = true;
			this.label106.Location = new System.Drawing.Point(192, 48);
			this.label106.Name = "label106";
			this.label106.Size = new System.Drawing.Size(63, 17);
			this.label106.TabIndex = 746;
			this.label106.Text = "Tipo Sueldo";
			this.label106.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// txtSDPasivos
			// 
			this.txtSDPasivos.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtSDPasivos.Enabled = false;
			this.txtSDPasivos.Location = new System.Drawing.Point(936, 24);
			this.txtSDPasivos.MaxValue = 99;
			this.txtSDPasivos.MinValue = 0;
			this.txtSDPasivos.Name = "txtSDPasivos";
			this.txtSDPasivos.PromptChar = ' ';
			this.txtSDPasivos.Size = new System.Drawing.Size(64, 22);
			this.txtSDPasivos.SpinButtonDisplayStyle = Infragistics.Win.ButtonDisplayStyle.Always;
			this.txtSDPasivos.TabIndex = 745;
			this.txtSDPasivos.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtSDPasivos_KeyDown);
			this.txtSDPasivos.Click += new System.EventHandler(this.txtSDPasivos_Click);
			this.txtSDPasivos.Enter += new System.EventHandler(this.txtSDPasivos_Enter);
			// 
			// txtSDActivos
			// 
			this.txtSDActivos.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtSDActivos.Enabled = false;
			this.txtSDActivos.Location = new System.Drawing.Point(864, 24);
			this.txtSDActivos.MaxValue = 99;
			this.txtSDActivos.MinValue = 0;
			this.txtSDActivos.Name = "txtSDActivos";
			this.txtSDActivos.PromptChar = ' ';
			this.txtSDActivos.Size = new System.Drawing.Size(64, 22);
			this.txtSDActivos.SpinButtonDisplayStyle = Infragistics.Win.ButtonDisplayStyle.Always;
			this.txtSDActivos.TabIndex = 744;
			this.txtSDActivos.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtSDActivos_KeyDown);
			this.txtSDActivos.Click += new System.EventHandler(this.txtSDActivos_Click);
			this.txtSDActivos.Enter += new System.EventHandler(this.txtSDActivos_Enter);
			// 
			// txtSDGastos
			// 
			this.txtSDGastos.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtSDGastos.Enabled = false;
			this.txtSDGastos.Location = new System.Drawing.Point(792, 24);
			this.txtSDGastos.MaskInput = "nnnn";
			this.txtSDGastos.MaxValue = 9999;
			this.txtSDGastos.MinValue = 0;
			this.txtSDGastos.Name = "txtSDGastos";
			this.txtSDGastos.PromptChar = ' ';
			this.txtSDGastos.Size = new System.Drawing.Size(64, 22);
			this.txtSDGastos.SpinButtonDisplayStyle = Infragistics.Win.ButtonDisplayStyle.Always;
			this.txtSDGastos.TabIndex = 743;
			this.txtSDGastos.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtSDGastos_KeyDown);
			this.txtSDGastos.Click += new System.EventHandler(this.txtSDGastos_Click);
			this.txtSDGastos.Enter += new System.EventHandler(this.txtSDGastos_Enter);
			// 
			// txtSDIngresos
			// 
			this.txtSDIngresos.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtSDIngresos.Enabled = false;
			this.txtSDIngresos.Location = new System.Drawing.Point(720, 24);
			this.txtSDIngresos.MaskInput = "nnnn";
			this.txtSDIngresos.MaxValue = 9999;
			this.txtSDIngresos.MinValue = 0;
			this.txtSDIngresos.Name = "txtSDIngresos";
			this.txtSDIngresos.PromptChar = ' ';
			this.txtSDIngresos.Size = new System.Drawing.Size(64, 22);
			this.txtSDIngresos.SpinButtonDisplayStyle = Infragistics.Win.ButtonDisplayStyle.Always;
			this.txtSDIngresos.TabIndex = 742;
			this.txtSDIngresos.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtSDIngresos_KeyDown);
			this.txtSDIngresos.Click += new System.EventHandler(this.txtSDIngresos_Click);
			this.txtSDIngresos.Enter += new System.EventHandler(this.txtSDIngresos_Enter);
			// 
			// label99
			// 
			this.label99.AutoSize = true;
			this.label99.BackColor = System.Drawing.SystemColors.Control;
			this.label99.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label99.Location = new System.Drawing.Point(936, 8);
			this.label99.Name = "label99";
			this.label99.Size = new System.Drawing.Size(48, 17);
			this.label99.TabIndex = 741;
			this.label99.Text = "Pasivos";
			this.label99.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// label101
			// 
			this.label101.AutoSize = true;
			this.label101.BackColor = System.Drawing.SystemColors.Control;
			this.label101.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label101.Location = new System.Drawing.Point(864, 8);
			this.label101.Name = "label101";
			this.label101.Size = new System.Drawing.Size(44, 17);
			this.label101.TabIndex = 739;
			this.label101.Text = "Activos";
			this.label101.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// label102
			// 
			this.label102.AutoSize = true;
			this.label102.BackColor = System.Drawing.SystemColors.Control;
			this.label102.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label102.Location = new System.Drawing.Point(792, 8);
			this.label102.Name = "label102";
			this.label102.Size = new System.Drawing.Size(44, 17);
			this.label102.TabIndex = 737;
			this.label102.Text = "Gastos";
			this.label102.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// label103
			// 
			this.label103.AutoSize = true;
			this.label103.BackColor = System.Drawing.SystemColors.Control;
			this.label103.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label103.Location = new System.Drawing.Point(720, 8);
			this.label103.Name = "label103";
			this.label103.Size = new System.Drawing.Size(52, 17);
			this.label103.TabIndex = 735;
			this.label103.Text = "Ingresos";
			this.label103.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// txtSDEmpresa
			// 
			appearance27.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtSDEmpresa.Appearance = appearance27;
			this.txtSDEmpresa.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtSDEmpresa.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtSDEmpresa.Enabled = false;
			this.txtSDEmpresa.Location = new System.Drawing.Point(152, 24);
			this.txtSDEmpresa.Name = "txtSDEmpresa";
			this.txtSDEmpresa.Size = new System.Drawing.Size(208, 22);
			this.txtSDEmpresa.TabIndex = 716;
			this.txtSDEmpresa.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtSDEmpresa_KeyDown);
			// 
			// cmbSDTipoEmpresa
			// 
			this.cmbSDTipoEmpresa.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbSDTipoEmpresa.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbSDTipoEmpresa.DataSource = this.ultraDataSource14;
			ultraGridBand40.AddButtonCaption = "RolSucursal";
			ultraGridColumn99.Header.VisiblePosition = 0;
			ultraGridColumn99.Hidden = true;
			ultraGridColumn100.Header.VisiblePosition = 1;
			ultraGridColumn100.Width = 168;
			ultraGridBand40.Columns.AddRange(new object[] {
																											ultraGridColumn99,
																											ultraGridColumn100});
			this.cmbSDTipoEmpresa.DisplayLayout.BandsSerializer.Add(ultraGridBand40);
			this.cmbSDTipoEmpresa.DisplayLayout.Override.RowSelectors = Infragistics.Win.DefaultableBoolean.True;
			this.cmbSDTipoEmpresa.DisplayMember = "Descripción";
			this.cmbSDTipoEmpresa.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbSDTipoEmpresa.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbSDTipoEmpresa.Enabled = false;
			this.cmbSDTipoEmpresa.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbSDTipoEmpresa.Location = new System.Drawing.Point(368, 24);
			this.cmbSDTipoEmpresa.Name = "cmbSDTipoEmpresa";
			this.cmbSDTipoEmpresa.Size = new System.Drawing.Size(168, 22);
			this.cmbSDTipoEmpresa.TabIndex = 715;
			this.cmbSDTipoEmpresa.ValueMember = "idTipoEmpresa";
			this.cmbSDTipoEmpresa.KeyDown += new System.Windows.Forms.KeyEventHandler(this.cmbSDTipoEmpresa_KeyDown);
			// 
			// ultraDataSource14
			// 
			ultraDataColumn52.DataType = typeof(int);
			this.ultraDataSource14.Band.Columns.AddRange(new object[] {
																																	ultraDataColumn52,
																																	ultraDataColumn53});
			// 
			// lblempres
			// 
			this.lblempres.AutoSize = true;
			this.lblempres.BackColor = System.Drawing.SystemColors.Control;
			this.lblempres.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.lblempres.Location = new System.Drawing.Point(152, 8);
			this.lblempres.Name = "lblempres";
			this.lblempres.Size = new System.Drawing.Size(54, 17);
			this.lblempres.TabIndex = 711;
			this.lblempres.Text = "Empresa";
			this.lblempres.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// label100
			// 
			this.label100.AutoSize = true;
			this.label100.BackColor = System.Drawing.SystemColors.Control;
			this.label100.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label100.Location = new System.Drawing.Point(368, 8);
			this.label100.Name = "label100";
			this.label100.Size = new System.Drawing.Size(98, 17);
			this.label100.TabIndex = 710;
			this.label100.Text = "Tipo de Empresa";
			this.label100.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// txtSDRuc
			// 
			appearance28.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtSDRuc.Appearance = appearance28;
			this.txtSDRuc.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtSDRuc.Enabled = false;
			this.txtSDRuc.Location = new System.Drawing.Point(8, 24);
			this.txtSDRuc.MaxLength = 13;
			this.txtSDRuc.Name = "txtSDRuc";
			this.txtSDRuc.Size = new System.Drawing.Size(136, 22);
			this.txtSDRuc.TabIndex = 364;
			this.txtSDRuc.Visible = false;
			this.txtSDRuc.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtSDRuc_KeyDown);
			this.txtSDRuc.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.txtSDRuc_KeyPress);
			this.txtSDRuc.Validating += new System.ComponentModel.CancelEventHandler(this.txtSDRuc_Validating);
			// 
			// label98
			// 
			this.label98.AutoSize = true;
			this.label98.Location = new System.Drawing.Point(8, 8);
			this.label98.Name = "label98";
			this.label98.Size = new System.Drawing.Size(25, 17);
			this.label98.TabIndex = 363;
			this.label98.Text = "RUC";
			this.label98.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// utcSolicitud
			// 
			this.utcSolicitud.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			appearance29.BackColor = System.Drawing.SystemColors.Control;
			appearance29.BackColor2 = System.Drawing.SystemColors.Control;
			appearance29.ForeColorDisabled = System.Drawing.Color.Black;
			this.utcSolicitud.Appearance = appearance29;
			appearance30.BackColor = System.Drawing.SystemColors.Control;
			appearance30.BackColor2 = System.Drawing.SystemColors.Control;
			this.utcSolicitud.ClientAreaAppearance = appearance30;
			this.utcSolicitud.Controls.Add(this.ultraTabSharedControlsPage1);
			this.utcSolicitud.Controls.Add(this.utcDatosGenerales);
			this.utcSolicitud.Controls.Add(this.utpEstudios);
			this.utcSolicitud.Controls.Add(this.utpImagenes);
			this.utcSolicitud.Controls.Add(this.utpAdicional);
			this.utcSolicitud.Controls.Add(this.ultraTabPageControl1);
			this.utcSolicitud.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.utcSolicitud.Location = new System.Drawing.Point(8, 112);
			this.utcSolicitud.Name = "utcSolicitud";
			this.utcSolicitud.SharedControlsPage = this.ultraTabSharedControlsPage1;
			this.utcSolicitud.Size = new System.Drawing.Size(1104, 344);
			this.utcSolicitud.Style = Infragistics.Win.UltraWinTabControl.UltraTabControlStyle.Excel;
			this.utcSolicitud.TabIndex = 679;
			ultraTab1.TabPage = this.utcDatosGenerales;
			ultraTab1.Text = "Datos Generales";
			ultraTab2.TabPage = this.utpEstudios;
			ultraTab2.Text = "Datos Conyuge";
			ultraTab3.TabPage = this.utpAdicional;
			ultraTab3.Text = "Referencias";
			ultraTab4.TabPage = this.utpImagenes;
			ultraTab4.Text = "Sección A - Negocio";
			ultraTab5.TabPage = this.ultraTabPageControl1;
			ultraTab5.Text = "Sección B - Dependiente";
			this.utcSolicitud.Tabs.AddRange(new Infragistics.Win.UltraWinTabControl.UltraTab[] {
																																													 ultraTab1,
																																													 ultraTab2,
																																													 ultraTab3,
																																													 ultraTab4,
																																													 ultraTab5});
			this.utcSolicitud.ViewStyle = Infragistics.Win.UltraWinTabControl.ViewStyle.Office2003;
			// 
			// ultraTabSharedControlsPage1
			// 
			this.ultraTabSharedControlsPage1.Location = new System.Drawing.Point(-10000, -10000);
			this.ultraTabSharedControlsPage1.Name = "ultraTabSharedControlsPage1";
			this.ultraTabSharedControlsPage1.Size = new System.Drawing.Size(1102, 323);
			// 
			// label6
			// 
			this.label6.AutoSize = true;
			this.label6.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label6.Location = new System.Drawing.Point(8, 0);
			this.label6.Name = "label6";
			this.label6.Size = new System.Drawing.Size(34, 17);
			this.label6.TabIndex = 681;
			this.label6.Text = "Fecha";
			this.label6.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// dtFecha
			// 
			appearance31.ForeColorDisabled = System.Drawing.Color.Black;
			this.dtFecha.Appearance = appearance31;
			this.dtFecha.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			dateButton6.Caption = "Today";
			this.dtFecha.DateButtons.Add(dateButton6);
			this.dtFecha.Enabled = false;
			this.dtFecha.Format = "dd/MM/yyyy";
			this.dtFecha.Location = new System.Drawing.Point(8, 24);
			this.dtFecha.Name = "dtFecha";
			this.dtFecha.NonAutoSizeHeight = 23;
			this.dtFecha.Size = new System.Drawing.Size(128, 21);
			this.dtFecha.SpinButtonsVisible = true;
			this.dtFecha.TabIndex = 680;
			this.dtFecha.Value = ((object)(resources.GetObject("dtFecha.Value")));
			// 
			// cmbTipoRuc
			// 
			this.cmbTipoRuc.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbTipoRuc.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbTipoRuc.DisplayMember = "Nombre";
			this.cmbTipoRuc.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbTipoRuc.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbTipoRuc.Enabled = false;
			this.cmbTipoRuc.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbTipoRuc.Location = new System.Drawing.Point(152, 24);
			this.cmbTipoRuc.Name = "cmbTipoRuc";
			this.cmbTipoRuc.Size = new System.Drawing.Size(112, 22);
			this.cmbTipoRuc.TabIndex = 777;
			this.cmbTipoRuc.ValueMember = "idTipoDoc";
			// 
			// lblBodega
			// 
			this.lblBodega.AutoSize = true;
			this.lblBodega.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.lblBodega.Location = new System.Drawing.Point(152, 0);
			this.lblBodega.Name = "lblBodega";
			this.lblBodega.Size = new System.Drawing.Size(86, 17);
			this.lblBodega.TabIndex = 774;
			this.lblBodega.Text = "Tipo Documento";
			this.lblBodega.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// txtDocumento
			// 
			appearance32.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtDocumento.Appearance = appearance32;
			this.txtDocumento.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtDocumento.Enabled = false;
			this.txtDocumento.Location = new System.Drawing.Point(272, 24);
			this.txtDocumento.MaxLength = 10;
			this.txtDocumento.Name = "txtDocumento";
			this.txtDocumento.Size = new System.Drawing.Size(150, 21);
			this.txtDocumento.TabIndex = 776;
			this.txtDocumento.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtDocumento_KeyDown);
			this.txtDocumento.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.txtDocumento_KeyPress);
			this.txtDocumento.Validating += new System.ComponentModel.CancelEventHandler(this.txtDocumento_Validating);
			this.txtDocumento.Validated += new System.EventHandler(this.txtDocumento_Validated);
			this.txtDocumento.ValueChanged += new System.EventHandler(this.txtDocumento_ValueChanged);
			// 
			// label5
			// 
			this.label5.AutoSize = true;
			this.label5.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label5.Location = new System.Drawing.Point(912, 0);
			this.label5.Name = "label5";
			this.label5.Size = new System.Drawing.Size(91, 17);
			this.label5.TabIndex = 781;
			this.label5.Text = "Segundo Nombre";
			this.label5.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// label4
			// 
			this.label4.AutoSize = true;
			this.label4.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label4.Location = new System.Drawing.Point(752, 0);
			this.label4.Name = "label4";
			this.label4.Size = new System.Drawing.Size(80, 17);
			this.label4.TabIndex = 780;
			this.label4.Text = "Primer Nombre";
			this.label4.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// label3
			// 
			this.label3.AutoSize = true;
			this.label3.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label3.Location = new System.Drawing.Point(592, 0);
			this.label3.Name = "label3";
			this.label3.Size = new System.Drawing.Size(88, 17);
			this.label3.TabIndex = 779;
			this.label3.Text = "Apellido Materno";
			this.label3.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// label2
			// 
			this.label2.AutoSize = true;
			this.label2.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label2.Location = new System.Drawing.Point(432, 0);
			this.label2.Name = "label2";
			this.label2.Size = new System.Drawing.Size(85, 17);
			this.label2.TabIndex = 778;
			this.label2.Text = "Apellido Paterno";
			this.label2.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// txtSnombre
			// 
			appearance33.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtSnombre.Appearance = appearance33;
			this.txtSnombre.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtSnombre.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtSnombre.Enabled = false;
			this.txtSnombre.Location = new System.Drawing.Point(912, 24);
			this.txtSnombre.Name = "txtSnombre";
			this.txtSnombre.Size = new System.Drawing.Size(150, 21);
			this.txtSnombre.TabIndex = 785;
			this.txtSnombre.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtSnombre_KeyDown);
			this.txtSnombre.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.txtSnombre_KeyPress);
			// 
			// txtPnombre
			// 
			appearance34.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtPnombre.Appearance = appearance34;
			this.txtPnombre.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtPnombre.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtPnombre.Enabled = false;
			this.txtPnombre.Location = new System.Drawing.Point(752, 24);
			this.txtPnombre.Name = "txtPnombre";
			this.txtPnombre.Size = new System.Drawing.Size(150, 21);
			this.txtPnombre.TabIndex = 784;
			this.txtPnombre.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtPnombre_KeyDown);
			this.txtPnombre.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.txtPnombre_KeyPress);
			// 
			// txtApellidoM
			// 
			appearance35.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtApellidoM.Appearance = appearance35;
			this.txtApellidoM.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtApellidoM.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtApellidoM.Enabled = false;
			this.txtApellidoM.Location = new System.Drawing.Point(592, 24);
			this.txtApellidoM.Name = "txtApellidoM";
			this.txtApellidoM.Size = new System.Drawing.Size(150, 21);
			this.txtApellidoM.TabIndex = 783;
			this.txtApellidoM.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtApellidoM_KeyDown);
			this.txtApellidoM.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.txtApellidoM_KeyPress);
			// 
			// txtApellidoP
			// 
			appearance36.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtApellidoP.Appearance = appearance36;
			this.txtApellidoP.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtApellidoP.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtApellidoP.Enabled = false;
			this.txtApellidoP.Location = new System.Drawing.Point(432, 24);
			this.txtApellidoP.Name = "txtApellidoP";
			this.txtApellidoP.Size = new System.Drawing.Size(150, 21);
			this.txtApellidoP.TabIndex = 782;
			this.txtApellidoP.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtApellidoP_KeyDown);
			this.txtApellidoP.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.txtApellidoP_KeyPress);
			// 
			// cmbCalificacion
			// 
			this.cmbCalificacion.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbCalificacion.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbCalificacion.DataSource = this.ultraDataSource6;
			ultraGridColumn101.Header.VisiblePosition = 0;
			ultraGridColumn101.Hidden = true;
			ultraGridColumn102.Header.Caption = "Calificacion";
			ultraGridColumn102.Header.VisiblePosition = 1;
			ultraGridBand41.Columns.AddRange(new object[] {
																											ultraGridColumn101,
																											ultraGridColumn102});
			this.cmbCalificacion.DisplayLayout.BandsSerializer.Add(ultraGridBand41);
			this.cmbCalificacion.DisplayLayout.Override.RowSelectors = Infragistics.Win.DefaultableBoolean.True;
			this.cmbCalificacion.DisplayMember = "Codigo";
			this.cmbCalificacion.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbCalificacion.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbCalificacion.Enabled = false;
			this.cmbCalificacion.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbCalificacion.Location = new System.Drawing.Point(120, 72);
			this.cmbCalificacion.Name = "cmbCalificacion";
			this.cmbCalificacion.Size = new System.Drawing.Size(104, 22);
			this.cmbCalificacion.TabIndex = 791;
			this.cmbCalificacion.ValueMember = "idTipoCalificacion";
			this.cmbCalificacion.KeyDown += new System.Windows.Forms.KeyEventHandler(this.cmbCalificacion_KeyDown);
			// 
			// ultraDataSource6
			// 
			ultraDataColumn54.DataType = typeof(int);
			this.ultraDataSource6.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn54,
																																 ultraDataColumn55});
			// 
			// cmbEstado
			// 
			this.cmbEstado.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbEstado.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbEstado.DataSource = this.ultraDataSource7;
			ultraGridColumn103.Header.VisiblePosition = 0;
			ultraGridColumn103.Hidden = true;
			ultraGridColumn104.Header.VisiblePosition = 1;
			ultraGridBand42.Columns.AddRange(new object[] {
																											ultraGridColumn103,
																											ultraGridColumn104});
			this.cmbEstado.DisplayLayout.BandsSerializer.Add(ultraGridBand42);
			this.cmbEstado.DisplayLayout.Override.RowSelectors = Infragistics.Win.DefaultableBoolean.True;
			this.cmbEstado.DisplayMember = "Estado";
			this.cmbEstado.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbEstado.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbEstado.Enabled = false;
			this.cmbEstado.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbEstado.Location = new System.Drawing.Point(8, 72);
			this.cmbEstado.Name = "cmbEstado";
			this.cmbEstado.Size = new System.Drawing.Size(104, 22);
			this.cmbEstado.TabIndex = 790;
			this.cmbEstado.ValueMember = "idEstado";
			this.cmbEstado.KeyDown += new System.Windows.Forms.KeyEventHandler(this.cmbEstado_KeyDown);
			// 
			// ultraDataSource7
			// 
			ultraDataColumn56.DataType = typeof(int);
			this.ultraDataSource7.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn56,
																																 ultraDataColumn57});
			// 
			// label16
			// 
			this.label16.AutoSize = true;
			this.label16.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label16.Location = new System.Drawing.Point(8, 48);
			this.label16.Name = "label16";
			this.label16.Size = new System.Drawing.Size(38, 17);
			this.label16.TabIndex = 789;
			this.label16.Text = "Estado";
			this.label16.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// label17
			// 
			this.label17.AutoSize = true;
			this.label17.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label17.Location = new System.Drawing.Point(232, 48);
			this.label17.Name = "label17";
			this.label17.Size = new System.Drawing.Size(32, 17);
			this.label17.TabIndex = 787;
			this.label17.Text = "Score";
			this.label17.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// label18
			// 
			this.label18.AutoSize = true;
			this.label18.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label18.Location = new System.Drawing.Point(120, 48);
			this.label18.Name = "label18";
			this.label18.Size = new System.Drawing.Size(60, 17);
			this.label18.TabIndex = 786;
			this.label18.Text = "Calificación";
			this.label18.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// txtScore
			// 
			this.txtScore.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtScore.Enabled = false;
			this.txtScore.FormatString = "#,##0";
			this.txtScore.Location = new System.Drawing.Point(232, 72);
			this.txtScore.MaskInput = "nnn";
			this.txtScore.MaxValue = 999;
			this.txtScore.MinValue = 0;
			this.txtScore.Name = "txtScore";
			this.txtScore.PromptChar = ' ';
			this.txtScore.Size = new System.Drawing.Size(80, 21);
			this.txtScore.SpinButtonDisplayStyle = Infragistics.Win.ButtonDisplayStyle.Always;
			this.txtScore.TabIndex = 788;
			this.txtScore.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtScore_KeyDown);
			this.txtScore.Click += new System.EventHandler(this.txtScore_Click);
			this.txtScore.Enter += new System.EventHandler(this.txtScore_Enter);
			// 
			// btnSalir
			// 
			this.btnSalir.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btnSalir.CausesValidation = false;
			this.btnSalir.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnSalir.ImageIndex = 62;
			this.btnSalir.Location = new System.Drawing.Point(96, 472);
			this.btnSalir.Name = "btnSalir";
			this.btnSalir.TabIndex = 794;
			this.btnSalir.Text = "Cancelar";
			this.btnSalir.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnSalir.Click += new System.EventHandler(this.btnSalir_Click);
			// 
			// btnGuardar
			// 
			this.btnGuardar.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btnGuardar.Image = ((System.Drawing.Image)(resources.GetObject("btnGuardar.Image")));
			this.btnGuardar.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnGuardar.Location = new System.Drawing.Point(8, 472);
			this.btnGuardar.Name = "btnGuardar";
			this.btnGuardar.Size = new System.Drawing.Size(76, 23);
			this.btnGuardar.TabIndex = 792;
			this.btnGuardar.Text = "&Guardar";
			this.btnGuardar.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnGuardar.Click += new System.EventHandler(this.btnGuardar_Click);
			// 
			// txtClienteIdCre_DatosUbicacion
			// 
			this.txtClienteIdCre_DatosUbicacion.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtClienteIdCre_DatosUbicacion.Enabled = false;
			this.txtClienteIdCre_DatosUbicacion.Location = new System.Drawing.Point(536, 72);
			this.txtClienteIdCre_DatosUbicacion.Name = "txtClienteIdCre_DatosUbicacion";
			this.txtClienteIdCre_DatosUbicacion.PromptChar = ' ';
			this.txtClienteIdCre_DatosUbicacion.Size = new System.Drawing.Size(64, 21);
			this.txtClienteIdCre_DatosUbicacion.SpinButtonAlignment = Infragistics.Win.ButtonAlignment.Left;
			this.txtClienteIdCre_DatosUbicacion.TabIndex = 798;
			this.txtClienteIdCre_DatosUbicacion.Visible = false;
			// 
			// txtIdCre_Conyugue
			// 
			this.txtIdCre_Conyugue.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtIdCre_Conyugue.Enabled = false;
			this.txtIdCre_Conyugue.Location = new System.Drawing.Point(608, 72);
			this.txtIdCre_Conyugue.Name = "txtIdCre_Conyugue";
			this.txtIdCre_Conyugue.PromptChar = ' ';
			this.txtIdCre_Conyugue.Size = new System.Drawing.Size(64, 21);
			this.txtIdCre_Conyugue.SpinButtonAlignment = Infragistics.Win.ButtonAlignment.Left;
			this.txtIdCre_Conyugue.TabIndex = 796;
			this.txtIdCre_Conyugue.Visible = false;
			// 
			// txtClienteIdCre_DatosGenerales
			// 
			this.txtClienteIdCre_DatosGenerales.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtClienteIdCre_DatosGenerales.Enabled = false;
			this.txtClienteIdCre_DatosGenerales.Location = new System.Drawing.Point(472, 72);
			this.txtClienteIdCre_DatosGenerales.Name = "txtClienteIdCre_DatosGenerales";
			this.txtClienteIdCre_DatosGenerales.PromptChar = ' ';
			this.txtClienteIdCre_DatosGenerales.Size = new System.Drawing.Size(56, 21);
			this.txtClienteIdCre_DatosGenerales.SpinButtonAlignment = Infragistics.Win.ButtonAlignment.Left;
			this.txtClienteIdCre_DatosGenerales.TabIndex = 795;
			this.txtClienteIdCre_DatosGenerales.Visible = false;
			// 
			// txtNegocioIdCre_DatosUbicacion
			// 
			this.txtNegocioIdCre_DatosUbicacion.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtNegocioIdCre_DatosUbicacion.Enabled = false;
			this.txtNegocioIdCre_DatosUbicacion.Location = new System.Drawing.Point(824, 72);
			this.txtNegocioIdCre_DatosUbicacion.Name = "txtNegocioIdCre_DatosUbicacion";
			this.txtNegocioIdCre_DatosUbicacion.PromptChar = ' ';
			this.txtNegocioIdCre_DatosUbicacion.Size = new System.Drawing.Size(64, 21);
			this.txtNegocioIdCre_DatosUbicacion.SpinButtonAlignment = Infragistics.Win.ButtonAlignment.Left;
			this.txtNegocioIdCre_DatosUbicacion.TabIndex = 804;
			this.txtNegocioIdCre_DatosUbicacion.Visible = false;
			// 
			// txtClienteIdCre_EmpresasTrabajo
			// 
			this.txtClienteIdCre_EmpresasTrabajo.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtClienteIdCre_EmpresasTrabajo.Enabled = false;
			this.txtClienteIdCre_EmpresasTrabajo.Location = new System.Drawing.Point(896, 72);
			this.txtClienteIdCre_EmpresasTrabajo.Name = "txtClienteIdCre_EmpresasTrabajo";
			this.txtClienteIdCre_EmpresasTrabajo.PromptChar = ' ';
			this.txtClienteIdCre_EmpresasTrabajo.Size = new System.Drawing.Size(64, 21);
			this.txtClienteIdCre_EmpresasTrabajo.SpinButtonAlignment = Infragistics.Win.ButtonAlignment.Left;
			this.txtClienteIdCre_EmpresasTrabajo.TabIndex = 803;
			this.txtClienteIdCre_EmpresasTrabajo.Visible = false;
			// 
			// txtConyugueIdCre_DatosGenerales
			// 
			this.txtConyugueIdCre_DatosGenerales.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtConyugueIdCre_DatosGenerales.Enabled = false;
			this.txtConyugueIdCre_DatosGenerales.Location = new System.Drawing.Point(680, 72);
			this.txtConyugueIdCre_DatosGenerales.Name = "txtConyugueIdCre_DatosGenerales";
			this.txtConyugueIdCre_DatosGenerales.PromptChar = ' ';
			this.txtConyugueIdCre_DatosGenerales.Size = new System.Drawing.Size(64, 21);
			this.txtConyugueIdCre_DatosGenerales.SpinButtonAlignment = Infragistics.Win.ButtonAlignment.Left;
			this.txtConyugueIdCre_DatosGenerales.TabIndex = 802;
			this.txtConyugueIdCre_DatosGenerales.Visible = false;
			// 
			// txtClienteIdCre_Negocio
			// 
			this.txtClienteIdCre_Negocio.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtClienteIdCre_Negocio.Enabled = false;
			this.txtClienteIdCre_Negocio.Location = new System.Drawing.Point(752, 72);
			this.txtClienteIdCre_Negocio.Name = "txtClienteIdCre_Negocio";
			this.txtClienteIdCre_Negocio.PromptChar = ' ';
			this.txtClienteIdCre_Negocio.Size = new System.Drawing.Size(64, 21);
			this.txtClienteIdCre_Negocio.SpinButtonAlignment = Infragistics.Win.ButtonAlignment.Left;
			this.txtClienteIdCre_Negocio.TabIndex = 800;
			this.txtClienteIdCre_Negocio.Visible = false;
			// 
			// txtEmpresaIdCre_DatosUbicacion
			// 
			this.txtEmpresaIdCre_DatosUbicacion.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtEmpresaIdCre_DatosUbicacion.Enabled = false;
			this.txtEmpresaIdCre_DatosUbicacion.Location = new System.Drawing.Point(968, 72);
			this.txtEmpresaIdCre_DatosUbicacion.Name = "txtEmpresaIdCre_DatosUbicacion";
			this.txtEmpresaIdCre_DatosUbicacion.PromptChar = ' ';
			this.txtEmpresaIdCre_DatosUbicacion.Size = new System.Drawing.Size(64, 21);
			this.txtEmpresaIdCre_DatosUbicacion.SpinButtonAlignment = Infragistics.Win.ButtonAlignment.Left;
			this.txtEmpresaIdCre_DatosUbicacion.TabIndex = 805;
			this.txtEmpresaIdCre_DatosUbicacion.Visible = false;
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
			// txtIdCre_Garante
			// 
			this.txtIdCre_Garante.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtIdCre_Garante.Enabled = false;
			this.txtIdCre_Garante.Location = new System.Drawing.Point(328, 72);
			this.txtIdCre_Garante.Name = "txtIdCre_Garante";
			this.txtIdCre_Garante.PromptChar = ' ';
			this.txtIdCre_Garante.Size = new System.Drawing.Size(64, 21);
			this.txtIdCre_Garante.SpinButtonAlignment = Infragistics.Win.ButtonAlignment.Left;
			this.txtIdCre_Garante.TabIndex = 806;
			this.txtIdCre_Garante.Visible = false;
			// 
			// label1
			// 
			this.label1.AutoSize = true;
			this.label1.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label1.Location = new System.Drawing.Point(272, 0);
			this.label1.Name = "label1";
			this.label1.Size = new System.Drawing.Size(61, 17);
			this.label1.TabIndex = 807;
			this.label1.Text = "Documento";
			this.label1.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// txtIdCre_Solicitud
			// 
			this.txtIdCre_Solicitud.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtIdCre_Solicitud.Enabled = false;
			this.txtIdCre_Solicitud.Location = new System.Drawing.Point(400, 72);
			this.txtIdCre_Solicitud.Name = "txtIdCre_Solicitud";
			this.txtIdCre_Solicitud.PromptChar = ' ';
			this.txtIdCre_Solicitud.Size = new System.Drawing.Size(56, 21);
			this.txtIdCre_Solicitud.SpinButtonAlignment = Infragistics.Win.ButtonAlignment.Left;
			this.txtIdCre_Solicitud.TabIndex = 819;
			this.txtIdCre_Solicitud.Visible = false;
			// 
			// frmCred_GaranteSolicitud
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 13);
			this.ClientSize = new System.Drawing.Size(1119, 501);
			this.Controls.Add(this.txtIdCre_Solicitud);
			this.Controls.Add(this.label1);
			this.Controls.Add(this.label16);
			this.Controls.Add(this.label17);
			this.Controls.Add(this.label18);
			this.Controls.Add(this.label5);
			this.Controls.Add(this.label4);
			this.Controls.Add(this.label3);
			this.Controls.Add(this.label2);
			this.Controls.Add(this.lblBodega);
			this.Controls.Add(this.label6);
			this.Controls.Add(this.txtIdCre_Garante);
			this.Controls.Add(this.txtEmpresaIdCre_DatosUbicacion);
			this.Controls.Add(this.txtNegocioIdCre_DatosUbicacion);
			this.Controls.Add(this.txtClienteIdCre_EmpresasTrabajo);
			this.Controls.Add(this.txtConyugueIdCre_DatosGenerales);
			this.Controls.Add(this.txtClienteIdCre_Negocio);
			this.Controls.Add(this.txtClienteIdCre_DatosUbicacion);
			this.Controls.Add(this.txtIdCre_Conyugue);
			this.Controls.Add(this.txtClienteIdCre_DatosGenerales);
			this.Controls.Add(this.btnSalir);
			this.Controls.Add(this.btnGuardar);
			this.Controls.Add(this.cmbCalificacion);
			this.Controls.Add(this.cmbEstado);
			this.Controls.Add(this.txtScore);
			this.Controls.Add(this.txtSnombre);
			this.Controls.Add(this.txtPnombre);
			this.Controls.Add(this.txtApellidoM);
			this.Controls.Add(this.txtApellidoP);
			this.Controls.Add(this.cmbTipoRuc);
			this.Controls.Add(this.txtDocumento);
			this.Controls.Add(this.dtFecha);
			this.Controls.Add(this.utcSolicitud);
			this.KeyPreview = true;
			this.Name = "frmCred_GaranteSolicitud";
			this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
			this.Text = "Garantes";
			this.KeyDown += new System.Windows.Forms.KeyEventHandler(this.frmCred_GaranteSolicitud_KeyDown);
			this.Load += new System.EventHandler(this.frmCred_GaranteSolicitud_Load);
			this.utcDatosGenerales.ResumeLayout(false);
			((System.ComponentModel.ISupportInitialize)(this.cmbSexo)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource17)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbNivel)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource19)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbNENacionalidad)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource26)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbCantonNacimiento)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource2)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbProvinciaNacimiento)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource1)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbSituacionLaboral)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource5)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbNEEstadoCivil)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource4)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.dtpNEFechaNacimiento)).EndInit();
			this.groupBox1.ResumeLayout(false);
			((System.ComponentModel.ISupportInitialize)(this.cmbBarrioDomicilio)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource10)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbParroquiaDomicilio)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource9)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbCantonDomicilio)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbProvinciaDomicilio)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.dtpTiempoVivienda)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbTipoVivienda)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource11)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtReferencia)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNEEMail)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNECelular)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNETelefono2)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNETelefono1)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtCalle2)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNumcasa)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtCalle1)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbProfesion)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource3)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNECargasFamiliares)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbActividadEcomica)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource8)).EndInit();
			this.utpEstudios.ResumeLayout(false);
			this.groupBox3.ResumeLayout(false);
			((System.ComponentModel.ISupportInitialize)(this.cmbCProfesion)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbCSexo)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbCTipoRuc)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbCNivel)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbCNacionalidad)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbCActividadEconomia)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbCSituacionLaboral)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbCProvinciaNacimiento)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbCCantonNacimiento)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.dpCFechaNacimiento)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtCSNombre)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtCPNombre)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtCApellidoM)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtCApellidoP)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtCDocumento)).EndInit();
			this.utpAdicional.ResumeLayout(false);
			((System.ComponentModel.ISupportInitialize)(this.uGridParentesco)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource20)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbRProvincia)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbRCanton)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbParentesco)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource18)).EndInit();
			this.utpImagenes.ResumeLayout(false);
			((System.ComponentModel.ISupportInitialize)(this.cmbSubActvidadEconomica)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource12)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtSNPasivos)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtSNActivos)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtSNGastos)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtSNReferencia)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtSNIngresos)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtSNEmail)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtSNCelular)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtSNTelefono2)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtSNTelefono1)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtSNCalle2)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtSNNumCasa)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtSNCalle1)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbSNProvincia)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbSNCanton)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbSNParroquia)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbSNBarrio)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbSCTipoNegocio)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtSNEmpleados)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtSNMetros)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.dtSNFechaInicio)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtSNNombre)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtSNRuc)).EndInit();
			this.ultraTabPageControl1.ResumeLayout(false);
			((System.ComponentModel.ISupportInitialize)(this.dtSDFecha)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbSDTipoContrato)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource13)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtSDReferencia)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtSDExt)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtSDEmail)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtSDCelular)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtSDTelefono)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtSDCalle2)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtSDNumCasa)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtSDCalle1)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbSDProvincia)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbSDCanton)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbSDParroquia)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbSDBarrio)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtSDPago)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtSDDepartamento)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbSDCargo)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource15)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbSDTipoSueldo)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource16)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtSDPasivos)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtSDActivos)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtSDGastos)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtSDIngresos)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtSDEmpresa)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbSDTipoEmpresa)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource14)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtSDRuc)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.utcSolicitud)).EndInit();
			this.utcSolicitud.ResumeLayout(false);
			((System.ComponentModel.ISupportInitialize)(this.dtFecha)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbTipoRuc)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtDocumento)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtSnombre)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtPnombre)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtApellidoM)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtApellidoP)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbCalificacion)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource6)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbEstado)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource7)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtScore)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtClienteIdCre_DatosUbicacion)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIdCre_Conyugue)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtClienteIdCre_DatosGenerales)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNegocioIdCre_DatosUbicacion)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtClienteIdCre_EmpresasTrabajo)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtConyugueIdCre_DatosGenerales)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtClienteIdCre_Negocio)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtEmpresaIdCre_DatosUbicacion)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIdCre_Garante)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIdCre_Solicitud)).EndInit();
			this.ResumeLayout(false);

		}
		#endregion

		private void cdsSeteoF_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			cdsSeteoF.Schema.Connections[0].ConnectionString = miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);	
		}

		private void frmCred_GaranteSolicitud_Load(object sender, System.EventArgs e)
		{
			bCarga= true;
			this.dtFecha.Value = DateTime.Today;
			this.txtDocumento.Focus();

			#region CargaListas

			#region Carga_Combos

			this.cmbProvinciaNacimiento.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "Exec Listas 0, 0, 0");
			this.cmbProvinciaDomicilio.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "Exec Listas 0, 0");
			this.cmbTipoRuc.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "Exec Listas 4, 0, 0");
			this.cmbTipoRuc.Value = 1;			
			this.cmbCalificacion.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "Exec Listas 8, 0, 0");
			this.cmbSexo.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "Exec Listas 10, 0, 0");
			this.cmbEstado.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "Exec Listas 9, 0, 0");
			this.cmbNENacionalidad.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "Exec Listas 11, 0, 0");
			this.cmbNEEstadoCivil.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "Exec Listas 12, 0, 0");
			this.cmbNivel.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "Exec Listas 13, 0");
			this.cmbProfesion.DataSource =Funcion.dvProcedimiento(cdsSeteoF, "Exec Listas 14, 0, 0");
			//this.cmbActividadEcomica.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "Exec Listas 16, 0");
			this.cmbTipoVivienda.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "Exec Listas 17, 0");

			if ( ientidad > 0)
			{
				this.cmbSituacionLaboral.DataSource = Funcion.dvProcedimiento(cdsSeteoF, string.Format( "Exec Listas 66, {0},0", ientidad));
				if ( ientidad == 2)
				{
					this.label56.Text = "Fuente de Ingreso";
					this.cmbActividadEcomica.Value = 0;
				}
				else
				{
					this.label56.Text = "Situación Laboral";
					this.cmbActividadEcomica.Value = 0;
				}
			}
			#endregion 

			#region Conyugue

			this.cmbCProvinciaNacimiento.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "Exec Listas 0, 0, 0");
			this.cmbCTipoRuc.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "Exec Listas 4, 0, 0");
			this.cmbCTipoRuc.Value = 1;		
			this.cmbCSexo.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "Exec Listas 10, 0, 0");
			this.cmbCNacionalidad.DataSource =Funcion.dvProcedimiento(cdsSeteoF, "Exec Listas 11, 0, 0");
			this.cmbCNivel.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "Exec Listas 13, 0, 0");
			this.cmbCProfesion.DataSource =Funcion.dvProcedimiento(cdsSeteoF, "Exec Listas 14, 0, 0");
//			this.cmbCSituacionLaboral.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "Exec Listas 18, 0, 0");
//			this.cmbCActividadEconomia.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "Exec Listas 16, 0, 0");


			if ( ientidad > 0)
			{
				this.cmbCSituacionLaboral.DataSource = Funcion.dvProcedimiento(cdsSeteoF, string.Format( "Exec Listas 66, {0},0", ientidad));
				if ( ientidad ==2)
				{
					this.label76.Text = "Fuente de Ingreso";
					this.cmbCActividadEconomia.Value = 0;
				}
				else
				{
					this.label76.Text = "Situación Laboral";
					this.cmbCActividadEconomia.Value = 0;
				}
			}

			#endregion 

			#region referencias

			this.cmbParentesco.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "Exec Listas 29, 0, 0");
			this.cmbRProvincia.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "Exec Listas 0, 0, 0");
			this.cmbRCanton.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "Exec Listas 1, 0, 0");

			#endregion 

			#region Negocio
			
			this.cmbSNProvincia.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "Exec Listas 0, 0, 0");
			this.cmbSCTipoNegocio.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "Exec Listas 17, 0, 0");
			this.cmbSubActvidadEconomica.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "Exec Listas 21, 0, 0");

			#endregion 

			#region Trabajo

			this.cmbSDTipoEmpresa.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "Exec Listas 28, 0, 0");
			this.cmbSDTipoContrato.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "Exec Listas 22, 0, 0"); 
			this.cmbSDTipoSueldo.DataSource  =  Funcion.dvProcedimiento(cdsSeteoF, "Exec Listas 23, 0, 0");
			this.cmbSDCargo.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "Exec Listas 20, 0, 0");
			this.cmbSDProvincia.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "Exec Listas 0, 0, 0");

			#endregion 

			#endregion CargaListas

			habilita_Control_Garante();

			//if (bPantGarante)
			//{
				if (idCre_Solicitud > 0)
				{
				
					string ssqlDo= string.Format ("Exec Cre_ConsultaGarantes {0}", idCre_Solicitud); 
					SqlDataReader drDO = Funcion.rEscalarSQL(cdsSeteoF,ssqlDo,true);
					drDO.Read();
					if(drDO.HasRows)
					{
						this.txtIdCre_Garante.Value = drDO.GetInt32 (0);
						this.txtIdCre_Solicitud.Value = drDO.GetInt32 (1);
					}
					drDO.Close();	
					Consulta((int) this.txtIdCre_Garante.Value, (int) this.txtIdCre_Solicitud.Value);
				}
			//}
//			else
//			{
//				Consulta(isolicitud, idCre_Solicitud);
//			}
			if (!bgarante ) habilita_Control_Garante();
			bCarga = false;
		}

		public void habilita_Control_Garante()
		{

			this.txtDocumento.Enabled = true;
			this.txtApellidoP.Enabled = true;
			this.txtApellidoM.Enabled = true;
			this.txtPnombre.Enabled = true;
			this.txtSnombre.Enabled = true;
			this.cmbEstado.Enabled = true;
			this.cmbCalificacion.Enabled = true;
			this.txtScore.Enabled = true;
			this.cmbNENacionalidad.Enabled = true;
			this.dtpNEFechaNacimiento.Enabled = true;
			this.cmbSexo.Enabled= true;
			this.cmbProvinciaNacimiento.Enabled = true;
			this.cmbCantonNacimiento.Enabled = true;
			this.cmbNEEstadoCivil.Enabled = true;
			this.txtNECargasFamiliares.Enabled = true;
			this.cmbNivel.Enabled = true;
			this.cmbProfesion.Enabled = true;
			this.cmbSituacionLaboral.Enabled = true;
			this.cmbActividadEcomica.Enabled = true;
			this.cmbProvinciaDomicilio.Enabled = true;
			this.cmbCantonDomicilio.Enabled = true;
			this.cmbParroquiaDomicilio.Enabled = true;
			this.cmbBarrioDomicilio.Enabled = true;
			this.txtCalle1.Enabled = true;
			this.txtNumcasa.Enabled = true;
			this.txtCalle2.Enabled = true;
			this.txtNETelefono1.Enabled = true;
			this.txtNETelefono2.Enabled = true;
			this.txtNECelular.Enabled = true;
			this.txtNEEMail.Enabled = true;
			this.cmbTipoVivienda.Enabled = true;
			this.dtpTiempoVivienda.Enabled = true;
			this.txtReferencia.Enabled = true;
			this.uGridParentesco.Enabled = true;
			this.btnTelf.Enabled= true;
			this.btnTelf2.Enabled = true;
			this.btnCelular.Enabled = true;
		}

		public void deshabilita_Control_Garante()
		{

			this.txtDocumento.Enabled = false;
			this.txtApellidoP.Enabled = false;
			this.txtApellidoM.Enabled = false;
			this.txtPnombre.Enabled = false;
			this.txtSnombre.Enabled = false;
			this.cmbEstado.Enabled = false;
			this.cmbCalificacion.Enabled = false;
			this.txtScore.Enabled = false;
			this.cmbNENacionalidad.Enabled = false;
			this.dtpNEFechaNacimiento.Enabled = false;
			this.cmbSexo.Enabled= false;
			this.cmbProvinciaNacimiento.Enabled = false;
			this.cmbCantonNacimiento.Enabled = false;
			this.cmbNEEstadoCivil.Enabled = false;
			this.txtNECargasFamiliares.Enabled = false;
			this.cmbNivel.Enabled = false;
			this.cmbProfesion.Enabled = false;
			this.cmbSituacionLaboral.Enabled = false;
			this.cmbActividadEcomica.Enabled = false;
			this.cmbProvinciaDomicilio.Enabled = false;
			this.cmbCantonDomicilio.Enabled = false;
			this.cmbParroquiaDomicilio.Enabled = false;
			this.cmbBarrioDomicilio.Enabled = false;
			this.txtCalle1.Enabled = false;
			this.txtNumcasa.Enabled = false;
			this.txtCalle2.Enabled = false;
			this.txtNETelefono1.Enabled = false;
			this.txtNETelefono2.Enabled = false;
			this.txtNECelular.Enabled = false;
			this.txtNEEMail.Enabled = false;
			this.cmbTipoVivienda.Enabled = false;
			this.dtpTiempoVivienda.Enabled = false;
			this.txtReferencia.Enabled = false;
			this.uGridParentesco.Enabled = false;
			this.uGridParentesco.Enabled = false;
			this.btnTelf.Enabled= false;
			this.btnTelf2.Enabled = false;
			this.btnCelular.Enabled = false;

		}
	
		public void habilita_Control_Conyugue()
		{
			this.txtCDocumento.Enabled= true;
			this.txtCApellidoP.Enabled= true;
			this.txtCApellidoM.Enabled= true;
			this.txtCPNombre.Enabled= true;
			this.txtCSNombre.Enabled= true;
			this.cmbCNacionalidad.Enabled= true;
			this.dpCFechaNacimiento.Enabled= true;
			this.cmbCSexo.Enabled= true;
			this.cmbCProvinciaNacimiento.Enabled= true;
			this.cmbCCantonNacimiento.Enabled= true;
			this.cmbCNivel.Enabled= true;
			this.cmbCProfesion.Enabled= true;
			this.cmbCSituacionLaboral.Enabled= true;
			this.cmbCActividadEconomia.Enabled= true;

		}

		public void deshabilita_Control_Conyugue()
		{

			this.txtCDocumento.Enabled= false;
			this.txtCApellidoP.Enabled=  false;
			this.txtCApellidoM.Enabled=  false;
			this.txtCPNombre.Enabled=  false;
			this.txtCSNombre.Enabled=  false;
			this.cmbCNacionalidad.Enabled=  false;
			this.dpCFechaNacimiento.Enabled=  false;
			this.cmbCSexo.Enabled=  false;
			this.cmbCProvinciaNacimiento.Enabled=  false;
			this.cmbCCantonNacimiento.Enabled=  false;
			this.cmbCNivel.Enabled=  false;
			this.cmbCProfesion.Enabled=  false;
			this.cmbCSituacionLaboral.Enabled=  false;
			this.cmbCActividadEconomia.Enabled=  false;
		}

		public void habilita_Control_Trabajo()
		{

			this.txtSNRuc.Enabled = true;
			this.txtSNNombre.Enabled = true;
			this.dtSNFechaInicio.Enabled = true;
			this.txtSNMetros.Enabled = true;
			this.txtSNEmpleados.Enabled = true;
			this.cmbSCTipoNegocio.Enabled = true;
			this.cmbSubActvidadEconomica.Enabled = true;
			this.cmbSNProvincia.Enabled = true;
			this.cmbSNCanton.Enabled = true;
			this.cmbSNParroquia.Enabled = true;
			this.cmbSNBarrio.Enabled = true;
			this.txtSNCalle1.Enabled = true;
			this.txtSNNumCasa.Enabled = true;
			this.txtSNCalle2.Enabled = true;
			this.txtSNTelefono1.Enabled = true;
			this.txtSNTelefono2.Enabled = true;
			this.txtSNCelular.Enabled = true;
			this.txtSNEmail.Enabled = true;
			this.txtSNReferencia.Enabled = true;
			this.txtSNIngresos.Enabled = true;
			this.txtSNGastos.Enabled = true;
			this.txtSNActivos.Enabled = true;
			this.txtSNPasivos.Enabled = true;
			this.chkSNContabilidad.Enabled = true;
			this.chkSNAfiliacion.Enabled = true;
			this.btnCTelf.Enabled = true;
			this.btnCTelf2.Enabled = true;
			this.btnCCelular.Enabled = true;
		}

		public void deshabilita_Control_Trabajo()
		{
			this.txtSNRuc.Enabled = false;
			this.txtSNNombre.Enabled = false;
			this.dtSNFechaInicio.Enabled = false;
			this.txtSNMetros.Enabled = false;
			this.txtSNEmpleados.Enabled = false;
			this.cmbSCTipoNegocio.Enabled = false;
			this.cmbSubActvidadEconomica.Enabled = false;
			this.cmbSNProvincia.Enabled = false;
			this.cmbSNCanton.Enabled = false;
			this.cmbSNParroquia.Enabled = false;
			this.cmbSNBarrio.Enabled = false;
			this.txtSNCalle1.Enabled = false;
			this.txtSNNumCasa.Enabled = false;
			this.txtSNCalle2.Enabled = false;
			this.txtSNTelefono1.Enabled = false;
			this.txtSNTelefono2.Enabled = false;
			this.txtSNCelular.Enabled = false;
			this.txtSNEmail.Enabled = false;
			this.txtSNReferencia.Enabled = false;
			this.txtSNIngresos.Enabled = false;
			this.txtSNGastos.Enabled = false;
			this.txtSNActivos.Enabled = false;
			this.txtSNPasivos.Enabled = false;
			this.chkSNContabilidad.Enabled = false;
			this.chkSNAfiliacion.Enabled = false;
			this.btnCTelf.Enabled = false;
			this.btnCTelf2.Enabled = false;
			this.btnCCelular.Enabled = false;

		}

		public void habilita_Control_Empresa()
		{
			this.txtSDRuc.Enabled =true;
			this.txtSDEmpresa.Enabled =true;
			this.cmbSDTipoEmpresa.Enabled =true;
			this.dtSDFecha.Enabled =true;
			this.txtSDIngresos.Enabled =true;
			this.txtSDGastos.Enabled =true;
			this.txtSDActivos.Enabled =true;
			this.txtSDPasivos.Enabled =true;
			this.cmbSDTipoContrato.Enabled =true;
			this.cmbSDTipoSueldo.Enabled =true;
			this.txtSDDepartamento.Enabled =true;
			this.cmbSDCargo.Enabled =true;
			this.txtSDPago.Enabled =true;
			this.chkSDAfiliacion.Enabled =true;
			this.cmbSDProvincia.Enabled =true;
			this.cmbSDCanton.Enabled =true;
			this.cmbSDParroquia.Enabled =true;
			this.cmbSDBarrio.Enabled =true;
			this.txtSDCalle1.Enabled =true;
			this.txtSDCalle2.Enabled =true;
			this.txtSDNumCasa.Enabled =true;
			this.txtSDTelefono.Enabled =true;
			this.txtSDExt.Enabled =true;
			this.txtSDCelular.Enabled =true;			
			this.txtSDEmail.Enabled =true;
			this.txtSDReferencia.Enabled = true;
			this.bntSDTelf.Enabled = true;
			this.btnSDCelular.Enabled = true;
		}

		public void deshabilita_Control_Empresa()
		{
			this.txtSDRuc.Enabled =false;
			this.txtSDEmpresa.Enabled =false;
			this.cmbSDTipoEmpresa.Enabled =false;
			this.dtSDFecha.Enabled =false;
			this.txtSDIngresos.Enabled =false;
			this.txtSDGastos.Enabled =false;
			this.txtSDActivos.Enabled =false;
			this.txtSDPasivos.Enabled =false;
			this.cmbSDTipoContrato.Enabled =false;
			this.cmbSDTipoSueldo.Enabled =false;
			this.txtSDDepartamento.Enabled =false;
			this.cmbSDCargo.Enabled =false;
			this.txtSDPago.Enabled =false;
			this.chkSDAfiliacion.Enabled =false;
			this.cmbSDProvincia.Enabled =false;
			this.cmbSDCanton.Enabled =false;
			this.cmbSDParroquia.Enabled =false;
			this.cmbSDBarrio.Enabled =false;
			this.txtSDCalle1.Enabled =false;
			this.txtSDCalle2.Enabled =false;
			this.txtSDNumCasa.Enabled =false;
			this.txtSDTelefono.Enabled =false;
			this.txtSDExt.Enabled =false;
			this.txtSDCelular.Enabled =false;		
			this.txtSDEmail.Enabled =false;
			this.txtSDReferencia.Enabled = false;
			this.bntSDTelf.Enabled = false;
			this.btnSDCelular.Enabled = false;
		}


		public void LimpiaControles()
		{
			#region Garante			
			this.txtClienteIdCre_DatosGenerales.Value = 0;
			this.txtIdCre_Garante.Value = 0;
			this.txtClienteIdCre_DatosUbicacion.Value = 0;
			this.txtIdCre_Conyugue.Value = 0;
			this.txtConyugueIdCre_DatosGenerales.Value = 0;
			this.txtClienteIdCre_Negocio.Value = 0;
			this.txtNegocioIdCre_DatosUbicacion.Value = 0;
			this.txtClienteIdCre_EmpresasTrabajo.Value = 0;
			this.txtEmpresaIdCre_DatosUbicacion.Value = 0;
			this.txtDocumento.Text = "";
			this.txtApellidoP.Text= "";
			this.txtApellidoM.Text= "";
			this.txtPnombre.Text= "";
			this.txtSnombre.Text= "";
			this.cmbEstado.Value = 0;
			this.cmbCalificacion.Value = 0;
			this.txtScore.Value = 0;
			this.cmbNENacionalidad.Value = 54;
			this.dtpNEFechaNacimiento.Value = System.DBNull.Value;
			this.cmbSexo.Value = 0;
			this.cmbProvinciaNacimiento.Value = 0;
			this.cmbCantonNacimiento.Value = System.DBNull.Value;
			this.cmbNEEstadoCivil.Value = 0;
			this.txtNECargasFamiliares.Value = 0;
			this.cmbNivel.Value = 0;
			this.cmbProfesion.Value = System.DBNull.Value;
			this.cmbSituacionLaboral.Value = 0;
			this.cmbActividadEcomica.Value = 0;
			this.cmbProvinciaDomicilio.Value = 0;
			this.cmbCantonDomicilio.Value = 0;
			this.cmbParroquiaDomicilio.Value = 0;
			this.cmbBarrioDomicilio.Value = 0;
			this.txtCalle1.Text= "";
			this.txtNumcasa.Text= "";
			this.txtCalle2.Text= "";
			this.txtNETelefono1.Text= "";
			this.txtNETelefono2.Text= "";
			this.txtNECelular.Text= "";
			this.txtNEEMail.Text= "";
			this.cmbTipoVivienda.Value = 0;
			this.dtpTiempoVivienda.Value = System.DBNull.Value;
			this.txtReferencia.Text= "";
			this.btnTelf.Enabled = false;
			this.btnTelf2.Enabled = false;
			this.btnCelular.Enabled = false;
			#endregion;

			#region Conyuge
			this.txtCDocumento.Text= "";
			this.txtCApellidoP.Text= "";
			this.txtCApellidoM.Text= "";
			this.txtCPNombre.Text= "";
			this.txtCSNombre.Text= "";
			this.cmbCNacionalidad.Value = 54;
			this.dpCFechaNacimiento.Value = System.DBNull.Value;
			this.cmbCSexo.Value = 0;
			this.cmbCProvinciaNacimiento.Value = 0;
			this.cmbCCantonNacimiento.Value = 0;
			this.cmbCNivel.Value = 0;
			this.cmbCProfesion.Value = 0;
			this.cmbCSituacionLaboral.Value = 0;
			this.cmbCActividadEconomia.Value = 0;
			#endregion 

			#region Referencias

			this.uGridParentesco.DataSource = Funcion.dvProcedimiento(cdsSeteoF, string.Format("Exec Cre_ConsultaReferencias 0,1"));

			#endregion Referencias

			#region Negocio
			this.txtSNRuc.Text = "";
			this.txtSNNombre.Text = "";
			this.dtSNFechaInicio.Value = System.DBNull.Value;
			this.txtSNMetros.Value = 1;
			this.txtSNEmpleados.Value = 0;
			this.cmbSCTipoNegocio.Value = 0;
			this.cmbSubActvidadEconomica.Value = 0;
			this.cmbSNProvincia.Value = 0;
			this.cmbSNCanton.Value = 0;
			this.cmbSNParroquia.Value = 0;
			this.cmbSNBarrio.Value = 0;
			this.txtSNEmail.Text = "";
			this.txtSNCalle1.Text ="";
			this.txtSNNumCasa.Text="";
			this.txtSNCalle2.Text = "";
			this.txtSNTelefono1.Text = "";
			this.txtSNTelefono2.Text = "";
			this.txtSNCelular.Text = "";
			this.txtSNEmail.Text= "";
			this.txtSNReferencia.Text = "";
			this.txtSNIngresos.Value = 0;
			this.txtSNGastos.Value = 0;
			this.txtSNActivos.Value = 0;
			this.txtSNPasivos.Value = 0;
			this.chkSNContabilidad.Checked = false;
			this.chkSNAfiliacion.Checked = false;
			#endregion Negocio

			#region Empresa
			this.txtSDRuc.Text ="";
			this.txtSDEmpresa.Text ="";
			this.cmbSDTipoEmpresa.Value = 0;
			this.dtSDFecha.Value = System.DBNull.Value; 
			this.txtSDIngresos.Value =0;
			this.txtSDGastos.Value =0;
			this.txtSDActivos.Value =0;
			this.txtSDPasivos.Value=0;
			this.cmbSDTipoContrato.Value =0;
			this.cmbSDTipoSueldo.Value = 0;
			this.txtSDDepartamento.Text ="";
			this.cmbSDCargo.Value = 0;
			this.txtSDPago.Value = 0;
			this.chkSDAfiliacion.Checked = false;
			this.cmbSDProvincia.Value = 0;
			this.cmbSDCanton.Value = 0;
			this.cmbSDParroquia.Value = 0;
			this.cmbSDBarrio.Value = 0;
			this.txtSDEmail.Text ="";
			this.txtSDCalle1.Text = "";
			this.txtSDNumCasa.Text = "";
			this.txtSDCalle2.Text = "";
			this.txtSDTelefono.Text= "";
			this.txtSDExt.Text = "";
			this.txtSDCelular.Text= "";
			this.txtSDReferencia.Text = "";
			this.bntSDTelf.Enabled = false;
			this.btnSDCelular.Enabled = false;
			#endregion Empresa

		}

		private void cmbNEEstadoCivil_ValueChanged(object sender, System.EventArgs e)
		{
			if( (int) this.cmbNEEstadoCivil.Value != 2 && (int) this.cmbNEEstadoCivil.Value != 3 && (int) this.cmbNEEstadoCivil.Value != 5)
			{
				this.cmbCNacionalidad.Value = 54;
				this.cmbCTipoRuc.Value = 1;
				habilita_Control_Conyugue();
			}
			else 
			{
				this.cmbCNacionalidad.Value = 54;
				this.cmbCTipoRuc.Value = 1;
				deshabilita_Control_Conyugue();
			}
		}

		private void cmbSituacionLaboral_ValueChanged(object sender, System.EventArgs e)
		{
			if (ientidad > 0)
			{
				SituacionLaboral(ientidad);
			}
		}

		private void SituacionLaboral(int idEntidadFinanciera)
		{
			if (idEntidadFinanciera != 2)
			{
				if (this.cmbSituacionLaboral.Value != null)
				{
					if ((int)this.cmbSituacionLaboral.Value == 1)
					{
						deshabilita_Control_Trabajo();
						habilita_Control_Empresa();
					}
					else if ((int)this.cmbSituacionLaboral.Value == 3)
					{
						habilita_Control_Empresa();
						habilita_Control_Trabajo();
					}
					else
					{
						habilita_Control_Trabajo();
						deshabilita_Control_Empresa();
					}
					this.cmbActividadEcomica.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "Exec Listas 27,1, 0");
					this.cmbActividadEcomica.Value = 0;
				}
			}
			else
			{
				if (this.cmbSituacionLaboral.Value != null)
				{
					if ((int)this.cmbSituacionLaboral.Value == 10)
					{
							deshabilita_Control_Trabajo();
							habilita_Control_Empresa();
					}
					else
					{
						habilita_Control_Trabajo();
						deshabilita_Control_Empresa();
					}

					int iSituacion = 0;
					if ((int) this.cmbSituacionLaboral.Value == 10) iSituacion = 1;
					if ((int) this.cmbSituacionLaboral.Value == 11) iSituacion = 2;
					if ((int) this.cmbSituacionLaboral.Value == 12) iSituacion = 3;
					this.cmbActividadEcomica.DataSource = Funcion.dvProcedimiento(cdsSeteoF, string.Format("Exec Listas 27, 2, {0}", iSituacion));
					this.cmbActividadEcomica.Value = 0;

				}
			}
		}

		#region SituacionLaboralConyugue

		private void SituacionLaboral1(int idEntidadFinanciera)
		{
			if (idEntidadFinanciera != 2)
			{
				if (this.cmbCSituacionLaboral.Value != null)
				{
					if ((int)this.cmbCSituacionLaboral.Value == 1)
					{
						
					}
					else if ((int)this.cmbCSituacionLaboral.Value == 3)
					{
					
					}
					else
					{

					}
					this.cmbCActividadEconomia.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "Exec Listas 27,1, 0");
					this.cmbCActividadEconomia.Value = 0;
				}
			}
			else
			{
				if (this.cmbCSituacionLaboral.Value != null)
				{
					if ((int)this.cmbCSituacionLaboral.Value == 10)
					{
				
					}
					else
					{

					}

					int iSituacion = 0;
					if ((int) this.cmbCSituacionLaboral.Value == 10) iSituacion = 1;
					if ((int) this.cmbCSituacionLaboral.Value == 11) iSituacion = 2;
					if ((int) this.cmbCSituacionLaboral.Value == 12) iSituacion = 3;
					this.cmbCActividadEconomia.DataSource = Funcion.dvProcedimiento(cdsSeteoF, string.Format("Exec Listas 27, 2, {0}", iSituacion));
					this.cmbCActividadEconomia.Value = 0;

				}
			}

		}

		#endregion SituacionLaboralConyugue

		private void txtDocumento_Validating(object sender, System.ComponentModel.CancelEventArgs e)
		{
			
			if (this.cmbTipoRuc.ActiveRow != null)
			{
				int iTipoRuc = 0;
				if ((int) this.cmbTipoRuc.Value == 1) iTipoRuc = 5;// cedula
				if ((int) this.cmbTipoRuc.Value == 2) iTipoRuc = 6;// pasaporte
				if ((int) this.cmbTipoRuc.Value == 3) iTipoRuc = 4;// ruc
				if (!Validacion.vbIdentificacion(this.txtDocumento, iTipoRuc, cdsSeteoF)) 
				{
					e.Cancel = true;    }
			}
		}

		// REGION KEY PRESS // DOWN

		#region KEY PRESS

		private void txtDocumento_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if(e.KeyCode == Keys.Enter) this.txtApellidoP.Focus();
		}

		private void txtDocumento_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
			validar.Solonumeros(e);
		}

		private void txtApellidoP_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if(e.KeyCode == Keys.Enter) this.txtApellidoM.Focus();
		}

		private void txtApellidoM_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if(e.KeyCode == Keys.Enter) this.txtPnombre.Focus();	
		}

		private void txtPnombre_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if(e.KeyCode == Keys.Enter) this.txtSnombre.Focus();	
		}

		private void txtSnombre_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if(e.KeyCode == Keys.Enter) this.cmbEstado.Focus();	
		}

		private void cmbEstado_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if(e.KeyCode == Keys.Enter) this.cmbCalificacion.Focus();
		}

		private void cmbCalificacion_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if(e.KeyCode == Keys.Enter) this.txtScore.Focus();
		}

		private void txtScore_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			this.utcSolicitud.Tabs[0].Selected = true;
			if(e.KeyCode == Keys.Enter) this.cmbNENacionalidad.Focus();
	
		}

		private void cmbNENacionalidad_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if(e.KeyCode == Keys.Enter) this.dtpNEFechaNacimiento.Focus();
		}

		private void dtpNEFechaNacimiento_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if(e.KeyCode == Keys.Enter) this.cmbSexo.Focus();
		}

		private void cmbSexo_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if(e.KeyCode == Keys.Enter) this.cmbProvinciaNacimiento.Focus();
		}

		private void cmbProvinciaNacimiento_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if(e.KeyCode == Keys.Enter) this.cmbCantonNacimiento.Focus();
		}

		private void cmbProvinciaNacimiento_ValueChanged(object sender, System.EventArgs e)
		{
			if (this.cmbProvinciaNacimiento.ActiveRow != null)
			{
				this.cmbCantonNacimiento.DataSource = Funcion.dvProcedimiento(cdsSeteoF, string.Format("Exec Listas 1, {0}, 0", (int) this.cmbProvinciaNacimiento.Value));
			}
		}

		private void cmbCantonNacimiento_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if(e.KeyCode == Keys.Enter) this.cmbNEEstadoCivil.Focus();
		}

		private void cmbNEEstadoCivil_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if(e.KeyCode == Keys.Enter) this.txtNECargasFamiliares.Focus();
		}

		private void txtNECargasFamiliares_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if(e.KeyCode == Keys.Enter) this.cmbNivel.Focus();
		}

		private void cmbNivel_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if(e.KeyCode == Keys.Enter) this.cmbProfesion.Focus();
		}

		private void cmbProfesion_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if(e.KeyCode == Keys.Enter)
			{
				if (ientidad == 2)
				{
					this.cmbSituacionLaboral.Focus();

				}
				else
				{
					this.cmbSituacionLaboral.Focus(); 
				
				}
			}
		}

		private void cmbSituacionLaboral_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if(e.KeyCode == Keys.Enter) this.cmbActividadEcomica.Focus();
		}

		private void cmbActividadEcomica_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if(e.KeyCode == Keys.Enter) this.cmbProvinciaDomicilio.Focus();
		}

		private void cmbProvinciaDomicilio_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if(e.KeyCode == Keys.Enter) this.cmbCantonDomicilio.Focus();
		}

		private void cmbProvinciaDomicilio_ValueChanged(object sender, System.EventArgs e)
		{
			if (this.cmbProvinciaDomicilio.ActiveRow != null)
			{
				this.cmbCantonDomicilio.DataSource = Funcion.dvProcedimiento(cdsSeteoF, string.Format("Exec Listas 1, {0}, 0", (int) this.cmbProvinciaDomicilio.Value));
			}
		}

		private void cmbCantonDomicilio_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if(e.KeyCode == Keys.Enter) this.cmbParroquiaDomicilio.Focus();
		}

		private void cmbCantonDomicilio_ValueChanged(object sender, System.EventArgs e)
		{
			if (this.cmbCantonDomicilio.ActiveRow != null)
			{
				this.cmbParroquiaDomicilio.DataSource = Funcion.dvProcedimiento(cdsSeteoF, string.Format("Exec Listas 2, {0}, 0", (int) this.cmbCantonDomicilio.Value));
			}
		}

		private void cmbParroquiaDomicilio_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if(e.KeyCode == Keys.Enter) this.cmbBarrioDomicilio.Focus();
		}

		private void cmbParroquiaDomicilio_ValueChanged(object sender, System.EventArgs e)
		{
			if (this.cmbParroquiaDomicilio.ActiveRow != null)
			{
				this.cmbBarrioDomicilio.DataSource = Funcion.dvProcedimiento(cdsSeteoF, string.Format("Exec Listas 3, {0}, 0", (int) this.cmbParroquiaDomicilio.Value));
			}
		}

		private void cmbBarrioDomicilio_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if(e.KeyCode == Keys.Enter) this.txtCalle1.Focus();
		}

		private void txtCalle1_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if(e.KeyCode == Keys.Enter) this.txtNumcasa.Focus();
		
		}

		private void txtNumcasa_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if(e.KeyCode == Keys.Enter) this.txtCalle2.Focus();
		}

		private void txtCalle2_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if(e.KeyCode == Keys.Enter) this.txtNETelefono1.Focus();
		}

		private void txtNETelefono1_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if(e.KeyCode == Keys.Enter) this.txtNETelefono2.Focus();
		}

		private void txtNETelefono2_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if(e.KeyCode == Keys.Enter) this.txtNECelular.Focus();
		}

		private void txtNECelular_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if(e.KeyCode == Keys.Enter) this.txtNEEMail.Focus();
		}

		private void txtNEEMail_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if(e.KeyCode == Keys.Enter) this.cmbTipoVivienda.Focus();
		}

		private void cmbTipoVivienda_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if(e.KeyCode == Keys.Enter) this.dtpTiempoVivienda.Focus();
		}

		private void dtpTiempoVivienda_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if(e.KeyCode == Keys.Enter) this.txtReferencia.Focus();
		}

		private void txtReferencia_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			try
			{
				if(e.KeyCode == Keys.Enter) 
					if ((int)this.cmbNEEstadoCivil.Value != 2 && (int)this.cmbNEEstadoCivil.Value != 3 && (int)this.cmbNEEstadoCivil.Value != 5)
					{
						this.utcSolicitud.Tabs[1].Selected = true;
						habilita_Control_Conyugue();
						this.txtCDocumento.Focus();
					}
					else 
					{
						this.utcSolicitud.Tabs[2].Selected = true;
						this.uGridParentesco.Focus();
						deshabilita_Control_Conyugue();
					}
			}
			catch(Exception){
				MessageBox.Show(string.Format("Seleccione Género y una Situacion / Fuente de Ingreso"), "Point Technology", MessageBoxButtons.OK,MessageBoxIcon.Stop);
				this.cmbNEEstadoCivil.Focus();
				this.utcSolicitud.Tabs[0].Selected = true;
				return;	
			}
			
		}

		private void txtCDocumento_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if(e.KeyCode == Keys.Enter) this.txtCApellidoP.Focus();
		}

		private void txtCApellidoP_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if(e.KeyCode == Keys.Enter) this.txtCApellidoM.Focus();
		}

		private void txtCApellidoM_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if(e.KeyCode == Keys.Enter) this.txtCPNombre.Focus();
		}

		private void txtCPNombre_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if(e.KeyCode == Keys.Enter) this.txtCSNombre.Focus();
		}

		private void txtCSNombre_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if(e.KeyCode == Keys.Enter) this.cmbCNacionalidad.Focus();
		}

		private void cmbCNacionalidad_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if(e.KeyCode == Keys.Enter) this.dpCFechaNacimiento.Focus();
		}

		private void dpCFechaNacimiento_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if(e.KeyCode == Keys.Enter) this.cmbCSexo.Focus();
		}

		private void cmbCSexo_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if(e.KeyCode == Keys.Enter) this.cmbCProvinciaNacimiento.Focus(); 
		}

		private void cmbCCantonNacimiento_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if(e.KeyCode == Keys.Enter) this.cmbCNivel.Focus(); 
		}

		private void cmbCProvinciaNacimiento_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if(e.KeyCode == Keys.Enter) this.cmbCCantonNacimiento.Focus(); 
		}

		private void cmbCNivel_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if(e.KeyCode == Keys.Enter) this.cmbCProfesion.Focus(); 
		}

		private void cmbCProfesion_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			
		}

		private void cmbCProfesion_KeyDown_1(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if(e.KeyCode == Keys.Enter) this.cmbCSituacionLaboral.Focus(); 
		}


		private void cmbCSituacionLaboral_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if(e.KeyCode == Keys.Enter) this.cmbCActividadEconomia.Focus(); 
		}

		private void cmbCActividadEconomia_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
		
		}

		
		private void txtSNRuc_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if(e.KeyCode == Keys.Enter) this.txtSNNombre.Focus();
		}

		private void txtSNNombre_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if(e.KeyCode == Keys.Enter) this.dtSNFechaInicio.Focus();
		}

		private void dtSNFechaInicio_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if(e.KeyCode == Keys.Enter) this.txtSNMetros.Focus();
		}

		private void txtSNMetros_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if(e.KeyCode == Keys.Enter) this.txtSNEmpleados.Focus();
		}

		private void txtSNEmpleados_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if(e.KeyCode == Keys.Enter) this.cmbSCTipoNegocio.Focus();
		}

		private void cmbSCTipoNegocio_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if(e.KeyCode == Keys.Enter) this.cmbSubActvidadEconomica.Focus();
		}

		private void cmbSubActvidadEconomica_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if(e.KeyCode == Keys.Enter) this.cmbSNProvincia.Focus();
		}

		private void cmbSNProvincia_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if(e.KeyCode == Keys.Enter) this.cmbSNCanton.Focus();
		}

		private void cmbSNProvincia_ValueChanged(object sender, System.EventArgs e)
		{
			if (this.cmbSNProvincia.ActiveRow != null)
			{
				this.cmbSNCanton.DataSource = Funcion.dvProcedimiento(cdsSeteoF, string.Format("Exec Listas 1, {0}, 0", (int) this.cmbSNProvincia.Value));
			}
		}

		private void cmbSNCanton_ValueChanged(object sender, System.EventArgs e)
		{
			if (this.cmbSNCanton.ActiveRow != null)
			{
				this.cmbSNParroquia.DataSource = Funcion.dvProcedimiento(cdsSeteoF, string.Format("Exec Listas 2, {0}, 0", (int) this.cmbSNCanton.Value));
			}
		}

		private void cmbSNParroquia_ValueChanged(object sender, System.EventArgs e)
		{
			if (this.cmbSNParroquia.ActiveRow != null)
			{
				this.cmbSNBarrio.DataSource = Funcion.dvProcedimiento(cdsSeteoF, string.Format("Exec Listas 3, {0}, 0", (int) this.cmbSNParroquia.Value));
			}
		}

		private void cmbSNCanton_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if(e.KeyCode == Keys.Enter) this.cmbSNParroquia.Focus();
		}



		private void cmbSNParroquia_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if(e.KeyCode == Keys.Enter) this.cmbSNBarrio.Focus();
		}

		private void cmbSNBarrio_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if(e.KeyCode == Keys.Enter) this.txtSNCalle1.Focus();
		}

		private void txtSNCalle1_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if(e.KeyCode == Keys.Enter) this.txtSNNumCasa.Focus();
		}

		private void txtSNNumCasa_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if(e.KeyCode == Keys.Enter) this.txtSNCalle2.Focus();
		}

		private void txtSNCalle2_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if(e.KeyCode == Keys.Enter) this.txtSNTelefono1.Focus();
		}

		private void txtSNTelefono1_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if(e.KeyCode == Keys.Enter) this.txtSNTelefono2.Focus();
		}

		private void txtSNTelefono2_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if(e.KeyCode == Keys.Enter) this.txtSNCelular.Focus();
		}

		private void txtSNCelular_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if(e.KeyCode == Keys.Enter) this.txtSNEmail.Focus();
		}

		private void txtSNEmail_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if(e.KeyCode == Keys.Enter) this.txtSNReferencia.Focus();
		}

		private void txtSNReferencia_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if(e.KeyCode == Keys.Enter) this.txtSNIngresos.Focus();
		}

		private void txtSNIngresos_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if(e.KeyCode == Keys.Enter) this.txtSNGastos.Focus();
		}

		private void txtSNGastos_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if(e.KeyCode == Keys.Enter) this.txtSNActivos.Focus();
		}

		private void txtSNActivos_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if(e.KeyCode == Keys.Enter) this.txtSNPasivos.Focus();
		}

		private void txtSNPasivos_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if(e.KeyCode == Keys.Enter) this.txtSNIngresos.Focus();
		}

		private void chkSNContabilidad_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if(e.KeyCode == Keys.Enter) this.chkSNAfiliacion.Focus();
		}

		private void chkSNAfiliacion_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
		
		}

		private void txtSDRuc_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if(e.KeyCode == Keys.Enter) this.txtSDEmpresa.Focus();
		}

		private void txtSDEmpresa_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if(e.KeyCode == Keys.Enter) this.cmbSDTipoEmpresa.Focus();
		}

		private void cmbSDTipoEmpresa_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if(e.KeyCode == Keys.Enter) this.dtSDFecha.Focus();
		}

		private void dtSDFecha_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if(e.KeyCode == Keys.Enter) this.txtSDIngresos.Focus();
		}

		private void txtSDIngresos_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if(e.KeyCode == Keys.Enter) this.txtSDGastos.Focus();
		}

		private void txtSDGastos_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if(e.KeyCode == Keys.Enter) this.txtSDActivos.Focus();
		}

		private void txtSDActivos_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if(e.KeyCode == Keys.Enter) this.txtSDPasivos.Focus();
		}

		private void txtSDPasivos_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if(e.KeyCode == Keys.Enter) this.cmbSDTipoContrato.Focus();
		}

		private void cmbSDTipoContrato_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if(e.KeyCode == Keys.Enter) this.cmbSDTipoSueldo.Focus();
		}

		private void cmbSDTipoSueldo_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if(e.KeyCode == Keys.Enter) this.txtSDDepartamento.Focus();
		}

		private void txtSDDepartamento_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if(e.KeyCode == Keys.Enter) this.cmbSDCargo.Focus();
		}

		private void cmbSDCargo_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if(e.KeyCode == Keys.Enter) this.txtSDPago.Focus();
		}

		private void txtSDPago_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if(e.KeyCode == Keys.Enter) this.chkSDAfiliacion.Focus();
		}

		private void chkSDAfiliacion_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if(e.KeyCode == Keys.Enter) this.cmbSDProvincia.Focus();
		}

		private void cmbSDProvincia_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if(e.KeyCode == Keys.Enter) this.cmbSDCanton.Focus();
		}

		private void cmbSDCanton_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if(e.KeyCode == Keys.Enter) this.cmbSDParroquia.Focus();
		}

		private void cmbSDParroquia_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if(e.KeyCode == Keys.Enter) this.cmbSDBarrio.Focus();
		}

		private void cmbSDBarrio_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if(e.KeyCode == Keys.Enter) this.txtSDCalle1.Focus();
		}

		private void txtSDCalle1_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if(e.KeyCode == Keys.Enter) this.txtSDNumCasa.Focus();
		}

		private void txtSDNumCasa_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if(e.KeyCode == Keys.Enter) this.txtSDCalle2.Focus();
		}

		private void txtSDCalle2_Paint(object sender, System.Windows.Forms.PaintEventArgs e)
		{
			
		}

		private void txtSDExt_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if(e.KeyCode == Keys.Enter) this.txtSDCelular.Focus();
		}

		private void txtSDCelular_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if(e.KeyCode == Keys.Enter) this.txtSDEmail.Focus();
		}

		private void txtSDEmail_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if(e.KeyCode == Keys.Enter) this.txtSDReferencia.Focus();
		}

		private void txtSDReferencia_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if(e.KeyCode == Keys.Enter) this.btnGuardar.Focus();
		}

		private void cmbSDProvincia_ValueChanged(object sender, System.EventArgs e)
		{
			if (this.cmbSDProvincia.ActiveRow != null)
			{
				this.cmbSDCanton.DataSource = Funcion.dvProcedimiento(cdsSeteoF, string.Format("Exec Listas 1, {0}, 0", (int) this.cmbSDProvincia.Value));
			}
		}

		private void cmbSDCanton_ValueChanged(object sender, System.EventArgs e)
		{
			if (this.cmbSDCanton.ActiveRow != null)
			{
				this.cmbSDParroquia.DataSource = Funcion.dvProcedimiento(cdsSeteoF, string.Format("Exec Listas 2, {0}, 0", (int) this.cmbSDCanton.Value));
			}
		}

		private void cmbSDParroquia_ValueChanged(object sender, System.EventArgs e)
		{
			if (this.cmbSDParroquia.ActiveRow != null)
			{
				this. cmbSDBarrio.DataSource = Funcion.dvProcedimiento(cdsSeteoF, string.Format("Exec Listas 3, {0}, 0", (int) this.cmbSDParroquia.Value));
			}
		}

		private void txtSDCalle2_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if(e.KeyCode == Keys.Enter) this.txtSDTelefono.Focus();
		}

		private void txtSDTelefono_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if(e.KeyCode == Keys.Enter) this.txtSDExt.Focus();
		}


		#endregion KEY PRESS

		private void btnTelf_Click(object sender, System.EventArgs e)
		{
			if (this.txtNETelefono1.Text.ToString().Trim().Length == 0)
			{
				MessageBox.Show(string.Format("Debe Ingresar un Número de Teléfono "), "Point Technology", MessageBoxButtons.OK,MessageBoxIcon.Stop);
				this.txtNETelefono1.Focus();
				return;	
			}

				if(!Validacion.vbValidaTelefono(this.txtNETelefono1, 1, false)) return;

				int iNuevoCel = 0;
				int iCelExistente = 0;
			
			iNuevoCel = Funcion.iEscalarSQL(cdsSeteoF, string.Format("select count(*) from Cre_DatosUbicacion where Origen = 2 and idOrigen = {0}and idCre_Ubicacion = {2} and Telefono1 like '%"+"{1}"+"%' ", (int)this.txtIdCre_Garante.Value, this.txtNETelefono1.Text.ToString(), (int)this.txtClienteIdCre_DatosUbicacion.Value));

				if((int)this.txtIdCre_Solicitud.Value > 0)
				{
					if (iNuevoCel > 0)
					{
						AbreVerificaciontelefonica (false, (int)this.txtIdCre_Garante.Value, 0,
							this.txtApellidoP.Text + " " + this.txtApellidoM.Text + " " + this.txtPnombre.Text + " " + this.txtSnombre.Text,
							this.txtNETelefono1.Text, 0); 
					}
					else
					{
						MessageBox.Show(string.Format("Guarde los datos del Garante para realizar la verificacion telefonica"), "Point Technology", MessageBoxButtons.OK,MessageBoxIcon.Stop);
						this.txtNETelefono1.Focus();
						return;	
					}

				}
				else
				{
					MessageBox.Show(string.Format("Guarde los datos del Garante para realizar la verificacion telefonica"), "Point Technology", MessageBoxButtons.OK,MessageBoxIcon.Stop);
					this.txtNETelefono1.Focus();
					return;	
			}
		}

		private void AbreVerificaciontelefonica (bool bClienteGarante, int idCre_Solicitud, int iOrigen, string sNombre, string sTelefono, int Indice)
		{
			using (frmCre_VerificacionTelefonica  consultaCR = new frmCre_VerificacionTelefonica(bClienteGarante, idCre_Solicitud, iOrigen,
							 sNombre,sTelefono, 0))
			{
				if (DialogResult.OK == consultaCR.ShowDialog())
				{

				}
			}
		}


		private void btnSalir_Click(object sender, System.EventArgs e)
		{
			this.DialogResult = DialogResult.No;
		}

		private void txtDocumento_Validated(object sender, System.EventArgs e)
		{
			 bCarga = true;
			if (sCreSolicitud == this.txtDocumento.Text)
			{
				MessageBox.Show(string.Format("No puede ser el Cliente Garante de si mismo"), "Point Technology", MessageBoxButtons.OK,MessageBoxIcon.Stop);
				this.txtDocumento.Focus();
				return;
			}
			int iContG = 0;
		
			iContG = Funcion.iEscalarSQL(cdsSeteoF, string.Format("select count(*) from Cre_Garantes c  Inner Join Cre_DatosGenerales d on c.idCre_Garante = d.idOrigen  where idCre_Solicitud = {0} and NumeroIdentificacion like '%"+"{1}"+"%'", (int)this.txtIdCre_Solicitud.Value, this.txtDocumento.Text.ToString()));
			if( iContG > 0 && (int) this.txtClienteIdCre_DatosGenerales.Value == 0)
			{
				MessageBox.Show("El Cliente ya tiene registrado a un garante con la siguiente Cedula " + this.txtDocumento.Text + ", ", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				this.txtDocumento.Text ="";
				this.txtDocumento.Focus ();
				return;	
			}
//			bool bRes = Funcion.bEscalarSQL(cdsSeteoF, string.Format("Exec ValidaDatosCognoScoring '{0}'",this.txtDocumento.Text), true);
//			if (!bRes)
//			{
//				MessageBox.Show(string.Format("Para registrar un Garante primero debe realizar la consulta en  CognoWare "), "Point Technology", MessageBoxButtons.OK,MessageBoxIcon.Stop);
//				this.txtDocumento.Focus();
//				return;
//			}

			if (idCre_Solicitud > 0 && bnuevo)
			{
				string sSQLDatosNovaScoring = string.Format("Exec ActualizaDatosNovaScoring '{0}'",this.txtDocumento.Text);
				Funcion.EjecutaSQL(cdsSeteoF, sSQLDatosNovaScoring);
				MessageBox.Show("Datos de Sincronizados con Cogno Ware", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Information);		

				string ssqlDo= string.Format ("Exec Cre_ConsultaGarantes {0}", idCre_Solicitud ); 
				SqlDataReader drDO = Funcion.rEscalarSQL(cdsSeteoF,ssqlDo,true);
				drDO.Read();
				if(drDO.HasRows)
				{
					this.txtIdCre_Garante.Value = drDO.GetInt32 (0);
					this.txtIdCre_Solicitud.Value = drDO.GetInt32 (1);
				}
				drDO.Close();
					this.txtIdCre_Garante.Value = isolicitud;

				Consulta((int)this.txtIdCre_Garante.Value, (int) this.txtIdCre_Solicitud.Value);
			}
			bCarga = false;
		}


		private void Consulta (int idGarante, int iCreSolicitud)
		{
			int iEstado = 0;
			int iCre_DatosGenerales = 0;
      this.txtIdCre_Solicitud.Value = idCre_Solicitud;
			if(idGarante > 0){
				this.txtIdCre_Garante.Value = idGarante;
				
				#region Carga Solicitudes

				string ssql= string.Format ("Exec Cre_ConsultaSolicitudes {0}",iCreSolicitud ); 
				SqlDataReader dr = Funcion.rEscalarSQL(cdsSeteoF,ssql,true);
				dr.Read();
				if (dr.HasRows)
				{		
					this.txtIdCre_Solicitud.Value = dr.GetInt32 (0);
				}
				dr.Close();

				#endregion Carga Solicitudes
			

				#region ValidaEntidad

				if (ientidad == 2)
				{
					this.label56.Text = "Fuente de Ingresos";
				}
				else
				{
					this.label56.Text = "Situación Laboral";
				}

				#endregion ValidaEntidad

				#region Datos Generales

				string ssqDG= string.Format ("Exec Cre_ConsultaDatosGenerales {0},2",idGarante); 
				SqlDataReader drDG = Funcion.rEscalarSQL(cdsSeteoF,ssqDG,true);
				drDG.Read();
				if (drDG.HasRows)
				{		
					this.txtClienteIdCre_DatosGenerales.Value = drDG.GetInt32 (0);
					if (drDG.GetValue (2) != System.DBNull.Value)this.txtClienteIdCre_DatosGenerales.Value = drDG.GetInt32 (0);
					if (drDG.GetValue (3) != System.DBNull.Value)this.cmbTipoRuc.Value = drDG.GetInt32 (3);
					if (drDG.GetValue (4) != System.DBNull.Value)this.txtDocumento.Text = drDG.GetString (4);
					if (drDG.GetValue (5)!= System.DBNull.Value)this.txtApellidoP.Text = drDG.GetString (5);
					if (drDG.GetValue (6)!= System.DBNull.Value)this.txtApellidoM.Text = drDG.GetString (6);
					if (drDG.GetValue (7)!= System.DBNull.Value)this.txtPnombre.Text = drDG.GetString (7);
					if (drDG.GetValue (8) != System.DBNull.Value)this.txtSnombre.Text = drDG.GetString (8);
					if (drDG.GetValue (9) != System.DBNull.Value)this.cmbCalificacion.Value = drDG.GetInt32 (9);
					if (drDG.GetValue (13) != System.DBNull.Value)this.cmbEstado.Value = drDG.GetInt32 (13);
					if (drDG.GetValue (11) != System.DBNull.Value)this.txtScore.Value = drDG.GetInt32 (11);
					if (drDG.GetValue (14) != System.DBNull.Value)this.cmbNENacionalidad.Value = drDG.GetInt32 (14);
					if (drDG.GetValue (15) != System.DBNull.Value)this.dtpNEFechaNacimiento.Value = drDG.GetDateTime (15);
					if (drDG.GetValue (16) != System.DBNull.Value)this.cmbSexo.Value = drDG.GetInt32 (16);
					if (drDG.GetValue (17) != System.DBNull.Value)this.cmbProvinciaNacimiento.Value = drDG.GetInt32 (17);
					if (drDG.GetValue (18) != System.DBNull.Value)this.cmbCantonNacimiento.Value = drDG.GetInt32 (18);
					if (drDG.GetValue (19) != System.DBNull.Value)this.cmbNEEstadoCivil.Value = drDG.GetInt32 (19);
					if (drDG.GetValue (20) != System.DBNull.Value)this.txtNECargasFamiliares.Value = drDG.GetInt32 (20);
					if (drDG.GetValue (21) != System.DBNull.Value)this.cmbNivel.Value = drDG.GetInt32 (21);
					if (drDG.GetValue (22) != System.DBNull.Value)this.cmbProfesion.Value = drDG.GetInt32 (22);
					if (drDG.GetValue (23) != System.DBNull.Value)this.cmbSituacionLaboral.Value = drDG.GetInt32 (23);
					if (drDG.GetValue (24) != System.DBNull.Value)this.cmbSituacionLaboral.Value = drDG.GetInt32 (24);
					SituacionLaboral(ientidad);
					if (drDG.GetValue (25) != System.DBNull.Value)this.cmbActividadEcomica.Value = drDG.GetInt32 (25);
				}
				drDG.Close();
				#endregion Datos Generales    
  
				#region Datos Ubicacion

				string sqlCDU= string.Format ("Exec Cre_ConsultaDatosUbicacion {0},2",idGarante ); 
				SqlDataReader drCub = Funcion.rEscalarSQL(cdsSeteoF,sqlCDU,true);
				drCub.Read();
				if (drCub.HasRows)
				{
					this.txtClienteIdCre_DatosUbicacion.Value = drCub.GetInt32 (0);
					if (drCub.GetValue (3) != System.DBNull.Value)	this.cmbProvinciaDomicilio.Value = drCub.GetInt32 (3);
					if (drCub.GetValue (4) != System.DBNull.Value)this.cmbCantonDomicilio.Value = drCub.GetInt32 (4);
					if (drCub.GetValue (5)!= System.DBNull.Value)this.cmbParroquiaDomicilio.Value = drCub.GetInt32 (5);
					if (drCub.GetValue (6) != System.DBNull.Value)this.cmbBarrioDomicilio.Value = drCub.GetInt32 (6);
					if (drCub.GetValue (7)!= System.DBNull.Value)this.txtCalle1.Text =  drCub.GetString (7);
					if (drCub.GetValue (8) != System.DBNull.Value)this.txtNumcasa.Text =  drCub.GetString (8);
					if (drCub.GetValue (9)!= System.DBNull.Value)this.txtCalle2.Text =  drCub.GetString (9);
					if (drCub.GetValue (10) != System.DBNull.Value)this.txtReferencia.Text = drCub.GetString (10);
					if (drCub.GetValue (11)!= System.DBNull.Value)this.txtNETelefono1.Text = drCub.GetString (11);
					if (drCub.GetValue (13)!= System.DBNull.Value)this.txtNETelefono2.Text = drCub.GetString (13);
					if (drCub.GetValue (15)!= System.DBNull.Value)this.txtNECelular.Text = drCub.GetString (15);
					if (drCub.GetValue (16)!= System.DBNull.Value)this.txtNEEMail.Text = drCub.GetString (16);
					if (drCub.GetValue (17) != System.DBNull.Value)this.cmbTipoVivienda.Value = drCub.GetInt32 (17);
					if (drCub.GetValue (18) != System.DBNull.Value)this.dtpTiempoVivienda.Value = drCub.GetDateTime (18);
					if (drCub.GetValue (2) != System.DBNull.Value)this.txtClienteIdCre_DatosUbicacion.Value = drCub.GetInt32 (0);
				}
				drCub.Close();
				#endregion Datos Ubicacion 

				#region Conyugue
				if(this.cmbNEEstadoCivil.ActiveRow != null)
				{
//					if((int)this.cmbNEEstadoCivil.Value != 2 && (int)this.cmbNEEstadoCivil.Value != 3 && (int)this.cmbNEEstadoCivil.Value != 5)
//					{
						string sqlDCY = string.Format ("Exec Cre_ConsultaConyugue {0}, 2",idGarante); 
						SqlDataReader drDCon= Funcion.rEscalarSQL(cdsSeteoF , sqlDCY, true);
						drDCon.Read();
						if (drDCon.HasRows)
						{
							if (drDCon.GetValue (0) != System.DBNull.Value)	this.txtIdCre_Conyugue.Value = drDCon.GetInt32 (0);
							if (drDCon.GetValue (2) != System.DBNull.Value)	this.txtIdCre_Solicitud.Value = drDCon.GetInt32 (2);
							if (drDCon.GetValue (4)!= System.DBNull.Value)this.cmbCActividadEconomia.Value = drDCon.GetInt32 (4);
						}
						drDCon.Close();

						string sqlDDY = string.Format ("Exec Cre_ConsultaDatosGenerales {0}, 1",(int)this.txtIdCre_Conyugue.Value); 
						SqlDataReader drDdot= Funcion.rEscalarSQL(cdsSeteoF,sqlDDY,true);
						drDdot.Read();
						if (drDdot.HasRows)
						{
							if (drDdot.GetValue (3) != System.DBNull.Value)this.cmbCTipoRuc.Value = drDdot.GetInt32 (3);
							if (drDdot.GetValue (4) != System.DBNull.Value)this.txtCDocumento.Text = drDdot.GetString (4);
							if (drDdot.GetValue (5)!= System.DBNull.Value)this.txtCApellidoP.Text = drDdot.GetString (5);
							if (drDdot.GetValue (6)!= System.DBNull.Value)this.txtCApellidoM.Text = drDdot.GetString (6);
							if (drDdot.GetValue (7)!= System.DBNull.Value)this.txtCPNombre.Text = drDdot.GetString (7);
							if (drDdot.GetValue (8) != System.DBNull.Value)this.txtCSNombre.Text = drDdot.GetString (8);
							if (drDdot.GetValue (14) != System.DBNull.Value)this.cmbCNacionalidad.Value = drDdot.GetInt32 (14);
							if (drDdot.GetValue (15) != System.DBNull.Value)this.dpCFechaNacimiento.Value = drDdot.GetDateTime (15);
							if (drDdot.GetValue (16) != System.DBNull.Value)this.cmbCSexo.Value = drDdot.GetInt32 (16);
							if (drDdot.GetValue (17) != System.DBNull.Value)this.cmbCProvinciaNacimiento.Value = drDdot.GetInt32 (17);
							if (drDdot.GetValue (18) != System.DBNull.Value)this.cmbCCantonNacimiento.Value = drDdot.GetInt32 (18);
							if (drDdot.GetValue (21) != System.DBNull.Value)this.cmbCNivel.Value = drDdot.GetInt32 (21);
							if (drDdot.GetValue (22) != System.DBNull.Value)this.cmbCProfesion.Value = drDdot.GetInt32 (22);
							if (drDdot.GetValue (23) != System.DBNull.Value)this.cmbCSituacionLaboral.Value = drDdot.GetInt32 (23);
							if (drDdot.GetValue (24) != System.DBNull.Value)this.cmbCSituacionLaboral.Value = drDdot.GetInt32 (24);
//							SituacionLaboral(ientidad);
							if (drDdot.GetValue (25) != System.DBNull.Value)this.cmbCActividadEconomia.Value = drDdot.GetInt32 (25);
						}
						drDdot.Close();
//					}
				}
			
				#endregion Conyugue

				#region Referencias

				string sqlGRP = string.Format ("Exec Cre_ConsultaReferencias {0}, 1", (int)this.txtClienteIdCre_DatosGenerales.Value);  
				this.uGridParentesco.DataSource = Funcion.dvProcedimiento(cdsSeteoF, sqlGRP);
//				if (	this.uGridParentesco.Rows.Count > 0) 	this.uGridParentesco.Focus(); else this.txtDocumento.Focus();

				#endregion Referencias 

				#region Negocio 

				string sqlNG = string.Format ("Exec Cre_ConsultaNegocions {0}", (int)this.txtClienteIdCre_DatosGenerales.Value);  
				SqlDataReader drNeg= Funcion.rEscalarSQL(cdsSeteoF, sqlNG , true);
				drNeg.Read();
				if (drNeg.HasRows)
				{
          this.txtClienteIdCre_Negocio.Value = drNeg.GetInt32(0);
					this.txtClienteIdCre_DatosGenerales.Value = drNeg.GetInt32(1);
					if (drNeg.GetValue (2) != System.DBNull.Value)	this.txtSNRuc.Value = drNeg.GetString (2);
					if (drNeg.GetValue (3) != System.DBNull.Value)	this.txtSNNombre.Value = drNeg.GetString (3);
					if (drNeg.GetValue (4)!= System.DBNull.Value)   this.dtSNFechaInicio.Value = drNeg.GetDateTime(4);
					if (drNeg.GetValue (7) != System.DBNull.Value)	this.txtSNMetros.Value = drNeg.GetInt32 (7);
					if (drNeg.GetValue (8) != System.DBNull.Value)	this.txtSNEmpleados.Value = drNeg.GetInt32 (8);
					if (drNeg.GetValue (5) != System.DBNull.Value)	this.cmbSCTipoNegocio.Value = drNeg.GetInt32 (5);
					if (drNeg.GetValue (6) != System.DBNull.Value)	this.cmbSubActvidadEconomica.Value = drNeg.GetInt32 (6);
					if (drNeg.GetValue (9) != System.DBNull.Value)	this.chkSNContabilidad.Checked = drNeg.GetBoolean (9);
					if (drNeg.GetValue (10) != System.DBNull.Value)	this.chkSNAfiliacion.Checked = drNeg.GetBoolean (10);
				}
			  drNeg.Close();

				string sqlDNG = string.Format ("Exec Cre_ConsultaDatosUbicacion {0}, 3", (int)this.txtClienteIdCre_Negocio.Value);  
				SqlDataReader drDNeg= Funcion.rEscalarSQL(cdsSeteoF, sqlDNG , true);
				drDNeg.Read();
				if (drDNeg.HasRows)
				{
          this.txtNegocioIdCre_DatosUbicacion.Value= drDNeg.GetInt32 (0);
					if (drDNeg.GetValue (3) != System.DBNull.Value)this.cmbSNProvincia.Value = drDNeg.GetInt32 (3);
		      if (drDNeg.GetValue (4) != System.DBNull.Value)this.cmbSNCanton.Value = drDNeg.GetInt32 (4);
					if (drDNeg.GetValue (5) != System.DBNull.Value)this.cmbSNParroquia.Value = drDNeg.GetInt32 (5);
					if (drDNeg.GetValue (6) != System.DBNull.Value)this.cmbSNBarrio.Value = drDNeg.GetInt32 (6);
					this.txtSNCalle1.Text =  drDNeg.GetString (7);
					this.txtSNNumCasa.Text =  drDNeg.GetString (8);
			  	this.txtSNCalle2.Text =  drDNeg.GetString (9);
					this.txtSNReferencia.Text = drDNeg.GetString (10);
				  this.txtSNTelefono1.Text = drDNeg.GetString (11);
				  this.txtSNTelefono2.Text = drDNeg.GetString (13);
					this.txtSNCelular.Text = drDNeg.GetString (15);
			    this.txtSNEmail.Text = drDNeg.GetString (16);
			    this.txtSNIngresos.Value = drDNeg.GetInt32 (21);
        	this.txtSNGastos.Value = drDNeg.GetInt32 (22);
				  this.txtSNActivos.Value = drDNeg.GetInt32 (23);
				  this.txtSNPasivos.Value = drDNeg.GetInt32 (24);
				}
				drDNeg.Close();
				#endregion Negocio

				#region Empresa
	
				string sqlEM = string.Format ("Exec Cre_ConsultaTrabajo  {0}", (int)this.txtClienteIdCre_DatosGenerales.Value);  
				SqlDataReader drEmp= Funcion.rEscalarSQL(cdsSeteoF, sqlEM , true);
				drEmp.Read();
				if (drEmp.HasRows)
				{
					this.txtClienteIdCre_EmpresasTrabajo.Value = drEmp.GetInt32 (0);
				  this.txtSDRuc.Value = drEmp.GetString (2);
			    this.txtSDEmpresa.Value = drEmp.GetString (3);
	        this.cmbSDTipoEmpresa.Value = drEmp.GetInt32 (4);
          this.dtSDFecha.Value = drEmp.GetDateTime(5);
					this.cmbSDTipoContrato.Value = drEmp.GetInt32 (7);
					this.cmbSDTipoSueldo.Value = drEmp.GetInt32 (8);
          this.chkSDAfiliacion.Checked = drEmp.GetBoolean (9);
          this.txtSDPago.Value = drEmp.GetInt32(10);
          this.txtSDDepartamento.Value = drEmp.GetString (11);
				  this.cmbSDCargo.Value = drEmp.GetInt32 (12);
				}
			 drEmp.Close();

				string sqlDEM = string.Format ("Exec Cre_ConsultaDatosUbicacion {0}, 4", (int)this.txtClienteIdCre_EmpresasTrabajo.Value);  
				SqlDataReader drDEmp= Funcion.rEscalarSQL(cdsSeteoF, sqlDEM   , true);
				drDEmp.Read();
				if (drDEmp.HasRows)
				{
					this.txtEmpresaIdCre_DatosUbicacion.Value = drDEmp.GetInt32 (0);
					if (drDEmp.GetValue (3) != System.DBNull.Value) this.cmbSDProvincia.Value =drDEmp.GetInt32 (3);
					if (drDEmp.GetValue (4) != System.DBNull.Value)this.cmbSDCanton.Value = drDEmp.GetInt32 (4);
				  if (drDEmp.GetValue (5) != System.DBNull.Value)this.cmbSDParroquia.Value = drDEmp.GetInt32 (5);
			   	if (drDEmp.GetValue (6) != System.DBNull.Value)this.cmbSDBarrio.Value = drDEmp.GetInt32 (6);
					this.txtSDCalle1.Text =  drDEmp.GetString (7);
					this.txtSDNumCasa.Text =  drDEmp.GetString (8);
					this.txtSDCalle2.Text =  drDEmp.GetString (9);
					this.txtSDReferencia.Text = drDEmp.GetString (10);
					this.txtSDTelefono.Text = drDEmp.GetString (11);
					this.txtSDExt.Text = drDEmp.GetString (12);
				  this.txtSDCelular.Text = drDEmp.GetString (15);
					this.txtSDEmail.Text = drDEmp.GetString (16);
				  this.txtSDIngresos.Value = drDEmp.GetInt32 (21);
					this.txtSDGastos.Value = drDEmp.GetInt32 (22);
					this.txtSDActivos.Value = drDEmp.GetInt32 (23);
				this.txtSDPasivos.Value = drDEmp.GetInt32 (24);
				}
				drDEmp.Close();
				#endregion Empresa
			}
		}


		private void btnGuardar_Click(object sender, System.EventArgs e)
		{
			#region Validacion Garante
			#region Campos

			if (!Validacion.vbComboVacio(this.cmbTipoRuc, "Seleccione el tipo de Documento "))
			{ 
				this.utcSolicitud.Tabs[0].Selected = true; 
				return;
			}

			if (!Validacion.vbTexto(this.txtDocumento, 3, 10, "Ingrese la Cedula del Garante ")) 
			{ 
				this.utcSolicitud.Tabs[0].Selected = true; 
				this.txtDocumento.Focus();
				return;
			}

			if (!Validacion.vbTexto(this.txtApellidoP, 4, 60, "Ingrese el Apellido Paterno"))
			{
				this.utcSolicitud.Tabs[0].Selected = true;
				this.txtApellidoP.Focus();
				return;
			} 
			if (!Validacion.vbTexto(this.txtPnombre, 4, 60, "Ingrese el Primer Nombre"))
			{
				this.utcSolicitud.Tabs[0].Selected = true;
				this.txtPnombre.Focus();
				return;
			}			

			if (!Validacion.vbComboVacio(this.cmbEstado, "Seleccione un Estado")) 
			{ 
				this.utcSolicitud.Tabs[0].Selected = true; 
				this.cmbEstado.Focus();
				return;
			}
			if (!Validacion.vbComboVacio(this.cmbCalificacion, "Seleccione una Calificacion"))
			{ 
				this.utcSolicitud.Tabs[0].Selected = true; 
				this.cmbCalificacion.Focus();
				return;
			}
			if(!Validacion.vbCampoEnteroVacio(this.txtScore,"Ingrese el Score",1,999))
			{ 
				this.utcSolicitud.Tabs[0].Selected = true; 
				this.txtScore.Focus();
				return;
			}
			if (!Validacion.vbComboVacio(this.cmbNENacionalidad, "Seleccione una Nacionalidad")) 
			{ 
				this.utcSolicitud.Tabs[0].Selected = true;
				this.cmbNENacionalidad.Focus();
				return;
			}

			//VALIDA FECHA
			if (!Validacion.vbFechaEnDocumentos(this.dtpNEFechaNacimiento, "Ingrese la fecha de Nacimiento", false, "", cdsSeteoF)) return;
			DateTime dtFecha = DateTime.Today.AddYears(-21);
			if (Convert.ToDateTime(this.dtpNEFechaNacimiento.Text) > dtFecha)
			{
				MessageBox.Show("La Edad Minima del Cliente debe de ser 21 años", "Validacion", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				this.utcSolicitud.Tabs[0].Selected = true; 
				this.dtpNEFechaNacimiento.Focus();
				return;
			}
			dtFecha = DateTime.Today.AddYears(-70);
			if (Convert.ToDateTime(this.dtpNEFechaNacimiento.Text) < dtFecha)
			{
				MessageBox.Show("La Edad Maxima del Cliente debe de ser 70 años", "Validacion", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				this.utcSolicitud.Tabs[0].Selected = true; 
				this.dtpNEFechaNacimiento.Focus();
				return;
			}

			if (!Validacion.vbComboVacio(this.cmbSexo, "Seleccione el Genero")) 
			{ 
				this.utcSolicitud.Tabs[0].Selected = true;
				this.cmbSexo.Focus();
				return;
			}
			if (!Validacion.vbComboVacio(this.cmbProvinciaNacimiento, "Seleccione Provincia de Nacimiento")) 
			{ 
				this.utcSolicitud.Tabs[0].Selected = true;
				this.cmbProvinciaNacimiento.Focus();
				return;
			}
			if (!Validacion.vbComboVacio(this.cmbCantonNacimiento, "Seleccione Canton de Nacimiento"))
			{ 
				this.utcSolicitud.Tabs[0].Selected = true; 
				this.cmbCantonNacimiento.Focus();
				return;
			}
			if (!Validacion.vbComboVacio(this.cmbNEEstadoCivil, "Seleccione Estado Civil")) 
			{ 
				this.utcSolicitud.Tabs[0].Selected = true; 
				this.cmbNEEstadoCivil.Focus();
				return;
			}
			if(!Validacion.vbCampoEnteroVacio(this.txtNECargasFamiliares,"Ingrese Numero de Dependientes",1,99))
			{ 
				this.utcSolicitud.Tabs[0].Selected = true; 
				this.txtNECargasFamiliares.Focus();
				return;
			}
			if (!Validacion.vbComboVacio(this.cmbNivel, "Seleccione el Nivel de Educación")) 
			{ 
				this.utcSolicitud.Tabs[0].Selected = true; 
				this.cmbNivel.Focus();
				return;
			}
			if (!Validacion.vbComboVacio(this.cmbProfesion, "Seleccione una Profesión")) 
			{ 
				this.utcSolicitud.Tabs[0].Selected = true; 
				this.cmbProfesion.Focus();
				return;
			}
			if (!Validacion.vbComboVacio(this.cmbSituacionLaboral, "Seleccione una Situación Laboral")) 
			{ 
				this.utcSolicitud.Tabs[0].Selected = true; 
				this.cmbSituacionLaboral.Focus();
				return;
			}
			if (!Validacion.vbComboVacio(this.cmbActividadEcomica, "Seleccione una Actividad Económica")) 
			{ 
				this.utcSolicitud.Tabs[0].Selected = true; 
				this.cmbActividadEcomica.Focus();
				return;
			}
			if (!Validacion.vbComboVacio(this.cmbProvinciaDomicilio, "Seleccione la Provincia de Domicilio"))
			{ 
				this.utcSolicitud.Tabs[0].Selected = true; 
				this.cmbProvinciaDomicilio.Focus();
				return;
			}
			if (!Validacion.vbComboVacio(this.cmbCantonDomicilio, "Seleccione el Canton de Domicilio")) 
			{ 
				this.utcSolicitud.Tabs[0].Selected = true; 
				this.cmbCantonDomicilio.Focus();
				return;
			}
			if (!Validacion.vbComboVacio(this.cmbParroquiaDomicilio, "Seleccione la Parroquia de Domicilio")) 
			{ 
				this.utcSolicitud.Tabs[0].Selected = true;
				this.cmbParroquiaDomicilio.Focus();
				return;
			}
			if (!Validacion.vbComboVacio(this.cmbBarrioDomicilio, "Seleccione el Barrio de Domicilio")) 
			{ 
				this.utcSolicitud.Tabs[0].Selected = true; 
				this.cmbParroquiaDomicilio.Focus();
				return;
			}
			if (!Validacion.vbTexto(this.txtCalle1, 4, 70, "Ingrese la calle Principal")) 
			{ 
				this.utcSolicitud.Tabs[0].Selected = true; 
				this.cmbParroquiaDomicilio.Focus();
				return;
			}
			if (!Validacion.vbTexto(this.txtNumcasa, 4, 30, "Ingrese el número de casa")) 
			{ 
				this.utcSolicitud.Tabs[0].Selected = true; 
				this.txtNumcasa.Focus();
				return;
			}
			if (!Validacion.vbTexto(this.txtCalle2, 4, 70, "Ingrese la calle Secundaria"))
			{ 
				this.utcSolicitud.Tabs[0].Selected = true;
				this.txtCalle2.Focus();
				return;
			}
			if (!Validacion.vbTexto(this.txtReferencia, 2, 100, "Ingrese una Referencia")) 
			{ 
				this.utcSolicitud.Tabs[0].Selected = true; 
				this.txtReferencia.Focus();
				return;
			}
			if (this.txtNEEMail.Text.ToString().Trim().Length > 0 )
			{
				if (!Validacion.vbTexto(this.txtNEEMail, 0, 60, "E-Mail")) return;
				if (!Validacion.bEMail(this.txtNEEMail.Text.ToString())) 
				{ 
					this.utcSolicitud.Tabs[0].Selected = true; 
					return;
				}
			}

			#endregion Campos

			#region Valida Telefono
			if (this.txtNETelefono1.Text.ToString().Trim().Length == 0  && this.txtNECelular.Text.ToString().Trim().Length == 0  )
			{
				MessageBox.Show(string.Format("Debe Ingresar el Telefono Convencional o el Telefono Celular"), "Point Technology", MessageBoxButtons.OK,MessageBoxIcon.Stop);
				this.txtNETelefono1.Focus();
				this.utcSolicitud.Tabs[0].Selected = true; 		

				return;	
			}
			if (!Validacion.vbValidaTelefono(this.txtNETelefono1, 1, false)) 
			{ 
				this.utcSolicitud.Tabs[0].Selected = true; 
				return;
			}
			if (!Validacion.vbValidaTelefono(this.txtNECelular, 2, false)) 
			{ 
				this.utcSolicitud.Tabs[0].Selected = true; 
				return;
			}
			#endregion Valida Telefono

			if (!Validacion.vbComboVacio(this.cmbTipoVivienda, "Seleccione el Tipo de Vivienda "))
			{
				this.utcSolicitud.Tabs[0].Selected = true; 
				return;
			}

			if (!Validacion.vbFechaEnDocumentos(this.dtpTiempoVivienda, "Ingrese el Tiempo de Vivienda", false, "", cdsSeteoF))
			{
				this.utcSolicitud.Tabs[0].Selected = true; 
				return;
			}

			if(!Validacion.vbComparaFechas((DateTime)this.dtpNEFechaNacimiento.Value,(DateTime)this.dtpTiempoVivienda.Value,this.dtpTiempoVivienda,"La Fecha de vivienda no puede ser menor a la Fecha de Nacimiento",false))
			{
				this.utcSolicitud.Tabs[0].Selected = true;
				return;
			}

			#endregion Validacion Garante 

			int iReferencia = 0;
			foreach(Infragistics.Win.UltraWinGrid.UltraGridRow drU in this.uGridParentesco.Rows.All)
			{
				iReferencia =  iReferencia + 1;
				drU.Cells["Indice"].Value = iReferencia;
			}

			#region ValidacionConyugue

			if (this.cmbNEEstadoCivil.ActiveRow != null)
			{
				if ((int)this.cmbNEEstadoCivil.Value != 2 && (int)this.cmbNEEstadoCivil.Value != 3 && (int)this.cmbNEEstadoCivil.Value != 5)
				{
					#region Valida_DatosConyuge

					if (!Validacion.vbComboVacio(this.cmbCTipoRuc, "Seleccione el Tipo de Documento")) 
					{ 
						this.utcSolicitud.Tabs[1].Selected = true; 
						return;
					}
					if (!Validacion.vbTexto(this.txtCDocumento, 3, 10, "Ingrese la Cedula del Conyuge ")) 
					{ 
						this.utcSolicitud.Tabs[1].Selected = true;
						this.txtCDocumento.Focus();
						return;
					}
					if (!Validacion.vbTexto(this.txtCApellidoP, 3, 60, "Ingrese el Apellido Paterno del Conyuge"))
					{ 
						this.utcSolicitud.Tabs[1].Selected = true; 
						this.txtCApellidoP.Focus();
						return;
					}
					if (!Validacion.vbTexto(this.txtCPNombre, 3, 60, "Ingrese el Primer Nombre del Conyuge")) 
					{ 
						this.utcSolicitud.Tabs[1].Selected = true; 
						this.txtCPNombre.Focus();
						return;
					}
					if (!Validacion.vbComboVacio(this.cmbCNacionalidad, "Seleccione la Nacionalidad del Conyuge "))
					{ 
						this.utcSolicitud.Tabs[1].Selected = true; 
						this.cmbCNacionalidad.Focus();
						return;
					}
					if (!Validacion.vbFechaEnDocumentos(this.dpCFechaNacimiento, "Ingrese la fecha de Nacimiento del Conyuge ", false, "", cdsSeteoF)) 
					{ 
						this.utcSolicitud.Tabs[1].Selected = true; 
						this.dpCFechaNacimiento.Focus();
						return;
					}
					if (!Validacion.vbComboVacio(this.cmbCSexo, "Seleccione el Genero del Conyuge "))
					{ 
						this.utcSolicitud.Tabs[1].Selected = true; 
						this.cmbCSexo.Focus();
						return;
					}
					if (!Validacion.vbComboVacio(this.cmbCProvinciaNacimiento, "Seleccione la Provincia de Nacimiento del Conyuge ")) 
					{ 
						this.utcSolicitud.Tabs[1].Selected = true; 
						this.cmbCProvinciaNacimiento.Focus();
						return;
					}
					if (!Validacion.vbComboVacio(this.cmbCCantonNacimiento, "Seleccione el Canton de Nacimiento del Conyuge "))
					{ 
						this.utcSolicitud.Tabs[1].Selected = true; 
						this.cmbCCantonNacimiento.Focus();
						return;
					}
					if (!Validacion.vbComboVacio(this.cmbCNivel, "Seleccione el Nivel de Eduación del Conyuge "))
					{ 
						this.utcSolicitud.Tabs[1].Selected = true; 
						this.cmbCNivel.Focus();
						return;
					}
					if (!Validacion.vbComboVacio(this.cmbCProfesion, "Seleccione la Profesión del Conyuge ")) 
					{ 
						this.utcSolicitud.Tabs[1].Selected = true; 
						this.cmbCProfesion.Focus();
						return;
					}
					if (!Validacion.vbComboVacio(this.cmbCSituacionLaboral, "Seleccione Situación Laboral del Conyugue ")) 
					{ 
						this.utcSolicitud.Tabs[1].Selected = true;
						this.cmbCSituacionLaboral.Focus();
						return;
					}
					if (!Validacion.vbComboVacio(this.cmbCActividadEconomia, "Seleccione Actividad Económica del Conyugue")) 
					{ 
						this.utcSolicitud.Tabs[1].Selected = true;
						this.cmbCActividadEconomia.Focus();
						return;
					}
					#endregion Valida_DatosConyuge	  
				}
			}
			#endregion ValidacionConyugue

			#region Validacion_Negocio/Empresa

			if (this.cmbSituacionLaboral.ActiveRow != null)		
			{
				#region Validacion_Negocio
	
				if ((ientidad == 2 && (int)this.cmbSituacionLaboral.Value > 10) || (ientidad != 2 && ((int)this.cmbSituacionLaboral.Value != 1 && (int) this.cmbSituacionLaboral.Value != 4)))
				{
					//					if (!Validacion.vbTexto(this.txtSNRuc, 3, 13, "Ingrese el Ruc del Negocio")) 	
					//					{ 
					//						this.utcSolicitud.Tabs[3].Selected = true;
					//						this.txtSNRuc.Focus();
					//						return;
					//					}
					if (!Validacion.vbTexto(this.txtSNNombre, 3, 200, "Ingrese el Nombre del Negocio"))	
					{ 
						this.utcSolicitud.Tabs[3].Selected = true; 
						this.txtSNNombre.Focus();
						return;
					}

					if (!Validacion.vbFechaEnDocumentos(this.dtSNFechaInicio, "Ingrese la fecha de inicio del Negocio ", false, "", cdsSeteoF)) 	
					{ 
						this.utcSolicitud.Tabs[3].Selected = true;
						this.dtSNFechaInicio.Focus();
						return;
					}

					if(!Validacion.vbComparaFechas((DateTime)this.dtpNEFechaNacimiento.Value,(DateTime)this.dtSNFechaInicio.Value,this.dtSNFechaInicio,"Inicio del Negocio no puede ser mayor a la edad del cliente",false))
					{
						this.utcSolicitud.Tabs[3].Selected = true;
						this.dtSNFechaInicio.Focus();
						return;
					}

					if (!Validacion.vbCampoEnteroVacio(this.txtSNMetros,"Ingrese los Metros del Negocio ",1,10000))	
					{ 
						this.utcSolicitud.Tabs[3].Selected = true; 
						this.txtSNMetros.Focus();
						return;
					}
					//					if (!Validacion.vbCampoEnteroVacio(this.txtSNEmpleados,"Ingrese el Número de Empleados del Negocio ",1,500))	
					//					{ 
					//						this.utcSolicitud.Tabs[3].Selected = true;
					//						this.txtSNEmpleados.Focus();
					//						return;
					//					}
					//					if (!Validacion.vbComboVacio(this.cmbSCTipoNegocio, "Seleccione Tipo de Negocio")) 	
					//					{ 
					//						this.utcSolicitud.Tabs[3].Selected = true; 
					//						this.cmbSCTipoNegocio.Focus();
					//						return;
					//					}
					//					if (!Validacion.vbComboVacio(this.cmbSubActvidadEconomica, "Seleccione Tipo de Negocio"))	
					//					{ 
					//						this.utcSolicitud.Tabs[3].Selected = true; 
					//						this.cmbSubActvidadEconomica.Focus();
					//						return;
					//					}
					if (!Validacion.vbComboVacio(this.cmbSNProvincia, "Seleccione la Provicia de la Ubicación del Negocio"))	
					{ 
						this.utcSolicitud.Tabs[3].Selected = true; 
						this.cmbSNProvincia.Focus();
						return;
					}
					if (!Validacion.vbComboVacio(this.cmbSNCanton, "Seleccione el Canton de la Ubicación del Negocio")) 	
					{ 
						this.utcSolicitud.Tabs[3].Selected = true;
						this.cmbSNCanton.Focus();
						return;
					}
					if (!Validacion.vbComboVacio(this.cmbSNParroquia, "Selecione la Parroquia de la Ubicación del Negocio")) 	
					{ 
						this.utcSolicitud.Tabs[3].Selected = true;
						this.cmbSNParroquia.Focus(); 
						return;
					}
					if (!Validacion.vbComboVacio(this.cmbSNBarrio, "Seleccione el Barrio de la Ubicación del Negocio")) 	
					{ 
						this.utcSolicitud.Tabs[3].Selected = true;
						this.cmbSNBarrio.Focus(); 
						return;
					}
					if (!Validacion.vbTexto(this.txtSNCalle1, 3, 100, "Ingrese la Calle Principal del Negocio")) 	
					{ 
						this.utcSolicitud.Tabs[3].Selected = true; 
						this.txtSNCalle1.Focus(); 
						return;
					}
					if (!Validacion.vbTexto(this.txtSNNumCasa, 3, 60, "Ingrese el Numero de Casa del Negocio")) 	
					{ 
						this.utcSolicitud.Tabs[3].Selected = true;
						this.txtSNNumCasa.Focus(); 
						return;
					}
					if (!Validacion.vbTexto(this.txtSNCalle2, 3, 100, "Ingrese la Calle Secundaria del Negocio")) 	
					{ 
						this.utcSolicitud.Tabs[3].Selected = true; 
						this.txtSNCalle2.Focus(); 
						return;
					}
					//					if (this.txtSNTelefono1.Text.ToString().Trim().Length == 0 &&  this.txtSNCelular.Text.ToString().Trim().Length == 0 )
					//					{
					//						MessageBox.Show(string.Format("Debe Ingresar el Telefono Convencional o el Telefono Celular del Negocio "), "Point Technology", MessageBoxButtons.OK,MessageBoxIcon.Stop);
					//						this.utcSolicitud.Tabs[3].Selected = true; 
					//						this.txtSNTelefono1.Focus();
					//						return;	
					//					}
					if (!Validacion.vbValidaTelefono(this.txtSNTelefono1, 1, false)) 	
					{ 
						this.utcSolicitud.Tabs[3].Selected = true; 
						return;
					}
					if (!Validacion.vbValidaTelefono(this.txtSNTelefono2, 1, false)) 	
					{ 
						this.utcSolicitud.Tabs[3].Selected = true; 
						return;
					}
					if (!Validacion.vbValidaTelefono(this.txtSNCelular, 2, false)) 	
					{ 
						this.utcSolicitud.Tabs[3].Selected = true; 
						return;
					}
					if (!Validacion.vbTexto(this.txtSNReferencia, 5, 300, "Ingrese la Referencia de la Ubicación del Negocio")) 	
					{ 
						this.utcSolicitud.Tabs[3].Selected = true; 
						this.txtSNReferencia.Focus();
						return;
					}
					if (!Validacion.vbCampoEnteroVacio(this.txtSNIngresos,"Ingrese los Ingresos del Negocio ",1,1000000))	
					{ 
						this.utcSolicitud.Tabs[3].Selected = true;
						this.txtSNIngresos.Focus(); 
						return;
					}
					if (!Validacion.vbCampoEnteroVacio(this.txtSNGastos,"Ingrese los Gastos del Negocio  ",1,1000000))	
					{ 
						this.utcSolicitud.Tabs[3].Selected = true;
						this.txtSNGastos.Focus(); 
						return;
					}
					if(Convert.ToInt32(this.txtSNIngresos.Value) < Convert.ToInt32(this.txtSNGastos.Value))
					{
						MessageBox.Show(string.Format("Los Ingresos no deben ser menores a los Gastos, mínimo un 25%  de su Sueldo Actual"), "Point Technology", MessageBoxButtons.OK,MessageBoxIcon.Stop);
						this.utcSolicitud.Tabs[3].Selected = true; 
						this.txtSNIngresos.Focus();
						return;
					}
					if (!Validacion.vbCampoEnteroVacio(this.txtSNActivos,"Ingrese los Activos del Negocio ",1,1000000))
					{ 
						this.utcSolicitud.Tabs[3].Selected = true;
						this.txtSNActivos.Focus();  
						return;
					}
					if (!Validacion.vbCampoEnteroVacio(this.txtSNPasivos,"Ingrese los  Pasivos del Negocio ",1,1000000))
					{ 
						this.utcSolicitud.Tabs[3].Selected = true;
						this.txtSNPasivos.Focus(); 
						return;
					}
				}
				#endregion Validacion_Negocio

				#region Validacion_Empresa

				if ((ientidad == 2 && Convert.ToInt32(this.cmbSituacionLaboral.Value) == 10) || 
					(ientidad != 2 && (Convert.ToInt32(this.cmbSituacionLaboral.Value) == 1 || Convert.ToInt32(this.cmbSituacionLaboral.Value) == 3)))		
				{
					//					if (!Validacion.vbTexto(this.txtSDRuc, 10, 13, "Ingrese el Ruc de la Empresa")) 
					//					{ 
					//						this.utcSolicitud.Tabs[4].Selected = true; 
					//						this.txtSDRuc.Focus();
					//						return;
					//					}
					if (!Validacion.vbTexto(this.txtSDEmpresa, 3, 200, "Ingrese el Nombre del Negocio")) 
					{ 
						this.utcSolicitud.Tabs[4].Selected = true; 
						this.txtSDEmpresa.Focus();
						return;
					}
          
					if (!Validacion.vbFechaEnDocumentos(this.dtSDFecha, "Ingrese la fecha de inicio del Negocio ", false, "", cdsSeteoF)) 	
					{ 
						this.utcSolicitud.Tabs[4].Selected = true; 
						this.dtSDFecha.Focus();
						return;
					}

					if(!Validacion.vbComparaFechas((DateTime)this.dtpNEFechaNacimiento.Value,(DateTime)this.dtSDFecha.Value,this.dtSDFecha,"Fecha de Ingreso al Trabajo no puede ser mayor a la edad del cliente",false))
					{
						this.utcSolicitud.Tabs[4].Selected = true;
						this.dtSDFecha.Focus();
						return;
					}

					if (!Validacion.vbComboVacio(this.cmbSDTipoEmpresa, "Seleccione Tipo de Empresa")) 	
					{ 
						this.utcSolicitud.Tabs[4].Selected = true; 
						this.cmbSDTipoEmpresa.Focus();
						return;
					}
					if (!Validacion.vbComboVacio(this.cmbSDTipoContrato, "Seleccione el Tipo de Contrato")) 	
					{ 
						this.utcSolicitud.Tabs[4].Selected = true; 
						this.cmbSDTipoContrato.Focus();
						return;
					}
					if (!Validacion.vbComboVacio(this.cmbSDTipoSueldo, "Seleccione el Tipo de Sueldo")) 	
					{ 
						this.utcSolicitud.Tabs[4].Selected = true; 
						this.cmbSDTipoSueldo.Focus();
						return;
					} 
					if (!Validacion.vbTexto(this.txtSDDepartamento, 4, 100, "Ingrese el Departamento")) 	
					{ 
						this.utcSolicitud.Tabs[4].Selected = true;
						this.txtSDDepartamento.Focus();
						return;
					}
					if (!Validacion.vbComboVacio(this.cmbSDCargo , "Seleccione Cargo")) 	
					{ 
						this.utcSolicitud.Tabs[4].Selected = true;
						this.cmbSDCargo.Focus();
						return;
					}
					if (!Validacion.vbComboVacio(this.cmbSDProvincia, "Seleccione la Provicia de la Ubicación de la Empresa")) 	
					{ 
						this.utcSolicitud.Tabs[4].Selected = true; 
						this.cmbSDCargo.Focus();
						return;
					}
					if (!Validacion.vbComboVacio(this.cmbSDCanton, "Seleccione el Canton de la Ubicación de la Empresa")) 	
					{ 
						this.utcSolicitud.Tabs[4].Selected = true;
						this.cmbSDCanton.Focus();
						return;
					}
					if (!Validacion.vbComboVacio(this.cmbSDParroquia, "Selecione la Parroquia de la Ubicación de la Empresa")) 	
					{ 
						this.utcSolicitud.Tabs[4].Selected = true; 
						this.cmbSDParroquia.Focus();
						return;
					}
					if (!Validacion.vbComboVacio(this.cmbSDBarrio, "Seleccione el Barrio de la Ubicación de la Empresa")) 	
					{ 
						this.utcSolicitud.Tabs[4].Selected = true;
						this.cmbSDBarrio.Focus();
						return;
					}
					if (!Validacion.vbTexto(this.txtSDCalle1, 3, 100, "Ingrese la Calle Principal de la Empresa")) 	
					{ 
						this.utcSolicitud.Tabs[4].Selected = true; 
						this.txtSDCalle1.Focus();
						return;
					}
					if (!Validacion.vbTexto(this.txtSDNumCasa, 3, 60, "Ingrese el Numero de Casa de la Empresa")) 	
					{ 
						this.utcSolicitud.Tabs[4].Selected = true;
						this.txtSDNumCasa.Focus();
						return;
					}
					if (!Validacion.vbTexto(this.txtSDCalle2, 3, 100, "Ingrese la Calle Secundaria de la Empresa")) 	
					{ 
						this.utcSolicitud.Tabs[4].Selected = true; 
						this.txtSDCalle2.Focus();
						return;
					}
					//					if (!Validacion.vbTexto(this.txtSDExt, 3, 100, "Ingrese la Calle Secundaria de la Empresa")) 	
					//					{ 
					//						this.utcSolicitud.Tabs[4].Selected = true; 
					//						this.txtSDCalle2.Focus();
					//						return;
					//					}
					if (this.txtSDTelefono.Text.ToString().Trim().Length == 0 &&  this.txtSDCelular.Text.ToString().Trim().Length == 0 )
					{
						MessageBox.Show(string.Format("Debe Ingresar el Telefono Convencional o el Telefono Celular del Negocio "), "Point Technology", MessageBoxButtons.OK,MessageBoxIcon.Stop);
						this.utcSolicitud.Tabs[4].Selected = true; 
						this.txtSDTelefono.Focus();
						return;	
					}
					if (!Validacion.vbValidaTelefono(this.txtSDTelefono, 1, false)) 	
					{ 
						this.utcSolicitud.Tabs[4].Selected = true; 
						return;
					}
					if (!Validacion.vbValidaTelefono(this.txtSDCelular, 2, false)) 	
					{ 
						this.utcSolicitud.Tabs[4].Selected = true; 
						return;
					}
					if (!Validacion.vbTexto(this.txtSDReferencia, 10, 300, "Ingrese la Referencia de la Ubicación de la Empresa")) 	
					{ 
						this.utcSolicitud.Tabs[4].Selected = true; 
						this.txtSDReferencia.Focus();
						return;
					}
					if (!Validacion.vbCampoEnteroVacio(this.txtSDPago,"Ingrese el Día de Pago",1,1000000))
					{ 
						this.utcSolicitud.Tabs[4].Selected = true;
						this.txtSDPago.Focus();
						return;
					}
					if (!Validacion.vbCampoEnteroVacio(this.txtSDIngresos,"Ingrese los Ingresos del Negocio ",1,1000000))	
					{ 
						this.utcSolicitud.Tabs[4].Selected = true;
						this.txtSDIngresos.Focus();
						return;
					}
					if (!Validacion.vbCampoEnteroVacio(this.txtSDGastos,"Ingrese los Gastos de la Empresa ",1,1000000))	
					{ 
						this.utcSolicitud.Tabs[4].Selected = true; 
						this.txtSDGastos.Focus();
						return;
					}
					if(Convert.ToInt32(this.txtSDIngresos.Value) < Convert.ToInt32(this.txtSDGastos.Value))
					{
						MessageBox.Show(string.Format("Los Ingresos no deben ser menores a los Gastos, mínimo un 25%  de su Sueldo Actual"), "Point Technology", MessageBoxButtons.OK,MessageBoxIcon.Stop);
						this.txtSDIngresos.Focus();
						return;
					}
					if (!Validacion.vbCampoEnteroVacio(this.txtSDActivos,"Ingrese los Activos de la Empresa ",1,1000000))	
					{ 
						this.utcSolicitud.Tabs[4].Selected = true;
						this.txtSDActivos.Focus();
						return;
					}
					if (!Validacion.vbCampoEnteroVacio(this.txtSDPasivos,"Ingrese los de la Empresa",1,1000000))
					{ 
						this.utcSolicitud.Tabs[4].Selected = true; 
						this.txtSDPasivos.Focus();
						return;
					}
				}
				#endregion Validacion_Empresa
			}	
			#endregion Validacion_Negocio/Empresa

			#region Referencias

		
			if (this.uGridParentesco.Rows.Count < 2)
			{
				MessageBox.Show("Debe Ingresar al menos 2 referencia", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				this.utcSolicitud.Tabs[2].Selected = true;
				this.uGridParentesco.Focus();
				return;
			}
			foreach(Infragistics.Win.UltraWinGrid.UltraGridRow drReferencias in this.uGridParentesco.Rows.All)
			{	
				#region Parentesco

				int idParentesco = 0;
				if (drReferencias.Cells["idParentesco"].Value != System.DBNull.Value) idParentesco = (int)drReferencias.Cells["idParentesco"].Value;	
				if (idParentesco == 0)
				{
					MessageBox.Show(string.Format("Seleccione un Parentesco de la Referencia N. {0}",drReferencias.Index + 1),"Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
					this.utcSolicitud.Tabs[2].Selected = true;
					this.uGridParentesco.ActiveRow = drReferencias;
					this.uGridParentesco.ActiveRow.Selected = true;
					this.uGridParentesco.Focus();
					return;
				}

				#endregion Parentesco

				#region Datos_Generales

				if (drReferencias.Cells["ApellidoPaterno"].Value.ToString().Trim().Length == 0)
				{
					MessageBox.Show(string.Format("Ingrese el Apellido Paterno de la Referencia Personal N. {0}",drReferencias.Index + 1),"Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
					this.utcSolicitud.Tabs[2].Selected = true;
					this.uGridParentesco.ActiveRow = drReferencias;
					this.uGridParentesco.ActiveRow.Selected = true;
					this.uGridParentesco.Focus();
					return;
				}

				//				if (drReferencias.Cells["PrimerNombre"].Value.ToString().Trim().Length == 0)
				//				{
				//					MessageBox.Show(string.Format("Ingrese el Primer Nombre de la Referencia Personal N. {0}",drReferencias.Index + 1),"Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				//					this.utcSolicitud.Tabs[2].Selected = true;
				//					this.uGridParentesco.ActiveRow = drReferencias;
				//					this.uGridParentesco.ActiveRow.Selected = true;
				//					this.uGridParentesco.Focus();
				//					return;
				//				}

				#endregion Datos_Generales

				#region Provincia

				int idProvincia = 0;
				if (drReferencias.Cells["idProvincia"].Value != System.DBNull.Value) idProvincia = (int)drReferencias.Cells["idProvincia"].Value;	
				if (idProvincia == 0)
				{
					MessageBox.Show(string.Format("Seleccione la Provincia de la Referencia Personal N. {0}",drReferencias.Index + 1),"Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
					this.utcSolicitud.Tabs[2].Selected = true;
					this.uGridParentesco.ActiveRow = drReferencias;
					this.uGridParentesco.ActiveRow.Selected = true;
					this.uGridParentesco.Focus();
					return;
				}

				#endregion Provincia

				#region Canton

				int idCanton = 0;
				if (drReferencias.Cells["idCanton"].Value != System.DBNull.Value) idCanton = (int)drReferencias.Cells["idCanton"].Value;	
				if (idCanton == 0)
				{
					MessageBox.Show(string.Format("Seleccione un Canton de la Referencia Personal N. {0}",drReferencias.Index + 1),"Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
					this.utcSolicitud.Tabs[2].Selected = true;
					this.uGridParentesco.ActiveRow = drReferencias;
					this.uGridParentesco.ActiveRow.Selected = true;
					this.uGridParentesco.Focus();
					return;
				}

				#endregion Canton

				#region Direccion

				//				if (drReferencias.Cells["Direccion"].Value.ToString().Trim().Length == 0)
				//				{
				//					MessageBox.Show(string.Format("Ingrese la Dirección de la Referencia Personal N. {0}",drReferencias.Index + 1),"Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				//					this.utcSolicitud.Tabs[2].Selected = true;
				//					this.uGridParentesco.ActiveRow = drReferencias;
				//					this.uGridParentesco.ActiveRow.Selected = true;
				//					this.uGridParentesco.Focus();
				//					return;
				//				}

				#endregion Direccion

				#region Telefonos Vacios
				//				if (drReferencias.Cells["Telefono"].Value.ToString().Trim().Length == 0 && drReferencias.Cells["Celular"].Value.ToString().Trim().Length == 0)
				//				{
				//					MessageBox.Show("Debe Ingresar el Telefono Convencional o el Telefono Celular de la Referencia Familiar Seleccionada", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				//					this.utcSolicitud.Tabs[2].Selected = true;
				//					this.uGridParentesco.ActiveRow = drReferencias;
				//					this.uGridParentesco.ActiveRow.Selected = true;
				//					this.uGridParentesco.Focus();
				//					return;
				//				}
				#endregion Telefonos Vacios

				#region Telefono
				//				if (drReferencias.Cells["Telefono"].Value.ToString().Trim().Length > 0)
				//				{
				//					if (drReferencias.Cells["Telefono"].Value.ToString().Trim().Length != 9)
				//					{
				//						MessageBox.Show("El Telefono Convencional debe tener 9 Digitos.\n\nLos dos primeros corresponden al Codigo de Provincia", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				//						this.utcSolicitud.Tabs[2].Selected = true;
				//						this.uGridParentesco.ActiveRow = drReferencias;
				//						this.uGridParentesco.ActiveRow.Selected = true;
				//						this.uGridParentesco.Focus();
				//						return;
				//					}
				//				}
				#endregion Telefono

				#region Celular
				if (drReferencias.Cells["Celular"].Value.ToString().Trim().Length > 0)
				{
					if (drReferencias.Cells["Celular"].Value.ToString().Trim().Length != 10)
					{
						MessageBox.Show("El Telefono Celular debe tener 10 digitos", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
						this.utcSolicitud.Tabs[2].Selected = true;
						this.uGridParentesco.ActiveRow = drReferencias;
						this.uGridParentesco.ActiveRow.Selected = true;
						this.uGridParentesco.Focus();
						return;
					}
				}
				#endregion Celular

			}
			
			#endregion Referencias
	
			#region Valida_Combo_Vacio_Conyuge
      
			int iDPCTipodocumento = 1;
			int iDPCNacionalidad = 54;
			int iDPCSexo = 0;
			int iDPCProvinciaN = 0;
			int iDPCCantonN = 0;
			int iDPCNivelEducacion = 0;
			int iDPCProfesion= 0;

			int iDPCSituacionL = 0;
			int iDPCActividadE = 0;
			int iDPCTActividadEmpresa = 0;


			if ( this.cmbCTipoRuc.ActiveRow != null) iDPCTipodocumento = (int) this.cmbCTipoRuc.Value;
			if ( this.cmbCNacionalidad.ActiveRow != null)  iDPCNacionalidad = (int) this.cmbCNacionalidad.Value;
			if ( this.cmbCSexo.ActiveRow != null)  iDPCSexo = (int) this.cmbCSexo.Value;
			if ( this.cmbCProvinciaNacimiento.ActiveRow != null) iDPCProvinciaN = (int) this.cmbCProvinciaNacimiento.Value;
			if ( this.cmbCCantonNacimiento.ActiveRow != null)  iDPCCantonN = (int) this.cmbCCantonNacimiento.Value;
			if ( this.cmbCNivel.ActiveRow != null) iDPCNivelEducacion = (int) this.cmbCNivel.Value;
			if ( this.cmbCProfesion.ActiveRow != null)   iDPCProfesion = (int) this.cmbCProfesion.Value;
			if ( this.cmbCSituacionLaboral.ActiveRow != null) iDPCSituacionL  = ( int) this.cmbCSituacionLaboral.Value;
			if ( this.cmbCActividadEconomia.ActiveRow != null) iDPCActividadE  = ( int) this.cmbCActividadEconomia.Value;


			#endregion Valida_Combo_Vacio_Conyuge

			#region Valida_Combo_Vacio_Negocio

			int iNegciotipo = 0;
			int iNEgocioActivdadEconomica = 0;
			int iNegocioProvincia = 0;
			int iNegocioCanton = 0;
			int iNegocioParroquia = 0;
			int iNegocioBarrio = 0;

			if ( this.cmbSCTipoNegocio.ActiveRow != null) iNegciotipo = (int) this.cmbSCTipoNegocio.Value;
			if ( this.cmbSubActvidadEconomica.ActiveRow != null) iNEgocioActivdadEconomica = (int) this.cmbSubActvidadEconomica.Value;
			if ( this.cmbSNProvincia.ActiveRow != null) iNegocioProvincia = (int) this.cmbSNProvincia.Value;
			if ( this.cmbSNCanton.ActiveRow != null) iNegocioCanton = (int) this.cmbSNCanton.Value;
			if ( this.cmbSNParroquia.ActiveRow != null) iNegocioParroquia = (int) this.cmbSNParroquia.Value;
			if ( this.cmbSNBarrio.ActiveRow != null) iNegocioBarrio = (int) this.cmbSNBarrio.Value;

			#endregion Valida_Combo_Vacio_Negocio

			#region Valida_Combo_Vacio_Dependiente

			int iDependienteTipoempresa = 0;
			int iDependienteTipocontrato = 0;
			int iDependienteTipoSueldo = 0;
			int iDependienteCargo = 0;
			int iDependienteProvincia = 0;
			int iDependienteCanton = 0;
			int iDependienteParroquia = 0;
			int iDependienteBarrio = 0;

			if ( this.cmbSDTipoEmpresa.ActiveRow != null) iDependienteTipoempresa = (int) this.cmbSDTipoEmpresa.Value;
			if ( this.cmbSDTipoContrato.ActiveRow != null) iDependienteTipocontrato = (int) this.cmbSDTipoContrato.Value;
			if ( this.cmbSDTipoSueldo.ActiveRow != null) iDependienteTipoSueldo = (int) this.cmbSDTipoSueldo.Value;
			if ( this.cmbSDCargo.ActiveRow != null) iDependienteCargo = (int) this.cmbSDCargo.Value;
			if ( this.cmbSDProvincia.ActiveRow != null) iDependienteProvincia = (int) this.cmbSDProvincia.Value;
			if ( this.cmbSDCanton.ActiveRow != null) iDependienteCanton = (int) this.cmbSDCanton.Value;
			if ( this.cmbSDParroquia.ActiveRow != null) iDependienteParroquia = (int) this.cmbSDParroquia.Value;
			if ( this.cmbSDBarrio.ActiveRow != null) iDependienteBarrio = (int) this.cmbSDBarrio.Value;

		
			#endregion Valida_Combo_Vacio_Dependiente

			#region Guarda_Datos
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

					#region Guarda Datos Garante

					string sSQLDetalle = string.Format("Exec Cre_GuardaGarantes {0}, {1}, '{2}', '{3}'",
						(int)this.txtIdCre_Garante.Value, 
						idCre_Solicitud, 
						Convert.ToDateTime(this.dtFecha.Value).ToString("yyyyMMdd"),
						"");
					oCmdActualiza.CommandText = sSQLDetalle;
					this.txtIdCre_Garante.Value = (int)oCmdActualiza.ExecuteScalar();
					#endregion Guarda Datos Garante

					#region Datos Personales Garante

					int Ingresos = Convert.ToInt32(this.txtSNIngresos.Value) + Convert.ToInt32(this.txtSDIngresos.Value);
					int Gastos = Convert.ToInt32(this.txtSNGastos.Value) + Convert.ToInt32(this.txtSDGastos.Value);
					int Activos = Convert.ToInt32(this.txtSNActivos.Value) + Convert.ToInt32(this.txtSDActivos.Value);
					int Pasivos = Convert.ToInt32(this.txtSNPasivos.Value) + Convert.ToInt32(this.txtSDPasivos.Value);

					string sSQLDatosGenerales = string.Format("Exec Cre_GuardaDatosGenerales {0}, {1}, {2}, {3}, '{4}', '{5}', '{6}', '{7}', '{8}', {9}, {10}, {11}, {12}, {13}, {14}, '{15}', {16}, {17},{18}, {19}, {20}, {21}, {22}, {23}, {24}, {25}, {26}, {27}, {28}, {29}, {30}, {31}, {32}, '{33}', '{34}', {35}, {36}, {37}",
						(int)this.txtClienteIdCre_DatosGenerales.Value, 
						2,
						(int)this.txtIdCre_Garante.Value, 
						this.cmbTipoRuc.Value,
						this.txtDocumento.Text, 
						this.txtApellidoP.Text, 
						this.txtApellidoM.Text,
						this.txtPnombre.Text.ToString(), 
						this.txtSnombre.Text.ToString(),
						this.cmbCalificacion.Value,  
						0, 
						(int)this.txtScore.Value, 
						0,
						this.cmbEstado.Value, 
						(int)this.cmbNENacionalidad.Value, 
						Convert.ToDateTime(this.dtpNEFechaNacimiento.Value).ToString("yyyyMMdd"),
						(int)this.cmbSexo.Value,   
						(int) this.cmbProvinciaNacimiento.Value,
						(int) this.cmbCantonNacimiento.Value,
						(int) this.cmbNEEstadoCivil.Value, 
						(int)this.txtNECargasFamiliares.Value, 
						(int) this.cmbNivel.Value, 
						(int) this.cmbProfesion.Value,
						0, 
						(int) this.cmbSituacionLaboral.Value, 
						(int) this.cmbActividadEcomica.Value,
						Ingresos, 
						Gastos, 
						Activos, 
						Pasivos,
						0,0,0,"","", 0,false,false);

					oCmdActualiza.CommandText = sSQLDatosGenerales;
					this.txtClienteIdCre_DatosGenerales.Value= (int)oCmdActualiza.ExecuteScalar();
	        
					#endregion Datos Personales Garante

					#region Guarda Datos Ubicacion Garante

					string sSQLDatosUbicacion = string.Format("Exec Cre_GuardaDatosUbicacion {0}, {1}, {2}, {3}, {4}, {5}, {6}, '{7}', '{8}', '{9}', '{10}', '{11}', '{12}', '{13}', '{14}', '{15}', '{16}', {17},'{18}', '{19}', '{20}', {21}, {22}, {23}, {24}",
						(int)this.txtClienteIdCre_DatosUbicacion.Value, 
						2, 
						(int)this.txtIdCre_Garante.Value,
						(int)	this.cmbProvinciaDomicilio.Value, 
						(int)this.cmbCantonDomicilio.Value, 
						(int)this.cmbParroquiaDomicilio.Value, 
						(int)this.cmbBarrioDomicilio.Value,
						this.txtCalle1.Text, 
						this.txtNumcasa.Text, 
						this.txtCalle2.Text, 
						this.txtReferencia.Text,
						this.txtNETelefono1.Text,
						"", 
						this.txtNETelefono2.Text.ToString(), 
						"", this.txtNECelular.Text.ToString(),
						this.txtNEEMail.Text,
						(int)this.cmbTipoVivienda.Value, 
						Convert.ToDateTime(this.dtpTiempoVivienda.Value).ToString("yyyyMMdd"),
						"", 
						"",
						0,
						0, 
						0, 
						0);     
					oCmdActualiza.CommandText = sSQLDatosUbicacion;
					this.txtClienteIdCre_DatosUbicacion.Value= (int)oCmdActualiza.ExecuteScalar();

					#endregion Guarda Datos Ubicacion Garante

					#region Conyugue Garante

					if ((int)this.cmbNEEstadoCivil.Value != 2 && (int)this.cmbNEEstadoCivil.Value != 3 && (int)this.cmbNEEstadoCivil.Value != 5)
					{
						string sSQLGuardaConyuge = string.Format("Exec Cre_GuardaConyugues {0}, {1}, {2}, '{3}', {4}, {5}",
							(int)this.txtIdCre_Conyugue.Value, 
							2, 
							(int)this.txtIdCre_Garante.Value,
							"", 
							0,
							0);   
						oCmdActualiza.CommandText = sSQLGuardaConyuge;
						this.txtIdCre_Conyugue.Value = (int)oCmdActualiza.ExecuteScalar();

						#region Datos Generales Conyugue Garante

						string sSQLDatosGeneralesconyuge = string.Format("Exec Cre_GuardaDatosGenerales {0}, {1}, {2}, {3}, '{4}', '{5}', '{6}', '{7}', '{8}', {9}, {10}, {11}, {12}, {13}, {14}, '{15}', {16}, {17},{18}, {19}, {20}, {21}, {22}, {23}, {24}, {25}, {26}, {27}, {28}, {29}, {30}, {31}, {32}, '{33}', '{34}', {35}, {36}, {37}",
							(int)this.txtConyugueIdCre_DatosGenerales.Value, 
							1, 
							(int)this.txtIdCre_Conyugue.Value,
							(int)this.cmbCTipoRuc.Value, 
							this.txtCDocumento.Text,
							this.txtCApellidoP.Text, 
							this.txtCApellidoM.Text,
							this.txtCPNombre.Text,
							this.txtCSNombre.Text, 
							0, 0, 0, 0,	0,
							(int)this.cmbCNacionalidad.Value, 
							Convert.ToDateTime(this.dpCFechaNacimiento.Value).ToString("yyyyMMdd"),
							(int)this.cmbCSexo.Value,
							(int)this.cmbCProvinciaNacimiento.Value, 
							(int)this.cmbCCantonNacimiento.Value,
							0, 0, 
							(int)this.cmbCNivel.Value, 
							(int)this.cmbCProfesion.Value,
							0, 
							(int)this.cmbCSituacionLaboral.Value, 
							(int)this.cmbCActividadEconomia.Value,
							0,0,0,0,0,0,0, "","",0,false,false);

						oCmdActualiza.CommandText = sSQLDatosGeneralesconyuge;
						this.txtConyugueIdCre_DatosGenerales.Value= (int)oCmdActualiza.ExecuteScalar();

						#endregion Datos Generales Conyugue
	
					}
					#endregion Conyuque Garante

					#region Referencias

					foreach(Infragistics.Win.UltraWinGrid.UltraGridRow drReferencias  in this.uGridParentesco.Rows.All)
					{	
						int idCre_ReferenciasClientes = 0;
						int Indice = 0;
						int idParentesco = 0;
						int idProvincia = 0;
						int idCanton = 0;

						if (drReferencias.Cells["idCre_ReferenciasClientes"].Value != System.DBNull.Value) idCre_ReferenciasClientes = (int)drReferencias.Cells["idCre_ReferenciasClientes"].Value;	
						if (drReferencias.Cells["Indice"].Value != System.DBNull.Value) Indice = Convert.ToInt32(drReferencias.Cells["Indice"].Value); else Indice =	drReferencias.Index;
						if (drReferencias.Cells["idParentesco"].Value != System.DBNull.Value) idParentesco = (int)drReferencias.Cells["idParentesco"].Value;
						if (drReferencias.Cells["idProvincia"].Value != System.DBNull.Value) idProvincia = (int)drReferencias.Cells["idProvincia"].Value;	
						if (drReferencias.Cells["idCanton"].Value != System.DBNull.Value) idCanton  = (int)drReferencias.Cells["idCanton"].Value;

						string sSQLReferencias = string.Format("Exec Cre_GuardaReferenciasClientes {0}, {1}, {2}, {3}, '{4}', '{5}', '{6}', '{7}', {8}, {9}, '{10}', '{11}', '{12}', '{13}', '{14}', '{15}', '{16}'",
							(int)drReferencias.Cells["idCre_ReferenciasClientes"].Value, 
							(int)this.txtClienteIdCre_DatosGenerales.Value, 
							Indice,
							(int)drReferencias.Cells["idParentesco"].Value,
							drReferencias.Cells["ApellidoPaterno"].Value.ToString(),
							drReferencias.Cells["ApellidoMaterno"].Value.ToString(), 
							drReferencias.Cells["PrimerNombre"].Value.ToString(),
							drReferencias.Cells["SegundoNombre"].Value.ToString(),
							(int)drReferencias.Cells["idProvincia"].Value,
							(int)drReferencias.Cells["idCanton"].Value,
							drReferencias.Cells["Direccion"].Value.ToString(), 
							drReferencias.Cells["Telefono"].Value.ToString(),
							drReferencias.Cells["Celular"].Value.ToString(),
							drReferencias.Cells["EmpresaTrabajo"].Value.ToString(),
							drReferencias.Cells["DireccionTrabajo"].Value.ToString(),
							drReferencias.Cells["TelefonoTrabajo"].Value.ToString(),
							drReferencias.Cells["ExtensionTrabajo"].Value.ToString());
						oCmdActualiza.CommandText = sSQLReferencias;
						oCmdActualiza.ExecuteNonQuery();
					}

					#endregion Referencias

					#region Negocios 

					int iSNEmpleados = 0;
					int iSCTipoNegocio = 0;
					int iSubActvidadEconomica = 0;


					if ((ientidad == 2 && (int)this.cmbSituacionLaboral.Value > 10) || ( ientidad != 2 && ((int)this.cmbSituacionLaboral.Value != 1 && (int) this.cmbSituacionLaboral.Value != 4)))
					{

						string sSQLGuardaNegocio = string.Format("Exec Cre_GuardaNegocios {0}, {1}, '{2}', '{3}', '{4}', {5}, {6}, {7}, {8}, {9}, {10}",
							(int)this.txtClienteIdCre_Negocio.Value, 
							(int)this.txtClienteIdCre_DatosGenerales.Value,
							this.txtSNRuc.Text.ToString(),
							this.txtSNNombre.Text.ToString(), 
							Convert.ToDateTime(this.dtSNFechaInicio.Value).ToString("yyyyMMdd"),
							iSCTipoNegocio, 
							iSubActvidadEconomica,
							(int)this.txtSNMetros.Value, 
							iSNEmpleados, 
							(bool)this.chkSNContabilidad.Checked,
							(bool)this.chkSNAfiliacion.Checked);   
						oCmdActualiza.CommandText = sSQLGuardaNegocio;
						this.txtClienteIdCre_Negocio.Value  = (int)oCmdActualiza.ExecuteScalar();


						string sSQLDatosUbicacionNegocio = string.Format("Exec Cre_GuardaDatosUbicacion {0}, {1}, {2}, {3}, {4}, {5}, {6}, '{7}', '{8}', '{9}', '{10}', '{11}', '{12}', '{13}', '{14}', '{15}', '{16}', {17},'{18}', '{19}', '{20}', {21}, {22}, {23}, {24}",
							(int)this.txtNegocioIdCre_DatosUbicacion.Value,
							3,
							(int)this.txtClienteIdCre_Negocio.Value,
							(int) this.cmbSNProvincia.Value, 
							(int) this.cmbSNCanton.Value,
							(int) this.cmbSNParroquia.Value,
							(int) 	this.cmbSNBarrio.Value,
							this.txtSNCalle1.Text.ToString(), 
							this.txtSNNumCasa.Text.ToString(),
							this.txtSNCalle2.Text.ToString(), 
							this.txtSNReferencia.Text.ToString(),
							this.txtSNTelefono1.Text.ToString(), 
							"", this.txtSNTelefono2.Text.ToString(), 
							"",  
							this.txtSNCelular.Text.ToString(),
							this.txtSNEmail.Text.ToString(),
							0,
							Convert.ToDateTime(DateTime.Today).ToString("yyyyMMdd"), 
							"", "", 
							(int)this.txtSNIngresos.Value,
							(int)this.txtSNGastos.Value, 
							(int)this.txtSNActivos.Value, 
							(int)this.txtSNPasivos.Value);     
						oCmdActualiza.CommandText = sSQLDatosUbicacionNegocio;
						this.txtNegocioIdCre_DatosUbicacion.Value= (int)oCmdActualiza.ExecuteScalar();
						
					}
					#endregion Negocios

					#region Empresa de Trabajo
					if ((ientidad == 2 && Convert.ToInt32(this.cmbSituacionLaboral.Value) == 10) || 
						(ientidad  != 2 && (Convert.ToInt32(this.cmbSituacionLaboral.Value) == 1 || 
						Convert.ToInt32(this.cmbSituacionLaboral.Value) == 3)))
					
					{
						string sSQLGuardaEmpresaTrabajo = string.Format("Exec Cre_GuardaEmpresaTrabajo {0}, {1}, '{2}', '{3}', {4}, '{5}', {6}, {7}, {8}, {9}, {10}, '{11}', {12}, {13}, {14}",
							(int)this.txtClienteIdCre_EmpresasTrabajo.Value, 
							(int)this.txtClienteIdCre_DatosGenerales.Value,
							this.txtSDRuc.Text, 
							this.txtSDEmpresa.Text, 
							(int)this.cmbSDTipoEmpresa.Value, 
							Convert.ToDateTime(this.dtSDFecha.Value).ToString("yyyyMMdd"),
							0,
							(int)this.cmbSDTipoContrato.Value, 
							(int)this.cmbSDTipoSueldo.Value, 
							this.chkSDAfiliacion.Checked, 
							(int)this.txtSDPago.Value,
							this.txtSDDepartamento.Text, 
							(int)this.cmbSDCargo.Value, 
							(int)this.txtSDPago.Value, 
							0);   
						oCmdActualiza.CommandText = sSQLGuardaEmpresaTrabajo;
						this.txtClienteIdCre_EmpresasTrabajo.Value  = (int)oCmdActualiza.ExecuteScalar();


						string sSQLDatosUbicacionUbicacion = string.Format("Exec Cre_GuardaDatosUbicacion {0}, {1}, {2}, {3}, {4}, {5}, {6}, '{7}', '{8}', '{9}', '{10}', '{11}', '{12}', '{13}', '{14}', '{15}', '{16}', {17},'{18}', '{19}', '{20}', {21}, {22}, {23}, {24}",
							(int)this.txtEmpresaIdCre_DatosUbicacion.Value, 
							4,
							(int)this.txtClienteIdCre_EmpresasTrabajo.Value,
							(int)this.cmbSDProvincia.Value, 
							(int)this.cmbSDCanton.Value,
							(int)this.cmbSDParroquia.Value, 
							(int)this.cmbSDBarrio.Value,

							this.txtSDCalle1.Text.ToString(), 
							this.txtSDNumCasa.Text.ToString(), 
							this.txtSDCalle2.Text.ToString(), 
							this.txtSDReferencia.Text.ToString(),
							this.txtSDTelefono.Text.ToString(), 
							this.txtSDExt.Text.ToString(), 
							"", "", 
							this.txtSDCelular.Text.ToString(), 
							this.txtSDEmail.Text.ToString(), 
							0, Convert.ToDateTime(DateTime.Today).ToString("yyyyMMdd"), "", "",
							(int)this.txtSDIngresos.Value, 
							(int)this.txtSDGastos.Value, 
							(int)this.txtSDActivos.Value, 
							(int)this.txtSDPasivos.Value);     
						oCmdActualiza.CommandText = sSQLDatosUbicacionUbicacion;
						this.txtEmpresaIdCre_DatosUbicacion.Value  = (int)oCmdActualiza.ExecuteScalar();
					}
					#endregion Empresa de Trabajo

					oTransaction.Commit();
					
				}
				catch (Exception ex)
				{
					oTransaction.Rollback();

					MessageBox.Show(string.Format("Commit Exception Type : {0} {1}", ex.GetType(), ex.Message), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Error);	
					return;
				}
				finally
				{
					oConexion.Close();
				}
			}

			#endregion Guarda_Datos

			MessageBox.Show("Registro Guardado Correctamente ", "Informacion", MessageBoxButtons.OK, MessageBoxIcon.Information);  

			this.DialogResult = DialogResult.OK;
		}

		private void txtDocumento_ValueChanged(object sender, System.EventArgs e)
		{
//			LimpiaControles();
		}

		private void uGridParentesco_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			try
			{
				FuncionesProcedimientos.DesplazamientoGrids(sender, e, this.uGridParentesco);
			}
			catch(Exception ex){


				if (this.cmbSituacionLaboral.ActiveRow != null)		
				{
	
					if ((ientidad == 2 && (int)this.cmbSituacionLaboral.Value > 10) || (ientidad != 2 && ((int)this.cmbSituacionLaboral.Value != 1 && (int) this.cmbSituacionLaboral.Value != 4)))
					{
//						if (!Validacion.vbTexto(this.txtSNRuc, 3, 13, "Ingrese el Ruc del Negocio")) 	
//						{ 
//							this.utcSolicitud.Tabs[3].Selected = true;
//							this.txtSNRuc.Focus();
//							return;
//						}
					}
				}
			}
			
		}

		private void uGridParentesco_AfterRowInsert(object sender, Infragistics.Win.UltraWinGrid.RowEventArgs e)
		{

				e.Row.Cells["idCre_ReferenciasClientes"].Value = 0;
				e.Row.Cells["idCre_DatosGenerales"].Value = 0;
		}

		private void uGridParentesco_BeforeRowsDeleted(object sender, Infragistics.Win.UltraWinGrid.BeforeRowsDeletedEventArgs e)
		{

		}

		private void uGridParentesco_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
			try
			{
				if (this.uGridParentesco.ActiveCell.Column.Key.ToString().ToUpper() == "TELEFONO" || this.uGridParentesco.ActiveCell.Column.Key.ToString().ToUpper() == "CELULAR")
					Funcion.CamposNumericos(sender, e);
			}
			catch(Exception ex){

				
				if (this.cmbSituacionLaboral.ActiveRow != null)		
				{
	
	
					if ((ientidad == 2 && (int)this.cmbSituacionLaboral.Value > 10) || (ientidad != 2 && ((int)this.cmbSituacionLaboral.Value != 1 && (int) this.cmbSituacionLaboral.Value != 4)))
					{
//						if (!Validacion.vbTexto(this.txtSNRuc, 3, 13, "Ingrese el Ruc del Negocio")) 	
//						{ 
//							this.utcSolicitud.Tabs[3].Selected = true;
//							this.txtSNRuc.Focus();
//							return;
//						}
					}
				}
			
			}
			
		}

		private void txtNETelefono1_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
			validar.Solonumeros(e);
		}

		private void txtNECelular_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
			validar.Solonumeros(e);
		}

		private void txtSNTelefono1_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
			validar.Solonumeros(e);
		}

		private void txtSNCelular_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
			validar.Solonumeros(e);
		}

		private void txtSDTelefono_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
			validar.Solonumeros(e);
		}

		private void txtSDCelular_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
			validar.Solonumeros(e);
		}

		private void uGridParentesco_ClickCellButton(object sender, Infragistics.Win.UltraWinGrid.CellEventArgs e)
		{
			if (e.Cell.Column.ToString() == "btnTelefono")
			{
				#region Validacion Telefono Referencia 
							
				if(e.Cell.Row.Cells["Telefono"].Value.ToString().Length==0)
				{
					MessageBox.Show("Ingrese el Telefono de la Referencia", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
					this.utcSolicitud.Tabs[2].Selected = true;
					this.uGridParentesco.ActiveRow.Selected = true;
					return;
				}

				if (e.Cell.Row.Cells["Telefono"].Value.ToString().Trim().Length > 0)
				{
					if (e.Cell.Row.Cells["Telefono"].Value.ToString().Trim().Length != 9)
					{
						MessageBox.Show("El Telefono Convencional debe tener 9 Digitos.\n\nLos dos primeros corresponden al Codigo de Provincia", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
						this.utcSolicitud.Tabs[2].Selected = true;
						this.uGridParentesco.ActiveRow.Selected = true;
						return;
					}
				}
 
				int iNuevoCel =0;
				iNuevoCel = Funcion.iEscalarSQL(cdsSeteoF, string.Format("select count(*) from Cre_ReferenciasClientes where idCre_DatosGenerales =  {0} and Telefono like '%"+"{1}"+"%' ", (int)this.txtClienteIdCre_DatosGenerales.Value, e.Cell.Row.Cells["Telefono"].Value.ToString()));

				if((int)e.Cell.Row.Cells["idCre_ReferenciasClientes"].Value > 0)
				{
					if (iNuevoCel > 0)
					{
						AbreVerificaciontelefonica (false, (int)this.txtIdCre_Garante.Value, 8,
							e.Cell.Row.Cells["ApellidoPaterno"].Value.ToString() + " " + e.Cell.Row.Cells["ApellidoMaterno"].Value.ToString() + " " 
							+ e.Cell.Row.Cells["PrimerNombre"].Value.ToString()+ " " + e.Cell.Row.Cells["SegundoNombre"].Value.ToString(),
							e.Cell.Row.Cells["Telefono"].Value.ToString(), 0); 
					}
					else
					{
						MessageBox.Show(string.Format("Guarde los datos del Garante para realizar la verificacion telefonica"), "Point Technology", MessageBoxButtons.OK,MessageBoxIcon.Stop);
						this.utcSolicitud.Tabs[2].Selected = true;
						this.uGridParentesco.ActiveRow.Selected = true;
						return;
					}

				}
				else
				{
					MessageBox.Show(string.Format("Guarde los datos del Garante para realizar la verificacion telefonica"), "Point Technology", MessageBoxButtons.OK,MessageBoxIcon.Stop);
					this.utcSolicitud.Tabs[2].Selected = true;
					this.uGridParentesco.ActiveRow.Selected = true;
					return;
				}
				#endregion Validacion Telefono Referencia 
			}

			if (e.Cell.Column.ToString() == "btnCelular")
			{			
					#region Celular
				
				if(e.Cell.Row.Cells["Celular"].Value.ToString().Length==0)
				{
					MessageBox.Show("Ingrese el Celular de la Referencia", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
					this.utcSolicitud.Tabs[2].Selected = true;
					this.uGridParentesco.ActiveRow.Selected = true;
					return;
				}

			
				if (e.Cell.Row.Cells["Celular"].Value.ToString().Trim().Length > 0)
				{
					if (e.Cell.Row.Cells["Celular"].ToString().Trim().Length < 10 )
					{
						MessageBox.Show("El Telefono Celular debe tener 10 digitos", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
						this.utcSolicitud.Tabs[2].Selected = true;
						this.uGridParentesco.ActiveRow.Selected = true;
						return;
					}
	
					if (e.Cell.Row.Cells["Celular"].Value.ToString().Trim().Length != 10)
					{
						MessageBox.Show("El Telefono Celular debe tener 10 digitos", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
						this.utcSolicitud.Tabs[2].Selected = true;
						this.uGridParentesco.ActiveRow.Selected = true;
						return;
					}
				}
			
				int iNuevoCel =0;
				iNuevoCel = Funcion.iEscalarSQL(cdsSeteoF, string.Format("select count(*) from Cre_ReferenciasClientes where idCre_DatosGenerales =  {0} and Celular like '%"+"{1}"+"%' ", (int)this.txtClienteIdCre_DatosGenerales.Value, e.Cell.Row.Cells["Celular"].Value.ToString()));

				if((int)e.Cell.Row.Cells["idCre_ReferenciasClientes"].Value > 0)
				{
					if (iNuevoCel > 0)
					{
						AbreVerificaciontelefonica (false, (int)this.txtIdCre_Garante.Value, 9,
							e.Cell.Row.Cells["ApellidoPaterno"].Value.ToString() + " " + e.Cell.Row.Cells["ApellidoMaterno"].Value.ToString() + " " 
							+ e.Cell.Row.Cells["PrimerNombre"].Value.ToString()+ " " + e.Cell.Row.Cells["SegundoNombre"].Value.ToString(),
							e.Cell.Row.Cells["Celular"].Value.ToString(), 0); 
					}
					else
					{
						MessageBox.Show(string.Format("Guarde los datos del Garante para realizar la verificacion telefonica"), "Point Technology", MessageBoxButtons.OK,MessageBoxIcon.Stop);
						this.utcSolicitud.Tabs[2].Selected = true;
						this.uGridParentesco.ActiveRow.Selected = true;
						return;
					}

				}
				else
				{
					MessageBox.Show(string.Format("Guarde los datos del Garante para realizar la verificacion telefonica"), "Point Technology", MessageBoxButtons.OK,MessageBoxIcon.Stop);
					this.utcSolicitud.Tabs[2].Selected = true;
					this.uGridParentesco.ActiveRow.Selected = true;
					return;
				}
					#endregion Celular
			}
			if (e.Cell.Column.ToString() == "btnTelTrabajo")
			{			
				#region Validacion Telefono Referencia 
							
				if(e.Cell.Row.Cells["TelefonoTrabajo"].Value.ToString().Length==0)
				{
					MessageBox.Show("Ingrese el Telefono de la Referencia del Trabajo", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
					this.utcSolicitud.Tabs[2].Selected = true;
					this.uGridParentesco.ActiveRow.Selected = true;
					return;
				}

				if (e.Cell.Row.Cells["TelefonoTrabajo"].Value.ToString().Trim().Length > 0)
				{
					if (e.Cell.Row.Cells["TelefonoTrabajo"].Value.ToString().Trim().Length != 9)
					{
						MessageBox.Show("El Telefono Convencional debe tener 9 Digitos.\n\nLos dos primeros corresponden al Codigo de Provincia", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
						this.utcSolicitud.Tabs[2].Selected = true;
						this.uGridParentesco.ActiveRow.Selected = true;
						return;
					}
				}
 
				int iNuevoCel =0;
				iNuevoCel = Funcion.iEscalarSQL(cdsSeteoF, string.Format("select count(*) from Cre_ReferenciasClientes where idCre_DatosGenerales =  {0} and TelefonoTrabajo like '%"+"{1}"+"%' ", (int)this.txtClienteIdCre_DatosGenerales.Value, e.Cell.Row.Cells["TelefonoTrabajo"].Value.ToString()));

				if((int)e.Cell.Row.Cells["idCre_ReferenciasClientes"].Value > 0)
				{
					if (iNuevoCel > 0)
					{
						AbreVerificaciontelefonica (false, (int)this.txtIdCre_Garante.Value, 10,
							e.Cell.Row.Cells["ApellidoPaterno"].Value.ToString() + " " + e.Cell.Row.Cells["ApellidoMaterno"].Value.ToString() + " " 
							+ e.Cell.Row.Cells["PrimerNombre"].Value.ToString()+ " " + e.Cell.Row.Cells["SegundoNombre"].Value.ToString(),
							e.Cell.Row.Cells["TelefonoTrabajo"].Value.ToString(), 0); 
					}
					else
					{
						MessageBox.Show(string.Format("Guarde los datos del Garante para realizar la verificacion telefonica"), "Point Technology", MessageBoxButtons.OK,MessageBoxIcon.Stop);
						this.utcSolicitud.Tabs[2].Selected = true;
						this.uGridParentesco.ActiveRow.Selected = true;
						return;
					}

				}
				else
				{
					MessageBox.Show(string.Format("Guarde los datos del cliente para realizar la verificacion telefonica"), "Point Technology", MessageBoxButtons.OK,MessageBoxIcon.Stop);
					this.utcSolicitud.Tabs[2].Selected = true;
					this.uGridParentesco.ActiveRow.Selected = true;
					return;
				}
				#endregion Validacion Telefono Referencia 
			}
		}

		private void btnTelf2_Click(object sender, System.EventArgs e)
		{

			if (this.txtNETelefono2.Text.ToString().Trim().Length == 0)
			{
				MessageBox.Show(string.Format("Debe Ingresar un Numero de Teléfono"), "Point Technology", MessageBoxButtons.OK,MessageBoxIcon.Stop);
				this.txtNETelefono2.Focus();
				return;	
			}
				if(!Validacion.vbValidaTelefono(this.txtNETelefono2, 1, false)) return;

				int iNuevoCel =0;
				int iCelExistente = 0;
			
				iNuevoCel = Funcion.iEscalarSQL(cdsSeteoF, string.Format("select count(*) from Cre_DatosUbicacion where Origen = 2 and idOrigen = {0} and Telefono2 like '%"+"{1}"+"%' ", (int)this.txtIdCre_Garante.Value, this.txtNETelefono2.Text.ToString()));

				if((int)this.txtIdCre_Solicitud.Value > 0)
				{
					if (iNuevoCel > 0)
					{
						AbreVerificaciontelefonica (false, (int)this.txtIdCre_Garante.Value, 1,
							this.txtApellidoP.Text + " " + this.txtApellidoM.Text + " " + this.txtPnombre.Text + " " + this.txtSnombre.Text,
							this.txtNETelefono2.Text, 0); 
					}
					else
					{
						MessageBox.Show(string.Format("Guarde los datos del Garante para realizar la verificacion telefonica"), "Point Technology", MessageBoxButtons.OK,MessageBoxIcon.Stop);
						this.txtNETelefono2.Focus();
						return;	
					}

				}
				else
				{
					MessageBox.Show(string.Format("Guarde los datos del Garante para realizar la verificacion telefonica"), "Point Technology", MessageBoxButtons.OK,MessageBoxIcon.Stop);
					this.txtNETelefono1.Focus();
					return;	
				}
		}

		private void btnCelular_Click(object sender, System.EventArgs e)
		{

			if (this.txtNECelular.Text.ToString().Trim().Length == 0  )
			{
				MessageBox.Show(string.Format("Ingrese un Número Celular "), "Point Technology", MessageBoxButtons.OK,MessageBoxIcon.Stop);
				this.txtNECelular.Focus();
				return;
			}

			if (!Validacion.vbValidaTelefono(this.txtNECelular, 2, false)) return;
			int iNuevoCel =0;
			int iCelExistente = 0;
			
			iNuevoCel = Funcion.iEscalarSQL(cdsSeteoF, string.Format("select count(*) from Cre_DatosUbicacion where Origen = 2 and idOrigen = {0} and Celular like '%"+"{1}"+"%' ", (int)this.txtIdCre_Garante.Value, this.txtNECelular.Text.ToString()));

			if((int)this.txtIdCre_Solicitud.Value  > 0)
			{
				if (iNuevoCel > 0)
				{
					AbreVerificaciontelefonica (false, (int)this.txtIdCre_Garante.Value, 2,
						this.txtApellidoP.Text + " " + this.txtApellidoM.Text + " " + this.txtPnombre.Text + " " + this.txtSnombre.Text,
						this.txtNECelular.Text, 0); 
				}
				else
				{
					MessageBox.Show(string.Format("Guarde los datos del Garante para realizar la verificacion telefonica"), "Point Technology", MessageBoxButtons.OK,MessageBoxIcon.Stop);
					this.txtNECelular.Focus();
					return;	
				}
			}
			else
			{
				MessageBox.Show(string.Format("Guarde los datos del Garante para realizar la verificacion telefonica"), "Point Technology", MessageBoxButtons.OK,MessageBoxIcon.Stop);
				this.txtNECelular.Focus();
				return;	
			}
		
		}

		private void btnCTelf_Click(object sender, System.EventArgs e)
		{
			if (this.txtSNTelefono1.Text.ToString().Trim().Length == 0  )
			{
				MessageBox.Show(string.Format("Ingrese un Número de Teléfono "), "Point Technology", MessageBoxButtons.OK,MessageBoxIcon.Stop);
				this.txtSNTelefono1.Focus();
				return;	
			}
			if (!Validacion.vbValidaTelefono(this.txtSNTelefono1, 1, false)) return;
			int iNuevoCel =0;
			iNuevoCel = Funcion.iEscalarSQL(cdsSeteoF, string.Format("select count(*) from Cre_DatosUbicacion where Origen = 3 and idOrigen = {0} and Telefono1 like '%"+"{1}"+"%' ", (int)this.txtClienteIdCre_Negocio.Value, this.txtSNTelefono1.Text.ToString()));

			if((int)this.txtIdCre_Solicitud.Value  > 0)
			{
				if (iNuevoCel > 0)
				{
					AbreVerificaciontelefonica (false, (int)this.txtIdCre_Garante.Value, 5,
						this.txtSNNombre.Text,this.txtSNTelefono1.Text, 0); 
				}
				else
				{
					MessageBox.Show(string.Format("Guarde los datos del Garante para realizar la verificacion telefonica"), "Point Technology", MessageBoxButtons.OK,MessageBoxIcon.Stop);
					this.txtSNTelefono1.Focus();
					return;	
				}

			}
			else
			{
				MessageBox.Show(string.Format("Guarde los datos del Garante para realizar la verificacion telefonica"), "Point Technology", MessageBoxButtons.OK,MessageBoxIcon.Stop);
				this.txtSNTelefono1.Focus();
				return;	
			}
		}

		private void btnCTelf2_Click(object sender, System.EventArgs e)
		{
			if (this.txtSNTelefono2.Text.ToString().Trim().Length == 0  )
			{
				MessageBox.Show(string.Format("Ingrese un Número Celular "), "Point Technology", MessageBoxButtons.OK,MessageBoxIcon.Stop);
				this.txtSNTelefono2.Focus();
				return;	 
			}
			if (!Validacion.vbValidaTelefono(this.txtSNTelefono2, 1, false)) return;
			int iNuevoCel =0;
			iNuevoCel = Funcion.iEscalarSQL(cdsSeteoF, string.Format("select count(*) from Cre_DatosUbicacion where Origen = 3 and idOrigen = {0} and Telefono1 like '%"+"{1}"+"%' ", (int)this.txtClienteIdCre_Negocio.Value, this.txtSNTelefono2.Text.ToString()));

			if((int)this.txtIdCre_Solicitud.Value  > 0)
			{
				if (iNuevoCel > 0)
				{
					AbreVerificaciontelefonica (false, (int)this.txtIdCre_Garante.Value, 6,
						this.txtSNNombre.Text,this.txtSNTelefono2.Text, 0); 
				}
				else
				{
					MessageBox.Show(string.Format("Guarde los datos del Garante para realizar la verificacion telefonica"), "Point Technology", MessageBoxButtons.OK,MessageBoxIcon.Stop);
					this.txtSNTelefono2.Focus();
					return;	
				}
			}
			else
			{
				MessageBox.Show(string.Format("Guarde los datos del Garante para realizar la verificacion telefonica"), "Point Technology", MessageBoxButtons.OK,MessageBoxIcon.Stop);
				this.txtSNTelefono2.Focus();
				return;	
			}

		}

		private void btnCCelular_Click(object sender, System.EventArgs e)
		{
			if (this.txtSNCelular.Text.ToString().Trim().Length == 0  )
			{
				MessageBox.Show(string.Format("Ingrese un Número Celular "), "Point Technology", MessageBoxButtons.OK,MessageBoxIcon.Stop);
				this.txtSNCelular.Focus();
				return;	  
			}

			if (!Validacion.vbValidaTelefono(this.txtSNCelular, 2, false)) return;
			int iNuevoCel =0;
			iNuevoCel = Funcion.iEscalarSQL(cdsSeteoF, string.Format("select count(*) from Cre_DatosUbicacion where Origen = 3 and idOrigen = {0} and Celular like '%"+"{1}"+"%' ", (int)this.txtClienteIdCre_Negocio.Value, this.txtSNCelular.Text.ToString()));

			if((int)this.txtIdCre_Solicitud.Value  > 0)
			{
				if (iNuevoCel > 0)
				{
					AbreVerificaciontelefonica (false, (int)this.txtIdCre_Garante.Value, 11,
						this.txtSNNombre.Text, this.txtSNCelular.Text, 0); 
				}
				else
				{
					MessageBox.Show(string.Format("Guarde los datos del Garante para realizar la verificacion telefonica"), "Point Technology", MessageBoxButtons.OK,MessageBoxIcon.Stop);
					this.txtSNCelular.Focus();
					return;	
				}
			}
			else
			{
				MessageBox.Show(string.Format("Guarde los datos del Garante para realizar la verificacion telefonica"), "Point Technology", MessageBoxButtons.OK,MessageBoxIcon.Stop);
				this.txtSNCelular.Focus();
				return;	
			}
		}

		private void bntSDTelf_Click(object sender, System.EventArgs e)
		{
			if (this.txtSDTelefono.Text.ToString().Trim().Length == 0  )
			{
				MessageBox.Show(string.Format("Ingrese un Número Celular "), "Point Technology", MessageBoxButtons.OK,MessageBoxIcon.Stop);
				this.txtSDTelefono.Focus();
				return;	
        
			}
			if (!Validacion.vbValidaTelefono(this.txtSDTelefono, 1, false)) return;
			int iNuevoCel =0;
			iNuevoCel = Funcion.iEscalarSQL(cdsSeteoF, string.Format("select count(*) from Cre_DatosUbicacion where Origen = 4 and idOrigen = {0} and Telefono1 like '%"+"{1}"+"%' ", (int)this.txtClienteIdCre_EmpresasTrabajo.Value, this.txtSDTelefono.Text.ToString()));

			if((int)this.txtIdCre_Solicitud.Value  > 0)
			{
				if (iNuevoCel > 0)
				{
					AbreVerificaciontelefonica (false, (int)this.txtIdCre_Garante.Value, 7,
						this.txtSDEmpresa.Text,this.txtSDTelefono.Text, 0); 
				}
				else
				{
					MessageBox.Show(string.Format("Guarde los datos del Garante para realizar la verificacion telefonica"), "Point Technology", MessageBoxButtons.OK,MessageBoxIcon.Stop);
					this.txtSDTelefono.Focus();
					return;	
				}
			}
			else
			{
				MessageBox.Show(string.Format("Guarde los datos del Garante para realizar la verificacion telefonica"), "Point Technology", MessageBoxButtons.OK,MessageBoxIcon.Stop);
				this.txtSDTelefono.Focus();
				return;	
			}
		}

		private void btnSDCelular_Click(object sender, System.EventArgs e)
		{
			if (this.txtSDCelular.Text.ToString().Trim().Length == 0  )
			{
				MessageBox.Show(string.Format("Ingrese un Número Celular "), "Point Technology", MessageBoxButtons.OK,MessageBoxIcon.Stop);
				this.txtSDCelular.Focus();
				return;	
        
			}
			if (!Validacion.vbValidaTelefono(this.txtSNCelular, 2, false)) return;
			int iNuevoCel =0;
			iNuevoCel = Funcion.iEscalarSQL(cdsSeteoF, string.Format("select count(*) from Cre_DatosUbicacion where Origen = 4 and idOrigen = {0} and Celular like '%"+"{1}"+"%' ", (int)this.txtClienteIdCre_EmpresasTrabajo.Value, this.txtSDCelular.Text.ToString()));

			if((int)this.txtIdCre_Solicitud.Value  > 0)
			{
				if (iNuevoCel > 0)
				{
					AbreVerificaciontelefonica (false, (int)this.txtIdCre_Garante.Value, 12,
						this.txtSDEmpresa.Text, this.txtSDCelular.Text, 0); 
				}
				else
				{
					MessageBox.Show(string.Format("Guarde los datos del Garante para realizar la verificacion telefonica"), "Point Technology", MessageBoxButtons.OK,MessageBoxIcon.Stop);
					this.txtSDCelular.Focus();
					return;	
				}
			}
			else
			{
				MessageBox.Show(string.Format("Guarde los datos del Garante para realizar la verificacion telefonica"), "Point Technology", MessageBoxButtons.OK,MessageBoxIcon.Stop);
				this.txtSDCelular.Focus();
				return;	
			}
		}


		private void cmbCSituacionLaboral_ValueChanged(object sender, System.EventArgs e)
		{
			if (ientidad > 0)
			{
				SituacionLaboral1(ientidad);
			}
		}

		private void cmbActividadEcomica_KeyDown_1(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if(e.KeyCode == Keys.Enter) this.cmbProvinciaDomicilio.Focus();		
		}

		private void txtApellidoP_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
		validar.SoloLetras(e);
		}

		private void txtApellidoM_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
			validar.SoloLetras(e);
		}

		private void txtPnombre_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
			validar.SoloLetras(e);
		}

		private void txtSnombre_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
			validar.SoloLetras(e);
		}

		private void txtCalle1_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
			validar.CaracteresEspeciales(e);
		}

		private void txtNumcasa_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
//			validar.CaracteresEspeciales(e);
		}

		private void txtCalle2_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
			validar.CaracteresEspeciales(e);
		}

		private void txtCApellidoP_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
		validar.SoloLetras(e);
		}

		private void txtCApellidoM_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
			validar.SoloLetras(e);
		}

		private void txtCPNombre_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
			validar.SoloLetras(e);
		}

		private void txtCSNombre_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
			validar.SoloLetras(e);
		}

		private void txtSNNombre_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
			validar.CaracteresEspeciales(e);
		}

		private void txtSNCalle1_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
			validar.CaracteresEspeciales(e);
		}

		private void txtSDCalle1_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
			validar.CaracteresEspeciales(e);
		}

		private void txtSNNumCasa_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
//			validar.CaracteresEspeciales(e);
		}

		private void txtSNCalle2_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
			validar.CaracteresEspeciales(e);
		}

		private void txtSNReferencia_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
			validar.CaracteresEspeciales(e);
		}

		private void txtSDNumCasa_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
//			validar.CaracteresEspeciales(e);
		}

		private void txtSDCalle2_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
			validar.CaracteresEspeciales(e);
		}

		private void txtSDReferencia_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
			validar.CaracteresEspeciales(e);
		}

		private void txtNECargasFamiliares_Enter(object sender, System.EventArgs e)
		{
		this.txtNECargasFamiliares.SelectAll();
		}

		private void txtNECargasFamiliares_Click(object sender, System.EventArgs e)
		{
			this.txtNECargasFamiliares.SelectAll();
		}

		private void txtSNMetros_Click(object sender, System.EventArgs e)
		{
			this.txtSNMetros.SelectAll();
		}

		private void txtSNMetros_Enter(object sender, System.EventArgs e)
		{
		this.txtSNMetros.SelectAll();
		}

		private void txtSNEmpleados_Enter(object sender, System.EventArgs e)
		{
			this.txtSNEmpleados.SelectAll();
		}

		private void txtSNEmpleados_Click(object sender, System.EventArgs e)
		{
		this.txtSNEmpleados.SelectAll();
		}

		private void txtSNIngresos_Click(object sender, System.EventArgs e)
		{
			this.txtSNIngresos.SelectAll();
		}

		private void txtSNIngresos_Enter(object sender, System.EventArgs e)
		{
			this.txtSNIngresos.SelectAll();
		}

		private void txtSNGastos_Enter(object sender, System.EventArgs e)
		{
			this.txtSNGastos.SelectAll();
		}

		private void txtSNGastos_Click(object sender, System.EventArgs e)
		{
			this.txtSNGastos.SelectAll();
		}

		private void txtSNActivos_Click(object sender, System.EventArgs e)
		{
			this.txtSNActivos.SelectAll();
		}

		private void txtSNActivos_Enter(object sender, System.EventArgs e)
		{
			this.txtSNActivos.SelectAll();
		}

		private void txtSNPasivos_Enter(object sender, System.EventArgs e)
		{
			this.txtSNPasivos.SelectAll();
		}

		private void txtSNPasivos_Click(object sender, System.EventArgs e)
		{
		this.txtSNPasivos.SelectAll();
		}

		private void txtSDIngresos_Click(object sender, System.EventArgs e)
		{
			this.txtSDIngresos.SelectAll();
		}

		private void txtSDIngresos_Enter(object sender, System.EventArgs e)
		{
			this.txtSDIngresos.SelectAll();
		}

		private void txtSDGastos_Enter(object sender, System.EventArgs e)
		{
			this.txtSDGastos.SelectAll();
		}

		private void txtSDGastos_Click(object sender, System.EventArgs e)
		{
		this.txtSDGastos.SelectAll();
		}

		private void txtSDActivos_Click(object sender, System.EventArgs e)
		{
			this.txtSDActivos.SelectAll();
		}

		private void txtSDActivos_Enter(object sender, System.EventArgs e)
		{
			this.txtSDActivos.SelectAll();
		}

		private void txtSDPasivos_Enter(object sender, System.EventArgs e)
		{
			this.txtSDPasivos.SelectAll();
		}

		private void txtSDPasivos_Click(object sender, System.EventArgs e)
		{
			this.txtSDPasivos.SelectAll();
		}

		private void frmCred_GaranteSolicitud_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Escape) this.Close();		
		}

		private void cmbCProvinciaNacimiento_ValueChanged(object sender, System.EventArgs e)
		{
			if (this.cmbCProvinciaNacimiento.ActiveRow != null)
			{
				this.cmbCCantonNacimiento.DataSource = Funcion.dvProcedimiento(cdsSeteoF, string.Format("Exec Listas 1, {0}, 0", (int) this.cmbCProvinciaNacimiento.Value));
			}
		}

		private void btnCancelar_Click(object sender, System.EventArgs e)
		{
			DialogResult res = 	MessageBox.Show("Seguro que desea Cancelar?","Point Technology",	MessageBoxButtons.YesNo);
			if(res == DialogResult.Yes){
				LimpiaControles();
				this.txtDocumento.Focus();
				this.utcSolicitud.Tabs[0].Selected = true; 
				habilita_Control_Garante();
				deshabilita_Control_Conyugue();
				deshabilita_Control_Trabajo();
				deshabilita_Control_Empresa();
			}
		}

		private void txtCDocumento_Validating(object sender, System.ComponentModel.CancelEventArgs e)
		{
						
			if (this.cmbCTipoRuc.ActiveRow != null)
			{
				int iTipoRuc = 0;
				if ((int) this.cmbTipoRuc.Value == 1) iTipoRuc = 5;// cedula
				if ((int) this.cmbTipoRuc.Value == 2) iTipoRuc = 6;// pasaporte
				if ((int) this.cmbTipoRuc.Value == 3) iTipoRuc = 4;// ruc
				if (!Validacion.vbIdentificacion(this.txtCDocumento, iTipoRuc, cdsSeteoF)) 
				{
					e.Cancel = true;    }
			}
		}

		private void txtCDocumento_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
		validar.Solonumeros(e);
		}

		private void txtSNRuc_Validating(object sender, System.ComponentModel.CancelEventArgs e)
		{
//		  int iTipoRuc = 4;
//			if (!Validacion.vbIdentificacion(this.txtSNRuc, iTipoRuc, cdsSeteoF)) 
//			{
//				e.Cancel = true;    
//			}
		}

		private void txtSNRuc_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
		validar.Solonumeros(e);
		}

		private void txtSDRuc_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
		validar.Solonumeros(e);
		}

		private void txtSDRuc_Validating(object sender, System.ComponentModel.CancelEventArgs e)
		{
			int iTipoRuc = 4;
//			if (!Validacion.vbIdentificacion(this.txtSDRuc, iTipoRuc, cdsSeteoF)) 
//			{
//				e.Cancel = true;    
//			}
		}

		private void txtScore_Click(object sender, System.EventArgs e)
		{
		this.txtScore.SelectAll();
		}

		private void txtScore_Enter(object sender, System.EventArgs e)
		{
			this.txtScore.SelectAll();
		}

		private void txtSDPago_Enter(object sender, System.EventArgs e)
		{
			this.txtSDPago.SelectAll();
		}

		private void txtSDPago_Click(object sender, System.EventArgs e)
		{
			this.txtSDPago.SelectAll();
		}

		private void cmbRProvincia_RowSelected(object sender, Infragistics.Win.UltraWinGrid.RowSelectedEventArgs e)
		{
			this.cmbRProvincia.Update();
			if (e.Row != null)
			{

				if ((int)e.Row.Cells["idProvincia"].Value > 0 )
				{
					int iProvincia = (int)e.Row.Cells["idProvincia"].Value;
					this.cmbRCanton.DataSource = Funcion.dvProcedimiento(cdsSeteoF, string.Format("Exec Listas 1, {0}, 0", iProvincia));
					this.cmbRProvincia.Update();
					this.cmbRCanton.Update();
				}
			}
			this.cmbRProvincia.Update();
			this.cmbRCanton.Update();
		}

		private void cmbRCanton_Validating(object sender, System.ComponentModel.CancelEventArgs e)
		{
			this.cmbRProvincia.Update();
			this.cmbRCanton.Update();
		}

		private void cmbRCanton_Validated(object sender, System.EventArgs e)
		{
			this.cmbRProvincia.Update();
			this.cmbRCanton.Update();
		}

		private void groupBox1_Enter(object sender, System.EventArgs e)
		{
		
		}

		private void groupBox3_Enter(object sender, System.EventArgs e)
		{
		
		}

		private void ultraTabPageControl1_Paint(object sender, System.Windows.Forms.PaintEventArgs e)
		{
		
		}

		private void utpImagenes_Paint(object sender, System.Windows.Forms.PaintEventArgs e)
		{
		
		}

		private void uGridParentesco_InitializeLayout(object sender, Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs e)
		{
		
		}


		class validar 
		{
			public static void Solonumeros (KeyPressEventArgs Nu)
			{
				if(char.IsDigit(Nu.KeyChar))
				{
			
					Nu.Handled = false;
			
				}
				else if (char.IsControl(Nu.KeyChar))
				{
			
					Nu.Handled = false;			} 
				else
				{
					Nu.Handled= true;
				}

			}

			public static void SoloLetras (KeyPressEventArgs Le)
			{
				if(char.IsLetter(Le.KeyChar))
				{
			
					Le.Handled = false;
			
				}
				else if (char.IsControl(Le.KeyChar))
				{
			
					Le.Handled = false;			
				} 
				else
				{
					Le.Handled= true;
				}
			}
			public static void CaracteresEspeciales (KeyPressEventArgs e)
			{

				if (!char.IsLetter(e.KeyChar) && !char.IsDigit(e.KeyChar) && !char.IsWhiteSpace(e.KeyChar) && e.KeyChar !='\b')
				{
					e.Handled = true;
				}
				else
				{
					e.Handled = false;
				}
			}
		}

	}
}