using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using System.Data;
using System.Data.SqlClient;
using System.Globalization;
using System.Runtime.InteropServices;
using System.Diagnostics;
using System.IO;
using Infragistics.Win.UltraWinGrid;
using CrystalDecisions;
using CrystalDecisions.CrystalReports.Engine;
using CrystalDecisions.ReportSource;
using CrystalDecisions.Shared;

namespace Latinium
{
	/// <summary>
	/// Descripción breve de frmCre_Solicitudes.
	/// </summary>
	public class frmCre_Solicitudes : System.Windows.Forms.Form
	{
		private System.Windows.Forms.Label label1;
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
		private System.Windows.Forms.Label label16;
		private System.Windows.Forms.Label lblEntidad;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbBodega;
		private System.Windows.Forms.Button btnGarante;
		private Infragistics.Win.UltraWinEditors.UltraCheckEditor chkGarante;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource7;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource4;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource8;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource9;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource2;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource6;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource3;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource1;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource5;
		private C1.Data.C1DataSet cdsSeteoF;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource10;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbEntidad;
		private Infragistics.Win.UltraWinTabControl.UltraTabControl tbBasico;
		private Infragistics.Win.UltraWinTabControl.UltraTabSharedControlsPage ultraTabSharedControlsPage1;
		private Infragistics.Win.UltraWinTabControl.UltraTabPageControl ultraTabPageControl1;
		private Infragistics.Win.UltraWinTabControl.UltraTabPageControl tpComplemento;
		private Infragistics.Win.UltraWinTabControl.UltraTabPageControl ultraTabPageControl5;
		private Infragistics.Win.UltraWinTabControl.UltraTabPageControl ultraTabPageControl7;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource26;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource11;
		private System.Windows.Forms.GroupBox groupBox2;
		private System.Windows.Forms.Label label21;
		private System.Windows.Forms.Label label22;
		private System.Windows.Forms.Label label23;
		private System.Windows.Forms.Label label24;
		private System.Windows.Forms.Label label25;
		private System.Windows.Forms.Label label26;
		private System.Windows.Forms.Label label27;
		private System.Windows.Forms.Label label28;
		private System.Windows.Forms.Label label29;
		private System.Windows.Forms.Label label30;
		private System.Windows.Forms.Label label31;
		private System.Windows.Forms.Label label32;
		private System.Windows.Forms.Label label33;
		private System.Windows.Forms.Label label34;
		private System.Windows.Forms.GroupBox groupBox1;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbDDBarrio;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbDDParroquia;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbDDCanton;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbDDProvincia;
		private System.Windows.Forms.Label label7;
		private System.Windows.Forms.Label label8;
		private System.Windows.Forms.Label label9;
		private System.Windows.Forms.Label label10;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtDDCallaPrincipal;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtDDNumeroCasa;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtDDCalleSecundaria;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtDDTelefono1;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtDDTelefono2;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtDDCelular;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtDDEmail;
		private System.Windows.Forms.Label label12;
		private System.Windows.Forms.Label label13;
		private System.Windows.Forms.Label label14;
		private System.Windows.Forms.Label label15;
		private System.Windows.Forms.Label label17;
		private System.Windows.Forms.Label label18;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtValorVehiculo;
		private Infragistics.Win.UltraWinEditors.UltraCheckEditor chkHipoteca;
		private Infragistics.Win.UltraWinEditors.UltraCheckEditor chkSeparacionBienes;
		public Infragistics.Win.UltraWinEditors.UltraTextEditor txtModelo;
		public Infragistics.Win.UltraWinEditors.UltraTextEditor txtVehiculoMarca;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtDependientes;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbEstadoCivil;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbCantonNacimiento;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbProvinciaNAcimiento;
		private Infragistics.Win.UltraWinSchedule.UltraCalendarCombo dtFechaNacimiento;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbNacionalidad;
		private System.Windows.Forms.Label label19;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtNombreArrendador;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtTelfArrendador;
		private System.Windows.Forms.Label label20;
		private System.Windows.Forms.Label label35;
		private System.Windows.Forms.Label label36;
		private System.Windows.Forms.Label label37;
		private System.Windows.Forms.Label label38;
		private System.Windows.Forms.Label label39;
		private System.Windows.Forms.Label label40;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbSexo;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbNivelEducacion;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbProfesion;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbSituacionLaboral;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbDDInmueble;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbDDCiudadInmueble;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtDDReferenciasU;
		private System.Windows.Forms.GroupBox groupBox3;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbDPCSituacionL;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbDPCProfesion;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbDPCNivelEducacion;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbDPCSexo;
		private System.Windows.Forms.Label label44;
		private System.Windows.Forms.Label label45;
		private System.Windows.Forms.Label label46;
		private System.Windows.Forms.Label label47;
		private System.Windows.Forms.Label label50;
		private System.Windows.Forms.Label label51;
		private System.Windows.Forms.Label label52;
		private System.Windows.Forms.Label label53;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbDPCCantonN;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbDPCProvinciaN;
		private Infragistics.Win.UltraWinSchedule.UltraCalendarCombo dtDPCFnacimiento;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbDPCNacionalidad;
		private System.Windows.Forms.Label label54;
		private System.Windows.Forms.Label label41;
		private System.Windows.Forms.Label label42;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbDPCTipodocumento;
		public Infragistics.Win.UltraWinEditors.UltraTextEditor txtDPCdocumento;
		private System.Windows.Forms.Label label43;
		private System.Windows.Forms.Label label55;
		private System.Windows.Forms.Label label56;
		private System.Windows.Forms.Label label57;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtDPCsegundoN;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtDPCPrimerN;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtDPCApellidoM;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtDPCApellidoP;
		private System.Windows.Forms.GroupBox groupBox4;
		private System.Windows.Forms.Label label63;
		private System.Windows.Forms.Label label64;
		private System.Windows.Forms.Label label65;
		private System.Windows.Forms.Label label66;
		private System.Windows.Forms.Label label67;
		private System.Windows.Forms.Label label68;
		private System.Windows.Forms.Label label69;
		private System.Windows.Forms.Label label70;
		private System.Windows.Forms.Label label71;
		private System.Windows.Forms.Label label72;
		private System.Windows.Forms.Label label73;
		private System.Windows.Forms.Label label74;
		private System.Windows.Forms.Label label48;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtDPCTNombreEmpresa;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbDPCTActividadEmpresa;
		private System.Windows.Forms.Label label49;
		private System.Windows.Forms.Label label58;
		private System.Windows.Forms.Label label59;
		private System.Windows.Forms.Label label60;
		private System.Windows.Forms.Label label61;
		private System.Windows.Forms.Label label62;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtDPCTReferencia;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtDPCTEmail;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtDPCTCelularC;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtDPCTExt;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtDPCTtelefono;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtDPCTCalleSecundaria;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtDPCTNumeroCasa;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtDPCTCallePrincipal;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbDPCTBarrio;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbDPCTParroquia;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbDPCTCanton;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbDPCTProvincia;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbDPCTCargo;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtDPCTIngresos;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtDPCTGastos;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtDPCTActivos;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtDPCTPasivos;
		private Infragistics.Win.UltraWinTabControl.UltraTabPageControl ultraTabPageControl2;
		private Infragistics.Win.UltraWinTabControl.UltraTabPageControl ultraTabPageControl3;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource12;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource13;
		private Infragistics.Win.UltraWinGrid.UltraDropDown cmbParentesco;
		private Infragistics.Win.UltraWinGrid.UltraDropDown cmbProvincia;
		private Infragistics.Win.UltraWinGrid.UltraDropDown ultraDropDown1;
		public Infragistics.Win.UltraWinGrid.UltraGrid ultgconsulta;
		private Infragistics.Win.UltraWinGrid.UltraDropDown cmbtipoReferencia;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource14;
		private Infragistics.Win.UltraWinGrid.UltraDropDown cmbTipoBanco;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource15;
		private Infragistics.Win.UltraWinGrid.UltraDropDown cmbtipoCuenta;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource16;
		public Infragistics.Win.UltraWinGrid.UltraGrid uGridRefernciaBancaria;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource17;
		private System.Windows.Forms.GroupBox groupBox5;
		private System.Windows.Forms.Label Nombre;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtNombreDelNegocio;
		private System.Windows.Forms.Label label81;
		public Infragistics.Win.UltraWinEditors.UltraTextEditor txtNegocioRuc;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbNegciotipo;
		private System.Windows.Forms.Label label89;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbNEgocioActivdadEconomica;
		private Infragistics.Win.UltraWinSchedule.UltraCalendarCombo dtpNegocioFechaInicio;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtNegocioMetros;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtNegocioEmpleados;
		private System.Windows.Forms.Label label75;
		private System.Windows.Forms.Label label76;
		private System.Windows.Forms.Label label77;
		private System.Windows.Forms.Label label78;
		private System.Windows.Forms.Label label79;
		private System.Windows.Forms.Label label80;
		private System.Windows.Forms.Label label82;
		private System.Windows.Forms.Label label83;
		private System.Windows.Forms.Label label84;
		private System.Windows.Forms.Label label85;
		private System.Windows.Forms.Label label86;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtNegocioTelefono1;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtNegocioCalleSecundaria;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtNegocioNumeroCasa;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtNegocioCallePrincipal;
		private System.Windows.Forms.Label label87;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtNegocioemail;
		private System.Windows.Forms.Label label88;
		private System.Windows.Forms.Label label90;
		private System.Windows.Forms.Label label91;
		private System.Windows.Forms.Label label92;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbNegocioBarrio;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbNegocioParroquia;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbNegocioCanton;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbNegocioProvincia;
		private System.Windows.Forms.Label label93;
		private System.Windows.Forms.Label label94;
		private System.Windows.Forms.Label label95;
		private System.Windows.Forms.Label label96;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtNegocioTelefono2;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtNEgocioPasivos;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtNegocioActivos;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtNEgocioGastos;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtNEgocioIngresos;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtNegocioReferncia;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtNEgocioCelular;
		private Infragistics.Win.UltraWinEditors.UltraCheckEditor chkNegocioObligadoaLlevar;
		private Infragistics.Win.UltraWinEditors.UltraCheckEditor chkAfiliadoTributario;
		private System.Windows.Forms.GroupBox groupBox6;
		private System.Windows.Forms.Label label101;
		private System.Windows.Forms.Label label102;
		private System.Windows.Forms.Label label103;
		private System.Windows.Forms.Label label104;
		private System.Windows.Forms.Label label105;
		private System.Windows.Forms.Label label106;
		private System.Windows.Forms.Label label107;
		private System.Windows.Forms.Label label108;
		private System.Windows.Forms.Label label109;
		private System.Windows.Forms.Label label110;
		private System.Windows.Forms.Label label111;
		private System.Windows.Forms.Label label112;
		private System.Windows.Forms.Label label113;
		private System.Windows.Forms.Label label114;
		private System.Windows.Forms.Label label115;
		private System.Windows.Forms.Label label116;
		private System.Windows.Forms.Label label117;
		private System.Windows.Forms.Label label118;
		private System.Windows.Forms.Label label119;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtDependienteActivos;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtDependienteGastos;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtDependienteIngresos;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtDependienteEmpresa;
		public Infragistics.Win.UltraWinEditors.UltraTextEditor txtDependienteRuc;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbDependienteTipocontrato;
		private Infragistics.Win.UltraWinSchedule.UltraCalendarCombo dtDependienteFechaIngreso;
		private System.Windows.Forms.Label label97;
		private System.Windows.Forms.Label label98;
		private System.Windows.Forms.Label label99;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtDependienteEXT;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtDependienteemail;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbDependienteBarrio;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbDependienteParroquia;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbDependienteCanton;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbDependienteProvincia;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtDependienteReferencia;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtDependienteCelular;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtDependienteTelefono;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtDependienteCalleSecundaria;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtDependienteNumeroCasa;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtDependienteCallePrinciapl;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtDependienteDiasPago;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtDependienteDepartamento;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbDependienteCargo;
		private System.Windows.Forms.Label label100;
		private System.Windows.Forms.Label label120;
		private System.Windows.Forms.Label label121;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtDependientePasivos;
		private System.Windows.Forms.GroupBox groupBox7;
		private System.Windows.Forms.Label label134;
		private System.Windows.Forms.Label label135;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbDependienteTipoSueldo;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource18;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource19;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource20;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource21;
		private System.Windows.Forms.Label label122;
		private System.Windows.Forms.Label label125;
		private Infragistics.Win.UltraWinSchedule.UltraCalendarCombo dtCreditoFechaSolicitud;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtCreditoScore;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbCreditoEstado;
		private System.Windows.Forms.Label label141;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtCupo;
		private System.Windows.Forms.GroupBox groupBox8;
		private System.Windows.Forms.Label label126;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtCreditoObservaciones;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource22;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource23;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource24;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbDependienteTipoempresa;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource25;
		private System.Windows.Forms.Label lblFISL;
		private System.Windows.Forms.Label lblEstado;
		private System.Windows.Forms.Label lblUsuarioCompleta;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbTipoRuc;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource27;
		private Infragistics.Win.UltraWinEditors.UltraCheckEditor chkAfiliadoIESS;
		public Infragistics.Win.UltraWinGrid.UltraGrid utgHistorialCredito;
		private System.Windows.Forms.Label label6;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtNegocioIdCre_DatosUbicacion;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtClienteIdCre_EmpresasTrabajo;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtEmpresaIdCre_DatosUbicacion;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtConyugueIdCre_DatosGenerales;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtConyugueIdCre_DatosUbicacion;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtClienteIdCre_Negocio;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtIdCre_Conyugue;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtIdCre_Solicitud;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtClienteIdCre_DatosGenerales;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtClienteIdCre_DatosUbicacion;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtEstado;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource28;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource29;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource30;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource31;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource32;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource33;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbDPCActividadE;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource34;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource35;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource36;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource37;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource38;
		private System.Windows.Forms.Button btnGuardar;
		private System.Windows.Forms.Button btnSalir;
		private System.Windows.Forms.Button btnConsultar;
		private System.Windows.Forms.Button btnImprimir;
		private Infragistics.Win.UltraWinSchedule.UltraCalendarCombo dtFecha;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource39;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource40;
		private System.Windows.Forms.Button btnCelularcliente;
		private System.Windows.Forms.Button btnClienteTelefonoCasa1;
		private System.Windows.Forms.Button btnClienteTelefonoCasa2;
		private System.Windows.Forms.Button btnConyugueTelefono;
		private System.Windows.Forms.Button btnConyugueCelular;
		private System.Windows.Forms.Button btnNegocioTelefono2;
		private System.Windows.Forms.Button btnNegocioCelular;
		private System.Windows.Forms.Button btnNegocioTelefono1;
		private System.Windows.Forms.Button btnTrabajoTelefono;
		private System.Windows.Forms.Button btnTrabajoCelular;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource41;
		private Infragistics.Win.UltraWinSchedule.UltraCalendarCombo dtTiempoVivienda;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbDDTipoVivienda;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbActividadEconomica;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtIdGarante;
		private Infragistics.Win.UltraWinTabControl.UltraTabPageControl ultraTabPageControl4;
		public Infragistics.Win.UltraWinGrid.UltraGrid ultraGrid1;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource42;
		private System.Windows.Forms.Label label11;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtCreditoScoreSobreEndeudamiento;
		private System.Windows.Forms.Label label127;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtCreditoScoreInclusion;
		private System.Windows.Forms.Label label128;
		private System.Windows.Forms.Label label129;
		private System.Windows.Forms.Label label130;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtCapacidadPago;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtTotalVencido;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtRangoIngresos;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtSegmentacion;
		private System.Windows.Forms.Label label131;
		private System.Windows.Forms.Label label132;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtResultado;
		private System.Windows.Forms.Label label133;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtCarteraCastigada;
		private System.Windows.Forms.Button btnIdValidator;
		private Infragistics.Win.UltraWinTabControl.UltraTabPageControl ultraTabPageControl6;
		private System.Windows.Forms.Label label136;
		private System.Windows.Forms.Label label137;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbidTipoCliente;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbdetallecliente;
		private System.Windows.Forms.Label label138;
		private System.Windows.Forms.GroupBox groupBox9;
		private Infragistics.Win.UltraWinGrid.UltraCombo ultraCombo2;
		private Infragistics.Win.UltraWinGrid.UltraCombo ultraCombo3;
		private System.Windows.Forms.Label label139;
		private System.Windows.Forms.Label label140;
		private Infragistics.Win.UltraWinEditors.UltraCheckEditor ChkOmite;
		private Infragistics.Win.UltraWinEditors.UltraCheckEditor chkNuovo;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource43;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource44;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource45;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbSituacionLaboral2;
		private Infragistics.Win.UltraWinGrid.UltraGrid gdObservaciones;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource46;
		private System.Windows.Forms.Label label143;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmdEstadoSolicitud;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource47;
		private System.Windows.Forms.Button btnImprimeVerificacion;
		private System.Windows.Forms.Button btnComandos;
		private System.Windows.Forms.ContextMenu mnuAcciones;
		private System.Windows.Forms.MenuItem mnuAccionesCopiarUltimoMes;
		private System.Windows.Forms.MenuItem mnuAccionesAgregarNuevoLocal;
		private System.Windows.Forms.MenuItem menuItem1;
		private System.Windows.Forms.MenuItem mnuAccionesActualizar;
		private System.Windows.Forms.Label label144;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtObsActividadEconomica;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtDDValorInmueble;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtCreditoCuotaAsiganada;
		private System.Windows.Forms.Label label124;
		private System.Windows.Forms.Label label123;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbCreditoCalificacion;
		private System.Windows.Forms.Label label142;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtObservaciones;
		private System.Windows.Forms.Label label145;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtNotas;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource48;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbTiempo;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource49;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbTiempoDependiente;
		private Infragistics.Win.UltraWinGrid.UltraCombo cmbTiempoVivienda;
		private System.Windows.Forms.Label label146;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtCelArrendador;
		private System.Windows.Forms.Label lblFechaActualizacion;
		private System.Windows.Forms.Button btnVerificacion;
		private System.Windows.Forms.Button btnUbicacion;
		private Infragistics.Win.UltraWinTabControl.UltraTabPageControl ultraTabPageControl8;
		private Infragistics.Win.UltraWinDataSource.UltraDataSource ultraDataSource50;
		public Infragistics.Win.UltraWinGrid.UltraGrid utgVerificacionTerrena;
		private Infragistics.Win.UltraWinTabControl.UltraTabControl ultraTabControl1;
		private Infragistics.Win.UltraWinTabControl.UltraTabSharedControlsPage ultraTabSharedControlsPage2;
		private Infragistics.Win.UltraWinTabControl.UltraTabPageControl ultraTabPageControl15;
		public Infragistics.Win.UltraWinGrid.UltraGrid ultraGrid5;
		private Infragistics.Win.UltraWinTabControl.UltraTabPageControl ultraTabPageControl9;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtVDTiempoMeses;
		private System.Windows.Forms.Label label147;
		private Infragistics.Win.UltraWinEditors.UltraOptionSet opdtVDVivienda;
		private System.Windows.Forms.Label label148;
		private Infragistics.Win.UltraWinEditors.UltraOptionSet opdtVDEstadoVivienda;
		private System.Windows.Forms.Label label149;
		private Infragistics.Win.UltraWinEditors.UltraOptionSet opdtVDZonaVivienda;
		private Infragistics.Win.UltraWinEditors.UltraOptionSet opdtVDPropiedad;
		private Infragistics.Win.UltraWinEditors.UltraOptionSet opdtVDAcceso;
		private System.Windows.Forms.Label label150;
		private System.Windows.Forms.Label label151;
		private System.Windows.Forms.Label label152;
		private Infragistics.Win.UltraWinEditors.UltraOptionSet opdtVDCobertura;
		private System.Windows.Forms.Label label153;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtVDPuntoReferencia;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtVDentrevistada;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtVDVecino;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtVDObservacion;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtVDSecundaria;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtVDPrincipal;
		private System.Windows.Forms.Label label154;
		private System.Windows.Forms.Label label155;
		private System.Windows.Forms.Label label156;
		private System.Windows.Forms.Label label157;
		private System.Windows.Forms.Label label158;
		private System.Windows.Forms.Label label159;
		private System.Windows.Forms.Label label160;
		public Infragistics.Win.UltraWinEditors.UltraTextEditor txtVDGPS;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtVDCCelArr;
		public Infragistics.Win.UltraWinEditors.UltraTextEditor txtVDLongitud;
		public Infragistics.Win.UltraWinEditors.UltraTextEditor txtVDLatitud;
		private System.Windows.Forms.Button btnImagenes;
		private System.Windows.Forms.Label label161;
		private Infragistics.Win.UltraWinEditors.UltraOptionSet opdtVTTipoTrabajo;
		private System.Windows.Forms.Label label162;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtVTTiempoMeses;
		private System.Windows.Forms.Label label163;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtVTTiempoYear;
		private Infragistics.Win.UltraWinEditors.UltraNumericEditor txtVTIngresos;
		private System.Windows.Forms.Label label164;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtVTActividadLaboral;
		private System.Windows.Forms.Label label165;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtVTTelefono;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtVTLatitud;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtVTGps;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtVTReferencia;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtVTPersonaEntrevistada;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtVTLongitud;
		private System.Windows.Forms.Label label166;
		private System.Windows.Forms.Label label167;
		private System.Windows.Forms.Label label168;
		private System.Windows.Forms.Label label169;
		private System.Windows.Forms.Label label170;
		private System.Windows.Forms.Label label171;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtVTPRincipal;
		private System.Windows.Forms.Label label172;
		private Infragistics.Win.UltraWinEditors.UltraTextEditor txtVTSecundaria;
		private System.Windows.Forms.Label label173;
		private System.Windows.Forms.Button btnImagenestrabajo;
		private System.Windows.Forms.Button button1;
		private System.Windows.Forms.Button btnEquifax;
		/// <summary>
		/// Variable del diseñador requerida.
		/// </summary>
		private System.ComponentModel.Container components = null;

		public frmCre_Solicitudes()
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
    bool bCarga = false;
		int idGaranteG = 0;
		int IdCre_SolicitudG = 0; 
		int IdCre_DatosGeneralesG = 0; 
		int txtIdCre_DatosUbicacionG = 0;
		bool bHabilitaPerfil = false;

		DateTime NegocioFechaInicio;

		#region Código generado por el Diseñador de Windows Forms
		/// <summary>
		/// Método necesario para admitir el Diseñador. No se puede modificar
		/// el contenido del método con el editor de código.
		/// </summary>
		private void InitializeComponent()
		{
			Infragistics.Win.Appearance appearance1 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance2 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance3 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance4 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance5 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance6 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance7 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance8 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance9 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance10 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance11 = new Infragistics.Win.Appearance();
			Infragistics.Win.ValueListItem valueListItem1 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.ValueListItem valueListItem2 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.ValueListItem valueListItem3 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.Appearance appearance12 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand1 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn1 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("institucion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn2 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("emisor");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn3 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("antiguedad");
			Infragistics.Win.Appearance appearance13 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn4 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("cupo");
			Infragistics.Win.Appearance appearance14 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance15 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance16 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance17 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance18 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance19 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn1 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("institucion");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn2 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("emisor");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn3 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("antiguedad");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn4 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("cupo");
			Infragistics.Win.Appearance appearance20 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance21 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance22 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance23 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance24 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance25 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance26 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance27 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance28 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance29 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance30 = new Infragistics.Win.Appearance();
			Infragistics.Win.ValueListItem valueListItem4 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.ValueListItem valueListItem5 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.Appearance appearance31 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance32 = new Infragistics.Win.Appearance();
			Infragistics.Win.ValueListItem valueListItem6 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.ValueListItem valueListItem7 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.Appearance appearance33 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance34 = new Infragistics.Win.Appearance();
			Infragistics.Win.ValueListItem valueListItem8 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.ValueListItem valueListItem9 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.ValueListItem valueListItem10 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.Appearance appearance35 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance36 = new Infragistics.Win.Appearance();
			Infragistics.Win.ValueListItem valueListItem11 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.ValueListItem valueListItem12 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.Appearance appearance37 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance38 = new Infragistics.Win.Appearance();
			Infragistics.Win.ValueListItem valueListItem13 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.ValueListItem valueListItem14 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.ValueListItem valueListItem15 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.Appearance appearance39 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance40 = new Infragistics.Win.Appearance();
			Infragistics.Win.ValueListItem valueListItem16 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.ValueListItem valueListItem17 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.ValueListItem valueListItem18 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.ValueListItem valueListItem19 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.ValueListItem valueListItem20 = new Infragistics.Win.ValueListItem();
			Infragistics.Win.Appearance appearance41 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand2 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn5 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idActEconomica");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn6 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn5 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idActEconomica");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn6 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Nombre");
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand3 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn7 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idSituacionLaboral");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn8 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Descripcion");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn7 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idSituacionLaboral");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn8 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Descripcion");
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand4 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn9 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idProfesion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn10 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn9 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idProfesion");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn10 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Nombre");
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand5 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn11 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idNivelEducacion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn12 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn11 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idNivelEducacion");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn12 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Nombre");
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand6 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn13 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idSexo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn14 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn13 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idSexo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn14 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Nombre");
			Infragistics.Win.Appearance appearance42 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance43 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance44 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance45 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand7 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn15 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idEdoCivil");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn16 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn15 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idEdoCivil");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn16 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Nombre");
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand8 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn17 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCanton");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn18 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn17 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idCanton");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn18 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Nombre");
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand9 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn19 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idProvincia");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn20 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn19 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idProvincia");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn20 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Nombre");
			Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton dateButton1 = new Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton();
			System.Resources.ResourceManager resources = new System.Resources.ResourceManager(typeof(frmCre_Solicitudes));
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand10 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn21 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idNacionalidad");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn22 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn21 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idNacionalidad");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn22 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Nombre");
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand11 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn23 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCre_Tiempo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn24 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Descripcion");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn23 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idCre_Tiempo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn24 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Descripcion");
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand12 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn25 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idTipoVivienda");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn26 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn25 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idTipoVivienda");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn26 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Nombre");
			Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton dateButton2 = new Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand13 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn27 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCanton");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn28 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn27 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idCanton");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn28 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Nombre");
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand14 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn29 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idInmueble");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn30 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn29 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idInmueble");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn30 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Nombre");
			Infragistics.Win.Appearance appearance46 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance47 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance48 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance49 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance50 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance51 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand15 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn31 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idBarrio");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn32 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn31 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idBarrio");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn32 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Nombre");
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand16 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn33 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idParroquia");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn34 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn33 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idParroquia");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn34 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Nombre");
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand17 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn35 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCanton");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn36 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand18 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn37 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idProvincia");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn38 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand19 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn39 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idActEconomica");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn40 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.Appearance appearance52 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance53 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance54 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance55 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance56 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand20 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn41 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idSituacionLaboral");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn42 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Descripcion");
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand21 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn43 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idProfesion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn44 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand22 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn45 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idNivelEducacion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn46 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn35 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idNivelEducacion");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn36 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Nombre");
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand23 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn47 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idSexo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn48 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand24 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn49 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCanton");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn50 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand25 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn51 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idProvincia");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn52 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton dateButton3 = new Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand26 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn53 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idNacionalidad");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn54 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand27 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn55 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCargo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn56 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn37 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idCargo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn38 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Nombre");
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand28 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn57 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idActEconomica");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn58 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn39 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idActEconomica");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn40 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Nombre");
			Infragistics.Win.Appearance appearance57 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance58 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance59 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance60 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance61 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance62 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand29 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn59 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idBarrio");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn60 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand30 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn61 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idParroquia");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn62 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand31 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn63 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCanton");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn64 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand32 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn65 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idProvincia");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn66 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand33 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn67 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idTipoCuentaBan");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn68 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn41 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idTipoCuentaBan");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn42 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Nombre");
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand34 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn69 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idBancos");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn70 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn43 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idBancos");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn44 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Nombre");
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand35 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn71 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idTipoRefBancaria");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn72 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn45 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idTipoRefBancaria");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn46 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Nombre");
			Infragistics.Win.Appearance appearance63 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand36 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn73 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCre_ReferenciasClientes");
			Infragistics.Win.Appearance appearance64 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn74 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCre_DatosGenerales");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn75 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Indice");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn76 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idParentesco", -1, "cmbParentesco");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn77 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("ApellidoPaterno");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn78 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("ApellidoMaterno");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn79 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("PrimerNombre");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn80 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("SegundoNombre");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn81 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idProvincia", -1, "cmbProvincia");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn82 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCanton", -1, "ultraDropDown1");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn83 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Direccion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn84 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Telefono");
			Infragistics.Win.Appearance appearance65 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn85 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Celular");
			Infragistics.Win.Appearance appearance66 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn86 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("EmpresaTrabajo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn87 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("DireccionTrabajo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn88 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("TelefonoTrabajo");
			Infragistics.Win.Appearance appearance67 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn89 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("ExtensionTrabajo");
			Infragistics.Win.Appearance appearance68 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn90 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("btnTelefono", 0);
			Infragistics.Win.Appearance appearance69 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn91 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("btnCelular", 1);
			Infragistics.Win.Appearance appearance70 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn92 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("btnTelefonoTrabj", 2);
			Infragistics.Win.Appearance appearance71 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance72 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance73 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance74 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance75 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn47 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idCre_ReferenciasClientes");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn48 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idCre_DatosGenerales");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn49 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Indice");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn50 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idParentesco");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn51 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("ApellidoPaterno");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn52 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("ApellidoMaterno");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn53 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("PrimerNombre");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn54 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("SegundoNombre");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn55 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idProvincia");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn56 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idCanton");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn57 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Direccion");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn58 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Telefono");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn59 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Celular");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn60 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("EmpresaTrabajo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn61 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("DireccionTrabajo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn62 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("TelefonoTrabajo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn63 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("ExtensionTrabajo");
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand37 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn93 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCanton");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn94 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand38 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn95 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idProvincia");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn96 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand39 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn97 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idParentesco");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn98 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn64 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idParentesco");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn65 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Nombre");
			Infragistics.Win.Appearance appearance76 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand40 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn99 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCre_ReferenciaBancaria");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn100 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCre_DatosGenerales");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn101 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idTipoRefBancaria", -1, "cmbtipoReferencia");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn102 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idBancos", -1, "cmbTipoBanco");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn103 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idTipoCuentaBan", -1, "cmbtipoCuenta");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn104 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("NumeroDeCuenta", -1, null, 0, Infragistics.Win.UltraWinGrid.SortIndicator.Ascending, false);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn105 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("CupoTarjeta");
			Infragistics.Win.Appearance appearance77 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance78 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance79 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance80 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn66 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idCre_ReferenciaBancaria");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn67 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idCre_DatosGenerales");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn68 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idTipoRefBancaria");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn69 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idBancos");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn70 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idTipoCuentaBan");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn71 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("NumeroDeCuenta");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn72 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("CupoTarjeta");
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand41 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn106 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCre_Tiempo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn107 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Descripcion");
			Infragistics.Win.Appearance appearance81 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance82 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance83 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance84 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand42 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn108 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idBarrio");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn109 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand43 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn110 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idParroquia");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn111 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand44 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn112 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCanton");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn113 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand45 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn114 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idProvincia");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn115 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.Appearance appearance85 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance86 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance87 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance88 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance89 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance90 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand46 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn116 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idTipoVivienda");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn117 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn73 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idTipoVivienda");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn74 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Nombre");
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand47 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn118 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idSubActEcoNegocio");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn119 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn75 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idSubActEcoNegocio");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn76 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Nombre");
			Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton dateButton4 = new Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand48 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn120 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCre_Tiempo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn121 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Descripcion");
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand49 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn122 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idTipoEmpresa");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn123 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Descripción");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn77 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idTipoEmpresa");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn78 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Descripción");
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand50 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn124 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idTipoSueldo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn125 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn79 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idTipoSueldo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn80 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Nombre");
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand51 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn126 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCargo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn127 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn81 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idCargo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn82 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Nombre");
			Infragistics.Win.Appearance appearance91 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance92 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance93 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand52 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn128 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idBarrio");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn129 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand53 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn130 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idParroquia");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn131 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand54 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn132 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCanton");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn133 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand55 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn134 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idProvincia");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn135 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.Appearance appearance94 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance95 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance96 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance97 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance98 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance99 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand56 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn136 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idTipoContrato");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn137 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn83 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idTipoContrato");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn84 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Nombre");
			Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton dateButton5 = new Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton();
			Infragistics.Win.Appearance appearance100 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand57 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn138 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Numero");
			Infragistics.Win.Appearance appearance101 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn139 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Ruc");
			Infragistics.Win.Appearance appearance102 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance103 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn140 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre", -1, null, 0, Infragistics.Win.UltraWinGrid.SortIndicator.Ascending, false);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn141 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Fecha");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn142 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Banco");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn143 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Lote");
			Infragistics.Win.Appearance appearance104 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance105 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn144 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Motivo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn145 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Bodega");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn146 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("NumCuotas");
			Infragistics.Win.Appearance appearance106 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance107 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn147 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Monto");
			Infragistics.Win.Appearance appearance108 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance109 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn148 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("FechaPrimerPago");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn149 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("ValorCuotas");
			Infragistics.Win.Appearance appearance110 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance111 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn150 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("CuotasPendientes");
			Infragistics.Win.Appearance appearance112 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance113 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn151 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("CuotasPagadas");
			Infragistics.Win.Appearance appearance114 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance115 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn152 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("SaldoCapital");
			Infragistics.Win.Appearance appearance116 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance117 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn153 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("PromedioRetraso");
			Infragistics.Win.Appearance appearance118 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance119 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn154 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Ultimo_Pago");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn155 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Estado");
			Infragistics.Win.Appearance appearance120 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance121 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance122 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance123 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn85 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Numero");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn86 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Ruc");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn87 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Nombre");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn88 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Fecha");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn89 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Banco");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn90 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Lote");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn91 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Motivo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn92 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Bodega");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn93 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("NumCuotas");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn94 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Monto");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn95 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("FechaPrimerPago");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn96 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("ValorCuotas");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn97 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("CuotasPendientes");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn98 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("CuotasPagadas");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn99 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("SaldoCapital");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn100 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("PromedioRetraso");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn101 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Ultimo_Pago");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn102 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Estado");
			Infragistics.Win.Appearance appearance124 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance125 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance126 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance127 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance128 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance129 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance130 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton dateButton6 = new Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand58 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn156 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idDetalleTipoCliente");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn157 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn103 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idDetalleTipoCliente");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn104 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Nombre");
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand59 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn158 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idTipoCliente");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn159 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn105 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idTipoCliente");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn106 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Nombre");
			Infragistics.Win.Appearance appearance131 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand60 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn160 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idTipoCalificacion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn161 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Codigo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn107 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idTipoCalificacion");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn108 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Codigo");
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand61 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn162 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idNegadoPendiente");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn163 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn109 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idNegadoPendiente");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn110 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Nombre");
			Infragistics.Win.Appearance appearance132 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand62 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn164 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Observacion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn165 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Usuario");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn166 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Fecha");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn167 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn168 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Banco");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn169 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Estado");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn170 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Codigo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn171 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Score");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn172 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Cupo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn173 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Pantalla");
			Infragistics.Win.Appearance appearance133 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance134 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance135 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance136 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance137 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance138 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn111 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Observacion");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn112 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Usuario");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn113 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Fecha");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn114 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Nombre");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn115 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Banco");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn116 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Estado");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn117 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Codigo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn118 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Score");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn119 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Cupo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn120 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Pantalla");
			Infragistics.Win.Appearance appearance139 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance140 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance141 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand63 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn174 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idEstado");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn175 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Estado");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn121 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idEstado");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn122 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Estado");
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand64 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn176 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idSituacionLaboral");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn177 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Descripcion");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn123 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idSituacionLaboral");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn124 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Descripcion");
			Infragistics.Win.Appearance appearance142 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand65 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn178 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("institucion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn179 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("emisor");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn180 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("antiguedad");
			Infragistics.Win.Appearance appearance143 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn181 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("cupo");
			Infragistics.Win.Appearance appearance144 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance145 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance146 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance147 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance148 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance149 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance150 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance151 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinTabControl.UltraTab ultraTab1 = new Infragistics.Win.UltraWinTabControl.UltraTab();
			Infragistics.Win.UltraWinTabControl.UltraTab ultraTab2 = new Infragistics.Win.UltraWinTabControl.UltraTab();
			Infragistics.Win.Appearance appearance152 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand66 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn182 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Numero");
			Infragistics.Win.Appearance appearance153 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance154 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn183 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCre_solicitud");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn184 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idCre_DatosGenerales");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn185 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Ruc");
			Infragistics.Win.Appearance appearance155 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance156 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn186 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombres");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn187 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("iEstado");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn188 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Estado");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn189 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Almacen");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn190 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("bDomicilio");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn191 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("bTrabajo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn192 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Fecha Envio");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn193 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Fecha Respuesta");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn194 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Usuario Asigna");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn195 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Verificador");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn196 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("UrlGoogle");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn197 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Domicilio Asignada");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn198 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Trabajo Asignada");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn199 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idClienteVerificacion");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn200 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idTerrenaGestionDomicilio");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn201 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idTerrenaGestionTrabajo");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn202 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Anular", 0);
			Infragistics.Win.Appearance appearance157 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn203 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Aprobar", 1);
			Infragistics.Win.Appearance appearance158 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn204 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Cobrador", 2);
			Infragistics.Win.Appearance appearance159 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn205 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Documento", 3);
			Infragistics.Win.Appearance appearance160 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn206 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Trabajo", 4);
			Infragistics.Win.Appearance appearance161 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn207 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Domicilio", 5);
			Infragistics.Win.Appearance appearance162 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance163 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance164 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance165 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance166 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn125 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Numero");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn126 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idCre_solicitud");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn127 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idCre_DatosGenerales");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn128 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Ruc");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn129 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Nombres");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn130 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("iEstado");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn131 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Estado");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn132 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Almacen");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn133 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("bDomicilio");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn134 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("bTrabajo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn135 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Fecha Envio");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn136 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Fecha Respuesta");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn137 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Usuario Asigna");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn138 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Verificador");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn139 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("UrlGoogle");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn140 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Domicilio Asignada");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn141 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Trabajo Asignada");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn142 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idClienteVerificacion");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn143 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idTerrenaGestionDomicilio");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn144 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idTerrenaGestionTrabajo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn145 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Bodega");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn146 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Nombre");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn147 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idCreActEconomicas");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn148 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Descripcion");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn149 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idFuenteIngreso");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn150 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Nombre");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn151 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idInmueble");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn152 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Nombre");
			Infragistics.Win.Appearance appearance167 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton dateButton7 = new Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton();
			Infragistics.Win.Appearance appearance168 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance169 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance170 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance171 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand67 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn208 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Bodega");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn209 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.UltraWinGrid.UltraGridBand ultraGridBand68 = new Infragistics.Win.UltraWinGrid.UltraGridBand("Band 0", -1);
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn210 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("idEntidadFinanciera");
			Infragistics.Win.UltraWinGrid.UltraGridColumn ultraGridColumn211 = new Infragistics.Win.UltraWinGrid.UltraGridColumn("Nombre");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn153 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idEntidadFinanciera");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn154 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Nombre");
			Infragistics.Win.Appearance appearance172 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn155 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idEstado");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn156 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Estado");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn157 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idNegadoPendiente");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn158 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Nombre");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn159 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idTipoCalificacion");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn160 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Codigo");
			Infragistics.Win.Appearance appearance173 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance174 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinTabControl.UltraTab ultraTab3 = new Infragistics.Win.UltraWinTabControl.UltraTab();
			Infragistics.Win.Appearance appearance175 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance176 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance177 = new Infragistics.Win.Appearance();
			Infragistics.Win.Appearance appearance178 = new Infragistics.Win.Appearance();
			Infragistics.Win.UltraWinTabControl.UltraTab ultraTab4 = new Infragistics.Win.UltraWinTabControl.UltraTab();
			Infragistics.Win.UltraWinTabControl.UltraTab ultraTab5 = new Infragistics.Win.UltraWinTabControl.UltraTab();
			Infragistics.Win.UltraWinTabControl.UltraTab ultraTab6 = new Infragistics.Win.UltraWinTabControl.UltraTab();
			Infragistics.Win.UltraWinTabControl.UltraTab ultraTab7 = new Infragistics.Win.UltraWinTabControl.UltraTab();
			Infragistics.Win.UltraWinTabControl.UltraTab ultraTab8 = new Infragistics.Win.UltraWinTabControl.UltraTab();
			Infragistics.Win.UltraWinTabControl.UltraTab ultraTab9 = new Infragistics.Win.UltraWinTabControl.UltraTab();
			Infragistics.Win.UltraWinTabControl.UltraTab ultraTab10 = new Infragistics.Win.UltraWinTabControl.UltraTab();
			Infragistics.Win.UltraWinTabControl.UltraTab ultraTab11 = new Infragistics.Win.UltraWinTabControl.UltraTab();
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn161 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idProvincia");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn162 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Nombre");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn163 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("idCre_Tiempo");
			Infragistics.Win.UltraWinDataSource.UltraDataColumn ultraDataColumn164 = new Infragistics.Win.UltraWinDataSource.UltraDataColumn("Descripcion");
			this.ultraTabPageControl15 = new Infragistics.Win.UltraWinTabControl.UltraTabPageControl();
			this.label173 = new System.Windows.Forms.Label();
			this.label172 = new System.Windows.Forms.Label();
			this.txtVTSecundaria = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.label171 = new System.Windows.Forms.Label();
			this.txtVTPRincipal = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.label170 = new System.Windows.Forms.Label();
			this.label169 = new System.Windows.Forms.Label();
			this.label168 = new System.Windows.Forms.Label();
			this.label167 = new System.Windows.Forms.Label();
			this.label166 = new System.Windows.Forms.Label();
			this.txtVTLongitud = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.txtVTPersonaEntrevistada = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.txtVTReferencia = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.txtVTGps = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.txtVTLatitud = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.txtVTTelefono = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.label165 = new System.Windows.Forms.Label();
			this.txtVTActividadLaboral = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.label164 = new System.Windows.Forms.Label();
			this.txtVTIngresos = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.label163 = new System.Windows.Forms.Label();
			this.txtVTTiempoYear = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.label162 = new System.Windows.Forms.Label();
			this.txtVTTiempoMeses = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.label161 = new System.Windows.Forms.Label();
			this.opdtVTTipoTrabajo = new Infragistics.Win.UltraWinEditors.UltraOptionSet();
			this.ultraGrid5 = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.ultraDataSource42 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.ultraTabPageControl9 = new Infragistics.Win.UltraWinTabControl.UltraTabPageControl();
			this.btnImagenes = new System.Windows.Forms.Button();
			this.txtVDLatitud = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.txtVDLongitud = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.txtVDCCelArr = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtVDGPS = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.label160 = new System.Windows.Forms.Label();
			this.label159 = new System.Windows.Forms.Label();
			this.label158 = new System.Windows.Forms.Label();
			this.label157 = new System.Windows.Forms.Label();
			this.label156 = new System.Windows.Forms.Label();
			this.label155 = new System.Windows.Forms.Label();
			this.label154 = new System.Windows.Forms.Label();
			this.txtVDSecundaria = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.txtVDPrincipal = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.txtVDVecino = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.txtVDObservacion = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.txtVDentrevistada = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.txtVDPuntoReferencia = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.label153 = new System.Windows.Forms.Label();
			this.opdtVDCobertura = new Infragistics.Win.UltraWinEditors.UltraOptionSet();
			this.label152 = new System.Windows.Forms.Label();
			this.label151 = new System.Windows.Forms.Label();
			this.label150 = new System.Windows.Forms.Label();
			this.opdtVDAcceso = new Infragistics.Win.UltraWinEditors.UltraOptionSet();
			this.opdtVDPropiedad = new Infragistics.Win.UltraWinEditors.UltraOptionSet();
			this.opdtVDZonaVivienda = new Infragistics.Win.UltraWinEditors.UltraOptionSet();
			this.label149 = new System.Windows.Forms.Label();
			this.opdtVDEstadoVivienda = new Infragistics.Win.UltraWinEditors.UltraOptionSet();
			this.label148 = new System.Windows.Forms.Label();
			this.opdtVDVivienda = new Infragistics.Win.UltraWinEditors.UltraOptionSet();
			this.label147 = new System.Windows.Forms.Label();
			this.txtVDTiempoMeses = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.ultraTabPageControl1 = new Infragistics.Win.UltraWinTabControl.UltraTabPageControl();
			this.groupBox2 = new System.Windows.Forms.GroupBox();
			this.label144 = new System.Windows.Forms.Label();
			this.txtObsActividadEconomica = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.cmbActividadEconomica = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.ultraDataSource40 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.cmbSituacionLaboral = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.ultraDataSource30 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.cmbProfesion = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.ultraDataSource29 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.cmbNivelEducacion = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.ultraDataSource28 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.cmbSexo = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.ultraDataSource24 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.label21 = new System.Windows.Forms.Label();
			this.label22 = new System.Windows.Forms.Label();
			this.label23 = new System.Windows.Forms.Label();
			this.chkHipoteca = new Infragistics.Win.UltraWinEditors.UltraCheckEditor();
			this.chkSeparacionBienes = new Infragistics.Win.UltraWinEditors.UltraCheckEditor();
			this.txtValorVehiculo = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtModelo = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.txtVehiculoMarca = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.label24 = new System.Windows.Forms.Label();
			this.label25 = new System.Windows.Forms.Label();
			this.label26 = new System.Windows.Forms.Label();
			this.label27 = new System.Windows.Forms.Label();
			this.label28 = new System.Windows.Forms.Label();
			this.txtDependientes = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.label29 = new System.Windows.Forms.Label();
			this.cmbEstadoCivil = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.ultraDataSource11 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.label30 = new System.Windows.Forms.Label();
			this.label31 = new System.Windows.Forms.Label();
			this.label32 = new System.Windows.Forms.Label();
			this.label33 = new System.Windows.Forms.Label();
			this.cmbCantonNacimiento = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.ultraDataSource2 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.cmbProvinciaNAcimiento = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.ultraDataSource1 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.dtFechaNacimiento = new Infragistics.Win.UltraWinSchedule.UltraCalendarCombo();
			this.cmbNacionalidad = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.ultraDataSource26 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.label34 = new System.Windows.Forms.Label();
			this.groupBox1 = new System.Windows.Forms.GroupBox();
			this.btnEquifax = new System.Windows.Forms.Button();
			this.button1 = new System.Windows.Forms.Button();
			this.btnUbicacion = new System.Windows.Forms.Button();
			this.label146 = new System.Windows.Forms.Label();
			this.txtCelArrendador = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.cmbTiempoVivienda = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.ultraDataSource49 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.txtDDValorInmueble = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.cmbDDTipoVivienda = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.ultraDataSource31 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.dtTiempoVivienda = new Infragistics.Win.UltraWinSchedule.UltraCalendarCombo();
			this.cmbDDCiudadInmueble = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.ultraDataSource38 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.cmbDDInmueble = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.ultraDataSource27 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.label40 = new System.Windows.Forms.Label();
			this.label39 = new System.Windows.Forms.Label();
			this.label38 = new System.Windows.Forms.Label();
			this.label37 = new System.Windows.Forms.Label();
			this.label36 = new System.Windows.Forms.Label();
			this.label35 = new System.Windows.Forms.Label();
			this.label20 = new System.Windows.Forms.Label();
			this.txtTelfArrendador = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.txtNombreArrendador = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.label19 = new System.Windows.Forms.Label();
			this.txtDDReferenciasU = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.btnCelularcliente = new System.Windows.Forms.Button();
			this.btnClienteTelefonoCasa2 = new System.Windows.Forms.Button();
			this.btnClienteTelefonoCasa1 = new System.Windows.Forms.Button();
			this.label18 = new System.Windows.Forms.Label();
			this.label17 = new System.Windows.Forms.Label();
			this.label15 = new System.Windows.Forms.Label();
			this.label14 = new System.Windows.Forms.Label();
			this.label13 = new System.Windows.Forms.Label();
			this.label12 = new System.Windows.Forms.Label();
			this.txtDDEmail = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.txtDDCelular = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.txtDDTelefono2 = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.txtDDTelefono1 = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.txtDDCalleSecundaria = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.txtDDNumeroCasa = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.txtDDCallaPrincipal = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.label10 = new System.Windows.Forms.Label();
			this.label9 = new System.Windows.Forms.Label();
			this.label8 = new System.Windows.Forms.Label();
			this.label7 = new System.Windows.Forms.Label();
			this.cmbDDBarrio = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.ultraDataSource4 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.cmbDDParroquia = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.ultraDataSource3 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.cmbDDCanton = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.cmbDDProvincia = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.label6 = new System.Windows.Forms.Label();
			this.tpComplemento = new Infragistics.Win.UltraWinTabControl.UltraTabPageControl();
			this.groupBox3 = new System.Windows.Forms.GroupBox();
			this.cmbDPCActividadE = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.label43 = new System.Windows.Forms.Label();
			this.label55 = new System.Windows.Forms.Label();
			this.label56 = new System.Windows.Forms.Label();
			this.label57 = new System.Windows.Forms.Label();
			this.txtDPCsegundoN = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.txtDPCPrimerN = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.txtDPCApellidoM = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.txtDPCApellidoP = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.label41 = new System.Windows.Forms.Label();
			this.label42 = new System.Windows.Forms.Label();
			this.cmbDPCTipodocumento = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.txtDPCdocumento = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.cmbDPCSituacionL = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.cmbDPCProfesion = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.cmbDPCNivelEducacion = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.ultraDataSource33 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.cmbDPCSexo = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.label44 = new System.Windows.Forms.Label();
			this.label45 = new System.Windows.Forms.Label();
			this.label46 = new System.Windows.Forms.Label();
			this.label47 = new System.Windows.Forms.Label();
			this.label50 = new System.Windows.Forms.Label();
			this.label51 = new System.Windows.Forms.Label();
			this.label52 = new System.Windows.Forms.Label();
			this.label53 = new System.Windows.Forms.Label();
			this.cmbDPCCantonN = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.cmbDPCProvinciaN = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.dtDPCFnacimiento = new Infragistics.Win.UltraWinSchedule.UltraCalendarCombo();
			this.cmbDPCNacionalidad = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.label54 = new System.Windows.Forms.Label();
			this.groupBox4 = new System.Windows.Forms.GroupBox();
			this.label62 = new System.Windows.Forms.Label();
			this.label61 = new System.Windows.Forms.Label();
			this.label60 = new System.Windows.Forms.Label();
			this.label59 = new System.Windows.Forms.Label();
			this.txtDPCTPasivos = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtDPCTActivos = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtDPCTGastos = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtDPCTIngresos = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.label58 = new System.Windows.Forms.Label();
			this.label49 = new System.Windows.Forms.Label();
			this.cmbDPCTCargo = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.ultraDataSource36 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.cmbDPCTActividadEmpresa = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.ultraDataSource35 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.label48 = new System.Windows.Forms.Label();
			this.txtDPCTNombreEmpresa = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.label63 = new System.Windows.Forms.Label();
			this.txtDPCTReferencia = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.btnConyugueCelular = new System.Windows.Forms.Button();
			this.btnConyugueTelefono = new System.Windows.Forms.Button();
			this.label64 = new System.Windows.Forms.Label();
			this.label65 = new System.Windows.Forms.Label();
			this.label66 = new System.Windows.Forms.Label();
			this.label67 = new System.Windows.Forms.Label();
			this.label68 = new System.Windows.Forms.Label();
			this.label69 = new System.Windows.Forms.Label();
			this.label70 = new System.Windows.Forms.Label();
			this.txtDPCTEmail = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.txtDPCTCelularC = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.txtDPCTExt = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.txtDPCTtelefono = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.txtDPCTCalleSecundaria = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.txtDPCTNumeroCasa = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.txtDPCTCallePrincipal = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.label71 = new System.Windows.Forms.Label();
			this.label72 = new System.Windows.Forms.Label();
			this.label73 = new System.Windows.Forms.Label();
			this.label74 = new System.Windows.Forms.Label();
			this.cmbDPCTBarrio = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.cmbDPCTParroquia = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.cmbDPCTCanton = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.cmbDPCTProvincia = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.ultraTabPageControl5 = new Infragistics.Win.UltraWinTabControl.UltraTabPageControl();
			this.cmbtipoCuenta = new Infragistics.Win.UltraWinGrid.UltraDropDown();
			this.ultraDataSource16 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.cmbTipoBanco = new Infragistics.Win.UltraWinGrid.UltraDropDown();
			this.ultraDataSource15 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.cmbtipoReferencia = new Infragistics.Win.UltraWinGrid.UltraDropDown();
			this.ultraDataSource14 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.ultgconsulta = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.ultraDataSource13 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.ultraDropDown1 = new Infragistics.Win.UltraWinGrid.UltraDropDown();
			this.cmbProvincia = new Infragistics.Win.UltraWinGrid.UltraDropDown();
			this.cmbParentesco = new Infragistics.Win.UltraWinGrid.UltraDropDown();
			this.ultraDataSource9 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.uGridRefernciaBancaria = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.ultraDataSource17 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.ultraTabPageControl7 = new Infragistics.Win.UltraWinTabControl.UltraTabPageControl();
			this.groupBox5 = new System.Windows.Forms.GroupBox();
			this.cmbTiempo = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.label145 = new System.Windows.Forms.Label();
			this.txtNotas = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.chkAfiliadoTributario = new Infragistics.Win.UltraWinEditors.UltraCheckEditor();
			this.chkNegocioObligadoaLlevar = new Infragistics.Win.UltraWinEditors.UltraCheckEditor();
			this.btnNegocioTelefono2 = new System.Windows.Forms.Button();
			this.txtNegocioTelefono2 = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.label96 = new System.Windows.Forms.Label();
			this.label95 = new System.Windows.Forms.Label();
			this.label94 = new System.Windows.Forms.Label();
			this.label93 = new System.Windows.Forms.Label();
			this.label87 = new System.Windows.Forms.Label();
			this.txtNegocioemail = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.label88 = new System.Windows.Forms.Label();
			this.label90 = new System.Windows.Forms.Label();
			this.label91 = new System.Windows.Forms.Label();
			this.label92 = new System.Windows.Forms.Label();
			this.cmbNegocioBarrio = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.cmbNegocioParroquia = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.cmbNegocioCanton = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.cmbNegocioProvincia = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.label75 = new System.Windows.Forms.Label();
			this.label76 = new System.Windows.Forms.Label();
			this.label77 = new System.Windows.Forms.Label();
			this.label78 = new System.Windows.Forms.Label();
			this.txtNEgocioPasivos = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtNegocioActivos = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtNEgocioGastos = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtNEgocioIngresos = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.label79 = new System.Windows.Forms.Label();
			this.txtNegocioReferncia = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.btnNegocioCelular = new System.Windows.Forms.Button();
			this.btnNegocioTelefono1 = new System.Windows.Forms.Button();
			this.label80 = new System.Windows.Forms.Label();
			this.label82 = new System.Windows.Forms.Label();
			this.label83 = new System.Windows.Forms.Label();
			this.label84 = new System.Windows.Forms.Label();
			this.label85 = new System.Windows.Forms.Label();
			this.label86 = new System.Windows.Forms.Label();
			this.txtNEgocioCelular = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.txtNegocioTelefono1 = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.txtNegocioCalleSecundaria = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.txtNegocioNumeroCasa = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.txtNegocioCallePrincipal = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.txtNegocioEmpleados = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtNegocioMetros = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.Nombre = new System.Windows.Forms.Label();
			this.txtNombreDelNegocio = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.label81 = new System.Windows.Forms.Label();
			this.txtNegocioRuc = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.cmbNegciotipo = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.ultraDataSource37 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.label89 = new System.Windows.Forms.Label();
			this.cmbNEgocioActivdadEconomica = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.ultraDataSource25 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.dtpNegocioFechaInicio = new Infragistics.Win.UltraWinSchedule.UltraCalendarCombo();
			this.ultraTabPageControl2 = new Infragistics.Win.UltraWinTabControl.UltraTabPageControl();
			this.groupBox6 = new System.Windows.Forms.GroupBox();
			this.cmbTiempoDependiente = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.cmbDependienteTipoempresa = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.ultraDataSource18 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.cmbDependienteTipoSueldo = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.ultraDataSource20 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.label121 = new System.Windows.Forms.Label();
			this.label120 = new System.Windows.Forms.Label();
			this.label100 = new System.Windows.Forms.Label();
			this.label99 = new System.Windows.Forms.Label();
			this.label98 = new System.Windows.Forms.Label();
			this.label97 = new System.Windows.Forms.Label();
			this.cmbDependienteCargo = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.ultraDataSource21 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.txtDependienteDepartamento = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.chkAfiliadoIESS = new Infragistics.Win.UltraWinEditors.UltraCheckEditor();
			this.btnTrabajoCelular = new System.Windows.Forms.Button();
			this.txtDependienteEXT = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.label101 = new System.Windows.Forms.Label();
			this.txtDependienteemail = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.label102 = new System.Windows.Forms.Label();
			this.label103 = new System.Windows.Forms.Label();
			this.label104 = new System.Windows.Forms.Label();
			this.label105 = new System.Windows.Forms.Label();
			this.cmbDependienteBarrio = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.cmbDependienteParroquia = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.cmbDependienteCanton = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.cmbDependienteProvincia = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.label106 = new System.Windows.Forms.Label();
			this.label107 = new System.Windows.Forms.Label();
			this.label108 = new System.Windows.Forms.Label();
			this.label109 = new System.Windows.Forms.Label();
			this.txtDependientePasivos = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtDependienteActivos = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtDependienteGastos = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtDependienteIngresos = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.label110 = new System.Windows.Forms.Label();
			this.txtDependienteReferencia = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.btnTrabajoTelefono = new System.Windows.Forms.Button();
			this.label111 = new System.Windows.Forms.Label();
			this.label112 = new System.Windows.Forms.Label();
			this.label113 = new System.Windows.Forms.Label();
			this.label114 = new System.Windows.Forms.Label();
			this.label115 = new System.Windows.Forms.Label();
			this.label116 = new System.Windows.Forms.Label();
			this.txtDependienteCelular = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.txtDependienteTelefono = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.txtDependienteCalleSecundaria = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.txtDependienteNumeroCasa = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.txtDependienteCallePrinciapl = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.txtDependienteDiasPago = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.label117 = new System.Windows.Forms.Label();
			this.txtDependienteEmpresa = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.label118 = new System.Windows.Forms.Label();
			this.txtDependienteRuc = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.cmbDependienteTipocontrato = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.ultraDataSource19 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.label119 = new System.Windows.Forms.Label();
			this.dtDependienteFechaIngreso = new Infragistics.Win.UltraWinSchedule.UltraCalendarCombo();
			this.ultraTabPageControl3 = new Infragistics.Win.UltraWinTabControl.UltraTabPageControl();
			this.groupBox8 = new System.Windows.Forms.GroupBox();
			this.utgHistorialCredito = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.ultraDataSource41 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.groupBox7 = new System.Windows.Forms.GroupBox();
			this.txtCarteraCastigada = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.label133 = new System.Windows.Forms.Label();
			this.txtResultado = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.label132 = new System.Windows.Forms.Label();
			this.txtSegmentacion = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.label131 = new System.Windows.Forms.Label();
			this.txtRangoIngresos = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.txtTotalVencido = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtCapacidadPago = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.label128 = new System.Windows.Forms.Label();
			this.label129 = new System.Windows.Forms.Label();
			this.label130 = new System.Windows.Forms.Label();
			this.label127 = new System.Windows.Forms.Label();
			this.txtCreditoScoreInclusion = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.label11 = new System.Windows.Forms.Label();
			this.txtCreditoScoreSobreEndeudamiento = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.lblUsuarioCompleta = new System.Windows.Forms.Label();
			this.txtCreditoObservaciones = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.label126 = new System.Windows.Forms.Label();
			this.label122 = new System.Windows.Forms.Label();
			this.txtCreditoScore = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.label134 = new System.Windows.Forms.Label();
			this.label135 = new System.Windows.Forms.Label();
			this.dtCreditoFechaSolicitud = new Infragistics.Win.UltraWinSchedule.UltraCalendarCombo();
			this.ultraTabPageControl6 = new Infragistics.Win.UltraWinTabControl.UltraTabPageControl();
			this.cmbdetallecliente = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.ultraDataSource44 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.cmbidTipoCliente = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.ultraDataSource43 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.label137 = new System.Windows.Forms.Label();
			this.groupBox9 = new System.Windows.Forms.GroupBox();
			this.label142 = new System.Windows.Forms.Label();
			this.txtObservaciones = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.label124 = new System.Windows.Forms.Label();
			this.label123 = new System.Windows.Forms.Label();
			this.cmbCreditoCalificacion = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.ultraDataSource22 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.txtCreditoCuotaAsiganada = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.btnImprimeVerificacion = new System.Windows.Forms.Button();
			this.label143 = new System.Windows.Forms.Label();
			this.cmdEstadoSolicitud = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.ultraDataSource47 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.gdObservaciones = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.ultraDataSource46 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.ChkOmite = new Infragistics.Win.UltraWinEditors.UltraCheckEditor();
			this.chkNuovo = new Infragistics.Win.UltraWinEditors.UltraCheckEditor();
			this.label141 = new System.Windows.Forms.Label();
			this.txtCupo = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.label125 = new System.Windows.Forms.Label();
			this.cmbCreditoEstado = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.ultraDataSource23 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.label138 = new System.Windows.Forms.Label();
			this.cmbSituacionLaboral2 = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.ultraDataSource45 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.label136 = new System.Windows.Forms.Label();
			this.ultraCombo2 = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.ultraCombo3 = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.label139 = new System.Windows.Forms.Label();
			this.label140 = new System.Windows.Forms.Label();
			this.ultraTabPageControl4 = new Infragistics.Win.UltraWinTabControl.UltraTabPageControl();
			this.ultraGrid1 = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.ultraTabPageControl8 = new Infragistics.Win.UltraWinTabControl.UltraTabPageControl();
			this.ultraTabControl1 = new Infragistics.Win.UltraWinTabControl.UltraTabControl();
			this.ultraTabSharedControlsPage2 = new Infragistics.Win.UltraWinTabControl.UltraTabSharedControlsPage();
			this.utgVerificacionTerrena = new Infragistics.Win.UltraWinGrid.UltraGrid();
			this.ultraDataSource50 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.ultraDataSource5 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.ultraDataSource34 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.ultraDataSource39 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.ultraDataSource32 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.label1 = new System.Windows.Forms.Label();
			this.lblBodega = new System.Windows.Forms.Label();
			this.txtDocumento = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.dtFecha = new Infragistics.Win.UltraWinSchedule.UltraCalendarCombo();
			this.label5 = new System.Windows.Forms.Label();
			this.label4 = new System.Windows.Forms.Label();
			this.label3 = new System.Windows.Forms.Label();
			this.label2 = new System.Windows.Forms.Label();
			this.txtSnombre = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.txtPnombre = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.txtApellidoM = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.txtApellidoP = new Infragistics.Win.UltraWinEditors.UltraTextEditor();
			this.label16 = new System.Windows.Forms.Label();
			this.lblEntidad = new System.Windows.Forms.Label();
			this.cmbBodega = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.cmbEntidad = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.ultraDataSource10 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.btnGarante = new System.Windows.Forms.Button();
			this.chkGarante = new Infragistics.Win.UltraWinEditors.UltraCheckEditor();
			this.ultraDataSource7 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.ultraDataSource8 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.ultraDataSource6 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.cdsSeteoF = new C1.Data.C1DataSet();
			this.tbBasico = new Infragistics.Win.UltraWinTabControl.UltraTabControl();
			this.ultraTabSharedControlsPage1 = new Infragistics.Win.UltraWinTabControl.UltraTabSharedControlsPage();
			this.ultraDataSource12 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.lblFISL = new System.Windows.Forms.Label();
			this.lblEstado = new System.Windows.Forms.Label();
			this.cmbTipoRuc = new Infragistics.Win.UltraWinGrid.UltraCombo();
			this.txtNegocioIdCre_DatosUbicacion = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtClienteIdCre_EmpresasTrabajo = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtEmpresaIdCre_DatosUbicacion = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtConyugueIdCre_DatosGenerales = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtConyugueIdCre_DatosUbicacion = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtClienteIdCre_Negocio = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtIdCre_Conyugue = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtIdCre_Solicitud = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtClienteIdCre_DatosGenerales = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtClienteIdCre_DatosUbicacion = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.txtEstado = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.btnGuardar = new System.Windows.Forms.Button();
			this.btnSalir = new System.Windows.Forms.Button();
			this.btnConsultar = new System.Windows.Forms.Button();
			this.btnImprimir = new System.Windows.Forms.Button();
			this.txtIdGarante = new Infragistics.Win.UltraWinEditors.UltraNumericEditor();
			this.btnIdValidator = new System.Windows.Forms.Button();
			this.btnComandos = new System.Windows.Forms.Button();
			this.mnuAcciones = new System.Windows.Forms.ContextMenu();
			this.mnuAccionesCopiarUltimoMes = new System.Windows.Forms.MenuItem();
			this.mnuAccionesAgregarNuevoLocal = new System.Windows.Forms.MenuItem();
			this.menuItem1 = new System.Windows.Forms.MenuItem();
			this.mnuAccionesActualizar = new System.Windows.Forms.MenuItem();
			this.ultraDataSource48 = new Infragistics.Win.UltraWinDataSource.UltraDataSource();
			this.lblFechaActualizacion = new System.Windows.Forms.Label();
			this.btnVerificacion = new System.Windows.Forms.Button();
			this.ultraTabPageControl15.SuspendLayout();
			((System.ComponentModel.ISupportInitialize)(this.txtVTSecundaria)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtVTPRincipal)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtVTLongitud)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtVTPersonaEntrevistada)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtVTReferencia)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtVTGps)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtVTLatitud)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtVTTelefono)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtVTActividadLaboral)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtVTIngresos)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtVTTiempoYear)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtVTTiempoMeses)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.opdtVTTipoTrabajo)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraGrid5)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource42)).BeginInit();
			this.ultraTabPageControl9.SuspendLayout();
			((System.ComponentModel.ISupportInitialize)(this.txtVDLatitud)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtVDLongitud)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtVDCCelArr)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtVDGPS)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtVDSecundaria)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtVDPrincipal)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtVDVecino)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtVDObservacion)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtVDentrevistada)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtVDPuntoReferencia)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.opdtVDCobertura)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.opdtVDAcceso)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.opdtVDPropiedad)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.opdtVDZonaVivienda)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.opdtVDEstadoVivienda)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.opdtVDVivienda)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtVDTiempoMeses)).BeginInit();
			this.ultraTabPageControl1.SuspendLayout();
			this.groupBox2.SuspendLayout();
			((System.ComponentModel.ISupportInitialize)(this.txtObsActividadEconomica)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbActividadEconomica)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource40)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbSituacionLaboral)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource30)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbProfesion)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource29)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbNivelEducacion)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource28)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbSexo)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource24)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtValorVehiculo)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtModelo)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtVehiculoMarca)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtDependientes)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbEstadoCivil)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource11)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbCantonNacimiento)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource2)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbProvinciaNAcimiento)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource1)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.dtFechaNacimiento)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbNacionalidad)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource26)).BeginInit();
			this.groupBox1.SuspendLayout();
			((System.ComponentModel.ISupportInitialize)(this.txtCelArrendador)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbTiempoVivienda)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource49)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtDDValorInmueble)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbDDTipoVivienda)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource31)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.dtTiempoVivienda)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbDDCiudadInmueble)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource38)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbDDInmueble)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource27)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtTelfArrendador)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNombreArrendador)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtDDReferenciasU)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtDDEmail)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtDDCelular)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtDDTelefono2)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtDDTelefono1)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtDDCalleSecundaria)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtDDNumeroCasa)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtDDCallaPrincipal)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbDDBarrio)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource4)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbDDParroquia)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource3)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbDDCanton)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbDDProvincia)).BeginInit();
			this.tpComplemento.SuspendLayout();
			this.groupBox3.SuspendLayout();
			((System.ComponentModel.ISupportInitialize)(this.cmbDPCActividadE)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtDPCsegundoN)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtDPCPrimerN)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtDPCApellidoM)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtDPCApellidoP)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbDPCTipodocumento)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtDPCdocumento)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbDPCSituacionL)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbDPCProfesion)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbDPCNivelEducacion)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource33)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbDPCSexo)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbDPCCantonN)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbDPCProvinciaN)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.dtDPCFnacimiento)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbDPCNacionalidad)).BeginInit();
			this.groupBox4.SuspendLayout();
			((System.ComponentModel.ISupportInitialize)(this.txtDPCTPasivos)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtDPCTActivos)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtDPCTGastos)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtDPCTIngresos)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbDPCTCargo)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource36)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbDPCTActividadEmpresa)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource35)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtDPCTNombreEmpresa)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtDPCTReferencia)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtDPCTEmail)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtDPCTCelularC)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtDPCTExt)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtDPCTtelefono)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtDPCTCalleSecundaria)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtDPCTNumeroCasa)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtDPCTCallePrincipal)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbDPCTBarrio)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbDPCTParroquia)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbDPCTCanton)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbDPCTProvincia)).BeginInit();
			this.ultraTabPageControl5.SuspendLayout();
			((System.ComponentModel.ISupportInitialize)(this.cmbtipoCuenta)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource16)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbTipoBanco)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource15)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbtipoReferencia)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource14)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultgconsulta)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource13)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDropDown1)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbProvincia)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbParentesco)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource9)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.uGridRefernciaBancaria)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource17)).BeginInit();
			this.ultraTabPageControl7.SuspendLayout();
			this.groupBox5.SuspendLayout();
			((System.ComponentModel.ISupportInitialize)(this.cmbTiempo)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNotas)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNegocioTelefono2)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNegocioemail)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbNegocioBarrio)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbNegocioParroquia)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbNegocioCanton)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbNegocioProvincia)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNEgocioPasivos)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNegocioActivos)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNEgocioGastos)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNEgocioIngresos)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNegocioReferncia)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNEgocioCelular)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNegocioTelefono1)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNegocioCalleSecundaria)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNegocioNumeroCasa)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNegocioCallePrincipal)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNegocioEmpleados)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNegocioMetros)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNombreDelNegocio)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNegocioRuc)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbNegciotipo)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource37)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbNEgocioActivdadEconomica)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource25)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.dtpNegocioFechaInicio)).BeginInit();
			this.ultraTabPageControl2.SuspendLayout();
			this.groupBox6.SuspendLayout();
			((System.ComponentModel.ISupportInitialize)(this.cmbTiempoDependiente)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbDependienteTipoempresa)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource18)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbDependienteTipoSueldo)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource20)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbDependienteCargo)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource21)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtDependienteDepartamento)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtDependienteEXT)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtDependienteemail)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbDependienteBarrio)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbDependienteParroquia)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbDependienteCanton)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbDependienteProvincia)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtDependientePasivos)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtDependienteActivos)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtDependienteGastos)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtDependienteIngresos)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtDependienteReferencia)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtDependienteCelular)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtDependienteTelefono)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtDependienteCalleSecundaria)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtDependienteNumeroCasa)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtDependienteCallePrinciapl)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtDependienteDiasPago)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtDependienteEmpresa)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtDependienteRuc)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbDependienteTipocontrato)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource19)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.dtDependienteFechaIngreso)).BeginInit();
			this.ultraTabPageControl3.SuspendLayout();
			this.groupBox8.SuspendLayout();
			((System.ComponentModel.ISupportInitialize)(this.utgHistorialCredito)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource41)).BeginInit();
			this.groupBox7.SuspendLayout();
			((System.ComponentModel.ISupportInitialize)(this.txtCarteraCastigada)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtResultado)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtSegmentacion)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtRangoIngresos)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtTotalVencido)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtCapacidadPago)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtCreditoScoreInclusion)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtCreditoScoreSobreEndeudamiento)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtCreditoObservaciones)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtCreditoScore)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.dtCreditoFechaSolicitud)).BeginInit();
			this.ultraTabPageControl6.SuspendLayout();
			((System.ComponentModel.ISupportInitialize)(this.cmbdetallecliente)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource44)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbidTipoCliente)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource43)).BeginInit();
			this.groupBox9.SuspendLayout();
			((System.ComponentModel.ISupportInitialize)(this.txtObservaciones)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbCreditoCalificacion)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource22)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtCreditoCuotaAsiganada)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmdEstadoSolicitud)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource47)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.gdObservaciones)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource46)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtCupo)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbCreditoEstado)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource23)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbSituacionLaboral2)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource45)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraCombo2)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraCombo3)).BeginInit();
			this.ultraTabPageControl4.SuspendLayout();
			((System.ComponentModel.ISupportInitialize)(this.ultraGrid1)).BeginInit();
			this.ultraTabPageControl8.SuspendLayout();
			((System.ComponentModel.ISupportInitialize)(this.ultraTabControl1)).BeginInit();
			this.ultraTabControl1.SuspendLayout();
			((System.ComponentModel.ISupportInitialize)(this.utgVerificacionTerrena)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource50)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource5)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource34)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource39)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource32)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtDocumento)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.dtFecha)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtSnombre)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtPnombre)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtApellidoM)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtApellidoP)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbBodega)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbEntidad)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource10)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource7)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource8)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource6)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.tbBasico)).BeginInit();
			this.tbBasico.SuspendLayout();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource12)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbTipoRuc)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNegocioIdCre_DatosUbicacion)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtClienteIdCre_EmpresasTrabajo)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtEmpresaIdCre_DatosUbicacion)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtConyugueIdCre_DatosGenerales)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtConyugueIdCre_DatosUbicacion)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtClienteIdCre_Negocio)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIdCre_Conyugue)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIdCre_Solicitud)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtClienteIdCre_DatosGenerales)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtClienteIdCre_DatosUbicacion)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtEstado)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIdGarante)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource48)).BeginInit();
			this.SuspendLayout();
			// 
			// ultraTabPageControl15
			// 
			this.ultraTabPageControl15.Controls.Add(this.label173);
			this.ultraTabPageControl15.Controls.Add(this.label172);
			this.ultraTabPageControl15.Controls.Add(this.txtVTSecundaria);
			this.ultraTabPageControl15.Controls.Add(this.label171);
			this.ultraTabPageControl15.Controls.Add(this.txtVTPRincipal);
			this.ultraTabPageControl15.Controls.Add(this.label170);
			this.ultraTabPageControl15.Controls.Add(this.label169);
			this.ultraTabPageControl15.Controls.Add(this.label168);
			this.ultraTabPageControl15.Controls.Add(this.label167);
			this.ultraTabPageControl15.Controls.Add(this.label166);
			this.ultraTabPageControl15.Controls.Add(this.txtVTLongitud);
			this.ultraTabPageControl15.Controls.Add(this.txtVTPersonaEntrevistada);
			this.ultraTabPageControl15.Controls.Add(this.txtVTReferencia);
			this.ultraTabPageControl15.Controls.Add(this.txtVTGps);
			this.ultraTabPageControl15.Controls.Add(this.txtVTLatitud);
			this.ultraTabPageControl15.Controls.Add(this.txtVTTelefono);
			this.ultraTabPageControl15.Controls.Add(this.label165);
			this.ultraTabPageControl15.Controls.Add(this.txtVTActividadLaboral);
			this.ultraTabPageControl15.Controls.Add(this.label164);
			this.ultraTabPageControl15.Controls.Add(this.txtVTIngresos);
			this.ultraTabPageControl15.Controls.Add(this.label163);
			this.ultraTabPageControl15.Controls.Add(this.txtVTTiempoYear);
			this.ultraTabPageControl15.Controls.Add(this.label162);
			this.ultraTabPageControl15.Controls.Add(this.txtVTTiempoMeses);
			this.ultraTabPageControl15.Controls.Add(this.label161);
			this.ultraTabPageControl15.Controls.Add(this.opdtVTTipoTrabajo);
			this.ultraTabPageControl15.Controls.Add(this.ultraGrid5);
			this.ultraTabPageControl15.Location = new System.Drawing.Point(1, 23);
			this.ultraTabPageControl15.Name = "ultraTabPageControl15";
			this.ultraTabPageControl15.Size = new System.Drawing.Size(1092, 198);
			// 
			// label173
			// 
			this.label173.AutoSize = true;
			this.label173.BackColor = System.Drawing.Color.Transparent;
			this.label173.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label173.Location = new System.Drawing.Point(792, 8);
			this.label173.Name = "label173";
			this.label173.Size = new System.Drawing.Size(89, 17);
			this.label173.TabIndex = 1068;
			this.label173.Text = "Actividad Laboral";
			this.label173.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// label172
			// 
			this.label172.AutoSize = true;
			this.label172.BackColor = System.Drawing.Color.Transparent;
			this.label172.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label172.Location = new System.Drawing.Point(544, 168);
			this.label172.Name = "label172";
			this.label172.Size = new System.Drawing.Size(87, 17);
			this.label172.TabIndex = 1067;
			this.label172.Text = "Calle Secundaria";
			this.label172.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// txtVTSecundaria
			// 
			appearance1.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtVTSecundaria.Appearance = appearance1;
			this.txtVTSecundaria.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtVTSecundaria.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtVTSecundaria.Enabled = false;
			this.txtVTSecundaria.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.txtVTSecundaria.Location = new System.Drawing.Point(640, 168);
			this.txtVTSecundaria.Name = "txtVTSecundaria";
			this.txtVTSecundaria.Size = new System.Drawing.Size(448, 22);
			this.txtVTSecundaria.TabIndex = 1066;
			// 
			// label171
			// 
			this.label171.AutoSize = true;
			this.label171.BackColor = System.Drawing.Color.Transparent;
			this.label171.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label171.Location = new System.Drawing.Point(8, 168);
			this.label171.Name = "label171";
			this.label171.Size = new System.Drawing.Size(73, 17);
			this.label171.TabIndex = 1065;
			this.label171.Text = "Calle Principal";
			this.label171.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// txtVTPRincipal
			// 
			appearance2.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtVTPRincipal.Appearance = appearance2;
			this.txtVTPRincipal.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtVTPRincipal.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtVTPRincipal.Enabled = false;
			this.txtVTPRincipal.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.txtVTPRincipal.Location = new System.Drawing.Point(96, 168);
			this.txtVTPRincipal.Name = "txtVTPRincipal";
			this.txtVTPRincipal.Size = new System.Drawing.Size(440, 22);
			this.txtVTPRincipal.TabIndex = 1064;
			// 
			// label170
			// 
			this.label170.AutoSize = true;
			this.label170.BackColor = System.Drawing.Color.Transparent;
			this.label170.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label170.Location = new System.Drawing.Point(552, 112);
			this.label170.Name = "label170";
			this.label170.Size = new System.Drawing.Size(109, 17);
			this.label170.TabIndex = 1063;
			this.label170.Text = "Persona Entrevistada";
			this.label170.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// label169
			// 
			this.label169.AutoSize = true;
			this.label169.BackColor = System.Drawing.Color.Transparent;
			this.label169.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label169.Location = new System.Drawing.Point(8, 112);
			this.label169.Name = "label169";
			this.label169.Size = new System.Drawing.Size(145, 17);
			this.label169.TabIndex = 1062;
			this.label169.Text = "Punto de Referencia Laboral";
			this.label169.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// label168
			// 
			this.label168.AutoSize = true;
			this.label168.BackColor = System.Drawing.Color.Transparent;
			this.label168.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label168.Location = new System.Drawing.Point(784, 64);
			this.label168.Name = "label168";
			this.label168.Size = new System.Drawing.Size(24, 17);
			this.label168.TabIndex = 1061;
			this.label168.Text = "GPS";
			this.label168.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// label167
			// 
			this.label167.AutoSize = true;
			this.label167.BackColor = System.Drawing.Color.Transparent;
			this.label167.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label167.Location = new System.Drawing.Point(472, 64);
			this.label167.Name = "label167";
			this.label167.Size = new System.Drawing.Size(47, 17);
			this.label167.TabIndex = 1060;
			this.label167.Text = "Longitud";
			this.label167.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// label166
			// 
			this.label166.AutoSize = true;
			this.label166.BackColor = System.Drawing.Color.Transparent;
			this.label166.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label166.Location = new System.Drawing.Point(256, 64);
			this.label166.Name = "label166";
			this.label166.Size = new System.Drawing.Size(39, 17);
			this.label166.TabIndex = 1059;
			this.label166.Text = "Latitud";
			this.label166.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// txtVTLongitud
			// 
			appearance3.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtVTLongitud.Appearance = appearance3;
			this.txtVTLongitud.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtVTLongitud.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtVTLongitud.Enabled = false;
			this.txtVTLongitud.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.txtVTLongitud.Location = new System.Drawing.Point(384, 88);
			this.txtVTLongitud.Name = "txtVTLongitud";
			this.txtVTLongitud.Size = new System.Drawing.Size(208, 22);
			this.txtVTLongitud.TabIndex = 1058;
			// 
			// txtVTPersonaEntrevistada
			// 
			appearance4.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtVTPersonaEntrevistada.Appearance = appearance4;
			this.txtVTPersonaEntrevistada.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtVTPersonaEntrevistada.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtVTPersonaEntrevistada.Enabled = false;
			this.txtVTPersonaEntrevistada.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.txtVTPersonaEntrevistada.Location = new System.Drawing.Point(552, 136);
			this.txtVTPersonaEntrevistada.Name = "txtVTPersonaEntrevistada";
			this.txtVTPersonaEntrevistada.Size = new System.Drawing.Size(536, 22);
			this.txtVTPersonaEntrevistada.TabIndex = 1057;
			// 
			// txtVTReferencia
			// 
			appearance5.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtVTReferencia.Appearance = appearance5;
			this.txtVTReferencia.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtVTReferencia.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtVTReferencia.Enabled = false;
			this.txtVTReferencia.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.txtVTReferencia.Location = new System.Drawing.Point(8, 136);
			this.txtVTReferencia.Name = "txtVTReferencia";
			this.txtVTReferencia.Size = new System.Drawing.Size(528, 22);
			this.txtVTReferencia.TabIndex = 1056;
			// 
			// txtVTGps
			// 
			appearance6.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtVTGps.Appearance = appearance6;
			this.txtVTGps.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtVTGps.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtVTGps.Enabled = false;
			this.txtVTGps.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.txtVTGps.Location = new System.Drawing.Point(600, 88);
			this.txtVTGps.Name = "txtVTGps";
			this.txtVTGps.Size = new System.Drawing.Size(448, 22);
			this.txtVTGps.TabIndex = 1055;
			// 
			// txtVTLatitud
			// 
			appearance7.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtVTLatitud.Appearance = appearance7;
			this.txtVTLatitud.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtVTLatitud.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtVTLatitud.Enabled = false;
			this.txtVTLatitud.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.txtVTLatitud.Location = new System.Drawing.Point(168, 88);
			this.txtVTLatitud.Name = "txtVTLatitud";
			this.txtVTLatitud.Size = new System.Drawing.Size(208, 22);
			this.txtVTLatitud.TabIndex = 1054;
			// 
			// txtVTTelefono
			// 
			appearance8.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtVTTelefono.Appearance = appearance8;
			this.txtVTTelefono.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtVTTelefono.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtVTTelefono.Enabled = false;
			this.txtVTTelefono.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.txtVTTelefono.Location = new System.Drawing.Point(8, 88);
			this.txtVTTelefono.Name = "txtVTTelefono";
			this.txtVTTelefono.Size = new System.Drawing.Size(152, 22);
			this.txtVTTelefono.TabIndex = 1053;
			// 
			// label165
			// 
			this.label165.AutoSize = true;
			this.label165.BackColor = System.Drawing.Color.Transparent;
			this.label165.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label165.Location = new System.Drawing.Point(8, 64);
			this.label165.Name = "label165";
			this.label165.Size = new System.Drawing.Size(88, 17);
			this.label165.TabIndex = 1052;
			this.label165.Text = "Teléfono Laboral";
			this.label165.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// txtVTActividadLaboral
			// 
			appearance9.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtVTActividadLaboral.Appearance = appearance9;
			this.txtVTActividadLaboral.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtVTActividadLaboral.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtVTActividadLaboral.Enabled = false;
			this.txtVTActividadLaboral.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.txtVTActividadLaboral.Location = new System.Drawing.Point(792, 32);
			this.txtVTActividadLaboral.Name = "txtVTActividadLaboral";
			this.txtVTActividadLaboral.Size = new System.Drawing.Size(296, 22);
			this.txtVTActividadLaboral.TabIndex = 1051;
			// 
			// label164
			// 
			this.label164.AutoSize = true;
			this.label164.BackColor = System.Drawing.Color.Transparent;
			this.label164.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label164.Location = new System.Drawing.Point(608, 8);
			this.label164.Name = "label164";
			this.label164.Size = new System.Drawing.Size(103, 17);
			this.label164.TabIndex = 1050;
			this.label164.Text = "Ingresos Mensuales";
			this.label164.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// txtVTIngresos
			// 
			this.txtVTIngresos.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtVTIngresos.Enabled = false;
			this.txtVTIngresos.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.txtVTIngresos.Location = new System.Drawing.Point(608, 32);
			this.txtVTIngresos.MaxValue = 1000000;
			this.txtVTIngresos.MinValue = 0;
			this.txtVTIngresos.Name = "txtVTIngresos";
			this.txtVTIngresos.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtVTIngresos.PromptChar = ' ';
			this.txtVTIngresos.Size = new System.Drawing.Size(176, 22);
			this.txtVTIngresos.SpinButtonDisplayStyle = Infragistics.Win.ButtonDisplayStyle.Always;
			this.txtVTIngresos.TabIndex = 1049;
			// 
			// label163
			// 
			this.label163.AutoSize = true;
			this.label163.BackColor = System.Drawing.Color.Transparent;
			this.label163.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label163.Location = new System.Drawing.Point(448, 8);
			this.label163.Name = "label163";
			this.label163.Size = new System.Drawing.Size(135, 17);
			this.label163.TabIndex = 1026;
			this.label163.Text = "Tiempo de Trabajo (Años)";
			this.label163.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// txtVTTiempoYear
			// 
			this.txtVTTiempoYear.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtVTTiempoYear.Enabled = false;
			this.txtVTTiempoYear.Location = new System.Drawing.Point(448, 32);
			this.txtVTTiempoYear.Name = "txtVTTiempoYear";
			this.txtVTTiempoYear.PromptChar = ' ';
			this.txtVTTiempoYear.Size = new System.Drawing.Size(152, 22);
			this.txtVTTiempoYear.SpinButtonAlignment = Infragistics.Win.ButtonAlignment.Left;
			this.txtVTTiempoYear.TabIndex = 1025;
			// 
			// label162
			// 
			this.label162.AutoSize = true;
			this.label162.BackColor = System.Drawing.Color.Transparent;
			this.label162.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label162.Location = new System.Drawing.Point(288, 8);
			this.label162.Name = "label162";
			this.label162.Size = new System.Drawing.Size(142, 17);
			this.label162.TabIndex = 1024;
			this.label162.Text = "Tiempo de Trabajo (Meses)";
			this.label162.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// txtVTTiempoMeses
			// 
			this.txtVTTiempoMeses.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtVTTiempoMeses.Enabled = false;
			this.txtVTTiempoMeses.Location = new System.Drawing.Point(288, 32);
			this.txtVTTiempoMeses.Name = "txtVTTiempoMeses";
			this.txtVTTiempoMeses.PromptChar = ' ';
			this.txtVTTiempoMeses.Size = new System.Drawing.Size(152, 22);
			this.txtVTTiempoMeses.SpinButtonAlignment = Infragistics.Win.ButtonAlignment.Left;
			this.txtVTTiempoMeses.TabIndex = 1023;
			// 
			// label161
			// 
			this.label161.AutoSize = true;
			this.label161.BackColor = System.Drawing.Color.Transparent;
			this.label161.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label161.Location = new System.Drawing.Point(8, 0);
			this.label161.Name = "label161";
			this.label161.Size = new System.Drawing.Size(87, 17);
			this.label161.TabIndex = 1022;
			this.label161.Text = "Tipo de Vivienda";
			this.label161.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// opdtVTTipoTrabajo
			// 
			appearance10.ForeColorDisabled = System.Drawing.Color.Black;
			this.opdtVTTipoTrabajo.Appearance = appearance10;
			this.opdtVTTipoTrabajo.BorderStyle = Infragistics.Win.UIElementBorderStyle.Solid;
			this.opdtVTTipoTrabajo.CheckedIndex = 0;
			this.opdtVTTipoTrabajo.Enabled = false;
			this.opdtVTTipoTrabajo.ItemAppearance = appearance11;
			this.opdtVTTipoTrabajo.ItemOrigin = new System.Drawing.Point(10, 0);
			valueListItem1.DataValue = 1;
			valueListItem1.DisplayText = "Dependiente";
			valueListItem2.DataValue = 2;
			valueListItem2.DisplayText = "Independiente";
			valueListItem3.DataValue = 3;
			valueListItem3.DisplayText = "Informal";
			this.opdtVTTipoTrabajo.Items.Add(valueListItem1);
			this.opdtVTTipoTrabajo.Items.Add(valueListItem2);
			this.opdtVTTipoTrabajo.Items.Add(valueListItem3);
			this.opdtVTTipoTrabajo.ItemSpacingVertical = 10;
			this.opdtVTTipoTrabajo.Location = new System.Drawing.Point(8, 24);
			this.opdtVTTipoTrabajo.Name = "opdtVTTipoTrabajo";
			this.opdtVTTipoTrabajo.Size = new System.Drawing.Size(272, 32);
			this.opdtVTTipoTrabajo.TabIndex = 1021;
			this.opdtVTTipoTrabajo.Text = "Dependiente";
			// 
			// ultraGrid5
			// 
			this.ultraGrid5.Cursor = System.Windows.Forms.Cursors.Default;
			this.ultraGrid5.DataSource = this.ultraDataSource42;
			appearance12.BackColor = System.Drawing.Color.White;
			appearance12.ForegroundAlpha = Infragistics.Win.Alpha.UseAlphaLevel;
			this.ultraGrid5.DisplayLayout.Appearance = appearance12;
			ultraGridColumn1.CellActivation = Infragistics.Win.UltraWinGrid.Activation.NoEdit;
			ultraGridColumn1.Header.Caption = "Institución";
			ultraGridColumn1.Header.VisiblePosition = 0;
			ultraGridColumn1.Width = 300;
			ultraGridColumn2.CellActivation = Infragistics.Win.UltraWinGrid.Activation.NoEdit;
			ultraGridColumn2.Header.Caption = "Emisor";
			ultraGridColumn2.Header.VisiblePosition = 1;
			ultraGridColumn2.Width = 200;
			ultraGridColumn3.CellActivation = Infragistics.Win.UltraWinGrid.Activation.NoEdit;
			appearance13.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn3.CellAppearance = appearance13;
			ultraGridColumn3.Header.Caption = "Antigüedad";
			ultraGridColumn3.Header.VisiblePosition = 2;
			ultraGridColumn3.Width = 200;
			ultraGridColumn4.CellActivation = Infragistics.Win.UltraWinGrid.Activation.NoEdit;
			appearance14.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn4.CellAppearance = appearance14;
			ultraGridColumn4.Format = "#,##0.00";
			appearance15.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn4.FormulaErrorAppearance = appearance15;
			ultraGridColumn4.Header.Caption = "Cupo";
			ultraGridColumn4.Header.VisiblePosition = 3;
			ultraGridColumn4.Width = 200;
			ultraGridBand1.Columns.AddRange(new object[] {
																										 ultraGridColumn1,
																										 ultraGridColumn2,
																										 ultraGridColumn3,
																										 ultraGridColumn4});
			this.ultraGrid5.DisplayLayout.BandsSerializer.Add(ultraGridBand1);
			this.ultraGrid5.DisplayLayout.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.TemplateOnBottom;
			this.ultraGrid5.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.True;
			this.ultraGrid5.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.True;
			appearance16.BackColor = System.Drawing.Color.Transparent;
			this.ultraGrid5.DisplayLayout.Override.CardAreaAppearance = appearance16;
			appearance17.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance17.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance17.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance17.FontData.BoldAsString = "True";
			appearance17.FontData.Name = "Arial";
			appearance17.FontData.SizeInPoints = 8.5F;
			appearance17.ForeColor = System.Drawing.Color.White;
			appearance17.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.ultraGrid5.DisplayLayout.Override.HeaderAppearance = appearance17;
			appearance18.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance18.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance18.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ultraGrid5.DisplayLayout.Override.RowSelectorAppearance = appearance18;
			appearance19.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance19.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance19.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ultraGrid5.DisplayLayout.Override.SelectedRowAppearance = appearance19;
			this.ultraGrid5.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.ultraGrid5.Location = new System.Drawing.Point(0, 8);
			this.ultraGrid5.Name = "ultraGrid5";
			this.ultraGrid5.Size = new System.Drawing.Size(0, 0);
			this.ultraGrid5.TabIndex = 801;
			// 
			// ultraDataSource42
			// 
			ultraDataColumn3.DataType = typeof(int);
			ultraDataColumn4.DataType = typeof(System.Decimal);
			this.ultraDataSource42.Band.Columns.AddRange(new object[] {
																																	ultraDataColumn1,
																																	ultraDataColumn2,
																																	ultraDataColumn3,
																																	ultraDataColumn4});
			this.ultraDataSource42.CellDataRequested += new Infragistics.Win.UltraWinDataSource.CellDataRequestedEventHandler(this.ultraDataSource42_CellDataRequested);
			// 
			// ultraTabPageControl9
			// 
			this.ultraTabPageControl9.Controls.Add(this.btnImagenes);
			this.ultraTabPageControl9.Controls.Add(this.txtVDLatitud);
			this.ultraTabPageControl9.Controls.Add(this.txtVDLongitud);
			this.ultraTabPageControl9.Controls.Add(this.txtVDCCelArr);
			this.ultraTabPageControl9.Controls.Add(this.txtVDGPS);
			this.ultraTabPageControl9.Controls.Add(this.label160);
			this.ultraTabPageControl9.Controls.Add(this.label159);
			this.ultraTabPageControl9.Controls.Add(this.label158);
			this.ultraTabPageControl9.Controls.Add(this.label157);
			this.ultraTabPageControl9.Controls.Add(this.label156);
			this.ultraTabPageControl9.Controls.Add(this.label155);
			this.ultraTabPageControl9.Controls.Add(this.label154);
			this.ultraTabPageControl9.Controls.Add(this.txtVDSecundaria);
			this.ultraTabPageControl9.Controls.Add(this.txtVDPrincipal);
			this.ultraTabPageControl9.Controls.Add(this.txtVDVecino);
			this.ultraTabPageControl9.Controls.Add(this.txtVDObservacion);
			this.ultraTabPageControl9.Controls.Add(this.txtVDentrevistada);
			this.ultraTabPageControl9.Controls.Add(this.txtVDPuntoReferencia);
			this.ultraTabPageControl9.Controls.Add(this.label153);
			this.ultraTabPageControl9.Controls.Add(this.opdtVDCobertura);
			this.ultraTabPageControl9.Controls.Add(this.label152);
			this.ultraTabPageControl9.Controls.Add(this.label151);
			this.ultraTabPageControl9.Controls.Add(this.label150);
			this.ultraTabPageControl9.Controls.Add(this.opdtVDAcceso);
			this.ultraTabPageControl9.Controls.Add(this.opdtVDPropiedad);
			this.ultraTabPageControl9.Controls.Add(this.opdtVDZonaVivienda);
			this.ultraTabPageControl9.Controls.Add(this.label149);
			this.ultraTabPageControl9.Controls.Add(this.opdtVDEstadoVivienda);
			this.ultraTabPageControl9.Controls.Add(this.label148);
			this.ultraTabPageControl9.Controls.Add(this.opdtVDVivienda);
			this.ultraTabPageControl9.Controls.Add(this.label147);
			this.ultraTabPageControl9.Controls.Add(this.txtVDTiempoMeses);
			this.ultraTabPageControl9.Location = new System.Drawing.Point(-10000, -10000);
			this.ultraTabPageControl9.Name = "ultraTabPageControl9";
			this.ultraTabPageControl9.Size = new System.Drawing.Size(1092, 198);
			// 
			// btnImagenes
			// 
			this.btnImagenes.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			this.btnImagenes.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnImagenes.Location = new System.Drawing.Point(8, 72);
			this.btnImagenes.Name = "btnImagenes";
			this.btnImagenes.Size = new System.Drawing.Size(80, 23);
			this.btnImagenes.TabIndex = 1053;
			this.btnImagenes.Text = "Imagenes";
			this.btnImagenes.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// txtVDLatitud
			// 
			appearance20.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtVDLatitud.Appearance = appearance20;
			this.txtVDLatitud.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtVDLatitud.Enabled = false;
			this.txtVDLatitud.Location = new System.Drawing.Point(824, 168);
			this.txtVDLatitud.MaxLength = 10;
			this.txtVDLatitud.Name = "txtVDLatitud";
			this.txtVDLatitud.Size = new System.Drawing.Size(264, 22);
			this.txtVDLatitud.TabIndex = 1050;
			// 
			// txtVDLongitud
			// 
			appearance21.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtVDLongitud.Appearance = appearance21;
			this.txtVDLongitud.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtVDLongitud.Enabled = false;
			this.txtVDLongitud.Location = new System.Drawing.Point(824, 144);
			this.txtVDLongitud.MaxLength = 10;
			this.txtVDLongitud.Name = "txtVDLongitud";
			this.txtVDLongitud.Size = new System.Drawing.Size(264, 22);
			this.txtVDLongitud.TabIndex = 1049;
			// 
			// txtVDCCelArr
			// 
			this.txtVDCCelArr.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtVDCCelArr.Enabled = false;
			this.txtVDCCelArr.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.txtVDCCelArr.Location = new System.Drawing.Point(904, 56);
			this.txtVDCCelArr.MaxValue = 1000000;
			this.txtVDCCelArr.MinValue = 0;
			this.txtVDCCelArr.Name = "txtVDCCelArr";
			this.txtVDCCelArr.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtVDCCelArr.PromptChar = ' ';
			this.txtVDCCelArr.Size = new System.Drawing.Size(176, 22);
			this.txtVDCCelArr.SpinButtonDisplayStyle = Infragistics.Win.ButtonDisplayStyle.Always;
			this.txtVDCCelArr.TabIndex = 1048;
			// 
			// txtVDGPS
			// 
			appearance22.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtVDGPS.Appearance = appearance22;
			this.txtVDGPS.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtVDGPS.Enabled = false;
			this.txtVDGPS.Location = new System.Drawing.Point(824, 120);
			this.txtVDGPS.MaxLength = 10;
			this.txtVDGPS.Name = "txtVDGPS";
			this.txtVDGPS.Size = new System.Drawing.Size(264, 22);
			this.txtVDGPS.TabIndex = 1045;
			// 
			// label160
			// 
			this.label160.AutoSize = true;
			this.label160.BackColor = System.Drawing.Color.Transparent;
			this.label160.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label160.Location = new System.Drawing.Point(904, 96);
			this.label160.Name = "label160";
			this.label160.Size = new System.Drawing.Size(111, 17);
			this.label160.TabIndex = 1044;
			this.label160.Text = "GPS/Latitud/Longitud";
			this.label160.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// label159
			// 
			this.label159.AutoSize = true;
			this.label159.BackColor = System.Drawing.Color.Transparent;
			this.label159.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label159.Location = new System.Drawing.Point(688, 104);
			this.label159.Name = "label159";
			this.label159.Size = new System.Drawing.Size(87, 17);
			this.label159.TabIndex = 1043;
			this.label159.Text = "Calle Secundaria";
			this.label159.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// label158
			// 
			this.label158.AutoSize = true;
			this.label158.BackColor = System.Drawing.Color.Transparent;
			this.label158.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label158.Location = new System.Drawing.Point(552, 104);
			this.label158.Name = "label158";
			this.label158.Size = new System.Drawing.Size(73, 17);
			this.label158.TabIndex = 1042;
			this.label158.Text = "Calle Principal";
			this.label158.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// label157
			// 
			this.label157.AutoSize = true;
			this.label157.BackColor = System.Drawing.Color.Transparent;
			this.label157.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label157.Location = new System.Drawing.Point(416, 104);
			this.label157.Name = "label157";
			this.label157.Size = new System.Drawing.Size(103, 17);
			this.label157.TabIndex = 1041;
			this.label157.Text = "Vecino Entrevistado";
			this.label157.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// label156
			// 
			this.label156.AutoSize = true;
			this.label156.BackColor = System.Drawing.Color.Transparent;
			this.label156.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label156.Location = new System.Drawing.Point(280, 104);
			this.label156.Name = "label156";
			this.label156.Size = new System.Drawing.Size(65, 17);
			this.label156.TabIndex = 1040;
			this.label156.Text = "Observación";
			this.label156.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// label155
			// 
			this.label155.AutoSize = true;
			this.label155.BackColor = System.Drawing.Color.Transparent;
			this.label155.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label155.Location = new System.Drawing.Point(144, 104);
			this.label155.Name = "label155";
			this.label155.Size = new System.Drawing.Size(109, 17);
			this.label155.TabIndex = 1039;
			this.label155.Text = "Persona Entrevistada";
			this.label155.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// label154
			// 
			this.label154.AutoSize = true;
			this.label154.BackColor = System.Drawing.Color.Transparent;
			this.label154.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label154.Location = new System.Drawing.Point(8, 104);
			this.label154.Name = "label154";
			this.label154.Size = new System.Drawing.Size(105, 17);
			this.label154.TabIndex = 1038;
			this.label154.Text = "Punto de Referencia";
			this.label154.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// txtVDSecundaria
			// 
			this.txtVDSecundaria.AcceptsReturn = true;
			appearance23.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtVDSecundaria.Appearance = appearance23;
			this.txtVDSecundaria.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtVDSecundaria.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtVDSecundaria.Enabled = false;
			this.txtVDSecundaria.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.txtVDSecundaria.Location = new System.Drawing.Point(688, 128);
			this.txtVDSecundaria.Multiline = true;
			this.txtVDSecundaria.Name = "txtVDSecundaria";
			this.txtVDSecundaria.Size = new System.Drawing.Size(128, 64);
			this.txtVDSecundaria.TabIndex = 1037;
			// 
			// txtVDPrincipal
			// 
			this.txtVDPrincipal.AcceptsReturn = true;
			appearance24.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtVDPrincipal.Appearance = appearance24;
			this.txtVDPrincipal.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtVDPrincipal.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtVDPrincipal.Enabled = false;
			this.txtVDPrincipal.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.txtVDPrincipal.Location = new System.Drawing.Point(552, 128);
			this.txtVDPrincipal.Multiline = true;
			this.txtVDPrincipal.Name = "txtVDPrincipal";
			this.txtVDPrincipal.Size = new System.Drawing.Size(128, 64);
			this.txtVDPrincipal.TabIndex = 1036;
			// 
			// txtVDVecino
			// 
			this.txtVDVecino.AcceptsReturn = true;
			appearance25.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtVDVecino.Appearance = appearance25;
			this.txtVDVecino.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtVDVecino.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtVDVecino.Enabled = false;
			this.txtVDVecino.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.txtVDVecino.Location = new System.Drawing.Point(416, 128);
			this.txtVDVecino.Multiline = true;
			this.txtVDVecino.Name = "txtVDVecino";
			this.txtVDVecino.Size = new System.Drawing.Size(128, 64);
			this.txtVDVecino.TabIndex = 1035;
			// 
			// txtVDObservacion
			// 
			this.txtVDObservacion.AcceptsReturn = true;
			appearance26.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtVDObservacion.Appearance = appearance26;
			this.txtVDObservacion.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtVDObservacion.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtVDObservacion.Enabled = false;
			this.txtVDObservacion.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.txtVDObservacion.Location = new System.Drawing.Point(280, 128);
			this.txtVDObservacion.Multiline = true;
			this.txtVDObservacion.Name = "txtVDObservacion";
			this.txtVDObservacion.Size = new System.Drawing.Size(128, 64);
			this.txtVDObservacion.TabIndex = 1034;
			// 
			// txtVDentrevistada
			// 
			this.txtVDentrevistada.AcceptsReturn = true;
			appearance27.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtVDentrevistada.Appearance = appearance27;
			this.txtVDentrevistada.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtVDentrevistada.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtVDentrevistada.Enabled = false;
			this.txtVDentrevistada.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.txtVDentrevistada.Location = new System.Drawing.Point(144, 128);
			this.txtVDentrevistada.Multiline = true;
			this.txtVDentrevistada.Name = "txtVDentrevistada";
			this.txtVDentrevistada.Size = new System.Drawing.Size(128, 64);
			this.txtVDentrevistada.TabIndex = 1033;
			// 
			// txtVDPuntoReferencia
			// 
			this.txtVDPuntoReferencia.AcceptsReturn = true;
			appearance28.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtVDPuntoReferencia.Appearance = appearance28;
			this.txtVDPuntoReferencia.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtVDPuntoReferencia.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtVDPuntoReferencia.Enabled = false;
			this.txtVDPuntoReferencia.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.txtVDPuntoReferencia.Location = new System.Drawing.Point(8, 128);
			this.txtVDPuntoReferencia.Multiline = true;
			this.txtVDPuntoReferencia.Name = "txtVDPuntoReferencia";
			this.txtVDPuntoReferencia.Size = new System.Drawing.Size(128, 64);
			this.txtVDPuntoReferencia.TabIndex = 1032;
			// 
			// label153
			// 
			this.label153.AutoSize = true;
			this.label153.BackColor = System.Drawing.Color.Transparent;
			this.label153.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label153.Location = new System.Drawing.Point(904, 0);
			this.label153.Name = "label153";
			this.label153.TabIndex = 1030;
			this.label153.Text = "Cobertura de Señal";
			this.label153.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// opdtVDCobertura
			// 
			appearance29.ForeColorDisabled = System.Drawing.Color.Black;
			this.opdtVDCobertura.Appearance = appearance29;
			this.opdtVDCobertura.BorderStyle = Infragistics.Win.UIElementBorderStyle.Solid;
			this.opdtVDCobertura.CheckedIndex = 0;
			this.opdtVDCobertura.Enabled = false;
			this.opdtVDCobertura.ItemAppearance = appearance30;
			this.opdtVDCobertura.ItemOrigin = new System.Drawing.Point(10, 0);
			valueListItem4.DataValue = 1;
			valueListItem4.DisplayText = "Llamada Movil";
			valueListItem5.DataValue = 2;
			valueListItem5.DisplayText = "WhatsApp";
			this.opdtVDCobertura.Items.Add(valueListItem4);
			this.opdtVDCobertura.Items.Add(valueListItem5);
			this.opdtVDCobertura.ItemSpacingVertical = 10;
			this.opdtVDCobertura.Location = new System.Drawing.Point(904, 24);
			this.opdtVDCobertura.Name = "opdtVDCobertura";
			this.opdtVDCobertura.Size = new System.Drawing.Size(184, 24);
			this.opdtVDCobertura.TabIndex = 1029;
			this.opdtVDCobertura.Text = "Llamada Movil";
			// 
			// label152
			// 
			this.label152.AutoSize = true;
			this.label152.BackColor = System.Drawing.Color.Transparent;
			this.label152.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label152.Location = new System.Drawing.Point(840, 16);
			this.label152.Name = "label152";
			this.label152.Size = new System.Drawing.Size(39, 17);
			this.label152.TabIndex = 1028;
			this.label152.Text = "Acceso";
			this.label152.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// label151
			// 
			this.label151.AutoSize = true;
			this.label151.BackColor = System.Drawing.Color.Transparent;
			this.label151.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label151.Location = new System.Drawing.Point(680, 16);
			this.label151.Name = "label151";
			this.label151.Size = new System.Drawing.Size(54, 17);
			this.label151.TabIndex = 1027;
			this.label151.Text = "Propiedad";
			this.label151.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// label150
			// 
			this.label150.AutoSize = true;
			this.label150.BackColor = System.Drawing.Color.Transparent;
			this.label150.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label150.Location = new System.Drawing.Point(592, 16);
			this.label150.Name = "label150";
			this.label150.Size = new System.Drawing.Size(75, 17);
			this.label150.TabIndex = 1026;
			this.label150.Text = "Zona Vivienda";
			this.label150.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// opdtVDAcceso
			// 
			appearance31.ForeColorDisabled = System.Drawing.Color.Black;
			this.opdtVDAcceso.Appearance = appearance31;
			this.opdtVDAcceso.BorderStyle = Infragistics.Win.UIElementBorderStyle.Solid;
			this.opdtVDAcceso.CheckedIndex = 0;
			this.opdtVDAcceso.Enabled = false;
			this.opdtVDAcceso.ItemAppearance = appearance32;
			this.opdtVDAcceso.ItemOrigin = new System.Drawing.Point(10, 0);
			valueListItem6.DataValue = 1;
			valueListItem6.DisplayText = "Facil";
			valueListItem7.DataValue = 3;
			valueListItem7.DisplayText = "Dificil";
			this.opdtVDAcceso.Items.Add(valueListItem6);
			this.opdtVDAcceso.Items.Add(valueListItem7);
			this.opdtVDAcceso.ItemSpacingVertical = 10;
			this.opdtVDAcceso.Location = new System.Drawing.Point(832, 40);
			this.opdtVDAcceso.Name = "opdtVDAcceso";
			this.opdtVDAcceso.Size = new System.Drawing.Size(64, 56);
			this.opdtVDAcceso.TabIndex = 1025;
			this.opdtVDAcceso.Text = "Facil";
			// 
			// opdtVDPropiedad
			// 
			appearance33.ForeColorDisabled = System.Drawing.Color.Black;
			this.opdtVDPropiedad.Appearance = appearance33;
			this.opdtVDPropiedad.BorderStyle = Infragistics.Win.UIElementBorderStyle.Solid;
			this.opdtVDPropiedad.CheckedIndex = 0;
			this.opdtVDPropiedad.Enabled = false;
			this.opdtVDPropiedad.ItemAppearance = appearance34;
			this.opdtVDPropiedad.ItemOrigin = new System.Drawing.Point(10, 0);
			valueListItem8.DataValue = 1;
			valueListItem8.DisplayText = "Propio";
			valueListItem9.DataValue = 3;
			valueListItem9.DisplayText = "Familiar";
			valueListItem10.DataValue = 2;
			valueListItem10.DisplayText = "Arrendado";
			this.opdtVDPropiedad.Items.Add(valueListItem8);
			this.opdtVDPropiedad.Items.Add(valueListItem9);
			this.opdtVDPropiedad.Items.Add(valueListItem10);
			this.opdtVDPropiedad.ItemSpacingVertical = 10;
			this.opdtVDPropiedad.Location = new System.Drawing.Point(664, 40);
			this.opdtVDPropiedad.Name = "opdtVDPropiedad";
			this.opdtVDPropiedad.Size = new System.Drawing.Size(160, 56);
			this.opdtVDPropiedad.TabIndex = 1024;
			this.opdtVDPropiedad.Text = "Propio";
			// 
			// opdtVDZonaVivienda
			// 
			appearance35.ForeColorDisabled = System.Drawing.Color.Black;
			this.opdtVDZonaVivienda.Appearance = appearance35;
			this.opdtVDZonaVivienda.BorderStyle = Infragistics.Win.UIElementBorderStyle.Solid;
			this.opdtVDZonaVivienda.CheckedIndex = 0;
			this.opdtVDZonaVivienda.Enabled = false;
			this.opdtVDZonaVivienda.ItemAppearance = appearance36;
			this.opdtVDZonaVivienda.ItemOrigin = new System.Drawing.Point(10, 0);
			valueListItem11.DataValue = 1;
			valueListItem11.DisplayText = "Urbano";
			valueListItem12.DataValue = 2;
			valueListItem12.DisplayText = "Rural";
			this.opdtVDZonaVivienda.Items.Add(valueListItem11);
			this.opdtVDZonaVivienda.Items.Add(valueListItem12);
			this.opdtVDZonaVivienda.ItemSpacingVertical = 10;
			this.opdtVDZonaVivienda.Location = new System.Drawing.Point(584, 40);
			this.opdtVDZonaVivienda.Name = "opdtVDZonaVivienda";
			this.opdtVDZonaVivienda.Size = new System.Drawing.Size(72, 56);
			this.opdtVDZonaVivienda.TabIndex = 1023;
			this.opdtVDZonaVivienda.Text = "Urbano";
			// 
			// label149
			// 
			this.label149.AutoSize = true;
			this.label149.BackColor = System.Drawing.Color.Transparent;
			this.label149.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label149.Location = new System.Drawing.Point(432, 16);
			this.label149.Name = "label149";
			this.label149.Size = new System.Drawing.Size(83, 17);
			this.label149.TabIndex = 1022;
			this.label149.Text = "Estado Vivienda";
			this.label149.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// opdtVDEstadoVivienda
			// 
			appearance37.ForeColorDisabled = System.Drawing.Color.Black;
			this.opdtVDEstadoVivienda.Appearance = appearance37;
			this.opdtVDEstadoVivienda.BorderStyle = Infragistics.Win.UIElementBorderStyle.Solid;
			this.opdtVDEstadoVivienda.CheckedIndex = 0;
			this.opdtVDEstadoVivienda.Enabled = false;
			this.opdtVDEstadoVivienda.ItemAppearance = appearance38;
			this.opdtVDEstadoVivienda.ItemOrigin = new System.Drawing.Point(10, 0);
			valueListItem13.DataValue = 1;
			valueListItem13.DisplayText = "Bueno";
			valueListItem14.DataValue = 3;
			valueListItem14.DisplayText = "Malo";
			valueListItem15.DataValue = 2;
			valueListItem15.DisplayText = "Muy Bueno";
			this.opdtVDEstadoVivienda.Items.Add(valueListItem13);
			this.opdtVDEstadoVivienda.Items.Add(valueListItem14);
			this.opdtVDEstadoVivienda.Items.Add(valueListItem15);
			this.opdtVDEstadoVivienda.ItemSpacingVertical = 10;
			this.opdtVDEstadoVivienda.Location = new System.Drawing.Point(424, 40);
			this.opdtVDEstadoVivienda.Name = "opdtVDEstadoVivienda";
			this.opdtVDEstadoVivienda.Size = new System.Drawing.Size(152, 56);
			this.opdtVDEstadoVivienda.TabIndex = 1021;
			this.opdtVDEstadoVivienda.Text = "Bueno";
			// 
			// label148
			// 
			this.label148.AutoSize = true;
			this.label148.BackColor = System.Drawing.Color.Transparent;
			this.label148.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label148.Location = new System.Drawing.Point(168, 16);
			this.label148.Name = "label148";
			this.label148.Size = new System.Drawing.Size(87, 17);
			this.label148.TabIndex = 1020;
			this.label148.Text = "Tipo de Vivienda";
			this.label148.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// opdtVDVivienda
			// 
			appearance39.ForeColorDisabled = System.Drawing.Color.Black;
			this.opdtVDVivienda.Appearance = appearance39;
			this.opdtVDVivienda.BorderStyle = Infragistics.Win.UIElementBorderStyle.Solid;
			this.opdtVDVivienda.CheckedIndex = 0;
			this.opdtVDVivienda.Enabled = false;
			this.opdtVDVivienda.ItemAppearance = appearance40;
			this.opdtVDVivienda.ItemOrigin = new System.Drawing.Point(10, 0);
			valueListItem16.DataValue = 1;
			valueListItem16.DisplayText = "Casa";
			valueListItem17.DataValue = 3;
			valueListItem17.DisplayText = "Villa";
			valueListItem18.DataValue = 4;
			valueListItem18.DisplayText = "Mixta";
			valueListItem19.DataValue = 2;
			valueListItem19.DisplayText = "Departamento";
			valueListItem20.DataValue = 5;
			valueListItem20.DisplayText = "Media Agua";
			this.opdtVDVivienda.Items.Add(valueListItem16);
			this.opdtVDVivienda.Items.Add(valueListItem17);
			this.opdtVDVivienda.Items.Add(valueListItem18);
			this.opdtVDVivienda.Items.Add(valueListItem19);
			this.opdtVDVivienda.Items.Add(valueListItem20);
			this.opdtVDVivienda.ItemSpacingVertical = 10;
			this.opdtVDVivienda.Location = new System.Drawing.Point(168, 40);
			this.opdtVDVivienda.Name = "opdtVDVivienda";
			this.opdtVDVivienda.Size = new System.Drawing.Size(248, 56);
			this.opdtVDVivienda.TabIndex = 1019;
			this.opdtVDVivienda.Text = "Casa";
			// 
			// label147
			// 
			this.label147.AutoSize = true;
			this.label147.BackColor = System.Drawing.Color.Transparent;
			this.label147.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label147.Location = new System.Drawing.Point(8, 16);
			this.label147.Name = "label147";
			this.label147.Size = new System.Drawing.Size(146, 17);
			this.label147.TabIndex = 829;
			this.label147.Text = "Tiempo de Vivienda (Meses)";
			this.label147.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// txtVDTiempoMeses
			// 
			this.txtVDTiempoMeses.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtVDTiempoMeses.Enabled = false;
			this.txtVDTiempoMeses.Location = new System.Drawing.Point(8, 40);
			this.txtVDTiempoMeses.Name = "txtVDTiempoMeses";
			this.txtVDTiempoMeses.PromptChar = ' ';
			this.txtVDTiempoMeses.Size = new System.Drawing.Size(152, 22);
			this.txtVDTiempoMeses.SpinButtonAlignment = Infragistics.Win.ButtonAlignment.Left;
			this.txtVDTiempoMeses.TabIndex = 828;
			// 
			// ultraTabPageControl1
			// 
			this.ultraTabPageControl1.Controls.Add(this.groupBox2);
			this.ultraTabPageControl1.Controls.Add(this.groupBox1);
			this.ultraTabPageControl1.Location = new System.Drawing.Point(1, 23);
			this.ultraTabPageControl1.Name = "ultraTabPageControl1";
			this.ultraTabPageControl1.Size = new System.Drawing.Size(1116, 366);
			this.ultraTabPageControl1.Paint += new System.Windows.Forms.PaintEventHandler(this.ultraTabPageControl1_Paint);
			// 
			// groupBox2
			// 
			this.groupBox2.BackColor = System.Drawing.SystemColors.Control;
			this.groupBox2.Controls.Add(this.label144);
			this.groupBox2.Controls.Add(this.txtObsActividadEconomica);
			this.groupBox2.Controls.Add(this.cmbActividadEconomica);
			this.groupBox2.Controls.Add(this.cmbSituacionLaboral);
			this.groupBox2.Controls.Add(this.cmbProfesion);
			this.groupBox2.Controls.Add(this.cmbNivelEducacion);
			this.groupBox2.Controls.Add(this.cmbSexo);
			this.groupBox2.Controls.Add(this.label21);
			this.groupBox2.Controls.Add(this.label22);
			this.groupBox2.Controls.Add(this.label23);
			this.groupBox2.Controls.Add(this.chkHipoteca);
			this.groupBox2.Controls.Add(this.chkSeparacionBienes);
			this.groupBox2.Controls.Add(this.txtValorVehiculo);
			this.groupBox2.Controls.Add(this.txtModelo);
			this.groupBox2.Controls.Add(this.txtVehiculoMarca);
			this.groupBox2.Controls.Add(this.label24);
			this.groupBox2.Controls.Add(this.label25);
			this.groupBox2.Controls.Add(this.label26);
			this.groupBox2.Controls.Add(this.label27);
			this.groupBox2.Controls.Add(this.label28);
			this.groupBox2.Controls.Add(this.txtDependientes);
			this.groupBox2.Controls.Add(this.label29);
			this.groupBox2.Controls.Add(this.cmbEstadoCivil);
			this.groupBox2.Controls.Add(this.label30);
			this.groupBox2.Controls.Add(this.label31);
			this.groupBox2.Controls.Add(this.label32);
			this.groupBox2.Controls.Add(this.label33);
			this.groupBox2.Controls.Add(this.cmbCantonNacimiento);
			this.groupBox2.Controls.Add(this.cmbProvinciaNAcimiento);
			this.groupBox2.Controls.Add(this.dtFechaNacimiento);
			this.groupBox2.Controls.Add(this.cmbNacionalidad);
			this.groupBox2.Controls.Add(this.label34);
			this.groupBox2.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.groupBox2.Location = new System.Drawing.Point(0, 0);
			this.groupBox2.Name = "groupBox2";
			this.groupBox2.Size = new System.Drawing.Size(1112, 144);
			this.groupBox2.TabIndex = 197;
			this.groupBox2.TabStop = false;
			this.groupBox2.Text = "Datos Personales";
			this.groupBox2.Enter += new System.EventHandler(this.groupBox2_Enter);
			// 
			// label144
			// 
			this.label144.AutoSize = true;
			this.label144.Location = new System.Drawing.Point(752, 96);
			this.label144.Name = "label144";
			this.label144.Size = new System.Drawing.Size(171, 17);
			this.label144.TabIndex = 820;
			this.label144.Text = "Observación Actividad Económica";
			// 
			// txtObsActividadEconomica
			// 
			appearance41.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtObsActividadEconomica.Appearance = appearance41;
			this.txtObsActividadEconomica.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtObsActividadEconomica.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtObsActividadEconomica.Enabled = false;
			this.txtObsActividadEconomica.Location = new System.Drawing.Point(752, 112);
			this.txtObsActividadEconomica.MaxLength = 100;
			this.txtObsActividadEconomica.Name = "txtObsActividadEconomica";
			this.txtObsActividadEconomica.Size = new System.Drawing.Size(344, 22);
			this.txtObsActividadEconomica.TabIndex = 819;
			// 
			// cmbActividadEconomica
			// 
			this.cmbActividadEconomica.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbActividadEconomica.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbActividadEconomica.DataSource = this.ultraDataSource40;
			ultraGridBand2.AddButtonCaption = "RolSucursal";
			ultraGridColumn5.Header.VisiblePosition = 0;
			ultraGridColumn5.Hidden = true;
			ultraGridColumn6.Header.VisiblePosition = 1;
			ultraGridColumn6.Width = 350;
			ultraGridBand2.Columns.AddRange(new object[] {
																										 ultraGridColumn5,
																										 ultraGridColumn6});
			this.cmbActividadEconomica.DisplayLayout.BandsSerializer.Add(ultraGridBand2);
			this.cmbActividadEconomica.DisplayMember = "Nombre";
			this.cmbActividadEconomica.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbActividadEconomica.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbActividadEconomica.Enabled = false;
			this.cmbActividadEconomica.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbActividadEconomica.Location = new System.Drawing.Point(752, 72);
			this.cmbActividadEconomica.Name = "cmbActividadEconomica";
			this.cmbActividadEconomica.Size = new System.Drawing.Size(350, 22);
			this.cmbActividadEconomica.TabIndex = 818;
			this.cmbActividadEconomica.ValueMember = "idActEconomica";
			this.cmbActividadEconomica.KeyDown += new System.Windows.Forms.KeyEventHandler(this.cmbActividadEconomica_KeyDown_1);
			// 
			// ultraDataSource40
			// 
			ultraDataColumn5.DataType = typeof(int);
			this.ultraDataSource40.Band.Columns.AddRange(new object[] {
																																	ultraDataColumn5,
																																	ultraDataColumn6});
			// 
			// cmbSituacionLaboral
			// 
			this.cmbSituacionLaboral.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbSituacionLaboral.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbSituacionLaboral.DataSource = this.ultraDataSource30;
			ultraGridColumn7.Header.VisiblePosition = 0;
			ultraGridColumn7.Hidden = true;
			ultraGridColumn8.Header.VisiblePosition = 1;
			ultraGridColumn8.Width = 150;
			ultraGridBand3.Columns.AddRange(new object[] {
																										 ultraGridColumn7,
																										 ultraGridColumn8});
			this.cmbSituacionLaboral.DisplayLayout.BandsSerializer.Add(ultraGridBand3);
			this.cmbSituacionLaboral.DisplayMember = "Descripcion";
			this.cmbSituacionLaboral.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbSituacionLaboral.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbSituacionLaboral.Enabled = false;
			this.cmbSituacionLaboral.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbSituacionLaboral.Location = new System.Drawing.Point(592, 72);
			this.cmbSituacionLaboral.Name = "cmbSituacionLaboral";
			this.cmbSituacionLaboral.Size = new System.Drawing.Size(150, 22);
			this.cmbSituacionLaboral.TabIndex = 816;
			this.cmbSituacionLaboral.ValueMember = "idSituacionLaboral";
			this.cmbSituacionLaboral.KeyDown += new System.Windows.Forms.KeyEventHandler(this.cmbSituacionLaboral_KeyDown);
			this.cmbSituacionLaboral.ValueChanged += new System.EventHandler(this.cmbSituacionLaboral_ValueChanged);
			// 
			// ultraDataSource30
			// 
			ultraDataColumn7.DataType = typeof(int);
			this.ultraDataSource30.Band.Columns.AddRange(new object[] {
																																	ultraDataColumn7,
																																	ultraDataColumn8});
			// 
			// cmbProfesion
			// 
			this.cmbProfesion.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbProfesion.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbProfesion.DataSource = this.ultraDataSource29;
			ultraGridColumn9.Header.VisiblePosition = 0;
			ultraGridColumn9.Hidden = true;
			ultraGridColumn10.Header.VisiblePosition = 1;
			ultraGridColumn10.Width = 150;
			ultraGridBand4.Columns.AddRange(new object[] {
																										 ultraGridColumn9,
																										 ultraGridColumn10});
			this.cmbProfesion.DisplayLayout.BandsSerializer.Add(ultraGridBand4);
			this.cmbProfesion.DisplayMember = "Nombre";
			this.cmbProfesion.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbProfesion.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbProfesion.Enabled = false;
			this.cmbProfesion.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbProfesion.Location = new System.Drawing.Point(432, 72);
			this.cmbProfesion.Name = "cmbProfesion";
			this.cmbProfesion.Size = new System.Drawing.Size(150, 22);
			this.cmbProfesion.TabIndex = 815;
			this.cmbProfesion.ValueMember = "idProfesion";
			this.cmbProfesion.KeyDown += new System.Windows.Forms.KeyEventHandler(this.cmbProfesion_KeyDown);
			// 
			// ultraDataSource29
			// 
			ultraDataColumn9.DataType = typeof(System.UInt32);
			this.ultraDataSource29.Band.Columns.AddRange(new object[] {
																																	ultraDataColumn9,
																																	ultraDataColumn10});
			// 
			// cmbNivelEducacion
			// 
			this.cmbNivelEducacion.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbNivelEducacion.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbNivelEducacion.DataSource = this.ultraDataSource28;
			ultraGridColumn11.Header.VisiblePosition = 0;
			ultraGridColumn11.Hidden = true;
			ultraGridColumn12.Header.VisiblePosition = 1;
			ultraGridColumn12.Width = 150;
			ultraGridBand5.Columns.AddRange(new object[] {
																										 ultraGridColumn11,
																										 ultraGridColumn12});
			this.cmbNivelEducacion.DisplayLayout.BandsSerializer.Add(ultraGridBand5);
			this.cmbNivelEducacion.DisplayMember = "Nombre";
			this.cmbNivelEducacion.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbNivelEducacion.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbNivelEducacion.Enabled = false;
			this.cmbNivelEducacion.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbNivelEducacion.Location = new System.Drawing.Point(272, 72);
			this.cmbNivelEducacion.Name = "cmbNivelEducacion";
			this.cmbNivelEducacion.Size = new System.Drawing.Size(150, 22);
			this.cmbNivelEducacion.TabIndex = 811;
			this.cmbNivelEducacion.ValueMember = "idNivelEducacion";
			this.cmbNivelEducacion.KeyDown += new System.Windows.Forms.KeyEventHandler(this.cmbNivelEducacion_KeyDown);
			// 
			// ultraDataSource28
			// 
			ultraDataColumn11.DataType = typeof(int);
			this.ultraDataSource28.Band.Columns.AddRange(new object[] {
																																	ultraDataColumn11,
																																	ultraDataColumn12});
			// 
			// cmbSexo
			// 
			this.cmbSexo.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbSexo.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbSexo.DataSource = this.ultraDataSource24;
			ultraGridColumn13.Header.VisiblePosition = 0;
			ultraGridColumn13.Hidden = true;
			ultraGridColumn14.Header.Caption = "Sexo";
			ultraGridColumn14.Header.VisiblePosition = 1;
			ultraGridColumn14.Width = 150;
			ultraGridBand6.Columns.AddRange(new object[] {
																										 ultraGridColumn13,
																										 ultraGridColumn14});
			this.cmbSexo.DisplayLayout.BandsSerializer.Add(ultraGridBand6);
			this.cmbSexo.DisplayMember = "Nombre";
			this.cmbSexo.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbSexo.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbSexo.Enabled = false;
			this.cmbSexo.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbSexo.Location = new System.Drawing.Point(272, 32);
			this.cmbSexo.Name = "cmbSexo";
			this.cmbSexo.Size = new System.Drawing.Size(150, 22);
			this.cmbSexo.TabIndex = 810;
			this.cmbSexo.ValueMember = "idSexo";
			this.cmbSexo.KeyDown += new System.Windows.Forms.KeyEventHandler(this.cmbSexo_KeyDown);
			this.cmbSexo.InitializeLayout += new Infragistics.Win.UltraWinGrid.InitializeLayoutEventHandler(this.cmbSexo_InitializeLayout);
			// 
			// ultraDataSource24
			// 
			ultraDataColumn13.DataType = typeof(System.UInt32);
			this.ultraDataSource24.Band.Columns.AddRange(new object[] {
																																	ultraDataColumn13,
																																	ultraDataColumn14});
			// 
			// label21
			// 
			this.label21.AutoSize = true;
			this.label21.BackColor = System.Drawing.Color.Transparent;
			this.label21.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label21.Location = new System.Drawing.Point(304, 96);
			this.label21.Name = "label21";
			this.label21.Size = new System.Drawing.Size(75, 17);
			this.label21.TabIndex = 809;
			this.label21.Text = "Valor Vehículo";
			this.label21.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.label21.Visible = false;
			// 
			// label22
			// 
			this.label22.AutoSize = true;
			this.label22.BackColor = System.Drawing.Color.Transparent;
			this.label22.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label22.Location = new System.Drawing.Point(168, 96);
			this.label22.Name = "label22";
			this.label22.Size = new System.Drawing.Size(40, 17);
			this.label22.TabIndex = 808;
			this.label22.Text = "Modelo";
			this.label22.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.label22.Visible = false;
			// 
			// label23
			// 
			this.label23.AutoSize = true;
			this.label23.BackColor = System.Drawing.Color.Transparent;
			this.label23.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label23.Location = new System.Drawing.Point(8, 96);
			this.label23.Name = "label23";
			this.label23.Size = new System.Drawing.Size(80, 17);
			this.label23.TabIndex = 807;
			this.label23.Text = "Vehículo Marca";
			this.label23.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.label23.Visible = false;
			// 
			// chkHipoteca
			// 
			appearance42.ForeColorDisabled = System.Drawing.Color.Black;
			this.chkHipoteca.Appearance = appearance42;
			this.chkHipoteca.BackColor = System.Drawing.Color.Transparent;
			this.chkHipoteca.CausesValidation = false;
			this.chkHipoteca.Enabled = false;
			this.chkHipoteca.Location = new System.Drawing.Point(584, 112);
			this.chkHipoteca.Name = "chkHipoteca";
			this.chkHipoteca.Size = new System.Drawing.Size(144, 22);
			this.chkHipoteca.TabIndex = 806;
			this.chkHipoteca.Text = "Hipoteca Inmuebles";
			this.chkHipoteca.Visible = false;
			// 
			// chkSeparacionBienes
			// 
			appearance43.ForeColorDisabled = System.Drawing.Color.Black;
			this.chkSeparacionBienes.Appearance = appearance43;
			this.chkSeparacionBienes.BackColor = System.Drawing.Color.Transparent;
			this.chkSeparacionBienes.CausesValidation = false;
			this.chkSeparacionBienes.Enabled = false;
			this.chkSeparacionBienes.Location = new System.Drawing.Point(408, 112);
			this.chkSeparacionBienes.Name = "chkSeparacionBienes";
			this.chkSeparacionBienes.Size = new System.Drawing.Size(168, 22);
			this.chkSeparacionBienes.TabIndex = 805;
			this.chkSeparacionBienes.Text = "Separación De Bienes";
			this.chkSeparacionBienes.Visible = false;
			// 
			// txtValorVehiculo
			// 
			this.txtValorVehiculo.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtValorVehiculo.Enabled = false;
			this.txtValorVehiculo.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.txtValorVehiculo.Location = new System.Drawing.Point(304, 112);
			this.txtValorVehiculo.MaxValue = 1000000;
			this.txtValorVehiculo.MinValue = 0;
			this.txtValorVehiculo.Name = "txtValorVehiculo";
			this.txtValorVehiculo.PromptChar = ' ';
			this.txtValorVehiculo.Size = new System.Drawing.Size(96, 22);
			this.txtValorVehiculo.SpinButtonDisplayStyle = Infragistics.Win.ButtonDisplayStyle.Always;
			this.txtValorVehiculo.TabIndex = 804;
			this.txtValorVehiculo.Visible = false;
			this.txtValorVehiculo.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtValorVehiculo_KeyDown);
			this.txtValorVehiculo.Click += new System.EventHandler(this.txtValorVehiculo_Click);
			this.txtValorVehiculo.ValueChanged += new System.EventHandler(this.txtValorVehiculo_ValueChanged);
			// 
			// txtModelo
			// 
			appearance44.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtModelo.Appearance = appearance44;
			this.txtModelo.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtModelo.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtModelo.Enabled = false;
			this.txtModelo.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.txtModelo.Location = new System.Drawing.Point(168, 112);
			this.txtModelo.MaxLength = 10;
			this.txtModelo.Name = "txtModelo";
			this.txtModelo.Size = new System.Drawing.Size(128, 22);
			this.txtModelo.TabIndex = 803;
			this.txtModelo.Visible = false;
			this.txtModelo.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtModelo_KeyDown);
			this.txtModelo.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.txtModelo_KeyPress);
			// 
			// txtVehiculoMarca
			// 
			appearance45.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtVehiculoMarca.Appearance = appearance45;
			this.txtVehiculoMarca.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtVehiculoMarca.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtVehiculoMarca.Enabled = false;
			this.txtVehiculoMarca.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.txtVehiculoMarca.Location = new System.Drawing.Point(8, 112);
			this.txtVehiculoMarca.MaxLength = 10;
			this.txtVehiculoMarca.Name = "txtVehiculoMarca";
			this.txtVehiculoMarca.Size = new System.Drawing.Size(128, 22);
			this.txtVehiculoMarca.TabIndex = 802;
			this.txtVehiculoMarca.Visible = false;
			this.txtVehiculoMarca.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtVehiculoMarca_KeyDown);
			this.txtVehiculoMarca.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.txtVehiculoMarca_KeyPress);
			// 
			// label24
			// 
			this.label24.AutoSize = true;
			this.label24.BackColor = System.Drawing.Color.Transparent;
			this.label24.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label24.Location = new System.Drawing.Point(752, 56);
			this.label24.Name = "label24";
			this.label24.Size = new System.Drawing.Size(106, 17);
			this.label24.TabIndex = 801;
			this.label24.Text = "Actividad Económica";
			this.label24.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// label25
			// 
			this.label25.AutoSize = true;
			this.label25.BackColor = System.Drawing.Color.Transparent;
			this.label25.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label25.Location = new System.Drawing.Point(592, 56);
			this.label25.Name = "label25";
			this.label25.Size = new System.Drawing.Size(90, 17);
			this.label25.TabIndex = 800;
			this.label25.Text = "Situación Laboral";
			this.label25.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// label26
			// 
			this.label26.AutoSize = true;
			this.label26.BackColor = System.Drawing.Color.Transparent;
			this.label26.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label26.Location = new System.Drawing.Point(440, 56);
			this.label26.Name = "label26";
			this.label26.Size = new System.Drawing.Size(50, 17);
			this.label26.TabIndex = 799;
			this.label26.Text = "Profesión";
			this.label26.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// label27
			// 
			this.label27.AutoSize = true;
			this.label27.BackColor = System.Drawing.Color.Transparent;
			this.label27.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label27.Location = new System.Drawing.Point(272, 56);
			this.label27.Name = "label27";
			this.label27.Size = new System.Drawing.Size(82, 17);
			this.label27.TabIndex = 798;
			this.label27.Text = "Nivel Educación";
			this.label27.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// label28
			// 
			this.label28.AutoSize = true;
			this.label28.BackColor = System.Drawing.Color.Transparent;
			this.label28.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label28.Location = new System.Drawing.Point(168, 56);
			this.label28.Name = "label28";
			this.label28.Size = new System.Drawing.Size(72, 17);
			this.label28.TabIndex = 794;
			this.label28.Text = "Dependientes";
			this.label28.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// txtDependientes
			// 
			this.txtDependientes.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtDependientes.Enabled = false;
			this.txtDependientes.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.txtDependientes.Location = new System.Drawing.Point(168, 72);
			this.txtDependientes.MaxValue = 100;
			this.txtDependientes.MinValue = 0;
			this.txtDependientes.Name = "txtDependientes";
			this.txtDependientes.PromptChar = ' ';
			this.txtDependientes.Size = new System.Drawing.Size(90, 22);
			this.txtDependientes.SpinButtonDisplayStyle = Infragistics.Win.ButtonDisplayStyle.Always;
			this.txtDependientes.TabIndex = 792;
			this.txtDependientes.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtDependientes_KeyDown);
			// 
			// label29
			// 
			this.label29.AutoSize = true;
			this.label29.BackColor = System.Drawing.Color.Transparent;
			this.label29.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label29.Location = new System.Drawing.Point(8, 56);
			this.label29.Name = "label29";
			this.label29.Size = new System.Drawing.Size(61, 17);
			this.label29.TabIndex = 791;
			this.label29.Text = "Estado Civil";
			this.label29.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// cmbEstadoCivil
			// 
			this.cmbEstadoCivil.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbEstadoCivil.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbEstadoCivil.DataSource = this.ultraDataSource11;
			ultraGridBand7.AddButtonCaption = "RolSucursal";
			ultraGridColumn15.Header.VisiblePosition = 0;
			ultraGridColumn15.Hidden = true;
			ultraGridColumn16.Header.Caption = "Estado Civil";
			ultraGridColumn16.Header.VisiblePosition = 1;
			ultraGridColumn16.Width = 150;
			ultraGridBand7.Columns.AddRange(new object[] {
																										 ultraGridColumn15,
																										 ultraGridColumn16});
			this.cmbEstadoCivil.DisplayLayout.BandsSerializer.Add(ultraGridBand7);
			this.cmbEstadoCivil.DisplayMember = "Nombre";
			this.cmbEstadoCivil.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbEstadoCivil.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbEstadoCivil.Enabled = false;
			this.cmbEstadoCivil.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbEstadoCivil.Location = new System.Drawing.Point(8, 72);
			this.cmbEstadoCivil.Name = "cmbEstadoCivil";
			this.cmbEstadoCivil.Size = new System.Drawing.Size(150, 22);
			this.cmbEstadoCivil.TabIndex = 790;
			this.cmbEstadoCivil.ValueMember = "idEdoCivil";
			this.cmbEstadoCivil.KeyDown += new System.Windows.Forms.KeyEventHandler(this.cmbEstadoCivil_KeyDown);
			this.cmbEstadoCivil.ValueChanged += new System.EventHandler(this.cmbEstadoCivil_ValueChanged);
			// 
			// ultraDataSource11
			// 
			ultraDataColumn15.DataType = typeof(int);
			this.ultraDataSource11.Band.Columns.AddRange(new object[] {
																																	ultraDataColumn15,
																																	ultraDataColumn16});
			// 
			// label30
			// 
			this.label30.AutoSize = true;
			this.label30.BackColor = System.Drawing.Color.Transparent;
			this.label30.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label30.Location = new System.Drawing.Point(584, 16);
			this.label30.Name = "label30";
			this.label30.Size = new System.Drawing.Size(99, 17);
			this.label30.TabIndex = 788;
			this.label30.Text = "Cantón Nacimiento";
			this.label30.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// label31
			// 
			this.label31.AutoSize = true;
			this.label31.BackColor = System.Drawing.Color.Transparent;
			this.label31.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label31.Location = new System.Drawing.Point(432, 16);
			this.label31.Name = "label31";
			this.label31.Size = new System.Drawing.Size(108, 17);
			this.label31.TabIndex = 787;
			this.label31.Text = "Provincia Nacimiento";
			this.label31.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// label32
			// 
			this.label32.AutoSize = true;
			this.label32.BackColor = System.Drawing.Color.Transparent;
			this.label32.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label32.Location = new System.Drawing.Point(168, 16);
			this.label32.Name = "label32";
			this.label32.Size = new System.Drawing.Size(73, 17);
			this.label32.TabIndex = 786;
			this.label32.Text = "F. Nacimiento";
			this.label32.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// label33
			// 
			this.label33.AutoSize = true;
			this.label33.BackColor = System.Drawing.Color.Transparent;
			this.label33.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label33.Location = new System.Drawing.Point(272, 16);
			this.label33.Name = "label33";
			this.label33.Size = new System.Drawing.Size(28, 17);
			this.label33.TabIndex = 785;
			this.label33.Text = "Sexo";
			this.label33.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// cmbCantonNacimiento
			// 
			this.cmbCantonNacimiento.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbCantonNacimiento.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbCantonNacimiento.DataSource = this.ultraDataSource2;
			ultraGridColumn17.Header.VisiblePosition = 0;
			ultraGridColumn17.Hidden = true;
			ultraGridColumn18.Header.Caption = "Canton";
			ultraGridColumn18.Header.VisiblePosition = 1;
			ultraGridColumn18.Width = 150;
			ultraGridBand8.Columns.AddRange(new object[] {
																										 ultraGridColumn17,
																										 ultraGridColumn18});
			this.cmbCantonNacimiento.DisplayLayout.BandsSerializer.Add(ultraGridBand8);
			this.cmbCantonNacimiento.DisplayLayout.Override.RowSelectors = Infragistics.Win.DefaultableBoolean.True;
			this.cmbCantonNacimiento.DisplayMember = "Nombre";
			this.cmbCantonNacimiento.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbCantonNacimiento.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbCantonNacimiento.Enabled = false;
			this.cmbCantonNacimiento.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbCantonNacimiento.Location = new System.Drawing.Point(592, 32);
			this.cmbCantonNacimiento.Name = "cmbCantonNacimiento";
			this.cmbCantonNacimiento.Size = new System.Drawing.Size(150, 22);
			this.cmbCantonNacimiento.TabIndex = 784;
			this.cmbCantonNacimiento.ValueMember = "idCanton";
			this.cmbCantonNacimiento.KeyDown += new System.Windows.Forms.KeyEventHandler(this.cmbCantonNacimiento_KeyDown);
			// 
			// ultraDataSource2
			// 
			ultraDataColumn17.DataType = typeof(int);
			this.ultraDataSource2.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn17,
																																 ultraDataColumn18});
			// 
			// cmbProvinciaNAcimiento
			// 
			this.cmbProvinciaNAcimiento.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbProvinciaNAcimiento.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbProvinciaNAcimiento.DataSource = this.ultraDataSource1;
			ultraGridColumn19.Header.VisiblePosition = 0;
			ultraGridColumn19.Hidden = true;
			ultraGridColumn19.Width = 50;
			ultraGridColumn20.Header.Caption = "Provincia";
			ultraGridColumn20.Header.VisiblePosition = 1;
			ultraGridColumn20.Width = 150;
			ultraGridBand9.Columns.AddRange(new object[] {
																										 ultraGridColumn19,
																										 ultraGridColumn20});
			this.cmbProvinciaNAcimiento.DisplayLayout.BandsSerializer.Add(ultraGridBand9);
			this.cmbProvinciaNAcimiento.DisplayMember = "Nombre";
			this.cmbProvinciaNAcimiento.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbProvinciaNAcimiento.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbProvinciaNAcimiento.Enabled = false;
			this.cmbProvinciaNAcimiento.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbProvinciaNAcimiento.Location = new System.Drawing.Point(432, 32);
			this.cmbProvinciaNAcimiento.Name = "cmbProvinciaNAcimiento";
			this.cmbProvinciaNAcimiento.Size = new System.Drawing.Size(150, 22);
			this.cmbProvinciaNAcimiento.TabIndex = 782;
			this.cmbProvinciaNAcimiento.ValueMember = "idProvincia";
			this.cmbProvinciaNAcimiento.KeyDown += new System.Windows.Forms.KeyEventHandler(this.cmbProvinciaNAcimiento_KeyDown);
			this.cmbProvinciaNAcimiento.ValueChanged += new System.EventHandler(this.cmbProvinciaNAcimiento_ValueChanged);
			// 
			// ultraDataSource1
			// 
			ultraDataColumn19.DataType = typeof(int);
			this.ultraDataSource1.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn19,
																																 ultraDataColumn20});
			// 
			// dtFechaNacimiento
			// 
			this.dtFechaNacimiento.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			dateButton1.Caption = "Today";
			this.dtFechaNacimiento.DateButtons.Add(dateButton1);
			this.dtFechaNacimiento.Enabled = false;
			this.dtFechaNacimiento.Format = "dd/MM/yyyy";
			this.dtFechaNacimiento.Location = new System.Drawing.Point(168, 32);
			this.dtFechaNacimiento.Name = "dtFechaNacimiento";
			this.dtFechaNacimiento.NonAutoSizeHeight = 23;
			this.dtFechaNacimiento.Size = new System.Drawing.Size(90, 21);
			this.dtFechaNacimiento.TabIndex = 779;
			this.dtFechaNacimiento.Value = ((object)(resources.GetObject("dtFechaNacimiento.Value")));
			this.dtFechaNacimiento.KeyDown += new System.Windows.Forms.KeyEventHandler(this.dtFechaNacimiento_KeyDown);
			// 
			// cmbNacionalidad
			// 
			this.cmbNacionalidad.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbNacionalidad.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbNacionalidad.DataSource = this.ultraDataSource26;
			ultraGridBand10.AddButtonCaption = "RolSucursal";
			ultraGridColumn21.Header.VisiblePosition = 0;
			ultraGridColumn21.Hidden = true;
			ultraGridColumn21.Width = 150;
			ultraGridColumn22.Header.Caption = "Nacionalidad";
			ultraGridColumn22.Header.VisiblePosition = 1;
			ultraGridColumn22.Width = 150;
			ultraGridBand10.Columns.AddRange(new object[] {
																											ultraGridColumn21,
																											ultraGridColumn22});
			this.cmbNacionalidad.DisplayLayout.BandsSerializer.Add(ultraGridBand10);
			this.cmbNacionalidad.DisplayMember = "Nombre";
			this.cmbNacionalidad.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbNacionalidad.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbNacionalidad.Enabled = false;
			this.cmbNacionalidad.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbNacionalidad.Location = new System.Drawing.Point(8, 32);
			this.cmbNacionalidad.Name = "cmbNacionalidad";
			this.cmbNacionalidad.Size = new System.Drawing.Size(150, 22);
			this.cmbNacionalidad.TabIndex = 354;
			this.cmbNacionalidad.ValueMember = "idNacionalidad";
			this.cmbNacionalidad.KeyDown += new System.Windows.Forms.KeyEventHandler(this.cmbNacionalidad_KeyDown);
			// 
			// ultraDataSource26
			// 
			ultraDataColumn21.DataType = typeof(int);
			this.ultraDataSource26.Band.Columns.AddRange(new object[] {
																																	ultraDataColumn21,
																																	ultraDataColumn22});
			// 
			// label34
			// 
			this.label34.AutoSize = true;
			this.label34.BackColor = System.Drawing.Color.Transparent;
			this.label34.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label34.Location = new System.Drawing.Point(8, 16);
			this.label34.Name = "label34";
			this.label34.Size = new System.Drawing.Size(68, 17);
			this.label34.TabIndex = 353;
			this.label34.Text = "Nacionalidad";
			this.label34.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// groupBox1
			// 
			this.groupBox1.BackColor = System.Drawing.SystemColors.Control;
			this.groupBox1.Controls.Add(this.btnEquifax);
			this.groupBox1.Controls.Add(this.button1);
			this.groupBox1.Controls.Add(this.btnUbicacion);
			this.groupBox1.Controls.Add(this.label146);
			this.groupBox1.Controls.Add(this.txtCelArrendador);
			this.groupBox1.Controls.Add(this.cmbTiempoVivienda);
			this.groupBox1.Controls.Add(this.txtDDValorInmueble);
			this.groupBox1.Controls.Add(this.cmbDDTipoVivienda);
			this.groupBox1.Controls.Add(this.dtTiempoVivienda);
			this.groupBox1.Controls.Add(this.cmbDDCiudadInmueble);
			this.groupBox1.Controls.Add(this.cmbDDInmueble);
			this.groupBox1.Controls.Add(this.label40);
			this.groupBox1.Controls.Add(this.label39);
			this.groupBox1.Controls.Add(this.label38);
			this.groupBox1.Controls.Add(this.label37);
			this.groupBox1.Controls.Add(this.label36);
			this.groupBox1.Controls.Add(this.label35);
			this.groupBox1.Controls.Add(this.label20);
			this.groupBox1.Controls.Add(this.txtTelfArrendador);
			this.groupBox1.Controls.Add(this.txtNombreArrendador);
			this.groupBox1.Controls.Add(this.label19);
			this.groupBox1.Controls.Add(this.txtDDReferenciasU);
			this.groupBox1.Controls.Add(this.btnCelularcliente);
			this.groupBox1.Controls.Add(this.btnClienteTelefonoCasa2);
			this.groupBox1.Controls.Add(this.btnClienteTelefonoCasa1);
			this.groupBox1.Controls.Add(this.label18);
			this.groupBox1.Controls.Add(this.label17);
			this.groupBox1.Controls.Add(this.label15);
			this.groupBox1.Controls.Add(this.label14);
			this.groupBox1.Controls.Add(this.label13);
			this.groupBox1.Controls.Add(this.label12);
			this.groupBox1.Controls.Add(this.txtDDEmail);
			this.groupBox1.Controls.Add(this.txtDDCelular);
			this.groupBox1.Controls.Add(this.txtDDTelefono2);
			this.groupBox1.Controls.Add(this.txtDDTelefono1);
			this.groupBox1.Controls.Add(this.txtDDCalleSecundaria);
			this.groupBox1.Controls.Add(this.txtDDNumeroCasa);
			this.groupBox1.Controls.Add(this.txtDDCallaPrincipal);
			this.groupBox1.Controls.Add(this.label10);
			this.groupBox1.Controls.Add(this.label9);
			this.groupBox1.Controls.Add(this.label8);
			this.groupBox1.Controls.Add(this.label7);
			this.groupBox1.Controls.Add(this.cmbDDBarrio);
			this.groupBox1.Controls.Add(this.cmbDDParroquia);
			this.groupBox1.Controls.Add(this.cmbDDCanton);
			this.groupBox1.Controls.Add(this.cmbDDProvincia);
			this.groupBox1.Controls.Add(this.label6);
			this.groupBox1.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.groupBox1.Location = new System.Drawing.Point(0, 144);
			this.groupBox1.Name = "groupBox1";
			this.groupBox1.Size = new System.Drawing.Size(1112, 232);
			this.groupBox1.TabIndex = 198;
			this.groupBox1.TabStop = false;
			this.groupBox1.Text = "Datos de Domicilio";
			this.groupBox1.Enter += new System.EventHandler(this.groupBox1_Enter);
			// 
			// btnEquifax
			// 
			this.btnEquifax.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			this.btnEquifax.Image = ((System.Drawing.Image)(resources.GetObject("btnEquifax.Image")));
			this.btnEquifax.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnEquifax.Location = new System.Drawing.Point(840, 64);
			this.btnEquifax.Name = "btnEquifax";
			this.btnEquifax.Size = new System.Drawing.Size(104, 23);
			this.btnEquifax.TabIndex = 1053;
			this.btnEquifax.Text = "EquiFax GPS";
			this.btnEquifax.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnEquifax.Visible = false;
			this.btnEquifax.Click += new System.EventHandler(this.btnEquifax_Click);
			// 
			// button1
			// 
			this.button1.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			this.button1.Image = ((System.Drawing.Image)(resources.GetObject("button1.Image")));
			this.button1.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.button1.Location = new System.Drawing.Point(952, 32);
			this.button1.Name = "button1";
			this.button1.Size = new System.Drawing.Size(128, 23);
			this.button1.TabIndex = 1052;
			this.button1.Text = "Ubicación-Trabajo";
			this.button1.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.button1.Click += new System.EventHandler(this.button1_Click);
			// 
			// btnUbicacion
			// 
			this.btnUbicacion.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			this.btnUbicacion.Image = ((System.Drawing.Image)(resources.GetObject("btnUbicacion.Image")));
			this.btnUbicacion.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnUbicacion.Location = new System.Drawing.Point(952, 64);
			this.btnUbicacion.Name = "btnUbicacion";
			this.btnUbicacion.Size = new System.Drawing.Size(128, 23);
			this.btnUbicacion.TabIndex = 1051;
			this.btnUbicacion.Text = "Ubicación- Domicilio";
			this.btnUbicacion.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnUbicacion.Click += new System.EventHandler(this.btnUbicacion_Click);
			// 
			// label146
			// 
			this.label146.AutoSize = true;
			this.label146.BackColor = System.Drawing.Color.Transparent;
			this.label146.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label146.Location = new System.Drawing.Point(640, 136);
			this.label146.Name = "label146";
			this.label146.Size = new System.Drawing.Size(98, 17);
			this.label146.TabIndex = 1049;
			this.label146.Text = "Celular Arrendador";
			this.label146.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// txtCelArrendador
			// 
			this.txtCelArrendador.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtCelArrendador.Enabled = false;
			this.txtCelArrendador.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.txtCelArrendador.Location = new System.Drawing.Point(640, 152);
			this.txtCelArrendador.MaxLength = 10;
			this.txtCelArrendador.Name = "txtCelArrendador";
			this.txtCelArrendador.Size = new System.Drawing.Size(100, 22);
			this.txtCelArrendador.TabIndex = 1048;
			this.txtCelArrendador.ValueChanged += new System.EventHandler(this.txtCelArrendador_ValueChanged);
			// 
			// cmbTiempoVivienda
			// 
			this.cmbTiempoVivienda.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbTiempoVivienda.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbTiempoVivienda.DataSource = this.ultraDataSource49;
			ultraGridColumn23.Header.VisiblePosition = 0;
			ultraGridColumn23.Hidden = true;
			ultraGridColumn23.Width = 8;
			ultraGridColumn24.Header.VisiblePosition = 1;
			ultraGridColumn24.Width = 150;
			ultraGridBand11.Columns.AddRange(new object[] {
																											ultraGridColumn23,
																											ultraGridColumn24});
			this.cmbTiempoVivienda.DisplayLayout.BandsSerializer.Add(ultraGridBand11);
			this.cmbTiempoVivienda.DisplayLayout.Override.RowSelectors = Infragistics.Win.DefaultableBoolean.True;
			this.cmbTiempoVivienda.DisplayMember = "Descripcion";
			this.cmbTiempoVivienda.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbTiempoVivienda.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbTiempoVivienda.Enabled = false;
			this.cmbTiempoVivienda.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbTiempoVivienda.Location = new System.Drawing.Point(888, 112);
			this.cmbTiempoVivienda.Name = "cmbTiempoVivienda";
			this.cmbTiempoVivienda.Size = new System.Drawing.Size(120, 22);
			this.cmbTiempoVivienda.TabIndex = 1047;
			this.cmbTiempoVivienda.ValueMember = "idCre_Tiempo";
			// 
			// ultraDataSource49
			// 
			ultraDataColumn23.DataType = typeof(int);
			this.ultraDataSource49.Band.Columns.AddRange(new object[] {
																																	ultraDataColumn23,
																																	ultraDataColumn24});
			// 
			// txtDDValorInmueble
			// 
			this.txtDDValorInmueble.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtDDValorInmueble.Enabled = false;
			this.txtDDValorInmueble.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.txtDDValorInmueble.Location = new System.Drawing.Point(640, 192);
			this.txtDDValorInmueble.MaxValue = 1000000;
			this.txtDDValorInmueble.MinValue = 0;
			this.txtDDValorInmueble.Name = "txtDDValorInmueble";
			this.txtDDValorInmueble.PromptChar = ' ';
			this.txtDDValorInmueble.Size = new System.Drawing.Size(96, 22);
			this.txtDDValorInmueble.SpinButtonDisplayStyle = Infragistics.Win.ButtonDisplayStyle.Always;
			this.txtDDValorInmueble.TabIndex = 852;
			// 
			// cmbDDTipoVivienda
			// 
			this.cmbDDTipoVivienda.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbDDTipoVivienda.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbDDTipoVivienda.DataSource = this.ultraDataSource31;
			ultraGridColumn25.Header.VisiblePosition = 0;
			ultraGridColumn25.Hidden = true;
			ultraGridColumn26.Header.VisiblePosition = 1;
			ultraGridColumn26.Width = 150;
			ultraGridBand12.Columns.AddRange(new object[] {
																											ultraGridColumn25,
																											ultraGridColumn26});
			this.cmbDDTipoVivienda.DisplayLayout.BandsSerializer.Add(ultraGridBand12);
			this.cmbDDTipoVivienda.DisplayLayout.Override.RowSelectors = Infragistics.Win.DefaultableBoolean.True;
			this.cmbDDTipoVivienda.DisplayMember = "Nombre";
			this.cmbDDTipoVivienda.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbDDTipoVivienda.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbDDTipoVivienda.Enabled = false;
			this.cmbDDTipoVivienda.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbDDTipoVivienda.Location = new System.Drawing.Point(728, 112);
			this.cmbDDTipoVivienda.Name = "cmbDDTipoVivienda";
			this.cmbDDTipoVivienda.Size = new System.Drawing.Size(150, 22);
			this.cmbDDTipoVivienda.TabIndex = 851;
			this.cmbDDTipoVivienda.ValueMember = "idTipoVivienda";
			this.cmbDDTipoVivienda.KeyDown += new System.Windows.Forms.KeyEventHandler(this.cmbDDTipoVivienda_KeyDown_1);
			this.cmbDDTipoVivienda.ValueChanged += new System.EventHandler(this.cmbDDTipoVivienda_ValueChanged_1);
			// 
			// ultraDataSource31
			// 
			ultraDataColumn25.DataType = typeof(System.UInt32);
			this.ultraDataSource31.Band.Columns.AddRange(new object[] {
																																	ultraDataColumn25,
																																	ultraDataColumn26});
			// 
			// dtTiempoVivienda
			// 
			this.dtTiempoVivienda.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			dateButton2.Caption = "Today";
			this.dtTiempoVivienda.DateButtons.Add(dateButton2);
			this.dtTiempoVivienda.Enabled = false;
			this.dtTiempoVivienda.Format = "dd/MM/yyyy";
			this.dtTiempoVivienda.Location = new System.Drawing.Point(944, 168);
			this.dtTiempoVivienda.Name = "dtTiempoVivienda";
			this.dtTiempoVivienda.NonAutoSizeHeight = 23;
			this.dtTiempoVivienda.Size = new System.Drawing.Size(104, 21);
			this.dtTiempoVivienda.TabIndex = 849;
			this.dtTiempoVivienda.Value = ((object)(resources.GetObject("dtTiempoVivienda.Value")));
			this.dtTiempoVivienda.Visible = false;
			this.dtTiempoVivienda.KeyDown += new System.Windows.Forms.KeyEventHandler(this.dtTiempoVivienda_KeyDown_1);
			// 
			// cmbDDCiudadInmueble
			// 
			this.cmbDDCiudadInmueble.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbDDCiudadInmueble.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbDDCiudadInmueble.DataSource = this.ultraDataSource38;
			ultraGridColumn27.Header.VisiblePosition = 0;
			ultraGridColumn27.Hidden = true;
			ultraGridColumn28.Header.Caption = "Ciudad Inmueble";
			ultraGridColumn28.Header.VisiblePosition = 1;
			ultraGridColumn28.Width = 150;
			ultraGridBand13.Columns.AddRange(new object[] {
																											ultraGridColumn27,
																											ultraGridColumn28});
			this.cmbDDCiudadInmueble.DisplayLayout.BandsSerializer.Add(ultraGridBand13);
			this.cmbDDCiudadInmueble.DisplayLayout.Override.RowSelectors = Infragistics.Win.DefaultableBoolean.True;
			this.cmbDDCiudadInmueble.DisplayMember = "Nombre";
			this.cmbDDCiudadInmueble.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbDDCiudadInmueble.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbDDCiudadInmueble.Enabled = false;
			this.cmbDDCiudadInmueble.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbDDCiudadInmueble.Location = new System.Drawing.Point(480, 192);
			this.cmbDDCiudadInmueble.Name = "cmbDDCiudadInmueble";
			this.cmbDDCiudadInmueble.Size = new System.Drawing.Size(150, 22);
			this.cmbDDCiudadInmueble.TabIndex = 848;
			this.cmbDDCiudadInmueble.ValueMember = "idCanton";
			this.cmbDDCiudadInmueble.KeyDown += new System.Windows.Forms.KeyEventHandler(this.cmbDDCiudadInmueble_KeyDown);
			// 
			// ultraDataSource38
			// 
			ultraDataColumn27.DataType = typeof(int);
			this.ultraDataSource38.Band.Columns.AddRange(new object[] {
																																	ultraDataColumn27,
																																	ultraDataColumn28});
			// 
			// cmbDDInmueble
			// 
			this.cmbDDInmueble.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbDDInmueble.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbDDInmueble.DataSource = this.ultraDataSource27;
			ultraGridColumn29.Header.VisiblePosition = 0;
			ultraGridColumn29.Hidden = true;
			ultraGridColumn30.Header.Caption = "Inmueble";
			ultraGridColumn30.Header.VisiblePosition = 1;
			ultraGridColumn30.Width = 150;
			ultraGridBand14.Columns.AddRange(new object[] {
																											ultraGridColumn29,
																											ultraGridColumn30});
			this.cmbDDInmueble.DisplayLayout.BandsSerializer.Add(ultraGridBand14);
			this.cmbDDInmueble.DisplayLayout.Override.RowSelectors = Infragistics.Win.DefaultableBoolean.True;
			this.cmbDDInmueble.DisplayMember = "Nombre";
			this.cmbDDInmueble.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbDDInmueble.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbDDInmueble.Enabled = false;
			this.cmbDDInmueble.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbDDInmueble.Location = new System.Drawing.Point(320, 192);
			this.cmbDDInmueble.Name = "cmbDDInmueble";
			this.cmbDDInmueble.Size = new System.Drawing.Size(150, 22);
			this.cmbDDInmueble.TabIndex = 847;
			this.cmbDDInmueble.ValueMember = "idInmueble";
			this.cmbDDInmueble.KeyDown += new System.Windows.Forms.KeyEventHandler(this.cmbDDInmueble_KeyDown);
			// 
			// ultraDataSource27
			// 
			ultraDataColumn29.DataType = typeof(int);
			this.ultraDataSource27.Band.Columns.AddRange(new object[] {
																																	ultraDataColumn29,
																																	ultraDataColumn30});
			// 
			// label40
			// 
			this.label40.AutoSize = true;
			this.label40.BackColor = System.Drawing.Color.Transparent;
			this.label40.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label40.Location = new System.Drawing.Point(640, 176);
			this.label40.Name = "label40";
			this.label40.Size = new System.Drawing.Size(80, 17);
			this.label40.TabIndex = 845;
			this.label40.Text = "Valor Inmueble";
			this.label40.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// label39
			// 
			this.label39.AutoSize = true;
			this.label39.BackColor = System.Drawing.Color.Transparent;
			this.label39.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label39.Location = new System.Drawing.Point(480, 176);
			this.label39.Name = "label39";
			this.label39.Size = new System.Drawing.Size(89, 17);
			this.label39.TabIndex = 844;
			this.label39.Text = "Ciudad Inmueble";
			this.label39.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// label38
			// 
			this.label38.AutoSize = true;
			this.label38.BackColor = System.Drawing.Color.Transparent;
			this.label38.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label38.Location = new System.Drawing.Point(320, 176);
			this.label38.Name = "label38";
			this.label38.Size = new System.Drawing.Size(51, 17);
			this.label38.TabIndex = 843;
			this.label38.Text = "Inmueble";
			this.label38.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// label37
			// 
			this.label37.AutoSize = true;
			this.label37.BackColor = System.Drawing.Color.Transparent;
			this.label37.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label37.Location = new System.Drawing.Point(504, 136);
			this.label37.Name = "label37";
			this.label37.Size = new System.Drawing.Size(82, 17);
			this.label37.TabIndex = 842;
			this.label37.Text = "Telf Arrendador";
			this.label37.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// label36
			// 
			this.label36.AutoSize = true;
			this.label36.BackColor = System.Drawing.Color.Transparent;
			this.label36.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label36.Location = new System.Drawing.Point(320, 136);
			this.label36.Name = "label36";
			this.label36.Size = new System.Drawing.Size(103, 17);
			this.label36.TabIndex = 841;
			this.label36.Text = "Nombre Arrendador";
			this.label36.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// label35
			// 
			this.label35.AutoSize = true;
			this.label35.BackColor = System.Drawing.Color.Transparent;
			this.label35.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label35.Location = new System.Drawing.Point(888, 96);
			this.label35.Name = "label35";
			this.label35.Size = new System.Drawing.Size(87, 17);
			this.label35.TabIndex = 840;
			this.label35.Text = "Tiempo Vivienda";
			this.label35.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// label20
			// 
			this.label20.AutoSize = true;
			this.label20.BackColor = System.Drawing.Color.Transparent;
			this.label20.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label20.Location = new System.Drawing.Point(728, 96);
			this.label20.Name = "label20";
			this.label20.Size = new System.Drawing.Size(71, 17);
			this.label20.TabIndex = 839;
			this.label20.Text = "Tipo Vivienda";
			this.label20.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// txtTelfArrendador
			// 
			this.txtTelfArrendador.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtTelfArrendador.Enabled = false;
			this.txtTelfArrendador.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.txtTelfArrendador.Location = new System.Drawing.Point(504, 152);
			this.txtTelfArrendador.MaxLength = 9;
			this.txtTelfArrendador.Name = "txtTelfArrendador";
			this.txtTelfArrendador.Size = new System.Drawing.Size(128, 22);
			this.txtTelfArrendador.TabIndex = 835;
			this.txtTelfArrendador.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtTelfArrendador_KeyDown);
			this.txtTelfArrendador.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.txtTelfArrendador_KeyPress);
			// 
			// txtNombreArrendador
			// 
			appearance46.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtNombreArrendador.Appearance = appearance46;
			this.txtNombreArrendador.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtNombreArrendador.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtNombreArrendador.Enabled = false;
			this.txtNombreArrendador.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.txtNombreArrendador.Location = new System.Drawing.Point(320, 152);
			this.txtNombreArrendador.Name = "txtNombreArrendador";
			this.txtNombreArrendador.Size = new System.Drawing.Size(180, 22);
			this.txtNombreArrendador.TabIndex = 834;
			this.txtNombreArrendador.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtNombreArrendador_KeyDown);
			this.txtNombreArrendador.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.txtNombreArrendador_KeyPress);
			// 
			// label19
			// 
			this.label19.AutoSize = true;
			this.label19.BackColor = System.Drawing.Color.Transparent;
			this.label19.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label19.Location = new System.Drawing.Point(16, 96);
			this.label19.Name = "label19";
			this.label19.Size = new System.Drawing.Size(108, 17);
			this.label19.TabIndex = 831;
			this.label19.Text = "Referencia Ubicacion";
			this.label19.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// txtDDReferenciasU
			// 
			this.txtDDReferenciasU.AcceptsReturn = true;
			appearance47.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtDDReferenciasU.Appearance = appearance47;
			this.txtDDReferenciasU.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtDDReferenciasU.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtDDReferenciasU.Enabled = false;
			this.txtDDReferenciasU.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.txtDDReferenciasU.Location = new System.Drawing.Point(16, 112);
			this.txtDDReferenciasU.Multiline = true;
			this.txtDDReferenciasU.Name = "txtDDReferenciasU";
			this.txtDDReferenciasU.Size = new System.Drawing.Size(300, 104);
			this.txtDDReferenciasU.TabIndex = 830;
			this.txtDDReferenciasU.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtDDReferenciasU_KeyDown);
			this.txtDDReferenciasU.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.txtDDReferenciasU_KeyPress);
			// 
			// btnCelularcliente
			// 
			this.btnCelularcliente.Image = ((System.Drawing.Image)(resources.GetObject("btnCelularcliente.Image")));
			this.btnCelularcliente.Location = new System.Drawing.Point(696, 112);
			this.btnCelularcliente.Name = "btnCelularcliente";
			this.btnCelularcliente.Size = new System.Drawing.Size(24, 23);
			this.btnCelularcliente.TabIndex = 829;
			this.btnCelularcliente.Click += new System.EventHandler(this.btnCelularcliente_Click);
			// 
			// btnClienteTelefonoCasa2
			// 
			this.btnClienteTelefonoCasa2.Image = ((System.Drawing.Image)(resources.GetObject("btnClienteTelefonoCasa2.Image")));
			this.btnClienteTelefonoCasa2.Location = new System.Drawing.Point(560, 112);
			this.btnClienteTelefonoCasa2.Name = "btnClienteTelefonoCasa2";
			this.btnClienteTelefonoCasa2.Size = new System.Drawing.Size(24, 23);
			this.btnClienteTelefonoCasa2.TabIndex = 828;
			this.btnClienteTelefonoCasa2.Click += new System.EventHandler(this.btnClienteTelefonoCasa2_Click);
			// 
			// btnClienteTelefonoCasa1
			// 
			this.btnClienteTelefonoCasa1.Image = ((System.Drawing.Image)(resources.GetObject("btnClienteTelefonoCasa1.Image")));
			this.btnClienteTelefonoCasa1.Location = new System.Drawing.Point(424, 112);
			this.btnClienteTelefonoCasa1.Name = "btnClienteTelefonoCasa1";
			this.btnClienteTelefonoCasa1.Size = new System.Drawing.Size(24, 23);
			this.btnClienteTelefonoCasa1.TabIndex = 827;
			this.btnClienteTelefonoCasa1.Click += new System.EventHandler(this.btnClienteTelefonoCasa1_Click);
			// 
			// label18
			// 
			this.label18.AutoSize = true;
			this.label18.BackColor = System.Drawing.Color.Transparent;
			this.label18.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label18.Location = new System.Drawing.Point(656, 16);
			this.label18.Name = "label18";
			this.label18.Size = new System.Drawing.Size(31, 17);
			this.label18.TabIndex = 826;
			this.label18.Text = "Email";
			this.label18.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// label17
			// 
			this.label17.AutoSize = true;
			this.label17.BackColor = System.Drawing.Color.Transparent;
			this.label17.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label17.Location = new System.Drawing.Point(592, 96);
			this.label17.Name = "label17";
			this.label17.Size = new System.Drawing.Size(76, 17);
			this.label17.TabIndex = 825;
			this.label17.Text = "Celular Cliente";
			this.label17.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// label15
			// 
			this.label15.AutoSize = true;
			this.label15.BackColor = System.Drawing.Color.Transparent;
			this.label15.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label15.Location = new System.Drawing.Point(456, 96);
			this.label15.Name = "label15";
			this.label15.Size = new System.Drawing.Size(85, 17);
			this.label15.TabIndex = 824;
			this.label15.Text = "Teléfono Casa 2";
			this.label15.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// label14
			// 
			this.label14.AutoSize = true;
			this.label14.BackColor = System.Drawing.Color.Transparent;
			this.label14.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label14.Location = new System.Drawing.Point(320, 96);
			this.label14.Name = "label14";
			this.label14.Size = new System.Drawing.Size(85, 17);
			this.label14.TabIndex = 823;
			this.label14.Text = "Teléfono Casa 1";
			this.label14.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// label13
			// 
			this.label13.AutoSize = true;
			this.label13.BackColor = System.Drawing.Color.Transparent;
			this.label13.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label13.Location = new System.Drawing.Point(456, 56);
			this.label13.Name = "label13";
			this.label13.Size = new System.Drawing.Size(87, 17);
			this.label13.TabIndex = 822;
			this.label13.Text = "Calle Secundaria";
			this.label13.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// label12
			// 
			this.label12.AutoSize = true;
			this.label12.BackColor = System.Drawing.Color.Transparent;
			this.label12.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label12.Location = new System.Drawing.Point(320, 56);
			this.label12.Name = "label12";
			this.label12.Size = new System.Drawing.Size(71, 17);
			this.label12.TabIndex = 821;
			this.label12.Text = "Número Casa";
			this.label12.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// txtDDEmail
			// 
			appearance48.BackColorDisabled = System.Drawing.Color.Silver;
			appearance48.BackColorDisabled2 = System.Drawing.Color.Silver;
			appearance48.BorderColor = System.Drawing.Color.Black;
			appearance48.BorderColor3DBase = System.Drawing.Color.Black;
			appearance48.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtDDEmail.Appearance = appearance48;
			this.txtDDEmail.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtDDEmail.Enabled = false;
			this.txtDDEmail.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.txtDDEmail.Location = new System.Drawing.Point(656, 32);
			this.txtDDEmail.Name = "txtDDEmail";
			this.txtDDEmail.Size = new System.Drawing.Size(168, 22);
			this.txtDDEmail.TabIndex = 819;
			this.txtDDEmail.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtDDEmail_KeyDown);
			// 
			// txtDDCelular
			// 
			this.txtDDCelular.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtDDCelular.Enabled = false;
			this.txtDDCelular.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.txtDDCelular.Location = new System.Drawing.Point(592, 112);
			this.txtDDCelular.MaxLength = 10;
			this.txtDDCelular.Name = "txtDDCelular";
			this.txtDDCelular.Size = new System.Drawing.Size(100, 22);
			this.txtDDCelular.TabIndex = 818;
			this.txtDDCelular.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtDDCelular_KeyDown);
			this.txtDDCelular.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.txtDDCelular_KeyPress);
			// 
			// txtDDTelefono2
			// 
			this.txtDDTelefono2.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtDDTelefono2.Enabled = false;
			this.txtDDTelefono2.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.txtDDTelefono2.Location = new System.Drawing.Point(456, 112);
			this.txtDDTelefono2.MaxLength = 10;
			this.txtDDTelefono2.Name = "txtDDTelefono2";
			this.txtDDTelefono2.Size = new System.Drawing.Size(100, 22);
			this.txtDDTelefono2.TabIndex = 817;
			this.txtDDTelefono2.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtDDTelefono2_KeyDown);
			this.txtDDTelefono2.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.txtDDTelefono2_KeyPress);
			// 
			// txtDDTelefono1
			// 
			this.txtDDTelefono1.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtDDTelefono1.Enabled = false;
			this.txtDDTelefono1.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.txtDDTelefono1.Location = new System.Drawing.Point(320, 112);
			this.txtDDTelefono1.MaxLength = 10;
			this.txtDDTelefono1.Name = "txtDDTelefono1";
			this.txtDDTelefono1.Size = new System.Drawing.Size(100, 22);
			this.txtDDTelefono1.TabIndex = 816;
			this.txtDDTelefono1.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtDDTelefono1_KeyDown);
			this.txtDDTelefono1.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.txtDDTelefono1_KeyPress);
			// 
			// txtDDCalleSecundaria
			// 
			appearance49.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtDDCalleSecundaria.Appearance = appearance49;
			this.txtDDCalleSecundaria.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtDDCalleSecundaria.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtDDCalleSecundaria.Enabled = false;
			this.txtDDCalleSecundaria.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.txtDDCalleSecundaria.Location = new System.Drawing.Point(456, 72);
			this.txtDDCalleSecundaria.Name = "txtDDCalleSecundaria";
			this.txtDDCalleSecundaria.Size = new System.Drawing.Size(300, 22);
			this.txtDDCalleSecundaria.TabIndex = 815;
			this.txtDDCalleSecundaria.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtDDCalleSecundaria_KeyDown);
			this.txtDDCalleSecundaria.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.txtDDCalleSecundaria_KeyPress);
			// 
			// txtDDNumeroCasa
			// 
			appearance50.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtDDNumeroCasa.Appearance = appearance50;
			this.txtDDNumeroCasa.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtDDNumeroCasa.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtDDNumeroCasa.Enabled = false;
			this.txtDDNumeroCasa.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.txtDDNumeroCasa.Location = new System.Drawing.Point(320, 72);
			this.txtDDNumeroCasa.Name = "txtDDNumeroCasa";
			this.txtDDNumeroCasa.Size = new System.Drawing.Size(128, 22);
			this.txtDDNumeroCasa.TabIndex = 814;
			this.txtDDNumeroCasa.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtDDNumeroCasa_KeyDown);
			this.txtDDNumeroCasa.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.txtDDNumeroCasa_KeyPress);
			// 
			// txtDDCallaPrincipal
			// 
			appearance51.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtDDCallaPrincipal.Appearance = appearance51;
			this.txtDDCallaPrincipal.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtDDCallaPrincipal.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtDDCallaPrincipal.Enabled = false;
			this.txtDDCallaPrincipal.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.txtDDCallaPrincipal.Location = new System.Drawing.Point(16, 72);
			this.txtDDCallaPrincipal.Name = "txtDDCallaPrincipal";
			this.txtDDCallaPrincipal.Size = new System.Drawing.Size(300, 22);
			this.txtDDCallaPrincipal.TabIndex = 813;
			this.txtDDCallaPrincipal.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtDDCallaPrincipal_KeyDown);
			this.txtDDCallaPrincipal.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.txtDDCallaPrincipal_KeyPress);
			// 
			// label10
			// 
			this.label10.AutoSize = true;
			this.label10.BackColor = System.Drawing.Color.Transparent;
			this.label10.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label10.Location = new System.Drawing.Point(496, 16);
			this.label10.Name = "label10";
			this.label10.Size = new System.Drawing.Size(34, 17);
			this.label10.TabIndex = 812;
			this.label10.Text = "Barrio";
			this.label10.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// label9
			// 
			this.label9.AutoSize = true;
			this.label9.BackColor = System.Drawing.Color.Transparent;
			this.label9.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label9.Location = new System.Drawing.Point(336, 16);
			this.label9.Name = "label9";
			this.label9.Size = new System.Drawing.Size(52, 17);
			this.label9.TabIndex = 811;
			this.label9.Text = "Parroquia";
			this.label9.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// label8
			// 
			this.label8.AutoSize = true;
			this.label8.BackColor = System.Drawing.Color.Transparent;
			this.label8.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label8.Location = new System.Drawing.Point(176, 16);
			this.label8.Name = "label8";
			this.label8.Size = new System.Drawing.Size(39, 17);
			this.label8.TabIndex = 810;
			this.label8.Text = "Cantón";
			this.label8.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// label7
			// 
			this.label7.AutoSize = true;
			this.label7.BackColor = System.Drawing.Color.Transparent;
			this.label7.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label7.Location = new System.Drawing.Point(16, 16);
			this.label7.Name = "label7";
			this.label7.Size = new System.Drawing.Size(49, 17);
			this.label7.TabIndex = 809;
			this.label7.Text = "Provincia";
			this.label7.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// cmbDDBarrio
			// 
			this.cmbDDBarrio.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbDDBarrio.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbDDBarrio.DataSource = this.ultraDataSource4;
			ultraGridColumn31.Header.VisiblePosition = 0;
			ultraGridColumn31.Hidden = true;
			ultraGridColumn32.Header.Caption = "Barrio";
			ultraGridColumn32.Header.VisiblePosition = 1;
			ultraGridColumn32.Width = 150;
			ultraGridBand15.Columns.AddRange(new object[] {
																											ultraGridColumn31,
																											ultraGridColumn32});
			this.cmbDDBarrio.DisplayLayout.BandsSerializer.Add(ultraGridBand15);
			this.cmbDDBarrio.DisplayLayout.Override.RowSelectors = Infragistics.Win.DefaultableBoolean.True;
			this.cmbDDBarrio.DisplayMember = "Nombre";
			this.cmbDDBarrio.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbDDBarrio.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbDDBarrio.Enabled = false;
			this.cmbDDBarrio.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbDDBarrio.Location = new System.Drawing.Point(496, 32);
			this.cmbDDBarrio.Name = "cmbDDBarrio";
			this.cmbDDBarrio.Size = new System.Drawing.Size(150, 22);
			this.cmbDDBarrio.TabIndex = 780;
			this.cmbDDBarrio.ValueMember = "idBarrio";
			this.cmbDDBarrio.KeyDown += new System.Windows.Forms.KeyEventHandler(this.cmbDDBarrio_KeyDown);
			// 
			// ultraDataSource4
			// 
			ultraDataColumn31.DataType = typeof(int);
			this.ultraDataSource4.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn31,
																																 ultraDataColumn32});
			// 
			// cmbDDParroquia
			// 
			this.cmbDDParroquia.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbDDParroquia.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbDDParroquia.DataSource = this.ultraDataSource3;
			ultraGridColumn33.Header.VisiblePosition = 0;
			ultraGridColumn33.Hidden = true;
			ultraGridColumn34.Header.Caption = "Provincia";
			ultraGridColumn34.Header.VisiblePosition = 1;
			ultraGridColumn34.Width = 150;
			ultraGridBand16.Columns.AddRange(new object[] {
																											ultraGridColumn33,
																											ultraGridColumn34});
			this.cmbDDParroquia.DisplayLayout.BandsSerializer.Add(ultraGridBand16);
			this.cmbDDParroquia.DisplayMember = "Nombre";
			this.cmbDDParroquia.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbDDParroquia.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbDDParroquia.Enabled = false;
			this.cmbDDParroquia.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbDDParroquia.Location = new System.Drawing.Point(336, 32);
			this.cmbDDParroquia.Name = "cmbDDParroquia";
			this.cmbDDParroquia.Size = new System.Drawing.Size(150, 22);
			this.cmbDDParroquia.TabIndex = 778;
			this.cmbDDParroquia.ValueMember = "idParroquia";
			this.cmbDDParroquia.KeyDown += new System.Windows.Forms.KeyEventHandler(this.cmbDDParroquia_KeyDown);
			this.cmbDDParroquia.ValueChanged += new System.EventHandler(this.cmbDDParroquia_ValueChanged);
			// 
			// ultraDataSource3
			// 
			ultraDataColumn33.DataType = typeof(int);
			this.ultraDataSource3.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn33,
																																 ultraDataColumn34});
			// 
			// cmbDDCanton
			// 
			this.cmbDDCanton.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbDDCanton.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbDDCanton.DataSource = this.ultraDataSource2;
			ultraGridColumn35.Header.VisiblePosition = 0;
			ultraGridColumn35.Hidden = true;
			ultraGridColumn36.Header.Caption = "Canton";
			ultraGridColumn36.Header.VisiblePosition = 1;
			ultraGridColumn36.Width = 150;
			ultraGridBand17.Columns.AddRange(new object[] {
																											ultraGridColumn35,
																											ultraGridColumn36});
			this.cmbDDCanton.DisplayLayout.BandsSerializer.Add(ultraGridBand17);
			this.cmbDDCanton.DisplayLayout.Override.RowSelectors = Infragistics.Win.DefaultableBoolean.True;
			this.cmbDDCanton.DisplayMember = "Nombre";
			this.cmbDDCanton.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbDDCanton.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbDDCanton.Enabled = false;
			this.cmbDDCanton.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbDDCanton.Location = new System.Drawing.Point(176, 32);
			this.cmbDDCanton.Name = "cmbDDCanton";
			this.cmbDDCanton.Size = new System.Drawing.Size(150, 22);
			this.cmbDDCanton.TabIndex = 776;
			this.cmbDDCanton.ValueMember = "idCanton";
			this.cmbDDCanton.KeyDown += new System.Windows.Forms.KeyEventHandler(this.cmbDDCanton_KeyDown);
			this.cmbDDCanton.ValueChanged += new System.EventHandler(this.cmbDDCanton_ValueChanged);
			// 
			// cmbDDProvincia
			// 
			this.cmbDDProvincia.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbDDProvincia.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbDDProvincia.DataSource = this.ultraDataSource1;
			ultraGridColumn37.Header.VisiblePosition = 0;
			ultraGridColumn37.Hidden = true;
			ultraGridColumn37.Width = 50;
			ultraGridColumn38.Header.VisiblePosition = 1;
			ultraGridColumn38.Width = 150;
			ultraGridBand18.Columns.AddRange(new object[] {
																											ultraGridColumn37,
																											ultraGridColumn38});
			this.cmbDDProvincia.DisplayLayout.BandsSerializer.Add(ultraGridBand18);
			this.cmbDDProvincia.DisplayMember = "Nombre";
			this.cmbDDProvincia.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbDDProvincia.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbDDProvincia.Enabled = false;
			this.cmbDDProvincia.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbDDProvincia.Location = new System.Drawing.Point(16, 32);
			this.cmbDDProvincia.Name = "cmbDDProvincia";
			this.cmbDDProvincia.Size = new System.Drawing.Size(150, 22);
			this.cmbDDProvincia.TabIndex = 774;
			this.cmbDDProvincia.ValueMember = "idProvincia";
			this.cmbDDProvincia.KeyDown += new System.Windows.Forms.KeyEventHandler(this.cmbDDProvincia_KeyDown);
			this.cmbDDProvincia.ValueChanged += new System.EventHandler(this.cmbDDProvincia_ValueChanged);
			// 
			// label6
			// 
			this.label6.AutoSize = true;
			this.label6.BackColor = System.Drawing.Color.Transparent;
			this.label6.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label6.Location = new System.Drawing.Point(16, 56);
			this.label6.Name = "label6";
			this.label6.Size = new System.Drawing.Size(73, 17);
			this.label6.TabIndex = 820;
			this.label6.Text = "Calle Principal";
			this.label6.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// tpComplemento
			// 
			this.tpComplemento.Controls.Add(this.groupBox3);
			this.tpComplemento.Controls.Add(this.groupBox4);
			this.tpComplemento.Location = new System.Drawing.Point(-10000, -10000);
			this.tpComplemento.Name = "tpComplemento";
			this.tpComplemento.Size = new System.Drawing.Size(1116, 366);
			this.tpComplemento.Paint += new System.Windows.Forms.PaintEventHandler(this.tpComplemento_Paint);
			// 
			// groupBox3
			// 
			this.groupBox3.BackColor = System.Drawing.SystemColors.Control;
			this.groupBox3.Controls.Add(this.cmbDPCActividadE);
			this.groupBox3.Controls.Add(this.label43);
			this.groupBox3.Controls.Add(this.label55);
			this.groupBox3.Controls.Add(this.label56);
			this.groupBox3.Controls.Add(this.label57);
			this.groupBox3.Controls.Add(this.txtDPCsegundoN);
			this.groupBox3.Controls.Add(this.txtDPCPrimerN);
			this.groupBox3.Controls.Add(this.txtDPCApellidoM);
			this.groupBox3.Controls.Add(this.txtDPCApellidoP);
			this.groupBox3.Controls.Add(this.label41);
			this.groupBox3.Controls.Add(this.label42);
			this.groupBox3.Controls.Add(this.cmbDPCTipodocumento);
			this.groupBox3.Controls.Add(this.txtDPCdocumento);
			this.groupBox3.Controls.Add(this.cmbDPCSituacionL);
			this.groupBox3.Controls.Add(this.cmbDPCProfesion);
			this.groupBox3.Controls.Add(this.cmbDPCNivelEducacion);
			this.groupBox3.Controls.Add(this.cmbDPCSexo);
			this.groupBox3.Controls.Add(this.label44);
			this.groupBox3.Controls.Add(this.label45);
			this.groupBox3.Controls.Add(this.label46);
			this.groupBox3.Controls.Add(this.label47);
			this.groupBox3.Controls.Add(this.label50);
			this.groupBox3.Controls.Add(this.label51);
			this.groupBox3.Controls.Add(this.label52);
			this.groupBox3.Controls.Add(this.label53);
			this.groupBox3.Controls.Add(this.cmbDPCCantonN);
			this.groupBox3.Controls.Add(this.cmbDPCProvinciaN);
			this.groupBox3.Controls.Add(this.dtDPCFnacimiento);
			this.groupBox3.Controls.Add(this.cmbDPCNacionalidad);
			this.groupBox3.Controls.Add(this.label54);
			this.groupBox3.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.groupBox3.Location = new System.Drawing.Point(0, 0);
			this.groupBox3.Name = "groupBox3";
			this.groupBox3.Size = new System.Drawing.Size(1120, 144);
			this.groupBox3.TabIndex = 198;
			this.groupBox3.TabStop = false;
			this.groupBox3.Text = "Datos Personales";
			this.groupBox3.Enter += new System.EventHandler(this.groupBox3_Enter);
			// 
			// cmbDPCActividadE
			// 
			this.cmbDPCActividadE.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbDPCActividadE.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbDPCActividadE.DataSource = this.ultraDataSource40;
			ultraGridColumn39.Header.VisiblePosition = 0;
			ultraGridColumn39.Hidden = true;
			ultraGridColumn40.Header.VisiblePosition = 1;
			ultraGridColumn40.Width = 485;
			ultraGridBand19.Columns.AddRange(new object[] {
																											ultraGridColumn39,
																											ultraGridColumn40});
			this.cmbDPCActividadE.DisplayLayout.BandsSerializer.Add(ultraGridBand19);
			this.cmbDPCActividadE.DisplayMember = "Nombre";
			this.cmbDPCActividadE.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbDPCActividadE.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbDPCActividadE.Enabled = false;
			this.cmbDPCActividadE.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbDPCActividadE.Location = new System.Drawing.Point(488, 112);
			this.cmbDPCActividadE.Name = "cmbDPCActividadE";
			this.cmbDPCActividadE.Size = new System.Drawing.Size(336, 22);
			this.cmbDPCActividadE.TabIndex = 830;
			this.cmbDPCActividadE.ValueMember = "idActEconomica";
			this.cmbDPCActividadE.Visible = false;
			this.cmbDPCActividadE.KeyDown += new System.Windows.Forms.KeyEventHandler(this.cmbDPCActividadE_KeyDown_2);
			this.cmbDPCActividadE.InitializeLayout += new Infragistics.Win.UltraWinGrid.InitializeLayoutEventHandler(this.cmbDPCActividadE_InitializeLayout);
			// 
			// label43
			// 
			this.label43.AutoSize = true;
			this.label43.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label43.Location = new System.Drawing.Point(673, 16);
			this.label43.Name = "label43";
			this.label43.Size = new System.Drawing.Size(91, 17);
			this.label43.TabIndex = 824;
			this.label43.Text = "Segundo Nombre";
			this.label43.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// label55
			// 
			this.label55.AutoSize = true;
			this.label55.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label55.Location = new System.Drawing.Point(493, 16);
			this.label55.Name = "label55";
			this.label55.Size = new System.Drawing.Size(80, 17);
			this.label55.TabIndex = 823;
			this.label55.Text = "Primer Nombre";
			this.label55.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// label56
			// 
			this.label56.AutoSize = true;
			this.label56.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label56.Location = new System.Drawing.Point(856, 16);
			this.label56.Name = "label56";
			this.label56.Size = new System.Drawing.Size(88, 17);
			this.label56.TabIndex = 822;
			this.label56.Text = "Apellido Materno";
			this.label56.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.label56.Visible = false;
			// 
			// label57
			// 
			this.label57.AutoSize = true;
			this.label57.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label57.Location = new System.Drawing.Point(304, 16);
			this.label57.Name = "label57";
			this.label57.Size = new System.Drawing.Size(85, 17);
			this.label57.TabIndex = 821;
			this.label57.Text = "Apellido Paterno";
			this.label57.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// txtDPCsegundoN
			// 
			appearance52.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtDPCsegundoN.Appearance = appearance52;
			this.txtDPCsegundoN.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtDPCsegundoN.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtDPCsegundoN.Enabled = false;
			this.txtDPCsegundoN.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.txtDPCsegundoN.Location = new System.Drawing.Point(672, 32);
			this.txtDPCsegundoN.Name = "txtDPCsegundoN";
			this.txtDPCsegundoN.Size = new System.Drawing.Size(180, 22);
			this.txtDPCsegundoN.TabIndex = 828;
			this.txtDPCsegundoN.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtDPCsegundoN_KeyDown);
			this.txtDPCsegundoN.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.txtDPCsegundoN_KeyPress);
			// 
			// txtDPCPrimerN
			// 
			appearance53.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtDPCPrimerN.Appearance = appearance53;
			this.txtDPCPrimerN.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtDPCPrimerN.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtDPCPrimerN.Enabled = false;
			this.txtDPCPrimerN.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.txtDPCPrimerN.Location = new System.Drawing.Point(486, 32);
			this.txtDPCPrimerN.Name = "txtDPCPrimerN";
			this.txtDPCPrimerN.Size = new System.Drawing.Size(180, 22);
			this.txtDPCPrimerN.TabIndex = 827;
			this.txtDPCPrimerN.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtDPCPrimerN_KeyDown);
			this.txtDPCPrimerN.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.txtDPCPrimerN_KeyPress);
			// 
			// txtDPCApellidoM
			// 
			appearance54.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtDPCApellidoM.Appearance = appearance54;
			this.txtDPCApellidoM.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtDPCApellidoM.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtDPCApellidoM.Enabled = false;
			this.txtDPCApellidoM.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.txtDPCApellidoM.Location = new System.Drawing.Point(856, 32);
			this.txtDPCApellidoM.Name = "txtDPCApellidoM";
			this.txtDPCApellidoM.Size = new System.Drawing.Size(180, 22);
			this.txtDPCApellidoM.TabIndex = 826;
			this.txtDPCApellidoM.Visible = false;
			this.txtDPCApellidoM.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtDPCApellidoM_KeyDown);
			this.txtDPCApellidoM.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.txtDPCApellidoM_KeyPress);
			this.txtDPCApellidoM.ValueChanged += new System.EventHandler(this.txtDPCApellidoM_ValueChanged);
			// 
			// txtDPCApellidoP
			// 
			appearance55.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtDPCApellidoP.Appearance = appearance55;
			this.txtDPCApellidoP.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtDPCApellidoP.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtDPCApellidoP.Enabled = false;
			this.txtDPCApellidoP.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.txtDPCApellidoP.Location = new System.Drawing.Point(299, 32);
			this.txtDPCApellidoP.Name = "txtDPCApellidoP";
			this.txtDPCApellidoP.Size = new System.Drawing.Size(180, 22);
			this.txtDPCApellidoP.TabIndex = 825;
			this.txtDPCApellidoP.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtDPCApellidoP_KeyDown);
			this.txtDPCApellidoP.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.txtDPCApellidoP_KeyPress);
			// 
			// label41
			// 
			this.label41.AutoSize = true;
			this.label41.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label41.Location = new System.Drawing.Point(168, 16);
			this.label41.Name = "label41";
			this.label41.Size = new System.Drawing.Size(61, 17);
			this.label41.TabIndex = 818;
			this.label41.Text = "Documento";
			this.label41.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// label42
			// 
			this.label42.AutoSize = true;
			this.label42.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label42.Location = new System.Drawing.Point(8, 16);
			this.label42.Name = "label42";
			this.label42.Size = new System.Drawing.Size(86, 17);
			this.label42.TabIndex = 817;
			this.label42.Text = "Tipo Documento";
			this.label42.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// cmbDPCTipodocumento
			// 
			this.cmbDPCTipodocumento.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbDPCTipodocumento.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbDPCTipodocumento.DisplayMember = "Nombre";
			this.cmbDPCTipodocumento.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbDPCTipodocumento.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbDPCTipodocumento.Enabled = false;
			this.cmbDPCTipodocumento.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbDPCTipodocumento.Location = new System.Drawing.Point(8, 32);
			this.cmbDPCTipodocumento.Name = "cmbDPCTipodocumento";
			this.cmbDPCTipodocumento.Size = new System.Drawing.Size(150, 22);
			this.cmbDPCTipodocumento.TabIndex = 820;
			this.cmbDPCTipodocumento.ValueMember = "idTipoDoc";
			// 
			// txtDPCdocumento
			// 
			appearance56.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtDPCdocumento.Appearance = appearance56;
			this.txtDPCdocumento.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtDPCdocumento.Enabled = false;
			this.txtDPCdocumento.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.txtDPCdocumento.Location = new System.Drawing.Point(165, 32);
			this.txtDPCdocumento.MaxLength = 10;
			this.txtDPCdocumento.Name = "txtDPCdocumento";
			this.txtDPCdocumento.Size = new System.Drawing.Size(127, 22);
			this.txtDPCdocumento.TabIndex = 819;
			this.txtDPCdocumento.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtDPCdocumento_KeyDown);
			// 
			// cmbDPCSituacionL
			// 
			this.cmbDPCSituacionL.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbDPCSituacionL.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbDPCSituacionL.DataSource = this.ultraDataSource30;
			ultraGridColumn41.Header.VisiblePosition = 0;
			ultraGridColumn41.Hidden = true;
			ultraGridColumn42.Header.VisiblePosition = 1;
			ultraGridColumn42.Width = 150;
			ultraGridBand20.Columns.AddRange(new object[] {
																											ultraGridColumn41,
																											ultraGridColumn42});
			this.cmbDPCSituacionL.DisplayLayout.BandsSerializer.Add(ultraGridBand20);
			this.cmbDPCSituacionL.DisplayMember = "Descripcion";
			this.cmbDPCSituacionL.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbDPCSituacionL.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbDPCSituacionL.Enabled = false;
			this.cmbDPCSituacionL.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbDPCSituacionL.Location = new System.Drawing.Point(328, 112);
			this.cmbDPCSituacionL.Name = "cmbDPCSituacionL";
			this.cmbDPCSituacionL.Size = new System.Drawing.Size(150, 22);
			this.cmbDPCSituacionL.TabIndex = 816;
			this.cmbDPCSituacionL.ValueMember = "idSituacionLaboral";
			this.cmbDPCSituacionL.Visible = false;
			this.cmbDPCSituacionL.KeyDown += new System.Windows.Forms.KeyEventHandler(this.cmbDPCSituacionL_KeyDown);
			this.cmbDPCSituacionL.ValueChanged += new System.EventHandler(this.cmbDPCSituacionL_ValueChanged);
			// 
			// cmbDPCProfesion
			// 
			this.cmbDPCProfesion.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbDPCProfesion.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbDPCProfesion.DataSource = this.ultraDataSource29;
			ultraGridColumn43.Header.VisiblePosition = 0;
			ultraGridColumn43.Hidden = true;
			ultraGridColumn44.Header.VisiblePosition = 1;
			ultraGridColumn44.Width = 150;
			ultraGridBand21.Columns.AddRange(new object[] {
																											ultraGridColumn43,
																											ultraGridColumn44});
			this.cmbDPCProfesion.DisplayLayout.BandsSerializer.Add(ultraGridBand21);
			this.cmbDPCProfesion.DisplayMember = "Nombre";
			this.cmbDPCProfesion.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbDPCProfesion.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbDPCProfesion.Enabled = false;
			this.cmbDPCProfesion.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbDPCProfesion.Location = new System.Drawing.Point(168, 112);
			this.cmbDPCProfesion.Name = "cmbDPCProfesion";
			this.cmbDPCProfesion.Size = new System.Drawing.Size(150, 22);
			this.cmbDPCProfesion.TabIndex = 815;
			this.cmbDPCProfesion.ValueMember = "idProfesion";
			this.cmbDPCProfesion.KeyDown += new System.Windows.Forms.KeyEventHandler(this.cmbDPCProfesion_KeyDown);
			// 
			// cmbDPCNivelEducacion
			// 
			this.cmbDPCNivelEducacion.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbDPCNivelEducacion.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbDPCNivelEducacion.DataSource = this.ultraDataSource33;
			ultraGridColumn45.Header.VisiblePosition = 0;
			ultraGridColumn45.Hidden = true;
			ultraGridColumn46.Header.VisiblePosition = 1;
			ultraGridColumn46.Width = 150;
			ultraGridBand22.Columns.AddRange(new object[] {
																											ultraGridColumn45,
																											ultraGridColumn46});
			this.cmbDPCNivelEducacion.DisplayLayout.BandsSerializer.Add(ultraGridBand22);
			this.cmbDPCNivelEducacion.DisplayMember = "Nombre";
			this.cmbDPCNivelEducacion.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbDPCNivelEducacion.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbDPCNivelEducacion.Enabled = false;
			this.cmbDPCNivelEducacion.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbDPCNivelEducacion.Location = new System.Drawing.Point(8, 112);
			this.cmbDPCNivelEducacion.Name = "cmbDPCNivelEducacion";
			this.cmbDPCNivelEducacion.Size = new System.Drawing.Size(150, 22);
			this.cmbDPCNivelEducacion.TabIndex = 811;
			this.cmbDPCNivelEducacion.ValueMember = "idNivelEducacion";
			this.cmbDPCNivelEducacion.KeyDown += new System.Windows.Forms.KeyEventHandler(this.cmbDPCNivelEducacion_KeyDown);
			this.cmbDPCNivelEducacion.InitializeRow += new Infragistics.Win.UltraWinGrid.InitializeRowEventHandler(this.cmbDPCNivelEducacion_InitializeRow);
			// 
			// ultraDataSource33
			// 
			ultraDataColumn35.DataType = typeof(int);
			this.ultraDataSource33.Band.Columns.AddRange(new object[] {
																																	ultraDataColumn35,
																																	ultraDataColumn36});
			// 
			// cmbDPCSexo
			// 
			this.cmbDPCSexo.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbDPCSexo.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbDPCSexo.DataSource = this.ultraDataSource24;
			ultraGridColumn47.Header.VisiblePosition = 0;
			ultraGridColumn47.Hidden = true;
			ultraGridColumn48.Header.VisiblePosition = 1;
			ultraGridColumn48.Width = 150;
			ultraGridBand23.Columns.AddRange(new object[] {
																											ultraGridColumn47,
																											ultraGridColumn48});
			this.cmbDPCSexo.DisplayLayout.BandsSerializer.Add(ultraGridBand23);
			this.cmbDPCSexo.DisplayMember = "Nombre";
			this.cmbDPCSexo.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbDPCSexo.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbDPCSexo.Enabled = false;
			this.cmbDPCSexo.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbDPCSexo.Location = new System.Drawing.Point(264, 72);
			this.cmbDPCSexo.Name = "cmbDPCSexo";
			this.cmbDPCSexo.Size = new System.Drawing.Size(150, 22);
			this.cmbDPCSexo.TabIndex = 810;
			this.cmbDPCSexo.ValueMember = "idSexo";
			this.cmbDPCSexo.KeyDown += new System.Windows.Forms.KeyEventHandler(this.cmbDPCSexo_KeyDown);
			// 
			// label44
			// 
			this.label44.AutoSize = true;
			this.label44.BackColor = System.Drawing.Color.Transparent;
			this.label44.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label44.Location = new System.Drawing.Point(488, 96);
			this.label44.Name = "label44";
			this.label44.Size = new System.Drawing.Size(106, 17);
			this.label44.TabIndex = 801;
			this.label44.Text = "Actividad Económica";
			this.label44.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.label44.Visible = false;
			// 
			// label45
			// 
			this.label45.AutoSize = true;
			this.label45.BackColor = System.Drawing.Color.Transparent;
			this.label45.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label45.Location = new System.Drawing.Point(328, 96);
			this.label45.Name = "label45";
			this.label45.Size = new System.Drawing.Size(90, 17);
			this.label45.TabIndex = 800;
			this.label45.Text = "Situación Laboral";
			this.label45.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.label45.Visible = false;
			// 
			// label46
			// 
			this.label46.AutoSize = true;
			this.label46.BackColor = System.Drawing.Color.Transparent;
			this.label46.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label46.Location = new System.Drawing.Point(168, 96);
			this.label46.Name = "label46";
			this.label46.Size = new System.Drawing.Size(50, 17);
			this.label46.TabIndex = 799;
			this.label46.Text = "Profesión";
			this.label46.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// label47
			// 
			this.label47.AutoSize = true;
			this.label47.BackColor = System.Drawing.Color.Transparent;
			this.label47.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label47.Location = new System.Drawing.Point(8, 96);
			this.label47.Name = "label47";
			this.label47.Size = new System.Drawing.Size(82, 17);
			this.label47.TabIndex = 798;
			this.label47.Text = "Nivel Educación";
			this.label47.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// label50
			// 
			this.label50.AutoSize = true;
			this.label50.BackColor = System.Drawing.Color.Transparent;
			this.label50.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label50.Location = new System.Drawing.Point(584, 56);
			this.label50.Name = "label50";
			this.label50.Size = new System.Drawing.Size(99, 17);
			this.label50.TabIndex = 788;
			this.label50.Text = "Cantón Nacimiento";
			this.label50.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.label50.Visible = false;
			// 
			// label51
			// 
			this.label51.AutoSize = true;
			this.label51.BackColor = System.Drawing.Color.Transparent;
			this.label51.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label51.Location = new System.Drawing.Point(424, 56);
			this.label51.Name = "label51";
			this.label51.Size = new System.Drawing.Size(108, 17);
			this.label51.TabIndex = 787;
			this.label51.Text = "Provincia Nacimiento";
			this.label51.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.label51.Visible = false;
			// 
			// label52
			// 
			this.label52.AutoSize = true;
			this.label52.BackColor = System.Drawing.Color.Transparent;
			this.label52.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label52.Location = new System.Drawing.Point(168, 56);
			this.label52.Name = "label52";
			this.label52.Size = new System.Drawing.Size(73, 17);
			this.label52.TabIndex = 786;
			this.label52.Text = "F. Nacimiento";
			this.label52.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// label53
			// 
			this.label53.AutoSize = true;
			this.label53.BackColor = System.Drawing.Color.Transparent;
			this.label53.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label53.Location = new System.Drawing.Point(264, 56);
			this.label53.Name = "label53";
			this.label53.Size = new System.Drawing.Size(28, 17);
			this.label53.TabIndex = 785;
			this.label53.Text = "Sexo";
			this.label53.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// cmbDPCCantonN
			// 
			this.cmbDPCCantonN.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbDPCCantonN.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbDPCCantonN.DataSource = this.ultraDataSource2;
			ultraGridColumn49.Header.VisiblePosition = 0;
			ultraGridColumn49.Hidden = true;
			ultraGridColumn50.Header.Caption = "Canton";
			ultraGridColumn50.Header.VisiblePosition = 1;
			ultraGridColumn50.Width = 150;
			ultraGridBand24.Columns.AddRange(new object[] {
																											ultraGridColumn49,
																											ultraGridColumn50});
			this.cmbDPCCantonN.DisplayLayout.BandsSerializer.Add(ultraGridBand24);
			this.cmbDPCCantonN.DisplayLayout.Override.RowSelectors = Infragistics.Win.DefaultableBoolean.True;
			this.cmbDPCCantonN.DisplayMember = "Nombre";
			this.cmbDPCCantonN.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbDPCCantonN.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbDPCCantonN.Enabled = false;
			this.cmbDPCCantonN.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbDPCCantonN.Location = new System.Drawing.Point(584, 72);
			this.cmbDPCCantonN.Name = "cmbDPCCantonN";
			this.cmbDPCCantonN.Size = new System.Drawing.Size(150, 22);
			this.cmbDPCCantonN.TabIndex = 784;
			this.cmbDPCCantonN.ValueMember = "idCanton";
			this.cmbDPCCantonN.Visible = false;
			this.cmbDPCCantonN.KeyDown += new System.Windows.Forms.KeyEventHandler(this.cmbDPCCantonN_KeyDown);
			// 
			// cmbDPCProvinciaN
			// 
			this.cmbDPCProvinciaN.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbDPCProvinciaN.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbDPCProvinciaN.DataSource = this.ultraDataSource1;
			ultraGridColumn51.Header.VisiblePosition = 0;
			ultraGridColumn51.Hidden = true;
			ultraGridColumn51.Width = 50;
			ultraGridColumn52.Header.VisiblePosition = 1;
			ultraGridColumn52.Width = 150;
			ultraGridBand25.Columns.AddRange(new object[] {
																											ultraGridColumn51,
																											ultraGridColumn52});
			this.cmbDPCProvinciaN.DisplayLayout.BandsSerializer.Add(ultraGridBand25);
			this.cmbDPCProvinciaN.DisplayMember = "Nombre";
			this.cmbDPCProvinciaN.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbDPCProvinciaN.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbDPCProvinciaN.Enabled = false;
			this.cmbDPCProvinciaN.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbDPCProvinciaN.Location = new System.Drawing.Point(424, 72);
			this.cmbDPCProvinciaN.Name = "cmbDPCProvinciaN";
			this.cmbDPCProvinciaN.Size = new System.Drawing.Size(150, 22);
			this.cmbDPCProvinciaN.TabIndex = 782;
			this.cmbDPCProvinciaN.ValueMember = "idProvincia";
			this.cmbDPCProvinciaN.Visible = false;
			this.cmbDPCProvinciaN.KeyDown += new System.Windows.Forms.KeyEventHandler(this.cmbDPCProvinciaN_KeyDown);
			this.cmbDPCProvinciaN.ValueChanged += new System.EventHandler(this.cmbDPCProvinciaN_ValueChanged);
			// 
			// dtDPCFnacimiento
			// 
			this.dtDPCFnacimiento.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			dateButton3.Caption = "Today";
			this.dtDPCFnacimiento.DateButtons.Add(dateButton3);
			this.dtDPCFnacimiento.Enabled = false;
			this.dtDPCFnacimiento.Format = "dd/MM/yyyy";
			this.dtDPCFnacimiento.Location = new System.Drawing.Point(168, 72);
			this.dtDPCFnacimiento.Name = "dtDPCFnacimiento";
			this.dtDPCFnacimiento.NonAutoSizeHeight = 23;
			this.dtDPCFnacimiento.Size = new System.Drawing.Size(90, 21);
			this.dtDPCFnacimiento.TabIndex = 779;
			this.dtDPCFnacimiento.Value = ((object)(resources.GetObject("dtDPCFnacimiento.Value")));
			this.dtDPCFnacimiento.KeyDown += new System.Windows.Forms.KeyEventHandler(this.dtDPCFnacimiento_KeyDown);
			// 
			// cmbDPCNacionalidad
			// 
			this.cmbDPCNacionalidad.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbDPCNacionalidad.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbDPCNacionalidad.DataSource = this.ultraDataSource26;
			ultraGridBand26.AddButtonCaption = "RolSucursal";
			ultraGridColumn53.Header.VisiblePosition = 0;
			ultraGridColumn53.Hidden = true;
			ultraGridColumn54.Header.VisiblePosition = 1;
			ultraGridColumn54.Width = 150;
			ultraGridBand26.Columns.AddRange(new object[] {
																											ultraGridColumn53,
																											ultraGridColumn54});
			this.cmbDPCNacionalidad.DisplayLayout.BandsSerializer.Add(ultraGridBand26);
			this.cmbDPCNacionalidad.DisplayMember = "Nombre";
			this.cmbDPCNacionalidad.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbDPCNacionalidad.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbDPCNacionalidad.Enabled = false;
			this.cmbDPCNacionalidad.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbDPCNacionalidad.Location = new System.Drawing.Point(8, 72);
			this.cmbDPCNacionalidad.Name = "cmbDPCNacionalidad";
			this.cmbDPCNacionalidad.Size = new System.Drawing.Size(150, 22);
			this.cmbDPCNacionalidad.TabIndex = 354;
			this.cmbDPCNacionalidad.ValueMember = "idNacionalidad";
			this.cmbDPCNacionalidad.KeyDown += new System.Windows.Forms.KeyEventHandler(this.cmbDPCNacionalidad_KeyDown);
			// 
			// label54
			// 
			this.label54.AutoSize = true;
			this.label54.BackColor = System.Drawing.Color.Transparent;
			this.label54.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label54.Location = new System.Drawing.Point(8, 56);
			this.label54.Name = "label54";
			this.label54.Size = new System.Drawing.Size(68, 17);
			this.label54.TabIndex = 353;
			this.label54.Text = "Nacionalidad";
			this.label54.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// groupBox4
			// 
			this.groupBox4.BackColor = System.Drawing.SystemColors.Control;
			this.groupBox4.Controls.Add(this.label62);
			this.groupBox4.Controls.Add(this.label61);
			this.groupBox4.Controls.Add(this.label60);
			this.groupBox4.Controls.Add(this.label59);
			this.groupBox4.Controls.Add(this.txtDPCTPasivos);
			this.groupBox4.Controls.Add(this.txtDPCTActivos);
			this.groupBox4.Controls.Add(this.txtDPCTGastos);
			this.groupBox4.Controls.Add(this.txtDPCTIngresos);
			this.groupBox4.Controls.Add(this.label58);
			this.groupBox4.Controls.Add(this.label49);
			this.groupBox4.Controls.Add(this.cmbDPCTCargo);
			this.groupBox4.Controls.Add(this.cmbDPCTActividadEmpresa);
			this.groupBox4.Controls.Add(this.label48);
			this.groupBox4.Controls.Add(this.txtDPCTNombreEmpresa);
			this.groupBox4.Controls.Add(this.label63);
			this.groupBox4.Controls.Add(this.txtDPCTReferencia);
			this.groupBox4.Controls.Add(this.btnConyugueCelular);
			this.groupBox4.Controls.Add(this.btnConyugueTelefono);
			this.groupBox4.Controls.Add(this.label64);
			this.groupBox4.Controls.Add(this.label65);
			this.groupBox4.Controls.Add(this.label66);
			this.groupBox4.Controls.Add(this.label67);
			this.groupBox4.Controls.Add(this.label68);
			this.groupBox4.Controls.Add(this.label69);
			this.groupBox4.Controls.Add(this.label70);
			this.groupBox4.Controls.Add(this.txtDPCTEmail);
			this.groupBox4.Controls.Add(this.txtDPCTCelularC);
			this.groupBox4.Controls.Add(this.txtDPCTExt);
			this.groupBox4.Controls.Add(this.txtDPCTtelefono);
			this.groupBox4.Controls.Add(this.txtDPCTCalleSecundaria);
			this.groupBox4.Controls.Add(this.txtDPCTNumeroCasa);
			this.groupBox4.Controls.Add(this.txtDPCTCallePrincipal);
			this.groupBox4.Controls.Add(this.label71);
			this.groupBox4.Controls.Add(this.label72);
			this.groupBox4.Controls.Add(this.label73);
			this.groupBox4.Controls.Add(this.label74);
			this.groupBox4.Controls.Add(this.cmbDPCTBarrio);
			this.groupBox4.Controls.Add(this.cmbDPCTParroquia);
			this.groupBox4.Controls.Add(this.cmbDPCTCanton);
			this.groupBox4.Controls.Add(this.cmbDPCTProvincia);
			this.groupBox4.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.groupBox4.Location = new System.Drawing.Point(0, 144);
			this.groupBox4.Name = "groupBox4";
			this.groupBox4.Size = new System.Drawing.Size(1120, 224);
			this.groupBox4.TabIndex = 199;
			this.groupBox4.TabStop = false;
			this.groupBox4.Text = "Datos de Trabajo";
			this.groupBox4.Visible = false;
			this.groupBox4.Enter += new System.EventHandler(this.groupBox4_Enter);
			// 
			// label62
			// 
			this.label62.AutoSize = true;
			this.label62.BackColor = System.Drawing.Color.Transparent;
			this.label62.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label62.Location = new System.Drawing.Point(608, 176);
			this.label62.Name = "label62";
			this.label62.Size = new System.Drawing.Size(41, 17);
			this.label62.TabIndex = 845;
			this.label62.Text = "Pasivos";
			this.label62.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.label62.Visible = false;
			// 
			// label61
			// 
			this.label61.AutoSize = true;
			this.label61.BackColor = System.Drawing.Color.Transparent;
			this.label61.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label61.Location = new System.Drawing.Point(512, 176);
			this.label61.Name = "label61";
			this.label61.Size = new System.Drawing.Size(39, 17);
			this.label61.TabIndex = 844;
			this.label61.Text = "Activos";
			this.label61.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.label61.Visible = false;
			// 
			// label60
			// 
			this.label60.AutoSize = true;
			this.label60.BackColor = System.Drawing.Color.Transparent;
			this.label60.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label60.Location = new System.Drawing.Point(416, 176);
			this.label60.Name = "label60";
			this.label60.Size = new System.Drawing.Size(38, 17);
			this.label60.TabIndex = 843;
			this.label60.Text = "Gastos";
			this.label60.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// label59
			// 
			this.label59.AutoSize = true;
			this.label59.BackColor = System.Drawing.Color.Transparent;
			this.label59.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label59.Location = new System.Drawing.Point(320, 176);
			this.label59.Name = "label59";
			this.label59.Size = new System.Drawing.Size(47, 17);
			this.label59.TabIndex = 842;
			this.label59.Text = "Ingresos";
			this.label59.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// txtDPCTPasivos
			// 
			this.txtDPCTPasivos.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtDPCTPasivos.Enabled = false;
			this.txtDPCTPasivos.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.txtDPCTPasivos.Location = new System.Drawing.Point(608, 192);
			this.txtDPCTPasivos.MaxValue = 10000;
			this.txtDPCTPasivos.MinValue = 0;
			this.txtDPCTPasivos.Name = "txtDPCTPasivos";
			this.txtDPCTPasivos.PromptChar = ' ';
			this.txtDPCTPasivos.Size = new System.Drawing.Size(90, 22);
			this.txtDPCTPasivos.SpinButtonDisplayStyle = Infragistics.Win.ButtonDisplayStyle.Always;
			this.txtDPCTPasivos.TabIndex = 841;
			this.txtDPCTPasivos.Value = 100;
			this.txtDPCTPasivos.Visible = false;
			this.txtDPCTPasivos.Click += new System.EventHandler(this.txtDPCTPasivos_Click);
			// 
			// txtDPCTActivos
			// 
			this.txtDPCTActivos.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtDPCTActivos.Enabled = false;
			this.txtDPCTActivos.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.txtDPCTActivos.Location = new System.Drawing.Point(512, 192);
			this.txtDPCTActivos.MaxValue = 10000;
			this.txtDPCTActivos.MinValue = 0;
			this.txtDPCTActivos.Name = "txtDPCTActivos";
			this.txtDPCTActivos.PromptChar = ' ';
			this.txtDPCTActivos.Size = new System.Drawing.Size(90, 22);
			this.txtDPCTActivos.SpinButtonDisplayStyle = Infragistics.Win.ButtonDisplayStyle.Always;
			this.txtDPCTActivos.TabIndex = 840;
			this.txtDPCTActivos.Value = 450;
			this.txtDPCTActivos.Visible = false;
			this.txtDPCTActivos.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtDPCTActivos_KeyDown);
			this.txtDPCTActivos.Click += new System.EventHandler(this.txtDPCTActivos_Click);
			// 
			// txtDPCTGastos
			// 
			this.txtDPCTGastos.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtDPCTGastos.Enabled = false;
			this.txtDPCTGastos.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.txtDPCTGastos.Location = new System.Drawing.Point(416, 192);
			this.txtDPCTGastos.MaxValue = 10000;
			this.txtDPCTGastos.MinValue = 0;
			this.txtDPCTGastos.Name = "txtDPCTGastos";
			this.txtDPCTGastos.PromptChar = ' ';
			this.txtDPCTGastos.Size = new System.Drawing.Size(90, 22);
			this.txtDPCTGastos.SpinButtonDisplayStyle = Infragistics.Win.ButtonDisplayStyle.Always;
			this.txtDPCTGastos.TabIndex = 839;
			this.txtDPCTGastos.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtDPCTGastos_KeyDown);
			this.txtDPCTGastos.Click += new System.EventHandler(this.txtDPCTGastos_Click);
			this.txtDPCTGastos.ValueChanged += new System.EventHandler(this.txtDPCTGastos_ValueChanged);
			// 
			// txtDPCTIngresos
			// 
			this.txtDPCTIngresos.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtDPCTIngresos.Enabled = false;
			this.txtDPCTIngresos.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.txtDPCTIngresos.Location = new System.Drawing.Point(320, 192);
			this.txtDPCTIngresos.MaxValue = 10000;
			this.txtDPCTIngresos.MinValue = 0;
			this.txtDPCTIngresos.Name = "txtDPCTIngresos";
			this.txtDPCTIngresos.PromptChar = ' ';
			this.txtDPCTIngresos.Size = new System.Drawing.Size(90, 22);
			this.txtDPCTIngresos.SpinButtonDisplayStyle = Infragistics.Win.ButtonDisplayStyle.Always;
			this.txtDPCTIngresos.TabIndex = 838;
			this.txtDPCTIngresos.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtDPCTIngresos_KeyDown);
			this.txtDPCTIngresos.Click += new System.EventHandler(this.txtDPCTIngresos_Click);
			// 
			// label58
			// 
			this.label58.AutoSize = true;
			this.label58.BackColor = System.Drawing.Color.Transparent;
			this.label58.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label58.Location = new System.Drawing.Point(496, 16);
			this.label58.Name = "label58";
			this.label58.Size = new System.Drawing.Size(33, 17);
			this.label58.TabIndex = 837;
			this.label58.Text = "Cargo";
			this.label58.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// label49
			// 
			this.label49.AutoSize = true;
			this.label49.BackColor = System.Drawing.Color.Transparent;
			this.label49.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label49.Location = new System.Drawing.Point(328, 16);
			this.label49.Name = "label49";
			this.label49.Size = new System.Drawing.Size(96, 17);
			this.label49.TabIndex = 836;
			this.label49.Text = "Actividad Empresa";
			this.label49.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// cmbDPCTCargo
			// 
			this.cmbDPCTCargo.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbDPCTCargo.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbDPCTCargo.DataSource = this.ultraDataSource36;
			ultraGridColumn55.Header.VisiblePosition = 0;
			ultraGridColumn55.Hidden = true;
			ultraGridColumn55.Width = 150;
			ultraGridColumn56.Header.VisiblePosition = 1;
			ultraGridBand27.Columns.AddRange(new object[] {
																											ultraGridColumn55,
																											ultraGridColumn56});
			this.cmbDPCTCargo.DisplayLayout.BandsSerializer.Add(ultraGridBand27);
			this.cmbDPCTCargo.DisplayMember = "Nombre";
			this.cmbDPCTCargo.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbDPCTCargo.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbDPCTCargo.Enabled = false;
			this.cmbDPCTCargo.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbDPCTCargo.Location = new System.Drawing.Point(496, 32);
			this.cmbDPCTCargo.Name = "cmbDPCTCargo";
			this.cmbDPCTCargo.Size = new System.Drawing.Size(150, 22);
			this.cmbDPCTCargo.TabIndex = 835;
			this.cmbDPCTCargo.ValueMember = "idCargo";
			this.cmbDPCTCargo.KeyDown += new System.Windows.Forms.KeyEventHandler(this.cmbDPCTCargo_KeyDown);
			this.cmbDPCTCargo.InitializeLayout += new Infragistics.Win.UltraWinGrid.InitializeLayoutEventHandler(this.cmbDPCCargo_InitializeLayout);
			// 
			// ultraDataSource36
			// 
			ultraDataColumn37.DataType = typeof(int);
			this.ultraDataSource36.Band.Columns.AddRange(new object[] {
																																	ultraDataColumn37,
																																	ultraDataColumn38});
			// 
			// cmbDPCTActividadEmpresa
			// 
			this.cmbDPCTActividadEmpresa.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbDPCTActividadEmpresa.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbDPCTActividadEmpresa.DataSource = this.ultraDataSource35;
			ultraGridColumn57.Header.VisiblePosition = 0;
			ultraGridColumn57.Hidden = true;
			ultraGridColumn57.Width = 150;
			ultraGridColumn58.Header.VisiblePosition = 1;
			ultraGridColumn58.Width = 150;
			ultraGridBand28.Columns.AddRange(new object[] {
																											ultraGridColumn57,
																											ultraGridColumn58});
			this.cmbDPCTActividadEmpresa.DisplayLayout.BandsSerializer.Add(ultraGridBand28);
			this.cmbDPCTActividadEmpresa.DisplayMember = "Nombre";
			this.cmbDPCTActividadEmpresa.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbDPCTActividadEmpresa.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbDPCTActividadEmpresa.Enabled = false;
			this.cmbDPCTActividadEmpresa.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbDPCTActividadEmpresa.Location = new System.Drawing.Point(328, 32);
			this.cmbDPCTActividadEmpresa.Name = "cmbDPCTActividadEmpresa";
			this.cmbDPCTActividadEmpresa.Size = new System.Drawing.Size(150, 22);
			this.cmbDPCTActividadEmpresa.TabIndex = 834;
			this.cmbDPCTActividadEmpresa.ValueMember = "idActEconomica";
			this.cmbDPCTActividadEmpresa.KeyDown += new System.Windows.Forms.KeyEventHandler(this.cmbDPCTActividadEmpresa_KeyDown);
			// 
			// ultraDataSource35
			// 
			ultraDataColumn39.DataType = typeof(int);
			this.ultraDataSource35.Band.Columns.AddRange(new object[] {
																																	ultraDataColumn39,
																																	ultraDataColumn40});
			// 
			// label48
			// 
			this.label48.AutoSize = true;
			this.label48.BackColor = System.Drawing.Color.Transparent;
			this.label48.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label48.Location = new System.Drawing.Point(16, 16);
			this.label48.Name = "label48";
			this.label48.Size = new System.Drawing.Size(132, 17);
			this.label48.TabIndex = 833;
			this.label48.Text = "Nombre Empresa Trabajo";
			this.label48.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// txtDPCTNombreEmpresa
			// 
			appearance57.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtDPCTNombreEmpresa.Appearance = appearance57;
			this.txtDPCTNombreEmpresa.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtDPCTNombreEmpresa.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtDPCTNombreEmpresa.Enabled = false;
			this.txtDPCTNombreEmpresa.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.txtDPCTNombreEmpresa.Location = new System.Drawing.Point(16, 32);
			this.txtDPCTNombreEmpresa.Name = "txtDPCTNombreEmpresa";
			this.txtDPCTNombreEmpresa.Size = new System.Drawing.Size(300, 22);
			this.txtDPCTNombreEmpresa.TabIndex = 832;
			this.txtDPCTNombreEmpresa.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtDPCTNombreEmpresa_KeyDown);
			this.txtDPCTNombreEmpresa.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.txtDPCTNombreEmpresa_KeyPress);
			// 
			// label63
			// 
			this.label63.AutoSize = true;
			this.label63.BackColor = System.Drawing.Color.Transparent;
			this.label63.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label63.Location = new System.Drawing.Point(16, 136);
			this.label63.Name = "label63";
			this.label63.Size = new System.Drawing.Size(108, 17);
			this.label63.TabIndex = 831;
			this.label63.Text = "Referencia Ubicacion";
			this.label63.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// txtDPCTReferencia
			// 
			this.txtDPCTReferencia.AcceptsReturn = true;
			appearance58.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtDPCTReferencia.Appearance = appearance58;
			this.txtDPCTReferencia.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtDPCTReferencia.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtDPCTReferencia.Enabled = false;
			this.txtDPCTReferencia.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.txtDPCTReferencia.Location = new System.Drawing.Point(16, 152);
			this.txtDPCTReferencia.Multiline = true;
			this.txtDPCTReferencia.Name = "txtDPCTReferencia";
			this.txtDPCTReferencia.Size = new System.Drawing.Size(300, 66);
			this.txtDPCTReferencia.TabIndex = 830;
			this.txtDPCTReferencia.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtDPCTReferencia_KeyDown);
			this.txtDPCTReferencia.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.txtDPCTReferencia_KeyPress);
			// 
			// btnConyugueCelular
			// 
			this.btnConyugueCelular.Enabled = false;
			this.btnConyugueCelular.Image = ((System.Drawing.Image)(resources.GetObject("btnConyugueCelular.Image")));
			this.btnConyugueCelular.Location = new System.Drawing.Point(656, 152);
			this.btnConyugueCelular.Name = "btnConyugueCelular";
			this.btnConyugueCelular.Size = new System.Drawing.Size(24, 23);
			this.btnConyugueCelular.TabIndex = 829;
			this.btnConyugueCelular.Click += new System.EventHandler(this.btnConyugueCelular_Click);
			// 
			// btnConyugueTelefono
			// 
			this.btnConyugueTelefono.Enabled = false;
			this.btnConyugueTelefono.Image = ((System.Drawing.Image)(resources.GetObject("btnConyugueTelefono.Image")));
			this.btnConyugueTelefono.Location = new System.Drawing.Point(424, 152);
			this.btnConyugueTelefono.Name = "btnConyugueTelefono";
			this.btnConyugueTelefono.Size = new System.Drawing.Size(24, 23);
			this.btnConyugueTelefono.TabIndex = 827;
			this.btnConyugueTelefono.Click += new System.EventHandler(this.btnConyugueTelefono_Click);
			// 
			// label64
			// 
			this.label64.AutoSize = true;
			this.label64.BackColor = System.Drawing.Color.Transparent;
			this.label64.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label64.Location = new System.Drawing.Point(656, 56);
			this.label64.Name = "label64";
			this.label64.Size = new System.Drawing.Size(31, 17);
			this.label64.TabIndex = 826;
			this.label64.Text = "Email";
			this.label64.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// label65
			// 
			this.label65.AutoSize = true;
			this.label65.BackColor = System.Drawing.Color.Transparent;
			this.label65.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label65.Location = new System.Drawing.Point(544, 136);
			this.label65.Name = "label65";
			this.label65.Size = new System.Drawing.Size(85, 17);
			this.label65.TabIndex = 825;
			this.label65.Text = "Celular Conyuge";
			this.label65.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// label66
			// 
			this.label66.AutoSize = true;
			this.label66.BackColor = System.Drawing.Color.Transparent;
			this.label66.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label66.Location = new System.Drawing.Point(456, 136);
			this.label66.Name = "label66";
			this.label66.Size = new System.Drawing.Size(20, 17);
			this.label66.TabIndex = 824;
			this.label66.Text = "Ext";
			this.label66.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// label67
			// 
			this.label67.AutoSize = true;
			this.label67.BackColor = System.Drawing.Color.Transparent;
			this.label67.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label67.Location = new System.Drawing.Point(320, 136);
			this.label67.Name = "label67";
			this.label67.Size = new System.Drawing.Size(48, 17);
			this.label67.TabIndex = 823;
			this.label67.Text = "Teléfono";
			this.label67.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// label68
			// 
			this.label68.AutoSize = true;
			this.label68.BackColor = System.Drawing.Color.Transparent;
			this.label68.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label68.Location = new System.Drawing.Point(504, 96);
			this.label68.Name = "label68";
			this.label68.Size = new System.Drawing.Size(87, 17);
			this.label68.TabIndex = 822;
			this.label68.Text = "Calle Secundaria";
			this.label68.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// label69
			// 
			this.label69.AutoSize = true;
			this.label69.BackColor = System.Drawing.Color.Transparent;
			this.label69.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label69.Location = new System.Drawing.Point(320, 96);
			this.label69.Name = "label69";
			this.label69.Size = new System.Drawing.Size(71, 17);
			this.label69.TabIndex = 821;
			this.label69.Text = "Número Casa";
			this.label69.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// label70
			// 
			this.label70.AutoSize = true;
			this.label70.BackColor = System.Drawing.Color.Transparent;
			this.label70.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label70.Location = new System.Drawing.Point(16, 96);
			this.label70.Name = "label70";
			this.label70.Size = new System.Drawing.Size(73, 17);
			this.label70.TabIndex = 820;
			this.label70.Text = "Calle Principal";
			this.label70.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// txtDPCTEmail
			// 
			appearance59.BackColorDisabled = System.Drawing.Color.Silver;
			appearance59.BackColorDisabled2 = System.Drawing.Color.Silver;
			appearance59.BorderColor = System.Drawing.Color.Black;
			appearance59.BorderColor3DBase = System.Drawing.Color.Black;
			appearance59.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtDPCTEmail.Appearance = appearance59;
			this.txtDPCTEmail.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtDPCTEmail.Enabled = false;
			this.txtDPCTEmail.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.txtDPCTEmail.Location = new System.Drawing.Point(656, 72);
			this.txtDPCTEmail.Name = "txtDPCTEmail";
			this.txtDPCTEmail.Size = new System.Drawing.Size(120, 22);
			this.txtDPCTEmail.TabIndex = 819;
			this.txtDPCTEmail.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtDPCTEmail_KeyDown);
			// 
			// txtDPCTCelularC
			// 
			this.txtDPCTCelularC.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtDPCTCelularC.Enabled = false;
			this.txtDPCTCelularC.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.txtDPCTCelularC.Location = new System.Drawing.Point(544, 152);
			this.txtDPCTCelularC.MaxLength = 10;
			this.txtDPCTCelularC.Name = "txtDPCTCelularC";
			this.txtDPCTCelularC.Size = new System.Drawing.Size(100, 22);
			this.txtDPCTCelularC.TabIndex = 818;
			this.txtDPCTCelularC.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtDPCTCelularC_KeyDown);
			this.txtDPCTCelularC.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.txtDPCTCelularC_KeyPress);
			// 
			// txtDPCTExt
			// 
			this.txtDPCTExt.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtDPCTExt.Enabled = false;
			this.txtDPCTExt.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.txtDPCTExt.Location = new System.Drawing.Point(456, 152);
			this.txtDPCTExt.MaxLength = 9;
			this.txtDPCTExt.Name = "txtDPCTExt";
			this.txtDPCTExt.Size = new System.Drawing.Size(80, 22);
			this.txtDPCTExt.TabIndex = 817;
			this.txtDPCTExt.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtDPCTExt_KeyDown);
			this.txtDPCTExt.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.txtDPCTExt_KeyPress);
			// 
			// txtDPCTtelefono
			// 
			this.txtDPCTtelefono.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtDPCTtelefono.Enabled = false;
			this.txtDPCTtelefono.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.txtDPCTtelefono.Location = new System.Drawing.Point(320, 152);
			this.txtDPCTtelefono.MaxLength = 10;
			this.txtDPCTtelefono.Name = "txtDPCTtelefono";
			this.txtDPCTtelefono.Size = new System.Drawing.Size(100, 22);
			this.txtDPCTtelefono.TabIndex = 816;
			this.txtDPCTtelefono.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtDPCTtelefono_KeyDown);
			this.txtDPCTtelefono.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.txtDPCTtelefono_KeyPress);
			// 
			// txtDPCTCalleSecundaria
			// 
			appearance60.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtDPCTCalleSecundaria.Appearance = appearance60;
			this.txtDPCTCalleSecundaria.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtDPCTCalleSecundaria.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtDPCTCalleSecundaria.Enabled = false;
			this.txtDPCTCalleSecundaria.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.txtDPCTCalleSecundaria.Location = new System.Drawing.Point(504, 112);
			this.txtDPCTCalleSecundaria.Name = "txtDPCTCalleSecundaria";
			this.txtDPCTCalleSecundaria.Size = new System.Drawing.Size(300, 22);
			this.txtDPCTCalleSecundaria.TabIndex = 815;
			this.txtDPCTCalleSecundaria.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtDPCTCalleSecundaria_KeyDown);
			this.txtDPCTCalleSecundaria.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.txtDPCTCalleSecundaria_KeyPress);
			// 
			// txtDPCTNumeroCasa
			// 
			appearance61.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtDPCTNumeroCasa.Appearance = appearance61;
			this.txtDPCTNumeroCasa.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtDPCTNumeroCasa.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtDPCTNumeroCasa.Enabled = false;
			this.txtDPCTNumeroCasa.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.txtDPCTNumeroCasa.Location = new System.Drawing.Point(320, 112);
			this.txtDPCTNumeroCasa.Name = "txtDPCTNumeroCasa";
			this.txtDPCTNumeroCasa.Size = new System.Drawing.Size(180, 22);
			this.txtDPCTNumeroCasa.TabIndex = 814;
			this.txtDPCTNumeroCasa.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtDPCTNumeroCasa_KeyDown);
			this.txtDPCTNumeroCasa.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.txtDPCTNumeroCasa_KeyPress);
			// 
			// txtDPCTCallePrincipal
			// 
			appearance62.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtDPCTCallePrincipal.Appearance = appearance62;
			this.txtDPCTCallePrincipal.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtDPCTCallePrincipal.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtDPCTCallePrincipal.Enabled = false;
			this.txtDPCTCallePrincipal.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.txtDPCTCallePrincipal.Location = new System.Drawing.Point(16, 112);
			this.txtDPCTCallePrincipal.Name = "txtDPCTCallePrincipal";
			this.txtDPCTCallePrincipal.Size = new System.Drawing.Size(300, 22);
			this.txtDPCTCallePrincipal.TabIndex = 813;
			this.txtDPCTCallePrincipal.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtDPCTCallePrincipal_KeyDown);
			this.txtDPCTCallePrincipal.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.txtDPCTCallePrincipal_KeyPress);
			// 
			// label71
			// 
			this.label71.AutoSize = true;
			this.label71.BackColor = System.Drawing.Color.Transparent;
			this.label71.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label71.Location = new System.Drawing.Point(496, 56);
			this.label71.Name = "label71";
			this.label71.Size = new System.Drawing.Size(34, 17);
			this.label71.TabIndex = 812;
			this.label71.Text = "Barrio";
			this.label71.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// label72
			// 
			this.label72.AutoSize = true;
			this.label72.BackColor = System.Drawing.Color.Transparent;
			this.label72.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label72.Location = new System.Drawing.Point(336, 56);
			this.label72.Name = "label72";
			this.label72.Size = new System.Drawing.Size(52, 17);
			this.label72.TabIndex = 811;
			this.label72.Text = "Parroquia";
			this.label72.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// label73
			// 
			this.label73.AutoSize = true;
			this.label73.BackColor = System.Drawing.Color.Transparent;
			this.label73.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label73.Location = new System.Drawing.Point(176, 56);
			this.label73.Name = "label73";
			this.label73.Size = new System.Drawing.Size(39, 17);
			this.label73.TabIndex = 810;
			this.label73.Text = "Cantón";
			this.label73.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// label74
			// 
			this.label74.AutoSize = true;
			this.label74.BackColor = System.Drawing.Color.Transparent;
			this.label74.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label74.Location = new System.Drawing.Point(16, 56);
			this.label74.Name = "label74";
			this.label74.Size = new System.Drawing.Size(49, 17);
			this.label74.TabIndex = 809;
			this.label74.Text = "Provincia";
			this.label74.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// cmbDPCTBarrio
			// 
			this.cmbDPCTBarrio.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbDPCTBarrio.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbDPCTBarrio.DataSource = this.ultraDataSource4;
			ultraGridColumn59.Header.VisiblePosition = 0;
			ultraGridColumn59.Hidden = true;
			ultraGridColumn60.Header.Caption = "Barrio";
			ultraGridColumn60.Header.VisiblePosition = 1;
			ultraGridColumn60.Width = 150;
			ultraGridBand29.Columns.AddRange(new object[] {
																											ultraGridColumn59,
																											ultraGridColumn60});
			this.cmbDPCTBarrio.DisplayLayout.BandsSerializer.Add(ultraGridBand29);
			this.cmbDPCTBarrio.DisplayLayout.Override.RowSelectors = Infragistics.Win.DefaultableBoolean.True;
			this.cmbDPCTBarrio.DisplayMember = "Nombre";
			this.cmbDPCTBarrio.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbDPCTBarrio.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbDPCTBarrio.Enabled = false;
			this.cmbDPCTBarrio.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbDPCTBarrio.Location = new System.Drawing.Point(496, 72);
			this.cmbDPCTBarrio.Name = "cmbDPCTBarrio";
			this.cmbDPCTBarrio.Size = new System.Drawing.Size(150, 22);
			this.cmbDPCTBarrio.TabIndex = 780;
			this.cmbDPCTBarrio.ValueMember = "idBarrio";
			this.cmbDPCTBarrio.KeyDown += new System.Windows.Forms.KeyEventHandler(this.cmbDPCTBarrio_KeyDown);
			// 
			// cmbDPCTParroquia
			// 
			this.cmbDPCTParroquia.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbDPCTParroquia.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbDPCTParroquia.DataSource = this.ultraDataSource3;
			ultraGridColumn61.Header.VisiblePosition = 0;
			ultraGridColumn61.Hidden = true;
			ultraGridColumn62.Header.Caption = "Provincia";
			ultraGridColumn62.Header.VisiblePosition = 1;
			ultraGridColumn62.Width = 150;
			ultraGridBand30.Columns.AddRange(new object[] {
																											ultraGridColumn61,
																											ultraGridColumn62});
			this.cmbDPCTParroquia.DisplayLayout.BandsSerializer.Add(ultraGridBand30);
			this.cmbDPCTParroquia.DisplayMember = "Nombre";
			this.cmbDPCTParroquia.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbDPCTParroquia.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbDPCTParroquia.Enabled = false;
			this.cmbDPCTParroquia.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbDPCTParroquia.Location = new System.Drawing.Point(336, 72);
			this.cmbDPCTParroquia.Name = "cmbDPCTParroquia";
			this.cmbDPCTParroquia.Size = new System.Drawing.Size(150, 22);
			this.cmbDPCTParroquia.TabIndex = 778;
			this.cmbDPCTParroquia.ValueMember = "idParroquia";
			this.cmbDPCTParroquia.KeyDown += new System.Windows.Forms.KeyEventHandler(this.cmbDPCTParroquia_KeyDown);
			this.cmbDPCTParroquia.ValueChanged += new System.EventHandler(this.cmbDPCTParroquia_ValueChanged);
			// 
			// cmbDPCTCanton
			// 
			this.cmbDPCTCanton.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbDPCTCanton.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbDPCTCanton.DataSource = this.ultraDataSource2;
			ultraGridColumn63.Header.VisiblePosition = 0;
			ultraGridColumn63.Hidden = true;
			ultraGridColumn64.Header.Caption = "Canton";
			ultraGridColumn64.Header.VisiblePosition = 1;
			ultraGridColumn64.Width = 150;
			ultraGridBand31.Columns.AddRange(new object[] {
																											ultraGridColumn63,
																											ultraGridColumn64});
			this.cmbDPCTCanton.DisplayLayout.BandsSerializer.Add(ultraGridBand31);
			this.cmbDPCTCanton.DisplayLayout.Override.RowSelectors = Infragistics.Win.DefaultableBoolean.True;
			this.cmbDPCTCanton.DisplayMember = "Nombre";
			this.cmbDPCTCanton.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbDPCTCanton.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbDPCTCanton.Enabled = false;
			this.cmbDPCTCanton.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbDPCTCanton.Location = new System.Drawing.Point(176, 72);
			this.cmbDPCTCanton.Name = "cmbDPCTCanton";
			this.cmbDPCTCanton.Size = new System.Drawing.Size(150, 22);
			this.cmbDPCTCanton.TabIndex = 776;
			this.cmbDPCTCanton.ValueMember = "idCanton";
			this.cmbDPCTCanton.KeyDown += new System.Windows.Forms.KeyEventHandler(this.cmbDPCTCanton_KeyDown);
			this.cmbDPCTCanton.ValueChanged += new System.EventHandler(this.cmbDPCTCanton_ValueChanged);
			// 
			// cmbDPCTProvincia
			// 
			this.cmbDPCTProvincia.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbDPCTProvincia.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbDPCTProvincia.DataSource = this.ultraDataSource1;
			ultraGridColumn65.Header.VisiblePosition = 0;
			ultraGridColumn65.Hidden = true;
			ultraGridColumn65.Width = 50;
			ultraGridColumn66.Header.VisiblePosition = 1;
			ultraGridColumn66.Width = 150;
			ultraGridBand32.Columns.AddRange(new object[] {
																											ultraGridColumn65,
																											ultraGridColumn66});
			this.cmbDPCTProvincia.DisplayLayout.BandsSerializer.Add(ultraGridBand32);
			this.cmbDPCTProvincia.DisplayMember = "Nombre";
			this.cmbDPCTProvincia.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbDPCTProvincia.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbDPCTProvincia.Enabled = false;
			this.cmbDPCTProvincia.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbDPCTProvincia.Location = new System.Drawing.Point(16, 72);
			this.cmbDPCTProvincia.Name = "cmbDPCTProvincia";
			this.cmbDPCTProvincia.Size = new System.Drawing.Size(150, 22);
			this.cmbDPCTProvincia.TabIndex = 774;
			this.cmbDPCTProvincia.ValueMember = "idProvincia";
			this.cmbDPCTProvincia.KeyDown += new System.Windows.Forms.KeyEventHandler(this.cmbDPCTProvincia_KeyDown);
			this.cmbDPCTProvincia.ValueChanged += new System.EventHandler(this.cmbDPCTProvincia_ValueChanged);
			// 
			// ultraTabPageControl5
			// 
			this.ultraTabPageControl5.Controls.Add(this.cmbtipoCuenta);
			this.ultraTabPageControl5.Controls.Add(this.cmbTipoBanco);
			this.ultraTabPageControl5.Controls.Add(this.cmbtipoReferencia);
			this.ultraTabPageControl5.Controls.Add(this.ultgconsulta);
			this.ultraTabPageControl5.Controls.Add(this.ultraDropDown1);
			this.ultraTabPageControl5.Controls.Add(this.cmbProvincia);
			this.ultraTabPageControl5.Controls.Add(this.cmbParentesco);
			this.ultraTabPageControl5.Controls.Add(this.uGridRefernciaBancaria);
			this.ultraTabPageControl5.Location = new System.Drawing.Point(-10000, -10000);
			this.ultraTabPageControl5.Name = "ultraTabPageControl5";
			this.ultraTabPageControl5.Size = new System.Drawing.Size(1116, 366);
			this.ultraTabPageControl5.Paint += new System.Windows.Forms.PaintEventHandler(this.ultraTabPageControl5_Paint);
			// 
			// cmbtipoCuenta
			// 
			this.cmbtipoCuenta.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbtipoCuenta.DataSource = this.ultraDataSource16;
			this.cmbtipoCuenta.DisplayLayout.AutoFitColumns = true;
			ultraGridColumn67.Header.VisiblePosition = 0;
			ultraGridColumn67.Width = 106;
			ultraGridColumn68.Header.VisiblePosition = 1;
			ultraGridColumn68.Width = 116;
			ultraGridBand33.Columns.AddRange(new object[] {
																											ultraGridColumn67,
																											ultraGridColumn68});
			ultraGridBand33.SummaryFooterCaption = "Grand Summaries";
			this.cmbtipoCuenta.DisplayLayout.BandsSerializer.Add(ultraGridBand33);
			this.cmbtipoCuenta.DisplayMember = "Nombre";
			this.cmbtipoCuenta.Font = new System.Drawing.Font("Tahoma", 8.25F);
			this.cmbtipoCuenta.Location = new System.Drawing.Point(872, 440);
			this.cmbtipoCuenta.Name = "cmbtipoCuenta";
			this.cmbtipoCuenta.Size = new System.Drawing.Size(224, 104);
			this.cmbtipoCuenta.TabIndex = 803;
			this.cmbtipoCuenta.ValueMember = "idTipoCuentaBan";
			this.cmbtipoCuenta.Visible = false;
			// 
			// ultraDataSource16
			// 
			ultraDataColumn41.DataType = typeof(int);
			this.ultraDataSource16.Band.Columns.AddRange(new object[] {
																																	ultraDataColumn41,
																																	ultraDataColumn42});
			// 
			// cmbTipoBanco
			// 
			this.cmbTipoBanco.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbTipoBanco.DataSource = this.ultraDataSource15;
			this.cmbTipoBanco.DisplayLayout.AutoFitColumns = true;
			ultraGridColumn69.Header.VisiblePosition = 0;
			ultraGridColumn69.Width = 72;
			ultraGridColumn70.Header.VisiblePosition = 1;
			ultraGridColumn70.Width = 133;
			ultraGridBand34.Columns.AddRange(new object[] {
																											ultraGridColumn69,
																											ultraGridColumn70});
			ultraGridBand34.SummaryFooterCaption = "Grand Summaries";
			this.cmbTipoBanco.DisplayLayout.BandsSerializer.Add(ultraGridBand34);
			this.cmbTipoBanco.DisplayMember = "Nombre";
			this.cmbTipoBanco.Font = new System.Drawing.Font("Tahoma", 8.25F);
			this.cmbTipoBanco.Location = new System.Drawing.Point(8, 376);
			this.cmbTipoBanco.Name = "cmbTipoBanco";
			this.cmbTipoBanco.Size = new System.Drawing.Size(224, 8);
			this.cmbTipoBanco.TabIndex = 802;
			this.cmbTipoBanco.ValueMember = "idBancos";
			this.cmbTipoBanco.Visible = false;
			// 
			// ultraDataSource15
			// 
			ultraDataColumn43.DataType = typeof(int);
			this.ultraDataSource15.Band.Columns.AddRange(new object[] {
																																	ultraDataColumn43,
																																	ultraDataColumn44});
			// 
			// cmbtipoReferencia
			// 
			this.cmbtipoReferencia.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbtipoReferencia.DataSource = this.ultraDataSource14;
			this.cmbtipoReferencia.DisplayLayout.AutoFitColumns = true;
			ultraGridColumn71.Header.VisiblePosition = 0;
			ultraGridColumn71.Width = 102;
			ultraGridColumn72.Header.VisiblePosition = 1;
			ultraGridColumn72.Width = 103;
			ultraGridBand35.Columns.AddRange(new object[] {
																											ultraGridColumn71,
																											ultraGridColumn72});
			ultraGridBand35.SummaryFooterCaption = "Grand Summaries";
			this.cmbtipoReferencia.DisplayLayout.BandsSerializer.Add(ultraGridBand35);
			this.cmbtipoReferencia.DisplayMember = "Nombre";
			this.cmbtipoReferencia.Font = new System.Drawing.Font("Tahoma", 8.25F);
			this.cmbtipoReferencia.Location = new System.Drawing.Point(8, 392);
			this.cmbtipoReferencia.Name = "cmbtipoReferencia";
			this.cmbtipoReferencia.Size = new System.Drawing.Size(224, 8);
			this.cmbtipoReferencia.TabIndex = 801;
			this.cmbtipoReferencia.ValueMember = "idTipoRefBancaria";
			this.cmbtipoReferencia.Visible = false;
			// 
			// ultraDataSource14
			// 
			ultraDataColumn45.DataType = typeof(int);
			this.ultraDataSource14.Band.Columns.AddRange(new object[] {
																																	ultraDataColumn45,
																																	ultraDataColumn46});
			// 
			// ultgconsulta
			// 
			this.ultgconsulta.Cursor = System.Windows.Forms.Cursors.Default;
			this.ultgconsulta.DataSource = this.ultraDataSource13;
			appearance63.BackColor = System.Drawing.Color.White;
			appearance63.ForegroundAlpha = Infragistics.Win.Alpha.UseAlphaLevel;
			this.ultgconsulta.DisplayLayout.Appearance = appearance63;
			ultraGridColumn73.Case = Infragistics.Win.UltraWinGrid.Case.Upper;
			appearance64.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn73.CellAppearance = appearance64;
			ultraGridColumn73.Header.VisiblePosition = 0;
			ultraGridColumn73.Hidden = true;
			ultraGridColumn74.Case = Infragistics.Win.UltraWinGrid.Case.Upper;
			ultraGridColumn74.Header.VisiblePosition = 1;
			ultraGridColumn74.Hidden = true;
			ultraGridColumn75.Case = Infragistics.Win.UltraWinGrid.Case.Upper;
			ultraGridColumn75.Header.VisiblePosition = 2;
			ultraGridColumn75.Hidden = true;
			ultraGridColumn76.Case = Infragistics.Win.UltraWinGrid.Case.Upper;
			ultraGridColumn76.Header.Caption = "Parentesco";
			ultraGridColumn76.Header.VisiblePosition = 3;
			ultraGridColumn76.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.DropDownList;
			ultraGridColumn76.Width = 150;
			ultraGridColumn77.Case = Infragistics.Win.UltraWinGrid.Case.Upper;
			ultraGridColumn77.Header.Caption = "Apellido Paterno";
			ultraGridColumn77.Header.VisiblePosition = 4;
			ultraGridColumn77.Width = 150;
			ultraGridColumn78.Case = Infragistics.Win.UltraWinGrid.Case.Upper;
			ultraGridColumn78.Header.Caption = "Apellido Materno";
			ultraGridColumn78.Header.VisiblePosition = 5;
			ultraGridColumn78.Hidden = true;
			ultraGridColumn78.Width = 150;
			ultraGridColumn79.Case = Infragistics.Win.UltraWinGrid.Case.Upper;
			ultraGridColumn79.Header.Caption = "Primer Nombre";
			ultraGridColumn79.Header.VisiblePosition = 6;
			ultraGridColumn79.Width = 150;
			ultraGridColumn80.Case = Infragistics.Win.UltraWinGrid.Case.Upper;
			ultraGridColumn80.Header.Caption = "Segundo Nombre";
			ultraGridColumn80.Header.VisiblePosition = 7;
			ultraGridColumn80.Width = 150;
			ultraGridColumn81.Case = Infragistics.Win.UltraWinGrid.Case.Upper;
			ultraGridColumn81.Header.Caption = "Provincia";
			ultraGridColumn81.Header.VisiblePosition = 8;
			ultraGridColumn81.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.DropDownList;
			ultraGridColumn81.Width = 150;
			ultraGridColumn82.Case = Infragistics.Win.UltraWinGrid.Case.Upper;
			ultraGridColumn82.Header.Caption = "Canton";
			ultraGridColumn82.Header.VisiblePosition = 9;
			ultraGridColumn82.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.DropDownList;
			ultraGridColumn82.Width = 150;
			ultraGridColumn83.Case = Infragistics.Win.UltraWinGrid.Case.Upper;
			ultraGridColumn83.Header.VisiblePosition = 10;
			ultraGridColumn83.Hidden = true;
			ultraGridColumn83.Width = 150;
			ultraGridColumn84.Case = Infragistics.Win.UltraWinGrid.Case.Upper;
			appearance65.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn84.CellAppearance = appearance65;
			ultraGridColumn84.Header.VisiblePosition = 11;
			ultraGridColumn84.Hidden = true;
			ultraGridColumn84.Width = 100;
			ultraGridColumn85.Case = Infragistics.Win.UltraWinGrid.Case.Upper;
			appearance66.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn85.CellAppearance = appearance66;
			ultraGridColumn85.Header.VisiblePosition = 13;
			ultraGridColumn85.Width = 100;
			ultraGridColumn86.Case = Infragistics.Win.UltraWinGrid.Case.Upper;
			ultraGridColumn86.Header.Caption = "Empresa Trabajo";
			ultraGridColumn86.Header.VisiblePosition = 15;
			ultraGridColumn86.Hidden = true;
			ultraGridColumn86.Width = 150;
			ultraGridColumn87.Case = Infragistics.Win.UltraWinGrid.Case.Upper;
			ultraGridColumn87.Header.Caption = "Direccion Trabajo";
			ultraGridColumn87.Header.VisiblePosition = 16;
			ultraGridColumn87.Hidden = true;
			ultraGridColumn87.Width = 150;
			ultraGridColumn88.Case = Infragistics.Win.UltraWinGrid.Case.Upper;
			appearance67.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn88.CellAppearance = appearance67;
			ultraGridColumn88.Header.Caption = "Telf Trabajo";
			ultraGridColumn88.Header.VisiblePosition = 17;
			ultraGridColumn88.Hidden = true;
			ultraGridColumn88.Width = 100;
			ultraGridColumn89.Case = Infragistics.Win.UltraWinGrid.Case.Upper;
			appearance68.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn89.CellAppearance = appearance68;
			ultraGridColumn89.Header.Caption = "Ext Trabajo";
			ultraGridColumn89.Header.VisiblePosition = 19;
			ultraGridColumn89.Hidden = true;
			ultraGridColumn89.Width = 90;
			ultraGridColumn90.ButtonDisplayStyle = Infragistics.Win.UltraWinGrid.ButtonDisplayStyle.Always;
			ultraGridColumn90.Case = Infragistics.Win.UltraWinGrid.Case.Upper;
			appearance69.Image = ((object)(resources.GetObject("appearance69.Image")));
			ultraGridColumn90.CellButtonAppearance = appearance69;
			ultraGridColumn90.Header.Caption = "....";
			ultraGridColumn90.Header.VisiblePosition = 12;
			ultraGridColumn90.Hidden = true;
			ultraGridColumn90.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.Button;
			ultraGridColumn90.Width = 30;
			ultraGridColumn91.ButtonDisplayStyle = Infragistics.Win.UltraWinGrid.ButtonDisplayStyle.Always;
			ultraGridColumn91.Case = Infragistics.Win.UltraWinGrid.Case.Upper;
			appearance70.Image = ((object)(resources.GetObject("appearance70.Image")));
			ultraGridColumn91.CellButtonAppearance = appearance70;
			ultraGridColumn91.Header.Caption = "....";
			ultraGridColumn91.Header.VisiblePosition = 14;
			ultraGridColumn91.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.Button;
			ultraGridColumn91.Width = 30;
			ultraGridColumn92.ButtonDisplayStyle = Infragistics.Win.UltraWinGrid.ButtonDisplayStyle.Always;
			ultraGridColumn92.Case = Infragistics.Win.UltraWinGrid.Case.Upper;
			appearance71.Image = ((object)(resources.GetObject("appearance71.Image")));
			ultraGridColumn92.CellButtonAppearance = appearance71;
			ultraGridColumn92.Header.Caption = "....";
			ultraGridColumn92.Header.VisiblePosition = 18;
			ultraGridColumn92.Hidden = true;
			ultraGridColumn92.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.Button;
			ultraGridColumn92.Width = 30;
			ultraGridBand36.Columns.AddRange(new object[] {
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
																											ultraGridColumn92});
			this.ultgconsulta.DisplayLayout.BandsSerializer.Add(ultraGridBand36);
			this.ultgconsulta.DisplayLayout.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.TemplateOnBottom;
			this.ultgconsulta.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.True;
			this.ultgconsulta.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.True;
			appearance72.BackColor = System.Drawing.Color.Transparent;
			this.ultgconsulta.DisplayLayout.Override.CardAreaAppearance = appearance72;
			appearance73.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance73.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance73.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance73.FontData.BoldAsString = "True";
			appearance73.FontData.Name = "Arial";
			appearance73.FontData.SizeInPoints = 8.5F;
			appearance73.ForeColor = System.Drawing.Color.White;
			appearance73.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.ultgconsulta.DisplayLayout.Override.HeaderAppearance = appearance73;
			appearance74.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance74.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance74.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ultgconsulta.DisplayLayout.Override.RowSelectorAppearance = appearance74;
			appearance75.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance75.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance75.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ultgconsulta.DisplayLayout.Override.SelectedRowAppearance = appearance75;
			this.ultgconsulta.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.ultgconsulta.Location = new System.Drawing.Point(8, 8);
			this.ultgconsulta.Name = "ultgconsulta";
			this.ultgconsulta.Size = new System.Drawing.Size(1104, 224);
			this.ultgconsulta.TabIndex = 800;
			this.ultgconsulta.CellChange += new Infragistics.Win.UltraWinGrid.CellEventHandler(this.ultgconsulta_CellChange);
			this.ultgconsulta.AfterRowInsert += new Infragistics.Win.UltraWinGrid.RowEventHandler(this.ultgconsulta_AfterRowInsert);
			this.ultgconsulta.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.ultgconsulta_KeyPress);
			this.ultgconsulta.KeyDown += new System.Windows.Forms.KeyEventHandler(this.ultgconsulta_KeyDown);
			this.ultgconsulta.ClickCellButton += new Infragistics.Win.UltraWinGrid.CellEventHandler(this.ultgconsulta_ClickCellButton);
			// 
			// ultraDataSource13
			// 
			ultraDataColumn47.DataType = typeof(int);
			ultraDataColumn48.DataType = typeof(int);
			ultraDataColumn49.DataType = typeof(int);
			ultraDataColumn50.DataType = typeof(int);
			ultraDataColumn55.DataType = typeof(int);
			ultraDataColumn56.DataType = typeof(int);
			this.ultraDataSource13.Band.Columns.AddRange(new object[] {
																																	ultraDataColumn47,
																																	ultraDataColumn48,
																																	ultraDataColumn49,
																																	ultraDataColumn50,
																																	ultraDataColumn51,
																																	ultraDataColumn52,
																																	ultraDataColumn53,
																																	ultraDataColumn54,
																																	ultraDataColumn55,
																																	ultraDataColumn56,
																																	ultraDataColumn57,
																																	ultraDataColumn58,
																																	ultraDataColumn59,
																																	ultraDataColumn60,
																																	ultraDataColumn61,
																																	ultraDataColumn62,
																																	ultraDataColumn63});
			// 
			// ultraDropDown1
			// 
			this.ultraDropDown1.Cursor = System.Windows.Forms.Cursors.Default;
			this.ultraDropDown1.DataSource = this.ultraDataSource2;
			this.ultraDropDown1.DisplayLayout.AutoFitColumns = true;
			ultraGridColumn93.Header.VisiblePosition = 0;
			ultraGridColumn93.Hidden = true;
			ultraGridColumn93.Width = 64;
			ultraGridColumn94.Header.VisiblePosition = 1;
			ultraGridColumn94.Width = 222;
			ultraGridBand37.Columns.AddRange(new object[] {
																											ultraGridColumn93,
																											ultraGridColumn94});
			ultraGridBand37.SummaryFooterCaption = "Grand Summaries";
			this.ultraDropDown1.DisplayLayout.BandsSerializer.Add(ultraGridBand37);
			this.ultraDropDown1.DisplayMember = "Nombre";
			this.ultraDropDown1.Font = new System.Drawing.Font("Tahoma", 8.25F);
			this.ultraDropDown1.Location = new System.Drawing.Point(136, 240);
			this.ultraDropDown1.Name = "ultraDropDown1";
			this.ultraDropDown1.Size = new System.Drawing.Size(224, 104);
			this.ultraDropDown1.TabIndex = 795;
			this.ultraDropDown1.ValueMember = "idCanton";
			this.ultraDropDown1.Visible = false;
			// 
			// cmbProvincia
			// 
			this.cmbProvincia.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbProvincia.DataSource = this.ultraDataSource1;
			this.cmbProvincia.DisplayLayout.AutoFitColumns = true;
			ultraGridColumn95.Header.VisiblePosition = 0;
			ultraGridColumn95.Hidden = true;
			ultraGridColumn95.Width = 69;
			ultraGridColumn96.Header.VisiblePosition = 1;
			ultraGridColumn96.Width = 205;
			ultraGridBand38.Columns.AddRange(new object[] {
																											ultraGridColumn95,
																											ultraGridColumn96});
			ultraGridBand38.SummaryFooterCaption = "Grand Summaries";
			this.cmbProvincia.DisplayLayout.BandsSerializer.Add(ultraGridBand38);
			this.cmbProvincia.DisplayMember = "Nombre";
			this.cmbProvincia.Font = new System.Drawing.Font("Tahoma", 8.25F);
			this.cmbProvincia.Location = new System.Drawing.Point(432, 288);
			this.cmbProvincia.Name = "cmbProvincia";
			this.cmbProvincia.Size = new System.Drawing.Size(224, 72);
			this.cmbProvincia.TabIndex = 794;
			this.cmbProvincia.ValueMember = "idProvincia";
			this.cmbProvincia.Visible = false;
			this.cmbProvincia.Validating += new System.ComponentModel.CancelEventHandler(this.cmbProvincia_Validating);
			this.cmbProvincia.RowSelected += new Infragistics.Win.UltraWinGrid.RowSelectedEventHandler(this.cmbProvincia_RowSelected);
			// 
			// cmbParentesco
			// 
			this.cmbParentesco.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbParentesco.DataSource = this.ultraDataSource9;
			this.cmbParentesco.DisplayLayout.AutoFitColumns = true;
			ultraGridColumn97.Header.VisiblePosition = 0;
			ultraGridColumn97.Hidden = true;
			ultraGridColumn97.Width = 117;
			ultraGridColumn98.Header.VisiblePosition = 1;
			ultraGridColumn98.Width = 222;
			ultraGridBand39.Columns.AddRange(new object[] {
																											ultraGridColumn97,
																											ultraGridColumn98});
			ultraGridBand39.SummaryFooterCaption = "Grand Summaries";
			this.cmbParentesco.DisplayLayout.BandsSerializer.Add(ultraGridBand39);
			this.cmbParentesco.DisplayMember = "Nombre";
			this.cmbParentesco.Font = new System.Drawing.Font("Tahoma", 8.25F);
			this.cmbParentesco.Location = new System.Drawing.Point(672, 248);
			this.cmbParentesco.Name = "cmbParentesco";
			this.cmbParentesco.Size = new System.Drawing.Size(224, 104);
			this.cmbParentesco.TabIndex = 793;
			this.cmbParentesco.ValueMember = "idParentesco";
			this.cmbParentesco.Visible = false;
			this.cmbParentesco.Validating += new System.ComponentModel.CancelEventHandler(this.cmbParentesco_Validating);
			this.cmbParentesco.Validated += new System.EventHandler(this.cmbParentesco_Validated);
			this.cmbParentesco.InitializeLayout += new Infragistics.Win.UltraWinGrid.InitializeLayoutEventHandler(this.cmbParentesco_InitializeLayout);
			// 
			// ultraDataSource9
			// 
			ultraDataColumn64.DataType = typeof(int);
			this.ultraDataSource9.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn64,
																																 ultraDataColumn65});
			// 
			// uGridRefernciaBancaria
			// 
			this.uGridRefernciaBancaria.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.uGridRefernciaBancaria.Cursor = System.Windows.Forms.Cursors.Default;
			this.uGridRefernciaBancaria.DataSource = this.ultraDataSource17;
			appearance76.BackColor = System.Drawing.Color.White;
			this.uGridRefernciaBancaria.DisplayLayout.Appearance = appearance76;
			ultraGridColumn99.Header.VisiblePosition = 0;
			ultraGridColumn100.Header.VisiblePosition = 1;
			ultraGridColumn101.Header.VisiblePosition = 2;
			ultraGridColumn101.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.DropDownList;
			ultraGridColumn102.Header.VisiblePosition = 3;
			ultraGridColumn102.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.DropDownList;
			ultraGridColumn103.Header.VisiblePosition = 4;
			ultraGridColumn103.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.DropDownList;
			ultraGridColumn104.Header.VisiblePosition = 5;
			ultraGridColumn105.Header.VisiblePosition = 6;
			ultraGridBand40.Columns.AddRange(new object[] {
																											ultraGridColumn99,
																											ultraGridColumn100,
																											ultraGridColumn101,
																											ultraGridColumn102,
																											ultraGridColumn103,
																											ultraGridColumn104,
																											ultraGridColumn105});
			this.uGridRefernciaBancaria.DisplayLayout.BandsSerializer.Add(ultraGridBand40);
			this.uGridRefernciaBancaria.DisplayLayout.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.TemplateOnBottom;
			this.uGridRefernciaBancaria.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.True;
			appearance77.BackColor = System.Drawing.Color.Transparent;
			this.uGridRefernciaBancaria.DisplayLayout.Override.CardAreaAppearance = appearance77;
			appearance78.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance78.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance78.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance78.FontData.BoldAsString = "True";
			appearance78.FontData.Name = "Arial";
			appearance78.FontData.SizeInPoints = 10F;
			appearance78.ForeColor = System.Drawing.Color.White;
			appearance78.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.uGridRefernciaBancaria.DisplayLayout.Override.HeaderAppearance = appearance78;
			this.uGridRefernciaBancaria.DisplayLayout.Override.HeaderClickAction = Infragistics.Win.UltraWinGrid.HeaderClickAction.SortMulti;
			appearance79.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance79.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance79.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridRefernciaBancaria.DisplayLayout.Override.RowSelectorAppearance = appearance79;
			appearance80.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance80.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance80.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.uGridRefernciaBancaria.DisplayLayout.Override.SelectedRowAppearance = appearance80;
			this.uGridRefernciaBancaria.DisplayLayout.Override.SelectTypeCell = Infragistics.Win.UltraWinGrid.SelectType.None;
			this.uGridRefernciaBancaria.DisplayLayout.Override.SelectTypeCol = Infragistics.Win.UltraWinGrid.SelectType.None;
			this.uGridRefernciaBancaria.DisplayLayout.Override.SelectTypeRow = Infragistics.Win.UltraWinGrid.SelectType.Single;
			this.uGridRefernciaBancaria.DisplayLayout.ScrollBounds = Infragistics.Win.UltraWinGrid.ScrollBounds.ScrollToFill;
			this.uGridRefernciaBancaria.DisplayLayout.ScrollStyle = Infragistics.Win.UltraWinGrid.ScrollStyle.Immediate;
			this.uGridRefernciaBancaria.DisplayLayout.TabNavigation = Infragistics.Win.UltraWinGrid.TabNavigation.NextControl;
			this.uGridRefernciaBancaria.DisplayLayout.ViewStyle = Infragistics.Win.UltraWinGrid.ViewStyle.SingleBand;
			this.uGridRefernciaBancaria.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.uGridRefernciaBancaria.Location = new System.Drawing.Point(8, 328);
			this.uGridRefernciaBancaria.Name = "uGridRefernciaBancaria";
			this.uGridRefernciaBancaria.RightToLeft = System.Windows.Forms.RightToLeft.No;
			this.uGridRefernciaBancaria.Size = new System.Drawing.Size(215, 1);
			this.uGridRefernciaBancaria.TabIndex = 787;
			this.uGridRefernciaBancaria.Visible = false;
			this.uGridRefernciaBancaria.AfterRowInsert += new Infragistics.Win.UltraWinGrid.RowEventHandler(this.uGridRefernciaBancaria_AfterRowInsert);
			// 
			// ultraDataSource17
			// 
			ultraDataColumn66.DataType = typeof(int);
			ultraDataColumn67.DataType = typeof(int);
			ultraDataColumn68.DataType = typeof(int);
			ultraDataColumn69.DataType = typeof(int);
			ultraDataColumn70.DataType = typeof(int);
			ultraDataColumn72.DataType = typeof(int);
			this.ultraDataSource17.Band.Columns.AddRange(new object[] {
																																	ultraDataColumn66,
																																	ultraDataColumn67,
																																	ultraDataColumn68,
																																	ultraDataColumn69,
																																	ultraDataColumn70,
																																	ultraDataColumn71,
																																	ultraDataColumn72});
			// 
			// ultraTabPageControl7
			// 
			this.ultraTabPageControl7.Controls.Add(this.groupBox5);
			this.ultraTabPageControl7.Location = new System.Drawing.Point(-10000, -10000);
			this.ultraTabPageControl7.Name = "ultraTabPageControl7";
			this.ultraTabPageControl7.Size = new System.Drawing.Size(1116, 366);
			// 
			// groupBox5
			// 
			this.groupBox5.BackColor = System.Drawing.SystemColors.Control;
			this.groupBox5.Controls.Add(this.cmbTiempo);
			this.groupBox5.Controls.Add(this.label145);
			this.groupBox5.Controls.Add(this.txtNotas);
			this.groupBox5.Controls.Add(this.chkAfiliadoTributario);
			this.groupBox5.Controls.Add(this.chkNegocioObligadoaLlevar);
			this.groupBox5.Controls.Add(this.btnNegocioTelefono2);
			this.groupBox5.Controls.Add(this.txtNegocioTelefono2);
			this.groupBox5.Controls.Add(this.label96);
			this.groupBox5.Controls.Add(this.label95);
			this.groupBox5.Controls.Add(this.label94);
			this.groupBox5.Controls.Add(this.label93);
			this.groupBox5.Controls.Add(this.label87);
			this.groupBox5.Controls.Add(this.txtNegocioemail);
			this.groupBox5.Controls.Add(this.label88);
			this.groupBox5.Controls.Add(this.label90);
			this.groupBox5.Controls.Add(this.label91);
			this.groupBox5.Controls.Add(this.label92);
			this.groupBox5.Controls.Add(this.cmbNegocioBarrio);
			this.groupBox5.Controls.Add(this.cmbNegocioParroquia);
			this.groupBox5.Controls.Add(this.cmbNegocioCanton);
			this.groupBox5.Controls.Add(this.cmbNegocioProvincia);
			this.groupBox5.Controls.Add(this.label75);
			this.groupBox5.Controls.Add(this.label76);
			this.groupBox5.Controls.Add(this.label77);
			this.groupBox5.Controls.Add(this.label78);
			this.groupBox5.Controls.Add(this.txtNEgocioPasivos);
			this.groupBox5.Controls.Add(this.txtNegocioActivos);
			this.groupBox5.Controls.Add(this.txtNEgocioGastos);
			this.groupBox5.Controls.Add(this.txtNEgocioIngresos);
			this.groupBox5.Controls.Add(this.label79);
			this.groupBox5.Controls.Add(this.txtNegocioReferncia);
			this.groupBox5.Controls.Add(this.btnNegocioCelular);
			this.groupBox5.Controls.Add(this.btnNegocioTelefono1);
			this.groupBox5.Controls.Add(this.label80);
			this.groupBox5.Controls.Add(this.label82);
			this.groupBox5.Controls.Add(this.label83);
			this.groupBox5.Controls.Add(this.label84);
			this.groupBox5.Controls.Add(this.label85);
			this.groupBox5.Controls.Add(this.label86);
			this.groupBox5.Controls.Add(this.txtNEgocioCelular);
			this.groupBox5.Controls.Add(this.txtNegocioTelefono1);
			this.groupBox5.Controls.Add(this.txtNegocioCalleSecundaria);
			this.groupBox5.Controls.Add(this.txtNegocioNumeroCasa);
			this.groupBox5.Controls.Add(this.txtNegocioCallePrincipal);
			this.groupBox5.Controls.Add(this.txtNegocioEmpleados);
			this.groupBox5.Controls.Add(this.txtNegocioMetros);
			this.groupBox5.Controls.Add(this.Nombre);
			this.groupBox5.Controls.Add(this.txtNombreDelNegocio);
			this.groupBox5.Controls.Add(this.label81);
			this.groupBox5.Controls.Add(this.txtNegocioRuc);
			this.groupBox5.Controls.Add(this.cmbNegciotipo);
			this.groupBox5.Controls.Add(this.label89);
			this.groupBox5.Controls.Add(this.cmbNEgocioActivdadEconomica);
			this.groupBox5.Controls.Add(this.dtpNegocioFechaInicio);
			this.groupBox5.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.groupBox5.Location = new System.Drawing.Point(0, 0);
			this.groupBox5.Name = "groupBox5";
			this.groupBox5.Size = new System.Drawing.Size(1112, 432);
			this.groupBox5.TabIndex = 830;
			this.groupBox5.TabStop = false;
			this.groupBox5.Enter += new System.EventHandler(this.groupBox5_Enter);
			// 
			// cmbTiempo
			// 
			this.cmbTiempo.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbTiempo.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbTiempo.DataSource = this.ultraDataSource49;
			ultraGridColumn106.Header.VisiblePosition = 0;
			ultraGridColumn106.Hidden = true;
			ultraGridColumn106.Width = 8;
			ultraGridColumn107.Header.VisiblePosition = 1;
			ultraGridColumn107.Width = 150;
			ultraGridBand41.Columns.AddRange(new object[] {
																											ultraGridColumn106,
																											ultraGridColumn107});
			this.cmbTiempo.DisplayLayout.BandsSerializer.Add(ultraGridBand41);
			this.cmbTiempo.DisplayLayout.Override.RowSelectors = Infragistics.Win.DefaultableBoolean.True;
			this.cmbTiempo.DisplayMember = "Descripcion";
			this.cmbTiempo.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbTiempo.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbTiempo.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbTiempo.Location = new System.Drawing.Point(200, 35);
			this.cmbTiempo.Name = "cmbTiempo";
			this.cmbTiempo.Size = new System.Drawing.Size(120, 22);
			this.cmbTiempo.TabIndex = 1045;
			this.cmbTiempo.ValueMember = "idCre_Tiempo";
			// 
			// label145
			// 
			this.label145.AutoSize = true;
			this.label145.BackColor = System.Drawing.SystemColors.Control;
			this.label145.Location = new System.Drawing.Point(344, 176);
			this.label145.Name = "label145";
			this.label145.Size = new System.Drawing.Size(111, 17);
			this.label145.TabIndex = 889;
			this.label145.Text = "Actividad del Negocio";
			this.label145.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// txtNotas
			// 
			this.txtNotas.AcceptsReturn = true;
			appearance81.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtNotas.Appearance = appearance81;
			this.txtNotas.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtNotas.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtNotas.Enabled = false;
			this.txtNotas.Location = new System.Drawing.Point(344, 200);
			this.txtNotas.Multiline = true;
			this.txtNotas.Name = "txtNotas";
			this.txtNotas.Size = new System.Drawing.Size(320, 64);
			this.txtNotas.TabIndex = 888;
			// 
			// chkAfiliadoTributario
			// 
			appearance82.ForeColorDisabled = System.Drawing.Color.Black;
			this.chkAfiliadoTributario.Appearance = appearance82;
			this.chkAfiliadoTributario.BackColor = System.Drawing.Color.Transparent;
			this.chkAfiliadoTributario.CausesValidation = false;
			this.chkAfiliadoTributario.Enabled = false;
			this.chkAfiliadoTributario.Location = new System.Drawing.Point(696, 152);
			this.chkAfiliadoTributario.Name = "chkAfiliadoTributario";
			this.chkAfiliadoTributario.Size = new System.Drawing.Size(160, 22);
			this.chkAfiliadoTributario.TabIndex = 887;
			this.chkAfiliadoTributario.Text = "Afiliado Tributario";
			// 
			// chkNegocioObligadoaLlevar
			// 
			appearance83.ForeColorDisabled = System.Drawing.Color.Black;
			this.chkNegocioObligadoaLlevar.Appearance = appearance83;
			this.chkNegocioObligadoaLlevar.BackColor = System.Drawing.Color.Transparent;
			this.chkNegocioObligadoaLlevar.CausesValidation = false;
			this.chkNegocioObligadoaLlevar.Enabled = false;
			this.chkNegocioObligadoaLlevar.Location = new System.Drawing.Point(856, 152);
			this.chkNegocioObligadoaLlevar.Name = "chkNegocioObligadoaLlevar";
			this.chkNegocioObligadoaLlevar.Size = new System.Drawing.Size(224, 22);
			this.chkNegocioObligadoaLlevar.TabIndex = 886;
			this.chkNegocioObligadoaLlevar.Text = "Obligado a Lllevar Contabilidad";
			// 
			// btnNegocioTelefono2
			// 
			this.btnNegocioTelefono2.Enabled = false;
			this.btnNegocioTelefono2.Image = ((System.Drawing.Image)(resources.GetObject("btnNegocioTelefono2.Image")));
			this.btnNegocioTelefono2.Location = new System.Drawing.Point(112, 152);
			this.btnNegocioTelefono2.Name = "btnNegocioTelefono2";
			this.btnNegocioTelefono2.Size = new System.Drawing.Size(24, 23);
			this.btnNegocioTelefono2.TabIndex = 885;
			this.btnNegocioTelefono2.Click += new System.EventHandler(this.btnNegocioTelefono2_Click);
			// 
			// txtNegocioTelefono2
			// 
			this.txtNegocioTelefono2.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtNegocioTelefono2.Enabled = false;
			this.txtNegocioTelefono2.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.txtNegocioTelefono2.Location = new System.Drawing.Point(8, 152);
			this.txtNegocioTelefono2.MaxLength = 9;
			this.txtNegocioTelefono2.Name = "txtNegocioTelefono2";
			this.txtNegocioTelefono2.Size = new System.Drawing.Size(100, 22);
			this.txtNegocioTelefono2.TabIndex = 884;
			this.txtNegocioTelefono2.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtNegocioTelefono2_KeyDown);
			this.txtNegocioTelefono2.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.txtNegocioTelefono2_KeyPress);
			// 
			// label96
			// 
			this.label96.AutoSize = true;
			this.label96.BackColor = System.Drawing.Color.Transparent;
			this.label96.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label96.Location = new System.Drawing.Point(680, 21);
			this.label96.Name = "label96";
			this.label96.Size = new System.Drawing.Size(125, 17);
			this.label96.TabIndex = 883;
			this.label96.Text = "Subactividad Económico";
			this.label96.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.label96.Visible = false;
			// 
			// label95
			// 
			this.label95.AutoSize = true;
			this.label95.BackColor = System.Drawing.Color.Transparent;
			this.label95.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label95.Location = new System.Drawing.Point(520, 21);
			this.label95.Name = "label95";
			this.label95.Size = new System.Drawing.Size(26, 17);
			this.label95.TabIndex = 882;
			this.label95.Text = "Tipo";
			this.label95.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.label95.Visible = false;
			// 
			// label94
			// 
			this.label94.AutoSize = true;
			this.label94.BackColor = System.Drawing.Color.Transparent;
			this.label94.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label94.Location = new System.Drawing.Point(432, 21);
			this.label94.Name = "label94";
			this.label94.Size = new System.Drawing.Size(58, 17);
			this.label94.TabIndex = 881;
			this.label94.Text = "Empleados";
			this.label94.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.label94.Visible = false;
			// 
			// label93
			// 
			this.label93.AutoSize = true;
			this.label93.BackColor = System.Drawing.Color.Transparent;
			this.label93.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label93.Location = new System.Drawing.Point(328, 21);
			this.label93.Name = "label93";
			this.label93.Size = new System.Drawing.Size(38, 17);
			this.label93.TabIndex = 880;
			this.label93.Text = "Metros";
			this.label93.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// label87
			// 
			this.label87.AutoSize = true;
			this.label87.BackColor = System.Drawing.Color.Transparent;
			this.label87.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label87.Location = new System.Drawing.Point(648, 56);
			this.label87.Name = "label87";
			this.label87.Size = new System.Drawing.Size(31, 17);
			this.label87.TabIndex = 879;
			this.label87.Text = "Email";
			this.label87.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.label87.Visible = false;
			// 
			// txtNegocioemail
			// 
			appearance84.BackColorDisabled = System.Drawing.Color.Silver;
			appearance84.BackColorDisabled2 = System.Drawing.Color.Silver;
			appearance84.BorderColor = System.Drawing.Color.Black;
			appearance84.BorderColor3DBase = System.Drawing.Color.Black;
			appearance84.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtNegocioemail.Appearance = appearance84;
			this.txtNegocioemail.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtNegocioemail.Enabled = false;
			this.txtNegocioemail.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.txtNegocioemail.Location = new System.Drawing.Point(648, 72);
			this.txtNegocioemail.Name = "txtNegocioemail";
			this.txtNegocioemail.Size = new System.Drawing.Size(120, 22);
			this.txtNegocioemail.TabIndex = 878;
			this.txtNegocioemail.Visible = false;
			this.txtNegocioemail.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtNegocioemail_KeyDown);
			// 
			// label88
			// 
			this.label88.AutoSize = true;
			this.label88.BackColor = System.Drawing.Color.Transparent;
			this.label88.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label88.Location = new System.Drawing.Point(488, 56);
			this.label88.Name = "label88";
			this.label88.Size = new System.Drawing.Size(34, 17);
			this.label88.TabIndex = 877;
			this.label88.Text = "Barrio";
			this.label88.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// label90
			// 
			this.label90.AutoSize = true;
			this.label90.BackColor = System.Drawing.Color.Transparent;
			this.label90.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label90.Location = new System.Drawing.Point(328, 56);
			this.label90.Name = "label90";
			this.label90.Size = new System.Drawing.Size(52, 17);
			this.label90.TabIndex = 876;
			this.label90.Text = "Parroquia";
			this.label90.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.label90.Click += new System.EventHandler(this.label90_Click);
			// 
			// label91
			// 
			this.label91.AutoSize = true;
			this.label91.BackColor = System.Drawing.Color.Transparent;
			this.label91.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label91.Location = new System.Drawing.Point(168, 56);
			this.label91.Name = "label91";
			this.label91.Size = new System.Drawing.Size(39, 17);
			this.label91.TabIndex = 875;
			this.label91.Text = "Cantón";
			this.label91.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// label92
			// 
			this.label92.AutoSize = true;
			this.label92.BackColor = System.Drawing.Color.Transparent;
			this.label92.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label92.Location = new System.Drawing.Point(8, 56);
			this.label92.Name = "label92";
			this.label92.Size = new System.Drawing.Size(49, 17);
			this.label92.TabIndex = 874;
			this.label92.Text = "Provincia";
			this.label92.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// cmbNegocioBarrio
			// 
			this.cmbNegocioBarrio.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbNegocioBarrio.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbNegocioBarrio.DataSource = this.ultraDataSource4;
			ultraGridColumn108.Header.VisiblePosition = 0;
			ultraGridColumn108.Hidden = true;
			ultraGridColumn109.Header.Caption = "Barrio";
			ultraGridColumn109.Header.VisiblePosition = 1;
			ultraGridColumn109.Width = 150;
			ultraGridBand42.Columns.AddRange(new object[] {
																											ultraGridColumn108,
																											ultraGridColumn109});
			this.cmbNegocioBarrio.DisplayLayout.BandsSerializer.Add(ultraGridBand42);
			this.cmbNegocioBarrio.DisplayLayout.Override.RowSelectors = Infragistics.Win.DefaultableBoolean.True;
			this.cmbNegocioBarrio.DisplayMember = "Nombre";
			this.cmbNegocioBarrio.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbNegocioBarrio.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbNegocioBarrio.Enabled = false;
			this.cmbNegocioBarrio.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbNegocioBarrio.Location = new System.Drawing.Point(488, 72);
			this.cmbNegocioBarrio.Name = "cmbNegocioBarrio";
			this.cmbNegocioBarrio.Size = new System.Drawing.Size(150, 22);
			this.cmbNegocioBarrio.TabIndex = 873;
			this.cmbNegocioBarrio.ValueMember = "idBarrio";
			this.cmbNegocioBarrio.KeyDown += new System.Windows.Forms.KeyEventHandler(this.cmbNegocioBarrio_KeyDown);
			// 
			// cmbNegocioParroquia
			// 
			this.cmbNegocioParroquia.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbNegocioParroquia.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbNegocioParroquia.DataSource = this.ultraDataSource3;
			ultraGridColumn110.Header.VisiblePosition = 0;
			ultraGridColumn110.Hidden = true;
			ultraGridColumn111.Header.Caption = "Provincia";
			ultraGridColumn111.Header.VisiblePosition = 1;
			ultraGridColumn111.Width = 150;
			ultraGridBand43.Columns.AddRange(new object[] {
																											ultraGridColumn110,
																											ultraGridColumn111});
			this.cmbNegocioParroquia.DisplayLayout.BandsSerializer.Add(ultraGridBand43);
			this.cmbNegocioParroquia.DisplayMember = "Nombre";
			this.cmbNegocioParroquia.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbNegocioParroquia.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbNegocioParroquia.Enabled = false;
			this.cmbNegocioParroquia.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbNegocioParroquia.Location = new System.Drawing.Point(328, 72);
			this.cmbNegocioParroquia.Name = "cmbNegocioParroquia";
			this.cmbNegocioParroquia.Size = new System.Drawing.Size(150, 22);
			this.cmbNegocioParroquia.TabIndex = 872;
			this.cmbNegocioParroquia.ValueMember = "idParroquia";
			this.cmbNegocioParroquia.KeyDown += new System.Windows.Forms.KeyEventHandler(this.cmbNegocioParroquia_KeyDown);
			this.cmbNegocioParroquia.ValueChanged += new System.EventHandler(this.cmbNegocioParroquia_ValueChanged);
			// 
			// cmbNegocioCanton
			// 
			this.cmbNegocioCanton.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbNegocioCanton.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbNegocioCanton.DataSource = this.ultraDataSource2;
			ultraGridColumn112.Header.VisiblePosition = 0;
			ultraGridColumn112.Hidden = true;
			ultraGridColumn113.Header.Caption = "Canton";
			ultraGridColumn113.Header.VisiblePosition = 1;
			ultraGridColumn113.Width = 150;
			ultraGridBand44.Columns.AddRange(new object[] {
																											ultraGridColumn112,
																											ultraGridColumn113});
			this.cmbNegocioCanton.DisplayLayout.BandsSerializer.Add(ultraGridBand44);
			this.cmbNegocioCanton.DisplayLayout.Override.RowSelectors = Infragistics.Win.DefaultableBoolean.True;
			this.cmbNegocioCanton.DisplayMember = "Nombre";
			this.cmbNegocioCanton.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbNegocioCanton.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbNegocioCanton.Enabled = false;
			this.cmbNegocioCanton.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbNegocioCanton.Location = new System.Drawing.Point(168, 72);
			this.cmbNegocioCanton.Name = "cmbNegocioCanton";
			this.cmbNegocioCanton.Size = new System.Drawing.Size(150, 22);
			this.cmbNegocioCanton.TabIndex = 871;
			this.cmbNegocioCanton.ValueMember = "idCanton";
			this.cmbNegocioCanton.KeyDown += new System.Windows.Forms.KeyEventHandler(this.cmbNegocioCanton_KeyDown);
			this.cmbNegocioCanton.ValueChanged += new System.EventHandler(this.cmbNegocioCanton_ValueChanged);
			// 
			// cmbNegocioProvincia
			// 
			this.cmbNegocioProvincia.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbNegocioProvincia.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbNegocioProvincia.DataSource = this.ultraDataSource1;
			ultraGridColumn114.Header.VisiblePosition = 0;
			ultraGridColumn114.Hidden = true;
			ultraGridColumn114.Width = 50;
			ultraGridColumn115.Header.VisiblePosition = 1;
			ultraGridColumn115.Width = 150;
			ultraGridBand45.Columns.AddRange(new object[] {
																											ultraGridColumn114,
																											ultraGridColumn115});
			this.cmbNegocioProvincia.DisplayLayout.BandsSerializer.Add(ultraGridBand45);
			this.cmbNegocioProvincia.DisplayMember = "Nombre";
			this.cmbNegocioProvincia.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbNegocioProvincia.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbNegocioProvincia.Enabled = false;
			this.cmbNegocioProvincia.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbNegocioProvincia.Location = new System.Drawing.Point(8, 72);
			this.cmbNegocioProvincia.Name = "cmbNegocioProvincia";
			this.cmbNegocioProvincia.Size = new System.Drawing.Size(150, 22);
			this.cmbNegocioProvincia.TabIndex = 870;
			this.cmbNegocioProvincia.ValueMember = "idProvincia";
			this.cmbNegocioProvincia.KeyDown += new System.Windows.Forms.KeyEventHandler(this.cmbNegocioProvincia_KeyDown);
			this.cmbNegocioProvincia.ValueChanged += new System.EventHandler(this.cmbNegocioProvincia_ValueChanged);
			// 
			// label75
			// 
			this.label75.AutoSize = true;
			this.label75.BackColor = System.Drawing.Color.Transparent;
			this.label75.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label75.Location = new System.Drawing.Point(592, 136);
			this.label75.Name = "label75";
			this.label75.Size = new System.Drawing.Size(41, 17);
			this.label75.TabIndex = 869;
			this.label75.Text = "Pasivos";
			this.label75.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.label75.Visible = false;
			// 
			// label76
			// 
			this.label76.AutoSize = true;
			this.label76.BackColor = System.Drawing.Color.Transparent;
			this.label76.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label76.Location = new System.Drawing.Point(488, 136);
			this.label76.Name = "label76";
			this.label76.Size = new System.Drawing.Size(39, 17);
			this.label76.TabIndex = 868;
			this.label76.Text = "Activos";
			this.label76.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.label76.Visible = false;
			// 
			// label77
			// 
			this.label77.AutoSize = true;
			this.label77.BackColor = System.Drawing.Color.Transparent;
			this.label77.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label77.Location = new System.Drawing.Point(392, 136);
			this.label77.Name = "label77";
			this.label77.Size = new System.Drawing.Size(38, 17);
			this.label77.TabIndex = 867;
			this.label77.Text = "Gastos";
			this.label77.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// label78
			// 
			this.label78.AutoSize = true;
			this.label78.BackColor = System.Drawing.Color.Transparent;
			this.label78.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label78.Location = new System.Drawing.Point(288, 136);
			this.label78.Name = "label78";
			this.label78.Size = new System.Drawing.Size(47, 17);
			this.label78.TabIndex = 866;
			this.label78.Text = "Ingresos";
			this.label78.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// txtNEgocioPasivos
			// 
			this.txtNEgocioPasivos.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtNEgocioPasivos.Enabled = false;
			this.txtNEgocioPasivos.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.txtNEgocioPasivos.Location = new System.Drawing.Point(592, 152);
			this.txtNEgocioPasivos.MaxValue = 1000000;
			this.txtNEgocioPasivos.MinValue = 0;
			this.txtNEgocioPasivos.Name = "txtNEgocioPasivos";
			this.txtNEgocioPasivos.PromptChar = ' ';
			this.txtNEgocioPasivos.Size = new System.Drawing.Size(90, 22);
			this.txtNEgocioPasivos.SpinButtonDisplayStyle = Infragistics.Win.ButtonDisplayStyle.Always;
			this.txtNEgocioPasivos.TabIndex = 865;
			this.txtNEgocioPasivos.Value = 100;
			this.txtNEgocioPasivos.Visible = false;
			this.txtNEgocioPasivos.Click += new System.EventHandler(this.txtNEgocioPasivos_Click);
			this.txtNEgocioPasivos.ValueChanged += new System.EventHandler(this.txtNEgocioPasivos_ValueChanged);
			// 
			// txtNegocioActivos
			// 
			this.txtNegocioActivos.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtNegocioActivos.Enabled = false;
			this.txtNegocioActivos.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.txtNegocioActivos.Location = new System.Drawing.Point(488, 152);
			this.txtNegocioActivos.MaxValue = 100000;
			this.txtNegocioActivos.MinValue = 0;
			this.txtNegocioActivos.Name = "txtNegocioActivos";
			this.txtNegocioActivos.PromptChar = ' ';
			this.txtNegocioActivos.Size = new System.Drawing.Size(90, 22);
			this.txtNegocioActivos.SpinButtonDisplayStyle = Infragistics.Win.ButtonDisplayStyle.Always;
			this.txtNegocioActivos.TabIndex = 864;
			this.txtNegocioActivos.Value = 450;
			this.txtNegocioActivos.Visible = false;
			this.txtNegocioActivos.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtNegocioActivos_KeyDown);
			this.txtNegocioActivos.Click += new System.EventHandler(this.txtNegocioActivos_Click);
			// 
			// txtNEgocioGastos
			// 
			this.txtNEgocioGastos.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtNEgocioGastos.Enabled = false;
			this.txtNEgocioGastos.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.txtNEgocioGastos.Location = new System.Drawing.Point(392, 152);
			this.txtNEgocioGastos.MaxValue = 100000;
			this.txtNEgocioGastos.MinValue = 0;
			this.txtNEgocioGastos.Name = "txtNEgocioGastos";
			this.txtNEgocioGastos.PromptChar = ' ';
			this.txtNEgocioGastos.Size = new System.Drawing.Size(90, 22);
			this.txtNEgocioGastos.SpinButtonDisplayStyle = Infragistics.Win.ButtonDisplayStyle.Always;
			this.txtNEgocioGastos.TabIndex = 863;
			this.txtNEgocioGastos.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtNEgocioGastos_KeyDown);
			this.txtNEgocioGastos.Click += new System.EventHandler(this.txtNEgocioGastos_Click);
			// 
			// txtNEgocioIngresos
			// 
			this.txtNEgocioIngresos.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtNEgocioIngresos.Enabled = false;
			this.txtNEgocioIngresos.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.txtNEgocioIngresos.Location = new System.Drawing.Point(288, 152);
			this.txtNEgocioIngresos.MaxValue = 100000;
			this.txtNEgocioIngresos.MinValue = 0;
			this.txtNEgocioIngresos.Name = "txtNEgocioIngresos";
			this.txtNEgocioIngresos.PromptChar = ' ';
			this.txtNEgocioIngresos.Size = new System.Drawing.Size(90, 22);
			this.txtNEgocioIngresos.SpinButtonDisplayStyle = Infragistics.Win.ButtonDisplayStyle.Always;
			this.txtNEgocioIngresos.TabIndex = 862;
			this.txtNEgocioIngresos.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtNEgocioIngresos_KeyDown);
			this.txtNEgocioIngresos.Click += new System.EventHandler(this.txtNEgocioIngresos_Click);
			// 
			// label79
			// 
			this.label79.AutoSize = true;
			this.label79.BackColor = System.Drawing.Color.Transparent;
			this.label79.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label79.Location = new System.Drawing.Point(8, 176);
			this.label79.Name = "label79";
			this.label79.Size = new System.Drawing.Size(108, 17);
			this.label79.TabIndex = 861;
			this.label79.Text = "Referencia Ubicacion";
			this.label79.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// txtNegocioReferncia
			// 
			this.txtNegocioReferncia.AcceptsReturn = true;
			appearance85.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtNegocioReferncia.Appearance = appearance85;
			this.txtNegocioReferncia.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtNegocioReferncia.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtNegocioReferncia.Enabled = false;
			this.txtNegocioReferncia.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.txtNegocioReferncia.Location = new System.Drawing.Point(8, 200);
			this.txtNegocioReferncia.Multiline = true;
			this.txtNegocioReferncia.Name = "txtNegocioReferncia";
			this.txtNegocioReferncia.Size = new System.Drawing.Size(300, 66);
			this.txtNegocioReferncia.TabIndex = 860;
			this.txtNegocioReferncia.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtNegocioReferncia_KeyDown);
			this.txtNegocioReferncia.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.txtNegocioReferncia_KeyPress);
			// 
			// btnNegocioCelular
			// 
			this.btnNegocioCelular.Enabled = false;
			this.btnNegocioCelular.Image = ((System.Drawing.Image)(resources.GetObject("btnNegocioCelular.Image")));
			this.btnNegocioCelular.Location = new System.Drawing.Point(248, 152);
			this.btnNegocioCelular.Name = "btnNegocioCelular";
			this.btnNegocioCelular.Size = new System.Drawing.Size(24, 23);
			this.btnNegocioCelular.TabIndex = 859;
			this.btnNegocioCelular.Click += new System.EventHandler(this.btnNegocioCelular_Click);
			// 
			// btnNegocioTelefono1
			// 
			this.btnNegocioTelefono1.Enabled = false;
			this.btnNegocioTelefono1.Image = ((System.Drawing.Image)(resources.GetObject("btnNegocioTelefono1.Image")));
			this.btnNegocioTelefono1.Location = new System.Drawing.Point(976, 208);
			this.btnNegocioTelefono1.Name = "btnNegocioTelefono1";
			this.btnNegocioTelefono1.Size = new System.Drawing.Size(24, 23);
			this.btnNegocioTelefono1.TabIndex = 858;
			this.btnNegocioTelefono1.Visible = false;
			this.btnNegocioTelefono1.Click += new System.EventHandler(this.btnNegocioTelefono1_Click);
			// 
			// label80
			// 
			this.label80.AutoSize = true;
			this.label80.BackColor = System.Drawing.Color.Transparent;
			this.label80.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label80.Location = new System.Drawing.Point(144, 136);
			this.label80.Name = "label80";
			this.label80.Size = new System.Drawing.Size(38, 17);
			this.label80.TabIndex = 857;
			this.label80.Text = "Celular";
			this.label80.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// label82
			// 
			this.label82.AutoSize = true;
			this.label82.BackColor = System.Drawing.Color.Transparent;
			this.label82.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label82.Location = new System.Drawing.Point(8, 136);
			this.label82.Name = "label82";
			this.label82.Size = new System.Drawing.Size(57, 17);
			this.label82.TabIndex = 856;
			this.label82.Text = "Teléfono 2";
			this.label82.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.label82.Click += new System.EventHandler(this.label82_Click);
			// 
			// label83
			// 
			this.label83.AutoSize = true;
			this.label83.BackColor = System.Drawing.Color.Transparent;
			this.label83.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label83.Location = new System.Drawing.Point(872, 192);
			this.label83.Name = "label83";
			this.label83.Size = new System.Drawing.Size(57, 17);
			this.label83.TabIndex = 855;
			this.label83.Text = "Teléfono 1";
			this.label83.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.label83.Visible = false;
			// 
			// label84
			// 
			this.label84.AutoSize = true;
			this.label84.BackColor = System.Drawing.Color.Transparent;
			this.label84.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label84.Location = new System.Drawing.Point(496, 96);
			this.label84.Name = "label84";
			this.label84.Size = new System.Drawing.Size(87, 17);
			this.label84.TabIndex = 854;
			this.label84.Text = "Calle Secundaria";
			this.label84.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// label85
			// 
			this.label85.AutoSize = true;
			this.label85.BackColor = System.Drawing.Color.Transparent;
			this.label85.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label85.Location = new System.Drawing.Point(312, 96);
			this.label85.Name = "label85";
			this.label85.Size = new System.Drawing.Size(71, 17);
			this.label85.TabIndex = 853;
			this.label85.Text = "Número Casa";
			this.label85.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// label86
			// 
			this.label86.AutoSize = true;
			this.label86.BackColor = System.Drawing.Color.Transparent;
			this.label86.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label86.Location = new System.Drawing.Point(8, 96);
			this.label86.Name = "label86";
			this.label86.Size = new System.Drawing.Size(73, 17);
			this.label86.TabIndex = 852;
			this.label86.Text = "Calle Principal";
			this.label86.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// txtNEgocioCelular
			// 
			this.txtNEgocioCelular.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtNEgocioCelular.Enabled = false;
			this.txtNEgocioCelular.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.txtNEgocioCelular.Location = new System.Drawing.Point(144, 152);
			this.txtNEgocioCelular.MaxLength = 10;
			this.txtNEgocioCelular.Name = "txtNEgocioCelular";
			this.txtNEgocioCelular.Size = new System.Drawing.Size(100, 22);
			this.txtNEgocioCelular.TabIndex = 851;
			this.txtNEgocioCelular.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtNEgocioCelular_KeyDown);
			this.txtNEgocioCelular.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.txtNEgocioCelular_KeyPress);
			// 
			// txtNegocioTelefono1
			// 
			this.txtNegocioTelefono1.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtNegocioTelefono1.Enabled = false;
			this.txtNegocioTelefono1.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.txtNegocioTelefono1.Location = new System.Drawing.Point(872, 208);
			this.txtNegocioTelefono1.MaxLength = 9;
			this.txtNegocioTelefono1.Name = "txtNegocioTelefono1";
			this.txtNegocioTelefono1.Size = new System.Drawing.Size(100, 22);
			this.txtNegocioTelefono1.TabIndex = 849;
			this.txtNegocioTelefono1.Visible = false;
			this.txtNegocioTelefono1.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtNegocioTelefono1_KeyDown);
			this.txtNegocioTelefono1.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.txtNegocioTelefono1_KeyPress);
			// 
			// txtNegocioCalleSecundaria
			// 
			appearance86.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtNegocioCalleSecundaria.Appearance = appearance86;
			this.txtNegocioCalleSecundaria.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtNegocioCalleSecundaria.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtNegocioCalleSecundaria.Enabled = false;
			this.txtNegocioCalleSecundaria.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.txtNegocioCalleSecundaria.Location = new System.Drawing.Point(496, 112);
			this.txtNegocioCalleSecundaria.Name = "txtNegocioCalleSecundaria";
			this.txtNegocioCalleSecundaria.Size = new System.Drawing.Size(300, 22);
			this.txtNegocioCalleSecundaria.TabIndex = 848;
			this.txtNegocioCalleSecundaria.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtNegocioCalleSecundaria_KeyDown);
			this.txtNegocioCalleSecundaria.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.txtNegocioCalleSecundaria_KeyPress);
			this.txtNegocioCalleSecundaria.ValueChanged += new System.EventHandler(this.txtNegocioCalleSecundaria_ValueChanged);
			// 
			// txtNegocioNumeroCasa
			// 
			appearance87.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtNegocioNumeroCasa.Appearance = appearance87;
			this.txtNegocioNumeroCasa.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtNegocioNumeroCasa.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtNegocioNumeroCasa.Enabled = false;
			this.txtNegocioNumeroCasa.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.txtNegocioNumeroCasa.Location = new System.Drawing.Point(312, 112);
			this.txtNegocioNumeroCasa.Name = "txtNegocioNumeroCasa";
			this.txtNegocioNumeroCasa.Size = new System.Drawing.Size(180, 22);
			this.txtNegocioNumeroCasa.TabIndex = 847;
			this.txtNegocioNumeroCasa.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtNegocioNumeroCasa_KeyDown);
			this.txtNegocioNumeroCasa.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.txtNegocioNumeroCasa_KeyPress);
			// 
			// txtNegocioCallePrincipal
			// 
			appearance88.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtNegocioCallePrincipal.Appearance = appearance88;
			this.txtNegocioCallePrincipal.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtNegocioCallePrincipal.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtNegocioCallePrincipal.Enabled = false;
			this.txtNegocioCallePrincipal.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.txtNegocioCallePrincipal.Location = new System.Drawing.Point(8, 112);
			this.txtNegocioCallePrincipal.Name = "txtNegocioCallePrincipal";
			this.txtNegocioCallePrincipal.Size = new System.Drawing.Size(300, 22);
			this.txtNegocioCallePrincipal.TabIndex = 846;
			this.txtNegocioCallePrincipal.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtNegocioCallePrincipal_KeyDown);
			this.txtNegocioCallePrincipal.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.txtNegocioCallePrincipal_KeyPress);
			// 
			// txtNegocioEmpleados
			// 
			this.txtNegocioEmpleados.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtNegocioEmpleados.Enabled = false;
			this.txtNegocioEmpleados.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.txtNegocioEmpleados.Location = new System.Drawing.Point(432, 35);
			this.txtNegocioEmpleados.MaxValue = 100;
			this.txtNegocioEmpleados.MinValue = 0;
			this.txtNegocioEmpleados.Name = "txtNegocioEmpleados";
			this.txtNegocioEmpleados.PromptChar = ' ';
			this.txtNegocioEmpleados.Size = new System.Drawing.Size(90, 22);
			this.txtNegocioEmpleados.SpinButtonDisplayStyle = Infragistics.Win.ButtonDisplayStyle.Always;
			this.txtNegocioEmpleados.TabIndex = 840;
			this.txtNegocioEmpleados.Visible = false;
			this.txtNegocioEmpleados.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtNegocioEmpleados_KeyDown);
			this.txtNegocioEmpleados.Click += new System.EventHandler(this.txtNegocioEmpleados_Click);
			// 
			// txtNegocioMetros
			// 
			this.txtNegocioMetros.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtNegocioMetros.Enabled = false;
			this.txtNegocioMetros.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.txtNegocioMetros.Location = new System.Drawing.Point(328, 35);
			this.txtNegocioMetros.MaxValue = 100;
			this.txtNegocioMetros.MinValue = 0;
			this.txtNegocioMetros.Name = "txtNegocioMetros";
			this.txtNegocioMetros.PromptChar = ' ';
			this.txtNegocioMetros.Size = new System.Drawing.Size(90, 22);
			this.txtNegocioMetros.SpinButtonDisplayStyle = Infragistics.Win.ButtonDisplayStyle.Always;
			this.txtNegocioMetros.TabIndex = 839;
			this.txtNegocioMetros.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtNegocioMetros_KeyDown);
			this.txtNegocioMetros.Click += new System.EventHandler(this.txtNegocioMetros_Click);
			// 
			// Nombre
			// 
			this.Nombre.AutoSize = true;
			this.Nombre.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.Nombre.Location = new System.Drawing.Point(7, 21);
			this.Nombre.Name = "Nombre";
			this.Nombre.Size = new System.Drawing.Size(103, 17);
			this.Nombre.TabIndex = 824;
			this.Nombre.Text = "Nombre de Negocio";
			this.Nombre.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// txtNombreDelNegocio
			// 
			appearance89.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtNombreDelNegocio.Appearance = appearance89;
			this.txtNombreDelNegocio.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtNombreDelNegocio.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtNombreDelNegocio.Enabled = false;
			this.txtNombreDelNegocio.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.txtNombreDelNegocio.Location = new System.Drawing.Point(7, 35);
			this.txtNombreDelNegocio.Name = "txtNombreDelNegocio";
			this.txtNombreDelNegocio.Size = new System.Drawing.Size(180, 22);
			this.txtNombreDelNegocio.TabIndex = 828;
			this.txtNombreDelNegocio.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtNombreDelNegocio_KeyDown);
			this.txtNombreDelNegocio.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.txtNombreDelNegocio_KeyPress);
			// 
			// label81
			// 
			this.label81.AutoSize = true;
			this.label81.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label81.Location = new System.Drawing.Point(907, 125);
			this.label81.Name = "label81";
			this.label81.Size = new System.Drawing.Size(23, 17);
			this.label81.TabIndex = 818;
			this.label81.Text = "Ruc";
			this.label81.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.label81.Visible = false;
			// 
			// txtNegocioRuc
			// 
			appearance90.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtNegocioRuc.Appearance = appearance90;
			this.txtNegocioRuc.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtNegocioRuc.Enabled = false;
			this.txtNegocioRuc.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.txtNegocioRuc.Location = new System.Drawing.Point(907, 139);
			this.txtNegocioRuc.MaxLength = 13;
			this.txtNegocioRuc.Name = "txtNegocioRuc";
			this.txtNegocioRuc.Size = new System.Drawing.Size(127, 22);
			this.txtNegocioRuc.TabIndex = 819;
			this.txtNegocioRuc.Visible = false;
			this.txtNegocioRuc.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtNegocioRuc_KeyDown);
			this.txtNegocioRuc.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.txtNegocioRuc_KeyPress);
			// 
			// cmbNegciotipo
			// 
			this.cmbNegciotipo.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbNegciotipo.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbNegciotipo.DataSource = this.ultraDataSource37;
			ultraGridColumn116.Header.VisiblePosition = 0;
			ultraGridColumn116.Hidden = true;
			ultraGridColumn117.Header.VisiblePosition = 1;
			ultraGridColumn117.Width = 150;
			ultraGridBand46.Columns.AddRange(new object[] {
																											ultraGridColumn116,
																											ultraGridColumn117});
			this.cmbNegciotipo.DisplayLayout.BandsSerializer.Add(ultraGridBand46);
			this.cmbNegciotipo.DisplayMember = "Nombre";
			this.cmbNegciotipo.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbNegciotipo.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbNegciotipo.Enabled = false;
			this.cmbNegciotipo.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbNegciotipo.Location = new System.Drawing.Point(520, 35);
			this.cmbNegciotipo.Name = "cmbNegciotipo";
			this.cmbNegciotipo.Size = new System.Drawing.Size(150, 22);
			this.cmbNegciotipo.TabIndex = 810;
			this.cmbNegciotipo.ValueMember = "idTipoVivienda";
			this.cmbNegciotipo.Visible = false;
			this.cmbNegciotipo.KeyDown += new System.Windows.Forms.KeyEventHandler(this.cmbNegciotipo_KeyDown);
			// 
			// ultraDataSource37
			// 
			ultraDataColumn73.DataType = typeof(int);
			this.ultraDataSource37.Band.Columns.AddRange(new object[] {
																																	ultraDataColumn73,
																																	ultraDataColumn74});
			// 
			// label89
			// 
			this.label89.AutoSize = true;
			this.label89.BackColor = System.Drawing.Color.Transparent;
			this.label89.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label89.Location = new System.Drawing.Point(200, 21);
			this.label89.Name = "label89";
			this.label89.Size = new System.Drawing.Size(103, 17);
			this.label89.TabIndex = 786;
			this.label89.Text = "Tiempo del Negocio";
			this.label89.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// cmbNEgocioActivdadEconomica
			// 
			this.cmbNEgocioActivdadEconomica.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbNEgocioActivdadEconomica.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbNEgocioActivdadEconomica.DataSource = this.ultraDataSource25;
			ultraGridColumn118.Header.VisiblePosition = 0;
			ultraGridColumn118.Hidden = true;
			ultraGridColumn119.Header.VisiblePosition = 1;
			ultraGridColumn119.Width = 248;
			ultraGridBand47.Columns.AddRange(new object[] {
																											ultraGridColumn118,
																											ultraGridColumn119});
			this.cmbNEgocioActivdadEconomica.DisplayLayout.BandsSerializer.Add(ultraGridBand47);
			this.cmbNEgocioActivdadEconomica.DisplayMember = "Nombre";
			this.cmbNEgocioActivdadEconomica.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbNEgocioActivdadEconomica.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbNEgocioActivdadEconomica.Enabled = false;
			this.cmbNEgocioActivdadEconomica.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbNEgocioActivdadEconomica.Location = new System.Drawing.Point(680, 35);
			this.cmbNEgocioActivdadEconomica.Name = "cmbNEgocioActivdadEconomica";
			this.cmbNEgocioActivdadEconomica.Size = new System.Drawing.Size(248, 22);
			this.cmbNEgocioActivdadEconomica.TabIndex = 782;
			this.cmbNEgocioActivdadEconomica.ValueMember = "idSubActEcoNegocio";
			this.cmbNEgocioActivdadEconomica.Visible = false;
			this.cmbNEgocioActivdadEconomica.KeyDown += new System.Windows.Forms.KeyEventHandler(this.cmbNEgocioActivdadEconomica_KeyDown);
			// 
			// ultraDataSource25
			// 
			ultraDataColumn75.DataType = typeof(int);
			this.ultraDataSource25.Band.Columns.AddRange(new object[] {
																																	ultraDataColumn75,
																																	ultraDataColumn76});
			// 
			// dtpNegocioFechaInicio
			// 
			this.dtpNegocioFechaInicio.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			dateButton4.Caption = "Today";
			this.dtpNegocioFechaInicio.DateButtons.Add(dateButton4);
			this.dtpNegocioFechaInicio.Enabled = false;
			this.dtpNegocioFechaInicio.Format = "dd/MM/yyyy";
			this.dtpNegocioFechaInicio.Location = new System.Drawing.Point(928, 80);
			this.dtpNegocioFechaInicio.Name = "dtpNegocioFechaInicio";
			this.dtpNegocioFechaInicio.NonAutoSizeHeight = 23;
			this.dtpNegocioFechaInicio.Size = new System.Drawing.Size(90, 21);
			this.dtpNegocioFechaInicio.TabIndex = 779;
			this.dtpNegocioFechaInicio.Value = ((object)(resources.GetObject("dtpNegocioFechaInicio.Value")));
			this.dtpNegocioFechaInicio.Visible = false;
			this.dtpNegocioFechaInicio.KeyDown += new System.Windows.Forms.KeyEventHandler(this.dtpNegocioFechaInicio_KeyDown);
			// 
			// ultraTabPageControl2
			// 
			this.ultraTabPageControl2.Controls.Add(this.groupBox6);
			this.ultraTabPageControl2.Location = new System.Drawing.Point(-10000, -10000);
			this.ultraTabPageControl2.Name = "ultraTabPageControl2";
			this.ultraTabPageControl2.Size = new System.Drawing.Size(1116, 366);
			// 
			// groupBox6
			// 
			this.groupBox6.BackColor = System.Drawing.SystemColors.Control;
			this.groupBox6.Controls.Add(this.cmbTiempoDependiente);
			this.groupBox6.Controls.Add(this.cmbDependienteTipoempresa);
			this.groupBox6.Controls.Add(this.cmbDependienteTipoSueldo);
			this.groupBox6.Controls.Add(this.label121);
			this.groupBox6.Controls.Add(this.label120);
			this.groupBox6.Controls.Add(this.label100);
			this.groupBox6.Controls.Add(this.label99);
			this.groupBox6.Controls.Add(this.label98);
			this.groupBox6.Controls.Add(this.label97);
			this.groupBox6.Controls.Add(this.cmbDependienteCargo);
			this.groupBox6.Controls.Add(this.txtDependienteDepartamento);
			this.groupBox6.Controls.Add(this.chkAfiliadoIESS);
			this.groupBox6.Controls.Add(this.btnTrabajoCelular);
			this.groupBox6.Controls.Add(this.txtDependienteEXT);
			this.groupBox6.Controls.Add(this.label101);
			this.groupBox6.Controls.Add(this.txtDependienteemail);
			this.groupBox6.Controls.Add(this.label102);
			this.groupBox6.Controls.Add(this.label103);
			this.groupBox6.Controls.Add(this.label104);
			this.groupBox6.Controls.Add(this.label105);
			this.groupBox6.Controls.Add(this.cmbDependienteBarrio);
			this.groupBox6.Controls.Add(this.cmbDependienteParroquia);
			this.groupBox6.Controls.Add(this.cmbDependienteCanton);
			this.groupBox6.Controls.Add(this.cmbDependienteProvincia);
			this.groupBox6.Controls.Add(this.label106);
			this.groupBox6.Controls.Add(this.label107);
			this.groupBox6.Controls.Add(this.label108);
			this.groupBox6.Controls.Add(this.label109);
			this.groupBox6.Controls.Add(this.txtDependientePasivos);
			this.groupBox6.Controls.Add(this.txtDependienteActivos);
			this.groupBox6.Controls.Add(this.txtDependienteGastos);
			this.groupBox6.Controls.Add(this.txtDependienteIngresos);
			this.groupBox6.Controls.Add(this.label110);
			this.groupBox6.Controls.Add(this.txtDependienteReferencia);
			this.groupBox6.Controls.Add(this.btnTrabajoTelefono);
			this.groupBox6.Controls.Add(this.label111);
			this.groupBox6.Controls.Add(this.label112);
			this.groupBox6.Controls.Add(this.label113);
			this.groupBox6.Controls.Add(this.label114);
			this.groupBox6.Controls.Add(this.label115);
			this.groupBox6.Controls.Add(this.label116);
			this.groupBox6.Controls.Add(this.txtDependienteCelular);
			this.groupBox6.Controls.Add(this.txtDependienteTelefono);
			this.groupBox6.Controls.Add(this.txtDependienteCalleSecundaria);
			this.groupBox6.Controls.Add(this.txtDependienteNumeroCasa);
			this.groupBox6.Controls.Add(this.txtDependienteCallePrinciapl);
			this.groupBox6.Controls.Add(this.txtDependienteDiasPago);
			this.groupBox6.Controls.Add(this.label117);
			this.groupBox6.Controls.Add(this.txtDependienteEmpresa);
			this.groupBox6.Controls.Add(this.label118);
			this.groupBox6.Controls.Add(this.txtDependienteRuc);
			this.groupBox6.Controls.Add(this.cmbDependienteTipocontrato);
			this.groupBox6.Controls.Add(this.label119);
			this.groupBox6.Controls.Add(this.dtDependienteFechaIngreso);
			this.groupBox6.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.groupBox6.Location = new System.Drawing.Point(2, 7);
			this.groupBox6.Name = "groupBox6";
			this.groupBox6.Size = new System.Drawing.Size(1110, 425);
			this.groupBox6.TabIndex = 831;
			this.groupBox6.TabStop = false;
			this.groupBox6.Enter += new System.EventHandler(this.groupBox6_Enter);
			// 
			// cmbTiempoDependiente
			// 
			this.cmbTiempoDependiente.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbTiempoDependiente.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbTiempoDependiente.DataSource = this.ultraDataSource49;
			ultraGridColumn120.Header.VisiblePosition = 0;
			ultraGridColumn120.Hidden = true;
			ultraGridColumn120.Width = 8;
			ultraGridColumn121.Header.VisiblePosition = 1;
			ultraGridColumn121.Width = 150;
			ultraGridBand48.Columns.AddRange(new object[] {
																											ultraGridColumn120,
																											ultraGridColumn121});
			this.cmbTiempoDependiente.DisplayLayout.BandsSerializer.Add(ultraGridBand48);
			this.cmbTiempoDependiente.DisplayLayout.Override.RowSelectors = Infragistics.Win.DefaultableBoolean.True;
			this.cmbTiempoDependiente.DisplayMember = "Descripcion";
			this.cmbTiempoDependiente.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbTiempoDependiente.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbTiempoDependiente.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbTiempoDependiente.Location = new System.Drawing.Point(904, 192);
			this.cmbTiempoDependiente.Name = "cmbTiempoDependiente";
			this.cmbTiempoDependiente.Size = new System.Drawing.Size(120, 22);
			this.cmbTiempoDependiente.TabIndex = 1046;
			this.cmbTiempoDependiente.ValueMember = "idCre_Tiempo";
			this.cmbTiempoDependiente.Visible = false;
			// 
			// cmbDependienteTipoempresa
			// 
			this.cmbDependienteTipoempresa.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbDependienteTipoempresa.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbDependienteTipoempresa.DataSource = this.ultraDataSource18;
			ultraGridColumn122.Header.VisiblePosition = 0;
			ultraGridColumn122.Hidden = true;
			ultraGridColumn123.Header.VisiblePosition = 1;
			ultraGridColumn123.Width = 150;
			ultraGridBand49.Columns.AddRange(new object[] {
																											ultraGridColumn122,
																											ultraGridColumn123});
			this.cmbDependienteTipoempresa.DisplayLayout.BandsSerializer.Add(ultraGridBand49);
			this.cmbDependienteTipoempresa.DisplayMember = "Descripción";
			this.cmbDependienteTipoempresa.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbDependienteTipoempresa.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbDependienteTipoempresa.Enabled = false;
			this.cmbDependienteTipoempresa.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbDependienteTipoempresa.Location = new System.Drawing.Point(192, 31);
			this.cmbDependienteTipoempresa.Name = "cmbDependienteTipoempresa";
			this.cmbDependienteTipoempresa.Size = new System.Drawing.Size(150, 22);
			this.cmbDependienteTipoempresa.TabIndex = 899;
			this.cmbDependienteTipoempresa.ValueMember = "idTipoEmpresa";
			this.cmbDependienteTipoempresa.KeyDown += new System.Windows.Forms.KeyEventHandler(this.cmbDependienteTipoempresa_KeyDown_1);
			// 
			// ultraDataSource18
			// 
			ultraDataColumn77.DataType = typeof(int);
			this.ultraDataSource18.Band.Columns.AddRange(new object[] {
																																	ultraDataColumn77,
																																	ultraDataColumn78});
			// 
			// cmbDependienteTipoSueldo
			// 
			this.cmbDependienteTipoSueldo.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbDependienteTipoSueldo.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbDependienteTipoSueldo.DataSource = this.ultraDataSource20;
			ultraGridColumn124.Header.VisiblePosition = 0;
			ultraGridColumn124.Hidden = true;
			ultraGridColumn125.Header.VisiblePosition = 1;
			ultraGridColumn125.Width = 150;
			ultraGridBand50.Columns.AddRange(new object[] {
																											ultraGridColumn124,
																											ultraGridColumn125});
			this.cmbDependienteTipoSueldo.DisplayLayout.BandsSerializer.Add(ultraGridBand50);
			this.cmbDependienteTipoSueldo.DisplayMember = "Nombre";
			this.cmbDependienteTipoSueldo.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbDependienteTipoSueldo.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbDependienteTipoSueldo.Enabled = false;
			this.cmbDependienteTipoSueldo.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbDependienteTipoSueldo.Location = new System.Drawing.Point(168, 72);
			this.cmbDependienteTipoSueldo.Name = "cmbDependienteTipoSueldo";
			this.cmbDependienteTipoSueldo.Size = new System.Drawing.Size(150, 22);
			this.cmbDependienteTipoSueldo.TabIndex = 897;
			this.cmbDependienteTipoSueldo.ValueMember = "idTipoSueldo";
			this.cmbDependienteTipoSueldo.KeyDown += new System.Windows.Forms.KeyEventHandler(this.cmbDependienteTipoSueldo_KeyDown_1);
			// 
			// ultraDataSource20
			// 
			ultraDataColumn79.DataType = typeof(int);
			this.ultraDataSource20.Band.Columns.AddRange(new object[] {
																																	ultraDataColumn79,
																																	ultraDataColumn80});
			// 
			// label121
			// 
			this.label121.AutoSize = true;
			this.label121.BackColor = System.Drawing.Color.Transparent;
			this.label121.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label121.Location = new System.Drawing.Point(656, 56);
			this.label121.Name = "label121";
			this.label121.Size = new System.Drawing.Size(54, 17);
			this.label121.TabIndex = 896;
			this.label121.Text = "Dias Pago";
			this.label121.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// label120
			// 
			this.label120.AutoSize = true;
			this.label120.BackColor = System.Drawing.Color.Transparent;
			this.label120.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label120.Location = new System.Drawing.Point(496, 56);
			this.label120.Name = "label120";
			this.label120.Size = new System.Drawing.Size(33, 17);
			this.label120.TabIndex = 895;
			this.label120.Text = "Cargo";
			this.label120.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.label120.Click += new System.EventHandler(this.label120_Click);
			// 
			// label100
			// 
			this.label100.AutoSize = true;
			this.label100.BackColor = System.Drawing.Color.Transparent;
			this.label100.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label100.Location = new System.Drawing.Point(328, 56);
			this.label100.Name = "label100";
			this.label100.Size = new System.Drawing.Size(76, 17);
			this.label100.TabIndex = 894;
			this.label100.Text = "Departamento";
			this.label100.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// label99
			// 
			this.label99.AutoSize = true;
			this.label99.BackColor = System.Drawing.Color.Transparent;
			this.label99.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label99.Location = new System.Drawing.Point(187, 7);
			this.label99.Name = "label99";
			this.label99.Size = new System.Drawing.Size(69, 17);
			this.label99.TabIndex = 893;
			this.label99.Text = "TipoEmpresa";
			this.label99.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// label98
			// 
			this.label98.AutoSize = true;
			this.label98.BackColor = System.Drawing.Color.Transparent;
			this.label98.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label98.Location = new System.Drawing.Point(168, 56);
			this.label98.Name = "label98";
			this.label98.Size = new System.Drawing.Size(62, 17);
			this.label98.TabIndex = 891;
			this.label98.Text = "Tipo sueldo";
			this.label98.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.label98.Click += new System.EventHandler(this.label98_Click);
			// 
			// label97
			// 
			this.label97.AutoSize = true;
			this.label97.BackColor = System.Drawing.Color.Transparent;
			this.label97.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label97.Location = new System.Drawing.Point(8, 56);
			this.label97.Name = "label97";
			this.label97.Size = new System.Drawing.Size(72, 17);
			this.label97.TabIndex = 890;
			this.label97.Text = "Tipo Contrato";
			this.label97.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// cmbDependienteCargo
			// 
			this.cmbDependienteCargo.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbDependienteCargo.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbDependienteCargo.DataSource = this.ultraDataSource21;
			ultraGridColumn126.Header.VisiblePosition = 0;
			ultraGridColumn126.Hidden = true;
			ultraGridColumn127.Header.Caption = "Provincia";
			ultraGridColumn127.Header.VisiblePosition = 1;
			ultraGridColumn127.Width = 150;
			ultraGridBand51.Columns.AddRange(new object[] {
																											ultraGridColumn126,
																											ultraGridColumn127});
			this.cmbDependienteCargo.DisplayLayout.BandsSerializer.Add(ultraGridBand51);
			this.cmbDependienteCargo.DisplayMember = "Nombre";
			this.cmbDependienteCargo.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbDependienteCargo.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbDependienteCargo.Enabled = false;
			this.cmbDependienteCargo.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbDependienteCargo.Location = new System.Drawing.Point(496, 72);
			this.cmbDependienteCargo.Name = "cmbDependienteCargo";
			this.cmbDependienteCargo.Size = new System.Drawing.Size(150, 22);
			this.cmbDependienteCargo.TabIndex = 889;
			this.cmbDependienteCargo.ValueMember = "idCargo";
			this.cmbDependienteCargo.KeyDown += new System.Windows.Forms.KeyEventHandler(this.cmbDependienteCargo_KeyDown);
			// 
			// ultraDataSource21
			// 
			ultraDataColumn81.DataType = typeof(int);
			this.ultraDataSource21.Band.Columns.AddRange(new object[] {
																																	ultraDataColumn81,
																																	ultraDataColumn82});
			// 
			// txtDependienteDepartamento
			// 
			appearance91.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtDependienteDepartamento.Appearance = appearance91;
			this.txtDependienteDepartamento.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtDependienteDepartamento.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtDependienteDepartamento.Enabled = false;
			this.txtDependienteDepartamento.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.txtDependienteDepartamento.Location = new System.Drawing.Point(328, 72);
			this.txtDependienteDepartamento.Name = "txtDependienteDepartamento";
			this.txtDependienteDepartamento.Size = new System.Drawing.Size(160, 22);
			this.txtDependienteDepartamento.TabIndex = 888;
			this.txtDependienteDepartamento.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtDependienteDepartamento_KeyDown);
			this.txtDependienteDepartamento.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.txtDependienteDepartamento_KeyPress);
			// 
			// chkAfiliadoIESS
			// 
			appearance92.ForeColorDisabled = System.Drawing.Color.Black;
			this.chkAfiliadoIESS.Appearance = appearance92;
			this.chkAfiliadoIESS.BackColor = System.Drawing.Color.Transparent;
			this.chkAfiliadoIESS.CausesValidation = false;
			this.chkAfiliadoIESS.Enabled = false;
			this.chkAfiliadoIESS.Location = new System.Drawing.Point(760, 72);
			this.chkAfiliadoIESS.Name = "chkAfiliadoIESS";
			this.chkAfiliadoIESS.Size = new System.Drawing.Size(160, 22);
			this.chkAfiliadoIESS.TabIndex = 887;
			this.chkAfiliadoIESS.Text = "Afiliado IESS";
			// 
			// btnTrabajoCelular
			// 
			this.btnTrabajoCelular.Enabled = false;
			this.btnTrabajoCelular.Image = ((System.Drawing.Image)(resources.GetObject("btnTrabajoCelular.Image")));
			this.btnTrabajoCelular.Location = new System.Drawing.Point(344, 192);
			this.btnTrabajoCelular.Name = "btnTrabajoCelular";
			this.btnTrabajoCelular.Size = new System.Drawing.Size(24, 23);
			this.btnTrabajoCelular.TabIndex = 885;
			this.btnTrabajoCelular.Click += new System.EventHandler(this.btnTrabajoCelular_Click);
			// 
			// txtDependienteEXT
			// 
			this.txtDependienteEXT.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtDependienteEXT.Enabled = false;
			this.txtDependienteEXT.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.txtDependienteEXT.Location = new System.Drawing.Point(144, 192);
			this.txtDependienteEXT.MaxLength = 9;
			this.txtDependienteEXT.Name = "txtDependienteEXT";
			this.txtDependienteEXT.Size = new System.Drawing.Size(90, 22);
			this.txtDependienteEXT.TabIndex = 884;
			this.txtDependienteEXT.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtDependienteEXT_KeyDown);
			this.txtDependienteEXT.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.txtDependienteEXT_KeyPress);
			// 
			// label101
			// 
			this.label101.AutoSize = true;
			this.label101.BackColor = System.Drawing.Color.Transparent;
			this.label101.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label101.Location = new System.Drawing.Point(656, 96);
			this.label101.Name = "label101";
			this.label101.Size = new System.Drawing.Size(31, 17);
			this.label101.TabIndex = 879;
			this.label101.Text = "Email";
			this.label101.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.label101.Visible = false;
			// 
			// txtDependienteemail
			// 
			appearance93.BackColorDisabled = System.Drawing.Color.Silver;
			appearance93.BackColorDisabled2 = System.Drawing.Color.Silver;
			appearance93.BorderColor = System.Drawing.Color.Black;
			appearance93.BorderColor3DBase = System.Drawing.Color.Black;
			appearance93.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtDependienteemail.Appearance = appearance93;
			this.txtDependienteemail.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtDependienteemail.Enabled = false;
			this.txtDependienteemail.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.txtDependienteemail.Location = new System.Drawing.Point(656, 112);
			this.txtDependienteemail.Name = "txtDependienteemail";
			this.txtDependienteemail.Size = new System.Drawing.Size(120, 22);
			this.txtDependienteemail.TabIndex = 878;
			this.txtDependienteemail.Visible = false;
			this.txtDependienteemail.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtDependienteemail_KeyDown);
			// 
			// label102
			// 
			this.label102.AutoSize = true;
			this.label102.BackColor = System.Drawing.Color.Transparent;
			this.label102.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label102.Location = new System.Drawing.Point(496, 96);
			this.label102.Name = "label102";
			this.label102.Size = new System.Drawing.Size(34, 17);
			this.label102.TabIndex = 877;
			this.label102.Text = "Barrio";
			this.label102.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// label103
			// 
			this.label103.AutoSize = true;
			this.label103.BackColor = System.Drawing.Color.Transparent;
			this.label103.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label103.Location = new System.Drawing.Point(328, 96);
			this.label103.Name = "label103";
			this.label103.Size = new System.Drawing.Size(52, 17);
			this.label103.TabIndex = 876;
			this.label103.Text = "Parroquia";
			this.label103.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// label104
			// 
			this.label104.AutoSize = true;
			this.label104.BackColor = System.Drawing.Color.Transparent;
			this.label104.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label104.Location = new System.Drawing.Point(168, 96);
			this.label104.Name = "label104";
			this.label104.Size = new System.Drawing.Size(39, 17);
			this.label104.TabIndex = 875;
			this.label104.Text = "Cantón";
			this.label104.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// label105
			// 
			this.label105.AutoSize = true;
			this.label105.BackColor = System.Drawing.Color.Transparent;
			this.label105.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label105.Location = new System.Drawing.Point(8, 96);
			this.label105.Name = "label105";
			this.label105.Size = new System.Drawing.Size(49, 17);
			this.label105.TabIndex = 874;
			this.label105.Text = "Provincia";
			this.label105.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// cmbDependienteBarrio
			// 
			this.cmbDependienteBarrio.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbDependienteBarrio.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbDependienteBarrio.DataSource = this.ultraDataSource4;
			ultraGridColumn128.Header.VisiblePosition = 0;
			ultraGridColumn128.Hidden = true;
			ultraGridColumn129.Header.Caption = "Barrio";
			ultraGridColumn129.Header.VisiblePosition = 1;
			ultraGridColumn129.Width = 150;
			ultraGridBand52.Columns.AddRange(new object[] {
																											ultraGridColumn128,
																											ultraGridColumn129});
			this.cmbDependienteBarrio.DisplayLayout.BandsSerializer.Add(ultraGridBand52);
			this.cmbDependienteBarrio.DisplayLayout.Override.RowSelectors = Infragistics.Win.DefaultableBoolean.True;
			this.cmbDependienteBarrio.DisplayMember = "Nombre";
			this.cmbDependienteBarrio.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbDependienteBarrio.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbDependienteBarrio.Enabled = false;
			this.cmbDependienteBarrio.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbDependienteBarrio.Location = new System.Drawing.Point(496, 112);
			this.cmbDependienteBarrio.Name = "cmbDependienteBarrio";
			this.cmbDependienteBarrio.Size = new System.Drawing.Size(150, 22);
			this.cmbDependienteBarrio.TabIndex = 873;
			this.cmbDependienteBarrio.ValueMember = "idBarrio";
			this.cmbDependienteBarrio.KeyDown += new System.Windows.Forms.KeyEventHandler(this.cmbDependienteBarrio_KeyDown);
			// 
			// cmbDependienteParroquia
			// 
			this.cmbDependienteParroquia.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbDependienteParroquia.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbDependienteParroquia.DataSource = this.ultraDataSource3;
			ultraGridColumn130.Header.VisiblePosition = 0;
			ultraGridColumn130.Hidden = true;
			ultraGridColumn131.Header.Caption = "Provincia";
			ultraGridColumn131.Header.VisiblePosition = 1;
			ultraGridColumn131.Width = 150;
			ultraGridBand53.Columns.AddRange(new object[] {
																											ultraGridColumn130,
																											ultraGridColumn131});
			this.cmbDependienteParroquia.DisplayLayout.BandsSerializer.Add(ultraGridBand53);
			this.cmbDependienteParroquia.DisplayMember = "Nombre";
			this.cmbDependienteParroquia.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbDependienteParroquia.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbDependienteParroquia.Enabled = false;
			this.cmbDependienteParroquia.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbDependienteParroquia.Location = new System.Drawing.Point(328, 112);
			this.cmbDependienteParroquia.Name = "cmbDependienteParroquia";
			this.cmbDependienteParroquia.Size = new System.Drawing.Size(150, 22);
			this.cmbDependienteParroquia.TabIndex = 872;
			this.cmbDependienteParroquia.ValueMember = "idParroquia";
			this.cmbDependienteParroquia.KeyDown += new System.Windows.Forms.KeyEventHandler(this.cmbDependienteParroquia_KeyDown);
			this.cmbDependienteParroquia.ValueChanged += new System.EventHandler(this.cmbDependienteParroquia_ValueChanged);
			// 
			// cmbDependienteCanton
			// 
			this.cmbDependienteCanton.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbDependienteCanton.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbDependienteCanton.DataSource = this.ultraDataSource2;
			ultraGridColumn132.Header.VisiblePosition = 0;
			ultraGridColumn132.Hidden = true;
			ultraGridColumn133.Header.Caption = "Canton";
			ultraGridColumn133.Header.VisiblePosition = 1;
			ultraGridColumn133.Width = 150;
			ultraGridBand54.Columns.AddRange(new object[] {
																											ultraGridColumn132,
																											ultraGridColumn133});
			this.cmbDependienteCanton.DisplayLayout.BandsSerializer.Add(ultraGridBand54);
			this.cmbDependienteCanton.DisplayLayout.Override.RowSelectors = Infragistics.Win.DefaultableBoolean.True;
			this.cmbDependienteCanton.DisplayMember = "Nombre";
			this.cmbDependienteCanton.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbDependienteCanton.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbDependienteCanton.Enabled = false;
			this.cmbDependienteCanton.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbDependienteCanton.Location = new System.Drawing.Point(168, 112);
			this.cmbDependienteCanton.Name = "cmbDependienteCanton";
			this.cmbDependienteCanton.Size = new System.Drawing.Size(150, 22);
			this.cmbDependienteCanton.TabIndex = 871;
			this.cmbDependienteCanton.ValueMember = "idCanton";
			this.cmbDependienteCanton.KeyDown += new System.Windows.Forms.KeyEventHandler(this.cmbDependienteCanton_KeyDown);
			this.cmbDependienteCanton.ValueChanged += new System.EventHandler(this.cmbDependienteCanton_ValueChanged);
			// 
			// cmbDependienteProvincia
			// 
			this.cmbDependienteProvincia.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbDependienteProvincia.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbDependienteProvincia.DataSource = this.ultraDataSource1;
			ultraGridColumn134.Header.VisiblePosition = 0;
			ultraGridColumn134.Hidden = true;
			ultraGridColumn134.Width = 50;
			ultraGridColumn135.Header.VisiblePosition = 1;
			ultraGridColumn135.Width = 150;
			ultraGridBand55.Columns.AddRange(new object[] {
																											ultraGridColumn134,
																											ultraGridColumn135});
			this.cmbDependienteProvincia.DisplayLayout.BandsSerializer.Add(ultraGridBand55);
			this.cmbDependienteProvincia.DisplayMember = "Nombre";
			this.cmbDependienteProvincia.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbDependienteProvincia.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbDependienteProvincia.Enabled = false;
			this.cmbDependienteProvincia.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbDependienteProvincia.Location = new System.Drawing.Point(8, 112);
			this.cmbDependienteProvincia.Name = "cmbDependienteProvincia";
			this.cmbDependienteProvincia.Size = new System.Drawing.Size(150, 22);
			this.cmbDependienteProvincia.TabIndex = 870;
			this.cmbDependienteProvincia.ValueMember = "idProvincia";
			this.cmbDependienteProvincia.KeyDown += new System.Windows.Forms.KeyEventHandler(this.cmbDependienteProvincia_KeyDown);
			this.cmbDependienteProvincia.ValueChanged += new System.EventHandler(this.cmbDependienteProvincia_ValueChanged);
			// 
			// label106
			// 
			this.label106.AutoSize = true;
			this.label106.BackColor = System.Drawing.Color.Transparent;
			this.label106.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label106.Location = new System.Drawing.Point(760, 7);
			this.label106.Name = "label106";
			this.label106.Size = new System.Drawing.Size(41, 17);
			this.label106.TabIndex = 869;
			this.label106.Text = "Pasivos";
			this.label106.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.label106.Visible = false;
			// 
			// label107
			// 
			this.label107.AutoSize = true;
			this.label107.BackColor = System.Drawing.Color.Transparent;
			this.label107.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label107.Location = new System.Drawing.Point(672, 7);
			this.label107.Name = "label107";
			this.label107.Size = new System.Drawing.Size(39, 17);
			this.label107.TabIndex = 868;
			this.label107.Text = "Activos";
			this.label107.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.label107.Visible = false;
			// 
			// label108
			// 
			this.label108.AutoSize = true;
			this.label108.BackColor = System.Drawing.Color.Transparent;
			this.label108.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label108.Location = new System.Drawing.Point(576, 7);
			this.label108.Name = "label108";
			this.label108.Size = new System.Drawing.Size(38, 17);
			this.label108.TabIndex = 867;
			this.label108.Text = "Gastos";
			this.label108.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// label109
			// 
			this.label109.AutoSize = true;
			this.label109.BackColor = System.Drawing.Color.Transparent;
			this.label109.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label109.Location = new System.Drawing.Point(472, 7);
			this.label109.Name = "label109";
			this.label109.Size = new System.Drawing.Size(47, 17);
			this.label109.TabIndex = 866;
			this.label109.Text = "Ingresos";
			this.label109.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// txtDependientePasivos
			// 
			this.txtDependientePasivos.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtDependientePasivos.Enabled = false;
			this.txtDependientePasivos.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.txtDependientePasivos.Location = new System.Drawing.Point(760, 31);
			this.txtDependientePasivos.MaxValue = 100000;
			this.txtDependientePasivos.MinValue = 0;
			this.txtDependientePasivos.Name = "txtDependientePasivos";
			this.txtDependientePasivos.PromptChar = ' ';
			this.txtDependientePasivos.Size = new System.Drawing.Size(90, 22);
			this.txtDependientePasivos.SpinButtonDisplayStyle = Infragistics.Win.ButtonDisplayStyle.Always;
			this.txtDependientePasivos.TabIndex = 865;
			this.txtDependientePasivos.Value = 100;
			this.txtDependientePasivos.Visible = false;
			this.txtDependientePasivos.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtDependientePasivos_KeyDown);
			this.txtDependientePasivos.Click += new System.EventHandler(this.txtDependientePasivos_Click);
			// 
			// txtDependienteActivos
			// 
			this.txtDependienteActivos.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtDependienteActivos.Enabled = false;
			this.txtDependienteActivos.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.txtDependienteActivos.Location = new System.Drawing.Point(664, 31);
			this.txtDependienteActivos.MaxValue = 1000000;
			this.txtDependienteActivos.MinValue = 0;
			this.txtDependienteActivos.Name = "txtDependienteActivos";
			this.txtDependienteActivos.PromptChar = ' ';
			this.txtDependienteActivos.Size = new System.Drawing.Size(90, 22);
			this.txtDependienteActivos.SpinButtonDisplayStyle = Infragistics.Win.ButtonDisplayStyle.Always;
			this.txtDependienteActivos.TabIndex = 864;
			this.txtDependienteActivos.Value = 450;
			this.txtDependienteActivos.Visible = false;
			this.txtDependienteActivos.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtDependienteActivos_KeyDown);
			this.txtDependienteActivos.Click += new System.EventHandler(this.txtDependienteActivos_Click);
			// 
			// txtDependienteGastos
			// 
			this.txtDependienteGastos.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtDependienteGastos.Enabled = false;
			this.txtDependienteGastos.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.txtDependienteGastos.Location = new System.Drawing.Point(568, 31);
			this.txtDependienteGastos.MaxValue = 1000000;
			this.txtDependienteGastos.MinValue = 0;
			this.txtDependienteGastos.Name = "txtDependienteGastos";
			this.txtDependienteGastos.PromptChar = ' ';
			this.txtDependienteGastos.Size = new System.Drawing.Size(90, 22);
			this.txtDependienteGastos.SpinButtonDisplayStyle = Infragistics.Win.ButtonDisplayStyle.Always;
			this.txtDependienteGastos.TabIndex = 863;
			this.txtDependienteGastos.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtDependienteGastos_KeyDown);
			this.txtDependienteGastos.Click += new System.EventHandler(this.txtDependienteGastos_Click);
			// 
			// txtDependienteIngresos
			// 
			this.txtDependienteIngresos.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtDependienteIngresos.Enabled = false;
			this.txtDependienteIngresos.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.txtDependienteIngresos.Location = new System.Drawing.Point(472, 31);
			this.txtDependienteIngresos.MaxValue = 1000000;
			this.txtDependienteIngresos.MinValue = 0;
			this.txtDependienteIngresos.Name = "txtDependienteIngresos";
			this.txtDependienteIngresos.PromptChar = ' ';
			this.txtDependienteIngresos.Size = new System.Drawing.Size(90, 22);
			this.txtDependienteIngresos.SpinButtonDisplayStyle = Infragistics.Win.ButtonDisplayStyle.Always;
			this.txtDependienteIngresos.TabIndex = 862;
			this.txtDependienteIngresos.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtDependienteIngresos_KeyDown);
			this.txtDependienteIngresos.Click += new System.EventHandler(this.txtDependienteIngresos_Click);
			// 
			// label110
			// 
			this.label110.AutoSize = true;
			this.label110.BackColor = System.Drawing.Color.Transparent;
			this.label110.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label110.Location = new System.Drawing.Point(8, 224);
			this.label110.Name = "label110";
			this.label110.Size = new System.Drawing.Size(108, 17);
			this.label110.TabIndex = 861;
			this.label110.Text = "Referencia Ubicacion";
			this.label110.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// txtDependienteReferencia
			// 
			this.txtDependienteReferencia.AcceptsReturn = true;
			appearance94.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtDependienteReferencia.Appearance = appearance94;
			this.txtDependienteReferencia.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtDependienteReferencia.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtDependienteReferencia.Enabled = false;
			this.txtDependienteReferencia.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.txtDependienteReferencia.Location = new System.Drawing.Point(8, 248);
			this.txtDependienteReferencia.Multiline = true;
			this.txtDependienteReferencia.Name = "txtDependienteReferencia";
			this.txtDependienteReferencia.Size = new System.Drawing.Size(300, 66);
			this.txtDependienteReferencia.TabIndex = 860;
			this.txtDependienteReferencia.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.txtDependienteReferencia_KeyPress);
			// 
			// btnTrabajoTelefono
			// 
			this.btnTrabajoTelefono.Enabled = false;
			this.btnTrabajoTelefono.Image = ((System.Drawing.Image)(resources.GetObject("btnTrabajoTelefono.Image")));
			this.btnTrabajoTelefono.Location = new System.Drawing.Point(112, 192);
			this.btnTrabajoTelefono.Name = "btnTrabajoTelefono";
			this.btnTrabajoTelefono.Size = new System.Drawing.Size(24, 23);
			this.btnTrabajoTelefono.TabIndex = 858;
			this.btnTrabajoTelefono.Click += new System.EventHandler(this.btnTrabajoTelefono_Click);
			// 
			// label111
			// 
			this.label111.AutoSize = true;
			this.label111.BackColor = System.Drawing.Color.Transparent;
			this.label111.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label111.Location = new System.Drawing.Point(240, 176);
			this.label111.Name = "label111";
			this.label111.Size = new System.Drawing.Size(38, 17);
			this.label111.TabIndex = 857;
			this.label111.Text = "Celular";
			this.label111.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// label112
			// 
			this.label112.AutoSize = true;
			this.label112.BackColor = System.Drawing.Color.Transparent;
			this.label112.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label112.Location = new System.Drawing.Point(144, 176);
			this.label112.Name = "label112";
			this.label112.Size = new System.Drawing.Size(20, 17);
			this.label112.TabIndex = 856;
			this.label112.Text = "Ext";
			this.label112.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// label113
			// 
			this.label113.AutoSize = true;
			this.label113.BackColor = System.Drawing.Color.Transparent;
			this.label113.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label113.Location = new System.Drawing.Point(8, 176);
			this.label113.Name = "label113";
			this.label113.Size = new System.Drawing.Size(57, 17);
			this.label113.TabIndex = 855;
			this.label113.Text = "Teléfono 1";
			this.label113.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// label114
			// 
			this.label114.AutoSize = true;
			this.label114.BackColor = System.Drawing.Color.Transparent;
			this.label114.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label114.Location = new System.Drawing.Point(496, 136);
			this.label114.Name = "label114";
			this.label114.Size = new System.Drawing.Size(87, 17);
			this.label114.TabIndex = 854;
			this.label114.Text = "Calle Secundaria";
			this.label114.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// label115
			// 
			this.label115.AutoSize = true;
			this.label115.BackColor = System.Drawing.Color.Transparent;
			this.label115.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label115.Location = new System.Drawing.Point(312, 136);
			this.label115.Name = "label115";
			this.label115.Size = new System.Drawing.Size(71, 17);
			this.label115.TabIndex = 853;
			this.label115.Text = "Número Casa";
			this.label115.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// label116
			// 
			this.label116.AutoSize = true;
			this.label116.BackColor = System.Drawing.Color.Transparent;
			this.label116.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label116.Location = new System.Drawing.Point(8, 136);
			this.label116.Name = "label116";
			this.label116.Size = new System.Drawing.Size(73, 17);
			this.label116.TabIndex = 852;
			this.label116.Text = "Calle Principal";
			this.label116.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// txtDependienteCelular
			// 
			this.txtDependienteCelular.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtDependienteCelular.Enabled = false;
			this.txtDependienteCelular.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.txtDependienteCelular.Location = new System.Drawing.Point(240, 192);
			this.txtDependienteCelular.MaxLength = 10;
			this.txtDependienteCelular.Name = "txtDependienteCelular";
			this.txtDependienteCelular.Size = new System.Drawing.Size(100, 22);
			this.txtDependienteCelular.TabIndex = 851;
			this.txtDependienteCelular.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtDependienteCelular_KeyDown);
			this.txtDependienteCelular.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.txtDependienteCelular_KeyPress);
			// 
			// txtDependienteTelefono
			// 
			this.txtDependienteTelefono.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtDependienteTelefono.Enabled = false;
			this.txtDependienteTelefono.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.txtDependienteTelefono.Location = new System.Drawing.Point(8, 192);
			this.txtDependienteTelefono.MaxLength = 9;
			this.txtDependienteTelefono.Name = "txtDependienteTelefono";
			this.txtDependienteTelefono.Size = new System.Drawing.Size(100, 22);
			this.txtDependienteTelefono.TabIndex = 849;
			this.txtDependienteTelefono.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtDependienteTelefono_KeyDown);
			this.txtDependienteTelefono.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.txtDependienteTelefono_KeyPress);
			// 
			// txtDependienteCalleSecundaria
			// 
			appearance95.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtDependienteCalleSecundaria.Appearance = appearance95;
			this.txtDependienteCalleSecundaria.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtDependienteCalleSecundaria.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtDependienteCalleSecundaria.Enabled = false;
			this.txtDependienteCalleSecundaria.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.txtDependienteCalleSecundaria.Location = new System.Drawing.Point(496, 152);
			this.txtDependienteCalleSecundaria.Name = "txtDependienteCalleSecundaria";
			this.txtDependienteCalleSecundaria.Size = new System.Drawing.Size(300, 22);
			this.txtDependienteCalleSecundaria.TabIndex = 848;
			this.txtDependienteCalleSecundaria.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtDependienteCalleSecundaria_KeyDown);
			this.txtDependienteCalleSecundaria.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.txtDependienteCalleSecundaria_KeyPress);
			// 
			// txtDependienteNumeroCasa
			// 
			appearance96.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtDependienteNumeroCasa.Appearance = appearance96;
			this.txtDependienteNumeroCasa.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtDependienteNumeroCasa.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtDependienteNumeroCasa.Enabled = false;
			this.txtDependienteNumeroCasa.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.txtDependienteNumeroCasa.Location = new System.Drawing.Point(312, 152);
			this.txtDependienteNumeroCasa.Name = "txtDependienteNumeroCasa";
			this.txtDependienteNumeroCasa.Size = new System.Drawing.Size(180, 22);
			this.txtDependienteNumeroCasa.TabIndex = 847;
			this.txtDependienteNumeroCasa.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtDependienteNumeroCasa_KeyDown);
			this.txtDependienteNumeroCasa.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.txtDependienteNumeroCasa_KeyPress);
			// 
			// txtDependienteCallePrinciapl
			// 
			appearance97.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtDependienteCallePrinciapl.Appearance = appearance97;
			this.txtDependienteCallePrinciapl.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtDependienteCallePrinciapl.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtDependienteCallePrinciapl.Enabled = false;
			this.txtDependienteCallePrinciapl.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.txtDependienteCallePrinciapl.Location = new System.Drawing.Point(8, 152);
			this.txtDependienteCallePrinciapl.Name = "txtDependienteCallePrinciapl";
			this.txtDependienteCallePrinciapl.Size = new System.Drawing.Size(300, 22);
			this.txtDependienteCallePrinciapl.TabIndex = 846;
			this.txtDependienteCallePrinciapl.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtDependienteCallePrinciapl_KeyDown);
			this.txtDependienteCallePrinciapl.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.txtDependienteCallePrinciapl_KeyPress);
			// 
			// txtDependienteDiasPago
			// 
			this.txtDependienteDiasPago.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtDependienteDiasPago.Enabled = false;
			this.txtDependienteDiasPago.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.txtDependienteDiasPago.Location = new System.Drawing.Point(656, 72);
			this.txtDependienteDiasPago.MaxValue = 31;
			this.txtDependienteDiasPago.MinValue = 0;
			this.txtDependienteDiasPago.Name = "txtDependienteDiasPago";
			this.txtDependienteDiasPago.PromptChar = ' ';
			this.txtDependienteDiasPago.Size = new System.Drawing.Size(90, 22);
			this.txtDependienteDiasPago.SpinButtonDisplayStyle = Infragistics.Win.ButtonDisplayStyle.Always;
			this.txtDependienteDiasPago.TabIndex = 839;
			this.txtDependienteDiasPago.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtDependienteDiasPago_KeyDown);
			this.txtDependienteDiasPago.Click += new System.EventHandler(this.txtDependienteDiasPago_Click);
			// 
			// label117
			// 
			this.label117.AutoSize = true;
			this.label117.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label117.Location = new System.Drawing.Point(7, 7);
			this.label117.Name = "label117";
			this.label117.Size = new System.Drawing.Size(47, 17);
			this.label117.TabIndex = 824;
			this.label117.Text = "Empresa";
			this.label117.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// txtDependienteEmpresa
			// 
			appearance98.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtDependienteEmpresa.Appearance = appearance98;
			this.txtDependienteEmpresa.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtDependienteEmpresa.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtDependienteEmpresa.Enabled = false;
			this.txtDependienteEmpresa.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.txtDependienteEmpresa.Location = new System.Drawing.Point(7, 31);
			this.txtDependienteEmpresa.Name = "txtDependienteEmpresa";
			this.txtDependienteEmpresa.Size = new System.Drawing.Size(180, 22);
			this.txtDependienteEmpresa.TabIndex = 828;
			this.txtDependienteEmpresa.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtDependienteEmpresa_KeyDown);
			this.txtDependienteEmpresa.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.txtDependienteEmpresa_KeyPress);
			// 
			// label118
			// 
			this.label118.AutoSize = true;
			this.label118.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label118.Location = new System.Drawing.Point(920, 118);
			this.label118.Name = "label118";
			this.label118.Size = new System.Drawing.Size(23, 17);
			this.label118.TabIndex = 818;
			this.label118.Text = "Ruc";
			this.label118.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.label118.Visible = false;
			// 
			// txtDependienteRuc
			// 
			appearance99.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtDependienteRuc.Appearance = appearance99;
			this.txtDependienteRuc.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtDependienteRuc.Enabled = false;
			this.txtDependienteRuc.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.txtDependienteRuc.Location = new System.Drawing.Point(920, 139);
			this.txtDependienteRuc.MaxLength = 13;
			this.txtDependienteRuc.Name = "txtDependienteRuc";
			this.txtDependienteRuc.Size = new System.Drawing.Size(127, 22);
			this.txtDependienteRuc.TabIndex = 819;
			this.txtDependienteRuc.Visible = false;
			this.txtDependienteRuc.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtDependienteRuc_KeyDown);
			this.txtDependienteRuc.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.txtDependienteRuc_KeyPress);
			// 
			// cmbDependienteTipocontrato
			// 
			this.cmbDependienteTipocontrato.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbDependienteTipocontrato.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbDependienteTipocontrato.DataSource = this.ultraDataSource19;
			ultraGridColumn136.Header.VisiblePosition = 0;
			ultraGridColumn136.Hidden = true;
			ultraGridColumn137.Header.VisiblePosition = 1;
			ultraGridColumn137.Width = 150;
			ultraGridBand56.Columns.AddRange(new object[] {
																											ultraGridColumn136,
																											ultraGridColumn137});
			this.cmbDependienteTipocontrato.DisplayLayout.BandsSerializer.Add(ultraGridBand56);
			this.cmbDependienteTipocontrato.DisplayMember = "Nombre";
			this.cmbDependienteTipocontrato.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbDependienteTipocontrato.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbDependienteTipocontrato.Enabled = false;
			this.cmbDependienteTipocontrato.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbDependienteTipocontrato.Location = new System.Drawing.Point(8, 72);
			this.cmbDependienteTipocontrato.Name = "cmbDependienteTipocontrato";
			this.cmbDependienteTipocontrato.Size = new System.Drawing.Size(150, 22);
			this.cmbDependienteTipocontrato.TabIndex = 810;
			this.cmbDependienteTipocontrato.ValueMember = "idTipoContrato";
			this.cmbDependienteTipocontrato.KeyDown += new System.Windows.Forms.KeyEventHandler(this.cmbDependienteTipocontrato_KeyDown);
			// 
			// ultraDataSource19
			// 
			ultraDataColumn83.DataType = typeof(int);
			this.ultraDataSource19.Band.Columns.AddRange(new object[] {
																																	ultraDataColumn83,
																																	ultraDataColumn84});
			// 
			// label119
			// 
			this.label119.AutoSize = true;
			this.label119.BackColor = System.Drawing.Color.Transparent;
			this.label119.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label119.Location = new System.Drawing.Point(347, 7);
			this.label119.Name = "label119";
			this.label119.Size = new System.Drawing.Size(89, 17);
			this.label119.TabIndex = 786;
			this.label119.Text = "Fecha de ingreso";
			this.label119.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// dtDependienteFechaIngreso
			// 
			this.dtDependienteFechaIngreso.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			dateButton5.Caption = "Today";
			this.dtDependienteFechaIngreso.DateButtons.Add(dateButton5);
			this.dtDependienteFechaIngreso.Enabled = false;
			this.dtDependienteFechaIngreso.Format = "dd/MM/yyyy";
			this.dtDependienteFechaIngreso.Location = new System.Drawing.Point(352, 32);
			this.dtDependienteFechaIngreso.Name = "dtDependienteFechaIngreso";
			this.dtDependienteFechaIngreso.NonAutoSizeHeight = 23;
			this.dtDependienteFechaIngreso.Size = new System.Drawing.Size(104, 21);
			this.dtDependienteFechaIngreso.TabIndex = 779;
			this.dtDependienteFechaIngreso.Value = ((object)(resources.GetObject("dtDependienteFechaIngreso.Value")));
			this.dtDependienteFechaIngreso.KeyDown += new System.Windows.Forms.KeyEventHandler(this.dtDependienteFechaIngreso_KeyDown);
			// 
			// ultraTabPageControl3
			// 
			this.ultraTabPageControl3.Controls.Add(this.groupBox8);
			this.ultraTabPageControl3.Controls.Add(this.groupBox7);
			this.ultraTabPageControl3.Location = new System.Drawing.Point(-10000, -10000);
			this.ultraTabPageControl3.Name = "ultraTabPageControl3";
			this.ultraTabPageControl3.Size = new System.Drawing.Size(1116, 366);
			// 
			// groupBox8
			// 
			this.groupBox8.BackColor = System.Drawing.SystemColors.Control;
			this.groupBox8.Controls.Add(this.utgHistorialCredito);
			this.groupBox8.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.groupBox8.Location = new System.Drawing.Point(8, 208);
			this.groupBox8.Name = "groupBox8";
			this.groupBox8.Size = new System.Drawing.Size(1104, 160);
			this.groupBox8.TabIndex = 200;
			this.groupBox8.TabStop = false;
			this.groupBox8.Text = "Historial Crediticio";
			// 
			// utgHistorialCredito
			// 
			this.utgHistorialCredito.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.utgHistorialCredito.Cursor = System.Windows.Forms.Cursors.Default;
			this.utgHistorialCredito.DataSource = this.ultraDataSource41;
			appearance100.BackColor = System.Drawing.Color.White;
			this.utgHistorialCredito.DisplayLayout.Appearance = appearance100;
			appearance101.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn138.CellAppearance = appearance101;
			ultraGridColumn138.Header.VisiblePosition = 0;
			appearance102.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn139.CellAppearance = appearance102;
			appearance103.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn139.FormulaErrorAppearance = appearance103;
			ultraGridColumn139.Header.VisiblePosition = 1;
			ultraGridColumn140.Header.Caption = "Cliente";
			ultraGridColumn140.Header.VisiblePosition = 3;
			ultraGridColumn141.Header.VisiblePosition = 2;
			ultraGridColumn141.Width = 80;
			ultraGridColumn142.Header.Caption = "E. Financiera";
			ultraGridColumn142.Header.VisiblePosition = 4;
			appearance104.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn143.CellAppearance = appearance104;
			appearance105.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn143.FormulaErrorAppearance = appearance105;
			ultraGridColumn143.Header.VisiblePosition = 5;
			ultraGridColumn143.Width = 45;
			ultraGridColumn144.Header.VisiblePosition = 6;
			ultraGridColumn145.Header.VisiblePosition = 7;
			appearance106.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn146.CellAppearance = appearance106;
			appearance107.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn146.FormulaErrorAppearance = appearance107;
			ultraGridColumn146.Header.Caption = "Cuotas";
			ultraGridColumn146.Header.VisiblePosition = 8;
			appearance108.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn147.CellAppearance = appearance108;
			ultraGridColumn147.Format = "#,##0.00";
			appearance109.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn147.FormulaErrorAppearance = appearance109;
			ultraGridColumn147.Header.VisiblePosition = 9;
			ultraGridColumn148.Header.Caption = "F. Pago";
			ultraGridColumn148.Header.VisiblePosition = 10;
			appearance110.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn149.CellAppearance = appearance110;
			ultraGridColumn149.Format = "#,##0.00";
			appearance111.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn149.FormulaErrorAppearance = appearance111;
			ultraGridColumn149.Header.Caption = "Valor Cuotas";
			ultraGridColumn149.Header.VisiblePosition = 11;
			appearance112.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn150.CellAppearance = appearance112;
			appearance113.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn150.FormulaErrorAppearance = appearance113;
			ultraGridColumn150.Header.Caption = "C. Pendientes";
			ultraGridColumn150.Header.VisiblePosition = 12;
			appearance114.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn151.CellAppearance = appearance114;
			appearance115.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn151.FormulaErrorAppearance = appearance115;
			ultraGridColumn151.Header.Caption = "C. Pagadas";
			ultraGridColumn151.Header.VisiblePosition = 13;
			appearance116.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn152.CellAppearance = appearance116;
			ultraGridColumn152.Format = "#,##0.00";
			appearance117.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn152.FormulaErrorAppearance = appearance117;
			ultraGridColumn152.Header.Caption = "Saldo C";
			ultraGridColumn152.Header.VisiblePosition = 14;
			appearance118.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn153.CellAppearance = appearance118;
			appearance119.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn153.FormulaErrorAppearance = appearance119;
			ultraGridColumn153.Header.Caption = "P. Retraso";
			ultraGridColumn153.Header.VisiblePosition = 15;
			ultraGridColumn154.Header.Caption = "Ult Pago";
			ultraGridColumn154.Header.VisiblePosition = 16;
			ultraGridColumn155.Header.VisiblePosition = 17;
			ultraGridBand57.Columns.AddRange(new object[] {
																											ultraGridColumn138,
																											ultraGridColumn139,
																											ultraGridColumn140,
																											ultraGridColumn141,
																											ultraGridColumn142,
																											ultraGridColumn143,
																											ultraGridColumn144,
																											ultraGridColumn145,
																											ultraGridColumn146,
																											ultraGridColumn147,
																											ultraGridColumn148,
																											ultraGridColumn149,
																											ultraGridColumn150,
																											ultraGridColumn151,
																											ultraGridColumn152,
																											ultraGridColumn153,
																											ultraGridColumn154,
																											ultraGridColumn155});
			this.utgHistorialCredito.DisplayLayout.BandsSerializer.Add(ultraGridBand57);
			this.utgHistorialCredito.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.False;
			this.utgHistorialCredito.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.False;
			appearance120.BackColor = System.Drawing.Color.Transparent;
			this.utgHistorialCredito.DisplayLayout.Override.CardAreaAppearance = appearance120;
			appearance121.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance121.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance121.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance121.FontData.BoldAsString = "True";
			appearance121.FontData.Name = "Arial";
			appearance121.FontData.SizeInPoints = 10F;
			appearance121.ForeColor = System.Drawing.Color.White;
			appearance121.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.utgHistorialCredito.DisplayLayout.Override.HeaderAppearance = appearance121;
			this.utgHistorialCredito.DisplayLayout.Override.HeaderClickAction = Infragistics.Win.UltraWinGrid.HeaderClickAction.SortMulti;
			appearance122.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance122.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance122.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.utgHistorialCredito.DisplayLayout.Override.RowSelectorAppearance = appearance122;
			appearance123.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance123.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance123.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.utgHistorialCredito.DisplayLayout.Override.SelectedRowAppearance = appearance123;
			this.utgHistorialCredito.DisplayLayout.Override.SelectTypeCell = Infragistics.Win.UltraWinGrid.SelectType.None;
			this.utgHistorialCredito.DisplayLayout.Override.SelectTypeCol = Infragistics.Win.UltraWinGrid.SelectType.None;
			this.utgHistorialCredito.DisplayLayout.Override.SelectTypeRow = Infragistics.Win.UltraWinGrid.SelectType.Single;
			this.utgHistorialCredito.DisplayLayout.ScrollBounds = Infragistics.Win.UltraWinGrid.ScrollBounds.ScrollToFill;
			this.utgHistorialCredito.DisplayLayout.ScrollStyle = Infragistics.Win.UltraWinGrid.ScrollStyle.Immediate;
			this.utgHistorialCredito.DisplayLayout.TabNavigation = Infragistics.Win.UltraWinGrid.TabNavigation.NextControl;
			this.utgHistorialCredito.DisplayLayout.ViewStyle = Infragistics.Win.UltraWinGrid.ViewStyle.SingleBand;
			this.utgHistorialCredito.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.utgHistorialCredito.Location = new System.Drawing.Point(8, 24);
			this.utgHistorialCredito.Name = "utgHistorialCredito";
			this.utgHistorialCredito.RightToLeft = System.Windows.Forms.RightToLeft.No;
			this.utgHistorialCredito.Size = new System.Drawing.Size(1088, 120);
			this.utgHistorialCredito.TabIndex = 788;
			// 
			// ultraDataSource41
			// 
			ultraDataColumn88.DataType = typeof(System.DateTime);
			ultraDataColumn90.DataType = typeof(int);
			ultraDataColumn93.DataType = typeof(int);
			ultraDataColumn94.DataType = typeof(System.Decimal);
			ultraDataColumn95.DataType = typeof(System.DateTime);
			ultraDataColumn96.DataType = typeof(System.Decimal);
			ultraDataColumn97.DataType = typeof(int);
			ultraDataColumn98.DataType = typeof(int);
			ultraDataColumn99.DataType = typeof(System.Decimal);
			ultraDataColumn100.DataType = typeof(int);
			ultraDataColumn101.DataType = typeof(System.DateTime);
			this.ultraDataSource41.Band.Columns.AddRange(new object[] {
																																	ultraDataColumn85,
																																	ultraDataColumn86,
																																	ultraDataColumn87,
																																	ultraDataColumn88,
																																	ultraDataColumn89,
																																	ultraDataColumn90,
																																	ultraDataColumn91,
																																	ultraDataColumn92,
																																	ultraDataColumn93,
																																	ultraDataColumn94,
																																	ultraDataColumn95,
																																	ultraDataColumn96,
																																	ultraDataColumn97,
																																	ultraDataColumn98,
																																	ultraDataColumn99,
																																	ultraDataColumn100,
																																	ultraDataColumn101,
																																	ultraDataColumn102});
			// 
			// groupBox7
			// 
			this.groupBox7.BackColor = System.Drawing.SystemColors.Control;
			this.groupBox7.Controls.Add(this.txtCarteraCastigada);
			this.groupBox7.Controls.Add(this.label133);
			this.groupBox7.Controls.Add(this.txtResultado);
			this.groupBox7.Controls.Add(this.label132);
			this.groupBox7.Controls.Add(this.txtSegmentacion);
			this.groupBox7.Controls.Add(this.label131);
			this.groupBox7.Controls.Add(this.txtRangoIngresos);
			this.groupBox7.Controls.Add(this.txtTotalVencido);
			this.groupBox7.Controls.Add(this.txtCapacidadPago);
			this.groupBox7.Controls.Add(this.label128);
			this.groupBox7.Controls.Add(this.label129);
			this.groupBox7.Controls.Add(this.label130);
			this.groupBox7.Controls.Add(this.label127);
			this.groupBox7.Controls.Add(this.txtCreditoScoreInclusion);
			this.groupBox7.Controls.Add(this.label11);
			this.groupBox7.Controls.Add(this.txtCreditoScoreSobreEndeudamiento);
			this.groupBox7.Controls.Add(this.lblUsuarioCompleta);
			this.groupBox7.Controls.Add(this.txtCreditoObservaciones);
			this.groupBox7.Controls.Add(this.label126);
			this.groupBox7.Controls.Add(this.label122);
			this.groupBox7.Controls.Add(this.txtCreditoScore);
			this.groupBox7.Controls.Add(this.label134);
			this.groupBox7.Controls.Add(this.label135);
			this.groupBox7.Controls.Add(this.dtCreditoFechaSolicitud);
			this.groupBox7.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.groupBox7.Location = new System.Drawing.Point(8, 8);
			this.groupBox7.Name = "groupBox7";
			this.groupBox7.Size = new System.Drawing.Size(1104, 200);
			this.groupBox7.TabIndex = 199;
			this.groupBox7.TabStop = false;
			this.groupBox7.Text = "Información";
			this.groupBox7.Enter += new System.EventHandler(this.groupBox7_Enter);
			// 
			// txtCarteraCastigada
			// 
			this.txtCarteraCastigada.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			appearance124.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtCarteraCastigada.Appearance = appearance124;
			this.txtCarteraCastigada.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtCarteraCastigada.Enabled = false;
			this.txtCarteraCastigada.Font = new System.Drawing.Font("Tahoma", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.txtCarteraCastigada.FormatString = "#,##0.00";
			this.txtCarteraCastigada.Location = new System.Drawing.Point(456, 64);
			this.txtCarteraCastigada.Name = "txtCarteraCastigada";
			this.txtCarteraCastigada.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtCarteraCastigada.PromptChar = ' ';
			this.txtCarteraCastigada.Size = new System.Drawing.Size(90, 23);
			this.txtCarteraCastigada.TabIndex = 848;
			// 
			// label133
			// 
			this.label133.AutoSize = true;
			this.label133.Location = new System.Drawing.Point(456, 40);
			this.label133.Name = "label133";
			this.label133.Size = new System.Drawing.Size(105, 17);
			this.label133.TabIndex = 847;
			this.label133.Text = "Cartera Castigada";
			this.label133.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// txtResultado
			// 
			appearance125.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtResultado.Appearance = appearance125;
			this.txtResultado.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtResultado.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtResultado.Enabled = false;
			this.txtResultado.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.txtResultado.Location = new System.Drawing.Point(352, 64);
			this.txtResultado.Name = "txtResultado";
			this.txtResultado.Size = new System.Drawing.Size(90, 22);
			this.txtResultado.TabIndex = 846;
			// 
			// label132
			// 
			this.label132.AutoSize = true;
			this.label132.Location = new System.Drawing.Point(352, 40);
			this.label132.Name = "label132";
			this.label132.Size = new System.Drawing.Size(105, 17);
			this.label132.TabIndex = 845;
			this.label132.Text = "Resultado Equifax";
			this.label132.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// txtSegmentacion
			// 
			appearance126.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtSegmentacion.Appearance = appearance126;
			this.txtSegmentacion.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtSegmentacion.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtSegmentacion.Enabled = false;
			this.txtSegmentacion.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.txtSegmentacion.Location = new System.Drawing.Point(568, 64);
			this.txtSegmentacion.Name = "txtSegmentacion";
			this.txtSegmentacion.Size = new System.Drawing.Size(168, 22);
			this.txtSegmentacion.TabIndex = 844;
			// 
			// label131
			// 
			this.label131.AutoSize = true;
			this.label131.Location = new System.Drawing.Point(568, 40);
			this.label131.Name = "label131";
			this.label131.Size = new System.Drawing.Size(128, 17);
			this.label131.TabIndex = 843;
			this.label131.Text = "Segmentación Equifax";
			this.label131.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// txtRangoIngresos
			// 
			appearance127.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtRangoIngresos.Appearance = appearance127;
			this.txtRangoIngresos.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtRangoIngresos.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtRangoIngresos.Enabled = false;
			this.txtRangoIngresos.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.txtRangoIngresos.Location = new System.Drawing.Point(27, 62);
			this.txtRangoIngresos.Name = "txtRangoIngresos";
			this.txtRangoIngresos.Size = new System.Drawing.Size(90, 22);
			this.txtRangoIngresos.TabIndex = 842;
			// 
			// txtTotalVencido
			// 
			this.txtTotalVencido.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			appearance128.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtTotalVencido.Appearance = appearance128;
			this.txtTotalVencido.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtTotalVencido.Enabled = false;
			this.txtTotalVencido.Font = new System.Drawing.Font("Tahoma", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.txtTotalVencido.FormatString = "#,##0.00";
			this.txtTotalVencido.Location = new System.Drawing.Point(240, 62);
			this.txtTotalVencido.Name = "txtTotalVencido";
			this.txtTotalVencido.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtTotalVencido.PromptChar = ' ';
			this.txtTotalVencido.Size = new System.Drawing.Size(90, 23);
			this.txtTotalVencido.TabIndex = 841;
			// 
			// txtCapacidadPago
			// 
			this.txtCapacidadPago.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			appearance129.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtCapacidadPago.Appearance = appearance129;
			this.txtCapacidadPago.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtCapacidadPago.Enabled = false;
			this.txtCapacidadPago.Font = new System.Drawing.Font("Tahoma", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.txtCapacidadPago.FormatString = "#,##0.00";
			this.txtCapacidadPago.Location = new System.Drawing.Point(133, 62);
			this.txtCapacidadPago.Name = "txtCapacidadPago";
			this.txtCapacidadPago.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtCapacidadPago.PromptChar = ' ';
			this.txtCapacidadPago.Size = new System.Drawing.Size(90, 23);
			this.txtCapacidadPago.TabIndex = 840;
			// 
			// label128
			// 
			this.label128.AutoSize = true;
			this.label128.Location = new System.Drawing.Point(133, 35);
			this.label128.Name = "label128";
			this.label128.Size = new System.Drawing.Size(96, 17);
			this.label128.TabIndex = 839;
			this.label128.Text = "Capacidad Pago";
			this.label128.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// label129
			// 
			this.label129.AutoSize = true;
			this.label129.Location = new System.Drawing.Point(240, 35);
			this.label129.Name = "label129";
			this.label129.Size = new System.Drawing.Size(80, 17);
			this.label129.TabIndex = 838;
			this.label129.Text = "Total Vencido";
			this.label129.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// label130
			// 
			this.label130.AutoSize = true;
			this.label130.Location = new System.Drawing.Point(27, 35);
			this.label130.Name = "label130";
			this.label130.Size = new System.Drawing.Size(92, 17);
			this.label130.TabIndex = 837;
			this.label130.Text = "Rango Ingresos";
			this.label130.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// label127
			// 
			this.label127.AutoSize = true;
			this.label127.Location = new System.Drawing.Point(136, 104);
			this.label127.Name = "label127";
			this.label127.Size = new System.Drawing.Size(89, 17);
			this.label127.TabIndex = 836;
			this.label127.Text = "Score Inclusion";
			this.label127.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.label127.Click += new System.EventHandler(this.label127_Click);
			// 
			// txtCreditoScoreInclusion
			// 
			this.txtCreditoScoreInclusion.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtCreditoScoreInclusion.Enabled = false;
			this.txtCreditoScoreInclusion.Font = new System.Drawing.Font("Tahoma", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.txtCreditoScoreInclusion.Location = new System.Drawing.Point(136, 120);
			this.txtCreditoScoreInclusion.MinValue = 0;
			this.txtCreditoScoreInclusion.Name = "txtCreditoScoreInclusion";
			this.txtCreditoScoreInclusion.PromptChar = ' ';
			this.txtCreditoScoreInclusion.Size = new System.Drawing.Size(90, 23);
			this.txtCreditoScoreInclusion.SpinButtonDisplayStyle = Infragistics.Win.ButtonDisplayStyle.Always;
			this.txtCreditoScoreInclusion.TabIndex = 835;
			// 
			// label11
			// 
			this.label11.AutoSize = true;
			this.label11.Location = new System.Drawing.Point(240, 104);
			this.label11.Name = "label11";
			this.label11.Size = new System.Drawing.Size(180, 17);
			this.label11.TabIndex = 834;
			this.label11.Text = "Score de Sobre Endeudamiento";
			this.label11.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// txtCreditoScoreSobreEndeudamiento
			// 
			this.txtCreditoScoreSobreEndeudamiento.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtCreditoScoreSobreEndeudamiento.Enabled = false;
			this.txtCreditoScoreSobreEndeudamiento.Font = new System.Drawing.Font("Tahoma", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.txtCreditoScoreSobreEndeudamiento.Location = new System.Drawing.Point(240, 120);
			this.txtCreditoScoreSobreEndeudamiento.MinValue = 0;
			this.txtCreditoScoreSobreEndeudamiento.Name = "txtCreditoScoreSobreEndeudamiento";
			this.txtCreditoScoreSobreEndeudamiento.PromptChar = ' ';
			this.txtCreditoScoreSobreEndeudamiento.Size = new System.Drawing.Size(90, 23);
			this.txtCreditoScoreSobreEndeudamiento.SpinButtonDisplayStyle = Infragistics.Win.ButtonDisplayStyle.Always;
			this.txtCreditoScoreSobreEndeudamiento.TabIndex = 833;
			this.txtCreditoScoreSobreEndeudamiento.ValueChanged += new System.EventHandler(this.txtCreditoScoreSobreEndeudamiento_ValueChanged);
			// 
			// lblUsuarioCompleta
			// 
			this.lblUsuarioCompleta.AutoSize = true;
			this.lblUsuarioCompleta.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.lblUsuarioCompleta.ForeColor = System.Drawing.Color.Firebrick;
			this.lblUsuarioCompleta.Location = new System.Drawing.Point(453, 159);
			this.lblUsuarioCompleta.Name = "lblUsuarioCompleta";
			this.lblUsuarioCompleta.Size = new System.Drawing.Size(0, 19);
			this.lblUsuarioCompleta.TabIndex = 832;
			this.lblUsuarioCompleta.TextAlign = System.Drawing.ContentAlignment.BottomRight;
			// 
			// txtCreditoObservaciones
			// 
			this.txtCreditoObservaciones.AcceptsReturn = true;
			appearance130.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtCreditoObservaciones.Appearance = appearance130;
			this.txtCreditoObservaciones.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtCreditoObservaciones.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtCreditoObservaciones.Enabled = false;
			this.txtCreditoObservaciones.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.txtCreditoObservaciones.Location = new System.Drawing.Point(760, 49);
			this.txtCreditoObservaciones.Multiline = true;
			this.txtCreditoObservaciones.Name = "txtCreditoObservaciones";
			this.txtCreditoObservaciones.Size = new System.Drawing.Size(333, 124);
			this.txtCreditoObservaciones.TabIndex = 831;
			this.txtCreditoObservaciones.Visible = false;
			this.txtCreditoObservaciones.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.txtCreditoObservaciones_KeyPress);
			// 
			// label126
			// 
			this.label126.AutoSize = true;
			this.label126.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label126.Location = new System.Drawing.Point(753, 21);
			this.label126.Name = "label126";
			this.label126.Size = new System.Drawing.Size(76, 17);
			this.label126.TabIndex = 795;
			this.label126.Text = "Observaciones";
			this.label126.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// label122
			// 
			this.label122.AutoSize = true;
			this.label122.Location = new System.Drawing.Point(24, 104);
			this.label122.Name = "label122";
			this.label122.Size = new System.Drawing.Size(37, 17);
			this.label122.TabIndex = 789;
			this.label122.Text = "Score";
			this.label122.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// txtCreditoScore
			// 
			this.txtCreditoScore.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtCreditoScore.Enabled = false;
			this.txtCreditoScore.Font = new System.Drawing.Font("Tahoma", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.txtCreditoScore.Location = new System.Drawing.Point(24, 120);
			this.txtCreditoScore.MinValue = 0;
			this.txtCreditoScore.Name = "txtCreditoScore";
			this.txtCreditoScore.PromptChar = ' ';
			this.txtCreditoScore.Size = new System.Drawing.Size(90, 23);
			this.txtCreditoScore.SpinButtonDisplayStyle = Infragistics.Win.ButtonDisplayStyle.Always;
			this.txtCreditoScore.TabIndex = 787;
			this.txtCreditoScore.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtCreditoScore_KeyDown);
			this.txtCreditoScore.Click += new System.EventHandler(this.txtCreditoScore_Click);
			// 
			// label134
			// 
			this.label134.AutoSize = true;
			this.label134.BackColor = System.Drawing.Color.Transparent;
			this.label134.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label134.Location = new System.Drawing.Point(27, 159);
			this.label134.Name = "label134";
			this.label134.Size = new System.Drawing.Size(149, 17);
			this.label134.TabIndex = 786;
			this.label134.Text = "Fecha Solicitud Completa ";
			this.label134.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// label135
			// 
			this.label135.AutoSize = true;
			this.label135.BackColor = System.Drawing.Color.Transparent;
			this.label135.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label135.Location = new System.Drawing.Point(893, 21);
			this.label135.Name = "label135";
			this.label135.Size = new System.Drawing.Size(92, 17);
			this.label135.TabIndex = 785;
			this.label135.Text = "Usuario Completa";
			this.label135.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// dtCreditoFechaSolicitud
			// 
			this.dtCreditoFechaSolicitud.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			dateButton6.Caption = "Today";
			this.dtCreditoFechaSolicitud.DateButtons.Add(dateButton6);
			this.dtCreditoFechaSolicitud.Enabled = false;
			this.dtCreditoFechaSolicitud.Format = "dd/MM/yyyy";
			this.dtCreditoFechaSolicitud.Location = new System.Drawing.Point(240, 159);
			this.dtCreditoFechaSolicitud.Name = "dtCreditoFechaSolicitud";
			this.dtCreditoFechaSolicitud.NonAutoSizeHeight = 23;
			this.dtCreditoFechaSolicitud.Size = new System.Drawing.Size(90, 21);
			this.dtCreditoFechaSolicitud.TabIndex = 779;
			this.dtCreditoFechaSolicitud.Value = ((object)(resources.GetObject("dtCreditoFechaSolicitud.Value")));
			this.dtCreditoFechaSolicitud.BeforeDropDown += new System.ComponentModel.CancelEventHandler(this.ultraCalendarCombo1_BeforeDropDown);
			// 
			// ultraTabPageControl6
			// 
			this.ultraTabPageControl6.Controls.Add(this.cmbdetallecliente);
			this.ultraTabPageControl6.Controls.Add(this.cmbidTipoCliente);
			this.ultraTabPageControl6.Controls.Add(this.label137);
			this.ultraTabPageControl6.Controls.Add(this.groupBox9);
			this.ultraTabPageControl6.Controls.Add(this.ultraCombo2);
			this.ultraTabPageControl6.Controls.Add(this.ultraCombo3);
			this.ultraTabPageControl6.Controls.Add(this.label139);
			this.ultraTabPageControl6.Controls.Add(this.label140);
			this.ultraTabPageControl6.Location = new System.Drawing.Point(-10000, -10000);
			this.ultraTabPageControl6.Name = "ultraTabPageControl6";
			this.ultraTabPageControl6.Size = new System.Drawing.Size(1116, 366);
			// 
			// cmbdetallecliente
			// 
			this.cmbdetallecliente.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbdetallecliente.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbdetallecliente.DataSource = this.ultraDataSource44;
			ultraGridColumn156.Header.VisiblePosition = 0;
			ultraGridColumn156.Hidden = true;
			ultraGridColumn156.Width = 8;
			ultraGridColumn157.Header.VisiblePosition = 1;
			ultraGridColumn157.Width = 150;
			ultraGridBand58.Columns.AddRange(new object[] {
																											ultraGridColumn156,
																											ultraGridColumn157});
			this.cmbdetallecliente.DisplayLayout.BandsSerializer.Add(ultraGridBand58);
			this.cmbdetallecliente.DisplayLayout.Override.RowSelectors = Infragistics.Win.DefaultableBoolean.True;
			this.cmbdetallecliente.DisplayMember = "Nombre";
			this.cmbdetallecliente.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbdetallecliente.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbdetallecliente.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbdetallecliente.Location = new System.Drawing.Point(408, 32);
			this.cmbdetallecliente.Name = "cmbdetallecliente";
			this.cmbdetallecliente.Size = new System.Drawing.Size(168, 22);
			this.cmbdetallecliente.TabIndex = 793;
			this.cmbdetallecliente.ValueMember = "idDetalleTipoCliente";
			// 
			// ultraDataSource44
			// 
			ultraDataColumn103.DataType = typeof(int);
			this.ultraDataSource44.Band.Columns.AddRange(new object[] {
																																	ultraDataColumn103,
																																	ultraDataColumn104});
			// 
			// cmbidTipoCliente
			// 
			this.cmbidTipoCliente.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbidTipoCliente.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbidTipoCliente.DataSource = this.ultraDataSource43;
			ultraGridColumn158.Header.VisiblePosition = 0;
			ultraGridColumn158.Hidden = true;
			ultraGridColumn158.Width = 8;
			ultraGridColumn159.Header.VisiblePosition = 1;
			ultraGridColumn159.Width = 150;
			ultraGridBand59.Columns.AddRange(new object[] {
																											ultraGridColumn158,
																											ultraGridColumn159});
			this.cmbidTipoCliente.DisplayLayout.BandsSerializer.Add(ultraGridBand59);
			this.cmbidTipoCliente.DisplayLayout.Override.RowSelectors = Infragistics.Win.DefaultableBoolean.True;
			this.cmbidTipoCliente.DisplayMember = "Nombre";
			this.cmbidTipoCliente.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbidTipoCliente.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbidTipoCliente.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbidTipoCliente.Location = new System.Drawing.Point(128, 32);
			this.cmbidTipoCliente.Name = "cmbidTipoCliente";
			this.cmbidTipoCliente.Size = new System.Drawing.Size(160, 22);
			this.cmbidTipoCliente.TabIndex = 792;
			this.cmbidTipoCliente.ValueMember = "idTipoCliente";
			// 
			// ultraDataSource43
			// 
			ultraDataColumn105.DataType = typeof(int);
			this.ultraDataSource43.Band.Columns.AddRange(new object[] {
																																	ultraDataColumn105,
																																	ultraDataColumn106});
			// 
			// label137
			// 
			this.label137.AutoSize = true;
			this.label137.Location = new System.Drawing.Point(24, 35);
			this.label137.Name = "label137";
			this.label137.Size = new System.Drawing.Size(26, 17);
			this.label137.TabIndex = 788;
			this.label137.Text = "Tipo";
			this.label137.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// groupBox9
			// 
			this.groupBox9.BackColor = System.Drawing.SystemColors.Control;
			this.groupBox9.Controls.Add(this.label142);
			this.groupBox9.Controls.Add(this.txtObservaciones);
			this.groupBox9.Controls.Add(this.label124);
			this.groupBox9.Controls.Add(this.label123);
			this.groupBox9.Controls.Add(this.cmbCreditoCalificacion);
			this.groupBox9.Controls.Add(this.txtCreditoCuotaAsiganada);
			this.groupBox9.Controls.Add(this.btnImprimeVerificacion);
			this.groupBox9.Controls.Add(this.label143);
			this.groupBox9.Controls.Add(this.cmdEstadoSolicitud);
			this.groupBox9.Controls.Add(this.gdObservaciones);
			this.groupBox9.Controls.Add(this.ChkOmite);
			this.groupBox9.Controls.Add(this.chkNuovo);
			this.groupBox9.Controls.Add(this.label141);
			this.groupBox9.Controls.Add(this.txtCupo);
			this.groupBox9.Controls.Add(this.label125);
			this.groupBox9.Controls.Add(this.cmbCreditoEstado);
			this.groupBox9.Controls.Add(this.label138);
			this.groupBox9.Controls.Add(this.cmbSituacionLaboral2);
			this.groupBox9.Controls.Add(this.label136);
			this.groupBox9.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.groupBox9.Location = new System.Drawing.Point(0, 0);
			this.groupBox9.Name = "groupBox9";
			this.groupBox9.Size = new System.Drawing.Size(1120, 368);
			this.groupBox9.TabIndex = 1040;
			this.groupBox9.TabStop = false;
			this.groupBox9.Text = "Información";
			this.groupBox9.Enter += new System.EventHandler(this.groupBox9_Enter);
			// 
			// label142
			// 
			this.label142.AutoSize = true;
			this.label142.Location = new System.Drawing.Point(24, 160);
			this.label142.Name = "label142";
			this.label142.Size = new System.Drawing.Size(87, 17);
			this.label142.TabIndex = 1065;
			this.label142.Text = "Observaciones";
			this.label142.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// txtObservaciones
			// 
			appearance131.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtObservaciones.Appearance = appearance131;
			this.txtObservaciones.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtObservaciones.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtObservaciones.Location = new System.Drawing.Point(120, 144);
			this.txtObservaciones.Multiline = true;
			this.txtObservaciones.Name = "txtObservaciones";
			this.txtObservaciones.Size = new System.Drawing.Size(456, 64);
			this.txtObservaciones.TabIndex = 1064;
			// 
			// label124
			// 
			this.label124.AutoSize = true;
			this.label124.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label124.Location = new System.Drawing.Point(592, 72);
			this.label124.Name = "label124";
			this.label124.Size = new System.Drawing.Size(68, 17);
			this.label124.TabIndex = 1063;
			this.label124.Text = "Calificación";
			this.label124.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// label123
			// 
			this.label123.AutoSize = true;
			this.label123.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label123.Location = new System.Drawing.Point(592, 32);
			this.label123.Name = "label123";
			this.label123.Size = new System.Drawing.Size(92, 17);
			this.label123.TabIndex = 1062;
			this.label123.Text = "Cuota Asignada";
			this.label123.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// cmbCreditoCalificacion
			// 
			this.cmbCreditoCalificacion.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbCreditoCalificacion.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbCreditoCalificacion.DataSource = this.ultraDataSource22;
			ultraGridColumn160.Header.VisiblePosition = 0;
			ultraGridColumn160.Hidden = true;
			ultraGridColumn161.Header.Caption = "Calificacion";
			ultraGridColumn161.Header.VisiblePosition = 1;
			ultraGridColumn161.Width = 150;
			ultraGridBand60.Columns.AddRange(new object[] {
																											ultraGridColumn160,
																											ultraGridColumn161});
			this.cmbCreditoCalificacion.DisplayLayout.BandsSerializer.Add(ultraGridBand60);
			this.cmbCreditoCalificacion.DisplayLayout.Override.RowSelectors = Infragistics.Win.DefaultableBoolean.True;
			this.cmbCreditoCalificacion.DisplayMember = "Codigo";
			this.cmbCreditoCalificacion.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbCreditoCalificacion.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbCreditoCalificacion.Enabled = false;
			this.cmbCreditoCalificacion.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbCreditoCalificacion.Location = new System.Drawing.Point(728, 69);
			this.cmbCreditoCalificacion.Name = "cmbCreditoCalificacion";
			this.cmbCreditoCalificacion.Size = new System.Drawing.Size(90, 22);
			this.cmbCreditoCalificacion.TabIndex = 1061;
			this.cmbCreditoCalificacion.ValueMember = "idTipoCalificacion";
			// 
			// ultraDataSource22
			// 
			ultraDataColumn107.DataType = typeof(int);
			this.ultraDataSource22.Band.Columns.AddRange(new object[] {
																																	ultraDataColumn107,
																																	ultraDataColumn108});
			// 
			// txtCreditoCuotaAsiganada
			// 
			this.txtCreditoCuotaAsiganada.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtCreditoCuotaAsiganada.Enabled = false;
			this.txtCreditoCuotaAsiganada.Font = new System.Drawing.Font("Tahoma", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.txtCreditoCuotaAsiganada.Location = new System.Drawing.Point(728, 29);
			this.txtCreditoCuotaAsiganada.MinValue = 0;
			this.txtCreditoCuotaAsiganada.Name = "txtCreditoCuotaAsiganada";
			this.txtCreditoCuotaAsiganada.PromptChar = ' ';
			this.txtCreditoCuotaAsiganada.Size = new System.Drawing.Size(90, 23);
			this.txtCreditoCuotaAsiganada.SpinButtonDisplayStyle = Infragistics.Win.ButtonDisplayStyle.Always;
			this.txtCreditoCuotaAsiganada.TabIndex = 1060;
			// 
			// btnImprimeVerificacion
			// 
			this.btnImprimeVerificacion.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btnImprimeVerificacion.CausesValidation = false;
			this.btnImprimeVerificacion.Image = ((System.Drawing.Image)(resources.GetObject("btnImprimeVerificacion.Image")));
			this.btnImprimeVerificacion.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnImprimeVerificacion.Location = new System.Drawing.Point(1008, 32);
			this.btnImprimeVerificacion.Name = "btnImprimeVerificacion";
			this.btnImprimeVerificacion.Size = new System.Drawing.Size(77, 23);
			this.btnImprimeVerificacion.TabIndex = 1059;
			this.btnImprimeVerificacion.Text = "&Imprimir";
			this.btnImprimeVerificacion.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnImprimeVerificacion.Click += new System.EventHandler(this.btnImprimeVerificacion_Click);
			// 
			// label143
			// 
			this.label143.AutoSize = true;
			this.label143.Location = new System.Drawing.Point(320, 104);
			this.label143.Name = "label143";
			this.label143.Size = new System.Drawing.Size(93, 17);
			this.label143.TabIndex = 1058;
			this.label143.Text = "Estado Solicitud";
			this.label143.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// cmdEstadoSolicitud
			// 
			this.cmdEstadoSolicitud.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			this.cmdEstadoSolicitud.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmdEstadoSolicitud.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmdEstadoSolicitud.DataSource = this.ultraDataSource47;
			ultraGridColumn162.Header.VisiblePosition = 0;
			ultraGridColumn162.Hidden = true;
			ultraGridColumn163.Header.Caption = "Estado";
			ultraGridColumn163.Header.VisiblePosition = 1;
			ultraGridColumn163.Width = 150;
			ultraGridBand61.Columns.AddRange(new object[] {
																											ultraGridColumn162,
																											ultraGridColumn163});
			this.cmdEstadoSolicitud.DisplayLayout.BandsSerializer.Add(ultraGridBand61);
			this.cmdEstadoSolicitud.DisplayLayout.Override.RowSelectors = Infragistics.Win.DefaultableBoolean.True;
			this.cmdEstadoSolicitud.DisplayMember = "Nombre";
			this.cmdEstadoSolicitud.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmdEstadoSolicitud.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmdEstadoSolicitud.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmdEstadoSolicitud.Location = new System.Drawing.Point(424, 104);
			this.cmdEstadoSolicitud.Name = "cmdEstadoSolicitud";
			this.cmdEstadoSolicitud.Size = new System.Drawing.Size(152, 22);
			this.cmdEstadoSolicitud.TabIndex = 1057;
			this.cmdEstadoSolicitud.ValueMember = "idNegadoPendiente";
			// 
			// ultraDataSource47
			// 
			ultraDataColumn109.DataType = typeof(int);
			this.ultraDataSource47.Band.Columns.AddRange(new object[] {
																																	ultraDataColumn109,
																																	ultraDataColumn110});
			// 
			// gdObservaciones
			// 
			this.gdObservaciones.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.gdObservaciones.Cursor = System.Windows.Forms.Cursors.Default;
			this.gdObservaciones.DataSource = this.ultraDataSource46;
			appearance132.BackColor = System.Drawing.Color.White;
			this.gdObservaciones.DisplayLayout.Appearance = appearance132;
			ultraGridColumn164.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn164.Header.VisiblePosition = 0;
			ultraGridColumn165.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn165.Header.VisiblePosition = 1;
			ultraGridColumn166.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn166.Header.VisiblePosition = 3;
			ultraGridColumn167.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn167.Header.Caption = "Almacen";
			ultraGridColumn167.Header.VisiblePosition = 2;
			ultraGridColumn167.Width = 140;
			ultraGridColumn168.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn168.Header.VisiblePosition = 4;
			ultraGridColumn169.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn169.Header.VisiblePosition = 5;
			ultraGridColumn170.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn170.Header.VisiblePosition = 6;
			ultraGridColumn171.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn171.Header.VisiblePosition = 7;
			ultraGridColumn172.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn172.Header.VisiblePosition = 8;
			ultraGridColumn173.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn173.Header.VisiblePosition = 9;
			ultraGridBand62.Columns.AddRange(new object[] {
																											ultraGridColumn164,
																											ultraGridColumn165,
																											ultraGridColumn166,
																											ultraGridColumn167,
																											ultraGridColumn168,
																											ultraGridColumn169,
																											ultraGridColumn170,
																											ultraGridColumn171,
																											ultraGridColumn172,
																											ultraGridColumn173});
			this.gdObservaciones.DisplayLayout.BandsSerializer.Add(ultraGridBand62);
			appearance133.ForeColor = System.Drawing.Color.Black;
			appearance133.ForeColorDisabled = System.Drawing.Color.Black;
			this.gdObservaciones.DisplayLayout.Override.ActiveRowAppearance = appearance133;
			this.gdObservaciones.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.False;
			appearance134.BackColor = System.Drawing.Color.Transparent;
			this.gdObservaciones.DisplayLayout.Override.CardAreaAppearance = appearance134;
			appearance135.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance135.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance135.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance135.FontData.BoldAsString = "True";
			appearance135.FontData.Name = "Arial";
			appearance135.FontData.SizeInPoints = 8F;
			appearance135.ForeColor = System.Drawing.Color.White;
			appearance135.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.gdObservaciones.DisplayLayout.Override.HeaderAppearance = appearance135;
			this.gdObservaciones.DisplayLayout.Override.HeaderClickAction = Infragistics.Win.UltraWinGrid.HeaderClickAction.SortSingle;
			appearance136.ForeColorDisabled = System.Drawing.Color.Black;
			this.gdObservaciones.DisplayLayout.Override.RowAppearance = appearance136;
			appearance137.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance137.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance137.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.gdObservaciones.DisplayLayout.Override.RowSelectorAppearance = appearance137;
			appearance138.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance138.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance138.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.gdObservaciones.DisplayLayout.Override.SelectedRowAppearance = appearance138;
			this.gdObservaciones.DisplayLayout.Override.SelectTypeRow = Infragistics.Win.UltraWinGrid.SelectType.Extended;
			this.gdObservaciones.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.gdObservaciones.Location = new System.Drawing.Point(0, 224);
			this.gdObservaciones.Name = "gdObservaciones";
			this.gdObservaciones.Size = new System.Drawing.Size(1104, 152);
			this.gdObservaciones.TabIndex = 1056;
			// 
			// ultraDataSource46
			// 
			ultraDataColumn113.DataType = typeof(System.DateTime);
			ultraDataColumn118.DataType = typeof(int);
			ultraDataColumn119.DataType = typeof(System.Decimal);
			this.ultraDataSource46.Band.Columns.AddRange(new object[] {
																																	ultraDataColumn111,
																																	ultraDataColumn112,
																																	ultraDataColumn113,
																																	ultraDataColumn114,
																																	ultraDataColumn115,
																																	ultraDataColumn116,
																																	ultraDataColumn117,
																																	ultraDataColumn118,
																																	ultraDataColumn119,
																																	ultraDataColumn120});
			// 
			// ChkOmite
			// 
			appearance139.ForeColorDisabled = System.Drawing.Color.Black;
			this.ChkOmite.Appearance = appearance139;
			this.ChkOmite.BackColor = System.Drawing.Color.Transparent;
			this.ChkOmite.Enabled = false;
			this.ChkOmite.Location = new System.Drawing.Point(840, 32);
			this.ChkOmite.Name = "ChkOmite";
			this.ChkOmite.Size = new System.Drawing.Size(128, 22);
			this.ChkOmite.TabIndex = 1055;
			this.ChkOmite.Text = "No aplica Nuovo";
			this.ChkOmite.Visible = false;
			// 
			// chkNuovo
			// 
			appearance140.ForeColorDisabled = System.Drawing.Color.Black;
			this.chkNuovo.Appearance = appearance140;
			this.chkNuovo.BackColor = System.Drawing.Color.Transparent;
			this.chkNuovo.Enabled = false;
			this.chkNuovo.Location = new System.Drawing.Point(840, 72);
			this.chkNuovo.Name = "chkNuovo";
			this.chkNuovo.Size = new System.Drawing.Size(68, 22);
			this.chkNuovo.TabIndex = 1054;
			this.chkNuovo.Text = "Nuovo";
			this.chkNuovo.Visible = false;
			// 
			// label141
			// 
			this.label141.AutoSize = true;
			this.label141.Location = new System.Drawing.Point(320, 72);
			this.label141.Name = "label141";
			this.label141.Size = new System.Drawing.Size(34, 17);
			this.label141.TabIndex = 1050;
			this.label141.Text = "Cupo";
			this.label141.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// txtCupo
			// 
			appearance141.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtCupo.Appearance = appearance141;
			this.txtCupo.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtCupo.Enabled = false;
			this.txtCupo.Location = new System.Drawing.Point(408, 69);
			this.txtCupo.MinValue = 0;
			this.txtCupo.Name = "txtCupo";
			this.txtCupo.NumericType = Infragistics.Win.UltraWinEditors.NumericType.Double;
			this.txtCupo.PromptChar = ' ';
			this.txtCupo.Size = new System.Drawing.Size(168, 23);
			this.txtCupo.SpinButtonDisplayStyle = Infragistics.Win.ButtonDisplayStyle.Always;
			this.txtCupo.TabIndex = 1051;
			// 
			// label125
			// 
			this.label125.AutoSize = true;
			this.label125.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label125.Location = new System.Drawing.Point(24, 72);
			this.label125.Name = "label125";
			this.label125.Size = new System.Drawing.Size(38, 17);
			this.label125.TabIndex = 1049;
			this.label125.Text = "Estado";
			this.label125.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// cmbCreditoEstado
			// 
			this.cmbCreditoEstado.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbCreditoEstado.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbCreditoEstado.DataSource = this.ultraDataSource23;
			ultraGridColumn174.Header.VisiblePosition = 0;
			ultraGridColumn174.Hidden = true;
			ultraGridColumn175.Header.VisiblePosition = 1;
			ultraGridColumn175.Width = 150;
			ultraGridBand63.Columns.AddRange(new object[] {
																											ultraGridColumn174,
																											ultraGridColumn175});
			this.cmbCreditoEstado.DisplayLayout.BandsSerializer.Add(ultraGridBand63);
			this.cmbCreditoEstado.DisplayLayout.Override.RowSelectors = Infragistics.Win.DefaultableBoolean.True;
			this.cmbCreditoEstado.DisplayMember = "Estado";
			this.cmbCreditoEstado.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbCreditoEstado.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbCreditoEstado.Enabled = false;
			this.cmbCreditoEstado.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbCreditoEstado.Location = new System.Drawing.Point(128, 69);
			this.cmbCreditoEstado.Name = "cmbCreditoEstado";
			this.cmbCreditoEstado.Size = new System.Drawing.Size(160, 22);
			this.cmbCreditoEstado.TabIndex = 1048;
			this.cmbCreditoEstado.ValueMember = "idEstado";
			this.cmbCreditoEstado.ValueChanged += new System.EventHandler(this.cmbCreditoEstado_ValueChanged);
			// 
			// ultraDataSource23
			// 
			ultraDataColumn121.DataType = typeof(int);
			this.ultraDataSource23.Band.Columns.AddRange(new object[] {
																																	ultraDataColumn121,
																																	ultraDataColumn122});
			// 
			// label138
			// 
			this.label138.AutoSize = true;
			this.label138.Location = new System.Drawing.Point(24, 104);
			this.label138.Name = "label138";
			this.label138.Size = new System.Drawing.Size(74, 17);
			this.label138.TabIndex = 1037;
			this.label138.Text = "Tipo Trabajo";
			this.label138.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// cmbSituacionLaboral2
			// 
			this.cmbSituacionLaboral2.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbSituacionLaboral2.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbSituacionLaboral2.DataSource = this.ultraDataSource45;
			ultraGridColumn176.Header.VisiblePosition = 0;
			ultraGridColumn176.Hidden = true;
			ultraGridColumn176.Width = 8;
			ultraGridColumn177.Header.VisiblePosition = 1;
			ultraGridColumn177.Width = 150;
			ultraGridBand64.Columns.AddRange(new object[] {
																											ultraGridColumn176,
																											ultraGridColumn177});
			this.cmbSituacionLaboral2.DisplayLayout.BandsSerializer.Add(ultraGridBand64);
			this.cmbSituacionLaboral2.DisplayLayout.Override.RowSelectors = Infragistics.Win.DefaultableBoolean.True;
			this.cmbSituacionLaboral2.DisplayMember = "Descripcion";
			this.cmbSituacionLaboral2.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbSituacionLaboral2.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbSituacionLaboral2.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbSituacionLaboral2.Location = new System.Drawing.Point(128, 104);
			this.cmbSituacionLaboral2.Name = "cmbSituacionLaboral2";
			this.cmbSituacionLaboral2.Size = new System.Drawing.Size(160, 22);
			this.cmbSituacionLaboral2.TabIndex = 1039;
			this.cmbSituacionLaboral2.ValueMember = "idSituacionLaboral";
			// 
			// ultraDataSource45
			// 
			ultraDataColumn123.DataType = typeof(int);
			this.ultraDataSource45.Band.Columns.AddRange(new object[] {
																																	ultraDataColumn123,
																																	ultraDataColumn124});
			// 
			// label136
			// 
			this.label136.AutoSize = true;
			this.label136.Location = new System.Drawing.Point(320, 32);
			this.label136.Name = "label136";
			this.label136.Size = new System.Drawing.Size(71, 17);
			this.label136.TabIndex = 790;
			this.label136.Text = "Tipo Cliente";
			this.label136.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// ultraCombo2
			// 
			this.ultraCombo2.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.ultraCombo2.Cursor = System.Windows.Forms.Cursors.Default;
			this.ultraCombo2.DisplayLayout.Override.RowSelectors = Infragistics.Win.DefaultableBoolean.True;
			this.ultraCombo2.DisplayMember = "Nombre";
			this.ultraCombo2.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.ultraCombo2.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.ultraCombo2.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.ultraCombo2.Location = new System.Drawing.Point(328, 24);
			this.ultraCombo2.Name = "ultraCombo2";
			this.ultraCombo2.Size = new System.Drawing.Size(168, 22);
			this.ultraCombo2.TabIndex = 793;
			this.ultraCombo2.ValueMember = "idDetalleTipoCliente";
			// 
			// ultraCombo3
			// 
			this.ultraCombo3.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.ultraCombo3.Cursor = System.Windows.Forms.Cursors.Default;
			this.ultraCombo3.DisplayLayout.Override.RowSelectors = Infragistics.Win.DefaultableBoolean.True;
			this.ultraCombo3.DisplayMember = "Nombre";
			this.ultraCombo3.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.ultraCombo3.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.ultraCombo3.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.ultraCombo3.Location = new System.Drawing.Point(64, 24);
			this.ultraCombo3.Name = "ultraCombo3";
			this.ultraCombo3.Size = new System.Drawing.Size(180, 22);
			this.ultraCombo3.TabIndex = 792;
			this.ultraCombo3.ValueMember = "idTipoCliente";
			// 
			// label139
			// 
			this.label139.AutoSize = true;
			this.label139.Location = new System.Drawing.Point(256, 24);
			this.label139.Name = "label139";
			this.label139.Size = new System.Drawing.Size(63, 17);
			this.label139.TabIndex = 790;
			this.label139.Text = "Tipo Cliente";
			this.label139.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// label140
			// 
			this.label140.AutoSize = true;
			this.label140.Location = new System.Drawing.Point(24, 24);
			this.label140.Name = "label140";
			this.label140.Size = new System.Drawing.Size(26, 17);
			this.label140.TabIndex = 788;
			this.label140.Text = "Tipo";
			this.label140.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// ultraTabPageControl4
			// 
			this.ultraTabPageControl4.Controls.Add(this.ultraGrid1);
			this.ultraTabPageControl4.Location = new System.Drawing.Point(-10000, -10000);
			this.ultraTabPageControl4.Name = "ultraTabPageControl4";
			this.ultraTabPageControl4.Size = new System.Drawing.Size(1116, 366);
			// 
			// ultraGrid1
			// 
			this.ultraGrid1.Cursor = System.Windows.Forms.Cursors.Default;
			this.ultraGrid1.DataSource = this.ultraDataSource42;
			appearance142.BackColor = System.Drawing.Color.White;
			appearance142.ForegroundAlpha = Infragistics.Win.Alpha.UseAlphaLevel;
			this.ultraGrid1.DisplayLayout.Appearance = appearance142;
			ultraGridColumn178.CellActivation = Infragistics.Win.UltraWinGrid.Activation.NoEdit;
			ultraGridColumn178.Header.Caption = "Institución";
			ultraGridColumn178.Header.VisiblePosition = 0;
			ultraGridColumn178.Width = 300;
			ultraGridColumn179.CellActivation = Infragistics.Win.UltraWinGrid.Activation.NoEdit;
			ultraGridColumn179.Header.Caption = "Emisor";
			ultraGridColumn179.Header.VisiblePosition = 1;
			ultraGridColumn179.Width = 200;
			ultraGridColumn180.CellActivation = Infragistics.Win.UltraWinGrid.Activation.NoEdit;
			appearance143.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn180.CellAppearance = appearance143;
			ultraGridColumn180.Header.Caption = "Antigüedad";
			ultraGridColumn180.Header.VisiblePosition = 2;
			ultraGridColumn180.Width = 200;
			ultraGridColumn181.CellActivation = Infragistics.Win.UltraWinGrid.Activation.NoEdit;
			appearance144.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn181.CellAppearance = appearance144;
			ultraGridColumn181.Format = "#,##0.00";
			appearance145.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn181.FormulaErrorAppearance = appearance145;
			ultraGridColumn181.Header.Caption = "Cupo";
			ultraGridColumn181.Header.VisiblePosition = 3;
			ultraGridColumn181.Width = 200;
			ultraGridBand65.Columns.AddRange(new object[] {
																											ultraGridColumn178,
																											ultraGridColumn179,
																											ultraGridColumn180,
																											ultraGridColumn181});
			this.ultraGrid1.DisplayLayout.BandsSerializer.Add(ultraGridBand65);
			this.ultraGrid1.DisplayLayout.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.TemplateOnBottom;
			this.ultraGrid1.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.True;
			this.ultraGrid1.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.True;
			appearance146.BackColor = System.Drawing.Color.Transparent;
			this.ultraGrid1.DisplayLayout.Override.CardAreaAppearance = appearance146;
			appearance147.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance147.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance147.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance147.FontData.BoldAsString = "True";
			appearance147.FontData.Name = "Arial";
			appearance147.FontData.SizeInPoints = 8.5F;
			appearance147.ForeColor = System.Drawing.Color.White;
			appearance147.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.ultraGrid1.DisplayLayout.Override.HeaderAppearance = appearance147;
			appearance148.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance148.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance148.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ultraGrid1.DisplayLayout.Override.RowSelectorAppearance = appearance148;
			appearance149.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance149.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance149.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.ultraGrid1.DisplayLayout.Override.SelectedRowAppearance = appearance149;
			this.ultraGrid1.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.ultraGrid1.Location = new System.Drawing.Point(0, 8);
			this.ultraGrid1.Name = "ultraGrid1";
			this.ultraGrid1.Size = new System.Drawing.Size(1103, 224);
			this.ultraGrid1.TabIndex = 801;
			// 
			// ultraTabPageControl8
			// 
			this.ultraTabPageControl8.Controls.Add(this.ultraTabControl1);
			this.ultraTabPageControl8.Controls.Add(this.utgVerificacionTerrena);
			this.ultraTabPageControl8.Location = new System.Drawing.Point(-10000, -10000);
			this.ultraTabPageControl8.Name = "ultraTabPageControl8";
			this.ultraTabPageControl8.Size = new System.Drawing.Size(1116, 366);
			// 
			// ultraTabControl1
			// 
			appearance150.BackColor = System.Drawing.Color.LightSteelBlue;
			appearance150.BackColor2 = System.Drawing.Color.LightSteelBlue;
			this.ultraTabControl1.Appearance = appearance150;
			this.ultraTabControl1.BackColor = System.Drawing.SystemColors.Control;
			appearance151.BackColor = System.Drawing.Color.LightSteelBlue;
			appearance151.BackColor2 = System.Drawing.Color.LightSteelBlue;
			this.ultraTabControl1.ClientAreaAppearance = appearance151;
			this.ultraTabControl1.Controls.Add(this.ultraTabSharedControlsPage2);
			this.ultraTabControl1.Controls.Add(this.ultraTabPageControl15);
			this.ultraTabControl1.Controls.Add(this.ultraTabPageControl9);
			this.ultraTabControl1.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.ultraTabControl1.Location = new System.Drawing.Point(8, 136);
			this.ultraTabControl1.Name = "ultraTabControl1";
			this.ultraTabControl1.SharedControlsPage = this.ultraTabSharedControlsPage2;
			this.ultraTabControl1.Size = new System.Drawing.Size(1096, 224);
			this.ultraTabControl1.TabIndex = 796;
			ultraTab1.TabPage = this.ultraTabPageControl15;
			ultraTab1.Text = "Trabajo";
			ultraTab2.TabPage = this.ultraTabPageControl9;
			ultraTab2.Text = "Domicilio";
			this.ultraTabControl1.Tabs.AddRange(new Infragistics.Win.UltraWinTabControl.UltraTab[] {
																																															 ultraTab1,
																																															 ultraTab2});
			this.ultraTabControl1.SelectedTabChanged += new Infragistics.Win.UltraWinTabControl.SelectedTabChangedEventHandler(this.ultraTabControl1_SelectedTabChanged);
			// 
			// ultraTabSharedControlsPage2
			// 
			this.ultraTabSharedControlsPage2.Location = new System.Drawing.Point(-10000, -10000);
			this.ultraTabSharedControlsPage2.Name = "ultraTabSharedControlsPage2";
			this.ultraTabSharedControlsPage2.Size = new System.Drawing.Size(1092, 198);
			// 
			// utgVerificacionTerrena
			// 
			this.utgVerificacionTerrena.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.utgVerificacionTerrena.Cursor = System.Windows.Forms.Cursors.Default;
			this.utgVerificacionTerrena.DataSource = this.ultraDataSource50;
			appearance152.BackColor = System.Drawing.Color.White;
			this.utgVerificacionTerrena.DisplayLayout.Appearance = appearance152;
			ultraGridColumn182.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance153.TextHAlign = Infragistics.Win.HAlign.Center;
			ultraGridColumn182.CellAppearance = appearance153;
			appearance154.TextHAlign = Infragistics.Win.HAlign.Center;
			ultraGridColumn182.Header.Appearance = appearance154;
			ultraGridColumn182.Header.VisiblePosition = 7;
			ultraGridColumn183.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn183.Header.VisiblePosition = 8;
			ultraGridColumn183.Hidden = true;
			ultraGridColumn184.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn184.Header.VisiblePosition = 10;
			ultraGridColumn184.Hidden = true;
			ultraGridColumn185.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			appearance155.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn185.CellAppearance = appearance155;
			appearance156.TextHAlign = Infragistics.Win.HAlign.Right;
			ultraGridColumn185.FormulaErrorAppearance = appearance156;
			ultraGridColumn185.Header.VisiblePosition = 9;
			ultraGridColumn186.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn186.Header.VisiblePosition = 11;
			ultraGridColumn186.Width = 180;
			ultraGridColumn187.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn187.Header.VisiblePosition = 12;
			ultraGridColumn187.Hidden = true;
			ultraGridColumn188.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn188.Header.VisiblePosition = 6;
			ultraGridColumn189.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn189.Header.VisiblePosition = 13;
			ultraGridColumn189.Width = 150;
			ultraGridColumn190.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn190.Header.Caption = "Domicilio";
			ultraGridColumn190.Header.VisiblePosition = 15;
			ultraGridColumn190.Width = 80;
			ultraGridColumn191.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn191.Header.Caption = "Trabajo";
			ultraGridColumn191.Header.VisiblePosition = 17;
			ultraGridColumn191.Width = 80;
			ultraGridColumn192.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn192.Format = "dd/MM/yyyy HH:mm";
			ultraGridColumn192.Header.VisiblePosition = 18;
			ultraGridColumn192.Width = 120;
			ultraGridColumn193.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn193.Format = "dd/MM/yyyy HH:mm";
			ultraGridColumn193.Header.VisiblePosition = 19;
			ultraGridColumn193.Width = 120;
			ultraGridColumn194.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn194.Header.VisiblePosition = 20;
			ultraGridColumn195.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn195.Header.VisiblePosition = 24;
			ultraGridColumn195.Width = 200;
			ultraGridColumn196.CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly;
			ultraGridColumn196.Header.VisiblePosition = 25;
			ultraGridColumn196.Hidden = true;
			ultraGridColumn197.Header.VisiblePosition = 14;
			ultraGridColumn197.Width = 130;
			ultraGridColumn198.Header.VisiblePosition = 16;
			ultraGridColumn198.Width = 130;
			ultraGridColumn199.Header.VisiblePosition = 22;
			ultraGridColumn199.Hidden = true;
			ultraGridColumn200.Header.VisiblePosition = 21;
			ultraGridColumn200.Hidden = true;
			ultraGridColumn201.Header.VisiblePosition = 23;
			ultraGridColumn201.Hidden = true;
			ultraGridColumn202.ButtonDisplayStyle = Infragistics.Win.UltraWinGrid.ButtonDisplayStyle.Always;
			appearance157.Image = ((object)(resources.GetObject("appearance157.Image")));
			appearance157.ImageHAlign = Infragistics.Win.HAlign.Center;
			ultraGridColumn202.CellButtonAppearance = appearance157;
			ultraGridColumn202.Header.VisiblePosition = 0;
			ultraGridColumn202.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.Button;
			ultraGridColumn202.Width = 40;
			ultraGridColumn203.ButtonDisplayStyle = Infragistics.Win.UltraWinGrid.ButtonDisplayStyle.Always;
			appearance158.Image = ((object)(resources.GetObject("appearance158.Image")));
			appearance158.ImageHAlign = Infragistics.Win.HAlign.Center;
			ultraGridColumn203.CellButtonAppearance = appearance158;
			ultraGridColumn203.Header.VisiblePosition = 1;
			ultraGridColumn203.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.Button;
			ultraGridColumn203.Width = 70;
			ultraGridColumn204.ButtonDisplayStyle = Infragistics.Win.UltraWinGrid.ButtonDisplayStyle.Always;
			appearance159.Image = ((object)(resources.GetObject("appearance159.Image")));
			appearance159.ImageHAlign = Infragistics.Win.HAlign.Center;
			ultraGridColumn204.CellButtonAppearance = appearance159;
			ultraGridColumn204.Header.VisiblePosition = 2;
			ultraGridColumn204.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.Button;
			ultraGridColumn204.Width = 70;
			ultraGridColumn205.ButtonDisplayStyle = Infragistics.Win.UltraWinGrid.ButtonDisplayStyle.Always;
			appearance160.Image = ((object)(resources.GetObject("appearance160.Image")));
			appearance160.ImageHAlign = Infragistics.Win.HAlign.Center;
			ultraGridColumn205.CellButtonAppearance = appearance160;
			ultraGridColumn205.Header.Caption = "...";
			ultraGridColumn205.Header.VisiblePosition = 3;
			ultraGridColumn205.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.Button;
			ultraGridColumn205.Width = 30;
			ultraGridColumn206.ButtonDisplayStyle = Infragistics.Win.UltraWinGrid.ButtonDisplayStyle.Always;
			appearance161.Image = ((object)(resources.GetObject("appearance161.Image")));
			appearance161.ImageHAlign = Infragistics.Win.HAlign.Center;
			ultraGridColumn206.CellButtonAppearance = appearance161;
			ultraGridColumn206.Header.Caption = "Img T";
			ultraGridColumn206.Header.VisiblePosition = 4;
			ultraGridColumn206.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.Button;
			ultraGridColumn206.Width = 50;
			ultraGridColumn207.ButtonDisplayStyle = Infragistics.Win.UltraWinGrid.ButtonDisplayStyle.Always;
			appearance162.Image = ((object)(resources.GetObject("appearance162.Image")));
			appearance162.ImageHAlign = Infragistics.Win.HAlign.Center;
			ultraGridColumn207.CellButtonAppearance = appearance162;
			ultraGridColumn207.Header.Caption = "Img D";
			ultraGridColumn207.Header.VisiblePosition = 5;
			ultraGridColumn207.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.Button;
			ultraGridColumn207.Width = 50;
			ultraGridBand66.Columns.AddRange(new object[] {
																											ultraGridColumn182,
																											ultraGridColumn183,
																											ultraGridColumn184,
																											ultraGridColumn185,
																											ultraGridColumn186,
																											ultraGridColumn187,
																											ultraGridColumn188,
																											ultraGridColumn189,
																											ultraGridColumn190,
																											ultraGridColumn191,
																											ultraGridColumn192,
																											ultraGridColumn193,
																											ultraGridColumn194,
																											ultraGridColumn195,
																											ultraGridColumn196,
																											ultraGridColumn197,
																											ultraGridColumn198,
																											ultraGridColumn199,
																											ultraGridColumn200,
																											ultraGridColumn201,
																											ultraGridColumn202,
																											ultraGridColumn203,
																											ultraGridColumn204,
																											ultraGridColumn205,
																											ultraGridColumn206,
																											ultraGridColumn207});
			this.utgVerificacionTerrena.DisplayLayout.BandsSerializer.Add(ultraGridBand66);
			this.utgVerificacionTerrena.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.False;
			this.utgVerificacionTerrena.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.False;
			appearance163.BackColor = System.Drawing.Color.Transparent;
			this.utgVerificacionTerrena.DisplayLayout.Override.CardAreaAppearance = appearance163;
			appearance164.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance164.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance164.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			appearance164.FontData.BoldAsString = "True";
			appearance164.FontData.Name = "Arial";
			appearance164.FontData.SizeInPoints = 10F;
			appearance164.ForeColor = System.Drawing.Color.White;
			appearance164.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent;
			this.utgVerificacionTerrena.DisplayLayout.Override.HeaderAppearance = appearance164;
			this.utgVerificacionTerrena.DisplayLayout.Override.HeaderClickAction = Infragistics.Win.UltraWinGrid.HeaderClickAction.SortMulti;
			appearance165.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(89)), ((System.Byte)(135)), ((System.Byte)(214)));
			appearance165.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(7)), ((System.Byte)(59)), ((System.Byte)(150)));
			appearance165.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.utgVerificacionTerrena.DisplayLayout.Override.RowSelectorAppearance = appearance165;
			appearance166.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(251)), ((System.Byte)(230)), ((System.Byte)(148)));
			appearance166.BackColor2 = System.Drawing.Color.FromArgb(((System.Byte)(238)), ((System.Byte)(149)), ((System.Byte)(21)));
			appearance166.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical;
			this.utgVerificacionTerrena.DisplayLayout.Override.SelectedRowAppearance = appearance166;
			this.utgVerificacionTerrena.DisplayLayout.Override.SelectTypeCell = Infragistics.Win.UltraWinGrid.SelectType.None;
			this.utgVerificacionTerrena.DisplayLayout.Override.SelectTypeCol = Infragistics.Win.UltraWinGrid.SelectType.None;
			this.utgVerificacionTerrena.DisplayLayout.Override.SelectTypeRow = Infragistics.Win.UltraWinGrid.SelectType.Single;
			this.utgVerificacionTerrena.DisplayLayout.ScrollBounds = Infragistics.Win.UltraWinGrid.ScrollBounds.ScrollToFill;
			this.utgVerificacionTerrena.DisplayLayout.ScrollStyle = Infragistics.Win.UltraWinGrid.ScrollStyle.Immediate;
			this.utgVerificacionTerrena.DisplayLayout.TabNavigation = Infragistics.Win.UltraWinGrid.TabNavigation.NextControl;
			this.utgVerificacionTerrena.DisplayLayout.ViewStyle = Infragistics.Win.UltraWinGrid.ViewStyle.SingleBand;
			this.utgVerificacionTerrena.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.utgVerificacionTerrena.Location = new System.Drawing.Point(8, 16);
			this.utgVerificacionTerrena.Name = "utgVerificacionTerrena";
			this.utgVerificacionTerrena.RightToLeft = System.Windows.Forms.RightToLeft.No;
			this.utgVerificacionTerrena.Size = new System.Drawing.Size(1088, 112);
			this.utgVerificacionTerrena.TabIndex = 789;
			this.utgVerificacionTerrena.BeforeSelectChange += new Infragistics.Win.UltraWinGrid.BeforeSelectChangeEventHandler(this.utgVerificacionTerrena_BeforeSelectChange);
			this.utgVerificacionTerrena.ClickCellButton += new Infragistics.Win.UltraWinGrid.CellEventHandler(this.utgVerificacionTerrena_ClickCellButton);
			// 
			// ultraDataSource50
			// 
			ultraDataColumn126.DataType = typeof(int);
			ultraDataColumn127.DataType = typeof(int);
			ultraDataColumn130.DataType = typeof(int);
			ultraDataColumn133.DataType = typeof(System.Byte);
			ultraDataColumn134.DataType = typeof(System.Byte);
			ultraDataColumn135.DataType = typeof(System.DateTime);
			ultraDataColumn136.DataType = typeof(System.DateTime);
			ultraDataColumn140.DataType = typeof(bool);
			ultraDataColumn141.DataType = typeof(bool);
			ultraDataColumn142.DataType = typeof(int);
			ultraDataColumn143.DataType = typeof(int);
			ultraDataColumn144.DataType = typeof(int);
			this.ultraDataSource50.Band.Columns.AddRange(new object[] {
																																	ultraDataColumn125,
																																	ultraDataColumn126,
																																	ultraDataColumn127,
																																	ultraDataColumn128,
																																	ultraDataColumn129,
																																	ultraDataColumn130,
																																	ultraDataColumn131,
																																	ultraDataColumn132,
																																	ultraDataColumn133,
																																	ultraDataColumn134,
																																	ultraDataColumn135,
																																	ultraDataColumn136,
																																	ultraDataColumn137,
																																	ultraDataColumn138,
																																	ultraDataColumn139,
																																	ultraDataColumn140,
																																	ultraDataColumn141,
																																	ultraDataColumn142,
																																	ultraDataColumn143,
																																	ultraDataColumn144});
			// 
			// ultraDataSource5
			// 
			ultraDataColumn145.DataType = typeof(int);
			this.ultraDataSource5.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn145,
																																 ultraDataColumn146});
			// 
			// ultraDataSource34
			// 
			ultraDataColumn147.DataType = typeof(int);
			this.ultraDataSource34.Band.Columns.AddRange(new object[] {
																																	ultraDataColumn147,
																																	ultraDataColumn148});
			// 
			// ultraDataSource39
			// 
			ultraDataColumn149.DataType = typeof(int);
			this.ultraDataSource39.Band.Columns.AddRange(new object[] {
																																	ultraDataColumn149,
																																	ultraDataColumn150});
			// 
			// ultraDataSource32
			// 
			ultraDataColumn151.DataType = typeof(int);
			this.ultraDataSource32.Band.Columns.AddRange(new object[] {
																																	ultraDataColumn151,
																																	ultraDataColumn152});
			// 
			// label1
			// 
			this.label1.AutoSize = true;
			this.label1.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label1.Location = new System.Drawing.Point(264, 0);
			this.label1.Name = "label1";
			this.label1.Size = new System.Drawing.Size(61, 17);
			this.label1.TabIndex = 775;
			this.label1.Text = "Documento";
			this.label1.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// lblBodega
			// 
			this.lblBodega.AutoSize = true;
			this.lblBodega.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.lblBodega.Location = new System.Drawing.Point(120, 0);
			this.lblBodega.Name = "lblBodega";
			this.lblBodega.Size = new System.Drawing.Size(86, 17);
			this.lblBodega.TabIndex = 774;
			this.lblBodega.Text = "Tipo Documento";
			this.lblBodega.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// txtDocumento
			// 
			appearance167.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtDocumento.Appearance = appearance167;
			this.txtDocumento.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtDocumento.Location = new System.Drawing.Point(264, 16);
			this.txtDocumento.MaxLength = 10;
			this.txtDocumento.Name = "txtDocumento";
			this.txtDocumento.Size = new System.Drawing.Size(128, 21);
			this.txtDocumento.TabIndex = 776;
			this.txtDocumento.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtDocumento_KeyDown);
			this.txtDocumento.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.txtDocumento_KeyPress);
			this.txtDocumento.Validating += new System.ComponentModel.CancelEventHandler(this.txtDocumento_Validating);
			this.txtDocumento.Validated += new System.EventHandler(this.txtDocumento_Validated);
			this.txtDocumento.TextChanged += new System.EventHandler(this.txtDocumento_TextChanged);
			this.txtDocumento.ValueChanged += new System.EventHandler(this.txtDocumento_ValueChanged);
			// 
			// dtFecha
			// 
			this.dtFecha.ButtonStyle = Infragistics.Win.UIElementButtonStyle.Office2003ToolbarButton;
			dateButton7.Caption = "Today";
			this.dtFecha.DateButtons.Add(dateButton7);
			this.dtFecha.Enabled = false;
			this.dtFecha.Format = "dd/MM/yyyy";
			this.dtFecha.Location = new System.Drawing.Point(8, 16);
			this.dtFecha.Name = "dtFecha";
			this.dtFecha.NonAutoSizeHeight = 23;
			this.dtFecha.Size = new System.Drawing.Size(104, 21);
			this.dtFecha.TabIndex = 778;
			this.dtFecha.Value = ((object)(resources.GetObject("dtFecha.Value")));
			// 
			// label5
			// 
			this.label5.AutoSize = true;
			this.label5.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label5.Location = new System.Drawing.Point(952, 0);
			this.label5.Name = "label5";
			this.label5.Size = new System.Drawing.Size(91, 17);
			this.label5.TabIndex = 782;
			this.label5.Text = "Segundo Nombre";
			this.label5.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// label4
			// 
			this.label4.AutoSize = true;
			this.label4.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label4.Location = new System.Drawing.Point(768, 0);
			this.label4.Name = "label4";
			this.label4.Size = new System.Drawing.Size(80, 17);
			this.label4.TabIndex = 781;
			this.label4.Text = "Primer Nombre";
			this.label4.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.label4.Click += new System.EventHandler(this.label4_Click);
			// 
			// label3
			// 
			this.label3.AutoSize = true;
			this.label3.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label3.Location = new System.Drawing.Point(584, 0);
			this.label3.Name = "label3";
			this.label3.Size = new System.Drawing.Size(88, 17);
			this.label3.TabIndex = 780;
			this.label3.Text = "Apellido Materno";
			this.label3.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// label2
			// 
			this.label2.AutoSize = true;
			this.label2.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label2.Location = new System.Drawing.Point(400, 0);
			this.label2.Name = "label2";
			this.label2.Size = new System.Drawing.Size(85, 17);
			this.label2.TabIndex = 779;
			this.label2.Text = "Apellido Paterno";
			this.label2.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// txtSnombre
			// 
			appearance168.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtSnombre.Appearance = appearance168;
			this.txtSnombre.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtSnombre.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtSnombre.Enabled = false;
			this.txtSnombre.Location = new System.Drawing.Point(952, 16);
			this.txtSnombre.Name = "txtSnombre";
			this.txtSnombre.Size = new System.Drawing.Size(180, 21);
			this.txtSnombre.TabIndex = 786;
			this.txtSnombre.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtSnombre_KeyDown);
			this.txtSnombre.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.txtSnombre_KeyPress);
			// 
			// txtPnombre
			// 
			appearance169.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtPnombre.Appearance = appearance169;
			this.txtPnombre.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtPnombre.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtPnombre.Enabled = false;
			this.txtPnombre.Location = new System.Drawing.Point(768, 16);
			this.txtPnombre.Name = "txtPnombre";
			this.txtPnombre.Size = new System.Drawing.Size(180, 21);
			this.txtPnombre.TabIndex = 785;
			this.txtPnombre.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtPnombre_KeyDown);
			this.txtPnombre.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.txtPnombre_KeyPress);
			// 
			// txtApellidoM
			// 
			appearance170.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtApellidoM.Appearance = appearance170;
			this.txtApellidoM.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtApellidoM.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtApellidoM.Enabled = false;
			this.txtApellidoM.Location = new System.Drawing.Point(584, 16);
			this.txtApellidoM.Name = "txtApellidoM";
			this.txtApellidoM.Size = new System.Drawing.Size(180, 21);
			this.txtApellidoM.TabIndex = 784;
			this.txtApellidoM.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtApellidoM_KeyDown);
			this.txtApellidoM.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.txtApellidoM_KeyPress);
			// 
			// txtApellidoP
			// 
			appearance171.ForeColorDisabled = System.Drawing.Color.Black;
			this.txtApellidoP.Appearance = appearance171;
			this.txtApellidoP.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
			this.txtApellidoP.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtApellidoP.Enabled = false;
			this.txtApellidoP.Location = new System.Drawing.Point(400, 16);
			this.txtApellidoP.Name = "txtApellidoP";
			this.txtApellidoP.Size = new System.Drawing.Size(180, 21);
			this.txtApellidoP.TabIndex = 783;
			this.txtApellidoP.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtApellidoP_KeyDown);
			this.txtApellidoP.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.txtApellidoP_KeyPress);
			// 
			// label16
			// 
			this.label16.AutoSize = true;
			this.label16.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label16.Location = new System.Drawing.Point(168, 40);
			this.label16.Name = "label16";
			this.label16.Size = new System.Drawing.Size(30, 17);
			this.label16.TabIndex = 789;
			this.label16.Text = "Local";
			this.label16.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
			// 
			// lblEntidad
			// 
			this.lblEntidad.AutoSize = true;
			this.lblEntidad.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.lblEntidad.Location = new System.Drawing.Point(8, 40);
			this.lblEntidad.Name = "lblEntidad";
			this.lblEntidad.Size = new System.Drawing.Size(96, 17);
			this.lblEntidad.TabIndex = 787;
			this.lblEntidad.Text = "Entidad Financiera";
			this.lblEntidad.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// cmbBodega
			// 
			this.cmbBodega.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbBodega.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbBodega.DataSource = this.ultraDataSource5;
			ultraGridColumn208.Header.VisiblePosition = 0;
			ultraGridColumn208.Hidden = true;
			ultraGridColumn209.Header.Caption = "Local";
			ultraGridColumn209.Header.VisiblePosition = 1;
			ultraGridColumn209.Width = 150;
			ultraGridBand67.Columns.AddRange(new object[] {
																											ultraGridColumn208,
																											ultraGridColumn209});
			this.cmbBodega.DisplayLayout.BandsSerializer.Add(ultraGridBand67);
			this.cmbBodega.DisplayLayout.Override.RowSelectors = Infragistics.Win.DefaultableBoolean.True;
			this.cmbBodega.DisplayMember = "Nombre";
			this.cmbBodega.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbBodega.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbBodega.Enabled = false;
			this.cmbBodega.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbBodega.Location = new System.Drawing.Point(168, 56);
			this.cmbBodega.Name = "cmbBodega";
			this.cmbBodega.Size = new System.Drawing.Size(150, 22);
			this.cmbBodega.TabIndex = 790;
			this.cmbBodega.ValueMember = "Bodega";
			this.cmbBodega.KeyDown += new System.Windows.Forms.KeyEventHandler(this.cmbBodega_KeyDown);
			// 
			// cmbEntidad
			// 
			this.cmbEntidad.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal;
			this.cmbEntidad.Cursor = System.Windows.Forms.Cursors.Default;
			this.cmbEntidad.DataSource = this.ultraDataSource10;
			ultraGridColumn210.Header.VisiblePosition = 0;
			ultraGridColumn210.Hidden = true;
			ultraGridColumn210.Width = 150;
			ultraGridColumn211.Header.Caption = "Entidad Financiera";
			ultraGridColumn211.Header.VisiblePosition = 1;
			ultraGridBand68.Columns.AddRange(new object[] {
																											ultraGridColumn210,
																											ultraGridColumn211});
			this.cmbEntidad.DisplayLayout.BandsSerializer.Add(ultraGridBand68);
			this.cmbEntidad.DisplayMember = "Nombre";
			this.cmbEntidad.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.cmbEntidad.DropDownStyle = Infragistics.Win.UltraWinGrid.UltraComboStyle.DropDownList;
			this.cmbEntidad.Enabled = false;
			this.cmbEntidad.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.cmbEntidad.Location = new System.Drawing.Point(8, 56);
			this.cmbEntidad.MaxDropDownItems = 30;
			this.cmbEntidad.Name = "cmbEntidad";
			this.cmbEntidad.Size = new System.Drawing.Size(150, 21);
			this.cmbEntidad.TabIndex = 788;
			this.cmbEntidad.ValueMember = "idEntidadFinanciera";
			this.cmbEntidad.KeyDown += new System.Windows.Forms.KeyEventHandler(this.cmbEntidad_KeyDown);
			this.cmbEntidad.ValueChanged += new System.EventHandler(this.cmbEntidad_ValueChanged);
			// 
			// ultraDataSource10
			// 
			ultraDataColumn153.DataType = typeof(int);
			this.ultraDataSource10.Band.Columns.AddRange(new object[] {
																																	ultraDataColumn153,
																																	ultraDataColumn154});
			// 
			// btnGarante
			// 
			this.btnGarante.Enabled = false;
			this.btnGarante.Location = new System.Drawing.Point(408, 56);
			this.btnGarante.Name = "btnGarante";
			this.btnGarante.Size = new System.Drawing.Size(56, 23);
			this.btnGarante.TabIndex = 794;
			this.btnGarante.Text = "Garante";
			this.btnGarante.Click += new System.EventHandler(this.btnGarante_Click);
			// 
			// chkGarante
			// 
			appearance172.ForeColorDisabled = System.Drawing.Color.Black;
			this.chkGarante.Appearance = appearance172;
			this.chkGarante.BackColor = System.Drawing.Color.Transparent;
			this.chkGarante.CausesValidation = false;
			this.chkGarante.Enabled = false;
			this.chkGarante.Location = new System.Drawing.Point(328, 56);
			this.chkGarante.Name = "chkGarante";
			this.chkGarante.Size = new System.Drawing.Size(80, 22);
			this.chkGarante.TabIndex = 793;
			this.chkGarante.Text = "Garante";
			this.chkGarante.CheckedChanged += new System.EventHandler(this.chkGarante_CheckedChanged);
			// 
			// ultraDataSource7
			// 
			ultraDataColumn155.DataType = typeof(int);
			this.ultraDataSource7.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn155,
																																 ultraDataColumn156});
			// 
			// ultraDataSource8
			// 
			ultraDataColumn157.DataType = typeof(int);
			this.ultraDataSource8.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn157,
																																 ultraDataColumn158});
			// 
			// ultraDataSource6
			// 
			ultraDataColumn159.DataType = typeof(int);
			this.ultraDataSource6.Band.Columns.AddRange(new object[] {
																																 ultraDataColumn159,
																																 ultraDataColumn160});
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
			// tbBasico
			// 
			appearance173.BackColor = System.Drawing.Color.LightSteelBlue;
			appearance173.BackColor2 = System.Drawing.Color.LightSteelBlue;
			this.tbBasico.Appearance = appearance173;
			this.tbBasico.BackColor = System.Drawing.SystemColors.Control;
			appearance174.BackColor = System.Drawing.Color.LightSteelBlue;
			appearance174.BackColor2 = System.Drawing.Color.LightSteelBlue;
			this.tbBasico.ClientAreaAppearance = appearance174;
			this.tbBasico.Controls.Add(this.ultraTabSharedControlsPage1);
			this.tbBasico.Controls.Add(this.ultraTabPageControl1);
			this.tbBasico.Controls.Add(this.tpComplemento);
			this.tbBasico.Controls.Add(this.ultraTabPageControl5);
			this.tbBasico.Controls.Add(this.ultraTabPageControl7);
			this.tbBasico.Controls.Add(this.ultraTabPageControl2);
			this.tbBasico.Controls.Add(this.ultraTabPageControl3);
			this.tbBasico.Controls.Add(this.ultraTabPageControl4);
			this.tbBasico.Controls.Add(this.ultraTabPageControl6);
			this.tbBasico.Controls.Add(this.ultraTabPageControl8);
			this.tbBasico.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.tbBasico.Location = new System.Drawing.Point(8, 88);
			this.tbBasico.Name = "tbBasico";
			this.tbBasico.SharedControlsPage = this.ultraTabSharedControlsPage1;
			this.tbBasico.Size = new System.Drawing.Size(1120, 392);
			this.tbBasico.TabIndex = 795;
			appearance175.BackColor = System.Drawing.Color.LightSteelBlue;
			appearance175.BackColor2 = System.Drawing.Color.LightSteelBlue;
			ultraTab3.ActiveAppearance = appearance175;
			appearance176.BackColor = System.Drawing.Color.LightSteelBlue;
			appearance176.BackColor2 = System.Drawing.Color.LightSteelBlue;
			ultraTab3.Appearance = appearance176;
			appearance177.BackColor = System.Drawing.Color.LightSteelBlue;
			appearance177.BackColor2 = System.Drawing.Color.LightSteelBlue;
			ultraTab3.ClientAreaAppearance = appearance177;
			appearance178.BackColor = System.Drawing.Color.LightSteelBlue;
			appearance178.BackColor2 = System.Drawing.Color.LightSteelBlue;
			ultraTab3.SelectedAppearance = appearance178;
			ultraTab3.TabPage = this.ultraTabPageControl1;
			ultraTab3.Text = "Datos Cliente";
			ultraTab4.TabPage = this.tpComplemento;
			ultraTab4.Text = "Datos Conyuge";
			ultraTab5.TabPage = this.ultraTabPageControl5;
			ultraTab5.Text = "Referencias";
			ultraTab6.TabPage = this.ultraTabPageControl7;
			ultraTab6.Text = "Sección A- Negocio";
			ultraTab7.TabPage = this.ultraTabPageControl2;
			ultraTab7.Text = "Sección B- Dependiente";
			ultraTab8.TabPage = this.ultraTabPageControl3;
			ultraTab8.Text = "Información de Crédito";
			ultraTab9.TabPage = this.ultraTabPageControl6;
			ultraTab9.Text = "Factores de Crédito ";
			ultraTab10.TabPage = this.ultraTabPageControl4;
			ultraTab10.Text = "Tarjetas";
			ultraTab10.Visible = false;
			ultraTab11.TabPage = this.ultraTabPageControl8;
			ultraTab11.Text = "Verificación";
			this.tbBasico.Tabs.AddRange(new Infragistics.Win.UltraWinTabControl.UltraTab[] {
																																											 ultraTab3,
																																											 ultraTab4,
																																											 ultraTab5,
																																											 ultraTab6,
																																											 ultraTab7,
																																											 ultraTab8,
																																											 ultraTab9,
																																											 ultraTab10,
																																											 ultraTab11});
			this.tbBasico.SelectedTabChanged += new Infragistics.Win.UltraWinTabControl.SelectedTabChangedEventHandler(this.tbBasico_SelectedTabChanged);
			// 
			// ultraTabSharedControlsPage1
			// 
			this.ultraTabSharedControlsPage1.Location = new System.Drawing.Point(-10000, -10000);
			this.ultraTabSharedControlsPage1.Name = "ultraTabSharedControlsPage1";
			this.ultraTabSharedControlsPage1.Size = new System.Drawing.Size(1116, 366);
			// 
			// ultraDataSource12
			// 
			ultraDataColumn161.DataType = typeof(int);
			this.ultraDataSource12.Band.Columns.AddRange(new object[] {
																																	ultraDataColumn161,
																																	ultraDataColumn162});
			// 
			// lblFISL
			// 
			this.lblFISL.AutoSize = true;
			this.lblFISL.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.lblFISL.ForeColor = System.Drawing.Color.Firebrick;
			this.lblFISL.Location = new System.Drawing.Point(856, 56);
			this.lblFISL.Name = "lblFISL";
			this.lblFISL.Size = new System.Drawing.Size(0, 19);
			this.lblFISL.TabIndex = 807;
			this.lblFISL.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
			// 
			// lblEstado
			// 
			this.lblEstado.AutoSize = true;
			this.lblEstado.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.lblEstado.ForeColor = System.Drawing.Color.Firebrick;
			this.lblEstado.Location = new System.Drawing.Point(688, 56);
			this.lblEstado.Name = "lblEstado";
			this.lblEstado.Size = new System.Drawing.Size(0, 19);
			this.lblEstado.TabIndex = 808;
			this.lblEstado.TextAlign = System.Drawing.ContentAlignment.BottomRight;
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
			this.cmbTipoRuc.Location = new System.Drawing.Point(120, 16);
			this.cmbTipoRuc.Name = "cmbTipoRuc";
			this.cmbTipoRuc.Size = new System.Drawing.Size(136, 22);
			this.cmbTipoRuc.TabIndex = 777;
			this.cmbTipoRuc.ValueMember = "idTipoDoc";
			// 
			// txtNegocioIdCre_DatosUbicacion
			// 
			this.txtNegocioIdCre_DatosUbicacion.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtNegocioIdCre_DatosUbicacion.Enabled = false;
			this.txtNegocioIdCre_DatosUbicacion.Location = new System.Drawing.Point(1032, 56);
			this.txtNegocioIdCre_DatosUbicacion.Name = "txtNegocioIdCre_DatosUbicacion";
			this.txtNegocioIdCre_DatosUbicacion.PromptChar = ' ';
			this.txtNegocioIdCre_DatosUbicacion.Size = new System.Drawing.Size(50, 21);
			this.txtNegocioIdCre_DatosUbicacion.SpinButtonAlignment = Infragistics.Win.ButtonAlignment.Left;
			this.txtNegocioIdCre_DatosUbicacion.TabIndex = 819;
			this.txtNegocioIdCre_DatosUbicacion.Visible = false;
			// 
			// txtClienteIdCre_EmpresasTrabajo
			// 
			this.txtClienteIdCre_EmpresasTrabajo.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtClienteIdCre_EmpresasTrabajo.Enabled = false;
			this.txtClienteIdCre_EmpresasTrabajo.Location = new System.Drawing.Point(1056, 56);
			this.txtClienteIdCre_EmpresasTrabajo.Name = "txtClienteIdCre_EmpresasTrabajo";
			this.txtClienteIdCre_EmpresasTrabajo.PromptChar = ' ';
			this.txtClienteIdCre_EmpresasTrabajo.Size = new System.Drawing.Size(50, 21);
			this.txtClienteIdCre_EmpresasTrabajo.SpinButtonAlignment = Infragistics.Win.ButtonAlignment.Left;
			this.txtClienteIdCre_EmpresasTrabajo.TabIndex = 818;
			this.txtClienteIdCre_EmpresasTrabajo.Visible = false;
			// 
			// txtEmpresaIdCre_DatosUbicacion
			// 
			this.txtEmpresaIdCre_DatosUbicacion.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtEmpresaIdCre_DatosUbicacion.Enabled = false;
			this.txtEmpresaIdCre_DatosUbicacion.Location = new System.Drawing.Point(1080, 56);
			this.txtEmpresaIdCre_DatosUbicacion.Name = "txtEmpresaIdCre_DatosUbicacion";
			this.txtEmpresaIdCre_DatosUbicacion.PromptChar = ' ';
			this.txtEmpresaIdCre_DatosUbicacion.Size = new System.Drawing.Size(50, 21);
			this.txtEmpresaIdCre_DatosUbicacion.SpinButtonAlignment = Infragistics.Win.ButtonAlignment.Left;
			this.txtEmpresaIdCre_DatosUbicacion.TabIndex = 817;
			this.txtEmpresaIdCre_DatosUbicacion.Visible = false;
			// 
			// txtConyugueIdCre_DatosGenerales
			// 
			this.txtConyugueIdCre_DatosGenerales.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtConyugueIdCre_DatosGenerales.Enabled = false;
			this.txtConyugueIdCre_DatosGenerales.Location = new System.Drawing.Point(976, 56);
			this.txtConyugueIdCre_DatosGenerales.Name = "txtConyugueIdCre_DatosGenerales";
			this.txtConyugueIdCre_DatosGenerales.PromptChar = ' ';
			this.txtConyugueIdCre_DatosGenerales.Size = new System.Drawing.Size(50, 21);
			this.txtConyugueIdCre_DatosGenerales.SpinButtonAlignment = Infragistics.Win.ButtonAlignment.Left;
			this.txtConyugueIdCre_DatosGenerales.TabIndex = 816;
			this.txtConyugueIdCre_DatosGenerales.Visible = false;
			// 
			// txtConyugueIdCre_DatosUbicacion
			// 
			this.txtConyugueIdCre_DatosUbicacion.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtConyugueIdCre_DatosUbicacion.Enabled = false;
			this.txtConyugueIdCre_DatosUbicacion.Location = new System.Drawing.Point(1000, 56);
			this.txtConyugueIdCre_DatosUbicacion.Name = "txtConyugueIdCre_DatosUbicacion";
			this.txtConyugueIdCre_DatosUbicacion.PromptChar = ' ';
			this.txtConyugueIdCre_DatosUbicacion.Size = new System.Drawing.Size(50, 21);
			this.txtConyugueIdCre_DatosUbicacion.SpinButtonAlignment = Infragistics.Win.ButtonAlignment.Left;
			this.txtConyugueIdCre_DatosUbicacion.TabIndex = 815;
			this.txtConyugueIdCre_DatosUbicacion.Visible = false;
			// 
			// txtClienteIdCre_Negocio
			// 
			this.txtClienteIdCre_Negocio.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtClienteIdCre_Negocio.Enabled = false;
			this.txtClienteIdCre_Negocio.Location = new System.Drawing.Point(1016, 56);
			this.txtClienteIdCre_Negocio.Name = "txtClienteIdCre_Negocio";
			this.txtClienteIdCre_Negocio.PromptChar = ' ';
			this.txtClienteIdCre_Negocio.Size = new System.Drawing.Size(50, 21);
			this.txtClienteIdCre_Negocio.SpinButtonAlignment = Infragistics.Win.ButtonAlignment.Left;
			this.txtClienteIdCre_Negocio.TabIndex = 814;
			this.txtClienteIdCre_Negocio.Visible = false;
			// 
			// txtIdCre_Conyugue
			// 
			this.txtIdCre_Conyugue.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtIdCre_Conyugue.Enabled = false;
			this.txtIdCre_Conyugue.Location = new System.Drawing.Point(960, 56);
			this.txtIdCre_Conyugue.Name = "txtIdCre_Conyugue";
			this.txtIdCre_Conyugue.PromptChar = ' ';
			this.txtIdCre_Conyugue.Size = new System.Drawing.Size(50, 21);
			this.txtIdCre_Conyugue.SpinButtonAlignment = Infragistics.Win.ButtonAlignment.Left;
			this.txtIdCre_Conyugue.TabIndex = 813;
			this.txtIdCre_Conyugue.Visible = false;
			// 
			// txtIdCre_Solicitud
			// 
			this.txtIdCre_Solicitud.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtIdCre_Solicitud.Enabled = false;
			this.txtIdCre_Solicitud.Location = new System.Drawing.Point(904, 56);
			this.txtIdCre_Solicitud.Name = "txtIdCre_Solicitud";
			this.txtIdCre_Solicitud.PromptChar = ' ';
			this.txtIdCre_Solicitud.Size = new System.Drawing.Size(50, 21);
			this.txtIdCre_Solicitud.SpinButtonAlignment = Infragistics.Win.ButtonAlignment.Left;
			this.txtIdCre_Solicitud.TabIndex = 812;
			this.txtIdCre_Solicitud.Visible = false;
			// 
			// txtClienteIdCre_DatosGenerales
			// 
			this.txtClienteIdCre_DatosGenerales.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtClienteIdCre_DatosGenerales.Enabled = false;
			this.txtClienteIdCre_DatosGenerales.Location = new System.Drawing.Point(920, 56);
			this.txtClienteIdCre_DatosGenerales.Name = "txtClienteIdCre_DatosGenerales";
			this.txtClienteIdCre_DatosGenerales.PromptChar = ' ';
			this.txtClienteIdCre_DatosGenerales.Size = new System.Drawing.Size(50, 21);
			this.txtClienteIdCre_DatosGenerales.SpinButtonAlignment = Infragistics.Win.ButtonAlignment.Left;
			this.txtClienteIdCre_DatosGenerales.TabIndex = 811;
			this.txtClienteIdCre_DatosGenerales.Visible = false;
			// 
			// txtClienteIdCre_DatosUbicacion
			// 
			this.txtClienteIdCre_DatosUbicacion.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtClienteIdCre_DatosUbicacion.Enabled = false;
			this.txtClienteIdCre_DatosUbicacion.Location = new System.Drawing.Point(944, 56);
			this.txtClienteIdCre_DatosUbicacion.Name = "txtClienteIdCre_DatosUbicacion";
			this.txtClienteIdCre_DatosUbicacion.PromptChar = ' ';
			this.txtClienteIdCre_DatosUbicacion.Size = new System.Drawing.Size(50, 21);
			this.txtClienteIdCre_DatosUbicacion.SpinButtonAlignment = Infragistics.Win.ButtonAlignment.Left;
			this.txtClienteIdCre_DatosUbicacion.TabIndex = 810;
			this.txtClienteIdCre_DatosUbicacion.Visible = false;
			// 
			// txtEstado
			// 
			this.txtEstado.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtEstado.Enabled = false;
			this.txtEstado.Location = new System.Drawing.Point(880, 56);
			this.txtEstado.Name = "txtEstado";
			this.txtEstado.PromptChar = ' ';
			this.txtEstado.Size = new System.Drawing.Size(50, 21);
			this.txtEstado.SpinButtonAlignment = Infragistics.Win.ButtonAlignment.Left;
			this.txtEstado.TabIndex = 809;
			this.txtEstado.Visible = false;
			// 
			// btnGuardar
			// 
			this.btnGuardar.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btnGuardar.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(230)), ((System.Byte)(230)), ((System.Byte)(230)));
			this.btnGuardar.Font = new System.Drawing.Font("Tahoma", 8.25F);
			this.btnGuardar.ForeColor = System.Drawing.Color.FromArgb(((System.Byte)(50)), ((System.Byte)(50)), ((System.Byte)(50)));
			this.btnGuardar.Image = ((System.Drawing.Image)(resources.GetObject("btnGuardar.Image")));
			this.btnGuardar.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnGuardar.Location = new System.Drawing.Point(8, 496);
			this.btnGuardar.Name = "btnGuardar";
			this.btnGuardar.Size = new System.Drawing.Size(72, 23);
			this.btnGuardar.TabIndex = 822;
			this.btnGuardar.Text = "Guardar";
			this.btnGuardar.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnGuardar.Click += new System.EventHandler(this.btnGuardar_Click);
			// 
			// btnSalir
			// 
			this.btnSalir.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btnSalir.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(230)), ((System.Byte)(230)), ((System.Byte)(230)));
			this.btnSalir.CausesValidation = false;
			this.btnSalir.Font = new System.Drawing.Font("Tahoma", 8.25F);
			this.btnSalir.ForeColor = System.Drawing.Color.FromArgb(((System.Byte)(50)), ((System.Byte)(50)), ((System.Byte)(50)));
			this.btnSalir.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnSalir.ImageIndex = 62;
			this.btnSalir.Location = new System.Drawing.Point(264, 496);
			this.btnSalir.Name = "btnSalir";
			this.btnSalir.Size = new System.Drawing.Size(72, 23);
			this.btnSalir.TabIndex = 825;
			this.btnSalir.Text = "Salir";
			this.btnSalir.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnSalir.Click += new System.EventHandler(this.btnSalir_Click);
			// 
			// btnConsultar
			// 
			this.btnConsultar.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btnConsultar.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(230)), ((System.Byte)(230)), ((System.Byte)(230)));
			this.btnConsultar.CausesValidation = false;
			this.btnConsultar.Font = new System.Drawing.Font("Tahoma", 8.25F);
			this.btnConsultar.ForeColor = System.Drawing.Color.FromArgb(((System.Byte)(50)), ((System.Byte)(50)), ((System.Byte)(50)));
			this.btnConsultar.Image = ((System.Drawing.Image)(resources.GetObject("btnConsultar.Image")));
			this.btnConsultar.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnConsultar.Location = new System.Drawing.Point(88, 496);
			this.btnConsultar.Name = "btnConsultar";
			this.btnConsultar.Size = new System.Drawing.Size(80, 23);
			this.btnConsultar.TabIndex = 821;
			this.btnConsultar.Text = "Consultar";
			this.btnConsultar.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnConsultar.Click += new System.EventHandler(this.btnConsultar_Click);
			// 
			// btnImprimir
			// 
			this.btnImprimir.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btnImprimir.CausesValidation = false;
			this.btnImprimir.Image = ((System.Drawing.Image)(resources.GetObject("btnImprimir.Image")));
			this.btnImprimir.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnImprimir.Location = new System.Drawing.Point(176, 496);
			this.btnImprimir.Name = "btnImprimir";
			this.btnImprimir.Size = new System.Drawing.Size(77, 23);
			this.btnImprimir.TabIndex = 826;
			this.btnImprimir.Text = "&Imprimir";
			this.btnImprimir.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnImprimir.Click += new System.EventHandler(this.btnImprimir_Click);
			// 
			// txtIdGarante
			// 
			this.txtIdGarante.DisplayStyle = Infragistics.Win.EmbeddableElementDisplayStyle.Office2003;
			this.txtIdGarante.Enabled = false;
			this.txtIdGarante.Location = new System.Drawing.Point(488, 496);
			this.txtIdGarante.Name = "txtIdGarante";
			this.txtIdGarante.PromptChar = ' ';
			this.txtIdGarante.Size = new System.Drawing.Size(50, 21);
			this.txtIdGarante.SpinButtonAlignment = Infragistics.Win.ButtonAlignment.Left;
			this.txtIdGarante.TabIndex = 827;
			this.txtIdGarante.Visible = false;
			// 
			// btnIdValidator
			// 
			this.btnIdValidator.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.btnIdValidator.CausesValidation = false;
			this.btnIdValidator.Image = ((System.Drawing.Image)(resources.GetObject("btnIdValidator.Image")));
			this.btnIdValidator.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnIdValidator.Location = new System.Drawing.Point(736, 488);
			this.btnIdValidator.Name = "btnIdValidator";
			this.btnIdValidator.Size = new System.Drawing.Size(96, 23);
			this.btnIdValidator.TabIndex = 828;
			this.btnIdValidator.Text = "& Id Validator";
			this.btnIdValidator.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnIdValidator.Visible = false;
			this.btnIdValidator.Click += new System.EventHandler(this.btnIdValidator_Click);
			// 
			// btnComandos
			// 
			this.btnComandos.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			this.btnComandos.Image = ((System.Drawing.Image)(resources.GetObject("btnComandos.Image")));
			this.btnComandos.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnComandos.Location = new System.Drawing.Point(640, 488);
			this.btnComandos.Name = "btnComandos";
			this.btnComandos.Size = new System.Drawing.Size(80, 23);
			this.btnComandos.TabIndex = 1050;
			this.btnComandos.Text = "Link Nuovo";
			this.btnComandos.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnComandos.Visible = false;
			this.btnComandos.Click += new System.EventHandler(this.btnComandos_Click);
			// 
			// mnuAcciones
			// 
			this.mnuAcciones.MenuItems.AddRange(new System.Windows.Forms.MenuItem[] {
																																								this.mnuAccionesCopiarUltimoMes,
																																								this.mnuAccionesAgregarNuevoLocal,
																																								this.menuItem1,
																																								this.mnuAccionesActualizar});
			// 
			// mnuAccionesCopiarUltimoMes
			// 
			this.mnuAccionesCopiarUltimoMes.Index = 0;
			this.mnuAccionesCopiarUltimoMes.Text = "Código QR Descarga Aplicativo";
			this.mnuAccionesCopiarUltimoMes.Click += new System.EventHandler(this.mnuAccionesCopiarUltimoMes_Click);
			// 
			// mnuAccionesAgregarNuevoLocal
			// 
			this.mnuAccionesAgregarNuevoLocal.Index = 1;
			this.mnuAccionesAgregarNuevoLocal.Text = "Código QR Enrolar";
			this.mnuAccionesAgregarNuevoLocal.Click += new System.EventHandler(this.mnuAccionesAgregarNuevoLocal_Click);
			// 
			// menuItem1
			// 
			this.menuItem1.Index = 2;
			this.menuItem1.Text = "Código QR Huawei Enrolar";
			this.menuItem1.Click += new System.EventHandler(this.menuItem1_Click);
			// 
			// mnuAccionesActualizar
			// 
			this.mnuAccionesActualizar.Index = 3;
			this.mnuAccionesActualizar.Text = "Dispositivos Secured by Knox";
			this.mnuAccionesActualizar.Click += new System.EventHandler(this.mnuAccionesActualizar_Click);
			// 
			// ultraDataSource48
			// 
			ultraDataColumn163.DataType = typeof(int);
			this.ultraDataSource48.Band.Columns.AddRange(new object[] {
																																	ultraDataColumn163,
																																	ultraDataColumn164});
			// 
			// lblFechaActualizacion
			// 
			this.lblFechaActualizacion.AutoSize = true;
			this.lblFechaActualizacion.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.lblFechaActualizacion.ForeColor = System.Drawing.Color.Firebrick;
			this.lblFechaActualizacion.Location = new System.Drawing.Point(488, 56);
			this.lblFechaActualizacion.Name = "lblFechaActualizacion";
			this.lblFechaActualizacion.Size = new System.Drawing.Size(0, 17);
			this.lblFechaActualizacion.TabIndex = 1051;
			this.lblFechaActualizacion.TextAlign = System.Drawing.ContentAlignment.BottomRight;
			// 
			// btnVerificacion
			// 
			this.btnVerificacion.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			this.btnVerificacion.Image = ((System.Drawing.Image)(resources.GetObject("btnVerificacion.Image")));
			this.btnVerificacion.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.btnVerificacion.Location = new System.Drawing.Point(1040, 488);
			this.btnVerificacion.Name = "btnVerificacion";
			this.btnVerificacion.Size = new System.Drawing.Size(80, 23);
			this.btnVerificacion.TabIndex = 1052;
			this.btnVerificacion.Text = "Verificacion Terrena";
			this.btnVerificacion.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnVerificacion.Visible = false;
			this.btnVerificacion.Click += new System.EventHandler(this.btnVerificacion_Click);
			// 
			// frmCre_Solicitudes
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 13);
			this.ClientSize = new System.Drawing.Size(1135, 525);
			this.Controls.Add(this.btnVerificacion);
			this.Controls.Add(this.lblFechaActualizacion);
			this.Controls.Add(this.btnComandos);
			this.Controls.Add(this.btnIdValidator);
			this.Controls.Add(this.txtIdGarante);
			this.Controls.Add(this.btnImprimir);
			this.Controls.Add(this.btnGuardar);
			this.Controls.Add(this.btnSalir);
			this.Controls.Add(this.btnConsultar);
			this.Controls.Add(this.txtNegocioIdCre_DatosUbicacion);
			this.Controls.Add(this.txtClienteIdCre_EmpresasTrabajo);
			this.Controls.Add(this.txtEmpresaIdCre_DatosUbicacion);
			this.Controls.Add(this.txtConyugueIdCre_DatosGenerales);
			this.Controls.Add(this.txtConyugueIdCre_DatosUbicacion);
			this.Controls.Add(this.txtClienteIdCre_Negocio);
			this.Controls.Add(this.txtIdCre_Conyugue);
			this.Controls.Add(this.txtIdCre_Solicitud);
			this.Controls.Add(this.txtClienteIdCre_DatosGenerales);
			this.Controls.Add(this.txtClienteIdCre_DatosUbicacion);
			this.Controls.Add(this.txtEstado);
			this.Controls.Add(this.lblEstado);
			this.Controls.Add(this.lblFISL);
			this.Controls.Add(this.tbBasico);
			this.Controls.Add(this.btnGarante);
			this.Controls.Add(this.chkGarante);
			this.Controls.Add(this.label16);
			this.Controls.Add(this.lblEntidad);
			this.Controls.Add(this.label5);
			this.Controls.Add(this.label4);
			this.Controls.Add(this.label3);
			this.Controls.Add(this.label2);
			this.Controls.Add(this.label1);
			this.Controls.Add(this.lblBodega);
			this.Controls.Add(this.cmbBodega);
			this.Controls.Add(this.cmbEntidad);
			this.Controls.Add(this.txtSnombre);
			this.Controls.Add(this.txtPnombre);
			this.Controls.Add(this.txtApellidoM);
			this.Controls.Add(this.txtApellidoP);
			this.Controls.Add(this.dtFecha);
			this.Controls.Add(this.cmbTipoRuc);
			this.Controls.Add(this.txtDocumento);
			this.KeyPreview = true;
			this.Name = "frmCre_Solicitudes";
			this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
			this.Text = "Solicitudes de Crédito Directo ";
			this.Load += new System.EventHandler(this.frmCre_Solicitudes_Load);
			this.ultraTabPageControl15.ResumeLayout(false);
			((System.ComponentModel.ISupportInitialize)(this.txtVTSecundaria)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtVTPRincipal)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtVTLongitud)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtVTPersonaEntrevistada)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtVTReferencia)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtVTGps)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtVTLatitud)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtVTTelefono)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtVTActividadLaboral)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtVTIngresos)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtVTTiempoYear)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtVTTiempoMeses)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.opdtVTTipoTrabajo)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraGrid5)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource42)).EndInit();
			this.ultraTabPageControl9.ResumeLayout(false);
			((System.ComponentModel.ISupportInitialize)(this.txtVDLatitud)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtVDLongitud)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtVDCCelArr)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtVDGPS)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtVDSecundaria)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtVDPrincipal)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtVDVecino)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtVDObservacion)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtVDentrevistada)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtVDPuntoReferencia)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.opdtVDCobertura)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.opdtVDAcceso)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.opdtVDPropiedad)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.opdtVDZonaVivienda)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.opdtVDEstadoVivienda)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.opdtVDVivienda)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtVDTiempoMeses)).EndInit();
			this.ultraTabPageControl1.ResumeLayout(false);
			this.groupBox2.ResumeLayout(false);
			((System.ComponentModel.ISupportInitialize)(this.txtObsActividadEconomica)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbActividadEconomica)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource40)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbSituacionLaboral)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource30)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbProfesion)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource29)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbNivelEducacion)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource28)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbSexo)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource24)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtValorVehiculo)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtModelo)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtVehiculoMarca)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtDependientes)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbEstadoCivil)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource11)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbCantonNacimiento)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource2)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbProvinciaNAcimiento)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource1)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.dtFechaNacimiento)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbNacionalidad)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource26)).EndInit();
			this.groupBox1.ResumeLayout(false);
			((System.ComponentModel.ISupportInitialize)(this.txtCelArrendador)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbTiempoVivienda)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource49)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtDDValorInmueble)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbDDTipoVivienda)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource31)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.dtTiempoVivienda)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbDDCiudadInmueble)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource38)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbDDInmueble)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource27)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtTelfArrendador)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNombreArrendador)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtDDReferenciasU)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtDDEmail)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtDDCelular)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtDDTelefono2)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtDDTelefono1)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtDDCalleSecundaria)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtDDNumeroCasa)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtDDCallaPrincipal)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbDDBarrio)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource4)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbDDParroquia)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource3)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbDDCanton)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbDDProvincia)).EndInit();
			this.tpComplemento.ResumeLayout(false);
			this.groupBox3.ResumeLayout(false);
			((System.ComponentModel.ISupportInitialize)(this.cmbDPCActividadE)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtDPCsegundoN)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtDPCPrimerN)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtDPCApellidoM)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtDPCApellidoP)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbDPCTipodocumento)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtDPCdocumento)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbDPCSituacionL)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbDPCProfesion)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbDPCNivelEducacion)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource33)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbDPCSexo)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbDPCCantonN)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbDPCProvinciaN)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.dtDPCFnacimiento)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbDPCNacionalidad)).EndInit();
			this.groupBox4.ResumeLayout(false);
			((System.ComponentModel.ISupportInitialize)(this.txtDPCTPasivos)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtDPCTActivos)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtDPCTGastos)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtDPCTIngresos)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbDPCTCargo)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource36)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbDPCTActividadEmpresa)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource35)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtDPCTNombreEmpresa)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtDPCTReferencia)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtDPCTEmail)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtDPCTCelularC)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtDPCTExt)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtDPCTtelefono)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtDPCTCalleSecundaria)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtDPCTNumeroCasa)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtDPCTCallePrincipal)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbDPCTBarrio)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbDPCTParroquia)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbDPCTCanton)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbDPCTProvincia)).EndInit();
			this.ultraTabPageControl5.ResumeLayout(false);
			((System.ComponentModel.ISupportInitialize)(this.cmbtipoCuenta)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource16)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbTipoBanco)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource15)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbtipoReferencia)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource14)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultgconsulta)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource13)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDropDown1)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbProvincia)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbParentesco)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource9)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.uGridRefernciaBancaria)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource17)).EndInit();
			this.ultraTabPageControl7.ResumeLayout(false);
			this.groupBox5.ResumeLayout(false);
			((System.ComponentModel.ISupportInitialize)(this.cmbTiempo)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNotas)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNegocioTelefono2)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNegocioemail)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbNegocioBarrio)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbNegocioParroquia)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbNegocioCanton)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbNegocioProvincia)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNEgocioPasivos)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNegocioActivos)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNEgocioGastos)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNEgocioIngresos)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNegocioReferncia)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNEgocioCelular)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNegocioTelefono1)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNegocioCalleSecundaria)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNegocioNumeroCasa)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNegocioCallePrincipal)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNegocioEmpleados)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNegocioMetros)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNombreDelNegocio)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNegocioRuc)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbNegciotipo)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource37)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbNEgocioActivdadEconomica)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource25)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.dtpNegocioFechaInicio)).EndInit();
			this.ultraTabPageControl2.ResumeLayout(false);
			this.groupBox6.ResumeLayout(false);
			((System.ComponentModel.ISupportInitialize)(this.cmbTiempoDependiente)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbDependienteTipoempresa)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource18)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbDependienteTipoSueldo)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource20)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbDependienteCargo)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource21)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtDependienteDepartamento)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtDependienteEXT)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtDependienteemail)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbDependienteBarrio)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbDependienteParroquia)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbDependienteCanton)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbDependienteProvincia)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtDependientePasivos)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtDependienteActivos)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtDependienteGastos)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtDependienteIngresos)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtDependienteReferencia)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtDependienteCelular)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtDependienteTelefono)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtDependienteCalleSecundaria)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtDependienteNumeroCasa)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtDependienteCallePrinciapl)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtDependienteDiasPago)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtDependienteEmpresa)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtDependienteRuc)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbDependienteTipocontrato)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource19)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.dtDependienteFechaIngreso)).EndInit();
			this.ultraTabPageControl3.ResumeLayout(false);
			this.groupBox8.ResumeLayout(false);
			((System.ComponentModel.ISupportInitialize)(this.utgHistorialCredito)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource41)).EndInit();
			this.groupBox7.ResumeLayout(false);
			((System.ComponentModel.ISupportInitialize)(this.txtCarteraCastigada)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtResultado)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtSegmentacion)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtRangoIngresos)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtTotalVencido)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtCapacidadPago)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtCreditoScoreInclusion)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtCreditoScoreSobreEndeudamiento)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtCreditoObservaciones)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtCreditoScore)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.dtCreditoFechaSolicitud)).EndInit();
			this.ultraTabPageControl6.ResumeLayout(false);
			((System.ComponentModel.ISupportInitialize)(this.cmbdetallecliente)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource44)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbidTipoCliente)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource43)).EndInit();
			this.groupBox9.ResumeLayout(false);
			((System.ComponentModel.ISupportInitialize)(this.txtObservaciones)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbCreditoCalificacion)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource22)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtCreditoCuotaAsiganada)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmdEstadoSolicitud)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource47)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.gdObservaciones)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource46)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtCupo)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbCreditoEstado)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource23)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbSituacionLaboral2)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource45)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraCombo2)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraCombo3)).EndInit();
			this.ultraTabPageControl4.ResumeLayout(false);
			((System.ComponentModel.ISupportInitialize)(this.ultraGrid1)).EndInit();
			this.ultraTabPageControl8.ResumeLayout(false);
			((System.ComponentModel.ISupportInitialize)(this.ultraTabControl1)).EndInit();
			this.ultraTabControl1.ResumeLayout(false);
			((System.ComponentModel.ISupportInitialize)(this.utgVerificacionTerrena)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource50)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource5)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource34)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource39)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource32)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtDocumento)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.dtFecha)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtSnombre)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtPnombre)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtApellidoM)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtApellidoP)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbBodega)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbEntidad)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource10)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource7)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource8)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource6)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cdsSeteoF)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.tbBasico)).EndInit();
			this.tbBasico.ResumeLayout(false);
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource12)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.cmbTipoRuc)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtNegocioIdCre_DatosUbicacion)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtClienteIdCre_EmpresasTrabajo)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtEmpresaIdCre_DatosUbicacion)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtConyugueIdCre_DatosGenerales)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtConyugueIdCre_DatosUbicacion)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtClienteIdCre_Negocio)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIdCre_Conyugue)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIdCre_Solicitud)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtClienteIdCre_DatosGenerales)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtClienteIdCre_DatosUbicacion)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtEstado)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.txtIdGarante)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.ultraDataSource48)).EndInit();
			this.ResumeLayout(false);

		}
		#endregion

		private void cdsSeteoF_BeforeFill(object sender, C1.Data.FillEventArgs e)
		{
			Funcion miFuncion = new Funcion();
			cdsSeteoF.Schema.Connections[0].ConnectionString = miFuncion.AccesoBase(MenuLatinium.stDirFacturacion);	
		}

		private void txtDocumento_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if(e.KeyCode == Keys.Enter) this.txtApellidoP.Focus();
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
			if(e.KeyCode == Keys.Enter) this.cmbEntidad.Focus();
		}

		private void cmbEntidad_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if(e.KeyCode == Keys.Enter) this.cmbBodega.Focus();
		}

		private void cmbBodega_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
//			if(e.KeyCode == Keys.Enter) this.cmbTipoTarjeta.Focus();
		}

		private void cmbTipoTarjeta_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if(e.KeyCode == Keys.Enter)
			{
				this.tbBasico.Tabs[0].Selected = true;
				this.cmbNacionalidad.Focus ();
			} 
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
		private Acceso miAcceso;

		private void UnloadMe()
		{
			this.Close();
		}


		private void frmCre_Solicitudes_Load(object sender, System.EventArgs e)
		{
			miAcceso = new Acceso(cdsSeteoF, "083404");

			if (!Funcion.Permiso("984", cdsSeteoF))
			{				
				MessageBox.Show("No puede ingresar a Solicitudes", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				BeginInvoke(new MethodInvoker(UnloadMe));
				return;
			}
			decimal dSueldoBasico = Funcion.decEscalarSQL(cdsSeteoF, "SELECT SueldoBasico FROM Seteo");

			 bCarga = true;
			#region Carga Combo

			#region Carga Datos Generales
			this.dtFecha.Value = DateTime.Today;
		

			this.cmbTipoRuc.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "Exec Listas 4, 0, 0");
			this.cmbEntidad.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "Exec Listas 95, 0, 0");
			this.cmbBodega.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "Exec Listas 7, 0, 0");
//			this.cmbTipoTarjeta.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "Exec Listas 6, 0, 0");
			this.cmbTipoRuc.Value = 1;
			//this.cmbTipoRuc.Value = 0;

			#endregion Carga Datos Generales

			#region Datos Cliente

			#region Carga Datos Personales

			this.cmbNacionalidad.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "Exec Listas 11, 0");
			this.cmbSexo.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "Exec Listas 10, 0");
			this.cmbProvinciaNAcimiento.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "Exec Listas 0, 0, 0");
			this.cmbEstadoCivil.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "Exec Listas 12, 0, 0");
			this.cmbNivelEducacion.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "Exec Listas 13, 0, 0");
			this.cmbProfesion.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "Exec Listas 14, 0, 0");
			

			

			#endregion Carga Datos Personales

			#region Carga Datos Domicilio

			this.cmbDDProvincia.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "Exec Listas 0, 0, 0");
			this.cmbDDTipoVivienda.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "Exec Listas 17, 0, 0");
			this.cmbDDCiudadInmueble.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "Exec Listas 26, 0, 0");
			this.cmbDDInmueble.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "Exec Listas 25, 0, 0");
			
			#endregion Carga Datos Domicilio

			#endregion Datos Cliente

			#region Datos Conyuge
			#region Datos Personales
			this.cmbDPCTipodocumento.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "Exec Listas 4, 0, 0");
			this.cmbDPCNacionalidad.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "Exec Listas 11, 0");
			this.cmbDPCSexo.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "Exec Listas 10, 0, 0");
			this.cmbDPCProvinciaN.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "Exec Listas 0, 0, 0");
			this.cmbDPCNivelEducacion.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "Exec Listas 13, 0, 0");
			this.cmbDPCProfesion.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "Exec Listas 14, 0, 0");
			this.cmbDPCSituacionL.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "Exec Listas 18, 0, 0");
//			this.cmbDPCActividadE.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "Exec Listas 16, 0, 0");
			#endregion Datos Personales
			
			
			#region datos de Trabajo 

			this.cmbDPCTActividadEmpresa.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "Exec Listas 27, 0, 0");
			this.cmbDPCTCargo.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "Exec Listas 20, 0, 0");
			this.cmbDPCTProvincia.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "Exec Listas 0, 0, 0");
		
			#region datos de Trabajo 
			 
							
								
			#endregion Datos Conyuge

			#endregion Carga combo

			#endregion Datos Conyuge

			

			#endregion Carga Combo
			this.ultgconsulta.DataSource = Funcion.dvProcedimiento(cdsSeteoF, string.Format("Exec Cre_ConsultaReferencias {0},{1}", 0,1));
			this.uGridRefernciaBancaria.DataSource = Funcion.dvProcedimiento(cdsSeteoF, string.Format("Exec Cre_ConsultaReferencias {0},{1}", 0,2));
			this.ultraGrid1.DataSource = Funcion.dvProcedimiento(cdsSeteoF, string.Format("Exec Cre_ConsultaReferencias {0},{1}", 0,3));
			this.cmbParentesco.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "Exec Listas 29, 0, 0");
			this.cmbProvincia.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "Exec Listas 0, 0, 0");
			this.cmbtipoReferencia.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "Exec Listas 32, 0, 0");
			this.cmbTipoBanco.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "Exec Listas 31, 0, 0");
			this.cmbtipoCuenta.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "Exec Listas 30, 0, 0");

			#region Negocio
			this.cmbNegciotipo.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "Exec Listas 17, 0, 0");
			this.cmbNEgocioActivdadEconomica.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "Exec Listas 21, 0, 0");
			this.cmbNegocioProvincia.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "Exec Listas 0, 0, 0");
			
      #endregion Negocio
      this.cmbDependienteTipoempresa.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "Exec Listas 28, 0, 0");
			this.cmbDependienteTipocontrato.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "Exec Listas 22, 0, 0");
			this.cmbDependienteTipoSueldo.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "Exec Listas 23, 0, 0");
      this.cmbDependienteCargo.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "Exec Listas 20, 0, 0");
			this.cmbDependienteProvincia.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "Exec Listas 0, 0, 0");
			this.cmbCreditoCalificacion.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "Exec Listas 8, 0, 0");
			this.cmbCreditoEstado.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "Exec Listas 9, 0, 0");

			this.cmbidTipoCliente.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "Exec Listas 78, 0, 0");
			this.cmbdetallecliente.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "Exec Listas 79, 0, 0");
			this.cmbSituacionLaboral2.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "Exec Listas 87, 0, 0");

			this.cmbTiempo.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "Exec Listas 94, 0, 0");
			this.cmbTiempoDependiente.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "Exec Listas 94, 0, 0");	
			this.cmbTiempoVivienda.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "Exec Listas 94, 0, 0");

//			this.cmbClienteFuenteDeIngresos.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "Exec Listas 15, 0, 0");
//			this.cmbSituacionLaboral.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "Exec Listas 66, 0, 0");
			this.cmbSituacionLaboral.DataSource = Funcion.dvProcedimiento(cdsSeteoF, string.Format("Exec Listas 66, {0}, 0",4));
			this.cmbDPCSituacionL.DataSource = Funcion.dvProcedimiento(cdsSeteoF, string.Format("Exec Listas 66, {0}, 0",4));

			this.tbBasico.Tabs[0].Selected = true;
			HabilitaControles ();
//			HabilitaConyuge ();
//			HabilitaDatosNegocio ();
//			Habilita_DatosDependiente();
			
			if (miAcceso.BHabilitaControles)
			{
				Habilita_InformacionCredito ();
				
			}
			if (!miAcceso.BAsignaCobrador) 
			{
				this.cmbidTipoCliente.Enabled = false;
			}

			if (miAcceso.BExportar) this.btnVerificacion.Visible = true;
				//this.dtpNegocioFechaInicio.CalendarInfo.MaxDate = DateTime.Today;
			
			FuncionesProcedimientos.RetornaFechaServidor(this.dtpNegocioFechaInicio, cdsSeteoF, false);

			FuncionesProcedimientos.RetornaFechaServidor(this.dtDependienteFechaIngreso, cdsSeteoF, false);
			FuncionesProcedimientos.RetornaFechaServidor(this.dtTiempoVivienda, cdsSeteoF, false);
			this.txtValorVehiculo.Value = 0;
			
			bCarga = false;

		}

		private void cmbProvinciaNAcimiento_ValueChanged(object sender, System.EventArgs e)
		{
			if (this.cmbProvinciaNAcimiento.ActiveRow != null)
			{
				this.cmbCantonNacimiento.DataSource = Funcion.dvProcedimiento(cdsSeteoF, string.Format("Exec Listas 1, {0}, 0", (int) this.cmbProvinciaNAcimiento.Value));
			}
		}

		private void groupBox1_Enter(object sender, System.EventArgs e)
		{
		
		}

		private void ultraTabPageControl1_Paint(object sender, System.Windows.Forms.PaintEventArgs e)
		{
		
		}

		private void cmbDDProvincia_ValueChanged(object sender, System.EventArgs e)
		{
			if (this.cmbDDProvincia.ActiveRow != null)
			{
				this.cmbDDCanton.DataSource = Funcion.dvProcedimiento(cdsSeteoF, string.Format("Exec Listas 1, {0}, 0", (int) this.cmbDDProvincia.Value));
				
			}
		}

		private void cmbDDCanton_ValueChanged(object sender, System.EventArgs e)
		{
			if (this.cmbDDCanton.ActiveRow != null)
			{
				this.cmbDDParroquia.DataSource = Funcion.dvProcedimiento(cdsSeteoF, string.Format("Exec Listas 2, {0}, 0", (int) this.cmbDDCanton.Value));
				
			}
		}

		private void cmbDDParroquia_ValueChanged(object sender, System.EventArgs e)
		{
			if (this.cmbDDParroquia.ActiveRow != null)
			{
				this.cmbDDBarrio.DataSource = Funcion.dvProcedimiento(cdsSeteoF, string.Format("Exec Listas 3, {0}, 0", (int) this.cmbDDParroquia.Value));
			}
		}

		private void cmbNacionalidad_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if(e.KeyCode == Keys.Enter) this.dtFechaNacimiento.Focus();
		}

		private void dtFechaNacimiento_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if(e.KeyCode == Keys.Enter) this.cmbSexo.Focus();
		}

		private void cmbSexo_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if(e.KeyCode == Keys.Enter) this.cmbProvinciaNAcimiento.Focus();
		}

		private void cmbProvinciaNAcimiento_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if(e.KeyCode == Keys.Enter) this.cmbCantonNacimiento.Focus();
		}

		private void cmbCantonNacimiento_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if(e.KeyCode == Keys.Enter) this.cmbEstadoCivil.Focus();
		}

		private void cmbEstadoCivil_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if(e.KeyCode == Keys.Enter) this.txtDependientes.Focus();
		}

		private void txtDependientes_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if(e.KeyCode == Keys.Enter) this.cmbNivelEducacion.Focus();
		}

		private void cmbNivelEducacion_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if(e.KeyCode == Keys.Enter) this.cmbProfesion.Focus();
		}

		private void cmbProfesion_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if(e.KeyCode == Keys.Enter)
			{
				if ((int)this.cmbEntidad.Value == 2)
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
			if(e.KeyCode == Keys.Enter) this.cmbActividadEconomica.Focus();

		}

		private void cmbActividadEconomica_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			//if(e.KeyCode == Keys.Enter) this.txtVehiculoMarca.Focus();
		}

		private void txtVehiculoMarca_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			//if(e.KeyCode == Keys.Enter) this.txtModelo.Focus();
		}

		private void txtModelo_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
//			if(e.KeyCode == Keys.Enter)
//			{
//					this.txtValorVehiculo.Focus();
//				this.txtValorVehiculo.SelectAll ();
//			}
		}

		private void txtValorVehiculo_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			//if(e.KeyCode == Keys.Enter) this.cmbDDProvincia.Focus();
		}

		private void cmbDDProvincia_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if(e.KeyCode == Keys.Enter) this.cmbDDCanton.Focus();
		}

		private void cmbDDCanton_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if(e.KeyCode == Keys.Enter) this.cmbDDParroquia.Focus();
		}

		private void cmbDDParroquia_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if(e.KeyCode == Keys.Enter) this.cmbDDBarrio.Focus();
		}

		private void cmbDDBarrio_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if(e.KeyCode == Keys.Enter) this.txtDDEmail.Focus();
		}

		private void txtDDEmail_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if(e.KeyCode == Keys.Enter) this.txtDDCallaPrincipal.Focus();
		}

		private void txtDDCallaPrincipal_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if(e.KeyCode == Keys.Enter) this.txtDDNumeroCasa.Focus();
		}

		private void txtDDNumeroCasa_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if(e.KeyCode == Keys.Enter) this.txtDDCalleSecundaria.Focus();
		}

		private void txtDDCalleSecundaria_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if(e.KeyCode == Keys.Enter) this.txtDDReferenciasU.Focus();
		}

		private void txtDDTelefono1_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if(e.KeyCode == Keys.Enter) this.txtDDTelefono2.Focus();
		}

		private void txtDDTelefono2_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if(e.KeyCode == Keys.Enter) this.txtDDCelular.Focus();
		}

		private void txtDDCelular_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if(e.KeyCode == Keys.Enter) this.cmbDDTipoVivienda.Focus();
		}

		private void txtDDReferenciasU_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if(e.KeyCode == Keys.Enter) this.txtDDTelefono1.Focus();
		}

		private void cmbDDTipoVivienda_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			//if(e.KeyCode == Keys.Enter) this.dtTiempoVivienda.Focus();
			if(e.KeyCode == Keys.Enter) this.cmbTiempoVivienda.Focus();
		}

		private void dtTiempoVivienda_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			
		}

		private void txtNombreArrendador_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if(e.KeyCode == Keys.Enter) this.txtTelfArrendador.Focus();
		}

		private void txtTelfArrendador_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if(e.KeyCode == Keys.Enter) 
			{
				this.tbBasico.Tabs[1].Selected = true;
				this.txtDPCdocumento.Focus ();
				if (!this.cmbDPCTipodocumento.Enabled) 
				{
					this.tbBasico.Tabs[2].Selected = true;
				}
			}
		}

		private void cmbDDInmueble_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if(e.KeyCode == Keys.Enter) this.cmbDDCiudadInmueble.Focus();
		}

		private void cmbDDCiudadInmueble_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			
			if(e.KeyCode == Keys.Enter) this.txtDDValorInmueble.Focus();
				this.txtDDValorInmueble.SelectAll ();
		}

		private void HabilitaControles ()
		{
			#region Datos cliente

			this.txtApellidoP.Enabled =  true;
			this.txtApellidoM.Enabled = true;
			this.txtPnombre.Enabled = true;
			this.txtSnombre.Enabled = true;
			//this.cmbEntidad.Enabled = true;
			this.cmbBodega.Enabled = true;
//			this.cmbTipoTarjeta.Enabled = true;
			this.chkGarante.Enabled = true;
//			this.cmbClienteFuenteDeIngresos.Enabled = true;

			this.cmbNacionalidad.Enabled = true;
			//this.dtFechaNacimiento.Enabled = true;
			this.cmbSexo.Enabled = true;
			this.cmbProvinciaNAcimiento.Enabled = true;
			this.cmbCantonNacimiento.Enabled = true;
			this.cmbEstadoCivil.Enabled = true;
			this.txtDependientes.Enabled = true;
			this.cmbNivelEducacion.Enabled = true;
			this.cmbProfesion.Enabled = true;
			this.cmbSituacionLaboral.Enabled = true;
			this.cmbActividadEconomica.Enabled = true;
			//this.txtVehiculoMarca.Enabled = true;
			//this.txtModelo.Enabled = true;
			//this.txtValorVehiculo.Enabled = true;
			this.chkSeparacionBienes.Enabled = true;
			this.chkHipoteca.Enabled = true;
			
			this.cmbDDProvincia.Enabled = true;
			this.cmbDDCanton.Enabled = true;
			this.cmbDDParroquia.Enabled = true;
			this.cmbDDBarrio.Enabled = true;
			this.txtDDEmail.Enabled = true;
			this.txtDDCallaPrincipal.Enabled = true;
			this.txtDDNumeroCasa.Enabled = true;
			this.txtDDCalleSecundaria.Enabled = true;
			this.txtDDTelefono1.Enabled = true;
			this.txtDDTelefono2.Enabled = true;
			this.txtDDCelular.Enabled = true;
			this.txtDDReferenciasU.Enabled = true;
			this.cmbDDTipoVivienda.Enabled = true;
			//this.dtTiempoVivienda.Enabled = true;
			this.cmbTiempoVivienda.Enabled = true;
			this.txtObsActividadEconomica.Enabled = true;
//			this.txtTelfArrendador.Enabled = true;
//			this.cmbDDInmueble.Enabled = true;
//			this.cmbDDCiudadInmueble.Enabled = true;
//			this.txtDDValorInmueble.Enabled = true;
			
			#endregion Datos cliente
		}

		private void Bloque_Controles ()
		{
			#region Datos cliente

			this.txtApellidoP.Enabled =  false;
			this.txtApellidoM.Enabled = false;
			this.txtPnombre.Enabled = false;
			this.txtSnombre.Enabled = false;
			this.cmbEntidad.Enabled = false;
			this.cmbBodega.Enabled = false;
			//			this.cmbTipoTarjeta.Enabled = true;
			this.chkGarante.Enabled = false;
			//			this.cmbClienteFuenteDeIngresos.Enabled = true;

			this.cmbNacionalidad.Enabled = false;
			this.dtFechaNacimiento.Enabled = false;
			this.cmbSexo.Enabled = false;
			this.cmbProvinciaNAcimiento.Enabled = false;
			this.cmbCantonNacimiento.Enabled = false;
			this.cmbEstadoCivil.Enabled = false;
			this.txtDependientes.Enabled = false;
			this.cmbNivelEducacion.Enabled = false;
			this.cmbProfesion.Enabled = false;
			this.cmbSituacionLaboral.Enabled = false;
			this.cmbActividadEconomica.Enabled = false;
			//this.txtVehiculoMarca.Enabled = false;
			//this.txtModelo.Enabled = false;
			//this.txtValorVehiculo.Enabled = false;
			this.chkSeparacionBienes.Enabled = false;
			this.chkHipoteca.Enabled = false;
			
			this.cmbDDProvincia.Enabled = false;
			this.cmbDDCanton.Enabled =false;
			this.cmbDDParroquia.Enabled = false;
			this.cmbDDBarrio.Enabled = false;
			this.txtDDEmail.Enabled = false;
			this.txtDDCallaPrincipal.Enabled = false;
			this.txtDDNumeroCasa.Enabled = false;
			this.txtDDCalleSecundaria.Enabled = false;
			this.txtDDTelefono1.Enabled = false;
			this.txtDDTelefono2.Enabled = false;
			this.txtDDCelular.Enabled = false;
			this.txtDDReferenciasU.Enabled = false;
			this.cmbDDTipoVivienda.Enabled = false;
			//this.dtTiempoVivienda.Enabled = false;
			this.cmbTiempoVivienda.Enabled = false;
			this.txtObsActividadEconomica.Enabled = false;
			//			this.txtNombreArrendador.Enabled = true;
			//			this.txtTelfArrendador.Enabled = true;
			//			this.cmbDDInmueble.Enabled = true;
			//			this.cmbDDCiudadInmueble.Enabled = true;
			//			this.txtDDValorInmueble.Enabled = true;
			
			#endregion Datos cliente
		}

		private void HabilitaConyuge ()
		{
			#region
			this.txtDPCdocumento.Enabled = true;
			this.txtDPCApellidoP.Enabled = true;
			this.txtDPCApellidoM.Enabled = true;
			this.txtDPCPrimerN.Enabled = true;
			this.txtDPCsegundoN.Enabled = true;
			this.cmbDPCNacionalidad.Enabled = true;
			this.dtDPCFnacimiento.Enabled = true;
			this.cmbDPCSexo.Enabled = true;
			this.cmbDPCProvinciaN.Enabled = true;
			this.cmbDPCCantonN.Enabled = true;
			this.cmbDPCNivelEducacion.Enabled = true;
			this.cmbDPCProfesion.Enabled = true;
			this.cmbDPCSituacionL.Enabled = true;
			this.cmbDPCActividadE.Enabled = true;

			this.txtDPCTNombreEmpresa.Enabled = true;
			this.cmbDPCTActividadEmpresa.Enabled = true;
			this.cmbDPCTCargo.Enabled = true;
			this.cmbDPCTProvincia.Enabled = true;
			this.cmbDPCTCanton.Enabled = true;
			this.cmbDPCTParroquia.Enabled = true;
			this.cmbDPCTBarrio.Enabled = true;
			this.txtDPCTEmail.Enabled = true;
			this.txtDPCTCallePrincipal.Enabled = true;
			this.txtDPCTNumeroCasa.Enabled = true;
			this.txtDPCTCalleSecundaria.Enabled = true;
			this.txtDPCTtelefono.Enabled = true;
			this.txtDPCTExt.Enabled = true;
			this.txtDPCTCelularC.Enabled = true;
			this.txtDPCTReferencia.Enabled = true;
			this.txtDPCTIngresos.Enabled = true;
			this.txtDPCTGastos.Enabled = true;
			//this.txtDPCTActivos.Enabled = true;
			//this.txtDPCTPasivos.Enabled = true;
			#endregion

		}
		private void DesabilitaConyuge ()
		{
			#region DesabilitaConyuge
			this.txtDPCdocumento.Enabled = false;
			this.txtDPCApellidoP.Enabled = false;
			this.txtDPCApellidoM.Enabled = false;
			this.txtDPCPrimerN.Enabled = false;
			this.txtDPCsegundoN.Enabled = false;
			this.cmbDPCNacionalidad.Enabled = false;
			this.dtDPCFnacimiento.Enabled = false;
			this.cmbDPCSexo.Enabled = false;
			this.cmbDPCProvinciaN.Enabled = false;
			this.cmbDPCCantonN.Enabled = false;
			this.cmbDPCNivelEducacion.Enabled = false;
			this.cmbDPCProfesion.Enabled = false;
			this.cmbDPCSituacionL.Enabled = false;
			this.cmbDPCActividadE.Enabled = false;
			this.txtDPCTNombreEmpresa.Enabled = false;
			this.cmbDPCTActividadEmpresa.Enabled = false;
			this.cmbDPCTCargo.Enabled = false;
			this.cmbDPCTProvincia.Enabled = false;
			this.cmbDPCTCanton.Enabled = false;
			this.cmbDPCTParroquia.Enabled = false;
			this.cmbDPCTBarrio.Enabled = false;
			this.txtDPCTEmail.Enabled = false;
			this.txtDPCTCallePrincipal.Enabled = false;
			this.txtDPCTNumeroCasa.Enabled = false;
			this.txtDPCTCalleSecundaria.Enabled = false;
			this.txtDPCTtelefono.Enabled = false;
			this.txtDPCTExt.Enabled = false;
			this.txtDPCTCelularC.Enabled = false;
			this.txtDPCTReferencia.Enabled = false;
			this.txtDPCTIngresos.Enabled =false;
			this.txtDPCTGastos.Enabled = false;
			this.txtDPCTActivos.Enabled = false;
			this.txtDPCTPasivos.Enabled = false;
			#endregion DesabilitaConyuge
		
		}

		private void HabilitaDatosNegocio ()
		{
			#region Negocio
			this.txtNegocioRuc.Enabled = true;
			this.txtNombreDelNegocio.Enabled = true;
			//this.dtpNegocioFechaInicio.Enabled = true;
			this.cmbTiempo.Enabled =true;
			this.txtNegocioMetros.Enabled = true;
			this.cmbNegciotipo.Enabled = true;
			this.txtNegocioEmpleados.Enabled = true;
			this.cmbNegciotipo.Enabled = true;
			this.cmbNEgocioActivdadEconomica.Enabled = true;
			this.cmbNegocioProvincia.Enabled = true;
			this.cmbNegocioCanton.Enabled = true;
			this.cmbNegocioParroquia.Enabled = true;
			this.cmbNegocioBarrio.Enabled = true;
			this.txtNegocioemail.Enabled = true;
			this.txtNegocioCallePrincipal.Enabled = true;
			this.txtNegocioNumeroCasa.Enabled = true;
			this.txtNegocioCalleSecundaria.Enabled = true;
			this.txtNegocioTelefono1.Enabled = true;
			this.txtNegocioTelefono2.Enabled = true;
			this.txtNEgocioCelular.Enabled = true;
			this.txtNegocioReferncia.Enabled = true;
			this.txtNEgocioIngresos.Enabled = true;
			this.txtNEgocioGastos.Enabled = true;
			//this.txtNegocioActivos.Enabled = true;
			//this.txtNEgocioPasivos.Enabled = true;
			this.chkNegocioObligadoaLlevar.Enabled = true;
			this.chkAfiliadoTributario.Enabled = true;
			this.txtNotas.Enabled=true;
			#endregion Negocio

		}

		private void DesabilitaDatosNegocio ()
		{
			#region NegocioD
			this.txtNegocioRuc.Enabled = false;
			this.txtNombreDelNegocio.Enabled = false;
			//this.dtpNegocioFechaInicio.Enabled = false;
			this.cmbTiempo.Enabled =false;
			this.txtNegocioMetros.Enabled = false;
			this.txtNegocioEmpleados.Enabled = false;
			this.cmbNegciotipo.Enabled = false;
			this.cmbNEgocioActivdadEconomica.Enabled = false;
			this.cmbNegocioProvincia.Enabled = false;
			this.cmbNegocioCanton.Enabled = false;
			this.cmbNegocioParroquia.Enabled = false;
			this.cmbNegocioBarrio.Enabled = false;
			this.txtNegocioemail.Enabled = false;
			this.txtNegocioCallePrincipal.Enabled = false;
			this.txtNegocioNumeroCasa.Enabled = false;
			this.txtNegocioCalleSecundaria.Enabled = false;
			this.txtNegocioTelefono1.Enabled = false;
			this.txtNegocioTelefono2.Enabled = false;
			this.txtNEgocioCelular.Enabled = false;
			this.txtNegocioReferncia.Enabled = false;
			this.txtNEgocioIngresos.Enabled =false;
			this.txtNEgocioGastos.Enabled = false;
			//this.txtNegocioActivos.Enabled = false;
			//this.txtNEgocioPasivos.Enabled = false;
			this.chkNegocioObligadoaLlevar.Enabled = false;
			this.chkAfiliadoTributario.Enabled = false;
			this.txtNotas.Enabled=false;
			#endregion NegocioD

		}

		private void Habilita_DatosDependiente ()
		{
			#region Habilita_Dependiente
						this.txtDependienteRuc.Enabled = true;
						this.txtDependienteEmpresa.Enabled = true;
						this.cmbDependienteTipoempresa.Enabled = true;
						this.dtDependienteFechaIngreso.Enabled = true;
						//this.cmbTiempoDependiente.Enabled =true;
						this.txtDependienteIngresos.Enabled = true;
						this.txtDependienteGastos.Enabled = true;
						//this.txtDependienteActivos.Enabled = true;
			      //this.txtDependientePasivos.Enabled = true;
						this.cmbDependienteTipocontrato.Enabled = true;
						this.cmbDependienteTipoSueldo.Enabled = true;
						this.txtDependienteDepartamento.Enabled = true;
						this.cmbDependienteCargo.Enabled = true;
						this.txtDependienteDiasPago.Enabled = true;
						this.cmbDependienteProvincia.Enabled = true;
						this.cmbDependienteCanton.Enabled = true;
						this.cmbDependienteParroquia.Enabled = true;
						this.cmbDependienteBarrio.Enabled = true;
						this.txtDependienteemail.Enabled = true;
						this.txtDependienteCallePrinciapl.Enabled = true;
						this.txtDependienteNumeroCasa.Enabled = true;
						this.txtDependienteCalleSecundaria.Enabled = true;
						this.txtDependienteTelefono.Enabled = true;
						this.txtDependienteEXT.Enabled = true;
						this.txtDependienteCelular.Enabled = true;
            this.txtDependienteReferencia.Enabled = true;
			      this.chkAfiliadoIESS.Enabled = true;

			#endregion Habilita_Dependiente
		}

		private void Desabilita_DatosDependiente ()
		{
			#region Desabilita_Dependiente
			this.txtDependienteRuc.Enabled = false;
			this.txtDependienteEmpresa.Enabled = false;
			this.cmbDependienteTipoempresa.Enabled = false;
			this.dtDependienteFechaIngreso.Enabled = false;
			//this.cmbTiempoDependiente.Enabled= false;
			this.txtDependienteIngresos.Enabled = false;
			this.txtDependienteGastos.Enabled = false;
			//this.txtDependienteActivos.Enabled = false;
			//this.txtDependientePasivos.Enabled = false;
			this.cmbDependienteTipocontrato.Enabled = false;
			this.cmbDependienteTipoSueldo.Enabled = false;
			this.txtDependienteDepartamento.Enabled = false;
			this.cmbDependienteCargo.Enabled = false;
			this.txtDependienteDiasPago.Enabled = false;
			this.cmbDependienteProvincia.Enabled = false;
			this.cmbDependienteCanton.Enabled = false;
			this.cmbDependienteParroquia.Enabled = false;
			this.cmbDependienteBarrio.Enabled = false;
			this.txtDependienteemail.Enabled = false;
			this.txtDependienteCallePrinciapl.Enabled =false;
			this.txtDependienteNumeroCasa.Enabled = false;
			this.txtDependienteCalleSecundaria.Enabled = false;
			this.txtDependienteTelefono.Enabled = false;
			this.txtDependienteEXT.Enabled = false;
			this.txtDependienteCelular.Enabled = false;
			this.txtDependienteReferencia.Enabled = false;
			this.chkAfiliadoIESS.Enabled = false;
			#endregion Desabilita_Dependiente
		}

		private void Habilita_InformacionCredito()
		{
				
					this.txtCreditoObservaciones.Enabled = true;
					this.cmbCreditoEstado.Enabled = true;
					this.txtCupo.Enabled = true;
					this.txtCreditoCuotaAsiganada.Enabled = true;
					this.cmbCreditoCalificacion.Enabled = true;
					this.txtObservaciones.Enabled = true;
					bHabilitaPerfil=true;
					//this.cmbSituacionLaboral2.Enabled = true;
					//this.ChkOmite.Enabled = true;
					//this.chkNuovo.Enabled = true;
					//this.cmbEntidad.Enabled = true;
					//this.cmbEntidad.Enabled = true;
					//this.cmbidTipoCliente.Enabled = true;

					//this.cmbdetallecliente.Enabled = true;
					//this.btnIdValidator.Visible=true;
					//this.btnImprimeVerificacion.Enabled = true;
					//this.tbBasico.Tabs[5].Visible = true;
					//this.tbBasico.Tabs[6].Visible = true;
					//this.tbBasico.Tabs[7].Visible = true;

		}

		private void groupBox4_Enter(object sender, System.EventArgs e)
		{
		
		}

		private void cmbDPCCargo_InitializeLayout(object sender, Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs e)
		{
		
		}

		private void groupBox3_Enter(object sender, System.EventArgs e)
		{
		
		}

		private void tpComplemento_Paint(object sender, System.Windows.Forms.PaintEventArgs e)
		{
		
		}

		private void cmbEstadoCivil_ValueChanged(object sender, System.EventArgs e)
		{
			EstadoCivil();
		}

		private void EstadoCivil()
		{
			if (this.cmbEstadoCivil.ActiveRow != null)
			{
				if ((int)this.cmbEstadoCivil.Value != 2 && (int)this.cmbEstadoCivil.Value != 3 && (int)this.cmbEstadoCivil.Value != 5)
				{
					this.cmbDPCNacionalidad.Value = 54;
					this.cmbDPCTipodocumento.Value = 1;
//					this.tbBasico.Tabs[1].Enabled = true;  
				 this.btnConyugueTelefono.Enabled = true;
					this.btnConyugueCelular.Enabled = true;
          HabilitaConyuge ()  ;     
				}
				else
				{
					this.cmbDPCNacionalidad.Value = 0;
					this.cmbDPCTipodocumento.Value = 0;
//					this.tbBasico.Tabs[1].Enabled = false;
					this.btnConyugueTelefono.Enabled = false;
					this.btnConyugueCelular.Enabled = false;
					DesabilitaConyuge ()  ; 
				}
			}
		}

//		private void NovaScoring( int IdCre_DatosGenerales)
//		{
//			bool bRes = Funcion.bEscalarSQL(cdsSeteoF, string.Format("Exec ValidaDatosGeneralesCognoScoring {0}",IdCre_DatosGenerales), true);
//			if (bRes)
//			{
//
//				#region Informacion General NovaScoring
//
//				string ssqlCTNG= string.Format ("Exec Cre_ConsultaInforcionEquifax {0}",IdCre_DatosGenerales ); 
//				SqlDataReader drDCTNG = Funcion.rEscalarSQL(cdsSeteoF,ssqlCTNG,true);
//				drDCTNG.Read();
//				if (drDCTNG.HasRows)
//				{	
//					if (drDCTNG.GetValue(0) != System.DBNull.Value)this.txtCreditoScoreInclusion.Value = drDCTNG.GetInt32(0);
//					if (drDCTNG.GetValue(1) != System.DBNull.Value)this.txtCreditoScoreSobreEndeudamiento.Value = drDCTNG.GetInt32(1);
//					if (drDCTNG.GetValue(2) != System.DBNull.Value)this.txtRangoIngresos.Text = drDCTNG.GetString (2);
//					if (drDCTNG.GetValue(3) != System.DBNull.Value)this.txtCapacidadPago.Value = drDCTNG.GetDecimal (3);
//					if (drDCTNG.GetValue(4) != System.DBNull.Value)this.txtTotalVencido.Value = drDCTNG.GetDecimal (4);
//
//				}
//				drDCTNG.Close();
//						
//						
//				#endregion Informacion General NovaScoring
//		
//			}
//			
//		}

		private void cmbDDTipoVivienda_ValueChanged(object sender, System.EventArgs e)
		{
			TipoVivienda();
		
		}
		private void TipoVivienda()
		{
			if (this.cmbDDTipoVivienda.ActiveRow != null)
			{
				if (!bCarga)
				{
					this.txtNombreArrendador.Text = "";
					this.txtTelfArrendador.Text = "";
					this.txtCelArrendador.Text = "";
					this.cmbDDInmueble.Value = 0;
					this.cmbDDCiudadInmueble.Value = 0;
					this.txtDDValorInmueble.Value = 0;
				}                       
				this.txtNombreArrendador.Enabled = false;
				this.txtTelfArrendador.Enabled = false;
				this.txtCelArrendador.Enabled = false;

				this.cmbDDInmueble.Enabled = false;
				this.cmbDDCiudadInmueble.Enabled = false;
				this.txtDDValorInmueble.Enabled = false;
                                
				if ((int)this.cmbDDTipoVivienda.Value == 1)
				{
					this.txtNombreArrendador.Enabled = true;
					this.txtTelfArrendador.Enabled = true;
					this.txtCelArrendador.Enabled = true;
					
				}
				else if ((int)this.cmbDDTipoVivienda.Value == 3 || (int)this.cmbDDTipoVivienda.Value == 4)
				{
					this.cmbDDInmueble.Enabled = true;
					this.cmbDDCiudadInmueble.Enabled = true;
					this.txtDDValorInmueble.Enabled = true;
				}
			}
		}

		private void txtDDValorInmueble_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if(e.KeyCode == Keys.Enter) 
			{
				this.tbBasico.Tabs[1].Selected = true;
				this.txtDPCdocumento.Focus ();
				if ( this.cmbDPCTipodocumento.Enabled == false) 
				{
						this.tbBasico.Tabs[2].Selected = true;
				}
			}
		}

		private void cmbDPCTProvincia_ValueChanged(object sender, System.EventArgs e)
		{
			if (this.cmbDPCTProvincia.ActiveRow != null)
			{
				this.cmbDPCTCanton.DataSource = Funcion.dvProcedimiento(cdsSeteoF, string.Format("Exec Listas 1, {0}, 0", (int) this.cmbDPCTProvincia.Value));
				
			}
		}

		private void cmbDPCTCanton_ValueChanged(object sender, System.EventArgs e)
		{
			if (this.cmbDPCTCanton.ActiveRow != null)
			{
				this.cmbDPCTParroquia.DataSource = Funcion.dvProcedimiento(cdsSeteoF, string.Format("Exec Listas 2, {0}, 0", (int) this.cmbDPCTCanton.Value));
				
			}
		}

		private void cmbDPCTParroquia_ValueChanged(object sender, System.EventArgs e)
		{
			if (this.cmbDPCTParroquia.ActiveRow != null)
			{
				this.cmbDPCTBarrio.DataSource = Funcion.dvProcedimiento(cdsSeteoF, string.Format("Exec Listas 3, {0}, 0", (int) this.cmbDPCTParroquia.Value));
				
			}
		}

		private void cmbDPCProvinciaN_ValueChanged(object sender, System.EventArgs e)
		{
			if (this.cmbDPCProvinciaN.ActiveRow != null)
			{
				this.cmbDPCCantonN.DataSource = Funcion.dvProcedimiento(cdsSeteoF, string.Format("Exec Listas 1, {0}, 0", (int) this.cmbDPCProvinciaN.Value));
				
			}
		}

		private void txtDPCdocumento_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if(e.KeyCode == Keys.Enter) this.txtDPCApellidoP.Focus();
		}

		private void txtDPCApellidoP_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if(e.KeyCode == Keys.Enter) this.txtDPCApellidoM.Focus();
		}

		private void txtDPCApellidoM_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if(e.KeyCode == Keys.Enter) this.txtDPCPrimerN.Focus();
		}

		private void txtDPCPrimerN_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
		 if(e.KeyCode == Keys.Enter) this.txtDPCsegundoN.Focus();
		}

		private void txtDPCsegundoN_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
		  if(e.KeyCode == Keys.Enter) this.cmbDPCNacionalidad.Focus();
		}

		private void cmbDPCNacionalidad_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			 if(e.KeyCode == Keys.Enter) this.dtDPCFnacimiento.Focus();
		}

		private void dtDPCFnacimiento_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if(e.KeyCode == Keys.Enter) this.cmbDPCSexo.Focus();
		}

		private void cmbDPCSexo_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			//if(e.KeyCode == Keys.Enter) this.cmbDPCProvinciaN.Focus();
		}

		private void cmbDPCProvinciaN_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if(e.KeyCode == Keys.Enter) this.cmbDPCCantonN.Focus();
		}

		private void cmbDPCCantonN_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if(e.KeyCode == Keys.Enter) this.cmbDPCNivelEducacion.Focus();
		}

		private void cmbDPCNivelEducacion_InitializeRow(object sender, Infragistics.Win.UltraWinGrid.InitializeRowEventArgs e)
		{
		
		}

		private void cmbDPCNivelEducacion_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
				if(e.KeyCode == Keys.Enter) this.cmbDPCProfesion.Focus();
		}

		private void cmbDPCProfesion_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
				if(e.KeyCode == Keys.Enter) this.cmbDPCSituacionL.Focus();
		}

		private void cmbDPCSituacionL_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if(e.KeyCode == Keys.Enter) this.cmbDPCActividadE.Focus();
		}

		private void cmbDPCActividadE_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if(e.KeyCode == Keys.Enter) this.txtDPCTNombreEmpresa.Focus();
		}

		private void txtDPCTNombreEmpresa_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if(e.KeyCode == Keys.Enter) this.cmbDPCTActividadEmpresa.Focus();
		}

		private void cmbDPCTActividadEmpresa_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
						if(e.KeyCode == Keys.Enter) this.cmbDPCTCargo.Focus();
		}

		private void cmbDPCTCargo_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
				if(e.KeyCode == Keys.Enter) this.cmbDPCTProvincia.Focus();
		}

		private void cmbDPCTProvincia_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
				if(e.KeyCode == Keys.Enter) this.cmbDPCTCanton.Focus();
		}

		private void cmbDPCTCanton_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if(e.KeyCode == Keys.Enter) this.cmbDPCTParroquia.Focus();
		}

		private void cmbDPCTParroquia_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if(e.KeyCode == Keys.Enter) this.cmbDPCTBarrio.Focus();
		}

		private void cmbDPCTBarrio_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if(e.KeyCode == Keys.Enter) this.txtDPCTEmail.Focus();
		}

		private void txtDPCTEmail_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if(e.KeyCode == Keys.Enter) this.txtDPCTCallePrincipal.Focus();
		}

		private void txtDPCTCallePrincipal_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
					if(e.KeyCode == Keys.Enter) this.txtDPCTNumeroCasa.Focus();
		}

		private void txtDPCTNumeroCasa_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
				if(e.KeyCode == Keys.Enter) this.txtDPCTCalleSecundaria.Focus();
		}

		private void txtDPCTCalleSecundaria_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if(e.KeyCode == Keys.Enter) this.txtDPCTtelefono.Focus();
		}

		private void txtDPCTtelefono_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if(e.KeyCode == Keys.Enter) this.txtDPCTExt.Focus();
		}

		private void txtDPCTExt_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if(e.KeyCode == Keys.Enter) this.txtDPCTCelularC.Focus();
		}

		private void txtDPCTCelularC_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if(e.KeyCode == Keys.Enter) this.txtDPCTReferencia.Focus();
		}

		private void txtDPCTReferencia_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if(e.KeyCode == Keys.Enter)
			{
					this.txtDPCTIngresos.Focus();
				this.txtDPCTIngresos.SelectAll ();
			}
		}

		private void txtDPCTIngresos_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if(e.KeyCode == Keys.Enter) 
			{
					this.txtDPCTGastos.Focus();
				this.txtDPCTGastos.SelectAll ();
			}
		}

		private void txtDPCTGastos_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
//			if(e.KeyCode == Keys.Enter)
//			{
//					this.txtDPCTActivos.Focus();
//				this.txtDPCTActivos.SelectAll ();
//			}
		}
		private void txtDPCTActivos_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
//			if(e.KeyCode == Keys.Enter)
//			{
//					this.txtDPCTPasivos.Focus();
//				this.txtDPCTPasivos.SelectAll();
//			}
		}

		private void cmbDPCActividadE_KeyDown_1(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if(e.KeyCode == Keys.Enter) this.txtDPCTNombreEmpresa.Focus();
		}

		private void tbBasico_SelectedTabChanged(object sender, Infragistics.Win.UltraWinTabControl.SelectedTabChangedEventArgs e)
		{
		
		}

		private void label4_Click(object sender, System.EventArgs e)
		{
		
		}

		private void cmbParentesco_InitializeLayout(object sender, Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs e)
		{
		
		}

		private void cmbProvincia_Validating(object sender, System.ComponentModel.CancelEventArgs e)
		{
		
		}

		private void cmbProvincia_RowSelected(object sender, Infragistics.Win.UltraWinGrid.RowSelectedEventArgs e)
		{
//			  this.cmbProvincia.Update();
			if (e.Row != null)
			{
//		    e.Row.Cells["Select"].Value = true;
			

				if ((int)e.Row.Cells["idProvincia"].Value > 0 )
				{
					int iProvincia = (int)e.Row.Cells["idProvincia"].Value;
					this.ultraDropDown1.DataSource = Funcion.dvProcedimiento(cdsSeteoF, string.Format("Exec Listas 1, {0}, 0", iProvincia));
//					this.cmbProvincia.Update();
					this.ultraDropDown1.Update();
				}
			}
				this.cmbProvincia.Update();
			 this.ultraDropDown1.Update();
		
		}

		private void ultraTabPageControl5_Paint(object sender, System.Windows.Forms.PaintEventArgs e)
		{
		
		}

		private void ultgconsulta_AfterRowInsert(object sender, Infragistics.Win.UltraWinGrid.RowEventArgs e)
		{
			e.Row.Cells["idCre_ReferenciasClientes"].Value = 0;
			e.Row.Cells["idCre_DatosGenerales"].Value = 0;

		}

		private void uGridRefernciaBancaria_AfterRowInsert(object sender, Infragistics.Win.UltraWinGrid.RowEventArgs e)
		{
			e.Row.Cells["idCre_ReferenciaBancaria"].Value = 0;
			e.Row.Cells["idCre_DatosGenerales"].Value = 0;
			e.Row.Cells["CupoTarjeta"].Value = 0;
		}

		private void label82_Click(object sender, System.EventArgs e)
		{
		
		}

		private void cmbNegocioProvincia_ValueChanged(object sender, System.EventArgs e)
		{
			if (this.cmbNegocioProvincia.ActiveRow != null)
			{
				this.cmbNegocioCanton.DataSource = Funcion.dvProcedimiento(cdsSeteoF, string.Format("Exec Listas 1, {0}, 0", (int) this.cmbNegocioProvincia.Value));
				
			}
		}

		private void cmbNegocioCanton_ValueChanged(object sender, System.EventArgs e)
		{
			if (this.cmbNegocioCanton.ActiveRow != null)
			{
				this.cmbNegocioParroquia.DataSource = Funcion.dvProcedimiento(cdsSeteoF, string.Format("Exec Listas 2, {0}, 0", (int) this.cmbNegocioCanton.Value));
				
			}
		}

		private void cmbNegocioParroquia_ValueChanged(object sender, System.EventArgs e)
		{
			if (this.cmbNegocioParroquia.ActiveRow != null)
			{
				this.cmbNegocioBarrio.DataSource = Funcion.dvProcedimiento(cdsSeteoF, string.Format("Exec Listas 3, {0}, 0", (int) this.cmbNegocioParroquia.Value));
				
			}
		}

		private void txtNegocioRuc_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
				if(e.KeyCode == Keys.Enter) this.txtNombreDelNegocio.Focus();
		}

		private void txtNombreDelNegocio_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
//			if(e.KeyCode == Keys.Enter) this.dtpNegocioFechaInicio.Focus();
		}

		private void dtpNegocioFechaInicio_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if(e.KeyCode == Keys.Enter)
			{
					this.txtNegocioMetros.Focus();
					this.txtNegocioMetros.SelectAll ();
			}
		}

		private void txtNegocioMetros_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if(e.KeyCode == Keys.Enter)
			{
					this.txtNegocioEmpleados.Focus();
				this.txtNegocioEmpleados.SelectAll ();
			}
			}

		private void txtNegocioEmpleados_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if(e.KeyCode == Keys.Enter) this.cmbNegciotipo.Focus();
		}

		private void cmbNegciotipo_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if(e.KeyCode == Keys.Enter) this.cmbNEgocioActivdadEconomica.Focus();
		}

		private void cmbNEgocioActivdadEconomica_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if(e.KeyCode == Keys.Enter) this.cmbNegocioProvincia.Focus();
		}

		private void cmbNegocioProvincia_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
		if(e.KeyCode == Keys.Enter) this.cmbNegocioCanton.Focus();
		}

		private void cmbNegocioCanton_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
				if(e.KeyCode == Keys.Enter) this.cmbNegocioParroquia.Focus();
		}

		private void cmbNegocioParroquia_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if(e.KeyCode == Keys.Enter) this.cmbNegocioBarrio.Focus();
		}

		private void cmbNegocioBarrio_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if(e.KeyCode == Keys.Enter) this.txtNegocioemail.Focus();
		}

		private void txtNegocioemail_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
				if(e.KeyCode == Keys.Enter) this.txtNegocioCallePrincipal.Focus();
		}

		private void txtNegocioCallePrincipal_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
					if(e.KeyCode == Keys.Enter) this.txtNegocioNumeroCasa.Focus();
		}

		private void txtNegocioNumeroCasa_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
				if(e.KeyCode == Keys.Enter) this.txtNegocioCalleSecundaria.Focus();
		}

		private void txtNegocioCalleSecundaria_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if(e.KeyCode == Keys.Enter) this.txtNegocioTelefono1.Focus();
		}

		private void txtNegocioTelefono1_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if(e.KeyCode == Keys.Enter) this.txtNegocioTelefono2.Focus();
		}

		private void txtNegocioTelefono2_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
				if(e.KeyCode == Keys.Enter) this.txtNEgocioCelular.Focus();
		}

		private void txtNEgocioCelular_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if(e.KeyCode == Keys.Enter) this.txtNegocioReferncia.Focus();
		}

		private void txtNegocioReferncia_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if(e.KeyCode == Keys.Enter) 
			{
					this.txtNEgocioIngresos.Focus();
				this.txtNEgocioIngresos.SelectAll();
			}
		}
		private void txtNEgocioIngresos_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if(e.KeyCode == Keys.Enter) 
			{
					this.txtNEgocioGastos.Focus();
				this.txtNEgocioGastos.SelectAll();
			}
		}
		private void txtNEgocioGastos_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
//			if(e.KeyCode == Keys.Enter)
//			{
//					this.txtNegocioActivos.Focus();
//				this.txtNegocioActivos.SelectAll ();
//			}
		}
		private void txtNegocioActivos_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
//			if(e.KeyCode == Keys.Enter) 
//			{
//					this.txtNEgocioPasivos.Focus();
//				this.txtNEgocioPasivos.SelectAll ();
//			}
		}
		private void ultraCombo6_InitializeLayout(object sender, Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs e)
		{
		
		}

		private void label98_Click(object sender, System.EventArgs e)
		{
		
		}

		private void txtDependienteRuc_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
		 if(e.KeyCode == Keys.Enter) this.txtDependienteEmpresa.Focus();
		}

		private void txtDependienteEmpresa_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if(e.KeyCode == Keys.Enter) this.cmbDependienteTipoempresa.Focus();
		}

		private void cmbDependienteTipoempresa_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if(e.KeyCode == Keys.Enter) this.dtDependienteFechaIngreso.Focus();
		}

		private void dtDependienteFechaIngreso_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if(e.KeyCode == Keys.Enter)
			{
					this.txtDependienteIngresos.Focus();
					this.txtDependienteIngresos.SelectAll ();
			}
		}

		private void txtDependienteGastos_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
//			if(e.KeyCode == Keys.Enter) 
//			{
//					this.txtDependienteActivos.Focus();
//				this.txtDependienteActivos.SelectAll ();
//			}
		}

		private void txtDependienteIngresos_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if(e.KeyCode == Keys.Enter) 
			{
					this.txtDependienteGastos.Focus();
				this.txtDependienteGastos.SelectAll ();
			}
		}

		private void txtDependienteActivos_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
//			if(e.KeyCode == Keys.Enter)
//			{
//					this.txtDependientePasivos.Focus();
//				this.txtDependientePasivos.SelectAll ();
//			}
		}

		private void txtDependientePasivos_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if(e.KeyCode == Keys.Enter) this.cmbDependienteTipocontrato.Focus();
		}

		private void cmbDependienteTipocontrato_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if(e.KeyCode == Keys.Enter) this.cmbDependienteTipoSueldo.Focus();
		}

		private void cmbDependienteTipoSueldo_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
				if(e.KeyCode == Keys.Enter) this.txtDependienteDepartamento.Focus();
		}

		private void txtDependienteDepartamento_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
		 if(e.KeyCode == Keys.Enter) this.cmbDependienteCargo.Focus();
		}

		private void cmbDependienteCargo_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if(e.KeyCode == Keys.Enter)
			{
					this.txtDependienteDiasPago.Focus();
				this.txtDependienteDiasPago.SelectAll ();
			}
		}

		private void txtDependienteDiasPago_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if(e.KeyCode == Keys.Enter) this.cmbDependienteProvincia.Focus();
		}

		private void cmbDependienteProvincia_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if(e.KeyCode == Keys.Enter) this.cmbDependienteCanton.Focus();
		}

		private void cmbDependienteCanton_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if(e.KeyCode == Keys.Enter) this.cmbDependienteParroquia.Focus();
		}

		private void cmbDependienteParroquia_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if(e.KeyCode == Keys.Enter) this.cmbDependienteBarrio.Focus();
		}

		private void cmbDependienteBarrio_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
		 if(e.KeyCode == Keys.Enter) this.txtDependienteemail.Focus();
		}

		private void txtDependienteemail_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
		 if(e.KeyCode == Keys.Enter) this.txtDependienteCallePrinciapl.Focus();
		}

		private void txtDependienteCallePrinciapl_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
					 if(e.KeyCode == Keys.Enter) this.txtDependienteNumeroCasa.Focus();
		}

		private void txtDependienteNumeroCasa_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			 if(e.KeyCode == Keys.Enter) this.txtDependienteCalleSecundaria.Focus();
		}

		private void txtDependienteCalleSecundaria_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if(e.KeyCode == Keys.Enter) this.txtDependienteTelefono.Focus();
		}

		private void txtDependienteTelefono_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if(e.KeyCode == Keys.Enter) this.txtDependienteEXT.Focus();
		}

		private void txtDependienteEXT_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
				if(e.KeyCode == Keys.Enter) this.txtDependienteCelular.Focus();
		}

		private void txtDependienteCelular_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
		 	if(e.KeyCode == Keys.Enter) this.txtDependienteReferencia.Focus();
		}

		private void cmbDependienteTipoSueldo_KeyDown_1(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if(e.KeyCode == Keys.Enter) this.cmbDependienteTipoSueldo.Focus();
			
		}

		private void cmbDependienteProvincia_ValueChanged(object sender, System.EventArgs e)
		{
			if (this.cmbDependienteProvincia.ActiveRow != null)
			{
				this.cmbDependienteCanton.DataSource = Funcion.dvProcedimiento(cdsSeteoF, string.Format("Exec Listas 1, {0}, 0", (int) this.cmbDependienteProvincia.Value));
				
			}
		}

		private void cmbDependienteCanton_ValueChanged(object sender, System.EventArgs e)
		{
			if (this.cmbDependienteCanton.ActiveRow != null)
			{
				this.cmbDependienteParroquia.DataSource = Funcion.dvProcedimiento(cdsSeteoF, string.Format("Exec Listas 2, {0}, 0", (int) this.cmbDependienteCanton.Value));
				
			}
		}

		private void cmbDependienteParroquia_ValueChanged(object sender, System.EventArgs e)
		{
			if (this.cmbDependienteParroquia.ActiveRow != null)
			{
				this.cmbDependienteBarrio.DataSource = Funcion.dvProcedimiento(cdsSeteoF, string.Format("Exec Listas 3, {0}, 0", (int) this.cmbDependienteParroquia.Value));
				
			}
		}

		private void ultraCalendarCombo1_BeforeDropDown(object sender, System.ComponentModel.CancelEventArgs e)
		{
		
		}

		private void ultraNumericEditor1_ValueChanged(object sender, System.EventArgs e)
		{
		
		}

		private void groupBox7_Enter(object sender, System.EventArgs e)
		{
		
		}

		private void txtCreditoScore_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if(e.KeyCode == Keys.Enter) 
			{
					this.txtCreditoCuotaAsiganada.Focus();
				this.txtCreditoCuotaAsiganada.SelectAll ();
			}
		}

		private void txtCreditoCuotaAsiganada_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if(e.KeyCode == Keys.Enter) this.cmbCreditoCalificacion.Focus();
		}

		private void cmbCreditoCalificacion_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if(e.KeyCode == Keys.Enter) this.cmbCreditoEstado.Focus();
		}

		private void cmbCreditoEstado_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if(e.KeyCode == Keys.Enter) this.txtCreditoObservaciones.Focus();
		}

		private void txtDocumento_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
			validar.Solonumeros(e);
		}

		private void txtDocumento_Validated(object sender, System.EventArgs e)
		{
			this.cmbEntidad.DataSource = Funcion.dvProcedimiento(cdsSeteoF, "Select idEntidadFinanciera, Nombre From Cre_EntidadFinanciera where idEntidadFinanciera not in (6,3) Order By idEntidadFinanciera");
	    VaciaCamposGenerales ();
					if (this.txtDocumento.Text.ToString().Trim().Length == 0  ) return;
//			using (frmCred_RevisiondeCedulas  consultaCR = new frmCred_RevisiondeCedulas(this.txtDocumento.Text))
//			{
//				if (DialogResult.OK == consultaCR.ShowDialog())
//				{
					bool bRes = Funcion.bEscalarSQL(cdsSeteoF, string.Format("Exec ValidaDatosCognoScoring '{0}'",this.txtDocumento.Text), true);
					if (bRes)
					{
						string sSQLDatosNovaScoring = string.Format("Exec ActualizaDatosNovaScoring '{0}'",this.txtDocumento.Text);
						Funcion.EjecutaSQL(cdsSeteoF, sSQLDatosNovaScoring);
						MessageBox.Show("Datos de Sincronizados con Cogno Ware", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Information);				
					}

					String sSqlConsulta= string.Format ("Exec Cre_ConsultaRevisionCedula '{0}'",this.txtDocumento.Text); 
					SqlDataReader drConsulta = Funcion.rEscalarSQL(cdsSeteoF,sSqlConsulta,true);
					drConsulta.Read();			
						
					if (drConsulta.HasRows)
					{		     
						if (drConsulta.GetValue(2) != System.DBNull.Value)  this.txtIdCre_Solicitud.Value = drConsulta.GetInt32 (2);													
					}
					drConsulta.Close();

					bCarga = true;
					//this.txtIdCre_Solicitud.Value = consultaCR.txtidCre_solicitud.Value;

					#region Solicitud

					string ssql= string.Format ("Exec Cre_ConsultaSolicitudes {0}",(int)this.txtIdCre_Solicitud.Value ); 
					SqlDataReader dr = Funcion.rEscalarSQL(cdsSeteoF,ssql,true);
					dr.Read();
					if (dr.HasRows)
					{		
						if (dr.GetValue(2) != System.DBNull.Value) this.chkGarante.Checked = dr.GetBoolean (2);	
						if (dr.GetValue(3) != System.DBNull.Value) this.cmbEntidad.Value = dr.GetInt32 (3);
						//						this.cmbTipoTarjeta.Value = dr.GetInt32 (4);	
						if (dr.GetValue(5) != System.DBNull.Value) this.cmbBodega.Value =  dr.GetInt32 (5);	
						if (dr.GetValue(6) != System.DBNull.Value) this.txtCreditoCuotaAsiganada.Value =  dr.GetInt32 (6);	
						if (dr.GetValue(15) != System.DBNull.Value) this.dtCreditoFechaSolicitud.Value =  dr.GetDateTime (15);	
						if (dr.GetValue(16) != System.DBNull.Value) this.lblUsuarioCompleta.Text =  dr.GetString (16);	
						if (dr.GetValue(17) != System.DBNull.Value) this.txtCreditoObservaciones.Text =  dr.GetString (17);	
						if (dr.GetValue(14) != System.DBNull.Value) this.txtEstado.Value =  dr.GetInt32 (14);	

						if (dr.GetValue(18) != System.DBNull.Value) this.cmbidTipoCliente.Value =  dr.GetInt32 (18);	
						if (dr.GetValue(19) != System.DBNull.Value) this.cmbdetallecliente.Value =  dr.GetInt32 (19);	
						if (dr.GetValue(20) != System.DBNull.Value) this.cmbSituacionLaboral2.Value =  dr.GetInt32 (20);	
						if (dr.GetValue(21) != System.DBNull.Value) this.txtCupo.Value =  dr.GetDecimal (21);	
						if (dr.GetValue(9) != System.DBNull.Value) this.cmdEstadoSolicitud.Value =  dr.GetInt32 (9);	
					}
					dr.Close();

					
					if ((int)this.cmbEntidad.Value == 2)
					{
						this.label25.Text = "Fuente de Ingresos";
					
						//												this.cmbSituacionLaboral.Visible = false;
						//												this.cmbClienteFuenteDeIngresos.Visible = true;
					}
					else
					{
						this.label25.Text = "Situación Laboral";
					
						//												this.cmbClienteFuenteDeIngresos.Visible = false;
						//												this.cmbSituacionLaboral.Visible = true;
					}
					#region Estados
					if (Convert.ToInt32(this.txtEstado.Value) == 2)
					{
						this.cmbEntidad.Enabled = false;
						this.chkGarante.Enabled = false;
						//						this.btnGarante.Enabled = false;
						this.lblEstado.Text = "NEGADO";
					}
					else if (Convert.ToInt32(this.txtEstado.Value) == 1)
					{
						//this.cmbEntidad.Enabled = true;
						//							this.btnImprimir.Enabled = true;
						this.chkGarante.Enabled = true;
						//						this.btnGarante.Enabled = true;
						this.lblEstado.Text = "COMPLETA";
					}
					else if (Convert.ToInt32(this.txtEstado.Value) == 0)
					{
						//this.cmbEntidad.Enabled = true;
						//							this.btnImprimir.Enabled = false;
						this.chkGarante.Enabled = true;
						this.btnGarante.Enabled = true;
						this.lblEstado.Text = "PENDIENTE";
					}                                
					#endregion Estados



					#endregion Solicitud

					#region Datos_Generales

					string ssqlDG= string.Format ("Exec Cre_ConsultaDatosGenerales {0},0",(int)this.txtIdCre_Solicitud.Value ); 

					SqlDataReader drDG = Funcion.rEscalarSQL(cdsSeteoF,ssqlDG,true);
					drDG.Read();
					if (drDG.HasRows)
					{	
						this.txtClienteIdCre_DatosGenerales.Value = drDG.GetInt32 (0);
						if (drDG.GetValue(3) != System.DBNull.Value)this.cmbTipoRuc.Value = drDG.GetInt32 (3);
						if (drDG.GetValue(4) != System.DBNull.Value)this.txtDocumento.Text =  drDG.GetString (4);	
						if (drDG.GetValue(5) != System.DBNull.Value)this.txtApellidoP.Text =  drDG.GetString (5);	
						if (drDG.GetValue(6) != System.DBNull.Value)this.txtApellidoM.Text =  drDG.GetString (6);	
						if (drDG.GetValue(7) != System.DBNull.Value)this.txtPnombre.Text =  drDG.GetString (7);	
						if (drDG.GetValue(8) != System.DBNull.Value)this.txtSnombre.Text =  drDG.GetString (8);	
						if (drDG.GetValue(9) != System.DBNull.Value)this.cmbCreditoCalificacion.Value = drDG.GetInt32 (9);
						if (drDG.GetValue(11) != System.DBNull.Value)this.txtCreditoScore.Value = drDG.GetInt32 (11);
						if (drDG.GetValue(13) != System.DBNull.Value)this.cmbCreditoEstado.Value = drDG.GetInt32 (13);
						if (drDG.GetValue(14) != System.DBNull.Value)this.cmbNacionalidad.Value = drDG.GetInt32 (14);
						if (drDG.GetValue(15) != System.DBNull.Value)this.dtFechaNacimiento.Value = drDG.GetDateTime (15);
						if (drDG.GetValue(16) != System.DBNull.Value)this.cmbSexo.Value = drDG.GetInt32 (16);
						if (drDG.GetValue(17) != System.DBNull.Value)this.cmbProvinciaNAcimiento.Value = drDG.GetInt32 (17);
						if (drDG.GetValue(18) != System.DBNull.Value)this.cmbCantonNacimiento.Value = drDG.GetInt32 (18);
						if (drDG.GetValue(19) != System.DBNull.Value)this.cmbEstadoCivil.Value = drDG.GetInt32 (19);
						if (drDG.GetValue(20) != System.DBNull.Value)this.txtDependientes.Value = drDG.GetInt32 (20);
						if (drDG.GetValue(21) != System.DBNull.Value)this.cmbNivelEducacion.Value = drDG.GetInt32 (21);
						if (drDG.GetValue(22) != System.DBNull.Value)this.cmbProfesion.Value = drDG.GetInt32 (22);
						if (drDG.GetValue(23) != System.DBNull.Value)this.cmbSituacionLaboral.Value = drDG.GetInt32 (23); //this.cmbClienteFuenteDeIngresos.Value = drDG.GetInt32 (23);
						if (drDG.GetValue(24) != System.DBNull.Value)this.cmbSituacionLaboral.Value = drDG.GetInt32 (24);
						SituacionLaboral((int)this.cmbSituacionLaboral.Value);
						if (drDG.GetValue(25) != System.DBNull.Value)this.cmbActividadEconomica.Value = drDG.GetInt32 (25);                             
						if (drDG.GetValue(33) != System.DBNull.Value)this.txtVehiculoMarca.Text = drDG.GetString (33);
						if (drDG.GetValue(34) != System.DBNull.Value)this.txtModelo.Text = drDG.GetString (34);
						if (drDG.GetValue(35) != System.DBNull.Value)this.txtValorVehiculo.Value = drDG.GetInt32 (35);
						if (drDG.GetValue(36) != System.DBNull.Value)this.chkSeparacionBienes.Checked = drDG.GetBoolean (36);	
						if (drDG.GetValue(37) != System.DBNull.Value)this.chkHipoteca.Checked = drDG.GetBoolean (37);	

						if (drDG.GetValue(30) != System.DBNull.Value)this.cmbDDInmueble.Value = drDG.GetInt32 (30);
						if (drDG.GetValue(31) != System.DBNull.Value)this.cmbDDCiudadInmueble.Value = drDG.GetInt32 (31);
						if (drDG.GetValue(32) != System.DBNull.Value)this.txtDDValorInmueble.Value = drDG.GetInt32 (32);
						
						if (drDG.GetValue(38) != System.DBNull.Value)this.txtCreditoScoreInclusion.Value = drDG.GetInt32 (38);
						if (drDG.GetValue(39) != System.DBNull.Value)this.txtCreditoScoreSobreEndeudamiento.Value = drDG.GetInt32(39);
						if (drDG.GetValue(40) != System.DBNull.Value)this.txtRangoIngresos.Text = drDG.GetString (40);
						if (drDG.GetValue(41) != System.DBNull.Value)this.txtCapacidadPago.Value = drDG.GetDecimal (41);
						if (drDG.GetValue(42) != System.DBNull.Value)this.txtTotalVencido.Value = drDG.GetDecimal (42);

						if (drDG.GetValue(43) != System.DBNull.Value)this.txtSegmentacion.Text = drDG.GetString (43);
						if (drDG.GetValue(44) != System.DBNull.Value)this.txtCarteraCastigada.Value = drDG.GetDecimal (44);
						if (drDG.GetValue(45) != System.DBNull.Value)this.txtResultado.Text = drDG.GetString (45);

						if (drDG.GetValue(46) != System.DBNull.Value)this.chkNuovo.Checked = drDG.GetBoolean (46);
						if (drDG.GetValue(47) != System.DBNull.Value)this.ChkOmite.Checked = drDG.GetBoolean (47);
						if (drDG.GetValue(48) != System.DBNull.Value)this.txtObsActividadEconomica.Text = drDG.GetString (48);
						if (drDG.GetValue(49) != System.DBNull.Value)this.lblFechaActualizacion.Text = drDG.GetString (49);
						
						
					}

					drDG.Close();

					EstadoCivil();

					if (this.txtResultado.Text.ToString().Trim().Length > 2) this.lblFISL.Text = "RESPUESTA POINTSCORING  : "+" "+ this.txtResultado.Text;

					
					if (!Validacion.vbValidaFacturasCliente(this.txtDocumento.Text, cdsSeteoF)) 
					{
						this.txtPnombre.Enabled = false;
						this.txtSnombre.Enabled = false;
						this.txtApellidoP.Enabled = false;
						this.txtApellidoM.Enabled = false;
					}

					#endregion Datos_Generales

					#region Datos_Ubicacion

					string ssqlDU= string.Format ("Exec Cre_ConsultaDatosUbicacion {0},0",(int)this.txtIdCre_Solicitud.Value ); 
					SqlDataReader drDU = Funcion.rEscalarSQL(cdsSeteoF,ssqlDU,true);
					drDU.Read();
					if (drDU.HasRows)
					{	
						this.txtClienteIdCre_DatosUbicacion.Value = drDU.GetInt32 (0);
						if (drDU.GetValue(3) != System.DBNull.Value)this.cmbDDProvincia.Value = drDU.GetInt32 (3);
						if (drDU.GetValue(4) != System.DBNull.Value)this.cmbDDCanton.Value = drDU.GetInt32 (4);
						if (drDU.GetValue(5) != System.DBNull.Value)this.cmbDDParroquia.Value = drDU.GetInt32 (5);
						if (drDU.GetValue(6) != System.DBNull.Value)this.cmbDDBarrio.Value = drDU.GetInt32 (6);
						if (drDU.GetValue(7) != System.DBNull.Value)this.txtDDCallaPrincipal.Text = drDU.GetString (7);
						if (drDU.GetValue(8) != System.DBNull.Value)this.txtDDNumeroCasa.Text = drDU.GetString (8);
						if (drDU.GetValue(9) != System.DBNull.Value)this.txtDDCalleSecundaria.Text = drDU.GetString (9);
						if (drDU.GetValue(10) != System.DBNull.Value)this.txtDDReferenciasU.Text = drDU.GetString (10);
						if (drDU.GetValue(11) != System.DBNull.Value)this.txtDDTelefono1.Text = drDU.GetString (11);
						if (drDU.GetValue(13) != System.DBNull.Value)this.txtDDTelefono2.Text = drDU.GetString (13);
						if (drDU.GetValue(15) != System.DBNull.Value)this.txtDDCelular.Text = drDU.GetString (15);
						if (drDU.GetValue(16) != System.DBNull.Value)this.txtDDEmail.Text = drDU.GetString (16);
						if (drDU.GetValue(17) != System.DBNull.Value)this.cmbDDTipoVivienda.Value = drDU.GetInt32 (17);
						if (drDU.GetValue(18) != System.DBNull.Value)this.dtTiempoVivienda.Value = drDU.GetDateTime (18);
						if (drDU.GetValue(19) != System.DBNull.Value)this.txtNombreArrendador.Text = drDU.GetString (19);
						if (drDU.GetValue(20) != System.DBNull.Value)this.txtTelfArrendador.Text = drDU.GetString (20);
						if (drDU.GetValue(25) != System.DBNull.Value)this.txtCelArrendador.Text = drDU.GetString (25);
						if (drDU.GetValue(26) != System.DBNull.Value)this.cmbTiempoVivienda.Value = drDU.GetInt32 (26);
					}
					drDU.Close();
					this.TipoVivienda();

					#endregion Datos_Ubicacion

					#region Conyugue
					if (this.cmbCreditoEstado.Value == null) this.cmbCreditoEstado.Value = 1;
					int idEstadoCivil = 0;
					if (this.cmbEstadoCivil.Value != null) idEstadoCivil = (int)this.cmbEstadoCivil.Value;
					//					if (idEstadoCivil != 2 && idEstadoCivil != 3 && idEstadoCivil != 5)
					//					{
					string ssqlDCY= string.Format ("Exec Cre_ConsultaConyugue {0},0",(int)this.txtIdCre_Solicitud.Value ); 
					SqlDataReader drDCY = Funcion.rEscalarSQL(cdsSeteoF,ssqlDCY,true);
					drDCY.Read();
					if (drDCY.HasRows)
					{	
						this.txtIdCre_Conyugue.Value = drDCY.GetInt32 (0);
						if (drDCY.GetValue(3) != System.DBNull.Value)this.txtDPCTNombreEmpresa.Text = drDCY.GetString (3);
						if (drDCY.GetValue(4) != System.DBNull.Value)this.cmbDPCTActividadEmpresa.Value = drDCY.GetInt32 (4);
						if (drDCY.GetValue(5) != System.DBNull.Value)this.cmbDPCTCargo.Value = drDCY.GetInt32 (5);
					}
					drDCY.Close();
					#region Datos Generales de Conyugue
						
					string ssqlDCGY= string.Format ("Exec Cre_ConsultaDatosGenerales {0},1",(int)this.txtIdCre_Conyugue.Value ); 
					SqlDataReader drDCGY = Funcion.rEscalarSQL(cdsSeteoF,ssqlDCGY,true);
					drDCGY.Read();

					if (drDCGY.HasRows)
					{	
							
						this.txtConyugueIdCre_DatosGenerales.Value = drDCGY.GetInt32 (0);
						if (drDCGY.GetValue(3) != System.DBNull.Value)this.cmbDPCTipodocumento.Value = drDCGY.GetInt32 (3);
						if (drDCGY.GetValue(4) != System.DBNull.Value)this.txtDPCdocumento.Text = drDCGY.GetString (4);
						if (drDCGY.GetValue(5) != System.DBNull.Value)this.txtDPCApellidoP.Text = drDCGY.GetString ( 5);
						if (drDCGY.GetValue(6) != System.DBNull.Value)this.txtDPCApellidoM.Text = drDCGY.GetString (6);
						if (drDCGY.GetValue(7) != System.DBNull.Value)this.txtDPCPrimerN.Text = drDCGY.GetString ( 7);
						if (drDCGY.GetValue(8) != System.DBNull.Value)this.txtDPCsegundoN.Text = drDCGY.GetString ( 8);
						if (drDCGY.GetValue(14) != System.DBNull.Value)this.cmbDPCNacionalidad.Value = drDCGY.GetInt32 (14);
						if (drDCGY.GetValue(15) != System.DBNull.Value)this.dtDPCFnacimiento.Value = drDCGY.GetDateTime ( 15);
						if (drDCGY.GetValue(16) != System.DBNull.Value)this.cmbDPCSexo.Value = drDCGY.GetInt32 (16);
						if (drDCGY.GetValue(17) != System.DBNull.Value)this.cmbDPCProvinciaN.Value = drDCGY.GetInt32 (17);
						if (drDCGY.GetValue(18) != System.DBNull.Value)this.cmbDPCCantonN.Value = drDCGY.GetInt32 (18);
						if (drDCGY.GetValue(21) != System.DBNull.Value)this.cmbDPCNivelEducacion.Value = drDCGY.GetInt32 (21);
						if (drDCGY.GetValue(22) != System.DBNull.Value)this.cmbDPCProfesion.Value = drDCGY.GetInt32 (22);
						if (drDCGY.GetValue(24) != System.DBNull.Value)this.cmbDPCSituacionL.Value = drDCGY.GetInt32 (24);
						SituacionLaboralConyuge((int)this.cmbDPCSituacionL.Value);
						if (drDCGY.GetValue(25) != System.DBNull.Value)this.cmbDPCActividadE.Value = drDCGY.GetInt32 (25);                                  
							
					}
					drDCGY.Close();
					#endregion Datos Generales de Conyugue

					#region Datos Ubicacion de Conyugue
						
	
					string ssqlDCGYT= string.Format ("Exec Cre_ConsultaDatosUbicacion {0},1",(int)this.txtIdCre_Conyugue.Value ); 
					SqlDataReader drDCGYT = Funcion.rEscalarSQL(cdsSeteoF,ssqlDCGYT,true);
					drDCGYT.Read();

						
					if (drDCGYT.HasRows)
					{
						this.txtConyugueIdCre_DatosUbicacion.Value = drDCGYT.GetInt32 (0);
						if (drDCGYT.GetValue (3) != System.DBNull.Value)this.cmbDPCTProvincia.Value = drDCGYT.GetInt32 (3);
						if (drDCGYT.GetValue (4) != System.DBNull.Value)this.cmbDPCTCanton.Value = drDCGYT.GetInt32 (4);
						if (drDCGYT.GetValue (5) != System.DBNull.Value)this.cmbDPCTParroquia.Value = drDCGYT.GetInt32 (5);
						if (drDCGYT.GetValue (6) != System.DBNull.Value)this.cmbDPCTBarrio.Value = drDCGYT.GetInt32 (6);

						if (drDCGYT.GetValue (7) != System.DBNull.Value)this.txtDPCTCallePrincipal.Text =  drDCGYT.GetString (7);
						if (drDCGYT.GetValue (8) != System.DBNull.Value)this.txtDPCTNumeroCasa.Text = drDCGYT.GetString (8);
						if (drDCGYT.GetValue (9) != System.DBNull.Value)this.txtDPCTCalleSecundaria.Text =  drDCGYT.GetString (9);
						if (drDCGYT.GetValue (10) != System.DBNull.Value)this.txtDPCTReferencia.Text = drDCGYT.GetString (10);
						if (drDCGYT.GetValue (11) != System.DBNull.Value)this.txtDPCTtelefono.Text = drDCGYT.GetString (11);
						if (drDCGYT.GetValue (12) != System.DBNull.Value)this.txtDPCTExt.Text = drDCGYT.GetString (12);
						if (drDCGYT.GetValue (15) != System.DBNull.Value)this.txtDPCTCelularC.Text = drDCGYT.GetString (15);
						if (drDCGYT.GetValue (16) != System.DBNull.Value)this.txtDPCTEmail.Text = drDCGYT.GetString (16);

						if (drDCGYT.GetValue(21) != System.DBNull.Value)this.txtDPCTIngresos.Value = drDCGYT.GetInt32 (21);
						if (drDCGYT.GetValue(22) != System.DBNull.Value)this.txtDPCTGastos.Value = drDCGYT.GetInt32 (22);
						if (drDCGYT.GetValue(23) != System.DBNull.Value)this.txtDPCTActivos.Value = drDCGYT.GetInt32 (23);
						if (drDCGYT.GetValue(24) != System.DBNull.Value)this.txtDPCTPasivos.Value = drDCGYT.GetInt32 (24);
					}

					drDCGYT.Close();
					#endregion Datos Ubicacion de Conyugue
					//						this.btnConyugueTelefono.Enabled = true;
					//						this.btnConyugueCelular.Enabled = true;
					//					}
					#endregion Conyugue

					string sSQLOSol = string.Format("Exec Cre_ConsultaReferencias {0},1", (int)this.txtClienteIdCre_DatosGenerales.Value);
					this.ultgconsulta.DataSource = FuncionesProcedimientos.dtGeneral(sSQLOSol);
					
					#region Tarjeta
					string sSQLOSolTar = string.Format("Exec Cre_ConsultaReferencias {0},3", (int)this.txtClienteIdCre_DatosGenerales.Value);
					this.ultraGrid1.DataSource = FuncionesProcedimientos.dtGeneral(sSQLOSolTar);
					#endregion Tarjeta

					//NovaScoring((int)this.txtClienteIdCre_DatosGenerales.Value);
          
			#region Verificacion Terrena
			string sSQLOSolVeri = string.Format("Exec ConsultaVerificacionTerrena {0}", (int)this.txtIdCre_Solicitud.Value);
			this.utgVerificacionTerrena.DataSource = FuncionesProcedimientos.dtGeneral(sSQLOSolVeri);
					
			#endregion Verificacion Terrena


					#region Negocios
					int idFuenteIngresos = 0;
					//					aqui va el el otro combo que no est acreado  cmbClienteFuenteDeIngresos
					if (this.cmbSituacionLaboral.Value != null) idFuenteIngresos = (int) this.cmbSituacionLaboral.Value;
					int idEntidadFinanciera = 0;
					if (this.cmbEntidad.Value != null) idEntidadFinanciera = (int)this.cmbEntidad.Value;

					//					if ((idEntidadFinanciera == 2 && idFuenteIngresos > 1) || (idEntidadFinanciera != 2 && (idFuenteIngresos != 1 && idFuenteIngresos != 4)))
					//					{
					#region Informacion General Negocios

					string ssqlCTNG= string.Format ("Exec Cre_ConsultaNegocions {0}",(int)this.txtClienteIdCre_DatosGenerales.Value ); 
					SqlDataReader drDCTNG = Funcion.rEscalarSQL(cdsSeteoF,ssqlCTNG,true);
					drDCTNG.Read();
			if (drDCTNG.HasRows)
			{	
				this.txtClienteIdCre_Negocio.Value = drDCTNG.GetInt32 (0);
				if (drDCTNG.GetValue(2) != System.DBNull.Value)this.txtNegocioRuc.Text = drDCTNG.GetString (2);
				if (drDCTNG.GetValue(3) != System.DBNull.Value)this.txtNombreDelNegocio.Text = drDCTNG.GetString (3);
				if (drDCTNG.GetValue(4) != System.DBNull.Value)this.dtpNegocioFechaInicio.Value = drDCTNG.GetDateTime (4);
				if (drDCTNG.GetValue(5) != System.DBNull.Value)this.cmbNegciotipo.Value = drDCTNG.GetInt32 ( 5);
				if (drDCTNG.GetValue(6) != System.DBNull.Value)this.cmbNEgocioActivdadEconomica.Value = drDCTNG.GetInt32 (6);
				if (drDCTNG.GetValue(7) != System.DBNull.Value)this.txtNegocioMetros.Value  = drDCTNG.GetInt32 (7);
				if (drDCTNG.GetValue(8) != System.DBNull.Value)this.txtNegocioEmpleados.Value = drDCTNG.GetInt32 (8);
				if (drDCTNG.GetValue(9) != System.DBNull.Value)this.chkNegocioObligadoaLlevar.Checked = drDCTNG.GetBoolean (9);
				if (drDCTNG.GetValue(10) != System.DBNull.Value)this.chkAfiliadoTributario.Checked =  drDCTNG.GetBoolean(10);
				if (drDCTNG.GetValue(11) != System.DBNull.Value)this.txtNotas.Text = drDCTNG.GetString (11);
				if (drDCTNG.GetValue(12) != System.DBNull.Value)this.cmbTiempo.Value = drDCTNG.GetInt32 (12);
			}
					drDCTNG.Close();
						
						
					#endregion Informacion General Negocios

					#region Datos Ubicacion de Negocio
				
					string ssqlCDUN= string.Format ("Exec Cre_ConsultaDatosUbicacion {0},3",(int)this.txtClienteIdCre_Negocio.Value ); 
					SqlDataReader drCDUN = Funcion.rEscalarSQL(cdsSeteoF,ssqlCDUN,true);
					drCDUN.Read();
			if (drCDUN.HasRows)
			{	
				this.txtNegocioIdCre_DatosUbicacion.Value = drCDUN.GetInt32 (0);
				if (drCDUN.GetValue(3) != System.DBNull.Value)this.cmbNegocioProvincia.Value = drCDUN.GetInt32  (3);
				if (drCDUN.GetValue(4) != System.DBNull.Value)	this.cmbNegocioCanton.Value = drCDUN.GetInt32 (4);
				if (drCDUN.GetValue(5) != System.DBNull.Value)this.cmbNegocioParroquia.Value = drCDUN.GetInt32  (5);
				if (drCDUN.GetValue(6) != System.DBNull.Value)this.cmbNegocioBarrio.Value = drCDUN.GetInt32 (6);
				if (drCDUN.GetValue(7) != System.DBNull.Value)this.txtNegocioCallePrincipal.Text = drCDUN.GetString  (7);
				if (drCDUN.GetValue(8) != System.DBNull.Value)this.txtNegocioNumeroCasa.Text = drCDUN.GetString (8);
				if (drCDUN.GetValue(9) != System.DBNull.Value)this.txtNegocioCalleSecundaria.Text = drCDUN.GetString  (9);
				if (drCDUN.GetValue(10) != System.DBNull.Value)this.txtNegocioReferncia.Text = drCDUN.GetString  (10);
				if (drCDUN.GetValue(11) != System.DBNull.Value)this.txtNegocioTelefono1.Text = drCDUN.GetString (11);
				if (drCDUN.GetValue(13) != System.DBNull.Value)this.txtNegocioTelefono2.Text = drCDUN.GetString (13);
				if (drCDUN.GetValue(15) != System.DBNull.Value)this.txtNEgocioCelular.Text = drCDUN.GetString  (15);
				if (drCDUN.GetValue(16) != System.DBNull.Value)this.txtNegocioemail.Text = drCDUN.GetString  (16);
				if (drCDUN.GetValue(21) != System.DBNull.Value)this.txtNEgocioIngresos.Value = drCDUN.GetInt32 (21);
				if (drCDUN.GetValue(22) != System.DBNull.Value)this.txtNEgocioGastos.Value = drCDUN.GetInt32 (22);
				if (drCDUN.GetValue(23) != System.DBNull.Value)this.txtNegocioActivos.Value = drCDUN.GetInt32 (23);
				if (drCDUN.GetValue(24) != System.DBNull.Value)this.txtNEgocioPasivos.Value = drCDUN.GetInt32 (24);
			}
					drCDUN.Close();
							
						
					#endregion Datos Ubicacion de Negocio

					//						this.btnNegocioTelefono1.Enabled = true;
					//						this.btnNegocioTelefono2.Enabled = true;
					//						this.btnNegocioCelular.Enabled = true;
					//					}
					#endregion Negocios

					#region Empresas
					//					if ((this.cmbEntidad.Value == 2 && this.cmbClienteFuenteDeIngresos. == 1) ||
					//					this.cmbSituacionLaboral.Value = 1;

					//					if (((int)this.cmbEntidad.Value == 2 && (int)this.cmbSituacionLaboral.Value == 1) || ((int)this.cmbEntidad.Value != 2 && ((int)this.cmbSituacionLaboral.Value == 1 || (int)this.cmbSituacionLaboral.Value == 3)))
					//					{
					#region Informacion General Empresas
					
					string sSqlIGES= string.Format ("Exec Cre_ConsultaTrabajo {0}",(int)this.txtClienteIdCre_DatosGenerales.Value ); 
					SqlDataReader drIGES = Funcion.rEscalarSQL(cdsSeteoF,sSqlIGES,true);
					drIGES.Read();

					if (drIGES.HasRows)
					{	

						this.txtClienteIdCre_EmpresasTrabajo.Value = drIGES.GetInt32 (0);
						if (drIGES.GetValue(2) != System.DBNull.Value)this.txtDependienteRuc.Text = drIGES.GetString (2);
						if (drIGES.GetValue(3) != System.DBNull.Value)this.txtDependienteEmpresa.Text = drIGES.GetString (3);
						if (drIGES.GetValue(4) != System.DBNull.Value)this.cmbDependienteTipoempresa.Value=  drIGES.GetInt32 (4);
						if (drIGES.GetValue(5) != System.DBNull.Value)this.dtDependienteFechaIngreso.Value = drIGES.GetDateTime (5);
						if (drIGES.GetValue(7) != System.DBNull.Value)this.cmbDependienteTipocontrato.Value = drIGES.GetInt32 (7);
						if (drIGES.GetValue(8) != System.DBNull.Value)this.cmbDependienteTipoSueldo.Value = drIGES.GetInt32 (8);
						if (drIGES.GetValue(9) != System.DBNull.Value)this.chkAfiliadoIESS.Checked = drIGES.GetBoolean (9);
						if (drIGES.GetValue(10) != System.DBNull.Value)this.txtDependienteDiasPago.Value = drIGES.GetInt32 (10);
						if (drIGES.GetValue(11) != System.DBNull.Value)this.txtDependienteDepartamento.Text = drIGES.GetString (11);
						if (drIGES.GetValue(12) != System.DBNull.Value)this.cmbDependienteCargo.Value = drIGES.GetInt32 (12);
						if (drIGES.GetValue(15) != System.DBNull.Value)this.cmbTiempoDependiente.Value = drIGES.GetInt32 (15);
					}
					drIGES.Close();
						
					#endregion Informacion General Empresas

					#region Datos Ubicacion de Trabajo
					

					string sSqlDUTs= string.Format ("Exec Cre_ConsultaDatosUbicacion {0},4 ",(int)this.txtClienteIdCre_EmpresasTrabajo.Value); 
					SqlDataReader drDUTs= Funcion.rEscalarSQL(cdsSeteoF,sSqlDUTs,true);
					drDUTs.Read();

					if (drDUTs.HasRows)
					{	
						this.txtEmpresaIdCre_DatosUbicacion.Value =  drDUTs.GetInt32(0);
						if (drDUTs.GetValue(3) != System.DBNull.Value)this.cmbDependienteProvincia.Value =   drDUTs.GetInt32(3);
						if (drDUTs.GetValue(4) != System.DBNull.Value)this.cmbDependienteCanton.Value =  drDUTs.GetInt32(4);
						if (drDUTs.GetValue(5) != System.DBNull.Value)this.cmbDependienteParroquia.Value =  drDUTs.GetInt32(5);
						if (drDUTs.GetValue(6) != System.DBNull.Value)this.cmbDependienteBarrio.Value =  drDUTs.GetInt32(6); 
							

						if (drDUTs.GetValue(7) != System.DBNull.Value)this.txtDependienteCallePrinciapl.Text =  drDUTs.GetString(7);
						if (drDUTs.GetValue(8) != System.DBNull.Value)this.txtDependienteNumeroCasa.Text =  drDUTs.GetString(8);
						if (drDUTs.GetValue(9) != System.DBNull.Value)this.txtDependienteCalleSecundaria.Text =   drDUTs.GetString(9);
						if (drDUTs.GetValue(10) != System.DBNull.Value)this.txtDependienteReferencia.Text =  drDUTs.GetString(10);
						if (drDUTs.GetValue(11) != System.DBNull.Value)this.txtDependienteTelefono.Text =   drDUTs.GetString(11);
						if (drDUTs.GetValue(12) != System.DBNull.Value)this.txtDependienteEXT.Text =  drDUTs.GetString(12);
						if (drDUTs.GetValue(15) != System.DBNull.Value)this.txtDependienteCelular.Text =   drDUTs.GetString(15);
						if (drDUTs.GetValue(16) != System.DBNull.Value)this.txtDependienteemail.Text =  drDUTs.GetString(16);

						if (drDUTs.GetValue(21) != System.DBNull.Value)this.txtDependienteIngresos.Value =   drDUTs.GetInt32(21);
						if (drDUTs.GetValue(22) != System.DBNull.Value)this.txtDependienteGastos.Value =  drDUTs.GetInt32(22);
						if (drDUTs.GetValue(23) != System.DBNull.Value)this.txtDependienteActivos.Value =  drDUTs.GetInt32(23);
						if (drDUTs.GetValue(24) != System.DBNull.Value)this.txtDependientePasivos.Value =   drDUTs.GetInt32(24);
					}
					drDUTs.Close();
					#endregion Datos Ubicacion de Trabajo

					//						this.btnTrabajoTelefono.Enabled = true;
					//						this.btnTrabajoCelular.Enabled = true;
					//					}
					#endregion Empresas
            
					string sSQLHCRED = string.Format("Exec Cre_ConsultaHistorialCreditos '{0}'", this.txtDocumento.Text);
					this.utgHistorialCredito.DataSource = FuncionesProcedimientos.dtGeneral(sSQLHCRED);
					bCarga = false; 
					
					this.gdObservaciones.DataSource = Funcion.dvProcedimiento(cdsSeteoF, string.Format("Exec Cre_ConsultaObservacionesSolicitud '{0}'",(int)this.txtIdCre_Solicitud.Value));
			#region Equifax
			 int iEquifax = Funcion.iEscalarSQL(cdsSeteoF, string.Format("select count(1) from coordenadasequifax where Identificacion ='{0}'",this.txtDocumento.Text), true);
			  if( iEquifax > 0 ) this.btnEquifax.Visible = true;
			#endregion Equifax
      
      
 
//				}
//
//				else
//				{
//					this.txtDocumento.Text = "";
//					this.btnImprimir.Enabled = false;
//					this.lblEstado.Text = "PENDIENTE";
//				}
//			}
		}

		private void txtDocumento_Validating(object sender, System.ComponentModel.CancelEventArgs e)
		{
			if (this.cmbTipoRuc.ActiveRow != null)
			{
				int iTipoRuc = 5;
				if ((int) this.cmbTipoRuc.Value == 1) iTipoRuc = 5;// cedula
				if ((int) this.cmbTipoRuc.Value == 2) iTipoRuc = 6;// pasaporte
				if ((int) this.cmbTipoRuc.Value == 3) iTipoRuc = 4;// ruc
				if (!Validacion.vbIdentificacion(this.txtDocumento, iTipoRuc, cdsSeteoF)) 
				{
					e.Cancel = true;   
				}
			}
		}
		#region Valida Caracteres Especiales
		private void groupBox6_Enter(object sender, System.EventArgs e)
		{
		
		}

		private void groupBox2_Enter(object sender, System.EventArgs e)
		{
		
		}

		private void label90_Click(object sender, System.EventArgs e)
		{
		
		}

		private void label120_Click(object sender, System.EventArgs e)
		{
		
		}

		private void label125_Click(object sender, System.EventArgs e)
		{
		
		}

		private void txtDependienteTelefono_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
			validar.Solonumeros(e);
		}

		private void txtDependienteEXT_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
			validar.Solonumeros(e);
		}

		private void txtDependienteCelular_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
			validar.Solonumeros(e);
		}

		private void txtNegocioTelefono1_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
			validar.Solonumeros(e);
		}

		private void txtNegocioTelefono2_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
			validar.Solonumeros(e);
		}

		private void txtNEgocioCelular_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
			validar.Solonumeros(e);
		}

		private void txtDPCTtelefono_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
			validar.Solonumeros(e);
		}

		private void txtDPCTExt_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
			validar.Solonumeros(e);
		}

		private void txtDPCTCelularC_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
			validar.Solonumeros(e);
		}

		private void txtDDTelefono1_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
			validar.Solonumeros(e);
		}

		private void txtDDTelefono2_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
			validar.Solonumeros(e);
		}

		private void txtDDCelular_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
			validar.Solonumeros(e);
		}

		private void txtDPCApellidoP_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
			validar.SoloLetras(e);
		}

		private void txtDPCTNombreEmpresa_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
			validar.CaracteresEspeciales(e);
		}

		private void txtDPCTCallePrincipal_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
			validar.CaracteresEspeciales(e);
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

		private void txtDPCApellidoM_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
			validar.SoloLetras(e);
		}

		private void txtDPCPrimerN_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
			validar.SoloLetras(e);
		}

		private void txtDPCsegundoN_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
			validar.SoloLetras(e);
		}

		private void txtNombreDelNegocio_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
		validar.CaracteresEspeciales(e);
		}

		private void txtDPCTNumeroCasa_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
			validar.CaracteresEspeciales(e);
		}

		private void txtDPCTCalleSecundaria_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
			validar.CaracteresEspeciales(e);
		}

		private void txtDDCallaPrincipal_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
			validar.CaracteresEspeciales(e);
		}

		private void txtDDNumeroCasa_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
			validar.CaracteresEspeciales(e);
		}

		private void txtDDCalleSecundaria_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
			validar.CaracteresEspeciales(e);
		}

		private void txtDDReferenciasU_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
			validar.CaracteresEspeciales(e);
		}

		private void txtNombreArrendador_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
			validar.CaracteresEspeciales(e);
		}

		private void txtTelfArrendador_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
			validar.Solonumeros(e);
		}

		private void txtDPCTReferencia_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
			validar.CaracteresEspeciales(e);
		}

		private void txtNegocioRuc_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
			validar.Solonumeros(e);
		}

		private void txtNegocioCallePrincipal_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
			validar.CaracteresEspeciales(e);
		}

		private void txtNegocioNumeroCasa_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
			validar.CaracteresEspeciales(e);
		}

		private void txtNegocioCalleSecundaria_ValueChanged(object sender, System.EventArgs e)
		{
		
		}

		private void txtNegocioCalleSecundaria_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
			validar.CaracteresEspeciales(e);
		}

		private void txtNegocioReferncia_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
			validar.CaracteresEspeciales(e);
		}

		private void txtDependienteRuc_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
			validar.Solonumeros(e);
		}

		private void txtDependienteEmpresa_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
			validar.CaracteresEspeciales(e);
		}

		private void txtDependienteDepartamento_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
			validar.CaracteresEspeciales(e);
		}

		private void txtDependienteCallePrinciapl_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
			validar.CaracteresEspeciales(e);
		}

		private void txtDependienteNumeroCasa_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
			validar.CaracteresEspeciales(e);
		}

		private void txtDependienteCalleSecundaria_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
			validar.CaracteresEspeciales(e);
		}

		private void txtDependienteReferencia_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
//			validar.CaracteresEspeciales(e);
		}

		private void txtCreditoObservaciones_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
			validar.CaracteresEspeciales(e);
		}

		private void txtVehiculoMarca_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
			//validar.CaracteresEspeciales(e);
		}

		private void txtModelo_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
			//validar.CaracteresEspeciales(e);
		}
			#endregion Valida Caracteres Especiales

		private void Valida_DatosCliente ()
		{
			#region Valida_DatosCliente

			if (!Validacion.vbTexto(this.txtApellidoP, 3, 60, "Ingrese el Apellido Paterno")) return;
			if (!Validacion.vbTexto(this.txtApellidoP, 3, 60, "Ingrese el Primer Nombre")) return;
			if (!Validacion.vbComboVacio(this.cmbEntidad, "Seleccione la Entidad Financiera ")) return;
			if (!Validacion.vbComboVacio(this.cmbBodega, "Seleccione el Local ")) return;
			if (!Validacion.vbComboVacio(this.cmbNacionalidad, "Seleccione la Nacionalidad ")) return;
			if (!Validacion.vbFechaEnDocumentos(this.dtFechaNacimiento, "Ingrese la fecha de Nacimiento", true, "Cobros", cdsSeteoF)) return;
			DateTime dtFecha = DateTime.Today.AddYears(-18);
			if (Convert.ToDateTime(this.dtFechaNacimiento.Text) > dtFecha)
			{
				MessageBox.Show("La Edad Minima del Cliente debe de ser 18 años", "Validacion", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				this.dtFechaNacimiento.Focus();
				return;
			}
			dtFecha = DateTime.Today.AddYears(-70);
			if (Convert.ToDateTime(this.dtFechaNacimiento.Text) < dtFecha)
			{
				MessageBox.Show("La Edad Maxima del Cliente debe de ser 70 años", "Validacion", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				this.dtFechaNacimiento.Focus();
				return;
			}
//			if(!Validacion.vbComparaFechas((DateTime)this.dtFechaNacimiento.Value,(DateTime)this.dtTiempoVivienda.Value,this.dtTiempoVivienda,"La Fecha de vivienda no puede ser menor a la Fecha de Nacimiento",false))return;
      if (!Validacion.vbComboVacio(this.cmbSexo, "Seleccione el Genero ")) return;
			if (!Validacion.vbComboVacio(this.cmbTiempoVivienda, "Seleccione el Tiempo Vivienda ")) return;
			if (!Validacion.vbComboVacio(this.cmbProvinciaNAcimiento, "Seleccione la Provincia de Nacimiento ")) return;
			if (!Validacion.vbComboVacio(this.cmbCantonNacimiento, "Seleccione el Canton de Nacimiento ")) return;
			if (!Validacion.vbComboVacio(this.cmbEstadoCivil, "Seleccione el Estado Civil ")) return;
			if (!Validacion.vbCampoEnteroVacio(this.txtDependientes,"Ingrese el Valor Dependiente ",0,10000))return;
			if (!Validacion.vbComboVacio(this.cmbNivelEducacion, "Seleccione el Nivel de Educación ")) return;
			if (!Validacion.vbComboVacio(this.cmbProfesion, "Seleccione la Profeción ")) return;
//			cmbSituacionLaboral
//			cmbActividadEconomica
			if (!Validacion.vbComboVacio(this.cmbDDProvincia, "Seleccione la Provincia del Domicilio")) return;
			if (!Validacion.vbComboVacio(this.cmbDDCanton, "Seleccione el Canton del Domicilio ")) return;
			if (!Validacion.vbComboVacio(this.cmbDDParroquia, "Seleccione la Parroquia del Domicilio ")) return;
			if (!Validacion.vbComboVacio(this.cmbDDBarrio, "Seleccione el Barrio del Domicilio ")) return;
			if (!Validacion.bEMail(this.txtDDEmail.Text.ToString())) return;
			if (!Validacion.vbTexto(this.txtDDCallaPrincipal, 3, 100, "Ingrese la Calle Principal")) return;
			if (!Validacion.vbTexto(this.txtDDNumeroCasa, 3, 60, "Ingrese el Numero de Casa")) return;
			if (!Validacion.vbTexto(this.txtDDCalleSecundaria, 3, 100, "Ingrese la Calle Secundaria")) return;
			if (this.txtDDTelefono1.Text.ToString().Trim().Length == 0 &&  this.txtDDCelular.Text.ToString().Trim().Length == 0 )
			{
				MessageBox.Show(string.Format("Debe Ingresar el Telefono Convencional o el Telefono Celular"), "Point Technology", MessageBoxButtons.OK,MessageBoxIcon.Stop);
				this.txtDDTelefono1.Focus();
				return;	
			}
			if (!Validacion.vbValidaTelefono(this.txtDDTelefono1, 1, false)) return;
			if (!Validacion.vbValidaTelefono(this.txtDDTelefono2, 1, false)) return;
			if (!Validacion.vbValidaTelefono(this.txtDDCelular, 2, false)) return;
			if (!Validacion.vbTexto(this.txtDDReferenciasU, 10, 999, "Ingrese la Referencia de Ubicación")) return;
			if (!Validacion.vbComboVacio(this.cmbDDTipoVivienda, "Seleccione el Tipo de Vivienda ")) return;
			if (!Validacion.vbComboVacio(this.cmbTiempoVivienda, "Seleccione el Tiempo Vivienda ")) return;
//			if (!Validacion.vbFechaEnDocumentos(this.dtTiempoVivienda, "Ingrese el Tiempo de Vivienda", true, "Cobros", cdsSeteoF)) return;
			//if (!Validacion.vbValidaTelefono(this.txtTelfArrendador, 1, false)) return;
			if (!Validacion.vbValidaTelefono(this.txtCelArrendador, 1, false)) return;

     #endregion Valida_DatosCliente
		}

		private void Valida_Conyuge ()
		{
			#region Valida_DatosConyuge

      if (!Validacion.vbComboVacio(this.cmbDPCTipodocumento, "Seleccione el Tipo de Documento")) return;
			if (!Validacion.vbTexto(this.txtDPCdocumento, 3, 60, "Ingrese la Cedula del Conyuge ")) return;
			if (!Validacion.vbTexto(this.txtDPCApellidoP, 3, 60, "Ingrese el Apellido Paterno del Conyuge")) return;
			if (!Validacion.vbTexto(this.txtDPCPrimerN, 3, 60, "Ingrese el Primer Nombre del Conyuge")) return;
			if (!Validacion.vbComboVacio(this.cmbDPCNacionalidad, "Seleccione la Nacionalidad del Conyuge ")) return;
			if (!Validacion.vbFechaEnDocumentos(this.dtDPCFnacimiento, "Ingrese la fecha de Nacimiento del Conyuge ", true, "Cobros", cdsSeteoF)) return;
			if (!Validacion.vbComboVacio(this.cmbDPCSexo, "Seleccione el Genero del Conyuge ")) return;
			if (!Validacion.vbComboVacio(this.cmbDPCProvinciaN, "Seleccione la Provincia de Nacimiento del Conyuge ")) return;
			if (!Validacion.vbComboVacio(this.cmbDPCCantonN, "Seleccione el Canton de Nacimiento del Conyuge ")) return;
			if (!Validacion.vbComboVacio(this.cmbDPCNivelEducacion, "Seleccione el Nivel de Eduación del Conyuge ")) return;
			if (!Validacion.vbComboVacio(this.cmbDPCProfesion, "Seleccione la Profesión del Conyuge ")) return;

			if (!Validacion.vbTexto(this.txtApellidoP, 3, 60, "Ingrese el Apellido Paterno")) return;
			if (!Validacion.vbTexto(this.txtApellidoP, 3, 60, "Ingrese el Primer Nombre")) return;
			if (!Validacion.vbComboVacio(this.cmbEntidad, "Seleccione la Entidad Financiera ")) return;
			if (!Validacion.vbComboVacio(this.cmbBodega, "Seleccione el Local ")) return;
			if (!Validacion.vbComboVacio(this.cmbNacionalidad, "Seleccione la Nacionalidad ")) return;

			if (this.txtDDTelefono1.Text.ToString().Trim().Length == 0  )
			{
				MessageBox.Show(string.Format("Debe Ingresar el Telefono Convencional del Conyuge"), "Point Technology", MessageBoxButtons.OK,MessageBoxIcon.Stop);
				this.txtDPCTtelefono.Focus();
				return;	
			}
			if ( this.txtDPCTCelularC.Text.ToString().Trim().Length == 0  )
			{
				MessageBox.Show(string.Format("Debe Ingresar el Celular  del Conyuge"), "Point Technology", MessageBoxButtons.OK,MessageBoxIcon.Stop);
				this.txtDPCTCelularC.Focus();
				return;	
			}
			if (!Validacion.vbValidaTelefono(this.txtDPCTtelefono, 1, false)) return;
			if (!Validacion.vbValidaTelefono(this.txtDPCTExt, 1, false)) return;
			if (!Validacion.vbValidaTelefono(this.txtDPCTCelularC, 2, false)) return;
			if (!Validacion.bEMail(this.txtDPCTEmail.Text.ToString())) return;

			#endregion Valida_DatosConyuge	

		}

		private void Valida_SeccionNegocio ()
		{
			#region Valida_SeccionNEgocio

			if (!Validacion.vbTexto(this.txtNegocioRuc, 3, 13, "Ingrese el Ruc del Negocio")) return;
			if (!Validacion.vbTexto(this.txtNombreDelNegocio, 3, 200, "Ingrese el Nombre del Negocio")) return;
//			if (!Validacion.vbFechaEnDocumentos(this.dtpNegocioFechaInicio, "Ingrese la fecha de inicio del Negocio ", true, "Cobros", cdsSeteoF)) return;
			if (!Validacion.vbComboVacio(this.cmbTiempo, "Seleccione el Fecha de Inicio ")) return;

			if (!Validacion.vbCampoEnteroVacio(this.txtNegocioMetros,"Ingrese los Metros del Negocio ",1,10000))return;
			if (!Validacion.vbCampoEnteroVacio(this.txtNegocioEmpleados,"Ingrese el Número de Empleados del Negocio ",1,500))return;
      if (!Validacion.bEMail(this.txtNegocioemail.Text.ToString())) return;
			if (!Validacion.vbTexto(this.txtNegocioCallePrincipal, 3, 100, "Ingrese la Calle Principal del Negocio")) return;
			if (!Validacion.vbTexto(this.txtNegocioNumeroCasa, 3, 60, "Ingrese el Numero de Casa del Negocio")) return;
			if (!Validacion.vbTexto(this.txtNegocioCalleSecundaria, 3, 100, "Ingrese la Calle Secundaria del Negocio")) return;

			if (this.txtNegocioTelefono1.Text.ToString().Trim().Length == 0 &&  this.txtNEgocioCelular.Text.ToString().Trim().Length == 0 )
			{
				MessageBox.Show(string.Format("Debe Ingresar el Telefono Convencional o el Telefono Celular del Negocio "), "Point Technology", MessageBoxButtons.OK,MessageBoxIcon.Stop);
				this.txtNegocioTelefono1.Focus();
				return;	
			}
			if (!Validacion.vbValidaTelefono(this.txtNegocioTelefono1, 1, false)) return;
			if (!Validacion.vbValidaTelefono(this.txtNegocioTelefono2, 1, false)) return;
			if (!Validacion.vbValidaTelefono(this.txtNEgocioCelular, 2, false)) return;
			if (!Validacion.vbTexto(this.txtNegocioReferncia, 10, 300, "Ingrese la Referencia de la Ubicación del Negocio")) return;
			if (!Validacion.vbCampoEnteroVacio(this.txtNEgocioIngresos,"Ingrese los Ingresos del Negocio ",450,1000000))return;
			if (!Validacion.vbCampoEnteroVacio(this.txtNEgocioGastos,"Ingrese los Gastos del Negocio  ",450,1000000))return;
			//if (!Validacion.vbCampoEnteroVacio(this.txtNegocioActivos,"Ingrese los Activos del Negocio ",450,1000000))return;
			//if (!Validacion.vbCampoEnteroVacio(this.txtNEgocioPasivos,"Ingrese los  PAsivos del Negocio ",450,1000000))return;

			#endregion Valida_SeccionNEgocio

		}

		private void Valida_SeccionDependiente ()
		{
			#region Valida_SeccionDependiente

      if (!Validacion.vbTexto(this.txtDependienteEmpresa, 10, 100, "Ingrese el Nombre de la Empresa")) return;
			if (!Validacion.vbComboVacio(this.cmbDependienteTipoempresa, "Seleccione el tipo de empresa ")) return;
			if (!Validacion.vbFechaEnDocumentos(this.dtDependienteFechaIngreso, "Ingrese la fecha de Ingreso a la Empresa ", true, "Cobros", cdsSeteoF)) return;
			//if (!Validacion.vbComboVacio(this.cmbTiempoDependiente, "Seleccione el Fecha de Ingreso ")) return;
			if (!Validacion.vbCampoEnteroVacio(this.txtDependienteIngresos,"Seleccione los Ingresos del Cliente  ",1,1000000))return;
			if (!Validacion.vbCampoEnteroVacio(this.txtDependienteGastos,"Ingrese los Gastos del Cliente ",1,1000000))return;
			//if (!Validacion.vbCampoEnteroVacio(this.txtDependienteActivos,"Ingrese los Activos del Cliente ",1,1000000))return;
			//if (!Validacion.vbCampoEnteroVacio(this.txtDependientePasivos,"Ingrese los  PAsivos del Cliente ",1,1000000))return;
			if (!Validacion.vbComboVacio(this.cmbDependienteTipocontrato, "Seleccione el Tipo de Contrato ")) return;
			if (!Validacion.vbComboVacio(this.cmbDependienteTipoSueldo, "Seleccione el Tipo de Sueldo ")) return;
			if (!Validacion.vbComboVacio(this.cmbDependienteTipoSueldo, "Seleccione el Cargo ")) return;
			if (!Validacion.vbCampoEnteroVacio(this.txtNegocioMetros,"Ingrese el Día de Pago ",1,31))return;
			if (!Validacion.vbComboVacio(this.cmbDependienteProvincia, "Seleccione la Provincia de su Trabajo ")) return;
			if (!Validacion.vbComboVacio(this.cmbDependienteCanton, "Seleccione el Canton de su Trabajo ")) return;
			if (!Validacion.vbComboVacio(this.cmbDependienteParroquia, "Seleccione la Parroquia de su Trabajo ")) return;
			if (!Validacion.vbComboVacio(this.cmbDependienteBarrio, "Seleccione el Barrio de su Trabajo ")) return;
			if (!Validacion.vbTexto(this.txtDependienteCallePrinciapl, 3, 100, "Ingrese la Calle Principal del Trabajo")) return;
			if (!Validacion.vbTexto(this.txtDependienteNumeroCasa, 3, 60, "Ingrese el Numero de Casa del Trabajo")) return;
			if (!Validacion.vbTexto(this.txtDependienteCalleSecundaria, 3, 100, "Ingrese la Calle Secundaria del Trabajo")) return;
		  if (!Validacion.bEMail(this.txtDependienteemail.Text.ToString())) return;
			if (this.txtDependienteTelefono.Text.ToString().Trim().Length == 0 &&  this.txtDependienteCelular.Text.ToString().Trim().Length == 0 )
			{
				MessageBox.Show(string.Format("Debe Ingresar el Telefono Convencional o el Telefono Celular del Trabajo "), "Point Technology", MessageBoxButtons.OK,MessageBoxIcon.Stop);
				this.txtNegocioTelefono1.Focus();
				return;	
			}
			if (!Validacion.vbValidaTelefono(this.txtDependienteTelefono, 1, false)) return;
			//if (!Validacion.vbValidaTelefono(this.txtDependienteEXT, 1, false)) return;
			if (!Validacion.vbValidaTelefono(this.txtDependienteCelular, 2, false)) return;
			if (!Validacion.vbTexto(this.txtNegocioReferncia, 10, 300, "Ingrese la Referencia de la Ubicación del Trabajo")) return;
     #endregion Valida_SeccionDependiente
		}

		private void btnGuardar_Click(object sender, System.EventArgs e)
		{
			#region Valida_DatosCliente
			if (!Validacion.vbTexto(this.txtDocumento, 1, 11, " Cedula ")) 
			{
				this.tbBasico.Tabs[0].Selected = true; 
				return;
			}

			if (!Validacion.vbTexto(this.txtApellidoP, 1, 60, " Paterno")) 
			{
				this.tbBasico.Tabs[0].Selected = true; 
				return;
			}
			if (!Validacion.vbTexto(this.txtPnombre,1, 60, " Nombre"))
			{
					this.tbBasico.Tabs[0].Selected = true; 
				return;
			}
			if (!Validacion.vbComboVacio(this.cmbEntidad, "Seleccione la Entidad Financiera "))
			{
				this.tbBasico.Tabs[0].Selected = true; 
				return;
			}
			if (!Validacion.vbComboVacio(this.cmbBodega, "Seleccione el Local ")) 
			{
					this.tbBasico.Tabs[0].Selected = true;
				return;
			}
			if (!Validacion.vbComboVacio(this.cmbNacionalidad, "Seleccione la Nacionalidad ")) 
			{
				this.tbBasico.Tabs[0].Selected = true;
				return;
			}
			if (!Validacion.vbFechaEnDocumentos(this.dtFechaNacimiento, "Ingrese la fecha de Nacimiento", true, "Cobros", cdsSeteoF))
			{
				this.tbBasico.Tabs[0].Selected = true; 
				return;
			}
			DateTime dtFecha = DateTime.Today.AddYears(-18);
			if (Convert.ToDateTime(this.dtFechaNacimiento.Text) > dtFecha)
			{
				MessageBox.Show("La Edad Minima del Cliente debe de ser 21 años", "Validacion", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				this.dtFechaNacimiento.Focus();
				this.tbBasico.Tabs[0].Selected = true;
				return;
			}
			dtFecha = DateTime.Today.AddYears(-75);
			if (Convert.ToDateTime(this.dtFechaNacimiento.Text) < dtFecha)
			{
				MessageBox.Show("La Edad Maxima del Cliente debe de ser 75 años", "Validacion", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				this.dtFechaNacimiento.Focus();
				this.tbBasico.Tabs[0].Selected = true;
				return;
			}
			if (!Validacion.vbComboVacio(this.cmbSexo, "Seleccione el Genero")) 
			{
					this.tbBasico.Tabs[0].Selected = true; 
				    return;
			}
			if (!Validacion.vbComboVacio(this.cmbProvinciaNAcimiento, "Seleccione la Provincia de Nacimiento ")) 
			{
				this.tbBasico.Tabs[0].Selected = true; 
				return;
			}
			if (!Validacion.vbComboVacio(this.cmbCantonNacimiento, "Seleccione el Canton de Nacimiento ")) 
			{
					this.tbBasico.Tabs[0].Selected = true; 
				return;
			}
			if (!Validacion.vbComboVacio(this.cmbEstadoCivil, "Seleccione el Estado Civil ")) 
			{
					this.tbBasico.Tabs[0].Selected = true; 
				    return;
			}
			if (!Validacion.vbCampoEnteroVacio(this.txtDependientes,"Ingrese el Valor Dependiente ",0,10000))
			{
					this.tbBasico.Tabs[0].Selected = true; 
				return;
			}
			if (!Validacion.vbComboVacio(this.cmbNivelEducacion, "Seleccione el Nivel de Educación "))
			{
				this.tbBasico.Tabs[0].Selected = true;  
				return;
			}
			if (!Validacion.vbComboVacio(this.cmbProfesion, "Seleccione la Profeción ")) 
			{
				this.tbBasico.Tabs[0].Selected = true; 
				return;
			}
			if ((int)this.cmbEntidad.Value == 2)
			{
				if (!Validacion.vbComboVacio(this.cmbSituacionLaboral, "Seleccione la Fuente de Ingreso  ")) 
				{
					this.tbBasico.Tabs[0].Selected = true; 
					return;
				}
			}
			else
			{
				if (!Validacion.vbComboVacio(this.cmbSituacionLaboral , "Seleccione la Situacion LAboral ")) 
				{
					this.tbBasico.Tabs[0].Selected = true; 
					return;
				}
				
			}

			if (!Validacion.vbComboVacio(this.cmbActividadEconomica , "Seleccione la Actividad Economica"))
			{
				this.tbBasico.Tabs[0].Selected = true; 
				return;
			}
			if (!Validacion.vbComboVacio(this.cmbDDProvincia, "Seleccione la Provincia del Domicilio"))
			{
				this.tbBasico.Tabs[0].Selected = true; 
				return;
			}
			if (!Validacion.vbComboVacio(this.cmbDDCanton, "Seleccione el Canton del Domicilio ")) 
			{
				this.tbBasico.Tabs[0].Selected = true; 
				return;
			}
			if (!Validacion.vbComboVacio(this.cmbDDParroquia, "Seleccione la Parroquia del Domicilio ")) 
			{
					this.tbBasico.Tabs[0].Selected = true; 
				return;
			}
			if (!Validacion.vbComboVacio(this.cmbDDBarrio, "Seleccione el Barrio del Domicilio "))
			{
					this.tbBasico.Tabs[0].Selected = true; 
				return;
			}
			if (this.txtDDEmail.Text.ToString().Trim().Length > 0 )
			{
				if (!Validacion.bEMail(this.txtDDEmail.Text.ToString())) 
				{
					this.tbBasico.Tabs[0].Selected = true;
					return;
				}
			
			}
			if (!Validacion.vbTexto(this.txtDDCallaPrincipal, 1, 100, "Ingrese la Calle Principal")) 
			{
				this.tbBasico.Tabs[0].Selected = true;
				return;
			}
			if (!Validacion.vbTexto(this.txtDDNumeroCasa, 1, 60, "Ingrese el Numero de Casa"))
			{
				this.tbBasico.Tabs[0].Selected = true;
				return;
			}
			if (!Validacion.vbTexto(this.txtDDCalleSecundaria, 1, 100, "Ingrese la Calle Secundaria")) 
			{
				this.tbBasico.Tabs[0].Selected = true;
				return;
			}
			if (this.txtDDTelefono1.Text.ToString().Trim().Length == 0 &&  this.txtDDCelular.Text.ToString().Trim().Length == 0 )
			{
				MessageBox.Show(string.Format("Debe Ingresar el Telefono Convencional o el Telefono Celular"), "Point Technology", MessageBoxButtons.OK,MessageBoxIcon.Stop);
				this.txtDDTelefono1.Focus();
				this.tbBasico.Tabs[0].Selected = true;
				return;	
			}
			if (!Validacion.vbValidaTelefono(this.txtDDTelefono1, 1, false))
			{
				this.tbBasico.Tabs[0].Selected = true;
				return;
			}
			if (!Validacion.vbValidaTelefono(this.txtDDTelefono2, 1, false))
			{
				this.tbBasico.Tabs[0].Selected = true; 
				return;
			}
			if (!Validacion.vbValidaTelefono(this.txtDDCelular, 2, false)) 
			{
				this.tbBasico.Tabs[0].Selected = true;
				return;
			}
			if (!Validacion.vbTexto(this.txtDDReferenciasU, 10, 1000, "Ingrese la Referencia de Ubicación"))
			{ 
				this.tbBasico.Tabs[0].Selected = true; 
				return;
			}
			if (!Validacion.vbComboVacio(this.cmbDDTipoVivienda, "Seleccione el Tipo de Vivienda "))
			{
				this.tbBasico.Tabs[0].Selected = true; 
				return;
			}
			if (!Validacion.vbComboVacio(this.cmbTiempoVivienda, "Seleccione el Tiempo Vivienda "))
			{
				this.tbBasico.Tabs[0].Selected = true; 
				return;
			}
//			if (!Validacion.vbFechaEnDocumentos(this.dtTiempoVivienda, "Ingrese el Tiempo de Vivienda", false, "", cdsSeteoF))
//			{
//				this.tbBasico.Tabs[0].Selected = true; 
//				return;
//			}
			if (!Validacion.vbValidaTelefono(this.txtCelArrendador, 2, false))
			{
				this.tbBasico.Tabs[0].Selected = true; 
				return;
			}
//			if(!Validacion.vbComparaFechas((DateTime)this.dtFechaNacimiento.Value,(DateTime)this.dtTiempoVivienda.Value,this.dtTiempoVivienda,"La Fecha de vivienda no puede ser menor a la Fecha de Nacimiento",false))
//			{
//				this.tbBasico.Tabs[0].Selected = true;
//				return;
//			}
       
			

			if ((int)this.cmbDDTipoVivienda.Value  == 3 || (int)this.cmbDDTipoVivienda.Value == 4)
			{
				if (!Validacion.vbComboVacio(this.cmbDDInmueble, "Seleccione el Inmueble "))
				{
					this.tbBasico.Tabs[0].Selected = true; 
					return;
				}
				if (!Validacion.vbComboVacio(this.cmbDDCiudadInmueble, "Seleccione la Ciudad Inmueble "))
				{
					this.tbBasico.Tabs[0].Selected = true; 
					return;
				}

				if (!Validacion.vbCampoEnteroVacio(this.txtDDValorInmueble,"Ingrese el valor del Inmueble ",0,1000000))
				{
					this.tbBasico.Tabs[0].Selected = true;
					return;
				}
				
			}

//			if ((int)this.cmbDDTipoVivienda.Value == 1)
//			{
////				if (!Validacion.vbTexto(this.txtNombreArrendador, 1, 50, "Ingrese el Nombre del Arrendador"))
////				{
////					this.tbBasico.Tabs[0].Selected = true;
////					return;
////				}
////				if (!Validacion.vbTexto(this.txtCelArrendador, 7, 10, "Ingrese el celular del Arrendador"))
////				{
////						MessageBox.Show("Para convencional :\n\nEl Número de Teléfono debe tener minimo 9 digitos \n\nLos dos primeros digitos corresponden al Codigo Provincial\n\n\nPara Celular :\n\nDebe tener minimo 10 digitos y los dos primeros deben ser '09'", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Error);
////						this.txtCelArrendador.Focus();
////					this.tbBasico.Tabs[0].Selected = true;
////					return;
////				}
//			}
			


			#endregion Valida_DatosCliente
			int iReferencia = 0;

			foreach(Infragistics.Win.UltraWinGrid.UltraGridRow drU in this.ultgconsulta.Rows.All)
			{
				iReferencia =  iReferencia + 1;
				drU.Cells["Indice"].Value = iReferencia;
			}
			

			if (this.cmbEstadoCivil.ActiveRow != null)
			{
				if ((int)this.cmbEstadoCivil.Value != 2 && (int)this.cmbEstadoCivil.Value != 3 && (int)this.cmbEstadoCivil.Value != 5)
				{
					#region Valida_DatosConyuge

					if (!Validacion.vbComboVacio(this.cmbDPCTipodocumento, "Seleccione el Tipo de Documento")) 
					{
							this.tbBasico.Tabs[1].Selected = true;
						this.cmbDPCTipodocumento.Focus();
						return;
					}
					
					if (!Validacion.vbIdentificacion(this.txtDPCdocumento, 5, cdsSeteoF))
					{
						this.tbBasico.Tabs[1].Selected = true; 
						this.txtDPCdocumento.Focus();
						return;
					}
						
					if (!Validacion.vbTexto(this.txtDPCdocumento, 10, 13, "Ingrese la Cedula del Conyuge "))
					{
							this.tbBasico.Tabs[1].Selected = true; 
						this.txtDPCdocumento.Focus();
						return;
					}
					if (!Validacion.vbTexto(this.txtDPCApellidoP, 1, 60, "Ingrese el Apellido Paterno del Conyuge"))
					{
						this.tbBasico.Tabs[1].Selected = true;
							this.txtDPCApellidoP.Focus();
						return;
					}
					if (!Validacion.vbTexto(this.txtDPCPrimerN, 1, 60, "Ingrese el Primer Nombre del Conyuge")) 
					{
							this.tbBasico.Tabs[1].Selected = true;
						this.txtDPCPrimerN.Focus();
						return;
					}
					if (!Validacion.vbComboVacio(this.cmbDPCNacionalidad, "Seleccione la Nacionalidad del Conyuge ")) 
					{
							this.tbBasico.Tabs[1].Selected = true;
						this.cmbDPCNacionalidad.Focus();
						return;
					}
					if (!Validacion.vbFechaEnDocumentos(this.dtDPCFnacimiento, "Ingrese la fecha de Nacimiento del Conyuge ", false, "", cdsSeteoF)) 
					{
							this.tbBasico.Tabs[1].Selected = true; 
						this.dtDPCFnacimiento.Focus();
						return;
					}
					if (!Validacion.vbComboVacio(this.cmbDPCSexo, "Seleccione el Genero del Conyuge "))
					{
							this.tbBasico.Tabs[1].Selected = true;
						this.cmbDPCSexo.Focus();
						return;
					}
//					if (!Validacion.vbComboVacio(this.cmbDPCProvinciaN, "Seleccione la Provincia de Nacimiento del Conyuge "))
//					{
//							this.tbBasico.Tabs[1].Selected = true; 
//								this.cmbDPCProvinciaN.Focus();
//						return;
//					}
//					if (!Validacion.vbComboVacio(this.cmbDPCCantonN, "Seleccione el Canton de Nacimiento del Conyuge ")) 
//					{
//							this.tbBasico.Tabs[1].Selected = true;
//						this.cmbDPCCantonN.Focus();
//						return;
//					}
					if (!Validacion.vbComboVacio(this.cmbDPCNivelEducacion, "Seleccione el Nivel de Eduación del Conyuge "))
					{
						this.tbBasico.Tabs[1].Selected = true;
						this.cmbDPCNivelEducacion.Focus();
						return;
					}
					if (!Validacion.vbComboVacio(this.cmbDPCProfesion, "Seleccione la Profesión del Conyuge ")) 
					{
							this.tbBasico.Tabs[1].Selected = true;
						this.cmbDPCProfesion.Focus();
						return;
					}
					if (!Validacion.vbTexto(this.txtApellidoP, 1, 60, "Ingrese el Apellido Paterno")) 
					{
							this.tbBasico.Tabs[1].Selected = true;
							this.txtApellidoP.Focus();
						return;
					}
					if (!Validacion.vbTexto(this.txtApellidoP, 1, 60, "Ingrese el Primer Nombre"))
					{
						this.tbBasico.Tabs[1].Selected = true;
						this.txtApellidoP.Focus();
						return;
					}
//					if (!Validacion.vbComboVacio(this.cmbEntidad, "Seleccione la Entidad Financiera ")) 
//					{
////							this.tbBasico.Tabs[1].Selected = true;
//						this.cmbEntidad.Focus();
//						return;
//					}
//					if (!Validacion.vbComboVacio(this.cmbBodega, "Seleccione el Local "))
//					{
//							this.tbBasico.Tabs[1].Selected = true; 
//						return;
//					}
//					if (!Validacion.vbComboVacio(this.cmbDPCNacionalidad, "Seleccione la Nacionalidad "))
//					{
//							this.tbBasico.Tabs[1].Selected = true; 
//						return;
//					}

//					if (this.txtDPCTtelefono.Text.ToString().Trim().Length == 0 && this.txtDPCTCelularC.Text.ToString().Trim().Length == 0   )
//					{
//						MessageBox.Show(string.Format("Debe Ingresar el Telefono Convencional del Conyuge"), "Point Technology", MessageBoxButtons.OK,MessageBoxIcon.Stop);
//						this.txtDPCTtelefono.Focus();
//						this.tbBasico.Tabs[1].Selected = true;
//						return;	
//					}
//					if ( this.txtDPCTCelularC.Text.ToString().Trim().Length == 0  )
//					{
//						MessageBox.Show(string.Format("Debe Ingresar el Celular  del Conyuge"), "Point Technology", MessageBoxButtons.OK,MessageBoxIcon.Stop);
//						this.txtDPCTCelularC.Focus();
//						this.tbBasico.Tabs[1].Selected = true;
//						return;	
//					}
//					if (!Validacion.vbValidaTelefono(this.txtDPCTtelefono, 1, false))
//					{
//							this.tbBasico.Tabs[1].Selected = true;
//						this.txtDPCTtelefono.Focus();
//						return;
//					}
//					if (!Validacion.vbValidaTelefono(this.txtDPCTExt, 1, false)) 
//					{
//						this.tbBasico.Tabs[1].Selected = true;
//						this.txtDPCTExt.Focus();
//						return;
//					}
//					if (!Validacion.vbValidaTelefono(this.txtDPCTCelularC, 2, false))
//					{
//						this.tbBasico.Tabs[1].Selected = true;
//						this.txtDPCTCelularC.Focus();
//						return;
//					}
//					if (this.txtDPCTEmail.Text.ToString().Trim().Length > 0 )
//					{
//						if (!Validacion.bEMail(this.txtDPCTEmail.Text.ToString()))
//						{
//								this.tbBasico.Tabs[1].Selected = true;
//							return;
//						}
//					}

					#endregion Valida_DatosConyuge	
					   
				}
			}
			#region Valida_Negocio
			if (((int)this.cmbEntidad.Value == 2 && (int)this.cmbSituacionLaboral.Value > 10) || ((int)this.cmbEntidad.Value != 2 && ((int)this.cmbSituacionLaboral.Value != 1 && (int) this.cmbSituacionLaboral.Value != 4)))
			{
				#region Valida_SeccionNEgocio

//				if (!Validacion.vbTexto(this.txtNegocioRuc, 10, 13, "Ingrese el Ruc del Negocio"))
//				{
//					this.tbBasico.Tabs[3].Selected = true;
//					this.txtNegocioRuc.Focus();
//					return;
//				}
//
//				if (!Validacion.vbIdentificacion(this.txtNegocioRuc, 4, cdsSeteoF))
//				{
//						this.tbBasico.Tabs[3].Selected = true;
//							this.txtNegocioRuc.Focus();
//						return;
//				}
				if (!Validacion.vbTexto(this.txtNombreDelNegocio, 3, 200, "Ingrese el Nombre del Negocio")) 
				{
					this.tbBasico.Tabs[3].Selected = true;
							this.txtNombreDelNegocio.Focus();
					return;
				}

//				if (!Validacion.vbFechaEnDocumentos(this.dtpNegocioFechaInicio, "Ingrese la fecha de inicio del Negocio ", false, "", cdsSeteoF))
//				{
//						this.tbBasico.Tabs[3].Selected = true;
//							this.dtpNegocioFechaInicio.Focus();
//					return;
//				}
				if (!Validacion.vbCampoEnteroVacio(this.txtNegocioMetros,"Ingrese los Metros del Negocio ",1,10000))
				{
						this.tbBasico.Tabs[3].Selected = true;
							this.txtNegocioMetros.Focus();
					return;
				}
//				if (!Validacion.vbCampoEnteroVacio(this.txtNegocioEmpleados,"Ingrese el Número de Empleados del Negocio ",1,500))
//				{
//						this.tbBasico.Tabs[3].Selected = true;
//							this.txtNegocioEmpleados.Focus();
//					return;
//				}
//				if(!Validacion.vbComparaFechas((DateTime)this.dtFechaNacimiento.Value,(DateTime)this.dtpNegocioFechaInicio.Value,this.dtpNegocioFechaInicio,"Inicio del Negocio no puede ser mayor a la edad del cliente",false))
//				{
//					this.tbBasico.Tabs[3].Selected = true;
//							this.dtpNegocioFechaInicio.Focus();
//					return;
//				}
				if (this.txtNegocioemail.Text.ToString().Trim().Length > 0 )
				{
					if (!Validacion.bEMail(this.txtNegocioemail.Text.ToString()))
					{
							this.tbBasico.Tabs[3].Selected = true;
								this.txtNegocioemail.Focus();
						return;
					}
				}
				if (!Validacion.vbTexto(this.txtNegocioCallePrincipal, 1, 100, "Ingrese la Calle Principal del Negocio")) 
				{
					this.tbBasico.Tabs[3].Selected = true;
							this.txtNegocioCallePrincipal.Focus();
					return;
				}
				if (!Validacion.vbTexto(this.txtNegocioNumeroCasa, 1, 60, "Ingrese el Numero de Casa del Negocio"))
				{
					this.tbBasico.Tabs[3].Selected = true;
							this.txtNegocioNumeroCasa.Focus();
					return;
				}
				if (!Validacion.vbTexto(this.txtNegocioCalleSecundaria, 3, 100, "Ingrese la Calle Secundaria del Negocio"))
				{
						this.tbBasico.Tabs[3].Selected = true;
							this.txtNegocioCalleSecundaria.Focus();
					return;
				}

//				if (!Validacion.vbComboVacio(this.cmbNegciotipo, "Seleccione el Tipo de Negocio")) 
//				{
//					this.tbBasico.Tabs[3].Selected = true;
//							this.cmbNegciotipo.Focus();
//					return;
//				}
//				if (!Validacion.vbComboVacio(this.cmbNEgocioActivdadEconomica, "Seleccione la Subactividad Economica"))
//				{
//						this.tbBasico.Tabs[3].Selected = true;
//							this.cmbNEgocioActivdadEconomica.Focus();
//					return;
//				}

				if (!Validacion.vbComboVacio(this.cmbNegocioProvincia  , "Seleccione la Provincia del Negocio")) 
				{
					this.tbBasico.Tabs[3].Selected = true;
							this.cmbNegocioProvincia.Focus();
					return;
				}
				if (!Validacion.vbComboVacio(this.cmbTiempo  , "Seleccione Fecha de Inicio")) 
				{
					this.tbBasico.Tabs[3].Selected = true;
					this.cmbTiempo.Focus();
					return;
				}
				if (!Validacion.vbComboVacio(this.cmbNegocioCanton , "Seleccione el Canton del Negocio")) 
				{
						this.tbBasico.Tabs[3].Selected = true;
							this.cmbNegocioCanton.Focus();
					return;
				}
				if (!Validacion.vbComboVacio(this.cmbNegocioParroquia , "Seleccione la Parroquia del Negocio"))
				{
					this.tbBasico.Tabs[3].Selected = true;
							this.cmbNegocioParroquia.Focus();
					return;
				}
				if (!Validacion.vbComboVacio(this.cmbNegocioBarrio, "Seleccione el Barrio del Negocio"))
				{
						this.tbBasico.Tabs[3].Selected = true; 
							this.cmbNegocioBarrio.Focus();
					return;
				}
				
				if (this.txtNegocioTelefono1.Text.ToString().Trim().Length == 0 &&  this.txtNEgocioCelular.Text.ToString().Trim().Length == 0 )
				{
					MessageBox.Show(string.Format("Debe Ingresar el Telefono Convencional o el Telefono Celular del Negocio "), "Point Technology", MessageBoxButtons.OK,MessageBoxIcon.Stop);
				
					this.tbBasico.Tabs[3].Selected = true;
						this.txtNegocioTelefono1.Focus();
					return;	
				}
//				if (!Validacion.vbValidaTelefono(this.txtNegocioTelefono1, 1, false))
//				{
//					this.tbBasico.Tabs[3].Selected = true; 
//							this.txtNegocioTelefono1.Focus();
//					return;
//				}
				if (!Validacion.vbValidaTelefono(this.txtNegocioTelefono2, 1, false)) 
				{
						this.tbBasico.Tabs[3].Selected = true; 
							this.txtNegocioTelefono2.Focus();
					return;
				}
				if (!Validacion.vbValidaTelefono(this.txtNEgocioCelular, 2, false)) 
				{
						this.tbBasico.Tabs[3].Selected = true; 
							this.txtNEgocioCelular.Focus();
					return;
				}
				if (!Validacion.vbTexto(this.txtNegocioReferncia, 10, 300, "Ingrese la Referencia de la Ubicación del Negocio"))
				{
						this.tbBasico.Tabs[3].Selected = true; 
							this.txtNegocioReferncia.Focus();
					return;
				}
				if (!Validacion.vbCampoEnteroVacio(this.txtNEgocioIngresos,"Ingrese los Ingresos del Negocio ",1,1000000))
				{
					this.tbBasico.Tabs[3].Selected = true; 
							this.txtNEgocioIngresos.Focus();
					return;
				}
				if (!Validacion.vbCampoEnteroVacio(this.txtNEgocioGastos,"Ingrese los Gastos del Negocio  ",1,1000000))
				{
						this.tbBasico.Tabs[3].Selected = true; 
							this.txtNEgocioGastos.Focus();
					return;
				}
//				if (!Validacion.vbCampoEnteroVacio(this.txtNegocioActivos,"Ingrese los Activos del Negocio ",1,1000000))
//				{
//						this.tbBasico.Tabs[3].Selected = true; 
//							this.txtNegocioActivos.Focus();
//					return;
//				}
//				if (!Validacion.vbCampoEnteroVacio(this.txtNEgocioPasivos,"Ingrese los  PAsivos del Negocio ",1,1000000))
//				{
//					this.tbBasico.Tabs[3].Selected = true; 
//							this.txtNEgocioPasivos.Focus();
//					return;
//				}

				decimal dSueldoBasico = Funcion.decEscalarSQL(cdsSeteoF, "SELECT SueldoBasico FROM Seteo");
				
				if ((int)this.txtNEgocioIngresos.Value < dSueldoBasico)
				{
					MessageBox.Show(string.Format("Los ingresos no pueden ser  menor a {0}", dSueldoBasico), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
					this.tbBasico.Tabs[3].Selected = true;
					this.txtNEgocioIngresos.Focus();
					return;
				}
				if ((int)this.txtNEgocioGastos.Value > (int)this.txtNEgocioIngresos.Value )
				{
					MessageBox.Show(string.Format("Los Gastos no pueden ser  mayor a los Ingresos "), "Point Technology", MessageBoxButtons.OK,MessageBoxIcon.Stop);
					
					this.tbBasico.Tabs[3].Selected = true;
					this.txtNEgocioGastos.Focus();
					return;
				}
			
				#endregion Valida_SeccionNEgocio

			}
			#endregion Valida_Negocio

			#region Valida Referencia

			if (this.ultgconsulta.Rows.Count < 2)
			{
				MessageBox.Show("Debe Ingresar al menos 2 referencia", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				this.tbBasico.Tabs[2].Selected = true;
				this.ultgconsulta.Focus();
				return;
			}

			foreach(Infragistics.Win.UltraWinGrid.UltraGridRow drReferencia in this.ultgconsulta.Rows.All)
			{
				if (drReferencia.Cells["idParentesco"].Value == System.DBNull.Value)
				{
					MessageBox.Show(string.Format("Seleccione el Parentesco de la Referencia"), "Point Technology", MessageBoxButtons.OK,MessageBoxIcon.Stop);
					this.ultgconsulta.ActiveRow = drReferencia;
					this.tbBasico.Tabs[2].Selected = true;
					return;	
				}
				if (drReferencia.Cells["idProvincia"].Value == System.DBNull.Value)
				{
					MessageBox.Show(string.Format("Seleccione la Provincia de la Referencia "), "Point Technology", MessageBoxButtons.OK,MessageBoxIcon.Stop);
									
					this.tbBasico.Tabs[2].Selected = true;
					return;	
				}
				if (drReferencia.Cells["idCanton"].Value == System.DBNull.Value)
				{
					MessageBox.Show(string.Format("Seleccione el Canton de la Referencia"), "Point Technology", MessageBoxButtons.OK,MessageBoxIcon.Stop);
									
					this.tbBasico.Tabs[2].Selected = true;
					return;	
				}
				if(drReferencia.Cells["ApellidoPaterno"].Value.ToString().Length == 0)
				{
					MessageBox.Show("Ingrese Apellido Paterno de la Referencia ", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
					this.tbBasico.Tabs[2].Selected = true;
					this.ultgconsulta.ActiveRow = drReferencia;
					this.ultgconsulta.ActiveRow.Selected = true;
					return;
				}
				if(drReferencia.Cells["PrimerNombre"].Value.ToString().Length == 0)
				{
					MessageBox.Show("Ingrese el Primer Nombre de la Referencia ", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
					this.tbBasico.Tabs[2].Selected = true;
					this.ultgconsulta.ActiveRow = drReferencia;
					this.ultgconsulta.ActiveRow.Selected = true;
					return;
				}
				//				if(drReferencia.Cells["Direccion"].Value.ToString().Length == 0)
				//				{
				//					MessageBox.Show("Ingrese la Direccion de la Referencia ", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				//					this.tbBasico.Tabs[2].Selected = true;
				//					this.ultgconsulta.ActiveRow = drReferencia;
				//					this.ultgconsulta.ActiveRow.Selected = true;
				//					return;
				//				}


				#region Valida Telefonos

				//				if (drReferencia.Cells["Telefono"].Value.ToString().Trim().Length == 0 &&  drReferencia.Cells["Celular"].Value.ToString().Trim().Length == 0)
				//				{
				//					MessageBox.Show(string.Format("Debe Ingresar el Telefono Convencional o el Telefono Celular de la Referencia"), "Point Technology", MessageBoxButtons.OK,MessageBoxIcon.Stop);
				//					this.ultgconsulta.ActiveRow = drReferencia;
				//					this.tbBasico.Tabs[2].Selected = true;
				//					return;	
				//
				//				}
				
				//				if (drReferencia.Cells["Telefono"].Value.ToString().Trim().Length > 0)
				//				{
				//					if (drReferencia.Cells["Telefono"].Value.ToString().Trim().Length != 9)
				//					{
				//						MessageBox.Show("El Telefono Convencional debe tener 9 Digitos.\n\nLos dos primeros corresponden al Codigo de Provincia", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				//						this.tbBasico.Tabs[2].Selected = true;
				//						this.ultgconsulta.ActiveRow = drReferencia;
				//						this.ultgconsulta.ActiveRow.Selected = true;
				//						return;
				//					}
				//				}
				

				#region Celular
				if (drReferencia.Cells["Celular"].Value.ToString().Trim().Length > 0)
				{
					
					if (drReferencia.Cells["Celular"].Value.ToString().Trim().Length != 10)
					{
						MessageBox.Show("El Telefono Celular debe tener 10 digitos", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
						this.tbBasico.Tabs[2].Selected = true;
						this.ultgconsulta.ActiveRow = drReferencia;
						this.ultgconsulta.ActiveRow.Selected = true;
						return;
					}
				}
				#endregion Celular

				//				if(drReferencia.Cells["EmpresaTrabajo"].Value.ToString().Length > 0)
				//				{
				//
				//					if(drReferencia.Cells["TelefonoTrabajo"].Value.ToString().Length==0)
				//					{
				//						MessageBox.Show("Ingrese el Telefono de la Referencia del Trabajo", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				//						this.tbBasico.Tabs[2].Selected = true;
				//						this.ultgconsulta.ActiveRow = drReferencia;
				//						this.ultgconsulta.ActiveRow.Selected = true;
				//						return;
				//					}

				//					if (drReferencia.Cells["TelefonoTrabajo"].Value.ToString().Trim().Length > 0)
				//					{
				//						if (drReferencia.Cells["TelefonoTrabajo"].Value.ToString().Trim().Length != 9)
				//						{
				//							MessageBox.Show("El Telefono Convencional debe tener 9 Digitos.\n\nLos dos primeros corresponden al Codigo de Provincia", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				//							this.tbBasico.Tabs[2].Selected = true;
				//							this.ultgconsulta.ActiveRow = drReferencia;
				//							this.ultgconsulta.ActiveRow.Selected = true;
				//							return;
				//						}
				//					}
				//				}
				#endregion Valida Telefonos
			}
			#endregion Valida Referencia

			if (((int)this.cmbEntidad.Value == 2 && Convert.ToInt32(this.cmbSituacionLaboral.Value) == 10) || 
				((int)this.cmbEntidad.Value != 2 && (Convert.ToInt32(this.cmbSituacionLaboral.Value) == 1 || Convert.ToInt32(this.cmbSituacionLaboral.Value) == 3)))					
			{				
				#region Valida_SeccionDependiente

				if (!Validacion.vbTexto(this.txtDependienteEmpresa, 1, 100, "Ingrese el Nombre de la Empresa"))
				{
					this.tbBasico.Tabs[4].Selected = true; 
					this.txtDependienteEmpresa.Focus();
					return;
				}
				if (!Validacion.vbComboVacio(this.cmbDependienteTipoempresa, "Seleccione el tipo de empresa ")) 
				{
					this.tbBasico.Tabs[4].Selected = true; 
						this.cmbDependienteTipoempresa.Focus();
					return;
				}
				if (!Validacion.vbFechaEnDocumentos(this.dtDependienteFechaIngreso, "Ingrese la fecha de Ingreso a la Empresa ", false, "", cdsSeteoF)) 
				{
					this.tbBasico.Tabs[4].Selected = true;
									this.dtDependienteFechaIngreso.Focus();
					return;
				}
				if (!Validacion.vbCampoEnteroVacio(this.txtDependienteIngresos,"Seleccione los Ingresos del Cliente  ",1,1000000))
				{
					this.tbBasico.Tabs[4].Selected = true;
						this.txtDependienteIngresos.Focus();
					return;
				}
				if (!Validacion.vbCampoEnteroVacio(this.txtDependienteGastos,"Ingrese los Gastos del Cliente ",1,1000000))
				{
					this.tbBasico.Tabs[4].Selected = true; 
						this.txtDependienteGastos.Focus();
					return;
				}
				if(!Validacion.vbComparaFechas((DateTime)this.dtFechaNacimiento.Value,(DateTime)this.dtDependienteFechaIngreso.Value,this.dtDependienteFechaIngreso,"Fecha de Ingreso al Trabajo no puede ser mayor a la edad del cliente",false))
				{
					this.tbBasico.Tabs[3].Selected = true;
						this.dtDependienteFechaIngreso.Focus();

					return;
				}
//				if (!Validacion.vbCampoEnteroVacio(this.txtDependienteActivos,"Ingrese los Activos del Cliente ",1,1000000))
//				{
//					this.tbBasico.Tabs[4].Selected = true; 
//					this.txtDependienteActivos.Focus();
//					return;
//				}
//				if (!Validacion.vbCampoEnteroVacio(this.txtDependientePasivos,"Ingrese los  PAsivos del Cliente ",1,1000000))
//				{
//					this.tbBasico.Tabs[4].Selected = true;
//					this.txtDependientePasivos.Focus();
//					return;
//				}
				if (!Validacion.vbComboVacio(this.cmbDependienteTipocontrato, "Seleccione el Tipo de Contrato ")) 
				{
					this.tbBasico.Tabs[4].Selected = true; 
					this.cmbDependienteTipocontrato.Focus();
					return;
				}
//				if (!Validacion.vbComboVacio(this.cmbTiempoDependiente, "Seleccione Fecha de Ingreso ")) 
//				{
//					this.tbBasico.Tabs[4].Selected = true; 
//					this.cmbTiempoDependiente.Focus();
//					return;
//				}
				if (!Validacion.vbComboVacio(this.cmbDependienteTipoSueldo, "Seleccione el Tipo de Sueldo "))
				{
					this.tbBasico.Tabs[4].Selected = true; 
					this.cmbDependienteTipoSueldo.Focus();
					return;
				}
				if (!Validacion.vbComboVacio(this.cmbDependienteTipoSueldo, "Seleccione el Cargo "))
				{
					this.tbBasico.Tabs[4].Selected = true; 
					this.cmbDependienteTipoSueldo.Focus();
					return;
				}
				//				if (!Validacion.vbCampoEnteroVacio(this.txtDependienteDiasPago,"Ingrese el Día de Pago ",1,31))return;
				if (!Validacion.vbComboVacio(this.cmbDependienteProvincia, "Seleccione la Provincia de su Trabajo ")) 
				{
					this.tbBasico.Tabs[4].Selected = true; 
					this.cmbDependienteProvincia.Focus();
					return;
				}
				if (!Validacion.vbComboVacio(this.cmbDependienteCanton, "Seleccione el Canton de su Trabajo ")) 
				{
					this.tbBasico.Tabs[4].Selected = true; 
					this.cmbDependienteCanton.Focus();
					return;
				}
				if (!Validacion.vbComboVacio(this.cmbDependienteParroquia, "Seleccione la Parroquia de su Trabajo "))
				{
					this.tbBasico.Tabs[4].Selected = true; 
					this.cmbDependienteParroquia.Focus();
					return;
				}
				if (!Validacion.vbComboVacio(this.cmbDependienteBarrio, "Seleccione el Barrio de su Trabajo ")) 
				{
					this.tbBasico.Tabs[4].Selected = true;
					this.cmbDependienteBarrio.Focus();
					return;
				}
				if (!Validacion.vbTexto(this.txtDependienteCallePrinciapl, 3, 100, "Ingrese la Calle Principal del Trabajo"))
				{
					this.tbBasico.Tabs[4].Selected = true;
					this.txtDependienteCallePrinciapl.Focus();
					return;
				}
				if (!Validacion.vbTexto(this.txtDependienteNumeroCasa, 3, 60, "Ingrese el Numero de Casa del Trabajo")) 
				{
					this.tbBasico.Tabs[4].Selected = true;
					this.txtDependienteNumeroCasa.Focus();
					return;
				}
				if (!Validacion.vbTexto(this.txtDependienteCalleSecundaria, 3, 100, "Ingrese la Calle Secundaria del Trabajo"))
				{
					this.tbBasico.Tabs[4].Selected = true; 
					this.txtDependienteCalleSecundaria.Focus();
					return;
				}
				if (this.txtDependienteemail.Text.ToString().Trim().Length > 0 )
				{
					if (!Validacion.bEMail(this.txtDependienteemail.Text.ToString()))
					{
						this.tbBasico.Tabs[4].Selected = true;
						this.txtDependienteemail.Focus();
						return;
					}
				}
				if (this.txtDependienteTelefono.Text.ToString().Trim().Length == 0 &&  this.txtDependienteCelular.Text.ToString().Trim().Length == 0 )
				{
					MessageBox.Show(string.Format("Debe Ingresar el Telefono Convencional o el Telefono Celular del Trabajo "), "Point Technology", MessageBoxButtons.OK,MessageBoxIcon.Stop);
		
					this.tbBasico.Tabs[4].Selected = true;
								this.txtDependienteTelefono.Focus();
					return;	
				}
				if (!Validacion.vbValidaTelefono(this.txtDependienteTelefono, 1, false))
				{
					this.tbBasico.Tabs[4].Selected = true; 
								this.txtDependienteTelefono.Focus();
					return;
				}
				if (!Validacion.vbValidaTelefono(this.txtDependienteEXT, 1, false))
				{
					this.tbBasico.Tabs[4].Selected = true; 
								this.txtDependienteEXT.Focus();
					return;
				}
				if (!Validacion.vbValidaTelefono(this.txtDependienteCelular, 2, false))
				{
					this.tbBasico.Tabs[4].Selected = true;
								this.txtDependienteCelular.Focus();
					return;
				}
				if (!Validacion.vbTexto(this.txtDependienteReferencia, 10, 300, "Ingrese la Referencia de la Ubicación del Trabajo")) 
				{
					this.tbBasico.Tabs[4].Selected = true; 
								this.txtDependienteReferencia.Focus();
					return;
				}

				decimal dSueldoBasico = Funcion.decEscalarSQL(cdsSeteoF, "SELECT SueldoBasico FROM Seteo");
				
				if ((int)this.txtDependienteIngresos.Value < dSueldoBasico)
				{
					MessageBox.Show(string.Format("Los ingresos no pueden ser  menor a {0}", dSueldoBasico), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
					this.tbBasico.Tabs[4].Selected = true;
					this.txtDependienteIngresos.Focus();
					return;
				}
				if ((int)this.txtDependienteGastos.Value > (int)this.txtDependienteIngresos.Value )
				{
					MessageBox.Show(string.Format("Los Gastos no pueden ser  mayor a los Ingresos "), "Point Technology", MessageBoxButtons.OK,MessageBoxIcon.Stop);
					
					this.tbBasico.Tabs[3].Selected = true;
					this.txtDependienteGastos.Focus();
					return;
				}

				#endregion Valida_SeccionDependiente
			}

			#region Valida_Combo_Vacio_Datos

			int iDDInmueble = 0;
			int iDDCiudadInmueble = 0;
			if ( this.cmbDDInmueble.ActiveRow != null) iDDInmueble  = ( int) this.cmbDDInmueble.Value;
			if ( this.cmbDDCiudadInmueble.ActiveRow != null) iDDCiudadInmueble  = ( int) this.cmbDDCiudadInmueble.Value;

			#endregion Valida_Combo_Vacio_Datos
		
			#region Valida_Combo_Vacio_Conyuge
      
				int iDPCTipodocumento = 0;
				int iDPCNacionalidad = 0;
				int iDPCSexo = 0;
				int iDPCProvinciaN = 0;
				int iDPCCantonN = 0;
				int iDPCNivelEducacion = 0;
				int iDPCProfesion= 0;

			int iDPCSituacionL = 0;
			int iDPCActividadE = 0;
			int iDPCTActividadEmpresa = 0;
			int iDPCTCargo = 0;
			int iDPCTProvincia = 0;
			int iDPCTCanton = 0;
			int iDPCTParroquia = 0;
			int iDPCTBarrio = 0;

			if ( this.cmbDPCTipodocumento.ActiveRow != null) iDPCTipodocumento = (int) this.cmbDPCTipodocumento.Value;
			if ( this.cmbDPCNacionalidad.ActiveRow != null)  iDPCNacionalidad = (int) this.cmbDPCNacionalidad.Value;
			if ( this.cmbDPCSexo.ActiveRow != null)  iDPCSexo = (int) this.cmbDPCSexo.Value;
			if ( this.cmbDPCProvinciaN.ActiveRow != null) iDPCProvinciaN = (int) this.cmbDPCProvinciaN.Value;
			if ( this.cmbDPCCantonN.ActiveRow != null)  iDPCCantonN = (int) this.cmbDPCCantonN.Value;
			if ( this.cmbDPCNivelEducacion.ActiveRow != null) iDPCNivelEducacion = (int) this.cmbDPCNivelEducacion.Value;
			if ( this.cmbDPCProfesion.ActiveRow != null)   iDPCProfesion = (int) this.cmbDPCProfesion.Value;
			if ( this.cmbDPCTActividadEmpresa.ActiveRow != null) iDPCTActividadEmpresa  = ( int) this.cmbDPCTActividadEmpresa.Value;
			if ( this.cmbDPCTCargo.ActiveRow != null) iDPCTCargo  = ( int) this.cmbDPCTCargo.Value;
			if ( this.cmbDPCSituacionL.ActiveRow != null) iDPCSituacionL  = ( int) this.cmbDPCSituacionL.Value;
			if ( this.cmbDPCActividadE.ActiveRow != null) iDPCActividadE  = ( int) this.cmbDPCActividadE.Value;
			if ( this.cmbDPCTProvincia.ActiveRow != null) iDPCTProvincia  = ( int) this.cmbDPCTProvincia.Value;
			if ( this.cmbDPCTCanton.ActiveRow != null) iDPCTCanton  = ( int) this.cmbDPCTCanton.Value;
			if ( this.cmbDPCTParroquia.ActiveRow != null) iDPCTParroquia  = ( int) this.cmbDPCTParroquia.Value;
			if ( this.cmbDPCTBarrio.ActiveRow != null) iDPCTBarrio  = ( int) this.cmbDPCTBarrio.Value;

			#endregion Valida_Combo_Vacio_Conyuge

			#region Valida_Combo_Vacio_Negocio

				int iNegciotipo = 0;
				int iNEgocioActivdadEconomica = 0;
				int iNegocioProvincia = 0;
				int iNegocioCanton = 0;
				int iNegocioParroquia = 0;
				int iNegocioBarrio = 0;

				if ( this.cmbNegciotipo.ActiveRow != null) iNegciotipo = (int) this.cmbNegciotipo.Value;
				if ( this.cmbNEgocioActivdadEconomica.ActiveRow != null) iNEgocioActivdadEconomica = (int) this.cmbNEgocioActivdadEconomica.Value;
				if ( this.cmbNegocioProvincia.ActiveRow != null) iNegocioProvincia = (int) this.cmbNegocioProvincia.Value;
				if ( this.cmbNegocioCanton.ActiveRow != null) iNegocioCanton = (int) this.cmbNegocioCanton.Value;
				if ( this.cmbNegocioParroquia.ActiveRow != null) iNegocioParroquia = (int) this.cmbNegocioParroquia.Value;
				if ( this.cmbNegocioBarrio.ActiveRow != null) iNegocioBarrio = (int) this.cmbNegocioBarrio.Value;

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

			if ( this.cmbDependienteTipoempresa.ActiveRow != null) iDependienteTipoempresa = (int) this.cmbDependienteTipoempresa.Value;
			if ( this.cmbDependienteTipocontrato.ActiveRow != null) iDependienteTipocontrato = (int) this.cmbDependienteTipocontrato.Value;
			if ( this.cmbDependienteTipoSueldo.ActiveRow != null) iDependienteTipoSueldo = (int) this.cmbDependienteTipoSueldo.Value;
			if ( this.cmbDependienteCargo.ActiveRow != null) iDependienteCargo = (int) this.cmbDependienteCargo.Value;
			if ( this.cmbDependienteProvincia.ActiveRow != null) iDependienteProvincia = (int) this.cmbDependienteProvincia.Value;
			if ( this.cmbDependienteCanton.ActiveRow != null) iDependienteCanton = (int) this.cmbDependienteCanton.Value;
			if ( this.cmbDependienteParroquia.ActiveRow != null) iDependienteParroquia = (int) this.cmbDependienteParroquia.Value;
			if ( this.cmbDependienteBarrio.ActiveRow != null) iDependienteBarrio = (int) this.cmbDependienteBarrio.Value;

		
			#endregion Valida_Combo_Vacio_Dependiente

			#region ValidaCoordenadas 
			int Existencia = Funcion.iEscalarSQL(cdsSeteoF,string.Format(" select count(1) from CoordenadasPrefactura where cedula = '{0}'",(string)this.txtDocumento.Text));
			if( Existencia >= 0)
			{
				if (DialogResult.No == MessageBox.Show("¿Desea continuar sin registrar la Ubicación.?, Si presiona SI puede continuar.", "Point Technology", MessageBoxButtons.YesNo, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button2))
				{
					return;
				}
			}
			#endregion ValidaCoordenadas

//			if (miAcceso.BHabilitaControles)
//			{
				if (!Validacion.vbTexto(this.txtObservaciones,2,1000,"Ingrese una Observacion en la pestaña de Factores de Crédito")) 
				{
					this.tbBasico.Tabs[6].Selected = true; 
					this.txtObservaciones.Focus();
					return;
				}
//			}
//			else {
//			this.txtObservaciones.Value="REGISTRO DE SOLICITIUD ";
//			}
		

		int iCreditoEstado = 0;
		int iCreditoCalificacion = 0;

		int cmbidTipoCliente = 0;
		int cmbdetallecliente = 0;
		int cmbSituacionLaboral2 = 0;
		int cmdEstadoSolicitud = 0;
		int cmbTiempo =0;
		int cmbTiempoDependiente =0;
		int cmbDDTipoVivienda =0;

		if ( this.cmbidTipoCliente.ActiveRow != null) cmbidTipoCliente  = ( int) this.cmbidTipoCliente.Value;
		if ( this.cmbdetallecliente.ActiveRow != null) cmbdetallecliente  = ( int) this.cmbdetallecliente.Value;
		if ( this.cmbSituacionLaboral2.ActiveRow != null) cmbSituacionLaboral2  = ( int) this.cmbSituacionLaboral2.Value;
		if ( this.cmdEstadoSolicitud.ActiveRow != null) cmdEstadoSolicitud  = ( int) this.cmdEstadoSolicitud.Value;
		if ( this.cmbDDTipoVivienda.ActiveRow != null) cmbDDTipoVivienda  = ( int) this.cmbDDTipoVivienda.Value;

		if ( this.cmbTiempo.ActiveRow != null) cmbTiempo  = ( int) this.cmbTiempo.Value;
		if ( this.cmbTiempoDependiente.ActiveRow != null) cmbTiempoDependiente  = ( int) this.cmbTiempoDependiente.Value;

		if ( this.cmbCreditoEstado.ActiveRow != null) iCreditoEstado  = ( int) this.cmbCreditoEstado.Value;
		if ( this.cmbCreditoCalificacion.ActiveRow != null) iCreditoCalificacion  = ( int) this.cmbCreditoCalificacion.Value;
    if ((int)this.txtEstado.Value < 2) this.txtEstado.Value = 1;
		if ((bool) this.chkGarante.Checked)
			{
			int iContG =0;
			iContG = Funcion.iEscalarSQL(cdsSeteoF, string.Format("select COUNT (*) from Cre_Garantes c  Inner Join Cre_DatosGenerales d on c.idCre_Garante = d.idOrigen where  Origen = 2 and idCre_Solicitud = {0}  ", (int)this.txtIdCre_Solicitud.Value));

				if( iContG <= 0)
				{
					MessageBox.Show("Ingrese el Garante ", "Validacion", MessageBoxButtons.OK, MessageBoxIcon.Stop);
					this.btnGuardar.Focus();
					return;
				}
			}

			#region Guarda Cuota
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

					#region Guarda Datos Credito

					string sSQLDetalle = string.Format("Exec Cre_GuardaSolicitudes {0}, '{1}', {2}, {3}, {4}, {5}, {6}, {7}, {8}, {9}, '{10}', '{11}', '{12}', '{13}', {14}, '{15}', '{16}', '{17}',{18},{19},{20},{21},{22}",
								(int)this.txtIdCre_Solicitud.Value,
								Convert.ToDateTime(this.dtFecha.Value).ToString("yyyyMMdd"),
								(bool)this.chkGarante.Checked,
								(int)this.cmbEntidad.Value,
								0,
								(int)this.cmbBodega.Value,
								(int)this.txtCreditoCuotaAsiganada.Value,
								0,0,
								cmdEstadoSolicitud,"","","","",																							
								this.txtEstado.Value,
								Convert.ToDateTime(this.dtCreditoFechaSolicitud.Value).ToString("yyyyMMdd"),											
								this.lblUsuarioCompleta.Text,
								this.txtObservaciones.Text,
							  cmbidTipoCliente,
							  cmbdetallecliente,
							  cmbSituacionLaboral2,
							  Convert.ToDecimal(this.txtCupo.Value),
								cmbDDTipoVivienda); 

					oCmdActualiza.CommandText = sSQLDetalle;
					this.txtIdCre_Solicitud.Value= (int)oCmdActualiza.ExecuteScalar();

						#endregion Guarda Datos Credito

					#region Datos Personales Cliente

					int Ingresos = Convert.ToInt32(this.txtNEgocioIngresos.Value) + Convert.ToInt32(this.txtDependienteIngresos.Value);
					int Gastos = Convert.ToInt32(this.txtNEgocioGastos.Value) + Convert.ToInt32(this.txtDependienteGastos.Value);
					int Activos = Convert.ToInt32(this.txtNegocioActivos.Value) + Convert.ToInt32(this.txtDependienteActivos.Value);
					int Pasivos = Convert.ToInt32(this.txtNEgocioPasivos.Value) + Convert.ToInt32(this.txtDependientePasivos.Value);

					string sSQLDatosGenerales = string.Format("Exec Cre_GuardaDatosGenerales {0}, {1}, {2}, {3}, '{4}', '{5}', '{6}', '{7}', '{8}', {9}, {10}, {11}, {12}, {13}, {14}, '{15}', {16}, {17},{18}, {19}, {20}, {21}, {22}, {23}, {24}, {25}, {26}, {27}, {28}, {29}, {30}, {31}, {32}, '{33}', '{34}', {35}, {36}, {37},{38}, {39},'{40}', {41}, {42},{43},{44},'{45}'",
													(int)this.txtClienteIdCre_DatosGenerales.Value, 
													0,
													(int)this.txtIdCre_Solicitud.Value, 
													this.cmbTipoRuc.Value,
													this.txtDocumento.Text, 
													this.txtApellidoP.Text, 
													this.txtApellidoM.Text,
													this.txtPnombre.Text.ToString(), 
													this.txtSnombre.Text.ToString(),
													iCreditoCalificacion, 
													0, 
													(int)this.txtCreditoScore.Value, 
													0,
													iCreditoEstado, 
													(int)this.cmbNacionalidad.Value, 
													Convert.ToDateTime(this.dtFechaNacimiento.Value).ToString("yyyyMMdd"),
													(int)this.cmbSexo.Value,   
													(int) this.cmbProvinciaNAcimiento.Value,
													(int) this.cmbCantonNacimiento.Value,
													(int) this.cmbEstadoCivil.Value, 
													(int)this.txtDependientes.Value, 
													(int) this.cmbNivelEducacion.Value, 
													(int) this.cmbProfesion.Value,
                             0, //	(int) this.cmbClienteFuenteDeIngresos.Value, 
													(int) this.cmbSituacionLaboral.Value, 
													(int) this.cmbActividadEconomica.Value,
													Ingresos, 
													Gastos, 
													Activos, 
													Pasivos,
													iDDInmueble, 
													iDDCiudadInmueble, 
													(int)this.txtDDValorInmueble.Value,
													this.txtVehiculoMarca.Text,
													this.txtModelo.Text, 
													(int)this.txtValorVehiculo.Value,
													(bool)this.chkSeparacionBienes.Checked,
													(bool)this.chkHipoteca.Checked,
													(int)this.txtCreditoScoreInclusion.Value,
													(int)this.txtCreditoScoreSobreEndeudamiento.Value,
													this.txtRangoIngresos.Text.ToString(),
													Convert.ToDecimal(this.txtCapacidadPago.Value),
													Convert.ToDecimal(this.txtTotalVencido.Value),
													Convert.ToInt32((bool) this.chkNuovo.Checked),
													Convert.ToInt32((bool) this.ChkOmite.Checked),
													this.txtObsActividadEconomica.Text.ToString());

					oCmdActualiza.CommandText = sSQLDatosGenerales;
					this.txtClienteIdCre_DatosGenerales.Value= (int)oCmdActualiza.ExecuteScalar();
	        
					 #endregion Datos Personales Cliente

					#region Guarda Datos Ubicacion cliente

					string sSQLDatosUbicacion = string.Format("Exec Cre_GuardaDatosUbicacion {0}, {1}, {2}, {3}, {4}, {5}, {6}, '{7}', '{8}', '{9}', '{10}', '{11}', '{12}', '{13}', '{14}', '{15}', '{16}', {17},'{18}', '{19}', '{20}', {21}, {22}, {23}, {24},'{25}',{26}",
								(int)this.txtClienteIdCre_DatosUbicacion.Value, 
								0, 
								(int)this.txtIdCre_Solicitud.Value,
						    (int)	this.cmbDDProvincia.Value, 
							  (int)this.cmbDDCanton.Value, 
								(int)this.cmbDDParroquia.Value, 
								(int)this.cmbDDBarrio.Value,
								this.txtDDCallaPrincipal.Text, 
								this.txtDDNumeroCasa.Text, 
								this.txtDDCalleSecundaria.Text, 
								this.txtDDReferenciasU.Text,
								this.txtDDTelefono1.Text,
								"", 
								this.txtDDTelefono2.Text.ToString(), 
								"", this.txtDDCelular.Text.ToString(),
								this.txtDDEmail.Text,
								(int)this.cmbDDTipoVivienda.Value,  
						    Convert.ToDateTime(this.dtTiempoVivienda.Value).ToString("yyyyMMdd"),
								this.txtNombreArrendador.Text, 
								this.txtTelfArrendador.Text,
								0,
								0, 
								0, 
								0,
								this.txtCelArrendador.Text,
								(int)this.cmbTiempoVivienda.Value);     
					oCmdActualiza.CommandText = sSQLDatosUbicacion;
					this.txtClienteIdCre_DatosUbicacion.Value= (int)oCmdActualiza.ExecuteScalar();

					#endregion Guarda Datos Ubicacion cliente

					#region Conyugue

					if ((int)this.cmbEstadoCivil.Value != 2 && (int)this.cmbEstadoCivil.Value != 3 && (int)this.cmbEstadoCivil.Value != 5)
					{

						string sSQLGuardaConyuge = string.Format("Exec Cre_GuardaConyugues {0}, {1}, {2}, '{3}', {4}, {5}",
											(int)this.txtIdCre_Conyugue.Value, 
											0, 
											(int)this.txtIdCre_Solicitud.Value,
											this.txtDPCTNombreEmpresa.Text, 
											iDPCTActividadEmpresa,
											iDPCTCargo);   
  
						oCmdActualiza.CommandText = sSQLGuardaConyuge;
						this.txtIdCre_Conyugue.Value = (int)oCmdActualiza.ExecuteScalar();

				

					#region Datos Generales Conyugue

						string sSQLDatosGeneralesconyuge = string.Format("Exec Cre_GuardaDatosGenerales {0}, {1}, {2}, {3}, '{4}', '{5}', '{6}', '{7}', '{8}', {9}, {10}, {11}, {12}, {13}, {14}, '{15}', {16}, {17},{18}, {19}, {20}, {21}, {22}, {23}, {24}, {25}, {26}, {27}, {28}, {29}, {30}, {31}, {32}, '{33}', '{34}', {35}, {36}, {37}",
							(int)this.txtConyugueIdCre_DatosGenerales.Value, 
						   1, 
							(int)this.txtIdCre_Conyugue.Value,
							iDPCTipodocumento, 
							this.txtDPCdocumento.Text,
							this.txtDPCApellidoP.Text, 
							this.txtDPCApellidoM.Text,
							this.txtDPCPrimerN.Text,
							this.txtDPCsegundoN.Text, 
							0, 0, 0, 0,	0,
							iDPCNacionalidad, 
						  Convert.ToDateTime(this.dtDPCFnacimiento.Value).ToString("yyyyMMdd"),
							iDPCSexo,
						  iDPCProvinciaN, 
							iDPCCantonN,
							0, 0, 
						  iDPCNivelEducacion, 
							iDPCProfesion,
							0, 
							iDPCSituacionL, 
							iDPCActividadE,
							0,0,0,0, 0,0, 0, "","",0,false,false);

						oCmdActualiza.CommandText = sSQLDatosGeneralesconyuge;
						this.txtConyugueIdCre_DatosGenerales.Value= (int)oCmdActualiza.ExecuteScalar();

					

						#endregion Datos Generales Conyugue


					#region Datos Ubicacion Trabajo del Conyugue

						string sSQLDatosTrabajoconyuge = string.Format("Exec Cre_GuardaDatosUbicacion {0}, {1}, {2}, {3}, {4}, {5}, {6}, '{7}', '{8}', '{9}', '{10}', '{11}', '{12}', '{13}', '{14}', '{15}', '{16}', {17},'{18}', '{19}', '{20}', {21}, {22}, {23}, {24}",
										(int)this.txtConyugueIdCre_DatosUbicacion.Value, 
										1, 
										(int)this.txtIdCre_Conyugue.Value,
										iDPCTProvincia, 
										iDPCTCanton ,
										iDPCTParroquia ,
										iDPCTBarrio,
										this.txtDPCTCallePrincipal.Text, 
										this.txtDPCTNumeroCasa.Text, 
										this.txtDPCTCalleSecundaria.Text, 
										this.txtDPCTReferencia.Text,
										this.txtDPCTtelefono.Text, 
										this.txtDPCTExt.Text, 
										"", "", 
										this.txtDPCTCelularC.Text, 
										this.txtDPCTEmail.Text,
										0, 
										 Convert.ToDateTime(DateTime.Today).ToString("yyyyMMdd"), 
										"", "", 
										(int)this.txtDPCTIngresos.Value,
										(int)this.txtDPCTGastos.Value, 
										(int)this.txtDPCTActivos.Value, 
										(int)this.txtDPCTPasivos.Value
										);     
						oCmdActualiza.CommandText = sSQLDatosTrabajoconyuge;
						this.txtConyugueIdCre_DatosUbicacion.Value= (int)oCmdActualiza.ExecuteScalar();


						
				
						#endregion Datos Ubicacion Trabajo del Conyugue
					}
					#endregion Conyugue

					#region Referencia

					

					foreach(Infragistics.Win.UltraWinGrid.UltraGridRow drReferencias in this.ultgconsulta.Rows.All)
					{
//revisar
						string sSQLReferencia = string.Format("Exec Cre_GuardaReferenciasClientes   {0}, {1}, {2}, {3}, '{4}', '{5}','{6}','{7}',{8},{9},'{10}', '{11}','{12}','{13}','{14}','{15}', '{16}'",
							(int)drReferencias.Cells["idCre_ReferenciasClientes"].Value,
							(int) this.txtClienteIdCre_DatosGenerales.Value,
//							(int)drReferencias.Cells["idCre_DatosGenerales"].Value,// MANDAR TXT
							(int)drReferencias.Cells["Indice"].Value,
							(int)drReferencias.Cells["idParentesco"].Value,
							drReferencias.Cells["ApellidoPaterno"].Value.ToString(),
							drReferencias.Cells["ApellidoMaterno"].Value.ToString(),
							drReferencias.Cells["PrimerNombre"].Value.ToString(),
							drReferencias.Cells["SegundoNombre"].Value.ToString(),
							drReferencias.Cells["idProvincia"].Value.ToString(),
							drReferencias.Cells["idCanton"].Value.ToString(),
							drReferencias.Cells["Direccion"].Value.ToString(),
							drReferencias.Cells["Telefono"].Value.ToString(),
							drReferencias.Cells["Celular"].Value.ToString(),
							drReferencias.Cells["EmpresaTrabajo"].Value.ToString(),
							drReferencias.Cells["DireccionTrabajo"].Value.ToString(),
							drReferencias.Cells["TelefonoTrabajo"].Value.ToString(),
							drReferencias.Cells["ExtensionTrabajo"].Value.ToString()); 
						oCmdActualiza.CommandText = sSQLReferencia;
						//  			this.txtidverificacion.Value= (int)oCmdActualiza.ExecuteScalar();
						oCmdActualiza.ExecuteNonQuery();	

					}
				
                
					#endregion Referencia

					#region Negocios 
//				if (((int)this.cmbEntidad.SelectedValue == 2 && Convert.ToInt32(this.cmbClienteFuenteDeIngresos.SelectedValue) > 1) || ((int)this.cmbEntidad.SelectedValue != 2 && (Convert.ToInt32(this.cmbClienteSituacionLaboral.SelectedValue) != 1 && Convert.ToInt32(this.cmbClienteSituacionLaboral.SelectedValue) != 4)))
					if (((int)this.cmbEntidad.Value == 2 && (int)this.cmbSituacionLaboral.Value > 10) || ((int)this.cmbEntidad.Value != 2 && ((int)this.cmbSituacionLaboral.Value != 1 && (int) this.cmbSituacionLaboral.Value != 4)))
					{

						string sSQLGuardaNegocio = string.Format("Exec Cre_GuardaNegocios {0}, {1}, '{2}', '{3}', '{4}', {5}, {6}, {7}, {8}, {9}, {10},'{11}',{12}",
							(int)this.txtClienteIdCre_Negocio.Value, 
							(int)this.txtClienteIdCre_DatosGenerales.Value,
							this.txtNegocioRuc.Text.ToString(),
							this.txtNombreDelNegocio.Text.ToString(), 
							Convert.ToDateTime(this.dtpNegocioFechaInicio.Value).ToString("yyyyMMdd"),
							iNegciotipo, 
							iNEgocioActivdadEconomica,
							(int)this.txtNegocioMetros.Value, 
							(int)this.txtNegocioEmpleados.Value, 
							(bool)this.chkNegocioObligadoaLlevar.Checked,
							(bool)this.chkAfiliadoTributario.Checked,
							this.txtNotas.Text,
							cmbTiempo);   
						oCmdActualiza.CommandText = sSQLGuardaNegocio;

						
					

						this.txtClienteIdCre_Negocio.Value  = (int)oCmdActualiza.ExecuteScalar();


						string sSQLDatosUbicacionNegocio = string.Format("Exec Cre_GuardaDatosUbicacion {0}, {1}, {2}, {3}, {4}, {5}, {6}, '{7}', '{8}', '{9}', '{10}', '{11}', '{12}', '{13}', '{14}', '{15}', '{16}', {17},'{18}', '{19}', '{20}', {21}, {22}, {23}, {24}",
							(int)this.txtNegocioIdCre_DatosUbicacion.Value,
							3,
							(int)this.txtClienteIdCre_Negocio.Value,
							iNegocioProvincia, 
							iNegocioCanton,
							iNegocioParroquia,
							iNegocioBarrio,
							this.txtNegocioCallePrincipal.Text.ToString(), 
							this.txtNegocioNumeroCasa.Text.ToString(),
							this.txtNegocioCalleSecundaria.Text.ToString(), 
							this.txtNegocioReferncia.Text.ToString(),
							this.txtNegocioTelefono1.Text.ToString(), 
							"", this.txtNegocioTelefono2.Text.ToString(), 
							"",  
							this.txtNEgocioCelular.Text.ToString(),
							this.txtNegocioemail.Text.ToString(),
							0, Convert.ToDateTime(DateTime.Today).ToString("yyyyMMdd"), "", "", 
							(int)this.txtNEgocioIngresos.Value,
							(int)this.txtNEgocioGastos.Value, 
							(int)this.txtNegocioActivos.Value, 
							(int)this.txtNEgocioPasivos.Value);     
						oCmdActualiza.CommandText = sSQLDatosUbicacionNegocio;

						this.txtNegocioIdCre_DatosUbicacion.Value= (int)oCmdActualiza.ExecuteScalar();
						
					}
					#endregion Negocios

					#region Empresa de Trabajo
//					if (((int)this.cmbEntidadFinanciera.SelectedValue == 2 && Convert.ToInt32(this.cmbClienteFuenteDeIngresos.SelectedValue) == 1) || 
//						((int)this.cmbEntidadFinanciera.SelectedValue != 2 && (Convert.ToInt32(this.cmbClienteSituacionLaboral.SelectedValue) == 1 || Convert.ToInt32(this.cmbClienteSituacionLaboral.SelectedValue) == 3)))

					if (((int)this.cmbEntidad.Value == 2 && Convert.ToInt32(this.cmbSituacionLaboral.Value) == 10) || 
						((int)this.cmbEntidad.Value != 2 && (Convert.ToInt32(this.cmbSituacionLaboral.Value) == 1 || 
						Convert.ToInt32(this.cmbSituacionLaboral.Value) == 3)))
					
				 {
						string sSQLGuardaEmpresaTrabajo = string.Format("Exec Cre_GuardaEmpresaTrabajo {0}, {1}, '{2}', '{3}', {4}, '{5}', {6}, {7}, {8}, {9}, {10}, '{11}', {12}, {13}, {14}, {15}",
											(int)this.txtClienteIdCre_EmpresasTrabajo.Value, 
											(int)this.txtClienteIdCre_DatosGenerales.Value,
											this.txtDependienteRuc.Text, 
											this.txtDependienteEmpresa.Text, 
											iDependienteTipoempresa, 
						         	Convert.ToDateTime(this.dtDependienteFechaIngreso.Value).ToString("yyyyMMdd"),
											0,
											iDependienteTipocontrato, 
											iDependienteTipoSueldo, 
											this.chkAfiliadoIESS.Checked, 
											(int)this.txtDependienteDiasPago.Value,
											this.txtDependienteDepartamento.Text, 
											iDependienteCargo, 
											(int)this.txtDependienteDiasPago.Value, 
											0,cmbTiempoDependiente);   

					
						

						oCmdActualiza.CommandText = sSQLGuardaEmpresaTrabajo;

						this.txtClienteIdCre_EmpresasTrabajo.Value  = (int)oCmdActualiza.ExecuteScalar();


						string sSQLDatosUbicacionUbicacion = string.Format("Exec Cre_GuardaDatosUbicacion {0}, {1}, {2}, {3}, {4}, {5}, {6}, '{7}', '{8}', '{9}', '{10}', '{11}', '{12}', '{13}', '{14}', '{15}', '{16}', {17},'{18}', '{19}', '{20}', {21}, {22}, {23}, {24}",
									(int)this.txtEmpresaIdCre_DatosUbicacion.Value, 
									4,
									(int)this.txtClienteIdCre_EmpresasTrabajo.Value,
									iDependienteProvincia, 
									iDependienteCanton,
									iDependienteParroquia, 
									iDependienteBarrio,
									this.txtDependienteCallePrinciapl.Text.ToString(), 
									this.txtDependienteNumeroCasa.Text.ToString(), 
									this.txtDependienteCalleSecundaria.Text.ToString(), 
									this.txtDependienteReferencia.Text.ToString(),
									this.txtDependienteTelefono.Text.ToString(), 
									this.txtDependienteEXT.Text.ToString(), 
									"", "", 
									this.txtDependienteCelular.Text.ToString(), 
									this.txtDependienteemail.Text.ToString(), 
									0, Convert.ToDateTime(DateTime.Today).ToString("yyyyMMdd"), "", "",
									(int)this.txtDependienteIngresos.Value, 
									(int)this.txtDependienteGastos.Value, 
									(int)this.txtDependienteActivos.Value, 
									(int)this.txtDependientePasivos.Value);     
						oCmdActualiza.CommandText = sSQLDatosUbicacionUbicacion;
					
						this.txtEmpresaIdCre_DatosUbicacion.Value  = (int)oCmdActualiza.ExecuteScalar();
	
					}
					#endregion Empresa de Trabajo

					#region Observaciones
					string  sSQLGuardarObservaciones = string.Format("Exec Cre_GuardarObservaciones {0}, {1}, {2}, {3}, {4}, {5}, {6}, {7}, '{8}', {9}, {10}, {11}" ,
						this.txtIdCre_Solicitud.Value, 
						(int)this.cmbEntidad.Value, 
						iCreditoCalificacion,
						0,
						(int) this.txtCreditoScore.Value,
						0,
						iCreditoEstado,
						(int) this.cmbBodega.Value,
						this.txtObservaciones.Text,
						2,
						Convert.ToDecimal (this.txtCupo.Value),
						0);

					oCmdActualiza.CommandText = sSQLGuardarObservaciones;
					oCmdActualiza.ExecuteNonQuery();	
					#endregion Observaciones


					string  sSQLDetwecopia = string.Format("Exec Cre_CopiaCliente '{0}'" ,
					this.txtDocumento.Text.ToString()  );

					oCmdActualiza.CommandText = sSQLDetwecopia;
					oCmdActualiza.ExecuteNonQuery();	



					oTransaction.Commit();

					this.lblEstado.Text = "COMPLETA";

					#region  Observaciones
					this.gdObservaciones.DataSource = Funcion.dvProcedimiento(cdsSeteoF, string.Format("Exec Cre_ConsultaObservacionesSolicitud '{0}'",(int)this.txtIdCre_Solicitud.Value));
					#endregion  Observaciones

					//Realiza la consulta de todo nuevamente.
					txtDocumento_Validated(sender, System.EventArgs.Empty);

//					this.btnImprimir.Enabled = true;

					MessageBox.Show("Registro Guardado Correctamente", "Información", MessageBoxButtons.OK, MessageBoxIcon.Information);
					
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
			#endregion Guarda Cuota



			//MessageBox.Show("Registro Guardado Correctamente ", "Informacion", MessageBoxButtons.OK, MessageBoxIcon.Information);  
			
		}

		private void btnConsultar_Click(object sender, System.EventArgs e)
		{
			using (frmConsultaSolicitudes consultaSoli = new frmConsultaSolicitudes())
			{
				if (DialogResult.OK == consultaSoli.ShowDialog())
				{
			
				}	
			}
		}

		private void txtDPCApellidoM_ValueChanged(object sender, System.EventArgs e)
		{
		
		}

		private void cmbEntidad_ValueChanged(object sender, System.EventArgs e)
		{
			if (this.cmbEntidad.ActiveRow != null)
			{
				//this.cmbSituacionLaboral.DataSource = Funcion.dvProcedimiento(cdsSeteoF, string.Format("Exec Listas 66, {0}, 0",4));
				//this.cmbDPCSituacionL.DataSource = Funcion.dvProcedimiento(cdsSeteoF, string.Format("Exec Listas 66, {0}, 0",4));
			
				if ((int)this.cmbEntidad.Value == 2)
				{
					this.label25.Text = "Fuente de Ingresos";
					this.label45.Text = "Fuente de Ingresos";
					this.cmbActividadEconomica.Value = 0;
						this.cmbDPCActividadE.Value = 0;
					
				}
				else
				{
					this.label25.Text = "Situación Laboral";
						this.label45.Text = "Situación Laboral";
					this.cmbActividadEconomica.Value = 0;
						this.cmbDPCActividadE.Value = 0;
				}
			}

		}

		private void cmbSituacionLaboral_ValueChanged(object sender, System.EventArgs e)
		{
			if (this.cmbEntidad.ActiveRow != null)
			{
				SituacionLaboral((int)this.cmbSituacionLaboral.Value);
			}
	
		}


		private void SituacionLaboral(int idSitucion)
		{
			if (idSitucion == 1)
			{
				if (this.cmbSituacionLaboral.Value != null)
				{

					if ((int)this.cmbSituacionLaboral.Value == 1)
					{
						DesabilitaDatosNegocio();
						Habilita_DatosDependiente ();
		        this.btnTrabajoTelefono.Enabled = true;
						this.btnTrabajoCelular.Enabled = true;

						this.btnNegocioTelefono1.Enabled = false;
						this.btnNegocioTelefono2.Enabled = false;
						this.btnNegocioCelular.Enabled = false;
					}
					else if ((int)this.cmbSituacionLaboral.Value == 3)
					{
							
					HabilitaDatosNegocio ();
					Habilita_DatosDependiente ();
			     	this.btnNegocioTelefono1.Enabled = true;
				    this.btnNegocioTelefono2.Enabled = true;
						this.btnNegocioCelular.Enabled = true;

				    this.btnTrabajoTelefono.Enabled = true;
						this.btnTrabajoCelular.Enabled = true;
					}
					else
					{
						HabilitaDatosNegocio ();
						Desabilita_DatosDependiente ();
						this.btnNegocioTelefono1.Enabled = true;
						this.btnNegocioTelefono2.Enabled = true;
							this.btnNegocioCelular.Enabled = true;
						this.btnTrabajoTelefono.Enabled = false;
						this.btnTrabajoCelular.Enabled = false;
					}
					//this.cmbActividadEconomica.DataSource = Funcion.dvProcedimiento(cdsSeteoF, string.Format("Exec Listas 27, 1, {0}", 0));
					this.cmbActividadEconomica.DataSource = Funcion.dvProcedimiento(cdsSeteoF, string.Format("Exec ListaActividadEconomica {0}", idSitucion));

	
					this.cmbActividadEconomica.Value = 0;
				}
			}
			else
			{
				if (this.cmbSituacionLaboral.Value != null)
				{
					if ((int)this.cmbSituacionLaboral.Value == 10)
					{
						  DesabilitaDatosNegocio ();
							Habilita_DatosDependiente ();
						this.btnTrabajoTelefono.Enabled = true;
						this.btnTrabajoCelular.Enabled = true;

						this.btnNegocioTelefono1.Enabled = false;
						this.btnNegocioTelefono2.Enabled = false;
						this.btnNegocioCelular.Enabled = false;
					}
					else
					{
						HabilitaDatosNegocio ();
						Desabilita_DatosDependiente ();
						this.btnNegocioTelefono1.Enabled = true;
						this.btnNegocioTelefono2.Enabled = true;
							this.btnNegocioCelular.Enabled = true;

						this.btnTrabajoTelefono.Enabled = false;
						this.btnTrabajoCelular.Enabled = false;
					}

					int iSituacion = 0;
					if ((int) this.cmbSituacionLaboral.Value == 10) iSituacion = 1;
					if ((int) this.cmbSituacionLaboral.Value == 11) iSituacion = 2;
					if ((int) this.cmbSituacionLaboral.Value == 12) iSituacion = 3;
					//this.cmbActividadEconomica.DataSource = Funcion.dvProcedimiento(cdsSeteoF, string.Format("Exec Listas 27, 2, {0}", iSituacion));
					this.cmbActividadEconomica.DataSource =  Funcion.dvProcedimiento(cdsSeteoF, string.Format("Exec ListaActividadEconomica {0}", idSitucion));
					this.cmbActividadEconomica.Value = 0;
				}
			}
		}
		private void SituacionLaboralConyuge(int idSitucion)
		{
//			if (idEntidadFinanciera != 2)
//			{
//				if (this.cmbDPCSituacionL.Value != null)
//				{
//
//					//this.cmbDPCActividadE.DataSource = Funcion.dvProcedimiento(cdsSeteoF, string.Format("Exec Listas 27, 1, {0}", 0));
//					this.cmbDPCActividadE.DataSource = Funcion.dvProcedimiento(cdsSeteoF, string.Format("Exec ListaActividadEconomica {0}", idSitucion));
//					this.cmbDPCActividadE.Value = 0;
//				}
//			}
//			else
//			{
				if (this.cmbDPCSituacionL.Value != null)
				{
					

					int iSituacion = 0;
					if ((int) this.cmbDPCSituacionL.Value == 10) iSituacion = 1;
					if ((int) this.cmbDPCSituacionL.Value == 11) iSituacion = 2;
					if ((int) this.cmbDPCSituacionL.Value == 12) iSituacion = 3;
					//this.cmbDPCActividadE.DataSource = Funcion.dvProcedimiento(cdsSeteoF, string.Format("Exec Listas 27, 2, {0}", iSituacion));
					this.cmbDPCActividadE.DataSource = Funcion.dvProcedimiento(cdsSeteoF, string.Format("Exec ListaActividadEconomica {0}", idSitucion));
					this.cmbDPCActividadE.Value = 0;
				}
			//}
		}

		private void cmbClienteFuenteDeIngresos_ValueChanged(object sender, System.EventArgs e)
		{
			if (this.cmbEntidad.ActiveRow != null)
			{
				//SituacionLaboral((int)this.cmbEntidad.Value);
				SituacionLaboral((int)this.cmbSituacionLaboral.Value);

			}
		}

		private void cmbClienteFuenteDeIngresos_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if(e.KeyCode == Keys.Enter) this.cmbActividadEconomica.Focus();	
		}

		private void ultgconsulta_CellChange(object sender, Infragistics.Win.UltraWinGrid.CellEventArgs e)
		{



		}

		private void btnCelularcliente_Click(object sender, System.EventArgs e)
		{
			if (this.txtDDCelular.Text.ToString().Trim().Length == 0  )
			{
				MessageBox.Show(string.Format("Ingrese un Número Celular "), "Point Technology", MessageBoxButtons.OK,MessageBoxIcon.Stop);
				this.txtDDCelular.Focus();
				return;	
        
			}
			if (!Validacion.vbValidaTelefono(this.txtDDCelular, 2, false)) return;
       int iNuevoCel =0;
			int iCelExistente = 0;
			
        iNuevoCel = Funcion.iEscalarSQL(cdsSeteoF, string.Format("select count(*) from Cre_DatosUbicacion where Origen = 0 and idOrigen = {0} and celular like '%"+"{1}"+"%' ", (int)this.txtIdCre_Solicitud.Value, this.txtDDCelular.Text.ToString()));

			if((int)this.txtIdCre_Solicitud.Value  > 0)
			{
				if (iNuevoCel > 0)
				{
											AbreVerificaciontelefonica (false, (int)this.txtIdCre_Solicitud.Value, 11,
												this.txtApellidoP.Text + " " + this.txtApellidoM.Text + " " + this.txtPnombre.Text + " " + this.txtSnombre.Text,
												this.txtDDCelular.Text, 0); 
				}
				else
				{
					MessageBox.Show(string.Format("Guarde los datos del cliente para realizar la verificacion telefonica"), "Point Technology", MessageBoxButtons.OK,MessageBoxIcon.Stop);
											this.txtDDCelular.Focus();
											return;	
				}

			}
			else
			{
				MessageBox.Show(string.Format("Guarde los datos del cliente para realizar la verificacion telefonica"), "Point Technology", MessageBoxButtons.OK,MessageBoxIcon.Stop);
				this.txtDDCelular.Focus();
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

		private void btnClienteTelefonoCasa1_Click(object sender, System.EventArgs e)
		{
			if (this.txtDDTelefono1.Text.ToString().Trim().Length == 0  )
			{
				MessageBox.Show(string.Format("Ingrese un Número Celular "), "Point Technology", MessageBoxButtons.OK,MessageBoxIcon.Stop);
				this.txtDDTelefono1.Focus();
				return;	
        
			}
			if (!Validacion.vbValidaTelefono(this.txtDDTelefono1, 1, false)) return;
			int iNuevoCel =0;
			int iCelExistente = 0;
			
			iNuevoCel = Funcion.iEscalarSQL(cdsSeteoF, string.Format("select count(*) from Cre_DatosUbicacion where Origen = 0 and idOrigen = {0} and Telefono1 like '%"+"{1}"+"%' ", (int)this.txtIdCre_Solicitud.Value, this.txtDDTelefono1.Text.ToString()));

			if((int)this.txtIdCre_Solicitud.Value  > 0)
			{
				if (iNuevoCel > 0)
				{
					AbreVerificaciontelefonica (false, (int)this.txtIdCre_Solicitud.Value, 0,
						this.txtApellidoP.Text + " " + this.txtApellidoM.Text + " " + this.txtPnombre.Text + " " + this.txtSnombre.Text,
						this.txtDDTelefono1.Text, 0); 
				}
				else
				{
					MessageBox.Show(string.Format("Guarde los datos del cliente para realizar la verificacion telefonica"), "Point Technology", MessageBoxButtons.OK,MessageBoxIcon.Stop);
					this.txtDDTelefono1.Focus();
					return;	
				}

			}
			else
			{
				MessageBox.Show(string.Format("Guarde los datos del cliente para realizar la verificacion telefonica"), "Point Technology", MessageBoxButtons.OK,MessageBoxIcon.Stop);
				this.txtDDTelefono1.Focus();
				return;	
			}
		}

		private void btnClienteTelefonoCasa2_Click(object sender, System.EventArgs e)
		{
			if (this.txtDDTelefono2.Text.ToString().Trim().Length == 0  )
			{
				MessageBox.Show(string.Format("Ingrese un Número Celular "), "Point Technology", MessageBoxButtons.OK,MessageBoxIcon.Stop);
				this.txtDDTelefono2.Focus();
				return;	
        
			}
			if (!Validacion.vbValidaTelefono(this.txtDDTelefono2, 1, false)) return;
			int iNuevoCel =0;
			iNuevoCel = Funcion.iEscalarSQL(cdsSeteoF, string.Format("select count(*) from Cre_DatosUbicacion where Origen = 0 and idOrigen = {0} and Telefono2 like '%"+"{1}"+"%' ", (int)this.txtIdCre_Solicitud.Value, this.txtDDTelefono2.Text.ToString()));

			if((int)this.txtIdCre_Solicitud.Value  > 0)
			{
				if (iNuevoCel > 0)
				{
					AbreVerificaciontelefonica (false, (int)this.txtIdCre_Solicitud.Value, 1,
						this.txtApellidoP.Text + " " + this.txtApellidoM.Text + " " + this.txtPnombre.Text + " " + this.txtSnombre.Text,
						this.txtDDTelefono2.Text, 0); 
				}
				else
				{
					MessageBox.Show(string.Format("Guarde los datos del cliente para realizar la verificacion telefonica"), "Point Technology", MessageBoxButtons.OK,MessageBoxIcon.Stop);
					this.txtDDTelefono2.Focus();
					return;	
				}

			}
			else
			{
				MessageBox.Show(string.Format("Guarde los datos del cliente para realizar la verificacion telefonica"), "Point Technology", MessageBoxButtons.OK,MessageBoxIcon.Stop);
				this.txtDDTelefono1.Focus();
				return;	
			}
		}

		private void btnConyugueTelefono_Click(object sender, System.EventArgs e)
		{
			if (this.txtDPCTtelefono.Text.ToString().Trim().Length == 0  )
			{
				MessageBox.Show(string.Format("Ingrese un Número Celular "), "Point Technology", MessageBoxButtons.OK,MessageBoxIcon.Stop);
				this.txtDPCTtelefono.Focus();
				return;	
        
			}
			if (!Validacion.vbValidaTelefono(this.txtDPCTtelefono, 1, false)) return;
			int iNuevoCel =0;
			iNuevoCel = Funcion.iEscalarSQL(cdsSeteoF, string.Format("select count(*) from Cre_DatosUbicacion where Origen = 1 and idOrigen = {0} and Telefono1 like '%"+"{1}"+"%' ", (int)this.txtIdCre_Conyugue.Value, this.txtDPCTtelefono.Text.ToString()));

			if((int)this.txtIdCre_Solicitud.Value  > 0)
			{
				if (iNuevoCel > 0)
				{
					AbreVerificaciontelefonica (false, (int)this.txtIdCre_Solicitud.Value, 3,
						this.txtApellidoP.Text + " " + this.txtApellidoM.Text + " " + this.txtPnombre.Text + " " + this.txtSnombre.Text,
						this.txtDPCTtelefono.Text, 0); 
				}
				else
				{
					MessageBox.Show(string.Format("Guarde los datos del cliente para realizar la verificacion telefonica"), "Point Technology", MessageBoxButtons.OK,MessageBoxIcon.Stop);
					this.txtDPCTtelefono.Focus();
					return;	
				}

			}
			else
			{
				MessageBox.Show(string.Format("Guarde los datos del cliente para realizar la verificacion telefonica"), "Point Technology", MessageBoxButtons.OK,MessageBoxIcon.Stop);
				this.txtDDTelefono1.Focus();
				return;	
			}
		}

		private void button3_Click(object sender, System.EventArgs e)
		{
		
		}

		private void btnConyugueCelular_Click(object sender, System.EventArgs e)
		{
			if (this.txtDPCTCelularC.Text.ToString().Trim().Length == 0  )
			{
				MessageBox.Show(string.Format("Ingrese un Número Celular "), "Point Technology", MessageBoxButtons.OK,MessageBoxIcon.Stop);
				this.txtDPCTCelularC.Focus();
				return;	
        
			}
			if (!Validacion.vbValidaTelefono(this.txtDPCTCelularC, 2, false)) return;
			int iNuevoCel =0;
			iNuevoCel = Funcion.iEscalarSQL(cdsSeteoF, string.Format("select count(*) from Cre_DatosUbicacion where Origen = 1 and idOrigen = {0} and Celular like '%"+"{1}"+"%' ", (int)this.txtIdCre_Conyugue.Value, this.txtDPCTCelularC.Text.ToString()));

			if((int)this.txtIdCre_Solicitud.Value  > 0)
			{
				if (iNuevoCel > 0)
				{
					AbreVerificaciontelefonica (false, (int)this.txtIdCre_Solicitud.Value, 4,
						this.txtApellidoP.Text + " " + this.txtApellidoM.Text + " " + this.txtPnombre.Text + " " + this.txtSnombre.Text,
						this.txtDPCTCelularC.Text, 0); 
				}
				else
				{
					MessageBox.Show(string.Format("Guarde los datos del cliente para realizar la verificacion telefonica"), "Point Technology", MessageBoxButtons.OK,MessageBoxIcon.Stop);
					this.txtDPCTCelularC.Focus();
					return;	
				}

			}
			else
			{
				MessageBox.Show(string.Format("Guarde los datos del cliente para realizar la verificacion telefonica"), "Point Technology", MessageBoxButtons.OK,MessageBoxIcon.Stop);
				this.txtDPCTCelularC.Focus();
				return;	
			}
		}

		private void btnTrabajoTelefono_Click(object sender, System.EventArgs e)
		{
			if (this.txtDependienteTelefono.Text.ToString().Trim().Length == 0  )
			{
				MessageBox.Show(string.Format("Ingrese un Número Celular "), "Point Technology", MessageBoxButtons.OK,MessageBoxIcon.Stop);
				this.txtDependienteTelefono.Focus();
				return;	
        
			}
			if (!Validacion.vbValidaTelefono(this.txtDependienteTelefono, 1, false)) return;
			int iNuevoCel =0;
			iNuevoCel = Funcion.iEscalarSQL(cdsSeteoF, string.Format("select count(*) from Cre_DatosUbicacion where Origen = 4 and idOrigen = {0} and Telefono1 like '%"+"{1}"+"%' ", (int)this.txtClienteIdCre_EmpresasTrabajo.Value, this.txtDependienteTelefono.Text.ToString()));

			if((int)this.txtIdCre_Solicitud.Value  > 0)
			{
				if (iNuevoCel > 0)
				{
					AbreVerificaciontelefonica (false, (int)this.txtIdCre_Solicitud.Value, 7,
						this.txtApellidoP.Text + " " + this.txtApellidoM.Text + " " + this.txtPnombre.Text + " " + this.txtSnombre.Text,
						this.txtDependienteTelefono.Text, 0); 
				}
				else
				{
					MessageBox.Show(string.Format("Guarde los datos del cliente para realizar la verificacion telefonica"), "Point Technology", MessageBoxButtons.OK,MessageBoxIcon.Stop);
					this.txtDependienteTelefono.Focus();
					return;	
				}

			}
			else
			{
				MessageBox.Show(string.Format("Guarde los datos del cliente para realizar la verificacion telefonica"), "Point Technology", MessageBoxButtons.OK,MessageBoxIcon.Stop);
				this.txtDDTelefono1.Focus();
				return;	
			}
		}

		private void btnTrabajoCelular_Click(object sender, System.EventArgs e)
		{
			if (this.txtDependienteCelular.Text.ToString().Trim().Length == 0  )
			{
				MessageBox.Show(string.Format("Ingrese un Número Celular "), "Point Technology", MessageBoxButtons.OK,MessageBoxIcon.Stop);
				this.txtDependienteCelular.Focus();
				return;	
        
			}
			if (!Validacion.vbValidaTelefono(this.txtDependienteCelular, 2, false)) return;
			int iNuevoCel =0;
			iNuevoCel = Funcion.iEscalarSQL(cdsSeteoF, string.Format("select count(*) from Cre_DatosUbicacion where Origen = 4 and idOrigen = {0} and Celular like '%"+"{1}"+"%' ", (int)this.txtClienteIdCre_EmpresasTrabajo.Value, this.txtDependienteCelular.Text.ToString()));

			if((int)this.txtIdCre_Solicitud.Value  > 0)
			{
				if (iNuevoCel > 0)
				{
					AbreVerificaciontelefonica (false, (int)this.txtIdCre_Solicitud.Value, 7,
						this.txtApellidoP.Text + " " + this.txtApellidoM.Text + " " + this.txtPnombre.Text + " " + this.txtSnombre.Text,
						this.txtDependienteCelular.Text, 0); 
				}
				else
				{
					MessageBox.Show(string.Format("Guarde los datos del cliente para realizar la verificacion telefonica"), "Point Technology", MessageBoxButtons.OK,MessageBoxIcon.Stop);
					this.txtDependienteCelular.Focus();
					return;	
				}

			}
			else
			{
				MessageBox.Show(string.Format("Guarde los datos del cliente para realizar la verificacion telefonica"), "Point Technology", MessageBoxButtons.OK,MessageBoxIcon.Stop);
				this.txtDependienteCelular.Focus();
				return;	
			}
		}

		private void btnNegocioTelefono1_Click(object sender, System.EventArgs e)
		{
			
					if (this.txtNegocioTelefono1.Text.ToString().Trim().Length == 0  )
					{
						MessageBox.Show(string.Format("Ingrese un Número Celular "), "Point Technology", MessageBoxButtons.OK,MessageBoxIcon.Stop);
						this.txtNegocioTelefono1.Focus();
						return;	
        
					}
			if (!Validacion.vbValidaTelefono(this.txtNegocioTelefono1, 1, false)) return;
			int iNuevoCel =0;
			iNuevoCel = Funcion.iEscalarSQL(cdsSeteoF, string.Format("select count(*) from Cre_DatosUbicacion where Origen = 3 and idOrigen = {0} and Telefono1 like '%"+"{1}"+"%' ", (int)this.txtClienteIdCre_Negocio.Value, this.txtNegocioTelefono1.Text.ToString()));

			if((int)this.txtIdCre_Solicitud.Value  > 0)
			{
				if (iNuevoCel > 0)
				{
					AbreVerificaciontelefonica (false, (int)this.txtIdCre_Solicitud.Value, 5,
						this.txtApellidoP.Text + " " + this.txtApellidoM.Text + " " + this.txtPnombre.Text + " " + this.txtSnombre.Text,
						this.txtNegocioTelefono1.Text, 0); 
				}
				else
				{
					MessageBox.Show(string.Format("Guarde los datos del cliente para realizar la verificacion telefonica"), "Point Technology", MessageBoxButtons.OK,MessageBoxIcon.Stop);
					this.txtNegocioTelefono1.Focus();
					return;	
				}

			}
			else
			{
				MessageBox.Show(string.Format("Guarde los datos del cliente para realizar la verificacion telefonica"), "Point Technology", MessageBoxButtons.OK,MessageBoxIcon.Stop);
				this.txtNegocioTelefono1.Focus();
				return;	
			}
		}

		private void btnNegocioTelefono2_Click(object sender, System.EventArgs e)
		{
			if (this.txtNegocioTelefono2.Text.ToString().Trim().Length == 0  )
			{
				MessageBox.Show(string.Format("Ingrese un Número Celular "), "Point Technology", MessageBoxButtons.OK,MessageBoxIcon.Stop);
				this.txtNegocioTelefono2.Focus();
				return;	
        
			}
			if (!Validacion.vbValidaTelefono(this.txtNegocioTelefono2, 1, false)) return;
			int iNuevoCel =0;
			iNuevoCel = Funcion.iEscalarSQL(cdsSeteoF, string.Format("select count(*) from Cre_DatosUbicacion where Origen = 3 and idOrigen = {0} and Telefono2 like '%"+"{1}"+"%' ", (int)this.txtClienteIdCre_Negocio.Value, this.txtNegocioTelefono2.Text.ToString()));

			if((int)this.txtIdCre_Solicitud.Value  > 0)
			{
				if (iNuevoCel > 0)
				{
					AbreVerificaciontelefonica (false, (int)this.txtIdCre_Solicitud.Value, 6,
						this.txtApellidoP.Text + " " + this.txtApellidoM.Text + " " + this.txtPnombre.Text + " " + this.txtSnombre.Text,
						this.txtNegocioTelefono2.Text, 0); 
				}
				else
				{
					MessageBox.Show(string.Format("Guarde los datos del cliente para realizar la verificacion telefonica"), "Point Technology", MessageBoxButtons.OK,MessageBoxIcon.Stop);
					this.txtNegocioTelefono2.Focus();
					return;	
				}

			}
			else
			{
				MessageBox.Show(string.Format("Guarde los datos del cliente para realizar la verificacion telefonica"), "Point Technology", MessageBoxButtons.OK,MessageBoxIcon.Stop);
				this.txtNegocioTelefono1.Focus();
				return;	
			}

		}

		private void btnNegocioCelular_Click(object sender, System.EventArgs e)
		{
			if (this.txtNEgocioCelular.Text.ToString().Trim().Length == 0  )
			{
				MessageBox.Show(string.Format("Ingrese un Número Celular "), "Point Technology", MessageBoxButtons.OK,MessageBoxIcon.Stop);
				this.txtNEgocioCelular.Focus();
				return;	
        
			}
			if (!Validacion.vbValidaTelefono(this.txtNEgocioCelular, 2, false)) return;
			int iNuevoCel =0;
			iNuevoCel = Funcion.iEscalarSQL(cdsSeteoF, string.Format("select count(*) from Cre_DatosUbicacion where Origen = 3 and idOrigen = {0} and Celular like '%"+"{1}"+"%' ", (int)this.txtClienteIdCre_Negocio.Value, this.txtNEgocioCelular.Text.ToString()));

			if((int)this.txtIdCre_Solicitud.Value  > 0)
			{
				if (iNuevoCel > 0)
				{
					AbreVerificaciontelefonica (false, (int)this.txtIdCre_Solicitud.Value, 11,
						this.txtApellidoP.Text + " " + this.txtApellidoM.Text + " " + this.txtPnombre.Text + " " + this.txtSnombre.Text,
						this.txtNEgocioCelular.Text, 0); 
				}
				else
				{
					MessageBox.Show(string.Format("Guarde los datos del cliente para realizar la verificacion telefonica"), "Point Technology", MessageBoxButtons.OK,MessageBoxIcon.Stop);
					this.txtNEgocioCelular.Focus();
					return;	
				}

			}
			else
			{
				MessageBox.Show(string.Format("Guarde los datos del cliente para realizar la verificacion telefonica"), "Point Technology", MessageBoxButtons.OK,MessageBoxIcon.Stop);
				this.txtNEgocioCelular.Focus();
				return;	
			}
		}

		private void ultgconsulta_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			FuncionesProcedimientos.DesplazamientoGrids(sender,e,this.ultgconsulta);
		}

		private void ultgconsulta_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
			if (ultgconsulta.ActiveCell.Column.Key.ToString().ToUpper() == "TELEFONO" || ultgconsulta.ActiveCell.Column.Key.ToString().ToUpper() == "CELULAR"  || ultgconsulta.ActiveCell.Column.Key.ToString().ToUpper() == "TELEFONOTRABAJO")
			{
				Funcion.CamposNumericos(sender, e);}
		}

		private void ultgconsulta_ClickCellButton(object sender, Infragistics.Win.UltraWinGrid.CellEventArgs e)
		{
			if (e.Cell.Column.ToString() == "btnTelefono")
			{
				 #region Validacion Telefono Referencia 
							
				if(e.Cell.Row.Cells["Telefono"].Value.ToString().Length==0)
				{
					MessageBox.Show("Ingrese el Telefono de la Referencia", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
					this.tbBasico.Tabs[2].Selected = true;
					
					this.ultgconsulta.ActiveRow.Selected = true;
					return;
				}

				if (e.Cell.Row.Cells["Telefono"].Value.ToString().Trim().Length > 0)
				{
					if (e.Cell.Row.Cells["Telefono"].Value.ToString().Trim().Length != 9)
					{
						MessageBox.Show("El Telefono Convencional debe tener 9 Digitos.\n\nLos dos primeros corresponden al Codigo de Provincia", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
						this.tbBasico.Tabs[2].Selected = true;
						this.ultgconsulta.ActiveRow.Selected = true;
						return;
					}
				}
 
				int iNuevoCel =0;
				iNuevoCel = Funcion.iEscalarSQL(cdsSeteoF, string.Format("select count(*) from Cre_ReferenciasClientes where idCre_DatosGenerales =  {0} and Telefono like '%"+"{1}"+"%' ", (int)this.txtClienteIdCre_DatosGenerales.Value, e.Cell.Row.Cells["Telefono"].Value.ToString()));

				if((int)e.Cell.Row.Cells["idCre_ReferenciasClientes"].Value > 0)
				{
					if (iNuevoCel > 0)
					{
						AbreVerificaciontelefonica (false, (int)this.txtIdCre_Solicitud.Value, 8,
							e.Cell.Row.Cells["ApellidoPaterno"].Value.ToString() + " " + e.Cell.Row.Cells["ApellidoMaterno"].Value.ToString() + " " 
							+ e.Cell.Row.Cells["PrimerNombre"].Value.ToString()+ " " + e.Cell.Row.Cells["SegundoNombre"].Value.ToString(),
							e.Cell.Row.Cells["Telefono"].Value.ToString(), 0); 
					}
					else
					{
						MessageBox.Show(string.Format("Guarde los datos del cliente para realizar la verificacion telefonica"), "Point Technology", MessageBoxButtons.OK,MessageBoxIcon.Stop);
						this.tbBasico.Tabs[2].Selected = true;
						this.ultgconsulta.ActiveRow.Selected = true;
						return;
					}

				}
				else
				{
					MessageBox.Show(string.Format("Guarde los datos del cliente para realizar la verificacion telefonica"), "Point Technology", MessageBoxButtons.OK,MessageBoxIcon.Stop);
					this.tbBasico.Tabs[2].Selected = true;
					
					this.ultgconsulta.ActiveRow.Selected = true;
					return;
				}
	#endregion Validacion Telefono Referencia 
			}

			if (e.Cell.Column.ToString() == "btnCelular")
			{			
				
				if(e.Cell.Row.Cells["Celular"].Value.ToString().Length==0)
				{
					MessageBox.Show("Ingrese el Celular de la Referencia", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
					this.tbBasico.Tabs[2].Selected = true;
					this.ultgconsulta.ActiveRow.Selected = true;
					return;
				}

				#region Celular
				if (e.Cell.Row.Cells["Celular"].Value.ToString().Trim().Length > 0)
				{
					if (e.Cell.Row.Cells["Celular"].ToString().Trim().Length < 10 )
					{
						MessageBox.Show("El Telefono Celular debe tener 10 digitos", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
						this.tbBasico.Tabs[2].Selected = true;
						this.ultgconsulta.ActiveRow.Selected = true;
						return;
					}
//					if ( e.Cell.Row.Cells["Celular"].ToString().Trim().Length > 10)
//					{
//						MessageBox.Show("El Telefono Celular debe tener 10 digitos", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
//						this.tbBasico.Tabs[2].Selected = true;
//						this.ultgconsulta.ActiveRow.Selected = true;
//						return;
//					}

					if (e.Cell.Row.Cells["Celular"].Value.ToString().Trim().Length != 10)
					{
						MessageBox.Show("El Telefono Celular debe tener 10 digitos", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
						this.tbBasico.Tabs[2].Selected = true;
						this.ultgconsulta.ActiveRow.Selected = true;
						return;
					}
				}
				#endregion Celular

				int iNuevoCel =0;
				iNuevoCel = Funcion.iEscalarSQL(cdsSeteoF, string.Format("select count(*) from Cre_ReferenciasClientes where idCre_DatosGenerales =  {0} and Celular like '%"+"{1}"+"%' ", (int)this.txtClienteIdCre_DatosGenerales.Value, e.Cell.Row.Cells["Celular"].Value.ToString()));

				if((int)e.Cell.Row.Cells["idCre_ReferenciasClientes"].Value > 0)
				{
					if (iNuevoCel > 0)
					{
						AbreVerificaciontelefonica (false, (int)this.txtIdCre_Solicitud.Value, 8,
							e.Cell.Row.Cells["ApellidoPaterno"].Value.ToString() + " " + e.Cell.Row.Cells["ApellidoMaterno"].Value.ToString() + " " 
							+ e.Cell.Row.Cells["PrimerNombre"].Value.ToString()+ " " + e.Cell.Row.Cells["SegundoNombre"].Value.ToString(),
							e.Cell.Row.Cells["Celular"].Value.ToString(), 0); 
					}
					else
					{
						MessageBox.Show(string.Format("Guarde los datos del cliente para realizar la verificacion telefonica"), "Point Technology", MessageBoxButtons.OK,MessageBoxIcon.Stop);
						this.tbBasico.Tabs[2].Selected = true;
						this.ultgconsulta.ActiveRow.Selected = true;
						return;
					}

				}
				else
				{
					MessageBox.Show(string.Format("Guarde los datos del cliente para realizar la verificacion telefonica"), "Point Technology", MessageBoxButtons.OK,MessageBoxIcon.Stop);
					this.tbBasico.Tabs[2].Selected = true;
					
					this.ultgconsulta.ActiveRow.Selected = true;
					return;
				}

			}
			if (e.Cell.Column.ToString() == "btnTelefonoTrabj")
			{			
				#region Validacion Telefono Referencia 
							
				if(e.Cell.Row.Cells["TelefonoTrabajo"].Value.ToString().Length==0)
				{
					MessageBox.Show("Ingrese el Telefono de la Referencia del Trabajo", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
					this.tbBasico.Tabs[2].Selected = true;
					this.ultgconsulta.ActiveRow.Selected = true;
					return;
				}

				if (e.Cell.Row.Cells["TelefonoTrabajo"].Value.ToString().Trim().Length > 0)
				{
					if (e.Cell.Row.Cells["TelefonoTrabajo"].Value.ToString().Trim().Length != 9)
					{
						MessageBox.Show("El Telefono Convencional debe tener 9 Digitos.\n\nLos dos primeros corresponden al Codigo de Provincia", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
						this.tbBasico.Tabs[2].Selected = true;
						this.ultgconsulta.ActiveRow.Selected = true;
						return;
					}
				}
 
				int iNuevoCel =0;
				iNuevoCel = Funcion.iEscalarSQL(cdsSeteoF, string.Format("select count(*) from Cre_ReferenciasClientes where idCre_DatosGenerales =  {0} and TelefonoTrabajo like '%"+"{1}"+"%' ", (int)this.txtClienteIdCre_DatosGenerales.Value, e.Cell.Row.Cells["Telefono"].Value.ToString()));

				if((int)e.Cell.Row.Cells["idCre_ReferenciasClientes"].Value > 0)
				{
					if (iNuevoCel > 0)
					{
						AbreVerificaciontelefonica (false, (int)this.txtIdCre_Solicitud.Value, 8,
							e.Cell.Row.Cells["ApellidoPaterno"].Value.ToString() + " " + e.Cell.Row.Cells["ApellidoMaterno"].Value.ToString() + " " 
							+ e.Cell.Row.Cells["PrimerNombre"].Value.ToString()+ " " + e.Cell.Row.Cells["SegundoNombre"].Value.ToString(),
							e.Cell.Row.Cells["TelefonoTrabajo"].Value.ToString(), 0); 
					}
					else
					{
						MessageBox.Show(string.Format("Guarde los datos del cliente para realizar la verificacion telefonica"), "Point Technology", MessageBoxButtons.OK,MessageBoxIcon.Stop);
						this.tbBasico.Tabs[2].Selected = true;
						this.ultgconsulta.ActiveRow.Selected = true;
						return;
					}

				}
				else
				{
					MessageBox.Show(string.Format("Guarde los datos del cliente para realizar la verificacion telefonica"), "Point Technology", MessageBoxButtons.OK,MessageBoxIcon.Stop);
					this.tbBasico.Tabs[2].Selected = true;
					
					this.ultgconsulta.ActiveRow.Selected = true;
					return;
				}
				#endregion Validacion Telefono Referencia 
			}

		}

		private void btnImprimir_Click(object sender, System.EventArgs e)
		{
			ParameterFields paramFields = new ParameterFields();

//			ParameterFields paramFields = new ParameterFields ();
				
			ParameterField paramFieldCedula = new ParameterField ();
			ParameterDiscreteValue ValCedula = new ParameterDiscreteValue ();
			paramFieldCedula.ParameterFieldName = "@Numero";
			ValCedula.Value = this.txtDocumento.Text;
			paramFieldCedula.CurrentValues.Add (ValCedula);
			paramFields.Add (paramFieldCedula);

			ParameterField pfIdOrigen = new ParameterField();
			ParameterDiscreteValue ValIdOrigen = new ParameterDiscreteValue();
			pfIdOrigen.ParameterFieldName = "@idOrigen";
			ValIdOrigen.Value = Convert.ToInt32(this.txtIdCre_Solicitud.Value);
			pfIdOrigen.CurrentValues.Add(ValIdOrigen);
			paramFields.Add(pfIdOrigen);

			ParameterField pfIdCreConyugue = new ParameterField();
			ParameterDiscreteValue ValIdCreConyugue = new ParameterDiscreteValue();
			pfIdCreConyugue.ParameterFieldName = "@idCre_Conyugue";
			ValIdCreConyugue.Value = Convert.ToInt32(this.txtIdCre_Conyugue.Value);
			pfIdCreConyugue.CurrentValues.Add(ValIdCreConyugue);
			paramFields.Add(pfIdCreConyugue);

			ParameterField pfidCreDatosgenerales = new ParameterField();
			ParameterDiscreteValue ValidCreDatosgenerales = new ParameterDiscreteValue();
			pfidCreDatosgenerales.ParameterFieldName = "@idCre_DatosGenerales";
			ValidCreDatosgenerales.Value = Convert.ToInt32(this.txtClienteIdCre_DatosGenerales.Value);
			pfidCreDatosgenerales.CurrentValues.Add(ValidCreDatosgenerales);
			paramFields.Add(pfidCreDatosgenerales);

			ParameterField pfidCreNegocios = new ParameterField();
			ParameterDiscreteValue ValidCreNegocios = new ParameterDiscreteValue();
			pfidCreNegocios.ParameterFieldName = "@idCre_Negocio";
			ValidCreNegocios.Value = Convert.ToInt32(this.txtClienteIdCre_Negocio.Value);
			pfidCreNegocios.CurrentValues.Add(ValidCreNegocios);
			paramFields.Add(pfidCreNegocios);

			ParameterField pfidCre_EmpresaTrabajo = new ParameterField();
			ParameterDiscreteValue ValidCre_EmpresaTrabajo = new ParameterDiscreteValue();
			pfidCre_EmpresaTrabajo.ParameterFieldName = "@idCre_EmpresaTrabajo";
			ValidCre_EmpresaTrabajo.Value = Convert.ToInt32(this.txtClienteIdCre_EmpresasTrabajo.Value);
			pfidCre_EmpresaTrabajo.CurrentValues.Add(ValidCre_EmpresaTrabajo);
			paramFields.Add(pfidCre_EmpresaTrabajo);

			Reporte miRep = new Reporte("RepSolicitud.rpt", "");
			miRep.MdiParent = this.MdiParent;
			miRep.Titulo("Solicitudes De Credito");
			miRep.crVista.ParameterFieldInfo = paramFields;
			miRep.Show();

		}

		private void btnSalir_Click(object sender, System.EventArgs e)
		{
			this.Close();
		}

		private void dtTiempoVivienda_KeyDown_1(object sender, System.Windows.Forms.KeyEventArgs e)
		{
//			if(e.KeyCode == Keys.Enter) 
//
//			{
//				if (this.cmbDDTipoVivienda.ActiveRow != null)
//				{
//					if ((int)this.cmbDDTipoVivienda.Value == 2 || (int)this.cmbDDTipoVivienda.Value == 6 )
//
//					{
//						this.tbBasico.Tabs[1].Selected = true;
//						this.txtDPCdocumento.Focus ();
//						if ( this.cmbDPCTipodocumento.Enabled == false) 
//						{ 
//							this.tbBasico.Tabs[2].Selected = true;
//						}
//
//					}
//					if ((int)this.cmbDDTipoVivienda.Value == 5 )
//
//					{
//						this.tbBasico.Tabs[1].Selected = true;
//						this.txtDPCdocumento.Focus ();
//						if ( this.cmbDPCTipodocumento.Enabled == false) 
//						{
//							this.tbBasico.Tabs[2].Selected = true;
//						}
//					}
//					if ((int)this.cmbDDTipoVivienda.Value == 1 )
//					{
//					
//						this.txtNombreArrendador.Focus ();
//					}
//					if ((int)this.cmbDDTipoVivienda.Value == 3 || (int)this.cmbDDTipoVivienda.Value == 4)
//					{
//					
//						this.cmbDDInmueble.Focus ();
//					}
//				}
//			}

		}

		private void cmbDDTipoVivienda_KeyDown_1(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			//if(e.KeyCode == Keys.Enter) this.dtTiempoVivienda.Focus();
		}

		private void cmbDDTipoVivienda_ValueChanged_1(object sender, System.EventArgs e)
		{
			TipoVivienda();
		}

		private void VaciaCamposGenerales()
		{
			this.cmbTipoRuc.Value=1;
			this.txtApellidoP.Text = "";
			this.lblFISL.Text = "";
			this.txtApellidoM.Text = "";
			this.txtPnombre.Text = "";
			this.txtSnombre.Text = "";
			this.cmbEntidad.Value=1;
			this.cmbBodega.Value=0;
			this.chkGarante.Checked = false;
			this.txtEstado.Value = 0;
			this.txtIdCre_Solicitud.Value = 0;
			this.txtClienteIdCre_DatosGenerales.Value = 0;
			this.txtClienteIdCre_DatosUbicacion.Value = 0;
			this.txtIdCre_Conyugue.Value = 0;
			this.txtIdGarante.Value = 0;
			this.txtConyugueIdCre_DatosGenerales.Value = 0;
			this.txtConyugueIdCre_DatosUbicacion.Value = 0;
			this.txtClienteIdCre_Negocio.Value = 0;
			this.txtNegocioIdCre_DatosUbicacion.Value = 0;
			this.txtClienteIdCre_EmpresasTrabajo.Value = 0;
			this.txtEmpresaIdCre_DatosUbicacion.Value = 0;

			this.cmbNacionalidad.Value = 54;
			this.dtFechaNacimiento.Value = System.DBNull.Value;
			this.cmbSexo.Value = 0;
			this.cmbProvinciaNAcimiento.Value = 0;
			this.cmbCantonNacimiento.Value = 0;
			this.cmbEstadoCivil.Value = 0;
			this.txtDependientes.Value = 0;
			this.cmbNivelEducacion.Value = 0;
			this.cmbProfesion.Value = 0;
			this.cmbSituacionLaboral.Value = 0;
			this.cmbActividadEconomica.Value = 0;
			//this.txtVehiculoMarca.Text ="";
			//this.txtModelo.Text ="";
			//this.txtValorVehiculo.Value = 0;
			this.chkSeparacionBienes.Checked = false;
			this.chkHipoteca.Checked = false;
			this.cmbDDProvincia.Value = 0;
			this.cmbDDCanton.Value = 0;
			this.cmbDDParroquia.Value = 0;
			this.cmbDDBarrio.Value = 0;
			this.txtDDEmail.Text ="";
			this.txtDDCallaPrincipal.Text ="";
			this.txtDDNumeroCasa.Text ="";
			this.txtDDCalleSecundaria.Text ="";
			this.txtDDReferenciasU.Text = "";
			this.txtDDTelefono1.Text ="";
			this.txtDDTelefono2.Text ="";
			this.txtDDCelular.Text = "";
			this.cmbDDTipoVivienda.Text = "";
			//this.dtTiempoVivienda.Text ="";
			this.cmbTiempoVivienda.Value = 0;
			this.txtNombreArrendador.Text ="";
			this.txtTelfArrendador.Text ="";
			this.txtCelArrendador.Text ="";
			

			this.cmbDDInmueble.Value =0;
			this.cmbDDCiudadInmueble.Value = 0;
			this.txtDDValorInmueble.Value = 0;

			this.cmbDPCTipodocumento.Value = 0;
			this.txtDPCdocumento.Text ="";
			this.txtDPCApellidoP.Text = "";
			this.txtDPCApellidoM.Text ="";
			this.txtDPCPrimerN.Text ="";
			this.txtDPCsegundoN.Text ="";
			this.cmbDPCNacionalidad.Value = 0;
			this.dtDPCFnacimiento.Value = System.DBNull.Value;
			this.cmbDPCSexo.Value = 0;
			this.cmbDPCProvinciaN.Value = 0;
			this.cmbDPCCantonN.Value = 0;
			this.cmbDPCNivelEducacion.Value = 0;
			this.cmbDPCProfesion.Value = 0;
			this.cmbDPCSituacionL.Value = 0;
			this.cmbDPCActividadE.Value = 0;
			this.txtDPCTNombreEmpresa.Text ="";
			this.cmbDPCTActividadEmpresa.Value = 0;
			this.cmbDPCTCargo.Value = 0;
			this.cmbDPCTProvincia.Value = 0;
			this.cmbDPCTCanton.Value = 0;
			this.cmbDPCTParroquia.Value = 0;
			this.cmbDPCTBarrio.Value = 0;
			this.txtDPCTEmail.Text = "";
			this.txtDPCTCallePrincipal.Text ="";
			this.txtDPCTNumeroCasa.Text ="";
			this.txtDPCTCalleSecundaria.Text ="";
			this.txtDPCTReferencia.Text ="";
			this.txtDPCTtelefono.Text="";
			this.txtDPCTExt.Text ="";
			this.txtDPCTCelularC.Text="";
			this.txtDPCTIngresos.Value = 0;
			this.txtDPCTGastos.Value = 0;
			//this.txtDPCTActivos.Value = 0;
			//this.txtDPCTPasivos.Value = 0;
			this.lblFechaActualizacion.Text = "";

//			this.ultgconsulta.DataSource = Funcion.dvProcedimiento(cdsSeteoF, string.Format("Exec Cre_ConsultaReferencias {0},{1}", 0,0));
			this.ultgconsulta.DataSource = Funcion.dvProcedimiento(cdsSeteoF, string.Format("Exec Cre_ConsultaReferencias 0,1"));

			this.txtNegocioRuc.Text = "";
			this.txtNombreDelNegocio.Text = "";
			//this.dtpNegocioFechaInicio.Value = System.DBNull.Value;
			this.txtNegocioMetros.Value = 1;
			this.txtNegocioEmpleados.Value = 0;
			this.cmbNegciotipo.Value = 0;
			this.cmbNEgocioActivdadEconomica.Value = 0;
			this.cmbNegocioProvincia.Value = 0;
			this.cmbNegocioCanton.Value = 0;
			this.cmbNegocioParroquia.Value = 0;
			this.cmbNegocioBarrio.Value = 0;
			this.txtNegocioemail.Text = "";
			this.txtNegocioCallePrincipal.Text ="";
			this.txtNegocioNumeroCasa.Text="";
			this.txtNegocioCalleSecundaria.Text = "";
			this.txtNegocioTelefono1.Text = "";
			this.txtNegocioTelefono2.Text = "";
			this.txtNEgocioCelular.Text = "";
			this.txtNegocioReferncia.Text = "";
			this.txtNEgocioIngresos.Value = 0;
			this.txtNEgocioGastos.Value = 0;
			//this.txtNegocioActivos.Value = 0;
			//this.txtNEgocioPasivos.Value = 0;
			this.chkNegocioObligadoaLlevar.Checked = false;
			this.chkAfiliadoTributario.Checked = false;

			this.txtDependienteRuc.Text ="";
			this.txtDependienteEmpresa.Text ="";
			this.cmbDependienteTipoempresa.Value = 0;
			this.dtDependienteFechaIngreso.Value = System.DBNull.Value; 
			this.txtDependienteIngresos.Value =0;
			this.txtDependienteGastos.Value =0;
			//this.txtDependienteActivos.Value =0;
			//this.txtDependientePasivos.Value=0;
			this.cmbDependienteTipocontrato.Value =0;
			this.cmbDependienteTipoSueldo.Value = 0;
			this.txtDependienteDepartamento.Text ="";
			this.cmbDependienteCargo.Value = 0;
			this.txtDependienteDiasPago.Value = 0;
			this.chkAfiliadoIESS.Checked = false;
			this.cmbDependienteProvincia.Value = 0;
			this.cmbDependienteCanton.Value = 0;
			this.cmbDependienteParroquia.Value = 0;
			this.cmbDependienteBarrio.Value = 0;
			this.txtDependienteemail.Text ="";
			this.txtDependienteCallePrinciapl.Text = "";
			this.txtDependienteNumeroCasa.Text = "";
			this.txtDependienteCalleSecundaria.Text = "";
			this.txtDependienteTelefono.Text= "";
			this.txtDependienteEXT.Text = "";
			this.txtDependienteCelular.Text= "";
			this.txtDependienteReferencia.Text = "";

			this.dtCreditoFechaSolicitud.Value = System.DBNull.Value; 
			this.txtCreditoScore.Value = 0;

			this.txtCreditoScoreInclusion.Value = 0;
			this.txtCreditoScoreSobreEndeudamiento.Value = 0;
			this.txtRangoIngresos.Text = "";
			this.txtCapacidadPago.Value = 0;
			this.txtTotalVencido.Value = 0;
			this.txtCarteraCastigada.Value = 0;
			this.txtSegmentacion.Text = "";
			this.txtResultado.Text = "";

			this.txtCreditoCuotaAsiganada.Value = 0;
			this.cmbCreditoCalificacion.Value = 0;
			this.cmbCreditoEstado.Value = 0;
			this.txtCreditoObservaciones.Value = 0;

			this.utgHistorialCredito.DataSource = Funcion.dvProcedimiento(cdsSeteoF, string.Format("Exec Cre_ConsultaHistorialCreditos 0 "));
			this.utgVerificacionTerrena.DataSource = Funcion.dvProcedimiento(cdsSeteoF, string.Format("Exec ConsultaVerificacionTerrena 0 "));

			#region Verificacion

   			 this.txtVDTiempoMeses.Value = 0;
				 this.txtVDPuntoReferencia.Text = "";
				 this.txtVDentrevistada.Text = "";
				this.txtVDObservacion.Text = "";
				 this.txtVDVecino.Text  = "";
				 this.txtVDGPS.Text = "";
				 this.txtVDLongitud.Text = "";
				 this.txtVDLatitud.Text = "";
				this.txtVDPrincipal.Text = "";
				 this.txtVDSecundaria.Text = "";
				 this.txtVDCCelArr.Value =  0;

				 this.opdtVTTipoTrabajo.Value = 0;
				 this.txtVTTiempoMeses.Value =  0;
				 this.txtVTTiempoYear.Value =  0;
				 this.txtVTIngresos.Value =  0;

				 this.txtVTActividadLaboral.Text = "";
				 this.txtVTTelefono.Text = "";
				this.txtVTReferencia.Text = "";
				 this.txtVTPersonaEntrevistada.Text = "";

			 this.txtVTGps.Text = "";
				 this.txtVTLatitud.Text = "";
				 this.txtVTLongitud.Text = "";

				 this.txtVTPRincipal.Text = "";
				 this.txtVTSecundaria.Text = "";

			#endregion Verificacion

			this.btnEquifax.Visible = false;

			


//		this.utgHistorialCredito.DataSource = Funcion.dvProcedimiento(cdsSeteoF, string.Format("Exec Cre_ConsultaHistorialCreditos '{0}'", ""));
      
			
		}

		private void cmbDependienteTipoempresa_KeyDown_1(object sender, System.Windows.Forms.KeyEventArgs e)
		{
				if(e.KeyCode == Keys.Enter) this.dtDependienteFechaIngreso.Focus();
		}

		private void cmbDPCActividadE_InitializeLayout(object sender, Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs e)
		{
		
		}

		private void cmbDPCSituacionL_ValueChanged(object sender, System.EventArgs e)
		{
			if (this.cmbEntidad.ActiveRow != null)
			{
				SituacionLaboralConyuge((int)this.cmbDPCSituacionL.Value);
			}
			
		}

		private void cmbActividadEconomica_InitializeLayout(object sender, Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs e)
		{
		}

		private void txtDDValorInmueble_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
		{
			this.txtDDValorInmueble.SelectAll ();
		}

		private void txtDDValorInmueble_Click(object sender, System.EventArgs e)
		{
				this.txtDDValorInmueble.SelectAll ();
		}

		private void txtValorVehiculo_Click(object sender, System.EventArgs e)
		{
		//this.txtValorVehiculo.SelectAll ();
		}

		private void txtDPCTIngresos_Click(object sender, System.EventArgs e)
		{
			this.txtDPCTIngresos.SelectAll ();
		}

		private void txtDPCTGastos_Click(object sender, System.EventArgs e)
		{
				this.txtDPCTGastos.SelectAll ();
		}

		private void txtDPCTActivos_Click(object sender, System.EventArgs e)
		{
		//this.txtDPCTActivos.SelectAll ();
		}

		private void txtDPCTPasivos_Click(object sender, System.EventArgs e)
		{
			//this.txtDPCTPasivos.SelectAll();
		}

		private void txtNEgocioIngresos_Click(object sender, System.EventArgs e)
		{
				this.txtNEgocioIngresos.SelectAll();
		}

		private void txtNEgocioGastos_Click(object sender, System.EventArgs e)
		{
				this.txtNEgocioGastos.SelectAll();
		}

		private void txtNegocioActivos_Click(object sender, System.EventArgs e)
		{
			//this.txtNegocioActivos.SelectAll ();
		}

		private void txtNEgocioPasivos_Click(object sender, System.EventArgs e)
		{
		//	this.txtNEgocioPasivos.SelectAll ();
		}

		private void txtNegocioMetros_Click(object sender, System.EventArgs e)
		{
					this.txtNegocioMetros.SelectAll ();
		}

		private void txtNegocioEmpleados_Click(object sender, System.EventArgs e)
		{
					this.txtNegocioEmpleados.SelectAll ();
		}

		private void txtCreditoScore_Click(object sender, System.EventArgs e)
		{
		 	this.txtCreditoScore.SelectAll ();
		}

		private void txtCreditoCuotaAsiganada_Click(object sender, System.EventArgs e)
		{
			this.txtCreditoCuotaAsiganada.SelectAll ();
		}

		private void txtDependienteIngresos_Click(object sender, System.EventArgs e)
		{
					this.txtDependienteIngresos.SelectAll ();
		}

		private void txtDependienteGastos_Click(object sender, System.EventArgs e)
		{
						this.txtDependienteGastos.SelectAll ();
		}

		private void txtDependienteActivos_Click(object sender, System.EventArgs e)
		{
				//this.txtDependienteActivos.SelectAll ();
		}

		private void txtDependientePasivos_Click(object sender, System.EventArgs e)
		{
			//this.txtDependientePasivos.SelectAll ();
		}

		private void txtDependienteDiasPago_Click(object sender, System.EventArgs e)
		{
				this.txtDependienteDiasPago.SelectAll ();
		}

		private void cmbActividadEconomica_KeyDown_1(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			//if(e.KeyCode == Keys.Enter) this.txtVehiculoMarca.Focus();
		}

		private void cmbDPCActividadE_KeyDown_2(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			if(e.KeyCode == Keys.Enter) this.txtDPCTNombreEmpresa.Focus();
		}

		private void btnGarante_Click(object sender, System.EventArgs e)
		{
			bool bGarante = true;
			if ((int) this.cmbEntidad.Value ==3)bGarante = false;
			if ((bool)this.chkGarante.Checked)
			{
       if ( (int) this.txtIdCre_Solicitud.Value <= 0) return;
//				int iGarante =  Funcion.iEscalarSQL(cdsSeteoF,string.Format("select COUNT (*) from Cre_Garantes where idCre_Solicitud = {0} ",(int)this.txtIdCre_Solicitud.Value));
				
//				if (iGarante == 0)
//				{
					using (frmCred_GaranteSolicitud consultaG = new frmCred_GaranteSolicitud(this.txtDocumento.Text, (int) this.txtIdCre_Solicitud.Value, bGarante, false, 0,false, 4))
					{
						if (DialogResult.OK == consultaG.ShowDialog())
						{
							this.txtIdGarante.Value = 1;
						}			
					}	

				//}
//				else
//				{
//					using (frmConsultaGarante consultab = new frmConsultaGarante(this.txtDocumento.Text, (int) this.txtIdCre_Solicitud.Value,bGarante, false, 4))
//					{                                                
//						if (DialogResult.OK == consultab.ShowDialog())
//						{
//         
//							//							idGaranteG =   (int)consultab.txtidGarante.Value;
//							//							IdCre_SolicitudG = (int)consultab.txtIdCre_Solicitud.Value; 
//							//							IdCre_DatosGeneralesG = (int)consultab.txtIdCre_DatosGenerales.Value; 
//							//							txtIdCre_DatosUbicacionG = (int)consultab.txtIdCre_DatosUbicacion.Value;
//							//							this.txtIdGarante.Value = (int)consultab.txtidGarante.Value;
//						}			
//					}
//				}
			}
		}

		private void chkGarante_CheckedChanged(object sender, System.EventArgs e)
		{
			if (this.chkGarante.Checked) this.btnGarante.Enabled = true; else this.btnGarante.Enabled = false;
		}

		private void cmbParentesco_Validating(object sender, System.ComponentModel.CancelEventArgs e)
		{
//			this.cmbProvincia.Update();
			this.ultraDropDown1.Update();
		}

		private void cmbParentesco_Validated(object sender, System.EventArgs e)
		{
//			this.cmbProvincia.Update();
			this.ultraDropDown1.Update();
		}

		private void ultraDataSource42_CellDataRequested(object sender, Infragistics.Win.UltraWinDataSource.CellDataRequestedEventArgs e)
		{
		
		}

		private void txtCreditoScoreSobreEndeudamiento_ValueChanged(object sender, System.EventArgs e)
		{
		
		}

		private void label127_Click(object sender, System.EventArgs e)
		{
		
		}

		private void groupBox5_Enter(object sender, System.EventArgs e)
		{
		
		}

		private void btnIdValidator_Click(object sender, System.EventArgs e)
		{
			int iIntentos = 0;
			int iEstadoValidacion = 0;


			#region Valida  
			string sSQLCredClte = string.Format("SELECT ISNULL(Intentos, 0), ISNULL(EstadoValidacion, 1) FROM CRE_DATOSGENERALES Where idCre_DatosGenerales = {0}", (int)this.txtClienteIdCre_DatosGenerales.Value);
			SqlDataReader drCredCliente = Funcion.rEscalarSQL(cdsSeteoF, sSQLCredClte, true);
			drCredCliente.Read();
			if (drCredCliente.HasRows)
			{
				iIntentos = (int)drCredCliente.GetInt32(0);
				iEstadoValidacion = (int)drCredCliente.GetInt32(1);
			}
			drCredCliente.Close();

			#endregion Valida  


			#region  Preguntas 
			using (frmCre_PreguntasValidacion Preguntas = new frmCre_PreguntasValidacion((int)this.txtClienteIdCre_DatosGenerales.Value,iIntentos,iEstadoValidacion))
			{
					if (DialogResult.OK == Preguntas.ShowDialog())
					{
						#region Actualiza Estados
						string sSQLidValidator = string.Format("Update Cre_DatosGenerales Set Intentos = ISNULL(Intentos, 0) + 1, EstadoValidacion = {0} Where idCre_DatosGenerales = {1}",Preguntas.iEstado,(int)this.txtClienteIdCre_DatosGenerales.Value);
						Funcion.EjecutaSQL(cdsSeteoF, sSQLidValidator, true);		
						#endregion Actualiza Estados		

					}
			}
			#endregion  Preguntas 
		}

		private void cmbSexo_InitializeLayout(object sender, Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs e)
		{
		
		}

		private void groupBox9_Enter(object sender, System.EventArgs e)
		{
		
		}

		private void cmbCreditoEstado_ValueChanged(object sender, System.EventArgs e)
		{
			if (this.cmbCreditoEstado.ActiveRow != null)
			{
				this.cmdEstadoSolicitud.DataSource = Funcion.dvProcedimiento(cdsSeteoF, string.Format("Exec Listas 24, {0}, 0", (int) this.cmbCreditoEstado.Value));
			}
		}
		private void imprimir ()
		{
			ParameterFields paramFields = new ParameterFields();
				
			ParameterField paramFieldCedula = new ParameterField ();
			ParameterDiscreteValue ValCedula = new ParameterDiscreteValue ();
			paramFieldCedula.ParameterFieldName = "@ruc";
			ValCedula.Value = this.txtDocumento.Text;
			paramFieldCedula.CurrentValues.Add (ValCedula);
			paramFields.Add (paramFieldCedula);


			Reporte miRep = new Reporte("ReporteDatosSolicitudes.rpt", "");
			miRep.MdiParent = this.MdiParent;
			miRep.Titulo("Datos cliente");
			miRep.crVista.ParameterFieldInfo = paramFields;
			miRep.Show();

		}

		private void btnImprimeVerificacion_Click(object sender, System.EventArgs e)
		{
			int iCont = 0;
			iCont = Funcion.iEscalarSQL(cdsSeteoF, string.Format("	select  count(1) from dbo.Cre_DatosGenerales c  Inner Join dbo.Cre_DatosUbicacion cu on cu.idOrigen = c.idOrigen Inner Join dbo.Cre_Provincia p on cu.idProvincia = p.idProvincia  Inner Join dbo.Cre_Canton ct on cu.idCanton = ct.idCanton Inner Join dbo.Cre_Parroquia cp on cu.idParroquia = cp.idParroquia  Inner Join dbo.Cre_Barrio v on v.idBarrio = cu.idBarrio  Inner Join dbo.Cre_ActividadEconomica ca on ca.idActEconomica = c.idCreActEconomicas where NumeroIdentificacion = '{0}' and cu.Origen = 0  and c.Origen = 0", this.txtDocumento.Text));
			if( iCont == 1)
			{
				imprimir ();
			}
		}

		private void txtDPCTGastos_ValueChanged(object sender, System.EventArgs e)
		{
		
		}

		private void btnComandos_Click(object sender, System.EventArgs e)
		{
			this.mnuAcciones.Show(this.btnComandos, new Point(1, 25));
		}

		private void mnuAccionesCopiarUltimoMes_Click(object sender, System.EventArgs e)
		{
		System.Diagnostics.Process.Start("chrome.exe", "https://storage.googleapis.com/sparta_bucket/Principal/DESAPOINT/1792605504001_000236550_1274424/QrDescaragNuovo.png");
		}

		private void mnuAccionesAgregarNuevoLocal_Click(object sender, System.EventArgs e)
		{
		System.Diagnostics.Process.Start("chrome.exe", "https://storage.googleapis.com/sparta_bucket/Principal/DESAPOINT/0916878820001_000001213_1279118/QR_NUOVO.PNG");
		}

		private void menuItem1_Click(object sender, System.EventArgs e)
		{
		System.Diagnostics.Process.Start("chrome.exe", "https://storage.googleapis.com/sparta_bucket/Principal/DESAPOINT/1204875023001_000001228_1279133/Qr_Huawei.png");
		}

		private void mnuAccionesActualizar_Click(object sender, System.EventArgs e)
		{
		System.Diagnostics.Process.Start("chrome.exe", "https://www.samsungknox.com/en/knox-platform/supported-devices");
		}

		private void txtValorVehiculo_ValueChanged(object sender, System.EventArgs e)
		{
		
		}

		private void txtNEgocioPasivos_ValueChanged(object sender, System.EventArgs e)
		{
		
		}

		private void txtCelArrendador_ValueChanged(object sender, System.EventArgs e)
		{
		
		}

		private void txtDocumento_ValueChanged(object sender, System.EventArgs e)
		{
		
		}

		private void txtDocumento_TextChanged(object sender, System.EventArgs e)
		{
		
		}

		private void btnVerificacion_Click(object sender, System.EventArgs e)
		{
			if (!Validacion.vbComboVacio(this.cmbBodega , "Seleccione la Bodega. ")) 
			{
				return;
			}
			if ((int)this.txtIdCre_Solicitud.Value == 0 ||(int) this.txtClienteIdCre_DatosGenerales.Value == 0 ) 
			{
				MessageBox.Show("Guarde la solicitud de Credito.", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				return;
			}

			if (!Validacion.vbTexto(this.txtDocumento, 1, 11, " Cedula ")) 
			{
				this.tbBasico.Tabs[0].Selected = true; 
				return;
			}
			/*int Existencia = Funcion.iEscalarSQL(cdsSeteoF,string.Format(" SELECT Count(1) idCoordenadasPrefactura  FROM CoordenadasPrefactura  WHERE cedula = '{0}' AND FechaSistema >= DATEADD(DAY, -2, GETDATE())",(string)this.txtDocumento.Text));
			if( Existencia == 0)
			{
				MessageBox.Show("Seleccione la ubicación del Cliente en el Mapa. \n Si ya lo genero intentelo en 2 minutos.", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
				return;
			}*/

			using (frmVerificacionTerrena miNota = new frmVerificacionTerrena((int)this.txtIdCre_Solicitud.Value, (int) this.txtClienteIdCre_DatosGenerales.Value, (string) this.txtDocumento.Text))
			{
				if (DialogResult.OK == miNota.ShowDialog())
				{	
					if(miNota.chkDomicilio.Checked )
					{
						#region Valida_DatosCliente
						if (!Validacion.vbTexto(this.txtDocumento, 1, 11, " Cedula ")) 
						{
							this.tbBasico.Tabs[0].Selected = true; 
							return;
						}
						if (!Validacion.vbTexto(this.txtApellidoP, 1, 60, " Paterno")) 
						{
							this.tbBasico.Tabs[0].Selected = true; 
							return;
						}
						if (!Validacion.vbTexto(this.txtPnombre,1, 60, " Nombre"))
						{
							this.tbBasico.Tabs[0].Selected = true; 
							return;
						}
						if (!Validacion.vbComboVacio(this.cmbBodega, "Seleccione el Local ")) 
						{
							this.tbBasico.Tabs[0].Selected = true;
							return;
						}

						if (!Validacion.vbComboVacio(this.cmbDDProvincia, "Seleccione la Provincia del Domicilio"))
						{
							this.tbBasico.Tabs[0].Selected = true; 
							return;
						}
						if (!Validacion.vbComboVacio(this.cmbDDCanton, "Seleccione el Canton del Domicilio ")) 
						{
							this.tbBasico.Tabs[0].Selected = true; 
							return;
						}
						if (!Validacion.vbComboVacio(this.cmbDDParroquia, "Seleccione la Parroquia del Domicilio ")) 
						{
							this.tbBasico.Tabs[0].Selected = true; 
							return;
						}
						if (!Validacion.vbComboVacio(this.cmbDDBarrio, "Seleccione el Barrio del Domicilio "))
						{
							this.tbBasico.Tabs[0].Selected = true; 
							return;
						}
						if (!Validacion.vbTexto(this.txtDDCallaPrincipal, 1, 100, "Ingrese la Calle Principal")) 
						{
							this.tbBasico.Tabs[0].Selected = true;
							return;
						}
						if (!Validacion.vbTexto(this.txtDDNumeroCasa, 1, 60, "Ingrese el Numero de Casa"))
						{
							this.tbBasico.Tabs[0].Selected = true;
							return;
						}
						if (!Validacion.vbTexto(this.txtDDCalleSecundaria, 1, 100, "Ingrese la Calle Secundaria")) 
						{
							this.tbBasico.Tabs[0].Selected = true;
							return;
						}
						if ( this.txtDDCelular.Text.ToString().Trim().Length == 0 )
						{
							MessageBox.Show(string.Format("Debe el Telefono Celular"), "Point Technology", MessageBoxButtons.OK,MessageBoxIcon.Stop);
							this.txtDDTelefono1.Focus();
							this.tbBasico.Tabs[0].Selected = true;
							return;	
						}
						if (!Validacion.vbValidaTelefono(this.txtDDCelular, 2, false)) 
						{
							this.tbBasico.Tabs[0].Selected = true;
							return;
						}
						if (!Validacion.vbTexto(this.txtDDReferenciasU, 10, 1000, "Ingrese la Referencia de Ubicación"))
						{ 
							this.tbBasico.Tabs[0].Selected = true; 
							return;
						}
						#endregion Valida_DatosCliente
					}
					if(miNota.chkTrabajo.Checked )
					{
						if (!Validacion.vbComboVacio(this.cmbSituacionLaboral , "Seleccione la Situacion LAboral ")) 
						{
							this.tbBasico.Tabs[0].Selected = true; 
							return;
						}
						if((int)this.cmbSituacionLaboral.Value == 1)
						{
							#region Valida_SeccionDependiente

							if (!Validacion.vbTexto(this.txtDependienteEmpresa, 1, 100, "Ingrese el Nombre de la Empresa"))
							{
								this.tbBasico.Tabs[4].Selected = true; 
								this.txtDependienteEmpresa.Focus();
								return;
							}

							if (!Validacion.vbComboVacio(this.cmbDependienteProvincia, "Seleccione la Provincia de su Trabajo ")) 
							{
								this.tbBasico.Tabs[4].Selected = true; 
								this.cmbDependienteProvincia.Focus();
								return;
							}
							if (!Validacion.vbComboVacio(this.cmbDependienteCanton, "Seleccione el Canton de su Trabajo ")) 
							{
								this.tbBasico.Tabs[4].Selected = true; 
								this.cmbDependienteCanton.Focus();
								return;
							}
							if (!Validacion.vbComboVacio(this.cmbDependienteParroquia, "Seleccione la Parroquia de su Trabajo "))
							{
								this.tbBasico.Tabs[4].Selected = true; 
								this.cmbDependienteParroquia.Focus();
								return;
							}
							if (!Validacion.vbComboVacio(this.cmbDependienteBarrio, "Seleccione el Barrio de su Trabajo ")) 
							{
								this.tbBasico.Tabs[4].Selected = true;
								this.cmbDependienteBarrio.Focus();
								return;
							}
							if (!Validacion.vbTexto(this.txtDependienteCallePrinciapl, 3, 100, "Ingrese la Calle Principal del Trabajo"))
							{
								this.tbBasico.Tabs[4].Selected = true;
								this.txtDependienteCallePrinciapl.Focus();
								return;
							}
							if (!Validacion.vbTexto(this.txtDependienteNumeroCasa, 3, 60, "Ingrese el Numero de Casa del Trabajo")) 
							{
								this.tbBasico.Tabs[4].Selected = true;
								this.txtDependienteNumeroCasa.Focus();
								return;
							}
							if (!Validacion.vbTexto(this.txtDependienteCalleSecundaria, 3, 100, "Ingrese la Calle Secundaria del Trabajo"))
							{
								this.tbBasico.Tabs[4].Selected = true; 
								this.txtDependienteCalleSecundaria.Focus();
								return;
							}
							if (this.txtDependienteTelefono.Text.ToString().Trim().Length == 0 &&  this.txtDependienteCelular.Text.ToString().Trim().Length == 0 )
							{
								MessageBox.Show(string.Format("Debe Ingresar el Telefono Convencional o el Telefono Celular del Trabajo "), "Point Technology", MessageBoxButtons.OK,MessageBoxIcon.Stop);
		
								this.tbBasico.Tabs[4].Selected = true;
								this.txtDependienteTelefono.Focus();
								return;	
							}
							if (!Validacion.vbValidaTelefono(this.txtDependienteTelefono, 1, false))
							{
								this.tbBasico.Tabs[4].Selected = true; 
								this.txtDependienteTelefono.Focus();
								return;
							}
							if (!Validacion.vbValidaTelefono(this.txtDependienteCelular, 2, false))
							{
								this.tbBasico.Tabs[4].Selected = true;
								this.txtDependienteCelular.Focus();
								return;
							}
							if (!Validacion.vbTexto(this.txtDependienteReferencia, 10, 300, "Ingrese la Referencia de la Ubicación del Trabajo")) 
							{
								this.tbBasico.Tabs[4].Selected = true; 
								this.txtDependienteReferencia.Focus();
								return;
							}

							#endregion Valida_SeccionDependiente
						}
						else
						{
							#region Valida_SeccionNEgocio

							if (!Validacion.vbTexto(this.txtNombreDelNegocio, 3, 200, "Ingrese el Nombre del Negocio")) 
							{
								this.tbBasico.Tabs[3].Selected = true;
								this.txtNombreDelNegocio.Focus();
								return;
							}

							if (!Validacion.vbTexto(this.txtNegocioCallePrincipal, 1, 100, "Ingrese la Calle Principal del Negocio")) 
							{
								this.tbBasico.Tabs[3].Selected = true;
								this.txtNegocioCallePrincipal.Focus();
								return;
							}
							if (!Validacion.vbTexto(this.txtNegocioNumeroCasa, 1, 60, "Ingrese el Numero de Casa del Negocio"))
							{
								this.tbBasico.Tabs[3].Selected = true;
								this.txtNegocioNumeroCasa.Focus();
								return;
							}
							if (!Validacion.vbTexto(this.txtNegocioCalleSecundaria, 3, 100, "Ingrese la Calle Secundaria del Negocio"))
							{
								this.tbBasico.Tabs[3].Selected = true;
								this.txtNegocioCalleSecundaria.Focus();
								return;
							}


							if (!Validacion.vbComboVacio(this.cmbNegocioProvincia  , "Seleccione la Provincia del Negocio")) 
							{
								this.tbBasico.Tabs[3].Selected = true;
								this.cmbNegocioProvincia.Focus();
								return;
							}

							if (!Validacion.vbComboVacio(this.cmbNegocioCanton , "Seleccione el Canton del Negocio")) 
							{
								this.tbBasico.Tabs[3].Selected = true;
								this.cmbNegocioCanton.Focus();
								return;
							}
							if (!Validacion.vbComboVacio(this.cmbNegocioParroquia , "Seleccione la Parroquia del Negocio"))
							{
								this.tbBasico.Tabs[3].Selected = true;
								this.cmbNegocioParroquia.Focus();
								return;
							}
							if (!Validacion.vbComboVacio(this.cmbNegocioBarrio, "Seleccione el Barrio del Negocio"))
							{
								this.tbBasico.Tabs[3].Selected = true; 
								this.cmbNegocioBarrio.Focus();
								return;
							}
				
							if (this.txtNegocioTelefono1.Text.ToString().Trim().Length == 0 &&  this.txtNEgocioCelular.Text.ToString().Trim().Length == 0 )
							{
								MessageBox.Show(string.Format("Debe Ingresar el Telefono Convencional o el Telefono Celular del Negocio "), "Point Technology", MessageBoxButtons.OK,MessageBoxIcon.Stop);
				
								this.tbBasico.Tabs[3].Selected = true;
								this.txtNegocioTelefono1.Focus();
								return;	
							}

							if (!Validacion.vbValidaTelefono(this.txtNegocioTelefono2, 1, false)) 
							{
								this.tbBasico.Tabs[3].Selected = true; 
								this.txtNegocioTelefono2.Focus();
								return;
							}
							if (!Validacion.vbValidaTelefono(this.txtNEgocioCelular, 2, false)) 
							{
								this.tbBasico.Tabs[3].Selected = true; 
								this.txtNEgocioCelular.Focus();
								return;
							}
							if (!Validacion.vbTexto(this.txtNegocioReferncia, 10, 300, "Ingrese la Referencia de la Ubicación del Negocio"))
							{
								this.tbBasico.Tabs[3].Selected = true; 
								this.txtNegocioReferncia.Focus();
								return;
							}
							#endregion Valida_SeccionNEgocio
						}
					}
					string sNumero ="";
					#region Crea Verificacion
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
							int idProvincia = 0;
							int idCanton = 0;
							int idParroquia = 0;
							int idBarrio = 0;

							if (this.cmbDDProvincia.ActiveRow != null) idProvincia = (int)this.cmbDDProvincia.Value;
							if (this.cmbDDCanton.ActiveRow != null) idCanton = (int)this.cmbDDCanton.Value;
							if (this.cmbDDParroquia.ActiveRow != null) idParroquia = (int)this.cmbDDParroquia.Value;
							if (this.cmbDDBarrio.ActiveRow != null) idBarrio = (int)this.cmbDDBarrio.Value;

							string empresa ="";
							int idProvinciaTrabajo = 0;
							int idCantonTrabajo = 0;
							int idParroquiaTrabajo = 0;
							int idBarrioTrabajo  = 0;
							string CalleprincipalTrabajo ="";
							string NumeracionTrabajo ="";
							string CalleSecundariaTrabajo ="";
							string ReferenciaTrabajo ="";
							string CelularReferenciaTrabajo ="";

							if((int)this.cmbSituacionLaboral.Value == 1)
							{
								empresa =	this.txtDependienteEmpresa.Text;
								if (this.cmbDependienteProvincia.ActiveRow != null) idProvinciaTrabajo = (int)this.cmbDependienteProvincia.Value;
								if (this.cmbDependienteCanton.ActiveRow != null) idCantonTrabajo = (int)this.cmbDependienteCanton.Value;
								if (this.cmbDependienteParroquia.ActiveRow != null) idParroquiaTrabajo = (int)this.cmbDependienteParroquia.Value;
								if (this.cmbDependienteBarrio.ActiveRow != null) idBarrioTrabajo = (int)this.cmbDependienteBarrio.Value;
								CalleprincipalTrabajo=	this.txtDependienteCallePrinciapl.Text;
								NumeracionTrabajo=	this.txtDependienteNumeroCasa.Text;
								CalleSecundariaTrabajo =	this.txtDependienteCalleSecundaria.Text;
								ReferenciaTrabajo =	this.txtDependienteReferencia.Text;
								CelularReferenciaTrabajo=	this.txtDependienteCelular.Text;
							}
							else
							{
								empresa =	this.txtNombreDelNegocio.Text;
								if (this.cmbNegocioProvincia.ActiveRow != null) idProvinciaTrabajo = (int)this.cmbNegocioProvincia.Value;
								if (this.cmbNegocioCanton.ActiveRow != null) idCantonTrabajo = (int)this.cmbNegocioCanton.Value;
								if (this.cmbNegocioParroquia.ActiveRow != null) idParroquiaTrabajo = (int)this.cmbNegocioParroquia.Value;
								if (this.cmbNegocioBarrio.ActiveRow != null) idBarrioTrabajo = (int)this.cmbNegocioBarrio.Value;
								CalleprincipalTrabajo=	this.txtNegocioCallePrincipal.Text;
								NumeracionTrabajo=	this.txtNegocioNumeroCasa.Text;
								CalleSecundariaTrabajo =	this.txtNegocioCalleSecundaria.Text;
								ReferenciaTrabajo =	this.txtNegocioReferncia.Text;
								CelularReferenciaTrabajo=	this.txtNEgocioCelular.Text;
							}



							#region Numeracion

							string ssSQL = string.Format("Exec NumeracionLocales 95, 0, 0");
							oCmdActualiza.CommandText = ssSQL;
							sNumero = oCmdActualiza.ExecuteScalar().ToString();						
							#endregion Numeracion

							#region Graba Ingreso Gastos

							string  sSQLDeta = string.Format("EXEC  CreateVerificacionTerrena {0},{1},'{2}','{3}','{4}',{5},{6},{7},'{8}',{9},{10},'{11}','{12}',{13},{14},{15},{16},'{17}','{18}','{19}','{20}','{21}','{22}',{23},{24},{25},{26},'{27}','{28}','{29}','{30}','{31}', '{32}'" ,
								(int)this.txtIdCre_Solicitud.Value, 
								(int) this.txtClienteIdCre_DatosGenerales.Value,
								this.txtDocumento.Text, 
								(this.txtApellidoP.Text + " "+ this.txtApellidoM.Text + " " + this.txtPnombre.Text + " " + this.txtSnombre.Text ),
								this.txtDDCelular.Text, miNota.cmbCobrador.Value, 0, (int) this.cmbBodega.Value, "", miNota.chkDomicilio.Checked,
								miNota.chkTrabajo.Checked, "", "",idProvincia, idCanton, idParroquia, idBarrio, this.txtDDCallaPrincipal.Text,
								this.txtDDNumeroCasa.Text, this.txtDDCalleSecundaria.Text, this.txtDDReferenciasU.Text, this.txtDDCelular.Text,
								empresa ,idProvinciaTrabajo , idCantonTrabajo , idParroquiaTrabajo ,idBarrioTrabajo  , CalleprincipalTrabajo ,
								NumeracionTrabajo , CalleSecundariaTrabajo , ReferenciaTrabajo, CelularReferenciaTrabajo , sNumero
								);
							oCmdActualiza.CommandText = sSQLDeta;
							oCmdActualiza.ExecuteNonQuery();	

							#endregion Graba Ingreso Gastos
          
							oTransaction.Commit();

							Cursor = Cursors.Default;	

							MessageBox.Show("Verificación Guardada Correctamente", "Información", MessageBoxButtons.OK, MessageBoxIcon.Information);

								
						}
						catch(Exception ex)
						{
							try
							{ 
								oTransaction.Rollback();

								MessageBox.Show(string.Format("Commit Exception Type : {0} {1}", ex.GetType(), ex.Message), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Error);
							}
							catch (Exception ex2)
							{
								MessageBox.Show(string.Format("Rollback Exception Type: {0} {1}", ex2.GetType(), ex2.Message), "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Error);							
							}
						}
						finally
						{
							oConexion.Close();
						}
					}
				
					#endregion Guardar
					#endregion Crea Verificacion
				}
			}
		}

		private void btnUbicacion_Click(object sender, System.EventArgs e)
		{
			#region Valida_DatosCliente
			if (!Validacion.vbTexto(this.txtDocumento, 1, 11, " Cedula ")) 
			{
				this.tbBasico.Tabs[0].Selected = true; 
				return;
			}
			if (!Validacion.vbTexto(this.txtApellidoP, 1, 60, " Paterno")) 
			{
				this.tbBasico.Tabs[0].Selected = true; 
				return;
			}
			if (!Validacion.vbTexto(this.txtPnombre,1, 60, " Nombre"))
			{
				this.tbBasico.Tabs[0].Selected = true; 
				return;
			}
			if (!Validacion.vbComboVacio(this.cmbDDProvincia, "Seleccione la Provincia del Domicilio"))
			{
				this.tbBasico.Tabs[0].Selected = true; 
				return;
			}
			
			if (!Validacion.vbComboVacio(this.cmbDDCanton, "Seleccione el Canton del Domicilio ")) 
			{
				this.tbBasico.Tabs[0].Selected = true; 
				return;
			}
			if (!Validacion.vbComboVacio(this.cmbDDParroquia, "Seleccione la Parroquia del Domicilio ")) 
			{
				this.tbBasico.Tabs[0].Selected = true; 
				return;
			}
			if (!Validacion.vbComboVacio(this.cmbDDBarrio, "Seleccione el Barrio del Domicilio "))
			{
				this.tbBasico.Tabs[0].Selected = true; 
				return;
			}
			if (!Validacion.vbTexto(this.txtDDCallaPrincipal, 1, 100, "Ingrese la Calle Principal")) 
			{
				this.tbBasico.Tabs[0].Selected = true;
				return;
			}
			if (!Validacion.vbTexto(this.txtDDNumeroCasa, 1, 60, "Ingrese el Numero de Casa"))
			{
				this.tbBasico.Tabs[0].Selected = true;
				return;
			}
			if (!Validacion.vbTexto(this.txtDDCalleSecundaria, 1, 100, "Ingrese la Calle Secundaria")) 
			{
				this.tbBasico.Tabs[0].Selected = true;
				return;
			}
			#endregion Valida_DatosCliente

			string Snombre = (this.txtPnombre.Text + " " +this.txtSnombre.Text + " " + this.txtApellidoP.Text + " " + this.txtApellidoM.Text );
			string sUrlNube = Funcion.sEscalarSQL(cdsSeteoF,string.Format("EXEC [RetornaURLCoordanadas] {0},{1},{2},{3},{4},'{5}', {6}",(string) this.txtDocumento.Text, (int) this.cmbDDProvincia.Value, (int) this.cmbDDCanton.Value, (int) this.cmbDDParroquia.Value, (int) this.cmbDDBarrio.Value, Snombre, 1)); 

			if (sUrlNube.Length > 10)
			{
				System.Diagnostics.Process.Start("chrome.exe", sUrlNube);				
			}
		}

		private void utgVerificacionTerrena_ClickCellButton(object sender, Infragistics.Win.UltraWinGrid.CellEventArgs e)
		{
			#region Anular
			if (e.Cell.Column.ToString() == "Anular")
			{	
				if( Convert.ToInt32( e.Cell.Row.Cells["iEstado"].Value ) == 3 )
				{
					MessageBox.Show("La Verificación esta Anulada.", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
					return;
				}
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

								string sSQLAutEs = string.Format("Update ClientesVerificionTerrena set iEstado = 3, NotaAnula ='{1}', FechaAnula = GETDATE (), UsuarioAnula =  SUSER_NAME() where idClienteVerificacion = {0}",
									(int) e.Cell.Row.Cells["idClienteVerificacion"].Value,miNota.txtNotas.Text);
								oCmdActualiza.CommandText = sSQLAutEs;
								oCmdActualiza.ExecuteNonQuery();					
						
		
								oTransaction.Commit();
								Cursor = Cursors.Default;	
								MessageBox.Show("Registro Anulado Correctamente", "Información", MessageBoxButtons.OK, MessageBoxIcon.Information);

								string sSQLOSolVeri = string.Format("Exec ConsultaVerificacionTerrena {0}", (int)this.txtIdCre_Solicitud.Value);
								this.utgVerificacionTerrena.DataSource = FuncionesProcedimientos.dtGeneral(sSQLOSolVeri);
				
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
			#endregion Anular

			#region Aprobar
			if (e.Cell.Column.ToString() == "Aprobar")
			{	
				if( Convert.ToInt32( e.Cell.Row.Cells["iEstado"].Value ) == 3 )
				{
					MessageBox.Show("La Verificación esta Anulada.", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
					return;
				}
				if( Convert.ToInt32( e.Cell.Row.Cells["iEstado"].Value ) == 2 )
				{
					MessageBox.Show("La Verificación esta Aprobada.", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
					return;
				}
				if( Convert.ToInt32( e.Cell.Row.Cells["iEstado"].Value ) == 0 )
				{
					MessageBox.Show("La Verificación esta pendiente de enviar por parte del Verificador.", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
					return;
				}
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

								string sSQLAutEs = string.Format("Update ClientesVerificionTerrena set iEstado = 2, NotaAprueba ='{1}', FechaAprueba = GETDATE (), UsuarioAprueba =  SUSER_NAME() where idClienteVerificacion = {0}",
									(int) e.Cell.Row.Cells["idClienteVerificacion"].Value,miNota.txtNotas.Text);
								oCmdActualiza.CommandText = sSQLAutEs;
								oCmdActualiza.ExecuteNonQuery();					
						
		
								oTransaction.Commit();
								Cursor = Cursors.Default;	
								MessageBox.Show("Registro Aprobado Correctamente", "Información", MessageBoxButtons.OK, MessageBoxIcon.Information);

								string sSQLOSolVeri = string.Format("Exec ConsultaVerificacionTerrena {0}", (int)this.txtIdCre_Solicitud.Value);
								this.utgVerificacionTerrena.DataSource = FuncionesProcedimientos.dtGeneral(sSQLOSolVeri);
				
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
			#endregion Aprobar

			#region Cobrador
			if (e.Cell.Column.ToString() == "Cobrador")
			{	
				if( Convert.ToInt32( e.Cell.Row.Cells["iEstado"].Value ) == 3 )
				{
					MessageBox.Show("La Verificación esta Anulada.", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
					return;
				}
				if( Convert.ToInt32( e.Cell.Row.Cells["iEstado"].Value ) == 2 )
				{
					MessageBox.Show("La Verificación esta Aprobada.", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
					return;
				}
				if( Convert.ToInt32( e.Cell.Row.Cells["iEstado"].Value ) == 1 )
				{
					MessageBox.Show("La Verificación esta Enviada", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
					return;
				}
				using (frmCambioVerificador miNota = new frmCambioVerificador())
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
								string sSQLAutEs = string.Format("Update ClientesVerificionTerrena set idVerificador = {1} where idClienteVerificacion = {0}",
									(int) e.Cell.Row.Cells["idClienteVerificacion"].Value, miNota.cmbCobrador.Value);
								oCmdActualiza.CommandText = sSQLAutEs;
								oCmdActualiza.ExecuteNonQuery();					
								oTransaction.Commit();
								Cursor = Cursors.Default;	
								MessageBox.Show("El Verificador fue Cambiado.", "Información", MessageBoxButtons.OK, MessageBoxIcon.Information);
								string sSQLOSolVeri = string.Format("Exec ConsultaVerificacionTerrena {0}", (int)this.txtIdCre_Solicitud.Value);
								this.utgVerificacionTerrena.DataSource = FuncionesProcedimientos.dtGeneral(sSQLOSolVeri);
				
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
				

			}
			#endregion Cobrador

			#region Documento
			if (e.Cell.Column.ToString() == "Documento")
			{
				if(Convert.ToString(e.Cell.Row.Cells["UrlGoogle"].Value).Length <=5)
				{
					MessageBox.Show("La Verificación esta Pendiente de su Verificación.", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
					
					return;
				}

				string sUrlNube = (string) e.Cell.Row.Cells["UrlGoogle"].Value;
				if (sUrlNube.Length > 10 )
				{
					Process.Start(sUrlNube);
					System.Diagnostics.Process.Start("chrome.exe", sUrlNube);
				}

			}
			#endregion Documento


			#region ImagenTrabajo
			if (e.Cell.Column.ToString() == "Trabajo")
			{
				if( Convert.ToInt32( e.Cell.Row.Cells["idTerrenaGestionTrabajo"].Value ) == 0 )
				{
					MessageBox.Show("No se ha cargo Imagenes", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
					return;
				}
				using (frmImagenesVerificacion miNota = new frmImagenesVerificacion(2,Convert.ToInt32( e.Cell.Row.Cells["idTerrenaGestionTrabajo"].Value ) ))
				{
					if (DialogResult.OK == miNota.ShowDialog())
					{	
						
					}
				}
			}
			#endregion ImagenTrabajo

			#region ImagenDomicilio
			if (e.Cell.Column.ToString() == "Domicilio")
			{
				if( Convert.ToInt32( e.Cell.Row.Cells["idTerrenaGestionDomicilio"].Value ) == 0 )
				{
					MessageBox.Show("No se ha cargo Imagenes", "Point Technology", MessageBoxButtons.OK, MessageBoxIcon.Stop);
					return;
				}
				using (frmImagenesVerificacion miNota = new frmImagenesVerificacion(1,Convert.ToInt32( e.Cell.Row.Cells["idTerrenaGestionDomicilio"].Value )))
				{
					if (DialogResult.OK == miNota.ShowDialog())
					{	
						
					}
				}
			}
			#endregion ImagenDomicilio
		}

		private void ultraTabControl1_SelectedTabChanged(object sender, Infragistics.Win.UltraWinTabControl.SelectedTabChangedEventArgs e)
		{
		
		}

		private void utgVerificacionTerrena_BeforeSelectChange(object sender, Infragistics.Win.UltraWinGrid.BeforeSelectChangeEventArgs e)
		{
			if (this.utgVerificacionTerrena.ActiveRow == null) return;
			if (this.utgVerificacionTerrena.ActiveRow.Cells["idClienteVerificacion"].Value == DBNull.Value) return;
				#region Domicilio
				string sSQLDomi = string.Format("Exec [ConsultaVerificacionTerrenaDomicilio] {0}, {1}", 
					(int) this.utgVerificacionTerrena.ActiveRow.Cells["idClienteVerificacion"].Value ,
					(int) this.utgVerificacionTerrena.ActiveRow.Cells["idTerrenaGestionDomicilio"].Value 
					);
				SqlDataReader dr = Funcion.rEscalarSQL(cdsSeteoF, sSQLDomi, true);
				dr.Read();
				if (dr.HasRows)
				{

					if (dr.GetValue(2) != System.DBNull.Value) this.txtVDTiempoMeses.Value = dr.GetInt32(2);
					if (dr.GetValue(3) != System.DBNull.Value) this.opdtVDVivienda.Value = dr.GetInt32(3);
					if (dr.GetValue(4) != System.DBNull.Value) this.opdtVDEstadoVivienda.Value = dr.GetInt32(4);
					if (dr.GetValue(5) != System.DBNull.Value) this.opdtVDZonaVivienda.Value = dr.GetInt32(5);
					if (dr.GetValue(6) != System.DBNull.Value) this.opdtVDPropiedad.Value = dr.GetInt32(6);
					if (dr.GetValue(7) != System.DBNull.Value) this.opdtVDAcceso.Value = dr.GetInt32(7); 
					if (dr.GetValue(8) != System.DBNull.Value) this.opdtVDCobertura.Value = dr.GetInt32(8);

					if (dr.GetValue(9) != System.DBNull.Value) this.txtVDPuntoReferencia.Text = dr.GetString(9);
					if (dr.GetValue(10) != System.DBNull.Value) this.txtVDentrevistada.Text = dr.GetString(10);
					if (dr.GetValue(11) != System.DBNull.Value) this.txtVDObservacion.Text = dr.GetString(11);
					if (dr.GetValue(12) != System.DBNull.Value) this.txtVDVecino.Text  = dr.GetString(12);
					if (dr.GetValue(13) != System.DBNull.Value) this.txtVDGPS.Text = dr.GetString(13);
					if (dr.GetValue(14) != System.DBNull.Value) this.txtVDLongitud.Text = dr.GetString(14);
					if (dr.GetValue(15) != System.DBNull.Value) this.txtVDLatitud.Text = dr.GetString(15);
					if (dr.GetValue(16) != System.DBNull.Value) this.txtVDPrincipal.Text = dr.GetString(16);
					if (dr.GetValue(17) != System.DBNull.Value) this.txtVDSecundaria.Text = dr.GetString(17);
					if (dr.GetValue(18) != System.DBNull.Value) this.txtVDCCelArr.Value = dr.GetDecimal(18); 
				}
				dr.Close();
				#endregion Domicilio

			#region trabajo
			string sSQLDomiTra = string.Format("Exec [ConsultaVerificacionTerrenaTrabajo] {0}, {1}", 
				(int) this.utgVerificacionTerrena.ActiveRow.Cells["idClienteVerificacion"].Value ,
				(int) this.utgVerificacionTerrena.ActiveRow.Cells["idTerrenaGestionTrabajo"].Value 
				);
			SqlDataReader drt = Funcion.rEscalarSQL(cdsSeteoF, sSQLDomiTra, true);
			drt.Read();
			if (drt.HasRows)
			{

				if (drt.GetValue(2) != System.DBNull.Value) this.opdtVTTipoTrabajo.Value = drt.GetInt32(2);
				if (drt.GetValue(3) != System.DBNull.Value) this.txtVTTiempoMeses.Value = drt.GetInt32(3);
				if (drt.GetValue(4) != System.DBNull.Value) this.txtVTTiempoYear.Value = drt.GetInt32(4);
				if (drt.GetValue(5) != System.DBNull.Value) this.txtVTIngresos.Value = drt.GetDecimal(5); 

				if (drt.GetValue(6) != System.DBNull.Value) this.txtVTActividadLaboral.Text = drt.GetString(6);
				if (drt.GetValue(7) != System.DBNull.Value) this.txtVTTelefono.Text = drt.GetString(7);
				if (drt.GetValue(8) != System.DBNull.Value) this.txtVTReferencia.Text = drt.GetString(8);
				if (drt.GetValue(9) != System.DBNull.Value) this.txtVTPersonaEntrevistada.Text = drt.GetString(9);

				if (drt.GetValue(10) != System.DBNull.Value) this.txtVTGps.Text = drt.GetString(10);
				if (drt.GetValue(11) != System.DBNull.Value) this.txtVTLatitud.Text = drt.GetString(11);
				if (drt.GetValue(12) != System.DBNull.Value) this.txtVTLongitud.Text = drt.GetString(12);

				if (drt.GetValue(14) != System.DBNull.Value) this.txtVTPRincipal.Text = drt.GetString(14);
				if (drt.GetValue(15) != System.DBNull.Value) this.txtVTSecundaria.Text = drt.GetString(15);
			}
			drt.Close();
			#endregion trabajo
		}

		private void button1_Click(object sender, System.EventArgs e)
		{
			#region Valida_DatosCliente
			if (!Validacion.vbTexto(this.txtDocumento, 1, 11, " Cedula ")) 
			{
				this.tbBasico.Tabs[0].Selected = true; 
				return;
			}
			if (!Validacion.vbTexto(this.txtApellidoP, 1, 60, " Paterno")) 
			{
				this.tbBasico.Tabs[0].Selected = true; 
				return;
			}
			if (!Validacion.vbTexto(this.txtPnombre,1, 60, " Nombre"))
			{
				this.tbBasico.Tabs[0].Selected = true; 
				return;
			}
			if (!Validacion.vbComboVacio(this.cmbDDProvincia, "Seleccione la Provincia del Domicilio"))
			{
				this.tbBasico.Tabs[0].Selected = true; 
				return;
			}
			
			if (!Validacion.vbComboVacio(this.cmbDDCanton, "Seleccione el Canton del Domicilio ")) 
			{
				this.tbBasico.Tabs[0].Selected = true; 
				return;
			}
			if (!Validacion.vbComboVacio(this.cmbDDParroquia, "Seleccione la Parroquia del Domicilio ")) 
			{
				this.tbBasico.Tabs[0].Selected = true; 
				return;
			}
			if (!Validacion.vbComboVacio(this.cmbDDBarrio, "Seleccione el Barrio del Domicilio "))
			{
				this.tbBasico.Tabs[0].Selected = true; 
				return;
			}
			if (!Validacion.vbTexto(this.txtDDCallaPrincipal, 1, 100, "Ingrese la Calle Principal")) 
			{
				this.tbBasico.Tabs[0].Selected = true;
				return;
			}
			if (!Validacion.vbTexto(this.txtDDNumeroCasa, 1, 60, "Ingrese el Numero de Casa"))
			{
				this.tbBasico.Tabs[0].Selected = true;
				return;
			}
			if (!Validacion.vbTexto(this.txtDDCalleSecundaria, 1, 100, "Ingrese la Calle Secundaria")) 
			{
				this.tbBasico.Tabs[0].Selected = true;
				return;
			}
			#endregion Valida_DatosCliente

			string Snombre = (this.txtPnombre.Text + " " +this.txtSnombre.Text + " " + this.txtApellidoP.Text + " " + this.txtApellidoM.Text );
			string sUrlNube = Funcion.sEscalarSQL(cdsSeteoF,string.Format("EXEC [RetornaURLCoordanadas] {0},{1},{2},{3},{4},'{5}', {6}",(string) this.txtDocumento.Text, (int) this.cmbDDProvincia.Value, (int) this.cmbDDCanton.Value, (int) this.cmbDDParroquia.Value, (int) this.cmbDDBarrio.Value, Snombre, 2)); 

			if (sUrlNube.Length > 10)
			{
				System.Diagnostics.Process.Start("chrome.exe", sUrlNube);				
			}
		}

		private void btnEquifax_Click(object sender, System.EventArgs e)
		{
			#region Valida_DatosCliente
			if (!Validacion.vbTexto(this.txtDocumento, 1, 11, " Cedula ")) 
			{
				this.tbBasico.Tabs[0].Selected = true; 
				return;
			}
			if (!Validacion.vbTexto(this.txtApellidoP, 1, 60, " Paterno")) 
			{
				this.tbBasico.Tabs[0].Selected = true; 
				return;
			}
			if (!Validacion.vbTexto(this.txtPnombre,1, 60, " Nombre"))
			{
				this.tbBasico.Tabs[0].Selected = true; 
				return;
			}
			#endregion Valida_DatosCliente

			string Snombre = (this.txtPnombre.Text + " " +this.txtSnombre.Text + " " + this.txtApellidoP.Text + " " + this.txtApellidoM.Text );
			string sUrlNube = Funcion.sEscalarSQL(cdsSeteoF,string.Format("EXEC [RetornaURLCoordanadasEquifax] '{0}','{1}'",(string) this.txtDocumento.Text, Snombre)); 

			if (sUrlNube.Length > 10)
			{
				System.Diagnostics.Process.Start("chrome.exe", sUrlNube);				
			}
		}

	}
}
